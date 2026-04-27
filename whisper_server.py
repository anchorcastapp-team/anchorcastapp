#!/usr/bin/env python3
"""
AnchorCast Whisper Transcription Server  — optimised for live sermon STT
Receives WAV audio chunks from the Electron app, returns clean transcript text.

Usage:  python whisper_server.py [--model small.en] [--port 7777]

Recommended models (auto-downloaded on first run):
  tiny.en   ~39 MB   — fast on any CPU, decent accuracy
  base.en   ~74 MB   — good balance of speed and accuracy
  small.en  ~244 MB  — best accuracy for CPUs that can handle it  ← default
"""

import sys, os

# Suppress HuggingFace Hub warnings that are not relevant to AnchorCast users:
# 1. HF_HUB_DISABLE_SYMLINKS_WARNING — Windows without Developer Mode can't create
#    symlinks; HF falls back to copying files which works fine, just uses more disk.
# 2. HF_HUB_DISABLE_PROGRESS_BARS — cleaner output in the Electron console.
# 3. TOKENIZERS_PARALLELISM — avoids fork warning from tokenizers library.
os.environ.setdefault('HF_HUB_DISABLE_SYMLINKS_WARNING', '1')
os.environ.setdefault('HF_HUB_DISABLE_PROGRESS_BARS', '1')
os.environ.setdefault('TOKENIZERS_PARALLELISM', 'false')

# Suppress Python warnings for cleaner logs
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='huggingface_hub')
warnings.filterwarnings('ignore', message='.*symlink.*')
warnings.filterwarnings('ignore', message='.*HF_TOKEN.*')

# ── Python version gate ────────────────────────────────────────────────────────
vi = sys.version_info
if not (vi.major == 3 and 8 <= vi.minor <= 13):
    print(f'[Whisper] ERROR: Python {vi.major}.{vi.minor} not supported.', flush=True)
    print('[Whisper] faster-whisper requires Python 3.8–3.12', flush=True)
    sys.exit(1)

import json, argparse, io, re, threading, time
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn
from collections import deque

# ── CLI args ──────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument('--model',           default='small.en')
parser.add_argument('--port',            type=int, default=7777)
parser.add_argument('--device',          default='auto')
parser.add_argument('--model_cache_dir', default=None,
                    help='Directory to store Whisper models (default: HuggingFace cache)')
args = parser.parse_args()

IS_EN_MODEL  = args.model.endswith('.en')
IS_SMALL_UP  = any(x in args.model for x in ('small', 'medium', 'large'))

# ── Load model ────────────────────────────────────────────────────────────────
print(f'[Whisper] Loading {args.model} model...', flush=True)
try:
    from faster_whisper import WhisperModel
    device = args.device
    if device == 'auto':
        try:
            import torch
            device = 'cuda' if torch.cuda.is_available() else 'cpu'
        except ImportError:
            device = 'cpu'

    # int8 on CPU is fastest with negligible quality loss for small/base
    # float16 on CUDA for best GPU throughput
    compute = 'float16' if device == 'cuda' else 'int8'

    # Use app-specific model directory if provided, otherwise HuggingFace default cache
    model_kwargs = dict(device=device, compute_type=compute)
    if args.model_cache_dir:
        import os
        os.makedirs(args.model_cache_dir, exist_ok=True)
        model_kwargs['download_root'] = args.model_cache_dir
        print(f'[Whisper] Model cache: {args.model_cache_dir}', flush=True)

    model = WhisperModel(args.model, **model_kwargs)
    print(f'[Whisper] Model ready — {args.model} on {device} ({compute})', flush=True)
except ImportError:
    print('[Whisper] ERROR: faster-whisper not installed.', flush=True)
    print('[Whisper] Run: pip install faster-whisper', flush=True)
    sys.exit(1)
except Exception as e:
    err = str(e)
    print(f'[Whisper] ERROR loading model: {err}', flush=True)
    # Emit a detectable marker so the Electron app can show the right message
    if 'No such file' in err or 'not found' in err.lower() or 'download' in err.lower() or 'connection' in err.lower():
        print('[Whisper] MARKER:model_not_found', flush=True)
    sys.exit(1)

# ── Sliding context window ─────────────────────────────────────────────────────
# Stores last N clean transcript chunks.
# Used to:  (1) build a dynamic initial_prompt  (2) strip already-seen phrases
CONTEXT_WINDOW  = deque(maxlen=8)   # last ~24s of sermon
CONTEXT_LOCK    = threading.Lock()

# ── Reinforcement layer ────────────────────────────────────────────────────────
# When the detection engine identifies a verse, that verse text is injected here.
# It primes Whisper to recognise the same passage vocabulary in subsequent chunks,
# creating a positive feedback loop:
#   transcript → detection → verse text → Whisper prompt → better transcript
REINFORCE_TEXT  = ''           # most recently detected verse text
REINFORCE_LOCK  = threading.Lock()
REINFORCE_TTL   = 0           # Unix timestamp when reinforcement expires (30s)

def set_reinforcement(text: str, ttl_seconds: int = 30):
    """Set a verse/phrase as high-priority context for the next Whisper chunks."""
    global REINFORCE_TEXT, REINFORCE_TTL
    with REINFORCE_LOCK:
        REINFORCE_TEXT = text.strip()
        REINFORCE_TTL  = time.time() + ttl_seconds

def get_reinforcement() -> str:
    """Return active reinforcement text, or empty string if expired."""
    with REINFORCE_LOCK:
        if REINFORCE_TEXT and time.time() < REINFORCE_TTL:
            return REINFORCE_TEXT
        return ''

def add_to_context(text: str):
    with CONTEXT_LOCK:
        CONTEXT_WINDOW.append(text.strip())

def get_recent_context() -> str:
    with CONTEXT_LOCK:
        return ' '.join(CONTEXT_WINDOW)

# ── initial_prompt ─────────────────────────────────────────────────────────────
# Whisper uses the initial_prompt to prime its language model for the domain.
# The prompt is rebuilt every chunk so it includes the last ~180 chars of actual
# transcript — this massively improves continuity across chunk boundaries.
_BASE_PROMPT = (
    # Genre signal — the single most impactful thing in initial_prompt
    'Church sermon. Preaching. Bible reading. Scripture. Prayer. Worship. '
    # Bible books — keeps Whisper from substituting common English words
    'Genesis Exodus Leviticus Numbers Deuteronomy Joshua Judges Ruth Samuel Kings '
    'Chronicles Ezra Nehemiah Esther Job Psalms Proverbs Ecclesiastes Isaiah '
    'Jeremiah Lamentations Ezekiel Daniel Hosea Joel Amos Obadiah Jonah Micah '
    'Nahum Habakkuk Zephaniah Haggai Zechariah Malachi Matthew Mark Luke John '
    'Acts Romans Corinthians Galatians Ephesians Philippians Colossians '
    'Thessalonians Timothy Titus Philemon Hebrews James Peter Jude Revelation. '
    # Bible characters — prevents name substitutions
    'Abraham Isaac Jacob Moses David Solomon Elijah Elisha '
    'Paul Barnabas Silas Nicodemus Lazarus Pilate. '
    # Unique theological terms Whisper commonly mishears
    'propitiation atonement sanctification justification '
    'hallelujah amen.'
)

def build_initial_prompt() -> str:
    parts = [_BASE_PROMPT]

    # Reinforcement: if a verse was recently detected, inject its text verbatim.
    # Whisper will strongly prefer words that appear in the initial_prompt,
    # so this biases decoding toward the exact vocabulary of the active passage.
    reinforce = get_reinforcement()
    if reinforce:
        # Trim to 180 chars — enough to bias without overloading the prompt
        parts.append(reinforce[:180])

    recent = get_recent_context()
    if recent:
        # Last 200 chars of what was just said — continuity signal
        tail = recent[-200:].strip()
        tail = re.sub(r'^\S+\s+', '', tail)
        parts.append(tail)

    return ' '.join(parts)

# ── Transcription parameters (tuned per model) ─────────────────────────────────
# Rules of thumb for CPU-based faster-whisper:
#   beam_size  : higher = more accurate but slower. 5 is optimal for small.en on CPU.
#   best_of    : only matters when temperature > 0. Keep 1 when temp=0.
#   patience   : 1.0 = standard beam search cutoff.
#   temperature: 0.0 = fully greedy (most stable, no randomness).
#   repetition_penalty: >1.0 suppresses the repetition loops Whisper falls into.
#   no_repeat_ngram_size: prevents any 4-gram from appearing twice in one output.
#   condition_on_previous_text: False = each chunk decoded independently.
#       This PREVENTS hallucination chains but we compensate with initial_prompt.
#   compression_ratio_threshold: segments with ratio > 2.4 are likely loops → drop.
#   no_speech_threshold: 0.5 = ignore segments where model thinks there is no speech.
def _transcribe_params():
    return dict(
        language         = None if IS_EN_MODEL else 'en',
        beam_size        = 5,
        best_of          = 1,       # irrelevant at temp=0, keep low to save memory
        patience         = 1.0,
        temperature      = 0.0,
        repetition_penalty      = 1.3,
        no_repeat_ngram_size    = 5,
        condition_on_previous_text = False,
        initial_prompt   = build_initial_prompt(),
        vad_filter       = True,
        vad_parameters   = dict(
            # VAD (Voice Activity Detection) silences non-speech segments.
            # Too aggressive = clipped words. Too permissive = noise transcribed.
            min_silence_duration_ms = 300,   # 300ms of silence = end of utterance
            speech_pad_ms           = 200,   # pad 200ms around speech for context
            threshold               = 0.25,  # 0.25 = sensitive, catches soft voices
        ),
        no_speech_threshold          = 0.45,
        compression_ratio_threshold  = 2.3,
        word_timestamps              = False,  # off = faster, we don't need word timing
        hotwords = (
            'Psalms Psalm Isaiah Jeremiah Ezekiel Matthew Mark Luke John Acts Romans '
            'Corinthians Galatians Ephesians Philippians Colossians Thessalonians '
            'Timothy Hebrews James Peter Jude Revelation Genesis Exodus Proverbs '
            'Deuteronomy Leviticus Numbers Joshua Judges Chronicles Nehemiah Ecclesiastes '
            'Habakkuk Zephaniah Zechariah Malachi Obadiah Micah Nahum Haggai Lamentations '
            'Jesus Christ Amen Hallelujah '
            'Abraham Isaac Jacob Moses David Solomon Elijah Elisha Paul Barnabas Silas '
            'Nicodemus Lazarus Pilate '
            'propitiation atonement sanctification justification '
            'begotten crucified resurrection'
        ),
    )

# ── Post-processing pipeline ───────────────────────────────────────────────────

# 1. Book-name normalisation
_BOOK_FIXES = [
    (r'\brevelations\b',           'Revelation'),
    (r'\brevealation\b',           'Revelation'),
    (r'\brevolution\b(?=\s+\d)',    'Revelation'),
    (r'\bsong of songs\b',         'Song of Solomon'),
    (r'\bsong of Salomon\b',       'Song of Solomon'),
    (r'\bphilippian\b',            'Philippians'),
    (r'\bephesian\b',              'Ephesians'),
    (r'\bgalations\b',             'Galatians'),
    (r'\bcolosians\b',             'Colossians'),
    (r'\bcolosiens\b',             'Colossians'),
    (r'\bphilippines\b',           'Philippians'),
    (r'\bphillipians\b',           'Philippians'),
    (r'\bphillipians\b',           'Philippians'),
    (r'\bthessolonians\b',         'Thessalonians'),
    (r'\bdueteronomy\b',           'Deuteronomy'),
    (r'\bduetoronomy\b',           'Deuteronomy'),
    (r'\blamentation\b(?!s)',      'Lamentations'),
    (r'\bchonicles\b',             'Chronicles'),
    (r'\beclesiastes\b',           'Ecclesiastes'),
    (r'\becclesiates\b',           'Ecclesiastes'),
    (r'\bhebrew\b(?!s)',           'Hebrews'),
    (r'\bfirst\s+corinthians\b',   '1 Corinthians'),
    (r'\bsecond\s+corinthians\b',  '2 Corinthians'),
    (r'\b1st\s+corinthians\b',     '1 Corinthians'),
    (r'\b2nd\s+corinthians\b',     '2 Corinthians'),
    (r'\bfirst\s+kings\b',         '1 Kings'),
    (r'\bsecond\s+kings\b',        '2 Kings'),
    (r'\b1st\s+kings\b',           '1 Kings'),
    (r'\b2nd\s+kings\b',           '2 Kings'),
    (r'\bfirst\s+samuel\b',        '1 Samuel'),
    (r'\bsecond\s+samuel\b',       '2 Samuel'),
    (r'\b1st\s+samuel\b',          '1 Samuel'),
    (r'\b2nd\s+samuel\b',          '2 Samuel'),
    (r'\bfirst\s+chronicles\b',    '1 Chronicles'),
    (r'\bsecond\s+chronicles\b',   '2 Chronicles'),
    (r'\b1st\s+chronicles\b',      '1 Chronicles'),
    (r'\b2nd\s+chronicles\b',      '2 Chronicles'),
    (r'\bfirst\s+peter\b',         '1 Peter'),
    (r'\bsecond\s+peter\b',        '2 Peter'),
    (r'\b1st\s+peter\b',           '1 Peter'),
    (r'\b2nd\s+peter\b',           '2 Peter'),
    (r'\bfirst\s+john\b',          '1 John'),
    (r'\bsecond\s+john\b',         '2 John'),
    (r'\bthird\s+john\b',          '3 John'),
    (r'\b1st\s+john\b',            '1 John'),
    (r'\b2nd\s+john\b',            '2 John'),
    (r'\b3rd\s+john\b',            '3 John'),
    (r'\bfirst\s+timothy\b',       '1 Timothy'),
    (r'\bsecond\s+timothy\b',      '2 Timothy'),
    (r'\b1st\s+timothy\b',         '1 Timothy'),
    (r'\b2nd\s+timothy\b',         '2 Timothy'),
    (r'\bfirst\s+thessalonians\b', '1 Thessalonians'),
    (r'\bsecond\s+thessalonians\b','2 Thessalonians'),
    (r'\b1st\s+thessalonians\b',   '1 Thessalonians'),
    (r'\b2nd\s+thessalonians\b',   '2 Thessalonians'),
]

def _fix_book_names(text: str) -> str:
    for pat, rep in _BOOK_FIXES:
        text = re.sub(pat, rep, text, flags=re.I)
    return text

# 2. Biblical word corrections — acoustically similar pairs Whisper confuses
# Ordered: longer/more-specific patterns first to avoid partial matches
_WORD_FIXES = [
    # Psalm 23 — most commonly garbled passage
    ('shall not watch',        'shall not want'),
    ('shall not wash',         'shall not want'),
    ('shall not won',          'shall not want'),
    ('shall not what',         'shall not want'),
    ('I shall not pass',       'I shall not want'),
    ('I shall not past',       'I shall not want'),
    ('I shall not want to',    'I shall not want'),
    ('green pass ',            'green pastures '),
    ('green pace ',            'green pastures '),
    ('green pastor ',          'green pastures '),
    ('green pastas ',          'green pastures '),
    ('green passes ',          'green pastures '),
    ('green patches ',         'green pastures '),
    ('still water.',           'still waters.'),
    ('still water,',           'still waters,'),
    ('still water ',           'still waters '),
    ('restoreth my son',       'restoreth my soul'),
    ('restored my son',        'restoreth my soul'),
    ('restore my son',         'restoreth my soul'),
    ('restores my son',        'restoreth my soul'),
    ('restores my sole',       'restoreth my soul'),
    ('leadeth my son',         'leadeth my soul'),
    ('leadeth my sole',        'leadeth my soul'),
    ('leads my sole',          'leadeth my soul'),
    ('rivers of righteousness','paths of righteousness'),
    ('path of righteousness',  'paths of righteousness'),
    ('pass of righteousness',  'paths of righteousness'),
    ('parts of righteousness', 'paths of righteousness'),
    ('walk through the back',  'walk through the valley'),
    ('through the back',       'through the valley'),
    ('through the valley of the shadow death', 'through the valley of the shadow of death'),
    ('shadow death',           'shadow of death'),
    ('I will fear and all evil','I will fear no evil'),
    ('I will fear and evil',   'I will fear no evil'),
    ('fear and all evil',      'fear no evil'),
    ('fear and evil',          'fear no evil'),
    ('my cup runny over',      'my cup runneth over'),
    ('my cup running over',    'my cup runneth over'),
    ('my cup runs over',       'my cup runneth over'),
    ('my cup runny',           'my cup runneth over'),
    ('my cup running',         'my cup runneth over'),
    ('anoint my head',         'anointest my head'),
    ('anoints my head',        'anointest my head'),
    ('goodness mercy',         'goodness and mercy'),
    ('rod staff',              'rod and staff'),
    ('thy rod and stuff',      'thy rod and staff'),
    # Lord's prayer
    ('hollow be thy name',     'hallowed be thy name'),
    ('hallow be thy name',     'hallowed be thy name'),
    ('hollowed be thy name',   'hallowed be thy name'),
    ('hello be thy name',      'hallowed be thy name'),
    ('our daily broad',        'our daily bread'),
    ('our daily braid',        'our daily bread'),
    ('our daily Brett',        'our daily bread'),
    ('trespassers',            'trespasses'),
    ('deliver us from eagle',  'deliver us from evil'),
    # John 3:16 — most quoted verse
    ('God so loved the word',  'God so loved the world'),
    ('God\'s so loved',        'God so loved'),
    ('he gave his only begotten sun', 'he gave his only begotten Son'),
    ('begotten sun',           'begotten Son'),
    ('should not perish but have ever lasting', 'should not perish but have everlasting'),
    ('ever lasting life',      'everlasting life'),
    # Isaiah 40:31
    ('mount up with wings as legal', 'mount up with wings as eagles'),
    ('wings as legal',         'wings as eagles'),
    ('wings like legal',       'wings like eagles'),
    ('wings as e-girls',       'wings as eagles'),
    ('wait upon the Lord shall renew', 'wait upon the LORD shall renew'),
    # Romans 8:28
    ('all things work together forget', 'all things work together for good'),
    ('work together forget',   'work together for good'),
    # Proverbs 3:5-6
    ('lean not unto thine own understanding', 'lean not unto thine own understanding'),
    ('lean not on your own understanding', 'lean not unto thine own understanding'),
    ('lean not under',         'lean not unto'),
    ('acknowledge hymn',       'acknowledge him'),
    # Philippians 4:13
    ('I can do all things through crisis', 'I can do all things through Christ'),
    ('all things through crisis', 'all things through Christ'),
    ('through crisis which strengthens', 'through Christ which strengtheneth'),
    # Jeremiah 29:11
    ('thoughts of piece',      'thoughts of peace'),
    ('thoughts of piece and not of eagle', 'thoughts of peace and not of evil'),
    ('and expected end',       'an expected end'),
    # Matthew 28:19-20
    ('go ye there for',        'go ye therefore'),
    ('go ye therefore and teach all nations', 'go ye therefore and teach all nations'),
    ('baptize in them',        'baptizing them'),
    # Common church word confusions
    ('thigh kingdom',          'thy kingdom'),
    ('thighs kingdom',         'thy kingdom'),
    ('thy king dumb',          'thy kingdom'),
    ('ethernet',               'eternal'),
    ('hollowly',               'hallowed'),
    ('sabour',                 'saviour'),
    ('the names sake',         "his name's sake"),
    ('for his names sake',     "for his name's sake"),
    ('name sake',              "name's sake"),
    (' savor ',                ' saviour '),
    (' righty ',               ' righteous '),
    (' rightness ',            ' righteousness '),
    ('rightness ',             'righteousness '),
    ('rightness.',             'righteousness.'),
    ('rightness,',             'righteousness,'),
    ('right to us',            'righteous'),
    ('right chess',            'righteous'),
    (' iniquity is',           ' iniquities'),
    (' in equity',             ' iniquity'),
    ('transgression is',       'transgressions'),
    ('trans gressions',        'transgressions'),
    ('in her it',              'inherit'),
    ('in haircut',             'inherit'),
    ('sanctify occasion',      'sanctification'),
    ('sanct if ication',       'sanctification'),
    ('just a vacation',        'justification'),
    (' reconcile Asian',       ' reconciliation'),
    ('propitiate ion',         'propitiation'),
    (' redeem shun',           ' redemption'),
    (' atonement for',         ' atonement for'),
    ('inter session',          'intercession'),
    ('inter cede',             'intercede'),
    # Archaic KJV pronoun confusions
    ('thine art',              'thou art'),
    ('died art',               'thou art'),
    (' thinned ',              ' thine '),
    ('for give us',            'forgive us'),
    # Common name confusions
    ('a Brahm',                'Abraham'),
    ('more says',              'Moses'),
    ('most is',                'Moses'),
    ('ice sake',               'Isaac'),
    ('jack up',                'Jacob'),
    ('just is',                'Jesus'),
    ('pie let',                'Pilate'),
    ('pilates',                'Pilate'),
    ('barrel bus',             'Barabbas'),
    ('bear a bus',             'Barabbas'),
    ('disciples hip',          'discipleship'),
    ('fellow ship',            'fellowship'),
    ('faith full',             'faithful'),
    ('faith full ness',        'faithfulness'),
    ('right just',             'righteous'),
    # Archaic KJV words Whisper modernises — keep originals
    ('leadeth me',             'leadeth me'),
    ('restoreth',              'restoreth'),
    ('saith',                  'saith'),
    ('hath',                   'hath'),
    ('verily verily i say unto', 'verily verily I say unto'),
    ('in the beginning was the word', 'In the beginning was the Word'),
    ('and the word was with god',     'and the Word was with God'),
    ('and the word was god',          'and the Word was God'),

    # ── Corrections from live Deepgram sermon (Apr 26, 2026) ──────────────
    # Biblical names misheard
    ('zakius',                 'Zacchaeus'),
    ('zakeus',                 'Zacchaeus'),
    ('zaceus',                 'Zacchaeus'),
    ('zacheus',                'Zacchaeus'),
    ('besali',                 'Bezalel'),
    ('desally',                'Bezalel'),
    ('besaly',                 'Bezalel'),
    ('postmortem',             'Paul'),    # "Paul" misheard as "postmortem"
    ('polymer',                'Paul'),    # "Paul" misheard as "polymer"

    # "Grace" misheard as other words (common in African English accents)
    ('decrease of god',        'grace of God'),
    ('decrease of law',        'grace of law'),  # "grace of the law"
    ('increase of gold',       'grace of God'),
    ('grade of god',           'grace of God'),
    ('graze of god',           'grace of God'),

    # "God" misheard in fast speech
    ('grace of job',           'grace of God'),
    ('grace of gold',          'grace of God'),
    ('word of go ',            'word of God '),
    ('kingdom of go ',         'kingdom of God '),
    ('people of go ',          'people of God '),
    ('children of go ',        'children of God '),
    ('man of go ',             'man of God '),

    # Profanity that is a Biblical phrase
    ('noah fuck grace',        'Noah found grace'),
    ('fuck grace',             'found grace'),
    ('what pussy',             'what passage'),

    # Book name mishears
    ('in fifteenth chapter',   'Ephesians chapter'),
    ('amos two was talking about', '1 Peter was talking about'),

    # Misc
    ('the antiseper',          'the answer is'),
    ('chris open',             'Christ upon'),
    ('postmortem was talking', 'Paul was talking'),
    ('grace of the job',       'grace of God'),

    # Sermon 2 corrections (Apr 19, 2026)
    ('sad vision',             'salvation'),
    ('for sad vision',         'for salvation'),
    ('affixure',               'apostle'),
    ('an affixure',            'an apostle'),
    ('bysecuting',             'persecuting'),
    ('cycling ship',           'keeping sheep'),
    ('kick ship',              'kingship'),
    ('messi says',             'mercy says'),
    ('messy says',             'mercy says'),
    ('unselfly',               'unselfishly'),
    ('unmerited female',       'unmerited favor'),
    ('scriptatory',            'scriptures'),
    ('grass of god',           'grace of God'),
    ('sad addition',           'salvation'),

    # Sermon 3 corrections (Apr 12, 2026) — FOCUS sermon
    # CRITICAL: 'focus' → 'fuck us' in African English
    ('fuck us',                'focus'),
    ('if you are watching, fuck us', 'if you are watching, focus'),
    ('fuck us on',             'focus on'),
    # 1 Kings misheard
    ('first kick from the',    '1 Kings chapter'),
    ('fourth kings',           '1 Kings'),
    ('first kick',             '1 Kings'),
    # Ephesians variant
    ('ephysians',              'Ephesians'),
    # Names
    ('zolom',                  'Solomon'),
    # Mishears
    ('fraud stealing in god',  'trusting in God'),
    ('fraud stealing',         'trusting'),
    ('why your son was pissed','why your son was busy'),
    ('jack of poultry',        'jack of all trades'),
    ('master of no ',          'master of none '),


    # Sermon 6 corrections (Apr 26, 2026) — Blessing of Work (Deepgram)
    ('locust placed man',          'Lord placed man'),
    ('placed man in the daddy',    'placed man in the garden'),
    ('hard water has plenty',      'hard worker has plenty'),
    ('delacing and become a slave','be lazy and become a slave'),
    ('may not live to poverty',    'mere talk leads to poverty'),
    ('lazy people work much',      'lazy people want much'),
    ('foursight',                  'foresight'),
    ('titty',                      'duty'),
    ('boiling the midnight oil',   'burning the midnight oil'),
    ('the copier',                 'the cupbearer'),
    ('a bandage scour seating',    'like a bandit'),

    # Sermon 5 corrections (Apr 26, 2026) — Blessing of Work
    ('procter',                'Proverbs'),
    ('proctor',                'Proverbs'),
    ('their lungs',            'their lamps'),
    ('our lungs',              'our lamps'),
    ('the lungs',              'the lamps'),
    ('lungs are going',        'lamps are going out'),
    ('having suicide',         'having foresight'),
    ('four sides',             'foresight'),
    ('suicide is the ability', 'foresight is the ability'),
    ('learning from the aunts','learning from the ants'),
    ('the aunts today',        'the ants today'),
    ('the aunt this morning',  'the ant this morning'),
    ('atonement is very weak', 'your amen is very weak'),
    ("it's delicious",         'diligence'),
    ('propitiation for yourself', 'provision for yourself'),
    ('propitiation for others',   'provision for others'),

    # Sermon 4 corrections (Apr 15, 2026) — New Life in Christ
    ('galicias chapter',       'Galatians chapter'),
    ("galicia's chapter",      'Galatians chapter'),
    ('collisions chapter',     'Colossians chapter'),
    ('collisions',             'Colossians'),
    ('converseius chapter',    'Colossians chapter'),
    ('fishes number',          'Ephesians'),
    ('joint ears',             'joint heirs'),
    ('ears of god',            'heirs of God'),
    ('we are afraid of god',   'we are heirs of God'),
    ('we are ears',            'we are heirs'),
    ('the fourth born of all creation', 'the firstborn of all creation'),
    ('fourth born',            'firstborn'),
    ('pebillion people',       'peculiar people'),
    ('save conscious',         'righteousness conscious'),
    ('kisos',                  'Jesus'),
    ('new payment',            'new man'),
]

def _fix_words(text: str) -> str:
    for old, new in _WORD_FIXES:
        text = text.replace(old, new)
        # also case-insensitive for start of sentence
        text = re.sub(re.escape(old), new, text, flags=re.I)
    return text

# 3. Verbal verse reference conversion: "James one twelve" → "James 1:12"
_WORD_NUMS = {
    'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,'eight':8,
    'nine':9,'ten':10,'eleven':11,'twelve':12,'thirteen':13,'fourteen':14,
    'fifteen':15,'sixteen':16,'seventeen':17,'eighteen':18,'nineteen':19,
    'twenty':20,'thirty':30,'forty':40,'fifty':50,'sixty':60,'seventy':70,
    'eighty':80,'ninety':90,'hundred':100,
    'first':1,'second':2,'third':3,'fourth':4,'fifth':5,'sixth':6,
    'seventh':7,'eighth':8,'ninth':9,'tenth':10,'eleventh':11,'twelfth':12,
    'thirteenth':13,'fourteenth':14,'fifteenth':15,'sixteenth':16,
    'seventeenth':17,'eighteenth':18,'nineteenth':19,'twentieth':20,
}
_NUM_ALT = '|'.join(sorted(_WORD_NUMS.keys(), key=len, reverse=True))
_TENS    = 'twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety'
_NUMTOK  = rf'(?:(?:{_TENS})\s+(?:{_NUM_ALT})|\d+|{_NUM_ALT})'
_BOOKS_FOR_REF = sorted([
    'Genesis','Exodus','Leviticus','Numbers','Deuteronomy','Joshua','Judges',
    'Ruth','Esther','Job','Psalms','Psalm','Proverbs','Ecclesiastes','Isaiah',
    'Jeremiah','Lamentations','Ezekiel','Daniel','Hosea','Joel','Amos',
    'Obadiah','Jonah','Micah','Nahum','Habakkuk','Zephaniah','Haggai',
    'Zechariah','Malachi','Matthew','Mark','Luke','John','Acts','Romans',
    'Galatians','Ephesians','Philippians','Colossians','Titus','Philemon',
    'Hebrews','James','Jude','Revelation','Song of Solomon',
    '1 Corinthians','2 Corinthians','1 Thessalonians','2 Thessalonians',
    '1 Timothy','2 Timothy','1 Samuel','2 Samuel','1 Kings','2 Kings',
    '1 Chronicles','2 Chronicles','1 Peter','2 Peter','1 John','2 John','3 John',
], key=len, reverse=True)

def _w2n(s: str):
    s = s.lower().strip()
    if re.match(r'^\d+$', s): return int(s)
    if s in _WORD_NUMS:       return _WORD_NUMS[s]
    parts = s.split()
    if len(parts) == 2 and parts[0] in _WORD_NUMS and parts[1] in _WORD_NUMS:
        return _WORD_NUMS[parts[0]] + _WORD_NUMS[parts[1]]
    return None

def _fix_verbal_refs(text: str) -> str:
    out = text
    for book in _BOOKS_FOR_REF:
        pat3 = re.compile(
            rf'\b({re.escape(book)})\s+({_NUMTOK})\s+(?:verse\s+)?({_NUMTOK})\s+(?:through|to|thru|dash)\s+({_NUMTOK})\b', re.I
        )
        def _repl_range(m):
            ch = _w2n(m.group(2).strip())
            vs = _w2n(m.group(3).strip())
            ve = _w2n(m.group(4).strip())
            if ch and vs and ve:
                return f'{m.group(1)} {ch}:{vs}-{ve}'
            return m.group(0)
        out = pat3.sub(_repl_range, out)

        pat = re.compile(
            rf'\b({re.escape(book)})\s+({_NUMTOK})\s+(?:verse\s+)?({_NUMTOK})\b', re.I
        )
        def _repl(m):
            ch = _w2n(m.group(2).strip())
            vs = _w2n(m.group(3).strip())
            return f'{m.group(1)} {ch}:{vs}' if ch and vs else m.group(0)
        out = pat.sub(_repl, out)

        pat_ch = re.compile(
            rf'\b({re.escape(book)})\s+chapter\s+({_NUMTOK})\b', re.I
        )
        def _repl_ch(m):
            ch = _w2n(m.group(2).strip())
            return f'{m.group(1)} {ch}' if ch else m.group(0)
        out = pat_ch.sub(_repl_ch, out)

    out = re.sub(
        rf'\bchapter\s+({_NUMTOK})\s+verse\s+({_NUMTOK})\s+(?:through|to|thru)\s+({_NUMTOK})\b',
        lambda m: f'chapter {_w2n(m.group(1))} verse {_w2n(m.group(2))}-{_w2n(m.group(3))}',
        out, flags=re.I
    )
    out = re.sub(
        rf'\bchapter\s+({_NUMTOK})\s+verse\s+({_NUMTOK})\b',
        lambda m: f'chapter {_w2n(m.group(1))} verse {_w2n(m.group(2))}',
        out, flags=re.I
    )
    out = re.sub(
        rf'\bverses?\s+({_NUMTOK})\s+(?:through|to|thru)\s+({_NUMTOK})\b',
        lambda m: f'verses {_w2n(m.group(1))}-{_w2n(m.group(2))}',
        out, flags=re.I
    )
    return out

# 4. Repetition removal — Whisper hallucination loops
def _fix_repetitions(text: str) -> str:
    # Pass 1: phrase repeats (4-80 chars), comma/period/space separated
    out = re.sub(r'\b(.{4,80}?)(?:[,.\s]+\1)+\b', r'\1', text, flags=re.I)
    # Pass 2: word-pair repeats
    out = re.sub(r'\b(\w+\s+\w+)[,. ]+\1\b', r'\1', out, flags=re.I)
    # Pass 3: single-word repeats
    out = re.sub(r'\b(\w{3,})[,. ]+\1\b', r'\1', out, flags=re.I)
    return re.sub(r'\s+', ' ', out).strip()

# 5. Context dedup — strip phrases already emitted in the last few chunks
def _dedup_context(text: str) -> str:
    """Strip overlap at the START of text that matches the END of recent context.
    Only trims the leading boundary overlap — does NOT remove phrases from
    the middle of the new chunk, which would delete legitimate repeated speech."""
    recent = get_recent_context().lower()
    if not recent:
        return text
    words = text.split()
    if len(words) < 4:
        return text

    # Only check for overlap at the START of the new chunk vs END of recent context
    # Try overlaps from longest (8 words) down to 3 words
    for overlap_len in range(min(8, len(words)), 2, -1):
        candidate = ' '.join(words[:overlap_len]).lower()
        if candidate in recent[-200:]:   # only check tail of recent context
            words = words[overlap_len:]
            break

    result = ' '.join(words).strip()
    return result if len(result) > 5 else text

# 6. Segment-level quality gate
def _good_segment(seg) -> bool:
    """Return False for segments that are likely hallucinations."""
    # High compression ratio → repeated tokens → hallucination
    cr = getattr(seg, 'compression_ratio', 0) or 0
    if cr > 2.3:
        return False
    # avg_logprob < -1.0 → model is very uncertain → noise
    lp = getattr(seg, 'avg_logprob', 0) or 0
    if lp < -1.0:
        return False
    return True

# ── Full post-processing pipeline ──────────────────────────────────────────────
def process_transcript(raw: str) -> str:
    if not raw or not raw.strip():
        return ''
    t = raw.strip()
    t = _fix_book_names(t)
    t = _fix_verbal_refs(t)
    t = _fix_words(t)
    t = _fix_repetitions(t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

# ── Transcription lock ─────────────────────────────────────────────────────────
_lock = threading.Lock()

# ── HTTP server ────────────────────────────────────────────────────────────────
class WhisperHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *a): pass  # suppress request logs

    def do_GET(self):
        if self.path == '/health':
            self._json(200, {'status': 'ok', 'model': args.model, 'device': device})
        else:
            self._json(404, {'error': 'not found'})

    def do_POST(self):
        # /reinforce — receive detected verse text from the app for context priming
        if self.path == '/reinforce':
            length = int(self.headers.get('Content-Length', 0))
            if length:
                body = self.rfile.read(length)
                try:
                    data = json.loads(body)
                    text = str(data.get('text', '')).strip()
                    ttl  = int(data.get('ttl', 30))
                    if text:
                        set_reinforcement(text, ttl)
                        add_to_context(text)  # also add to rolling context
                        print(f'[Whisper] Reinforcement set: "{text[:60]}"', flush=True)
                        self._json(200, {'ok': True})
                        return
                except Exception:
                    pass
            self._json(400, {'error': 'invalid body'})
            return

        if self.path != '/transcribe':
            self._json(404, {'error': 'not found'}); return

        length = int(self.headers.get('Content-Length', 0))
        if not length:
            self._json(400, {'error': 'no audio data'}); return

        audio_data = self.rfile.read(length)
        transcript  = ''

        try:
            with _lock:
                buf  = io.BytesIO(audio_data)
                params = _transcribe_params()
                segments, _info = model.transcribe(buf, **params)

                # Collect high-quality segments only
                parts = []
                for seg in segments:
                    t = seg.text.strip()
                    if t and _good_segment(seg):
                        parts.append(t)

                raw = ' '.join(parts).strip()
                if raw:
                    transcript = process_transcript(raw)
                    transcript = _dedup_context(transcript)
                    if transcript and len(transcript) > 8:
                        add_to_context(transcript)

        except Exception as e:
            print(f'[Whisper] Transcription error: {e}', flush=True)
            try:
                self._json(500, {'error': str(e)}); return
            except OSError:
                return

        if transcript:
            print(f'[Whisper] "{transcript[:100]}"', flush=True)

        try:
            self._json(200, {'text': transcript})
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError, OSError):
            pass

    def _json(self, code, data):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', len(body))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def handle_error(self, request, client_address):
        import traceback
        tb = traceback.format_exc()
        if any(x in tb for x in ['BrokenPipe','ConnectionAbort','ConnectionReset','10053','10054']):
            return
        super().handle_error(request, client_address)

class ThreadedServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True

# ── Start ──────────────────────────────────────────────────────────────────────
server = ThreadedServer(('127.0.0.1', args.port), WhisperHandler)
print(f'[Whisper] Server listening on http://127.0.0.1:{args.port}', flush=True)
server.serve_forever()

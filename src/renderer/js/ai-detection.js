// AnchorCast AI Detection Engine
// Detects Bible verses from live sermon transcript
// Works offline (rule-based + keyword + content search) and online (Claude AI)

window.AIDetection = (() => {
  let apiKey        = '';
  let onDetection   = null;
  let aiEnabled     = true;   // controls only the online Claude AI — offline always runs
  let debounceTimer = null;
  let detectionCache = new Map();
  let recentEmits    = new Map();
  let learnedPhrases = [];

  // Rolling window — 6 lines gives ~18s of context, enough for cross-chunk refs
  // without accumulating too much stale text that causes false positives
  const WINDOW_SIZE = 6;
  let recentLines = [];
  let lastWindowHash = '';

  // Last-known book+chapter context for orphan verse references.
  // When the preacher says "John 11" and later "verse 43", the orphan
  // "verse 43" is connected back to John 11 → John 11:43.
  let lastBookChapter = null;  // { book, chapter, timestamp }
  let lastEmittedVerse = null;  // { book, chapter, verse, timestamp }
  const LAST_CONTEXT_TTL = 90000; // 90s — context expires after this


  async function init(key, callback) {
    apiKey      = key;
    onDetection = callback;
    await reloadLearnedPhrases();
  }

  function setEnabled(val) {
    aiEnabled = val; // only gates the Claude API call — offline always runs
  }

  // ── Emit with smart deduplication ────────────────────────────────────────
  // Cooldown is keyed on ref ONLY (not type) so the same verse doesn't
  // fire from verbal, keyword, AND content search within the same window.
  function emit(det) {
    if (!onDetection || !det?.ref) return;
    // skip low-confidence detections from noisy methods
    if ((det.confidence || 0) < 0.55) return;
    const key = det.ref;
    const now  = Date.now();
    const last = recentEmits.get(key) || 0;
    // direct = 10s cooldown (preacher may re-emphasise); others = 45s
    const cooldownMs = det.type === 'direct' ? 10000 : 45000;
    if (now - last < cooldownMs) return;
    recentEmits.set(key, now);
    if (recentEmits.size > 150) recentEmits.delete(recentEmits.keys().next().value);
    if (det.book && det.chapter && det.verse) {
      lastEmittedVerse = { book: det.book, chapter: det.chapter, verse: det.verse, timestamp: now };
    }
    onDetection(det);
  }

  // ── Learned phrases ───────────────────────────────────────────────────────
  async function reloadLearnedPhrases() {
    try {
      const data = await window.electronAPI?.loadDetectionReviewData?.();
      learnedPhrases = Array.isArray(data?.phrases) ? data.phrases : [];
    } catch (_) { learnedPhrases = []; }
    return learnedPhrases;
  }

  function detectLearnedPhrases(text) {
    if (!window.BibleDB) return [];
    const hay = normalizeDetectionText(text);
    const results = [];
    for (const item of (learnedPhrases || [])) {
      const phrase = normalizeDetectionText(item.phrase || '');
      if (!phrase || !hay.includes(phrase)) continue;
      const parsed = window.BibleDB.parseReference?.(String(item.ref || ''));
      if (!parsed) continue;
      const verseText = window.BibleDB.getVerse(parsed.book, parsed.chapter, parsed.verse || 1);
      if (!verseText) continue;
      results.push({
        book: parsed.book, chapter: parsed.chapter, verse: parsed.verse || 1,
        ref: `${parsed.book} ${parsed.chapter}:${parsed.verse || 1}`,
        text: verseText, type: 'learned', confidence: 0.96,
        sourcePhrase: item.phrase || '',
      });
    }
    return results;
  }

  // ── Text normalisation ────────────────────────────────────────────────────
  function normalizeDetectionText(text) {
    return String(text || '')
      .toLowerCase()
      .replace(/\brevelations\b/g, 'revelation')
      .replace(/\bresolutions\s+(?=chapter|chap|\d)/g, 'revelation ')
      .replace(/\brevolution\s+(?=chapter|chap|\d)/g, 'revelation ')
      .replace(/\baphesians\b/g, 'ephesians')
      .replace(/\befesians\b/g, 'ephesians')
      .replace(/\bfilipians\b/g, 'philippians')
      .replace(/\bfilippians\b/g, 'philippians')
      .replace(/\bcollosians\b/g, 'colossians')
      .replace(/\bcolosians\b/g, 'colossians')
      .replace(/\bthessalonions\b/g, 'thessalonians')
      .replace(/\bthesalonians\b/g, 'thessalonians')
      .replace(/\bduteronomy\b/g, 'deuteronomy')
      .replace(/\blevitikus\b/g, 'leviticus')
      .replace(/\becclesiasties\b/g, 'ecclesiastes')
      .replace(/\becclesiast\b/g, 'ecclesiastes')
      .replace(/\becclesiat\b/g, 'ecclesiastes')
      .replace(/\bfirst\s+/g, '1 ')
      .replace(/\bsecond\s+/g, '2 ')
      .replace(/\bthird\s+/g, '3 ')
      .replace(/[^a-z0-9:\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // ── 1. DIRECT REFERENCE DETECTION ────────────────────────────────────────
  // Catches both numeric ("Romans 8:5") and spoken-number ("Romans eight five")
  function detectDirect(text) {
    if (!window.BibleDB) return [];
    const results = window.BibleDB.detectReferences(text) || [];

    // Also detect spoken-number references:
    //   "Romans eight five", "Psalm twenty three one",
    //   "Ephesians chapter one verse seven",
    //   "Ephesians chapter number one verse seven"
    const lower = text.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
    const spokenWord = '(?:one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)';
    const numToken = `(?:\\d+|${spokenWord}(?:\\s+${spokenWord})?)`;
    const bookPattern = Object.keys(BOOK_NAMES)
      .sort((a, b) => b.length - a.length)
      .map(b => b.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
      .join('|');
    const chPrefix = '(?:chapter\\s+(?:number\\s+)?|chap(?:ter)?\\s+)?';
    const vPrefix = '(?:verses?\\s+(?:number\\s+)?|v\\.?\\s+)?';
    const spokenRefRe = new RegExp(`(${bookPattern})\\s+${chPrefix}(${numToken})(?:\\s+${vPrefix}(?:and\\s+)?(${numToken}))?(?:\\s|$|,|\\.)`, 'gi');
    const loosVerseRe = new RegExp(`verses?\\s+(?:number\\s+)?(${numToken})`, 'gi');
    let m;
    while ((m = spokenRefRe.exec(lower)) !== null) {
      const bookStr = m[1]?.trim();
      const chStr = m[2]?.trim();
      let vStr = m[3]?.trim();
      if (!bookStr || !chStr) continue;
      let ch = wordsToNumber(chStr);
      if (!ch) continue;
      let v = vStr ? wordsToNumber(vStr) : null;

      // Smart number-splitting: when Whisper merges "7 9" into "79" and
      // the book doesn't have that many chapters, try splitting the digits.
      // e.g., "Job 79" → Job has 42 chapters → try ch=7, v=9.
      if (!v && ch > 9 && String(ch).length >= 2) {
        const chDigits = String(ch);
        const testRef = window.BibleDB.parseReference?.(`${bookStr} ${ch}:1`);
        const hasChapter = testRef && window.BibleDB.getVerse(testRef.book, ch, 1);
        if (!hasChapter) {
          for (let splitAt = 1; splitAt < chDigits.length; splitAt++) {
            const tryC = parseInt(chDigits.slice(0, splitAt));
            const tryV = parseInt(chDigits.slice(splitAt));
            if (!tryC || !tryV) continue;
            const splitRef = window.BibleDB.parseReference?.(`${bookStr} ${tryC}:${tryV}`);
            if (splitRef && window.BibleDB.getVerse(splitRef.book, tryC, tryV)) {
              ch = tryC;
              v = tryV;
              break;
            }
          }
        }
      }

      // If no verse found adjacent, look ahead on the same text for "verse N"
      // e.g., "John 1, I will live from verse one" → connects verse one to John 1
      if (!v) {
        const afterMatch = lower.slice(m.index + m[0].length, m.index + m[0].length + 120);
        loosVerseRe.lastIndex = 0;
        const lvm = loosVerseRe.exec(afterMatch);
        if (lvm) {
          const candidateV = wordsToNumber(lvm[1]?.trim());
          if (candidateV) v = candidateV;
        }
      }
      const bibleRef = window.BibleDB.parseReference?.(`${bookStr} ${ch}:${v || 1}`);
      if (!bibleRef) continue;
      const verseText = window.BibleDB.getVerse(bibleRef.book, ch, v || 1);
      if (!verseText) continue;
      if (results.some(r => r.book === bibleRef.book && r.chapter === ch && r.verse === (v || 1))) continue;
      // Guard: "Genesis 5 and 6" — two chapters connected by 'and' without
      // a verse keyword is a CHAPTER RANGE announcement, not Book 5:6.
      // Detect by checking if the match text contains "and N" as the verse
      // and no "verse" keyword preceded the number.
      const matchedText = m[0] || '';
      const hasVerseKeyword = /\bverses?\s/i.test(matchedText) || /\bv\.?\s*\d/i.test(matchedText);
      const isChapterRange = !hasVerseKeyword && /\band\s+\d/i.test(matchedText) && !v;
      // Also suppress: verse captured was literally from "and N" pattern with no verse keyword
      const verseFromAnd = !hasVerseKeyword && m[3] && /^(and\s+)?\d+$/.test(m[3].trim()) && !v;
      if (isChapterRange || verseFromAnd) {
        // Treat as chapter-only — emit verse 1 but with very low confidence
        // so the navigation buffer + dedup can suppress it
        results.push({
          bookIdx: BOOK_NAMES[bookStr], book: bibleRef.book, chapter: ch, verse: 1,
          ref: `${bibleRef.book} ${ch}:1`,
          confidence: 0.60, type: 'direct',
          _chapterOnly: true, _chapterRange: true,
        });
        continue;
      }

      results.push({
        bookIdx: BOOK_NAMES[bookStr], book: bibleRef.book, chapter: ch, verse: v || 1,
        ref: `${bibleRef.book} ${ch}:${v || 1}`,
        confidence: v ? 0.95 : 0.92, type: 'direct',
        _chapterOnly: !v,
      });
    }
    return results;
  }

  // ── 2. VERBAL MENTION DETECTION ──────────────────────────────────────────
  const BOOK_NAMES = {
    'genesis':1,'exodus':2,'leviticus':3,'numbers':4,'deuteronomy':5,
    'joshua':6,'judges':7,'ruth':8,'first samuel':9,'1 samuel':9,'2 samuel':10,
    'second samuel':10,'first kings':11,'1 kings':11,'second kings':12,'2 kings':12,
    'first chronicles':13,'1 chronicles':13,'second chronicles':14,'2 chronicles':14,
    'ezra':15,'nehemiah':16,'esther':17,'job':18,'psalms':19,'psalm':19,
    'proverbs':20,'ecclesiastes':21,'song of solomon':22,'isaiah':23,'jeremiah':24,
    'lamentations':25,'ezekiel':26,'daniel':27,'hosea':28,'joel':29,'amos':30,
    'obadiah':31,'jonah':32,'micah':33,'nahum':34,'habakkuk':35,'zephaniah':36,
    'haggai':37,'zechariah':38,'malachi':39,'matthew':40,'mark':41,'luke':42,
    'john':43,'acts':44,'romans':45,'first corinthians':46,'1 corinthians':46,
    'second corinthians':47,'2 corinthians':47,'galatians':48,'ephesians':49,
    'philippians':50,'colossians':51,'first thessalonians':52,'1 thessalonians':52,
    'second thessalonians':53,'2 thessalonians':53,'first timothy':54,'1 timothy':54,
    'second timothy':55,'2 timothy':55,'titus':56,'philemon':57,'hebrews':58,
    'james':59,'first peter':60,'1 peter':60,'second peter':61,'2 peter':61,
    'first john':62,'1 john':62,'second john':63,'2 john':63,'third john':64,
    '3 john':64,'jude':65,'revelation':66,
  };

  const NUMBER_WORDS = {
    'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,'eight':8,
    'nine':9,'ten':10,'eleven':11,'twelve':12,'thirteen':13,'fourteen':14,
    'fifteen':15,'sixteen':16,'seventeen':17,'eighteen':18,'nineteen':19,
    'twenty':20,'thirty':30,'forty':40,'fifty':50,'sixty':60,
    'first':1,'second':2,'third':3,'fourth':4,'fifth':5,
    'sixth':6,'seventh':7,'eighth':8,'ninth':9,'tenth':10,
  };

  function wordsToNumber(str) {
    if (!str) return null;
    const n = parseInt(str);
    if (!isNaN(n)) return n;
    const lower = str.toLowerCase().trim();
    if (NUMBER_WORDS[lower]) return NUMBER_WORDS[lower];
    const parts = lower.split(/\s+/);
    if (parts.length === 2 && NUMBER_WORDS[parts[0]] && NUMBER_WORDS[parts[1]])
      return NUMBER_WORDS[parts[0]] + NUMBER_WORDS[parts[1]];
    return null;
  }

  function detectVerbal(text) {
    if (!window.BibleDB) return [];
    const results = [];
    const lower = normalizeDetectionText(text);

    const bookPattern = Object.keys(BOOK_NAMES)
      .sort((a, b) => b.length - a.length)
      .map(b => b.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
      .join('|');

    const patterns = [
      new RegExp(`(?:turn to|go to|look at|open to|read from|found in|\\bin)\\s+(${bookPattern})(?:\\s+(?:chapter|chap|ch\\.?))?\\s+(\\w+)(?:\\s+(?:verse|v\\.?|and verse)\\s+(\\w+))?`, 'gi'),
      new RegExp(`(${bookPattern})\\s+(?:chapter|chap|ch\\.?)\\s+(\\w+)(?:\\s+(?:verse|v\\.?)\\s+(\\w+))?`, 'gi'),
      new RegExp(`(?:bible says|scripture says|it says|word says|god says|jesus said|jesus says|the lord says)(?:\\s+in)?\\s+(${bookPattern})\\s+(\\w+)(?:\\s+(\\w+))?`, 'gi'),
      new RegExp(`(?:written in|according to|as \\w+ (?:said|wrote|declared) in)\\s+(${bookPattern})\\s+(\\w+)(?:\\s+(\\w+))?`, 'gi'),
      new RegExp(`(?:the psalmist (?:said|says|asked|wrote)|as david said)\\s+(?:in\\s+)?(${bookPattern})\\s+(\\w+)(?:\\s+(\\w+))?`, 'gi'),
    ];

    for (const pattern of patterns) {
      let match;
      pattern.lastIndex = 0;
      while ((match = pattern.exec(lower)) !== null) {
        const bookStr = match[1]?.toLowerCase().trim();
        const chStr   = match[2]?.trim();
        const vStr    = match[3]?.trim();
        if (!bookStr || !chStr) continue;
        const ch = wordsToNumber(chStr);
        const v  = vStr ? wordsToNumber(vStr) : 1;
        if (!ch || !v) continue;
        const bibleRef = window.BibleDB.parseReference?.(`${bookStr} ${ch}:${v}`);
        if (bibleRef) {
          const verseText = window.BibleDB.getVerse(bibleRef.book, ch, v);
          if (verseText) results.push({
            book: bibleRef.book, chapter: ch, verse: v,
            ref: `${bibleRef.book} ${ch}:${v}`,
            text: verseText, type: 'verbal', confidence: 0.90,
          });
        }
      }
    }
    return results;
  }

  // ── 3. KEYWORD DETECTION ─────────────────────────────────────────────────
  function detectKeyword(text) {
    if (!window.BibleDB) return [];
    const results = [];
    const signatures = [
      { pattern: /god so loved the world/i,                    book:'John',          ch:3,  v:16 },
      { pattern: /lord is my shepherd/i,                       book:'Psalms',        ch:23, v:1  },
      { pattern: /all things work together for good/i,         book:'Romans',        ch:8,  v:28 },
      { pattern: /mount up with wings.{0,20}eagles/i,          book:'Isaiah',        ch:40, v:31 },
      { pattern: /can do all things through christ/i,          book:'Philippians',   ch:4,  v:13 },
      { pattern: /nothing.{0,15}impossible.{0,10}with god/i,    book:'Luke',          ch:1,  v:37 },
      { pattern: /i know the plans i have/i,                   book:'Jeremiah',      ch:29, v:11 },
      { pattern: /plans.{0,20}prosper.{0,10}not.{0,10}harm/i,  book:'Jeremiah',      ch:29, v:11 },
      { pattern: /i am the way.{0,10}truth.{0,10}life/i,       book:'John',          ch:14, v:6  },
      { pattern: /fear not.{0,20}i am with (thee|you)/i,       book:'Isaiah',        ch:41, v:10 },
      { pattern: /by grace.{0,20}saved through faith/i,        book:'Ephesians',     ch:2,  v:8  },
      { pattern: /trust in the lord.{0,30}heart/i,             book:'Proverbs',      ch:3,  v:5  },
      { pattern: /seek first the kingdom/i,                    book:'Matthew',       ch:6,  v:33 },
      { pattern: /faith is the substance/i,                    book:'Hebrews',       ch:11, v:1  },
      { pattern: /faith.{0,20}assurance.{0,20}hoped/i,         book:'Hebrews',       ch:11, v:1  },
      { pattern: /love is patient.{0,20}kind/i,                book:'1 Corinthians', ch:13, v:4  },
      { pattern: /no condemnation.{0,30}christ jesus/i,        book:'Romans',        ch:8,  v:1  },
      { pattern: /put on the.{0,15}armor of god/i,               book:'Ephesians',     ch:6,  v:11 },
      { pattern: /nothing.{0,20}shall.{0,20}separate us/i,     book:'Romans',        ch:8,  v:38 },
      { pattern: /be anxious for nothing/i,                    book:'Philippians',   ch:4,  v:6  },
      { pattern: /peace.{0,20}surpasses.{0,20}understanding/i, book:'Philippians',   ch:4,  v:7  },
      { pattern: /stand at the door.{0,20}knock/i,             book:'Revelation',    ch:3,  v:20 },
      { pattern: /god.{0,20}wipe away.{0,20}tears/i,            book:'Revelation',    ch:21, v:4  },
      { pattern: /valley of the shadow of death/i,             book:'Psalms',        ch:23, v:4  },
      { pattern: /be strong in the lord.{0,20}power/i,          book:'Ephesians',     ch:6,  v:10 },
      { pattern: /renewed.{0,20}renewing of.{0,10}mind/i,      book:'Romans',        ch:12, v:2  },
      { pattern: /do not conform.{0,30}pattern.{0,20}world/i,  book:'Romans',        ch:12, v:2  },
      { pattern: /your body is a temple/i,                     book:'1 Corinthians', ch:6,  v:19 },
      { pattern: /holy spirit.{0,30}live in you/i,             book:'Romans',        ch:8,  v:11 },
      { pattern: /wages of sin is death/i,                     book:'Romans',        ch:6,  v:23 },
      { pattern: /must be born again/i,                         book:'John',          ch:3,  v:3  },
      { pattern: /new creation.{0,30}old.{0,20}passed/i,       book:'2 Corinthians', ch:5,  v:17 },
      { pattern: /lamp.{0,20}feet.{0,20}light.{0,20}path/i,    book:'Psalms',        ch:119,v:105 },
      { pattern: /be still.{0,20}know.{0,20}i am god/i,        book:'Psalms',        ch:46, v:10 },
      { pattern: /weeping.{0,20}endure.{0,20}night.{0,20}joy/i,book:'Psalms',        ch:30, v:5  },
      { pattern: /come to me.{0,30}heavy laden/i,              book:'Matthew',       ch:11, v:28 },
      { pattern: /ask.{0,20}shall be given.{0,20}seek/i,       book:'Matthew',       ch:7,  v:7  },
      { pattern: /go.{0,20}make disciples.{0,20}all nations/i, book:'Matthew',       ch:28, v:19 },
      { pattern: /the great commission/i,                       book:'Matthew',       ch:28, v:19 },
      { pattern: /he is not here.{0,20}he is risen/i,           book:'Matthew',       ch:28, v:6  },
      { pattern: /every knee.{0,20}shall bow.{0,20}every tongue/i, book:'Philippians', ch:2,  v:10 },
      { pattern: /word of god.{0,30}living.{0,20}active/i,     book:'Hebrews',       ch:4,  v:12 },
      { pattern: /greater.{0,20}he that is in you/i,           book:'1 John',        ch:4,  v:4  },
      { pattern: /cast all your.{0,20}(care|anxiety)/i,        book:'1 Peter',       ch:5,  v:7  },
      { pattern: /if my people.{0,20}humble.{0,20}pray/i,      book:'2 Chronicles',  ch:7,  v:14 },
      { pattern: /why art thou cast down.{0,20}(my soul|o my)/i, book:'Psalms',      ch:42, v:5  },
      { pattern: /hope thou in god.{0,20}praise him/i,         book:'Psalms',        ch:42, v:5  },
      { pattern: /why.{0,10}disquieted.{0,15}(within me|in me)/i, book:'Psalms',     ch:42, v:5  },
      { pattern: /as the (hart|deer).{0,20}(panteth|pants).{0,20}water/i, book:'Psalms', ch:42, v:1 },
      { pattern: /my soul.{0,15}(thirsteth|thirsts).{0,15}god/i, book:'Psalms',      ch:42, v:2  },
      { pattern: /god is our refuge.{0,15}strength/i,          book:'Psalms',        ch:46, v:1  },
      { pattern: /a very present help in trouble/i,            book:'Psalms',        ch:46, v:1  },
      { pattern: /he is risen/i,                                book:'Matthew',       ch:28, v:6, confidence: 0.85 },
      { pattern: /angel answered.{0,20}(unto|to) the women/i,  book:'Matthew',       ch:28, v:5  },
      { pattern: /break this temple.{0,20}(three|3) days/i,    book:'John',          ch:2,  v:19 },
      { pattern: /destroy this temple.{0,20}(three|3) days/i,  book:'John',          ch:2,  v:19 },
      { pattern: /delivered the poor.{0,15}(that |who )?cri/i, book:'Job',           ch:29, v:12 },
      { pattern: /fatherless.{0,20}none to help/i,             book:'Job',           ch:29, v:12 },
      { pattern: /i.{0,10}sent.{0,10}(my |mine )?angel/i, book:'Revelation', ch:22, v:16, confidence: 0.82 },
      { pattern: /the joy of the lord.{0,15}(my |your )?strength/i, book:'Nehemiah', ch:8,  v:10 },
      { pattern: /create in me a clean heart/i,                book:'Psalms',        ch:51, v:10 },
      { pattern: /the earth is the lord.{0,5}s/i,              book:'Psalms',        ch:24, v:1  },
      { pattern: /the lord is my light.{0,15}salvation/i,      book:'Psalms',        ch:27, v:1  },
      { pattern: /make a joyful noise/i,                        book:'Psalms',        ch:100,v:1  },
      { pattern: /enter.{0,15}gates.{0,15}thanksgiving/i,     book:'Psalms',        ch:100,v:4  },
      { pattern: /his courts with praise/i,                     book:'Psalms',        ch:100,v:4  },
      { pattern: /love the lord.{0,15}(thy|your) god.{0,20}(all|whole).{0,10}heart/i, book:'Deuteronomy', ch:6, v:5 },
      { pattern: /the lord.{0,15}(my|our) shepherd/i,          book:'Psalms',        ch:23, v:1  },
      { pattern: /give thanks.{0,15}lord.{0,15}(he is |for he.{0,5})good/i, book:'Psalms', ch:136, v:1 },
      { pattern: /his mercy endure.{0,10}forever/i,            book:'Psalms',        ch:136,v:1  },
      { pattern: /bless the lord.{0,10}o my soul/i,            book:'Psalms',        ch:103,v:1  },
      { pattern: /forget not all his benefits/i,               book:'Psalms',        ch:103,v:2  },
      { pattern: /eye hath not seen.{0,15}ear.{0,15}heard/i,  book:'1 Corinthians', ch:2,  v:9  },
      { pattern: /the blood of jesus.{0,15}cleanse/i,          book:'1 John',        ch:1,  v:7  },
      { pattern: /if we confess our sins/i,                     book:'1 John',        ch:1,  v:9  },
      { pattern: /faithful and just to forgive/i,              book:'1 John',        ch:1,  v:9  },
      { pattern: /delight.{0,10}(thyself|yourself) .{0,15}lord/i, book:'Psalms',     ch:37, v:4  },
      { pattern: /commit thy way.{0,15}lord/i,                 book:'Psalms',        ch:37, v:5  },
      { pattern: /train up a child.{0,20}(way|should go)/i,    book:'Proverbs',      ch:22, v:6  },
      { pattern: /old.{0,10}(he|they) will not depart/i,      book:'Proverbs',      ch:22, v:6  },
      { pattern: /heaven and earth.{0,15}(shall |will )?pass away/i, book:'Matthew', ch:24, v:35 },
      { pattern: /my words.{0,15}(shall |will )?not pass away/i, book:'Matthew',     ch:24, v:35 },
    ];

    for (const sig of signatures) {
      if (sig.pattern.test(text)) {
        const verseText = window.BibleDB.getVerse(sig.book, sig.ch, sig.v);
        if (verseText) results.push({
          book: sig.book, chapter: sig.ch, verse: sig.v,
          ref: `${sig.book} ${sig.ch}:${sig.v}`,
          text: verseText, type: 'keyword', confidence: sig.confidence || 0.88,
        });
      }
    }
    return results;
  }

  // ── 4. OFFLINE CONTENT SEARCH ────────────────────────────────────────────
  // Full-text Bible search when Claude AI is unavailable.
  // Very strict thresholds — this method is inherently noisy because
  // common English words appear throughout the Bible.
  function detectContentSearch(text) {
    if (!window.BibleDB?.searchVerses) return [];
    const snippet = text.trim().slice(-140);
    if (snippet.split(/\s+/).length < 8) return [];
    const results = window.BibleDB.searchVerses(snippet, 'KJV', 2);
    return results
      .filter(r => r.score >= 18)
      .map(r => ({
        book: r.book, chapter: r.chapter, verse: r.verse, ref: r.ref,
        text: r.text, type: 'content', confidence: Math.min(0.56 + (r.score * 0.02), 0.75),
      }));
  }

  // ── 4b. TRIGGERED CONTENT SEARCH ─────────────────────────────────────────
  // When a preacher uses a Bible-introduction phrase ("the bible says",
  // "the scripture tells me", etc.) the text that follows is very likely a
  // Bible quote.  Run content search on the post-trigger text with a lower
  // threshold so verses are detected even without naming book/chapter/verse.
  const BIBLE_TRIGGERS = [
    /the bible says/i,
    /the bible recorded/i,
    /the bible tells? (?:me|us)/i,
    /the scripture says/i,
    /the scripture tells? (?:me|us)/i,
    /according to the scripture/i,
    /according to the bible/i,
    /according to/i,
    /open to the book/i,
    /the book of/i,
    /it is written/i,
    /as it is written/i,
    /the word says/i,
    /the word of god says/i,
    /jesus said/i,
    /jesus says/i,
    /jesus tells? (?:me|us|them|his)/i,
    /the psalmist (?:said|says|asked|wrote|declared)/i,
    /(?:paul|david|moses|peter|james|john|isaiah|jeremiah) (?:said|wrote|declared) in/i,
    /the verse says/i,
  ];

  function detectTriggeredContent(windowText) {
    if (!window.BibleDB?.searchVerses) return [];
    const lower = windowText.toLowerCase();
    let bestPostTrigger = '';
    for (const trigger of BIBLE_TRIGGERS) {
      const m = lower.match(trigger);
      if (m) {
        const after = windowText.slice(m.index + m[0].length).trim();
        if (after.length > bestPostTrigger.length) bestPostTrigger = after;
      }
    }
    if (!bestPostTrigger || bestPostTrigger.split(/\s+/).length < 5) return [];
    const snippet = bestPostTrigger.slice(0, 180);
    const results = window.BibleDB.searchVerses(snippet, 'KJV', 2);
    return results
      .filter(r => r.score >= 12)
      .map(r => ({
        book: r.book, chapter: r.chapter, verse: r.verse, ref: r.ref,
        text: r.text, type: 'content', confidence: Math.min(0.62 + (r.score * 0.02), 0.82),
      }));
  }

  // ── 5. CLAUDE AI DETECTION ────────────────────────────────────────────────
  async function detectWithClaude(windowText) {
    if (!apiKey) return [];
    const cacheKey = windowText.slice(-300);
    if (detectionCache.has(cacheKey)) return detectionCache.get(cacheKey);

    const prompt = `You are a Bible verse detector for a live church sermon app.

A preacher may quote, paraphrase, or verbally reference Bible verses. Detect any:
- Direct quotes (even partial)
- Paraphrases of well-known verses
- Verbal mentions like "turn to John 3" or "as Paul wrote in Romans 8"
- Thematic references that clearly point to a specific verse

TRANSCRIPT (last ~25 seconds of sermon):
"${windowText}"

Respond ONLY with a JSON array. Each detected verse:
{
  "ref": "Book Chapter:Verse",
  "confidence": 0.0-1.0,
  "type": "quote" | "paraphrase" | "mention"
}

Return [] if nothing detected. No markdown, no extra text.`;

    try {
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
        },
        body: JSON.stringify({
          model: 'claude-haiku-4-5-20251001',
          max_tokens: 300,
          system: 'You detect Bible verse references in sermon transcripts. Respond with valid JSON only.',
          messages: [{ role: 'user', content: prompt }],
        }),
      });

      if (!response.ok) { console.warn('[AIDetection] Claude error', response.status); return []; }

      const data  = await response.json();
      const text  = data.content?.[0]?.text || '[]';
      const clean = text.replace(/```json|```/g, '').trim();
      const items = JSON.parse(clean);

      const enriched = (Array.isArray(items) ? items : []).map(item => {
        if (!item?.ref) return null;
        const parsed = window.BibleDB?.parseReference?.(item.ref);
        if (!parsed) return null;
        const verseText = window.BibleDB.getVerse(parsed.book, parsed.chapter, parsed.verse);
        if (!verseText) return null;
        return {
          book: parsed.book, chapter: parsed.chapter, verse: parsed.verse,
          ref: `${parsed.book} ${parsed.chapter}:${parsed.verse}`,
          text: verseText, type: item.type || 'ai',
          confidence: Math.min(1, Math.max(0, parseFloat(item.confidence) || 0.75)),
        };
      }).filter(Boolean);

      detectionCache.set(cacheKey, enriched);
      if (detectionCache.size > 50) detectionCache.delete(detectionCache.keys().next().value);
      return enriched;

    } catch(e) {
      console.warn('[AIDetection] Claude detection error:', e.message);
      return [];
    }
  }

  // ── MAIN: processText ─────────────────────────────────────────────────────
  function processText(text, isOnline) {
    const cleanLine = String(text || '').trim();
    if (!cleanLine) return;

    recentLines.push(cleanLine);
    if (recentLines.length > WINDOW_SIZE) recentLines.shift();
    const windowText = recentLines.join(' ');

    const useAI = !!(apiKey && isOnline !== false && aiEnabled !== false);

    // Direct detection always runs — explicit "Book Ch:V" citations are reliable
    // Also check the join of last 2 lines so references split across lines
    // (e.g., "Romans" on line 1, "5:12" on line 2) are caught at the boundary.
    const directLine = detectDirect(cleanLine);
    const directAll = [...directLine];
    if (recentLines.length >= 2) {
      const prevLine = recentLines[recentLines.length - 2];
      const boundary = prevLine + ' ' + cleanLine;
      const directBoundary = detectDirect(boundary);
      const directPrev = detectDirect(prevLine);
      for (const bd of directBoundary) {
        // Skip if this book was already detected from either individual line —
        // prevents false refs when "Revelation" ends line 1 and "2-6" starts line 2
        if (directLine.some(d => d.book === bd.book)) continue;
        if (directPrev.some(d => d.book === bd.book)) continue;
        if (!directAll.some(d => d.book === bd.book && d.chapter === bd.chapter && d.verse === bd.verse)) {
          directAll.push(bd);
        }
      }
    }
    for (const d of directAll) {
      const verseText = window.BibleDB?.getVerse(d.book, d.chapter, d.verse);
      if (!verseText) continue;
      emit({ ...d, text: verseText, type: 'direct',
        confidence: d._chapterOnly ? 0.95 : 0.97 });
      lastBookChapter = { book: d.book, chapter: d.chapter, timestamp: Date.now() };
    }

    // ── Orphan verse detection ──────────────────────────────────────────────
    // When the preacher says "verse 43" or "verse 43 of 44" without naming
    // a book, connect it to the most recently detected book+chapter context.
    if (lastBookChapter && (Date.now() - lastBookChapter.timestamp) < LAST_CONTEXT_TTL) {
      const lower = cleanLine.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
      const spokenWord = '(?:one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)';
      const numToken = `(?:\\d+|${spokenWord}(?:\\s+${spokenWord})?)`;
      const orphanRe = new RegExp(`(?:verses?|v\\.?)\\s+(?:number\\s+)?(${numToken})(?:\\s+(?:of|and|to|through)\\s+(${numToken}))?`, 'gi');
      let om;
      while ((om = orphanRe.exec(lower)) !== null) {
        const v = wordsToNumber(om[1]?.trim());
        if (!v) continue;
        if (directAll.some(d => d.verse === v && d.book === lastBookChapter.book && d.chapter === lastBookChapter.chapter)) continue;
        const verseText = window.BibleDB?.getVerse(lastBookChapter.book, lastBookChapter.chapter, v);
        if (!verseText) continue;
        emit({
          book: lastBookChapter.book, chapter: lastBookChapter.chapter, verse: v,
          ref: `${lastBookChapter.book} ${lastBookChapter.chapter}:${v}`,
          text: verseText, type: 'direct', confidence: 0.93,
        });
      }
    }

    // ── "Next verse" / "previous verse" detection ────────────────────────
    // When the preacher says "next verse", "the following verse", "read on",
    // etc., advance from the last emitted verse to the next one.
    if (lastEmittedVerse && (Date.now() - lastEmittedVerse.timestamp) < LAST_CONTEXT_TTL) {
      const lower = cleanLine.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
      const nextPat = /\b(?:next verse|the next verse|following verse|the following verse|verse after|read on|let us read on|lets read on|continue reading|read the next|and the next verse|move to the next verse|go to the next verse|look at the next verse)\b/i;
      const prevPat = /\b(?:previous verse|the previous verse|verse before|go back a verse|back a verse|the verse before)\b/i;
      const nextMatch = nextPat.test(lower);
      const prevMatch = !nextMatch && prevPat.test(lower);
      if (nextMatch || prevMatch) {
        const delta = nextMatch ? 1 : -1;
        const newVerse = lastEmittedVerse.verse + delta;
        if (newVerse >= 1) {
          const verseText = window.BibleDB?.getVerse(lastEmittedVerse.book, lastEmittedVerse.chapter, newVerse);
          if (verseText) {
            emit({
              book: lastEmittedVerse.book, chapter: lastEmittedVerse.chapter, verse: newVerse,
              ref: `${lastEmittedVerse.book} ${lastEmittedVerse.chapter}:${newVerse}`,
              text: verseText, type: 'verbal', confidence: 0.93,
            });
            lastBookChapter = { book: lastEmittedVerse.book, chapter: lastEmittedVerse.chapter, timestamp: Date.now() };
          }
        }
      }
    }

    // Learned phrases always run — user-curated, low false-positive risk
    const windowHash = windowText.length + ':' + windowText.slice(-80);
    if (windowHash !== lastWindowHash) {
      lastWindowHash = windowHash;
      const learned = detectLearnedPhrases(windowText);
      for (const l of learned) emit(l);
    }

    // Triggered content search runs in both modes — when the preacher
    // uses phrases like "the bible says" followed by actual Bible text,
    // search the Bible for matching verses regardless of online/offline.
    clearTimeout(debounceTimer);
    const snapshot = windowText;

    if (useAI) {
      // ── ONLINE MODE: let Claude handle detection ──
      // Skip keyword/verbal — they produce false positives and Claude
      // does the same job with contextual understanding.
      debounceTimer = setTimeout(async () => {
        const triggered = detectTriggeredContent(snapshot);
        for (const t of triggered) emit(t);
        const aiResults = await detectWithClaude(snapshot);
        for (const a of aiResults) emit(a);
      }, 2000);
    } else {
      // ── OFFLINE MODE: fall back to rule-based methods ──
      const verbal = detectVerbal(windowText);
      for (const v of verbal) emit(v);

      const keyword = detectKeyword(cleanLine);
      for (const k of keyword) emit(k);

      debounceTimer = setTimeout(() => {
        const triggered = detectTriggeredContent(snapshot);
        for (const t of triggered) emit(t);
        const contentResults = detectContentSearch(snapshot);
        for (const c of contentResults) emit(c);
      }, 3000);
    }
  }

  function clearCache() {
    detectionCache.clear();
    recentEmits.clear();
    clearTimeout(debounceTimer);
    debounceTimer = null;
    recentLines   = [];
    lastWindowHash = '';
    lastBookChapter = null;
    lastEmittedVerse = null;
  }

  async function generateSermonNotes(transcript, title = '') {
    if (!apiKey) throw new Error('API key required');
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 2000,
        messages: [{ role: 'user', content:
          `Generate structured sermon notes from this transcript.\n${title ? `Title: ${title}\n` : ''}\nTRANSCRIPT:\n${transcript}\n\nReturn JSON: { title, topic, summary, mainPoints:[{heading,content,scriptures}], keyVerses:[{ref,text}], practicalApplications:[string], closingThought }\n\nJSON only, no markdown.`
        }],
      }),
    });
    const data = await response.json();
    const text = data.content?.[0]?.text || '{}';
    return JSON.parse(text.replace(/```json|```/g, '').trim());
  }

  return {
    init, setEnabled, processText, clearCache, generateSermonNotes,
    detectDirect, detectKeyword, detectVerbal, detectLearnedPhrases, reloadLearnedPhrases,
  };
})();

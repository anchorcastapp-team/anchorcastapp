'use strict';
// v8.0.0: the reusable scripts are data, never commands or HTML.
const crypto = require('node:crypto');
const INTRO = "Welcome, and thank you for joining us. Prepare your heart for a message from God's Word, bringing encouragement, understanding, and practical guidance for everyday life. As you listen, take a moment to reflect, and allow the truth of Scripture to strengthen your faith and draw you closer to God.";
const OUTRO = "Thank you for listening. We pray this message has encouraged your heart and strengthened your faith. Take time to reflect on what you have heard and put God's Word into practice. Join us again for another message, and share this encouragement with someone who needs it. May God's grace and peace be with you.";
// These legacy strings identify untouched starter text during a one-time upgrade.
// Custom scripts and imported narration are never replaced automatically.
const LEGACY_INTRO = 'God’s plan for you is that you are sustained by His word. That is why, at The Redeemed Christian Church of God, Lighthouse Fremont, we are committed to preaching and teaching the word of God in a way that is clear, easy to understand, and speaks to your needs. As the word comes through His anointed vessel, let your heart open to the new thing it brings.';
const LEGACY_OUTRO = 'We hope you met God through this message. He is not done with you yet. If you would like to know more about Him, or you have questions from today’s message or anything else, please call us on [phone number]. You can also email us at [email address], or follow us on [platform] @[handle].';
const LEGACY_STYLE = 'Use a formal, warm male narration style with a calm lower register, clear diction and measured phrasing. Speak naturally and respectfully for a church sermon introduction or closing. Do not imitate any real person. Read the supplied text exactly, including its AI narration disclosure. Do not add music, sound effects, words or commentary. Read phone numbers slowly in distinct digits and pronounce email addresses clearly.';
const SPEECH_FORMAT = 'script-only-v1';
const SAMPLES_REVISION = 2;
const VOICES = ['onyx', 'cedar', 'ash', 'echo', 'fable', 'alloy', 'marin'];
const STYLE = 'Use a formal, warm male narration style with a calm lower register, clear diction and measured phrasing. Speak naturally and respectfully for a church sermon introduction or closing. Do not imitate any real person. Read the supplied text exactly, without any added preface, disclosure announcement or sign-off. Do not add music, sound effects, words or commentary. Read phone numbers slowly in distinct digits and pronounce email addresses clearly.';
function defaults() {
  return { introText: INTRO, outroText: OUTRO, phone: '', email: '', platform: '', handle: '',
    voice: 'onyx', speed: 0.92, instructions: STYLE, music: 'warm', musicVolume: 18,
    leadSeconds: 1, tailSeconds: 1.5, gapSeconds: 0.4 };
}
function string(value, label, max) {
  if (typeof value !== 'string' || value.length > max || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(value)) throw Error(`Invalid ${label}. Maximum ${max} characters.`);
  return value.trim();
}
function number(value, label, min, max) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) throw Error(`${label} must be between ${min} and ${max}.`);
  return value;
}
function validate(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw Error('Recording settings are invalid.');
  const c = {};
  for (const [key, max] of Object.entries({ introText: 3500, outroText: 3500, phone: 100, email: 200, platform: 100, handle: 100, instructions: 1800 })) c[key] = string(input[key] ?? defaults()[key], key, max);
  c.handle = c.handle.replace(/^@+/, '');
  c.voice = input.voice ?? 'onyx'; if (!VOICES.includes(c.voice)) throw Error('Select a supported narration voice.');
  c.speed = number(input.speed ?? .92, 'Voice speed', .75, 1.2);
  c.music = input.music ?? 'warm'; if (!['none', 'warm', 'custom'].includes(c.music)) throw Error('Select a background music option.');
  c.musicVolume = number(input.musicVolume ?? 18, 'Music level', 0, 35);
  c.leadSeconds = number(input.leadSeconds ?? 1, 'Music lead-in', 0, 4);
  c.tailSeconds = number(input.tailSeconds ?? 1.5, 'Music ending', 0, 5);
  c.gapSeconds = number(input.gapSeconds ?? .4, 'Transition gap', 0, 2);
  return c;
}
function migrateLegacyDefaults(input, clips = {}) {
  const config = validate(input), changedKeys = [];
  // Formatting-only differences are accepted; edited wording stays untouched.
  const equivalent = (a, b) => a.replace(/[\u2018\u2019]/g, "'").replace(/\s+/g, ' ').trim() === b.replace(/[\u2018\u2019]/g, "'").replace(/\s+/g, ' ').trim();
  for (const [slot, oldText, newText] of [['intro', LEGACY_INTRO, INTRO], ['outro', LEGACY_OUTRO, OUTRO]]) {
    const key = slot + 'Text';
    if (clips?.[slot]?.source !== 'imported' && equivalent(config[key], oldText)) {
      config[key] = newText; changedKeys.push(key);
    }
  }
  if (equivalent(config.instructions, LEGACY_STYLE)) { config.instructions = STYLE; changedKeys.push('instructions'); }
  return { config, changedKeys };
}
function slotName(slot) { if (!['intro', 'outro'].includes(slot)) throw Error('Choose introduction or closing.'); return slot; }
function resolvedText(config, slot, { allowEmpty = false } = {}) {
  slotName(slot);
  let text = config[slot + 'Text'];
  const fields = { 'phone number': config.phone, 'email address': config.email, platform: config.platform, handle: config.handle };
  text = text.replace(/\[([^\]]+)\]/g, (match, name) => fields[name.toLowerCase().trim()] || match);
  const unresolved = text.match(/\[[^\]]*\]/g);
  if (unresolved) throw Error('Fill in or remove these placeholders before preparing audio: ' + [...new Set(unresolved)].join(', '));
  if (!allowEmpty && !text.trim()) throw Error('Enter the ' + (slot === 'intro' ? 'introduction' : 'closing') + ' text first.');
  return text.trim();
}
const digest = value => crypto.createHash('sha256').update(JSON.stringify(value)).digest('hex');
function textSignature(c, slot, imported = false) {
  // Imported narration is tied to text/contact edits, but not to a synthetic voice choice.
  const text = resolvedText(c, slot, { allowEmpty: imported });
  return digest(imported ? { text } : { text, voice: c.voice, speed: c.speed, instructions: c.instructions, model: 'gpt-4o-mini-tts', speechFormat: SPEECH_FORMAT });
}
function mixSignature(c, voice, customMusic) {
  return digest({ voice: voice?.file, music: c.music, customMusic: c.music === 'custom' ? customMusic?.file : null,
    volume: c.musicVolume, lead: c.leadSeconds, tail: c.tailSeconds });
}
function speechPayload(c, slot) {
  // The spoken content is exactly the resolved editor text. Listener disclosure
  // belongs in the publishing description/player notice, not an injected sentence.
  const input = resolvedText(c, slot);
  if (input.length > 4096) throw Error('The completed script including contacts is too long. Shorten it before generating speech.');
  return { model: 'gpt-4o-mini-tts', voice: c.voice, input,
    instructions: c.instructions || STYLE, speed: c.speed, response_format: 'mp3' };
}
module.exports = { INTRO, OUTRO, STYLE, VOICES, SPEECH_FORMAT, SAMPLES_REVISION, migrateLegacyDefaults, defaults, validate, slotName, resolvedText, digest, textSignature, mixSignature, speechPayload };

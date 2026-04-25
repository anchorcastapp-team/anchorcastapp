/**
 * AnchorCast — Self-Improving Transcript Memory
 * Post-Whisper adaptive correction layer.
 *
 * Architecture:
 *   Raw Whisper text
 *     → normalize()          — lowercase, spacing
 *     → applyVocabAliases()  — church names, pastor names, ministry terms
 *     → applyCorrectionRules() — phrase substitutions (bible books, local corrections)
 *     → output corrected text + metadata
 *
 * All data stored as JSON via IPC to main process.
 * No model retraining. No runtime training. Pure post-processing.
 */

'use strict';

const TranscriptMemory = (function () {

  // ── Built-in Bible normalization rules (always active, no learning needed) ──
  const BUILTIN_BIBLE_RULES = [
    // Book name corrections
    { src: /\brevelations\b/gi,                   tgt: 'Revelation' },
    { src: /\bgeneses\b/gi,                        tgt: 'Genesis' },
    { src: /\bpsalms\b/gi,                         tgt: 'Psalms' },
    { src: /\bpsalm\b/gi,                          tgt: 'Psalm' },
    // Ordinal prefixes: spoken → numeric
    { src: /\bfirst\s+(corinthians|kings|samuel|chronicles|thessalonians|timothy|peter|john)\b/gi,
      tgt: (_, b) => `1 ${b.charAt(0).toUpperCase()}${b.slice(1)}` },
    { src: /\bsecond\s+(corinthians|kings|samuel|chronicles|thessalonians|timothy|peter|john)\b/gi,
      tgt: (_, b) => `2 ${b.charAt(0).toUpperCase()}${b.slice(1)}` },
    { src: /\bthird\s+(john)\b/gi,
      tgt: (_, b) => `3 ${b.charAt(0).toUpperCase()}${b.slice(1)}` },
    // Verbal verse references: "chapter three verse sixteen" → "3:16"
    { src: /\bchapter\s+(\w+)\s+verse\s+(\w+)\b/gi,
      tgt: (_, ch, v) => `${_wordToNum(ch)}:${_wordToNum(v)}` },
    // "psalm twenty three" → "Psalm 23"
    { src: /\bpsalm\s+(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred|twenty.one|twenty.two|twenty.three|twenty.four|twenty.five|twenty.six|twenty.seven|twenty.eight|twenty.nine|thirty.one|thirty.two|thirty.three|one.hundred|one.hundred.and.nineteen)\b/gi,
      tgt: (_, n) => `Psalm ${_wordToNum(n)}` },
  ];

  // Word → number map for verbal references
  const WORD_NUMS = {
    one:1,two:2,three:3,four:4,five:5,six:6,seven:7,eight:8,nine:9,ten:10,
    eleven:11,twelve:12,thirteen:13,fourteen:14,fifteen:15,sixteen:16,
    seventeen:17,eighteen:18,nineteen:19,twenty:20,thirty:30,forty:40,
    fifty:50,sixty:60,seventy:70,eighty:80,ninety:90,hundred:100,
  };

  function _wordToNum(word) {
    const w = String(word || '').toLowerCase().trim();
    if (/^\d+$/.test(w)) return parseInt(w, 10);
    // handle "twenty three" → 23, "one hundred and nineteen" → 119
    const parts = w.replace(/\s+and\s+/g, ' ').replace(/[-]/g,' ').split(/\s+/);
    let total = 0, current = 0;
    for (const p of parts) {
      const n = WORD_NUMS[p];
      if (!n) continue;
      if (n === 100) { current = (current || 1) * 100; }
      else if (n >= 20) { current += n; }
      else { current += n; }
    }
    return (total + current) || word;
  }

  // ── State ─────────────────────────────────────────────────────────────────
  let _rules        = [];   // correction_rules from storage
  let _vocab        = [];   // custom_vocabulary from storage
  let _profiles     = [];   // speaker_profiles from storage
  let _activeProfile = null; // currently selected speaker profile id
  let _sessionId    = null;  // current transcript session id
  let _chunkIndex   = 0;
  let _settings     = { enabled: true, learningEnabled: true };
  let _pendingEvents = []; // correction events not yet flushed

  // ── Public API ────────────────────────────────────────────────────────────

  /**
   * Process a raw transcript chunk.
   * Returns { raw, corrected, appliedRules, sessionId, chunkId }
   */
  function process(rawText) {
    if (!rawText?.trim()) return { raw: rawText, corrected: rawText };
    if (!_settings.enabled) return { raw: rawText, corrected: rawText };

    const raw = rawText.trim();

    // Step 1 — Normalize
    let text = _normalize(raw);

    // Step 2 — Apply vocabulary aliases (church names, people, ministries)
    const { text: afterVocab } = _applyVocabAliases(text);
    text = afterVocab;

    // Step 3 — Apply built-in Bible normalization (always on)
    text = _applyBuiltinBibleRules(text);

    // Step 4 — Apply learned correction rules (sorted by priority)
    const { text: afterRules, appliedRules } = _applyCorrectionRules(text);
    text = afterRules;

    // Step 5 — Title-case proper nouns and restore sentence structure
    const corrected = _restoreCase(raw, text);

    // Record chunk for later learning
    const chunkId = _recordChunk(raw, corrected, appliedRules);

    return { raw, corrected, appliedRules, chunkId, sessionId: _sessionId };
  }

  /** Start a new transcript session. Call when recording starts. */
  function startSession(speakerProfileId) {
    _sessionId = `session_${Date.now()}_${Math.random().toString(36).slice(2,6)}`;
    _chunkIndex = 0;
    _activeProfile = speakerProfileId || _profiles.find(p => p.isDefault)?.id || null;
    _persistSession({ id: _sessionId, speakerProfileId: _activeProfile, startedAt: new Date().toISOString() });
    return _sessionId;
  }

  /** End the current session. */
  function endSession() {
    if (_sessionId) {
      _persistSessionEnd(_sessionId);
      _flushPendingEvents();
    }
    _sessionId = null;
    _chunkIndex = 0;
  }

  /** User manually corrected a transcript line. Record for learning. */
  function recordUserCorrection(chunkId, beforeText, afterText) {
    if (!chunkId || !beforeText || !afterText || beforeText === afterText) return;
    const event = {
      id: `evt_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      chunkId, beforeText, afterText,
      wasUserEdited: true,
      accepted: 1,
      createdAt: new Date().toISOString(),
    };
    _pendingEvents.push(event);
    _flushPendingEvents();
    // Check if this correction should be auto-promoted to a rule suggestion
    _checkForRulePromotion(beforeText, afterText);
  }

  /** User accepted/rejected an auto-correction. */
  function recordCorrectionFeedback(chunkId, ruleId, accepted) {
    const event = {
      id: `evt_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      chunkId, ruleId,
      wasUserEdited: false,
      accepted: accepted ? 1 : 0,
      createdAt: new Date().toISOString(),
    };
    _pendingEvents.push(event);
    _flushPendingEvents();
    // Update rule hit/approve/reject counts in memory
    const rule = _rules.find(r => r.id === ruleId);
    if (rule) {
      rule.hitCount = (rule.hitCount || 0) + 1;
      if (accepted) rule.approvedCount = (rule.approvedCount || 0) + 1;
      else rule.rejectedCount = (rule.rejectedCount || 0) + 1;
      // Deactivate if rejection rate too high
      if ((rule.rejectedCount || 0) >= 3 && (rule.approvedCount || 0) < (rule.rejectedCount || 0)) {
        rule.isActive = 0;
      }
      _persistRules();
    }
  }

  /** Set active speaker profile. */
  function setProfile(profileId) {
    _activeProfile = profileId;
  }

  /** Get all speaker profiles. */
  function getProfiles() { return _profiles; }

  /** Get all custom vocabulary. */
  function getVocabulary() { return _vocab; }

  /** Get all correction rules. */
  function getRules() { return _rules; }

  /** Get pending learning suggestions (rules with enough evidence but not yet approved). */
  function getLearningQueue() {
    return (window._tmLearningQueue || []);
  }

  /** Approve a suggested rule from the learning queue. */
  function approveSuggestion(suggestionId) {
    const queue = window._tmLearningQueue || [];
    const s = queue.find(x => x.id === suggestionId);
    if (!s) return;
    const newRule = {
      id: `rule_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      scope: s.scope || 'global',
      speakerProfileId: s.speakerProfileId || null,
      ruleType: s.ruleType || 'phrase',
      sourceText: s.sourceText,
      targetText: s.targetText,
      priority: s.priority || 100,
      confidence: 1.0,
      hitCount: s.hitCount || 0,
      approvedCount: (s.approvedCount || 0) + 1,
      rejectedCount: 0,
      isActive: 1,
      createdBy: 'user_approved',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    _rules.push(newRule);
    _sortRules();
    _persistRules();
    window._tmLearningQueue = queue.filter(x => x.id !== suggestionId);
    _persistLearningQueue();
  }

  /** Reject a suggested rule. */
  function rejectSuggestion(suggestionId) {
    const queue = window._tmLearningQueue || [];
    window._tmLearningQueue = queue.filter(x => x.id !== suggestionId);
    _persistLearningQueue();
  }

  /** Add a custom vocabulary term. */
  function addVocabTerm(term) {
    const existing = _vocab.find(v =>
      v.canonicalTerm.toLowerCase() === (term.canonicalTerm || '').toLowerCase()
    );
    if (existing) {
      // Merge aliases
      const existingAliases = existing.aliases || [];
      const newAliases = term.aliases || [];
      existing.aliases = [...new Set([...existingAliases, ...newAliases])];
      existing.updatedAt = new Date().toISOString();
    } else {
      _vocab.push({
        id: `vocab_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
        scope: term.scope || 'global',
        speakerProfileId: term.speakerProfileId || null,
        canonicalTerm: term.canonicalTerm,
        aliases: term.aliases || [],
        category: term.category || 'general',
        weight: 1.0,
        usageCount: 0,
        isActive: 1,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
    }
    _persistVocab();
  }

  /** Remove a vocabulary term. */
  function removeVocabTerm(id) {
    _vocab = _vocab.filter(v => v.id !== id);
    _persistVocab();
  }

  /** Add a correction rule directly. */
  function addRule(rule) {
    const existing = _rules.find(r =>
      r.sourceText?.toLowerCase() === (rule.sourceText || '').toLowerCase() &&
      r.scope === (rule.scope || 'global')
    );
    if (existing) {
      existing.targetText = rule.targetText;
      existing.updatedAt = new Date().toISOString();
      existing.isActive = 1;
    } else {
      _rules.push({
        id: `rule_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
        scope: rule.scope || 'global',
        speakerProfileId: rule.speakerProfileId || null,
        ruleType: rule.ruleType || 'phrase',
        sourceText: rule.sourceText,
        targetText: rule.targetText,
        priority: rule.priority || 100,
        confidence: 1.0,
        hitCount: 0,
        approvedCount: 0,
        rejectedCount: 0,
        isActive: 1,
        createdBy: rule.createdBy || 'user',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
    }
    _sortRules();
    _persistRules();
  }

  /** Remove a rule. */
  function removeRule(id) {
    _rules = _rules.filter(r => r.id !== id);
    _persistRules();
  }

  /** Add or update a speaker profile. */
  function saveProfile(profile) {
    const existing = _profiles.find(p => p.id === profile.id);
    if (existing) {
      Object.assign(existing, profile, { updatedAt: new Date().toISOString() });
    } else {
      _profiles.push({
        id: profile.id || `prof_${Date.now()}`,
        name: profile.name || 'Unnamed',
        description: profile.description || '',
        isDefault: profile.isDefault || 0,
        isActive: 1,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
    }
    // Only one default
    if (profile.isDefault) {
      _profiles.forEach(p => { if (p.id !== profile.id) p.isDefault = 0; });
    }
    _persistProfiles();
  }

  /** Delete a speaker profile. */
  function deleteProfile(id) {
    _profiles = _profiles.filter(p => p.id !== id);
    if (_activeProfile === id) _activeProfile = null;
    _persistProfiles();
  }

  /** Run the offline learning job — mines corrections for rule suggestions. */
  function runLearningJob() {
    if (!window.electronAPI?.loadLearningData) return;
    window.electronAPI.loadLearningData().then(data => {
      if (!data) return;
      const events = data.correctionEvents || [];
      const grouped = {};
      for (const ev of events) {
        if (!ev.wasUserEdited || !ev.beforeText || !ev.afterText) continue;
        const key = `${ev.beforeText.toLowerCase()}|||${ev.afterText}`;
        if (!grouped[key]) grouped[key] = { before: ev.beforeText, after: ev.afterText, events: [] };
        grouped[key].events.push(ev);
      }
      const queue = window._tmLearningQueue || [];
      const existingKeys = new Set(queue.map(s => `${s.sourceText.toLowerCase()}|||${s.targetText}`));
      const existingRuleKeys = new Set(_rules.map(r => `${(r.sourceText||'').toLowerCase()}|||${r.targetText}`));

      for (const [key, info] of Object.entries(grouped)) {
        const total = info.events.length;
        const accepted = info.events.filter(e => e.accepted !== 0).length;
        const rate = total > 0 ? accepted / total : 0;
        // Threshold: seen ≥3 times, accepted ≥70%
        if (total >= 3 && rate >= 0.7 && !existingKeys.has(key) && !existingRuleKeys.has(key)) {
          queue.push({
            id: `sug_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
            sourceText: info.before,
            targetText: info.after,
            ruleType: 'phrase',
            scope: 'global',
            hitCount: total,
            approvedCount: accepted,
            evidence: `Seen ${total}× — accepted ${Math.round(rate*100)}%`,
            createdAt: new Date().toISOString(),
          });
        }
      }
      window._tmLearningQueue = queue;
      _persistLearningQueue();
    }).catch(() => {});
  }

  /** Load all stored data. Call once on app init. */
  async function init() {
    if (!window.electronAPI?.loadAdaptiveMemory) return;
    try {
      const data = await window.electronAPI.loadAdaptiveMemory();
      _rules    = (data?.rules    || []).filter(r => r.isActive !== 0);
      _vocab    = (data?.vocab    || []).filter(v => v.isActive !== 0);
      _profiles = data?.profiles  || [];
      _settings = { ...{ enabled: true, learningEnabled: true }, ...(data?.settings || {}) };
      window._tmLearningQueue = data?.learningQueue || [];
      _sortRules();
      // Ensure a default profile exists
      if (!_profiles.length) {
        _profiles = [
          { id:'default', name:'Default', description:'General speaker', isDefault:1, isActive:1,
            createdAt:new Date().toISOString(), updatedAt:new Date().toISOString() },
          { id:'pastor',  name:'Pastor',  description:'Main pastor',     isDefault:0, isActive:1,
            createdAt:new Date().toISOString(), updatedAt:new Date().toISOString() },
          { id:'choir',   name:'Choir / Worship Leader', description:'Worship team', isDefault:0, isActive:1,
            createdAt:new Date().toISOString(), updatedAt:new Date().toISOString() },
        ];
        _persistProfiles();
      }
    } catch (e) {
      console.warn('[TranscriptMemory] init failed', e);
    }
  }

  // ── Internal helpers ───────────────────────────────────────────────────────

  function _normalize(text) {
    return text
      .replace(/\s+/g, ' ')
      .trim();
  }

  function _applyBuiltinBibleRules(text) {
    let out = text;
    for (const rule of BUILTIN_BIBLE_RULES) {
      out = out.replace(rule.src, rule.tgt);
    }
    return out;
  }

  function _applyVocabAliases(text) {
    let out = text;
    // Filter by active profile or global
    const relevant = _vocab.filter(v =>
      v.isActive !== 0 &&
      (v.scope === 'global' || !v.speakerProfileId || v.speakerProfileId === _activeProfile)
    );
    for (const v of relevant) {
      const aliases = v.aliases || [];
      for (const alias of aliases) {
        if (!alias?.trim()) continue;
        const escaped = alias.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const re = new RegExp(`\\b${escaped}\\b`, 'gi');
        if (re.test(out)) {
          out = out.replace(re, v.canonicalTerm);
          v.usageCount = (v.usageCount || 0) + 1;
        }
      }
    }
    return { text: out };
  }

  function _applyCorrectionRules(text) {
    let out = text;
    const appliedRules = [];
    // Filter rules: global + active profile rules, sorted by priority
    const relevant = _rules.filter(r =>
      r.isActive !== 0 &&
      (r.scope === 'global' || !r.speakerProfileId || r.speakerProfileId === _activeProfile)
    );
    for (const rule of relevant) {
      if (!rule.sourceText || !rule.targetText) continue;
      let re;
      try {
        if (rule.ruleType === 'regex') {
          re = new RegExp(rule.sourceText, 'gi');
        } else {
          const escaped = rule.sourceText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          re = new RegExp(`\\b${escaped}\\b`, 'gi');
        }
      } catch { continue; }
      if (re.test(out)) {
        out = out.replace(re, rule.targetText);
        rule.hitCount = (rule.hitCount || 0) + 1;
        appliedRules.push(rule.id);
      }
    }
    return { text: out, appliedRules };
  }

  function _restoreCase(original, corrected) {
    // Restore sentence-start capitals
    return corrected.replace(/(^\s*\w|[.!?]\s+\w)/g, c => c.toUpperCase());
  }

  function _sortRules() {
    _rules.sort((a, b) => (a.priority || 100) - (b.priority || 100));
  }

  function _recordChunk(raw, corrected, appliedRules) {
    const id = `chunk_${Date.now()}_${_chunkIndex++}`;
    if (!_sessionId) return id;
    const chunk = {
      id, sessionId: _sessionId,
      speakerProfileId: _activeProfile,
      chunkIndex: _chunkIndex,
      rawText: raw,
      correctedText: corrected,
      appliedRules,
      createdAt: new Date().toISOString(),
    };
    // Non-blocking persist
    window.electronAPI?.persistTranscriptChunk?.(chunk);
    return id;
  }

  function _checkForRulePromotion(before, after) {
    // Add to a pending suggestions buffer; full promotion happens in runLearningJob
    if (!window._tmCorrectionBuffer) window._tmCorrectionBuffer = {};
    const key = `${before.toLowerCase()}|||${after}`;
    window._tmCorrectionBuffer[key] = (window._tmCorrectionBuffer[key] || 0) + 1;
  }

  function _persistSession(session) {
    window.electronAPI?.persistTranscriptSession?.(session);
  }

  function _persistSessionEnd(sessionId) {
    window.electronAPI?.persistTranscriptSessionEnd?.({ sessionId, endedAt: new Date().toISOString() });
  }

  function _flushPendingEvents() {
    if (!_pendingEvents.length) return;
    window.electronAPI?.persistCorrectionEvents?.([..._pendingEvents]);
    _pendingEvents = [];
  }

  function _persistRules() {
    window.electronAPI?.saveAdaptiveMemory?.({ type: 'rules', data: _rules });
  }

  function _persistVocab() {
    window.electronAPI?.saveAdaptiveMemory?.({ type: 'vocab', data: _vocab });
  }

  function _persistProfiles() {
    window.electronAPI?.saveAdaptiveMemory?.({ type: 'profiles', data: _profiles });
  }

  function _persistLearningQueue() {
    window.electronAPI?.saveAdaptiveMemory?.({ type: 'learningQueue', data: window._tmLearningQueue || [] });
  }

  // ── Public interface ───────────────────────────────────────────────────────
  return {
    init,
    process,
    startSession,
    endSession,
    setProfile,
    getProfiles,
    getVocabulary,
    getRules,
    getLearningQueue,
    recordUserCorrection,
    recordCorrectionFeedback,
    approveSuggestion,
    rejectSuggestion,
    addVocabTerm,
    removeVocabTerm,
    addRule,
    removeRule,
    saveProfile,
    deleteProfile,
    runLearningJob,
    get activeProfile() { return _activeProfile; },
    get enabled() { return _settings.enabled; },
    set enabled(v) { _settings.enabled = !!v; },
  };

})();

// Make globally available
window.TranscriptMemory = TranscriptMemory;

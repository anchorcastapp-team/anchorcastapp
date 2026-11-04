/**
 * Detection discussion/passage model (1.8.1 activity-order-v1).
 * Occurrence records are never discarded. This model is shared by the panel,
 * Auto scheduler and exports, not a visual hide-duplicates filter.
 * Explicit re-announcements and intervening confirmed passages start new groups.
 * Unconfirmed thematic suggestions cannot move the active reading or boost its score.
 */
(function (root) {
  'use strict';
  const CLOSED = new Set(['rejected', 'dismissed', 'superseded']);
  const RANK = {reference: 3, quotation: 2, learned: 1};
  const CONTINUITY_MS = 120000; // fallback boundary, not a service-wide verse ban
  const key = d => d.eventKey || d.id;
  const time = d => Number(d.sourceTime || d.createdAt || 0);
  // Presentation recency is deliberately separate from spoken passage order.
  // Rebuilding a view or changing queue state must never make an old match new.
  const finiteTime = value => typeof value === 'number' && Number.isFinite(value) && value >= 0 ? value : 0;
  const activityTime = d => finiteTime(d.updatedAt) || finiteTime(d.createdAt) || finiteTime(d.sourceTime);
  const activitySequence = d => Number.isSafeInteger(d.activitySequence) && d.activitySequence > 0 ? d.activitySequence : 0;
  const EVIDENCE_FIELDS = ['ref','book','chapter','verse','confidence','confirmation','autoEligible',
    'holdReason','sourceText','sourceLineId','sourceStartLineId','sourceStart','sourceEnd','sourceTime',
    'firstSeq','lastSeq','text','translation','fallbackTranslation','type','evidenceTypes',
    'passageId','passageVerses','rangeStart','rangeEnd','announcedPassage'];
  function evidenceChanged(before, after) {
    return !before || EVIDENCE_FIELDS.some(field => JSON.stringify(before[field]) !== JSON.stringify(after[field]));
  }
  function nextActivity(records, now = Date.now()) {
    let updatedAt = finiteTime(now), sequence = 0;
    for (const d of records || []) {
      if (!d) continue;
      updatedAt = Math.max(updatedAt, activityTime(d));
      sequence = Math.max(sequence, activitySequence(d));
    }
    // The sequence makes simultaneous deliveries stable; the logical clock also
    // prevents a system clock correction from burying the latest accepted update.
    return {updatedAt, activitySequence: sequence < Number.MAX_SAFE_INTEGER ? sequence + 1 : sequence};
  }
  function recentFirst(a, b) {
    return b.latestUpdatedAt - a.latestUpdatedAt || b.latestActivitySequence - a.latestActivitySequence ||
      b.lastTime - a.lastTime || b.sourceStart - a.sourceStart ||
      (a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
  }
  function position(d) {
    if (Number.isFinite(d.sourceStart)) return d.sourceStart;
    const eventKey=String(d.eventKey || '');
    const tail = eventKey.startsWith('ref:') ? eventKey.match(/:(\d+)$/) : eventKey.match(/:(\d+):\d+$/);
    return tail ? Number(tail[1]) : time(d);
  }
  const confirmed = d => !CLOSED.has(d.status) && !!RANK[d.confirmation] && d.autoEligible !== false;
  function order(a, b) { return position(a) - position(b) || time(a) - time(b) || Number(a.verse) - Number(b.verse); }
  function better(a, b) {
    if (!b) return true;
    return (RANK[a.confirmation] || 0) > (RANK[b.confirmation] || 0) ||
      ((RANK[a.confirmation] || 0) === (RANK[b.confirmation] || 0) && a.confidence > b.confidence);
  }
  function samePassage(g, d) {
    return g.book === d.book && g.chapter === Number(d.chapter) && g.verses.includes(Number(d.verse));
  }
  function passageLabel(book, chapter, verses) {
    const runs = []; let start, end;
    for (const v of [...new Set(verses)].sort((a,b)=>a-b)) {
      if (start === undefined) start = end = v;
      else if (v === end + 1) end = v;
      else { runs.push(start === end ? String(start) : `${start}–${end}`); start = end = v; }
    }
    if (start !== undefined) runs.push(start === end ? String(start) : `${start}–${end}`);
    return `${book} ${chapter}:${runs.join(', ')}`;
  }
  function announcement(d) {
    if (d.type !== 'direct' || d.confirmation !== 'reference') return null;
    if(d.rangeStart && d.rangeStart===d.rangeEnd)return null;
    let verses = Array.isArray(d.passageVerses) ? d.passageVerses : null;
    // Compatibility for existing 1.7.2 range records. Never infer intermediate
    // verses from MIN/MAX of a non-contiguous list.
    if (!verses && d.sourceText && root.ScriptureReferences && root.BibleDB) {
      const ev = root.ScriptureReferences.create(root.BibleDB).scan(d.sourceText)
        .find(e => !e.incomplete && e.book === d.book && e.chapter === Number(d.chapter) && e.verses?.includes(Number(d.verse)));
      if (ev?.verses?.length > 1) verses = ev.verses;
    }
    if (!verses?.length) return null;
    verses = [...new Set(verses.map(Number).filter(v=>Number.isInteger(v)&&v>0&&v<=176))].sort((a,b)=>a-b);
    if (verses.length < 2 || !verses.includes(Number(d.verse))) return null;
    return {id: d.passageId || `passage:${d.book}:${d.chapter}:${position(d)}`, verses};
  }
  function build(records = [], settings = {}) {
    const threshold = root.DetectionPolicy?.threshold(settings) ?? .75;
    const all = [], byEvent = new Map(), byAnnouncement = new Map();
    let active = null;
    // Threshold changes affect the active view only, never erase History.
    const eligibleRecords = records.filter(d=>d && key(d) && Number.isFinite(d.confidence) && d.confidence >= threshold && !CLOSED.has(d.status));
    for (const d of eligibleRecords.slice().sort(order)) {
      const ann = announcement(d), isConfirmed = confirmed(d);
      let group = ann ? byAnnouncement.get(ann.id) : null;
      const startsNewCitation = d.type === 'direct' && d.confirmation === 'reference';
      // A separately spoken direct reference within an announced range advances
      // that reading. Re-announcing the same single verse remains a new occurrence.
      const rangeStep = active?.kind === 'passage' && !ann && samePassage(active,d) && isConfirmed;
      if (!group && active && samePassage(active,d) && (!startsNewCitation || rangeStep) &&
          time(d) >= active.firstTime && time(d) - active.lastTime <= CONTINUITY_MS) group = active;
      if (!group && !isConfirmed) {
        // Related matches aggregate with their active passage, but cannot revive
        // a previous passage after a confirmed topic change.
        group = [...all].reverse().find(g=>g.section==='suggestions' && samePassage(g,d) && time(d)-g.lastTime<=CONTINUITY_MS);
      }
      if (!group) {
        group = {id: ann ? ann.id : `discussion:${key(d)}`, book:d.book, chapter:Number(d.chapter),
          verses:ann ? ann.verses : [Number(d.verse)], kind:ann?'passage':'verse',
          section:isConfirmed?'confirmed':'suggestions', firstTime:time(d), lastTime:time(d),
          sourceStart:position(d), members:[], representative:null, confirmedRepresentative:null};
        all.push(group);
      }
      if (ann) byAnnouncement.set(ann.id,group);
      group.members.push(d); group.firstTime = Math.min(group.firstTime,time(d)); group.lastTime = Math.max(group.lastTime,time(d));
      if (better(d,group.representative)) group.representative=d;
      if (isConfirmed) {
        group.section='confirmed';
        if (better(d,group.confirmedRepresentative)) group.confirmedRepresentative=d;
        active=group;
      }
      byEvent.set(key(d),group);
    }
    for (const g of all) {
      g.label=passageLabel(g.book,g.chapter,g.verses);
      g.representative=g.confirmedRepresentative || g.representative;
      // First referenced verse, not the newest/last range-member card.
      g.initial=g.kind==='passage' ? g.members.find(d=>Number(d.verse)===g.verses[0] && confirmed(d)) || g.representative : g.representative;
      g.count=g.members.length;
      g.methods=[...new Set(g.members.flatMap(d=>d.evidenceTypes || [d.type]).filter(Boolean))];
      g.latestUpdatedAt = Math.max(...g.members.map(activityTime));
      g.latestActivitySequence = Math.max(...g.members.filter(d=>activityTime(d)===g.latestUpdatedAt).map(activitySequence));
    }
    all.sort(recentFirst);
    return {groups:all, confirmed:all.filter(g=>g.section==='confirmed'), suggestions:all.filter(g=>g.section==='suggestions'), byEvent};
  }
  function annotate(records, model) {
    for (const d of records) {
      const g=model.byEvent.get(key(d)); if(!g)continue;
      d.discussionId=g.id; d.discussionLabel=g.label;
      const ann=announcement(d);
      if(ann) { d.passageId=ann.id; d.passageVerses=ann.verses; d.rangeStart=ann.verses[0];d.rangeEnd=ann.verses.at(-1);d.announcedPassage=true; }
    }
    return model;
  }
  function autoCandidate(group, now, settings) {
    // Never queue unread tail members of an announced passage. A new quotation
    // or separately announced verse is still eligible to advance the reading.
    return group.members.filter(d=>root.DetectionPolicy.eligible(d,settings,now).ok)
      .filter(d=>!d.announcedPassage || Number(d.verse)===d.passageVerses?.[0])
      .sort((a,b)=>position(b)-position(a) || time(b)-time(a) || a.verse-b.verse)[0] || null;
  }
  const api={build,annotate,confirmed,announcement,passageLabel,autoCandidate,CONTINUITY_MS,
    activityTime,evidenceChanged,nextActivity,compareReadingOrder:order};
  root.DetectionGroups=api;
  if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(typeof window!=='undefined'?window:globalThis);

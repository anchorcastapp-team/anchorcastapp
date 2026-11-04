/* v8.0.0 — non-destructive sermon trimming, optional voiceovers, explicit Cancel. */
(function () {
  'use strict';
  const byId = id => document.getElementById(id);
  const api = () => window.electronAPI?.recordingBranding;
  let dialog, current = null, resolveOffer = null, working = false, loading = false, cancelling = false;
  let finished = null, previewURL = null, generation = 0, previewGeneration = 0, clipSources = {}, clips = {};
  let original = null, activeRequest = null, activeTask = null, selectionTimer = null, selectionPlaying = false;
  const requestId = () => 'finish-' + (crypto.randomUUID ? crypto.randomUUID() : Date.now() + '-' + Math.random().toString(36).slice(2));
  async function call(method, data) {
    if (typeof api()?.[method] !== 'function') throw Error('Apply the recording trim patch, run node scripts/install-recording-trim.cjs, and restart AnchorCast.');
    const result = await api()[method](data);
    if (!result?.success) throw Error(result?.error || 'The audio operation failed.');
    return result;
  }
  function status(text, error = false) {
    byId('rbFinishStatus').textContent = text;
    byId('rbFinishStatus').classList.toggle('rb-error', error);
  }
  function timeString(value) {
    const ms = Math.max(0, Math.round(Number(value || 0) * 1000));
    return [Math.floor(ms / 3600000), Math.floor(ms / 60000) % 60, Math.floor(ms / 1000) % 60]
      .map(n => String(n).padStart(2, '0')).join(':') + '.' + String(ms % 1000).padStart(3, '0');
  }
  function parseTime(text) {
    const parts = String(text).trim().split(':');
    if (!parts.length || parts.length > 3 || !parts.every((p, i) => (i === parts.length - 1 ? /^\d+(?:\.\d{1,3})?$/ : /^\d+$/).test(p))) return NaN;
    const values = parts.map(Number);
    if (values.slice(1).some(n => n >= 60) || values.some(n => !Number.isFinite(n))) return NaN;
    return values.reduce((seconds, n) => seconds * 60 + n, 0);
  }
  function selection() {
    if (!original) return { valid: false, changed: false, error: 'Recording duration is not available yet.' };
    const start = parseTime(byId('rbTrimStart').value), enteredEnd = parseTime(byId('rbTrimEnd').value);
    if (!Number.isFinite(start) || !Number.isFinite(enteredEnd)) return { valid: false, changed: false, error: 'Enter seconds, MM:SS, or HH:MM:SS (up to three decimal places).' };
    const end = Math.min(original.seconds, enteredEnd);
    if (start < 0 || enteredEnd > original.seconds + .0011 || start >= original.seconds) return { valid: false, changed: false, error: 'Start and end must be within the original recording.' };
    if (end - start < .099999) return { valid: false, changed: false, error: 'End must be after start. Keep at least 0.1 second.' };
    return { valid: true, start, end, kept: end - start, changed: start > .0005 || original.seconds - end > .0011 };
  }
  function updatePublishingNote() {
    const slots = ['Intro', 'Outro'].filter(s => byId('rbFinish' + s)?.checked && clipSources[s] === 'openai');
    byId('rbFinishPublishingNote').hidden = slots.length === 0;
    const which = slots.length === 2 ? 'introduction and closing voiceovers are' : slots[0] === 'Intro' ? 'introduction voiceover is' : 'closing voiceover is';
    byId('rbFinishDisclosure').textContent = slots.length ? 'The ' + which + ' AI-generated. The sermon is the original recording' + (selection().changed ? ', trimmed at the beginning and/or end.' : '.') : '';
  }
  function flow() {
    const s = selection(), busy = working || loading || cancelling;
    updatePublishingNote();
    byId('rbFinishFlow').textContent = [byId('rbFinishIntro').checked ? 'Introduction + selected music' : '', s.valid && s.changed ? 'Trimmed sermon audio' : 'Original sermon audio', byId('rbFinishOutro').checked ? 'Closing + selected music' : ''].filter(Boolean).join(' → ');
    byId('rbFinishCreate').disabled = busy || !!finished || (original && !s.valid) || !(s.changed || byId('rbFinishIntro').checked || byId('rbFinishOutro').checked);
    byId('rbTrimFields').disabled = busy || !original;
    byId('rbTrimPlay').disabled = busy || !s.valid;
    byId('rbTrimReset').disabled = busy || !original;
    byId('rbTrimRetry').disabled = busy;
    byId('rbFinishKeep').disabled = busy;
    byId('rbFinishKeep').textContent = finished ? 'Done' : 'Keep original only';
    byId('rbFinishCancel').disabled = cancelling;
    byId('rbFinishCancel').textContent = cancelling ? 'Cancelling…' : 'Cancel';
    for (const suffix of ['Intro', 'Outro']) {
      byId('rbFinish' + suffix).disabled = busy || !clips[suffix]?.ready;
      byId('rbPreview' + suffix).disabled = busy || !clips[suffix]?.ready;
    }
    byId('rbFinishSettings').disabled = busy;
    byId('rbFinishCopy').disabled = busy || !finished;
    byId('rbFinishFolder').disabled = cancelling;
    const invalid = original && !s.valid;
    byId('rbTrimError').textContent = invalid ? s.error : '';
    byId('rbTrimStart').setAttribute('aria-invalid', invalid ? 'true' : 'false');
    byId('rbTrimEnd').setAttribute('aria-invalid', invalid ? 'true' : 'false');
    if (s.valid) {
      byId('rbTrimSummary').textContent = 'Keep ' + timeString(s.kept) + ' · Remove beginning ' + timeString(s.start) + ' · Remove ending ' + timeString(original.seconds - s.end);
      byId('rbTrimStartSlider').value = String(s.start); byId('rbTrimEndSlider').value = String(s.end);
      byId('rbTrimStartSlider').setAttribute('aria-valuetext', timeString(s.start));
      byId('rbTrimEndSlider').setAttribute('aria-valuetext', timeString(s.end));
    } else byId('rbTrimSummary').textContent = '';
  }
  function pauseSelection() {
    selectionPlaying = false; clearInterval(selectionTimer); selectionTimer = null;
    byId('rbTrimPlayer')?.pause(); if (byId('rbTrimPlay')) byId('rbTrimPlay').textContent = '▶ Play kept section';
  }
  function stopPreview() {
    previewGeneration++; const p = byId('rbFinishPlayer');
    if (p) { p.pause(); p.removeAttribute('src'); p.load(); p.hidden = true; }
    if (previewURL) URL.revokeObjectURL(previewURL); previewURL = null;
  }
  function releasePreview(info) { if (info?.previewToken) void call('releaseRecordingPreview', { previewToken: info.previewToken }).catch(() => {}); }
  function clearOriginal() {
    pauseSelection(); const p = byId('rbTrimPlayer');
    if (p) { p.pause(); p.removeAttribute('src'); p.load(); p.hidden = true; }
    releasePreview(original); original = null;
  }
  function complete() {
    if (working || loading) return;
    generation++; stopPreview(); clearOriginal();
    if (dialog.open) dialog.close('done');
    const result = finished, resolve = resolveOffer; current = null; resolveOffer = null; cancelling = false;
    resolve?.(result);
  }
  async function cancelOrDone() {
    if (cancelling) return;
    if (!working && !loading) { complete(); return; }
    cancelling = true; pauseSelection(); stopPreview(); flow(); status('Cancelling this operation. The original recording is safe…');
    const task = activeTask, id = activeRequest;
    try {
      if (id) await call('cancel', { requestId: id });
      // Wait for the compose/inspection response as well, not just the cancel IPC.
      // This closes cleanly even when a successful export races the Cancel click.
      await task?.catch(() => {}); working = false; loading = false; complete();
    } catch (e) { cancelling = false; flow(); status('Could not cancel: ' + e.message + ' Keep AnchorCast open until the operation finishes.', true); }
  }
  async function refresh() {
    const epoch = generation, s = await call('get');
    if (epoch !== generation || !dialog.open || working || cancelling || finished) return;
    for (const slot of ['intro', 'outro']) {
      const suffix = slot === 'intro' ? 'Intro' : 'Outro', clip = s.clips[slot];
      clips[suffix] = clip; clipSources[suffix] = clip.source;
      if (!clip.ready) byId('rbFinish' + suffix).checked = false;
      byId('rbFinish' + suffix + 'Info').textContent = clip.ready ? 'Ready · ' + Math.round(clip.seconds) + ' seconds · ' + (clip.source === 'openai' ? 'AI-narrated' : 'Imported narration') : clip.error || clip.state;
    }
    flow();
  }
  function editSelection() {
    pauseSelection();
    if (finished) { finished = null; byId('rbFinishCopy').hidden = true; status('Selection changed. The previous finished MP3 is kept; create a new copy for these edits.'); }
    flow();
  }
  async function inspect() {
    if (!current || working || loading || cancelling) return;
    const epoch = generation, name = current, id = requestId(); activeRequest = id;
    loading = true; clearOriginal(); byId('rbTrimRetry').hidden = true;
    byId('rbTrimInfo').textContent = 'Reading recording duration…'; flow();
    const task = (async () => {
      try {
        const info = await call('recordingInfo', { name, requestId: id });
        if (epoch !== generation || !dialog.open || cancelling) { releasePreview(info); return; }
        if (!Number.isFinite(info.seconds) || info.seconds <= 0 || !/^http:\/\/127\.0\.0\.1:\d+\/recording\/[a-f0-9]{64}$/.test(info.previewURL || '')) { releasePreview(info); throw Error('Recording preview could not be prepared.'); }
        original = info; byId('rbTrimInfo').textContent = 'Original duration: ' + timeString(info.seconds);
        byId('rbTrimStart').value = timeString(0); byId('rbTrimEnd').value = timeString(info.seconds);
        for (const suffix of ['Start', 'End']) byId('rbTrim' + suffix + 'Slider').max = String(info.seconds);
        const player = byId('rbTrimPlayer'); player.src = info.previewURL; player.hidden = false; player.load();
      } catch (e) {
        if (epoch === generation && !cancelling) { byId('rbTrimInfo').textContent = 'Trim preview unavailable: ' + e.message; byId('rbTrimRetry').hidden = false; }
      } finally { if (epoch === generation) { loading = false; if (activeRequest === id) activeRequest = null; flow(); } }
    })();
    activeTask = task; await task; if (activeTask === task) activeTask = null;
  }
  async function compose() {
    const s = selection();
    if (!current || working || loading || cancelling || finished || (original && !s.valid)) return;
    if (!(s.changed || byId('rbFinishIntro').checked || byId('rbFinishOutro').checked)) return;
    const epoch = generation, id = requestId(); activeRequest = id; working = true;
    pauseSelection(); stopPreview(); flow(); status('Creating the finished MP3 locally. Your original is safe…');
    const data = { name: current, intro: byId('rbFinishIntro').checked, outro: byId('rbFinishOutro').checked, requestId: id };
    if (original) data.sourceIdentity = original.sourceIdentity;
    if (s.valid && s.changed) data.trim = { startSeconds: s.start, endSeconds: s.end };
    const task = (async () => {
      try {
        const r = await call('compose', data); if (epoch !== generation) return;
        finished = r; byId('rbFinishCopy').hidden = false;
        status('Finished MP3 saved: ' + r.path + '. The original recording is unchanged.');
      } catch (e) { if (epoch === generation && !cancelling) status(e.message, true); }
      finally {
        if (epoch === generation) { working = false; if (activeRequest === id) activeRequest = null; flow(); if (!finished && !cancelling) void refresh().catch(e => status(e.message, true)); }
      }
    })();
    activeTask = task; await task; if (activeTask === task) activeTask = null;
  }
  function mount() {
    if (dialog) return;
    dialog = document.createElement('dialog'); dialog.id = 'recordingFinishDialog'; dialog.className = 'sermon-audio-dialog rb-finish';
    dialog.setAttribute('aria-labelledby', 'rbFinishTitle');
    dialog.innerHTML = `<div class="sermon-audio-body"><h2 id="rbFinishTitle">Trim and finish recording</h2>
      <p class="rb-note"><strong>Your original recording is already saved.</strong> Keep the section you want, then optionally add the prepared introduction and closing. All edits create a separate MP3.</p>
      <section class="rb-trim" aria-labelledby="rbTrimTitle"><h3 id="rbTrimTitle">1 · Trim the sermon <span>(optional)</span></h3>
        <p class="rb-note">Set the beginning and ending of the part to keep. Intro and closing clips are not trimmed.</p>
        <p id="rbTrimInfo" class="rb-note" role="status">Reading recording duration…</p>
        <audio id="rbTrimPlayer" controls preload="metadata" hidden aria-label="Original sermon recording preview"></audio>
        <fieldset id="rbTrimFields" disabled><legend class="rb-sr-only">Sermon start and end positions</legend><div class="rb-trim-grid">
          <div><label for="rbTrimStart">Start at</label><input id="rbTrimStart" type="text" inputmode="decimal" value="00:00:00.000" aria-describedby="rbTrimFormat rbTrimError" autocomplete="off" spellcheck="false">
            <input id="rbTrimStartSlider" type="range" min="0" max="1" step="0.001" value="0" aria-label="Trim start position">
            <button type="button" id="rbTrimSetStart">Set start to playhead</button></div>
          <div><label for="rbTrimEnd">End at</label><input id="rbTrimEnd" type="text" inputmode="decimal" value="00:00:00.000" aria-describedby="rbTrimFormat rbTrimError" autocomplete="off" spellcheck="false">
            <input id="rbTrimEndSlider" type="range" min="0" max="1" step="0.001" value="1" aria-label="Trim end position">
            <button type="button" id="rbTrimSetEnd">Set end to playhead</button></div>
        </div></fieldset>
        <p id="rbTrimFormat" class="rb-note">Times: HH:MM:SS.mmm. You may also enter seconds or MM:SS.</p>
        <p id="rbTrimError" class="rb-error" role="alert"></p><p id="rbTrimSummary" class="rb-trim-summary" aria-live="polite"></p>
        <div class="rb-actions"><button type="button" id="rbTrimPlay" disabled>▶ Play kept section</button><button type="button" id="rbTrimReset" disabled>Reset trim</button><button type="button" id="rbTrimRetry" hidden>Retry loading</button></div>
      </section>
      <h3 class="rb-finish-step">2 · Add an introduction or closing <span>(optional)</span></h3>
      <label class="rb-choice"><input id="rbFinishIntro" type="checkbox" disabled><span><strong>Add introduction</strong><small id="rbFinishIntroInfo">Loading…</small></span></label>
      <label class="rb-choice"><input id="rbFinishOutro" type="checkbox" disabled><span><strong>Add closing message</strong><small id="rbFinishOutroInfo">Loading…</small></span></label>
      <div class="rb-flow" id="rbFinishFlow">Original sermon audio</div>
      <div id="rbFinishPublishingNote" class="rb-note" hidden><strong>Publishing note — not added to the audio</strong><p>Include this disclosure in the episode description or a visible player notice for listeners:</p><p id="rbFinishDisclosure"></p></div>
      <div class="rb-actions"><button type="button" id="rbPreviewIntro" disabled>▶ Preview introduction</button><button type="button" id="rbPreviewOutro" disabled>▶ Preview closing</button></div>
      <audio controls hidden id="rbFinishPlayer" preload="none" aria-label="Prepared voiceover preview"></audio>
      <p id="rbFinishStatus" role="status" aria-live="polite">Prepared clips are reused locally; trimming sends no audio online.</p>
      <div class="rb-actions"><button type="button" id="rbFinishSettings">Recording settings</button><button type="button" id="rbFinishCopy" hidden>Save finished MP3 As…</button><button type="button" id="rbFinishFolder">Open finished messages</button></div>
      <p class="rb-note">Music stays in the introduction and closing, not underneath the preacher. Cancel keeps the original and any finished copies already saved.</p>
      </div><footer class="rb-finish-footer"><button type="button" id="rbFinishCreate" class="rb-primary" disabled>Create finished MP3</button><button type="button" id="rbFinishKeep">Keep original only</button><button type="button" id="rbFinishCancel">Cancel</button></footer>`;
    document.body.appendChild(dialog);
    for (const suffix of ['Intro', 'Outro']) {
      byId('rbFinish' + suffix).onchange = editSelection;
      byId('rbPreview' + suffix).onclick = async () => {
        pauseSelection(); stopPreview(); const epoch = generation, id = previewGeneration;
        try {
          const r = await call('preview', { slot: suffix === 'Intro' ? 'intro' : 'outro' });
          if (!dialog.open || working || cancelling || generation !== epoch || previewGeneration !== id) return;
          previewURL = URL.createObjectURL(new Blob([r.bytes], { type: 'audio/mpeg' })); const p = byId('rbFinishPlayer'); p.src = previewURL; p.hidden = false; await p.play();
        } catch (e) { if (dialog.open && epoch === generation && id === previewGeneration) status(e.message, true); }
      };
    }
    for (const suffix of ['Start', 'End']) {
      byId('rbTrim' + suffix).oninput = editSelection;
      byId('rbTrim' + suffix).onchange = () => { const v = parseTime(byId('rbTrim' + suffix).value); if (Number.isFinite(v)) byId('rbTrim' + suffix).value = timeString(v); editSelection(); };
      byId('rbTrim' + suffix + 'Slider').oninput = () => { byId('rbTrim' + suffix).value = timeString(Number(byId('rbTrim' + suffix + 'Slider').value)); editSelection(); };
      byId('rbTrimSet' + suffix).onclick = () => { byId('rbTrim' + suffix).value = timeString(Math.min(original.seconds, byId('rbTrimPlayer').currentTime)); editSelection(); };
    }
    byId('rbTrimReset').onclick = () => { byId('rbTrimStart').value = timeString(0); byId('rbTrimEnd').value = timeString(original.seconds); editSelection(); };
    byId('rbTrimRetry').onclick = () => void inspect();
    const player = byId('rbTrimPlayer');
    const stopAtEnd = () => { const s = selection(); if (selectionPlaying && (!s.valid || player.currentTime >= s.end - .01)) { pauseSelection(); if (s.valid) player.currentTime = s.end; } };
    player.addEventListener('timeupdate', stopAtEnd);
    player.addEventListener('ended', pauseSelection);
    player.addEventListener('pause', () => { clearInterval(selectionTimer); selectionTimer = null; selectionPlaying = false; byId('rbTrimPlay').textContent = '▶ Play kept section'; });
    player.addEventListener('play', () => {
      const s = selection(); if (!s.valid || working || loading || cancelling) { pauseSelection(); return; }
      stopPreview(); if (player.currentTime < s.start || player.currentTime >= s.end) player.currentTime = s.start;
      selectionPlaying = true; byId('rbTrimPlay').textContent = 'Pause preview'; clearInterval(selectionTimer); selectionTimer = setInterval(stopAtEnd, 40);
    });
    player.addEventListener('seeking', () => { const s = selection(); if (selectionPlaying && s.valid && player.currentTime < s.start) player.currentTime = s.start; stopAtEnd(); });
    player.addEventListener('error', () => { if (original && dialog.open) status('Audio preview could not play. Retry loading or check the saved recording. The original has not been changed.', true); });
    byId('rbTrimPlay').onclick = async () => {
      if (selectionPlaying) { pauseSelection(); return; }
      const s = selection(); if (!s.valid) return;
      stopPreview(); player.currentTime = s.start;
      try { await player.play(); } catch (e) { status('Could not play the selection: ' + e.message, true); }
    };
    byId('rbFinishKeep').onclick = complete;
    byId('rbFinishCancel').onclick = () => void cancelOrDone();
    byId('rbFinishSettings').onclick = async () => { pauseSelection(); stopPreview(); try { await call('openSettings'); status('Prepare the clips in Recording settings. Return here after saving; your trim selection stays in place.'); } catch (e) { status(e.message, true); } };
    byId('rbFinishFolder').onclick = () => void call('openExports').catch(e => status(e.message, true));
    byId('rbFinishCopy').onclick = async () => { if (!finished) return; try { const r = await call('saveExportAs', { name: finished.name }); if (!r.cancelled && dialog.open) status('Finished MP3 copied to: ' + r.path); } catch (e) { if (dialog.open) status(e.message, true); } };
    byId('rbFinishCreate').onclick = () => void compose();
    dialog.addEventListener('cancel', e => { e.preventDefault(); e.stopPropagation(); void cancelOrDone(); });
    api()?.onChanged?.(() => { if (dialog.open && !working && !finished && !cancelling) void refresh().catch(e => status(e.message, true)); });
    api()?.onProgress?.(p => {
      if (!dialog.open || cancelling || !activeRequest || p.requestId !== activeRequest || p.phase === 'idle') return;
      if (loading) byId('rbTrimInfo').textContent = 'Reading recording duration…' + (p.seconds ? ' ' + timeString(p.seconds) + ' scanned' : '');
      else if (working) status(p.phase + (p.seconds ? ' · ' + Math.floor(p.seconds) + ' seconds processed' : ''));
    });
  }
  function offer(name) {
    mount(); if (!api()) return Promise.resolve(null);
    if (current) return Promise.reject(Error('Finish or dismiss the current audio packaging dialog first.'));
    current = name; generation++; finished = null; working = false; loading = false; cancelling = false; clipSources = {}; clips = {}; activeRequest = null;
    stopPreview(); clearOriginal();
    for (const suffix of ['Intro', 'Outro']) { byId('rbFinish' + suffix).checked = false; byId('rbFinish' + suffix).disabled = true; byId('rbFinish' + suffix + 'Info').textContent = 'Loading…'; }
    byId('rbFinishCopy').hidden = true; byId('rbTrimRetry').hidden = true;
    status('Keep the full recording or trim its beginning and end. Add voiceovers only when selected.'); flow();
    dialog.showModal(); byId('rbFinishCancel').focus();
    const result = new Promise(resolve => { resolveOffer = resolve; });
    void refresh().catch(e => { if (dialog.open) status(e.message, true); }); void inspect();
    return result;
  }
  window.RecordingBrandingUI = { offer };
})();

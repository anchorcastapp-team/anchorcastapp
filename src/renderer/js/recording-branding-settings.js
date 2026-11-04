/* v8.0.0 — reusable narration. This renderer never holds or sends the API key. */
(function () {
  'use strict';
  let mounted = false, data = null, dirty = false, busy = false, requested = false, loaded = null, objectURL = null;
  const api = () => window.electronAPI?.recordingBranding;
  const el = id => document.getElementById(id);
  const keys = ['introText', 'outroText', 'phone', 'email', 'platform', 'handle', 'voice', 'speed', 'instructions', 'music', 'musicVolume', 'leadSeconds', 'tailSeconds', 'gapSeconds'];
  const numeric = new Set(['speed', 'musicVolume', 'leadSeconds', 'tailSeconds', 'gapSeconds']);
  function message(text, error = false) { if (!el('rbStatus')) return; el('rbStatus').textContent = text; el('rbStatus').classList.toggle('rb-error', error); }
  async function call(method, value) { if (!api()?.[method]) throw Error('Recording settings bridge is not installed. Apply the complete v8.0.0 update and restart AnchorCast.'); const r = await api()[method](value); if (!r?.success) throw Error(r?.error || 'Recording settings operation failed.'); return r; }
  function selectTab() {
    requested = true;
    if (!mounted || typeof setSection !== 'function') return;
    const sidebar = document.querySelector('.sidebar-item[data-section="audio"]');
    if (sidebar) setSection(sidebar, 'audio');
    if (typeof setTab === 'function') { const tab = [...document.querySelectorAll('.mod-tab')].find(t => t.textContent.trim().toLowerCase() === 'recording'); if (tab) setTab(tab, 'recording'); }
  }
  function draft() { const c = {}; for (const key of keys) c[key] = numeric.has(key) ? Number(el('rb-' + key).value) : el('rb-' + key).value; return c; }
  function updateLabels() {
    el('rbSpeedLabel').textContent = Number(el('rb-speed').value).toFixed(2) + '×';
    el('rbMusicLevelLabel').textContent = el('rb-musicVolume').value + '%';
  }
  function renderState() {
    if (!data) return;
    updateLabels();
    el('rbKeyStatus').textContent = data.keyAvailable ? 'Saved OpenAI key available. Generation uses your API account.' : 'No OpenAI key found. Save it in Audio & AI → General, or import narration instead.';
    el('rbMusicName').textContent = data.hasCustomMusic ? 'Imported: ' + data.customMusicName : 'No custom track imported.';
    for (const slot of ['intro', 'outro']) {
      const clip = data.clips[slot], tag = el('rb-' + slot + '-state');
      tag.textContent = dirty ? 'Unsaved changes — save before preview' : (clip.error || clip.state) + (clip.ready ? ' · ' + Math.round(clip.seconds) + ' sec' : '');
      tag.classList.toggle('rb-ready', clip.ready && !dirty);
      el('rb-' + slot + '-generate').disabled = busy;
      el('rb-' + slot + '-sample').disabled = busy || !data.samples?.[slot + 'Text'];
      el('rb-' + slot + '-import').disabled = busy;
      el('rb-' + slot + '-mix').disabled = busy || !clip.voiceCurrent;
      el('rb-' + slot + '-preview').disabled = busy || dirty || !clip.ready;
    }
    for (const key of keys) el('rb-' + key).disabled = busy;
    for (const id of ['rbSave', 'rbReload', 'rbImportMusic', 'rbPreviewMusic']) el(id).disabled = busy;
    el('rbCancel').hidden = !busy;
    el('rbDirty').textContent = dirty ? 'Unsaved Recording settings' : 'Recording settings saved on this computer';
  }
  async function load(force = false) {
    if (dirty && !force) return;
    data = await call('get');
    for (const key of keys) el('rb-' + key).value = data.config[key];
    dirty = false; renderState();
  }
  async function saveIfDirty() {
    if (!dirty) return;
    if (busy) throw Error('Wait for the audio operation before saving Recording settings.');
    await call('save', { config: draft(), revision: data.revision }); await load(true);
    message('Recording settings saved. Text changes require fresh narration; music changes only require Rebuild mix.');
  }
  function stopPreview() { const audio = el('rbPlayer'); if (audio) { audio.pause(); audio.removeAttribute('src'); audio.load(); } if (objectURL) URL.revokeObjectURL(objectURL); objectURL = null; }
  async function playResult(r) {
    stopPreview(); objectURL = URL.createObjectURL(new Blob([r.bytes], { type: r.mime || 'audio/mpeg' }));
    const audio = el('rbPlayer'); audio.src = objectURL; audio.hidden = false;
    await audio.play();
  }
  async function action(method, slot) {
    if (busy) return;
    try {
      if (method === 'generate' && !el('rbConsent').checked) throw Error('Tick the voice-generation consent box before generating speech.');
      if ((method === 'importVoice' || method === 'importMusic') && !el('rbImportPermission').checked) throw Error('Confirm permission to use your imported narration or music.');
      await saveIfDirty(); stopPreview(); busy = true; renderState(); message(method === 'generate' ? 'Generating voice. Only this script is sent to OpenAI…' : 'Preparing audio…');
      const r = await call(method, { slot, consent: el('rbConsent').checked, permission: el('rbImportPermission').checked, revision: data.revision });
      busy = false; await load(true);
      message(r.cancelled ? 'Import cancelled. Existing audio is unchanged.' : method === 'importMusic' ? 'Music imported. Rebuild each prepared mix to use this track; no new voice generation is needed.' : 'Audio prepared. Preview it before using it in a sermon.');
    } catch (e) { busy = false; message(e.message, true); if (data) renderState(); }
  }
  function mount() {
    if (mounted) return;
    const pane = el('pane-audio-recording'); if (!pane) return; mounted = true;
    pane.innerHTML = `<div class="rb-settings">
      <div class="rb-heading"><span class="rb-eyebrow">RECORDING · INTRO &amp; CLOSING</span><h2>Give every message a welcome.</h2><p>Prepare your voiceovers once. Choose whether to add them when a sermon is saved.</p><span class="rb-pill">Original recording always retained</span></div>
      <div class="rb-section"><h3>Church contact details</h3><p>Used only where the script contains a matching placeholder. Leave out unused sentences, or complete every placeholder before generating.</p>
        <div class="rb-grid">${[['phone','Phone number','tel'],['email','Email address','text'],['platform','Social platform','text'],['handle','Social handle (without @)','text']].map(([k,l,t])=>`<label>${l}<input id="rb-${k}" type="${t}" maxlength="${k==='email'?200:100}" autocomplete="off"></label>`).join('')}</div>
      </div>
      <div class="rb-section"><h3>Narration voice</h3><p>Formal male narration preset. Preview pronunciation and delivery before using a clip; the voice is synthetic, not an impersonation.</p>
        <div class="rb-grid"><label>Voice<select id="rb-voice"><option value="onyx">Onyx — default</option value="cedar">Cedar</option><option value="ash">Ash</option><option value="echo">Echo</option><option value="fable">Fable</option><option value="alloy">Alloy</option><option value="marin">Marin</option></select></label><label>Speaking speed <output id="rbSpeedLabel">0.92×</output><input id="rb-speed" type="range" min="0.75" max="1.2" step="0.01"></label></div>
        <details><summary>Voice direction</summary><label>Delivery instructions<textarea id="rb-instructions" rows="4" maxlength="1800"></textarea></label></details>
        <p id="rbKeyStatus"></p><p id="rbSpeechProviderNote"><strong>Voice audio provider: OpenAI.</strong> A Claude (Anthropic) API key cannot generate MP3 voiceovers through the current public Claude API. It produces text, not speech audio. Use an OpenAI key for Generate voice, or Import audio for narration prepared elsewhere.</p><p><strong>Generate voice uses OpenAI text-to-speech and may incur API charges.</strong> Only the selected script, filled contact details and voice direction are sent. Saved clips can be reused offline, without another API request.</p>
        <label class="rb-check"><input id="rbConsent" type="checkbox"><span>I agree to send a script to OpenAI when I click Generate voice. Only the text in the selected script is spoken; no announcement is added automatically.</span></label>
        <p id="rbPublishingNote" class="rb-note"><strong>Listener notice — not spoken:</strong> Generated narration is AI-generated. When sharing it, clearly disclose this to listeners in the episode description or a visible player notice. The application does not add these words to the audio.</p>
      </div>
      ${['intro','outro'].map(slot=>`<section class="rb-section rb-script"><div class="rb-section-top"><h3>${slot==='intro'?'Introduction':'Closing message'}</h3><span class="rb-tag" id="rb-${slot}-state">Not prepared</span></div>
        <label>${slot==='intro'?'Opening':'Closing'} text<textarea id="rb-${slot}Text" rows="6" maxlength="3500"></textarea></label>
        <p>Editable sample text — customize it for your message. Up to 3,500 characters / 3 minutes. You may also import your own narration. Clear the script first when it does not describe your imported audio.</p>
        <div class="rb-actions"><button type="button" id="rb-${slot}-generate" class="rb-primary">Generate voice</button><button type="button" id="rb-${slot}-import">Import audio…</button><button type="button" id="rb-${slot}-mix">Rebuild mix</button><button type="button" id="rb-${slot}-preview">▶ Preview</button><button type="button" id="rb-${slot}-sample" disabled>Use sample text</button></div></section>`).join('')}
      <div class="rb-section"><h3>Background sound</h3><p>Music is mixed under the introduction and closing only — never under the preacher’s recording.</p>
        <label>Track<select id="rb-music"><option value="warm">Warm Welcome — included instrumental</option><option value="none">None — voice only</option><option value="custom">My imported background music</option></select></label>
        <p id="rbMusicName"></p><div class="rb-actions"><button type="button" id="rbImportMusic">Import music…</button><button type="button" id="rbPreviewMusic">▶ Preview track</button></div>
        <label>Music level beneath voice <output id="rbMusicLevelLabel">18%</output><input id="rb-musicVolume" type="range" min="0" max="35" step="1"></label>
        <div class="rb-grid"><label>Music lead-in (seconds)<input id="rb-leadSeconds" type="number" min="0" max="4" step="0.1"></label><label>Music ending (seconds)<input id="rb-tailSeconds" type="number" min="0" max="5" step="0.1"></label><label>Gap between sections (seconds)<input id="rb-gapSeconds" type="number" min="0" max="2" step="0.1"></label></div>
        <p>Changing music, level or fades does not require another paid voice generation. Save settings, then Rebuild mix for each prepared clip. A gap change applies at the next export.</p>
        <label class="rb-check"><input id="rbImportPermission" type="checkbox"><span>I have permission to use imported narration/music in recordings I distribute.</span></label>
      </div>
      <audio controls id="rbPlayer" hidden preload="none" aria-label="Voiceover or background preview"></audio><p class="rb-note">Preview plays on this computer’s normal audio output. Use headphones during recording to avoid capturing the preview through a microphone.</p>
      <div class="rb-save"><span id="rbDirty">Loading…</span><div class="rb-actions"><button type="button" id="rbSave" class="rb-primary">Save Recording settings</button><button type="button" id="rbReload">Reload saved</button><button type="button" id="rbCancel" hidden>Cancel operation</button></div></div>
      <p id="rbStatus" role="status" aria-live="polite">Loading Recording settings…</p></div>`;
    for (const key of keys) el('rb-' + key).addEventListener('input', () => { dirty = true; stopPreview(); renderState(); });
    el('rbSave').onclick = () => void saveIfDirty().catch(e => message(e.message, true));
    el('rbReload').onclick = () => { if (!dirty || window.confirm('Discard unsaved Recording settings?')) void load(true).then(() => message('Saved Recording settings loaded.')).catch(e => message(e.message, true)); };
    el('rbCancel').onclick = () => void call('cancel').then(() => message('Cancelled. Original recording and previous audio retained.')).catch(e => message(e.message, true));
    el('rbImportMusic').onclick = () => void action('importMusic');
    el('rbPreviewMusic').onclick = async () => { try { await saveIfDirty(); await playResult(await call('previewMusic')); } catch (e) { message(e.message, true); } };
    for (const slot of ['intro', 'outro']) {
      el('rb-' + slot + '-sample').onclick = () => {
        const sample = data?.samples?.[slot + 'Text'], field = el('rb-' + slot + 'Text');
        if (busy || !sample || field.value === sample) return;
        if (field.value.trim() && !window.confirm('Replace this script with the generic sample text? Existing audio is kept, but you will need to prepare it again after saving.')) return;
        field.value = sample; dirty = true; stopPreview(); renderState();
        message('Sample text loaded. Edit it, then save and Generate voice when ready. No audio has been generated.');
      };
      el('rb-' + slot + '-generate').onclick = () => void action('generate', slot);
      el('rb-' + slot + '-import').onclick = () => void action('importVoice', slot);
      el('rb-' + slot + '-mix').onclick = () => void action('prepare', slot);
      el('rb-' + slot + '-preview').onclick = async () => { try { await playResult(await call('preview', { slot })); } catch (e) { message(e.message, true); } };
    }
    api()?.onProgress?.(p => { if (p.phase !== 'idle') message(p.phase); });
    api()?.onChanged?.(() => { if (!dirty && !busy) void load().catch(e => message(e.message, true)); });
    api()?.onOpenRecording?.(selectTab);
    loaded = load().then(() => message('Edit the scripts, save, and prepare the audio. Nothing is sent until you click Generate voice.')).catch(e => message(e.message, true));
    if (requested) selectTab();
  }
  window.RecordingBrandingSettings = { saveIfDirty, afterInit() { if (requested) selectTab(); }, open: selectTab, isDirty: () => dirty };
  document.addEventListener('DOMContentLoaded', mount);
  window.addEventListener('beforeunload', event => { stopPreview(); if (dirty && !window.confirm('Discard unsaved Recording settings?')) { event.preventDefault(); event.returnValue = false; } });
})();

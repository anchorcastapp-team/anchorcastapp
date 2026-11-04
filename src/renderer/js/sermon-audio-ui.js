/* Recorder controls: explicit Minimize and guarded Close. UI revision recorder-window-v3-branding.
 * recorder-indicator-v1: persistent REC dot, production-theme-safe state colours. */
(function () {
  'use strict';

  function time(seconds) {
    const n = Math.max(0, Math.floor(seconds || 0));
    return [Math.floor(n / 3600), Math.floor(n % 3600 / 60), n % 60]
      .map(v => String(v).padStart(2, '0')).join(':');
  }

  document.addEventListener('DOMContentLoaded', () => {
    const launch = document.getElementById('sermonAudioBtn');
    if (!launch || document.getElementById('sermonAudioDialog')) return;
    const byId = id => document.getElementById(id);
    const bridge = window.electronAPI?.sermonAudio;
    const dialog = document.createElement('dialog');
    dialog.id = 'sermonAudioDialog';
    dialog.className = 'sermon-audio-dialog';
    dialog.dataset.uiRevision = 'recorder-window-v3-branding';
    dialog.setAttribute('aria-labelledby', 'sermonAudioTitle');
    dialog.innerHTML = `
      <header>
        <div class="sermon-audio-heading"><h2 id="sermonAudioTitle">Sermon Audio Recorder</h2><p>Microphone recording · MP3 · Saved on this computer</p></div>
        <div class="sermon-audio-window-actions" role="group" aria-label="Recorder window controls">
          <button type="button" id="sermonAudioMinimize" title="Minimize controls; keep recording" aria-label="Minimize recorder controls; keep recording"><span aria-hidden="true">−</span></button>
          <button type="button" id="sermonAudioClose" title="Close recorder; active recording asks to stop and save" aria-label="Close recorder controls"><span aria-hidden="true">×</span></button>
        </div>
      </header>
      <div class="sermon-audio-body">
        <p class="sermon-audio-note">Uses the microphone selected in <strong>Settings → Audio &amp; AI</strong>. Recording is separate from Start Transcript.</p>
        <div class="sermon-audio-monitor"><span id="sermonAudioState">Ready</span><strong id="sermonAudioElapsed">00:00:00</strong><meter id="sermonAudioLevel" min="0" max="1" value="0" aria-label="Microphone input level"></meter><small id="sermonAudioLevelText">Input level appears when recording</small></div>
        <p id="sermonAudioStatus" role="status" aria-live="polite">Record only with the speaker’s knowledge and permission.</p>
        <div class="sermon-audio-actions"><button type="button" id="sermonAudioStart">● Start Recording</button><button type="button" id="sermonAudioPause" disabled>Pause</button><button type="button" id="sermonAudioStop" disabled>Stop &amp; Save</button></div>
        <p class="sermon-audio-note">128 kbps mono MP3. Pause affects this audio file only. <strong>Minimize (−)</strong> hides these controls without stopping or resuming recording. Use the <strong>Audio</strong> button in the main window to restore them. <strong>Close (×)</strong> asks to stop and save an active recording.</p>
        <hr><label for="sermonAudioFiles">Saved recordings</label><select id="sermonAudioFiles" aria-label="Saved recordings"><option value="">No saved recordings</option></select>
        <div class="sermon-audio-actions"><button type="button" id="sermonAudioCopy" disabled>Save MP3 As…</button><button type="button" id="sermonAudioFolder">Open Recordings</button><button type="button" id="sermonAudioBrand" disabled>Trim / Add Intro / Closing…</button></div>
        <p class="sermon-audio-note">Audio is saved continuously to the app’s Recordings folder. An interrupted recording may be incomplete; listen before sharing.</p>
      </div>`;
    document.body.appendChild(dialog);

    // Separate alert dialog: Escape/Cancel dismisses the question, never the recording.
    const closeDialog = document.createElement('dialog');
    closeDialog.id = 'sermonAudioCloseDialog';
    closeDialog.className = 'sermon-audio-dialog sermon-audio-close-dialog';
    closeDialog.setAttribute('role', 'alertdialog');
    closeDialog.setAttribute('aria-labelledby', 'sermonAudioCloseTitle');
    closeDialog.setAttribute('aria-describedby', 'sermonAudioCloseMessage');
    closeDialog.innerHTML = `<div class="sermon-audio-body">
      <h2 id="sermonAudioCloseTitle">Stop recording and close?</h2>
      <p id="sermonAudioCloseMessage"></p>
      <div class="sermon-audio-actions"><button type="button" id="sermonAudioCloseCancel" autofocus>Cancel</button><button type="button" id="sermonAudioCloseSave">Stop &amp; Save, then Close</button></div>
    </div>`;
    document.body.appendChild(closeDialog);
    const notice = document.createElement('span');
    notice.className = 'sermon-audio-sr-only';
    notice.setAttribute('role', 'status');
    notice.setAttribute('aria-live', 'polite');
    launch.insertAdjacentElement('afterend', notice);
    launch.setAttribute('aria-haspopup', 'dialog');
    launch.setAttribute('aria-controls', dialog.id);
    launch.setAttribute('aria-expanded', 'false');
    launch.dataset.indicatorRevision = 'recorder-indicator-v1';
    dialog.dataset.indicatorRevision = 'recorder-indicator-v1';
    // Keep the animated element mounted. Recreating it for every PCM update would
    // continually restart its pulse. The visible word also conveys state without colour.
    const recordDot = document.createElement('span');
    recordDot.className = 'sermon-audio-record-dot';
    recordDot.setAttribute('aria-hidden', 'true');
    const buttonText = document.createElement('span');
    buttonText.className = 'sermon-audio-button-text';
    launch.replaceChildren(recordDot, buttonText);
    const stateDot = recordDot.cloneNode(false);
    const stateText = document.createElement('span');
    byId('sermonAudioState').replaceChildren(stateDot, stateText);
    notice.id = 'sermonAudioAnnouncement';
    notice.setAttribute('aria-atomic', 'true');

    let lastState = 'idle', refreshBusy = false, recordings = [];
    let packaging = false;
    let minimized = false, closePending = false, closeProblem = false, closeGeneration = 0;
    const recorder = new window.SermonAudioRecorder({
      bridge,
      getMicrophone: () => typeof State !== 'undefined' ? State.settings?.microphoneId || 'default' : 'default',
      onChange: update,
    });
    const busy = () => packaging || ['starting', 'recording', 'paused', 'stopping'].includes(recorder.state);

    function status(message, error = false) {
      byId('sermonAudioStatus').textContent = message;
      byId('sermonAudioStatus').classList.toggle('error', error);
    }
    async function refresh() {
      if (!bridge) {
        status('Recorder bridge not installed. Run the setup script and fully restart AnchorCast.', true);
        byId('sermonAudioStart').disabled = true;
        return;
      }
      if (refreshBusy) return;
      refreshBusy = true;
      try {
        const r = await bridge.info();
        if (!r?.success) throw Error(r?.error || 'Recorder unavailable');
        const selected = recorder.saved?.name || byId('sermonAudioFiles').value;
        recordings = r.recordings || [];
        const select = byId('sermonAudioFiles');
        select.replaceChildren();
        for (const row of recordings) {
          const option = document.createElement('option');
          option.value = row.name;
          option.textContent = `${row.interrupted ? 'Interrupted · ' : ''}${row.name} · ${(row.bytes / 1048576).toFixed(1)} MB`;
          select.appendChild(option);
        }
        if (!recordings.length) {
          const option = document.createElement('option');
          option.value = '';
          option.textContent = 'No saved recordings';
          select.appendChild(option);
        }
        if (recordings.some(row => row.name === selected)) select.value = selected;
        byId('sermonAudioCopy').disabled = !select.value || busy();
        byId('sermonAudioBrand').disabled = !select.value || busy();
        if (!r.available) {
          status(r.error, true);
          byId('sermonAudioStart').disabled = true;
        }
      } catch (e) { status(e.message, true); }
      finally { refreshBusy = false; }
    }

    function renderClosePrompt() {
      const state = recorder.state;
      const title = byId('sermonAudioCloseTitle');
      const message = byId('sermonAudioCloseMessage');
      const save = byId('sermonAudioCloseSave');
      const cancel = byId('sermonAudioCloseCancel');
      save.disabled = closePending;
      cancel.textContent = closePending ? 'Keep controls open' : 'Cancel';
      if (closePending) {
        title.textContent = state === 'starting' ? 'Cancelling microphone startup…' : 'Saving MP3 before closing…';
        message.textContent = 'The controls will close after this operation succeeds. Keep controls open cancels closing only; it does not restart or resume the recording.';
        save.textContent = state === 'starting' ? 'Cancelling…' : 'Saving…';
      } else if (closeProblem) {
        title.textContent = 'Recording needs attention';
        message.textContent = recorder.message || 'The recording could not be saved normally. Check Recordings for available audio.';
        save.textContent = 'Back to recorder';
        cancel.textContent = 'Keep controls open';
      } else if (state === 'starting') {
        title.textContent = 'Cancel microphone startup and close?';
        message.textContent = 'Microphone startup will be cancelled. To keep it starting in the background, choose Cancel, then Minimize (−).';
        save.textContent = 'Cancel startup & close';
      } else if (state === 'stopping') {
        title.textContent = 'MP3 is being saved';
        message.textContent = 'Choose Close when saved to wait for completion, or Cancel to keep the recorder controls open.';
        save.textContent = 'Close when saved';
      } else if (state === 'recording' || state === 'paused') {
        title.textContent = state === 'paused' ? 'Save paused recording and close?' : 'Stop recording and close?';
        message.textContent = 'The MP3 will be saved before these controls close. To keep the recording ' + (state === 'paused' ? 'paused' : 'running') + ' in the background, choose Cancel, then Minimize (−).';
        save.textContent = 'Stop & Save, then Close';
      } else {
        title.textContent = 'Close recorder controls?';
        message.textContent = recorder.message || 'No recording is running. Saved files will not be removed.';
        save.textContent = 'Close controls';
      }
    }

    function dismissClosePrompt() {
      // Invalidates an earlier asynchronous close so it cannot close a restored window.
      closeGeneration++;
      closePending = false;
      closeProblem = false;
      if (closeDialog.open) closeDialog.close('cancelled');
      if (dialog.open) byId('sermonAudioClose').focus();
    }
    function hideControls(reason) {
      dismissClosePrompt();
      minimized = reason === 'minimized';
      if (dialog.open) dialog.close(reason);
      launch.setAttribute('aria-expanded', 'false');
      recorder.emit();
      notice.textContent = minimized
        ? recorder.state === 'recording' ? 'Recorder minimized. Audio is still recording. Use the Audio button to restore controls.'
          : recorder.state === 'paused' ? 'Recorder minimized. Audio remains paused. Use the Audio button to restore controls.'
            : 'Recorder controls minimized. Use the Audio button to restore them.'
        : recorder.message || 'Recorder controls closed.';
      launch.focus();
    }
    function requestClose() {
      if (closeDialog.open) return;
      if (!busy()) { hideControls('closed'); return; }
      closeProblem = false;
      renderClosePrompt();
      closeDialog.showModal();
      byId('sermonAudioCloseCancel').focus();
    }
    async function confirmClose() {
      if (closePending) return;
      if (closeProblem) { dismissClosePrompt(); return; }
      if (!busy()) { hideControls('closed'); return; }
      const generation = ++closeGeneration;
      const cancellingStartup = recorder.state === 'starting';
      closePending = true;
      renderClosePrompt();
      // Stop() is idempotent and resolves only after encoder finalization/cleanup.
      try {
        await recorder.stop();
        if (generation !== closeGeneration) return;
        const safeSaved = recorder.state === 'saved' && recorder.saved && !recorder.saved.interrupted && !recorder.saved.error;
        const safeIdle = recorder.state === 'idle' &&
          (cancellingStartup || recorder.message === 'No audio was captured.');
        if (safeSaved || safeIdle) {
          if (safeSaved && window.RecordingBrandingUI) {
            if (closeDialog.open) closeDialog.close('saved');
            packaging = true;
            try { await window.RecordingBrandingUI.offer(recorder.saved.name); }
            finally { packaging = false; recorder.emit(); }
          }
          if (generation === closeGeneration) hideControls('closed');
          return;
        }
        closePending = false;
        closeProblem = true;
        renderClosePrompt();
        byId('sermonAudioCloseSave').focus();
      } catch (e) {
        if (generation !== closeGeneration) return;
        closePending = false;
        closeProblem = true;
        renderClosePrompt();
        byId('sermonAudioCloseMessage').textContent = e.message || 'Recording could not be saved. The controls remain open.';
        status(e.message, true);
      }
    }

    function update(s) {
      launch.classList.toggle('is-recording', s.state === 'recording');
      launch.classList.toggle('is-paused', s.state === 'paused');
      launch.dataset.recorderState = s.state;
      dialog.dataset.recorderState = s.state;
      const label = s.state === 'recording' ? `REC ${time(s.seconds)}`
        : s.state === 'paused' ? `Paused ${time(s.seconds)}`
        : s.state === 'stopping' ? 'Saving audio…'
        : s.state === 'starting' ? 'Starting audio…'
        : s.state === 'error' ? 'Audio — needs attention' : 'Record Audio';
      if (buttonText.textContent !== label) buttonText.textContent = label;
      launch.setAttribute('aria-busy', String(['starting', 'stopping'].includes(s.state)));
      const action = minimized ? 'Restore recorder controls.' : 'Open recorder controls.';
      launch.setAttribute('aria-label', s.state === 'recording' ? `Audio recording ${time(s.seconds)}. ${action}` : s.state === 'paused' ? `Audio paused at ${time(s.seconds)}. ${action}` : action);
      launch.title = s.state === 'recording' ? `Still recording ${time(s.seconds)} — ${action}` : s.state === 'paused' ? `Audio paused — ${action}` : action;
      const stateLabel = ({ idle: 'Ready', starting: 'Starting', recording: 'Recording', paused: 'Paused', stopping: 'Saving MP3', saved: 'Saved', error: 'Needs attention' })[s.state] || 'Needs attention';
      if (stateText.textContent !== stateLabel) stateText.textContent = stateLabel;
      // Announce state transitions, not every microphone sample or timer tick.
      if (lastState !== s.state) notice.textContent = s.state === 'recording'
        ? `Audio recording. ${minimized ? 'Recorder controls are minimized.' : ''}`
        : s.state === 'paused' ? 'Audio paused. No audio is being added to the file.'
        : s.message || `Audio recorder: ${stateLabel}.`;
      byId('sermonAudioElapsed').textContent = time(s.seconds);
      byId('sermonAudioLevel').value = s.peak || 0;
      byId('sermonAudioLevelText').textContent = s.state === 'recording' ? (s.peak > .98 ? 'Input is clipping — lower microphone gain' : s.peak > .003 ? 'Microphone signal received' : 'Low/no signal — check the microphone') : s.state === 'paused' ? 'Paused — audio is not added to the recording' : 'Input level appears when recording';
      byId('sermonAudioStart').disabled = busy() || !bridge;
      byId('sermonAudioPause').disabled = !['recording', 'paused'].includes(s.state);
      byId('sermonAudioPause').textContent = s.state === 'paused' ? 'Resume' : 'Pause';
      byId('sermonAudioStop').disabled = !['starting', 'recording', 'paused'].includes(s.state);
      byId('sermonAudioStop').textContent = s.state === 'starting' ? 'Cancel' : 'Stop & Save';
      byId('sermonAudioCopy').disabled = busy() || !byId('sermonAudioFiles').value;
      byId('sermonAudioBrand').disabled = busy() || !byId('sermonAudioFiles').value;
      status(s.message || 'Record only with the speaker’s knowledge and permission.', s.state === 'error' || !!s.saved?.interrupted);
      if (closeDialog.open) renderClosePrompt();
      if (lastState !== s.state && ['saved', 'idle', 'error'].includes(s.state)) {
        void refresh();
        if (minimized) notice.textContent = s.message || 'Audio recorder status changed. Open recording controls.';
      }
      lastState = s.state;
    }

    launch.onclick = () => {
      minimized = false;
      if (!dialog.open) dialog.showModal();
      launch.setAttribute('aria-expanded', 'true');
      recorder.emit();
      void refresh();
    };
    byId('sermonAudioMinimize').onclick = () => hideControls('minimized');
    byId('sermonAudioClose').onclick = requestClose;
    byId('sermonAudioCloseCancel').onclick = dismissClosePrompt;
    byId('sermonAudioCloseSave').onclick = () => void confirmClose();
    dialog.addEventListener('cancel', e => { e.preventDefault(); e.stopPropagation(); requestClose(); });
    closeDialog.addEventListener('cancel', e => { e.preventDefault(); e.stopPropagation(); dismissClosePrompt(); });
    dialog.addEventListener('close', () => {
      // A queued close event from Minimize must not disturb an immediate restoration.
      if (!dialog.open) { launch.setAttribute('aria-expanded', 'false'); launch.focus(); }
    });
    // Backdrop clicks deliberately do nothing. They are not a hidden recording action.
    byId('sermonAudioStart').onclick = () => void recorder.start();
    byId('sermonAudioPause').onclick = () => recorder.pause();
    byId('sermonAudioStop').onclick = async () => {
      if (packaging) return;
      packaging = true;
      try {
        await recorder.stop();
        if (recorder.state === 'saved' && recorder.saved && !recorder.saved.interrupted && !recorder.saved.error) {
          await window.RecordingBrandingUI?.offer(recorder.saved.name);
        }
      } catch (e) { status(e.message, true); }
      finally { packaging = false; recorder.emit(); }
    };
    byId('sermonAudioBrand').onclick = async () => {
      if (busy() || !byId('sermonAudioFiles').value) return;
      packaging = true; recorder.emit();
      try { await window.RecordingBrandingUI?.offer(byId('sermonAudioFiles').value); }
      catch (e) { status(e.message, true); }
      finally { packaging = false; recorder.emit(); }
    };
    byId('sermonAudioFiles').onchange = () => { byId('sermonAudioCopy').disabled = !byId('sermonAudioFiles').value || busy(); byId('sermonAudioBrand').disabled = byId('sermonAudioCopy').disabled; };
    byId('sermonAudioCopy').onclick = async () => {
      byId('sermonAudioCopy').disabled = true;
      try {
        const r = await bridge.saveCopy({ name: byId('sermonAudioFiles').value });
        if (!r.success) throw Error(r.error);
        if (!r.cancelled) status(`MP3 copy saved: ${r.path}`);
      } catch (e) { status(e.message, true); }
      finally { byId('sermonAudioCopy').disabled = busy() || !byId('sermonAudioFiles').value;
      byId('sermonAudioBrand').disabled = busy() || !byId('sermonAudioFiles').value; }
    };
    byId('sermonAudioFolder').onclick = async () => {
      if (!bridge) return;
      try { const r = await bridge.openFolder(); if (!r.success) status(r.error, true); }
      catch (e) { status(e.message, true); }
    };
    bridge?.onStopRequest?.(() => void recorder.stop());
    bridge?.onFault?.(message => {
      if (['starting', 'recording', 'paused'].includes(recorder.state)) void recorder.stop(message);
      else status(message, true);
    });
    window.addEventListener('beforeunload', () => {
      if (['recording', 'paused'].includes(recorder.state)) void recorder.stop('Window is closing.');
    });
    update({ state: 'idle', seconds: 0 });
  });
})();

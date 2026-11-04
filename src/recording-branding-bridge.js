'use strict';
// Privileged recording settings / generation / file dialogs are only exposed to the
// application's own main frame and its Settings frame. Never the projection/remote UI.
const fs = require('node:fs'), fsp = fs.promises, path = require('node:path'), crypto = require('node:crypto');
const { fileURLToPath } = require('node:url');
const { BrandingService } = require('./recording-branding');
function isTrusted(event, windows, rendererPort) {
  if (!event?.senderFrame || event.senderFrame !== event.sender?.mainFrame) return false;
  const match = windows.find(([win]) => win && !win.isDestroyed() && win.webContents === event.sender);
  if (!match) return false;
  try {
    const u = new URL(event.senderFrame.url), filename = match[1];
    return u.protocol === 'http:' ? u.hostname === '127.0.0.1' && u.port === String(rendererPort) && u.pathname === '/' + filename
      : u.protocol === 'file:' && path.resolve(fileURLToPath(u)) === path.join(__dirname, 'renderer', filename);
  } catch (_) { return false; }
}
function install({ app, ipcMain, dialog, shell, powerSaveBlocker, getMainWindow, getSettingsWindow, getRendererPort, getSettings, recordings, openRecordingSettings, service: injected }) {
  const service = injected || new BrandingService({ directory: path.join(app.getPath('userData'), 'AnchorCastData', 'AudioBranding'), recordings });
  const windows = () => [[getMainWindow(), 'index.html'], [getSettingsWindow(), 'settings.html']];
  const parent = e => windows().find(([w]) => w?.webContents === e.sender)?.[0] || getMainWindow();
  let power = null, quitting = false;
  function send(channel, data) { for (const [w] of windows()) if (w && !w.isDestroyed()) w.webContents.send(channel, data); }
  service.on('progress', info => {
    if (info.phase !== 'idle' && power === null) { try { power = powerSaveBlocker.start('prevent-app-suspension'); } catch (_) {} }
    if (info.phase === 'idle' && power !== null) { try { powerSaveBlocker.stop(power); } catch (_) {} power = null; }
    send('audio-branding:progress', info);
  });
  function handler(name, work, mainOnly = false) {
    ipcMain.handle('audio-branding:' + name, async (e, data) => {
      const allowed = mainOnly ? [[getMainWindow(), 'index.html']] : windows();
      if (!isTrusted(e, allowed, getRendererPort())) return { success: false, error: 'This action is available only in AnchorCast recording controls or Settings.' };
      try { return { success: true, ...(await work(data || {}, e)) }; }
      catch (error) { return { success: false, error: error.message || 'The audio operation could not be completed.' }; }
    });
  }
  const changed = async result => { send('audio-branding:changed', {}); return result; };
  handler('get', async () => {
    const s = await service.state(); const settings = getSettings() || {};
    return { ...s, keyAvailable: !!(process.env.ANCHORCAST_TTS_KEY || settings.openAiKey) };
  });
  handler('save', async data => changed(await service.save(data)));
  handler('generate', async data => changed(await service.generate(data, process.env.ANCHORCAST_TTS_KEY || (getSettings() || {}).openAiKey || '')));
  async function pick(e, title) {
    service.assertFree();
    const r = await dialog.showOpenDialog(parent(e), { title, properties: ['openFile'], filters: [{ name: 'Local audio (up to 3 minutes / 30 MB)', extensions: ['mp3', 'wav', 'm4a', 'aac', 'ogg', 'flac', 'aiff', 'aif'] }] });
    if (r.canceled || !r.filePaths?.length) return null;
    const file = r.filePaths[0]; if (!path.isAbsolute(file) || !/\.(mp3|wav|m4a|aac|ogg|flac|aiff|aif)$/i.test(file)) throw Error('Select a supported local audio file.');
    return file;
  }
  handler('import-voice', async (d, e) => {
    if (d.permission !== true) throw Error('Confirm permission to use this narration before importing it.');
    const input = await pick(e, 'Import ' + (d.slot === 'intro' ? 'introduction' : 'closing') + ' narration');
    return input ? changed(await service.importVoice({ slot: d.slot, revision: d.revision, input })) : { cancelled: true };
  });
  handler('import-music', async (d, e) => {
    if (d.permission !== true) throw Error('Confirm that you may use this music in distributed recordings.');
    const input = await pick(e, 'Import background music you have permission to use');
    return input ? changed(await service.importMusic({ revision: d.revision, input })) : { cancelled: true };
  });
  handler('prepare', async d => changed(await service.prepare(d)));
  handler('preview', d => service.preview(d));
  handler('preview-music', () => service.previewMusic());
  handler('cancel', d => service.cancel(d));
  handler('recording-info', d => service.recordingInfo(d), true);
  handler('release-recording-preview', d => service.releaseRecordingPreview(d), true);
  handler('compose', d => service.compose(d), true);
  handler('open-settings', async () => { openRecordingSettings(); return {}; }, true);
  handler('open-exports', async () => { await fsp.mkdir(service.exports, { recursive: true, mode: 0o700 }); const message = await shell.openPath(service.exports); if (message) throw Error(message); return {}; }, true);
  handler('save-export-as', async ({ name }, e) => {
    const src = await service.exportFile(name);
    const r = await dialog.showSaveDialog(parent(e), { title: 'Save finished sermon MP3', defaultPath: path.join(app.getPath('music'), name),
      filters: [{ name: 'MP3 audio', extensions: ['mp3'] }], properties: ['showOverwriteConfirmation', 'createDirectory'] });
    if (r.canceled || !r.filePath) return { cancelled: true };
    let dest = r.filePath; if (path.extname(dest).toLowerCase() !== '.mp3') dest += '.mp3';
    if (path.resolve(dest) === path.resolve(src)) return { path: src };
    // Never overwrite an original managed recording with a branded derivative.
    if (path.dirname(path.resolve(dest)) === path.resolve(recordings.directory)) throw Error('Choose a folder outside the original Recordings folder. Original recordings are preserved.');
    if (dest !== r.filePath && fs.existsSync(dest)) {
      const confirm = await dialog.showMessageBox(parent(e), { type: 'question', message: 'Replace this existing MP3?', detail: dest, buttons: ['Cancel', 'Replace'], defaultId: 0, cancelId: 0 });
      if (confirm.response !== 1) return { cancelled: true };
    }
    const temp = dest + '.' + crypto.randomBytes(8).toString('hex') + '.tmp';
    try { await fsp.copyFile(src, temp, fs.constants.COPYFILE_EXCL); await fsp.rename(temp, dest); }
    finally { await fsp.unlink(temp).catch(() => {}); }
    return { path: dest };
  }, true);
  function guardExit(e, kind) {
    if (!service.job || quitting || recordings.busy || e.defaultPrevented) return;
    e.preventDefault(); const w = getMainWindow();
    if (!w || w.isDestroyed()) { void service.cancel(); return; }
    const response = dialog.showMessageBoxSync(w, { type: 'question', title: 'Audio preparation is running', message: 'Cancel audio preparation and close?',
      detail: 'The original sermon recording and previously prepared clips are retained. Any unfinished export will be removed.', buttons: ['Keep AnchorCast open', 'Cancel operation and close'], defaultId: 0, cancelId: 0 });
    if (response === 1) { quitting = true; void service.cancel().finally(() => {
      quitting = false;
      if (kind === 'quit') app.quit(); else if (!w.isDestroyed()) w.close();
    }); }
  }
  app.prependListener('before-quit', e => guardExit(e, 'quit'));
  app.on('browser-window-created', (_event, win) => {
    win.on('close', e => { if (win === getMainWindow()) guardExit(e, 'window'); });
  });
  app.on('will-quit', () => { void service.dispose?.(); });
  return service;
}
module.exports = { install, isTrusted };

// AnchorCast Timer — Standalone Entry Point
// Runs as a completely independent Electron app.
// Opens the Timer control window + its own projection screen.

'use strict';

const { app, BrowserWindow, ipcMain, screen, Menu, shell } = require('electron');
const path = require('path');
const fs   = require('fs');
const os   = require('os');

// Force userData to AnchorCast (capital) to match main app and NSIS paths
if (process.platform === 'win32') {
  const _userData = path.join(
    process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'),
    'AnchorCast'
  );
  try { app.setPath('userData', _userData); } catch(_) {}
}

const APP_ICON = path.join(__dirname, '..', 'assets', 'icon.ico');
const IS_WIN   = process.platform === 'win32';

let timerWin      = null;
let projWin       = null;
let _activeTimer  = null;

// ── Single instance lock ──────────────────────────────────────────────────────
if (!app.requestSingleInstanceLock()) { app.quit(); }
app.on('second-instance', () => {
  if (timerWin) { timerWin.isMinimized() && timerWin.restore(); timerWin.focus(); }
});

// ── Create timer control window ───────────────────────────────────────────────
function createTimerWindow() {
  timerWin = new BrowserWindow({
    icon: APP_ICON,
    width: 1060, height: 820,
    minWidth: 860, minHeight: 680,
    title: 'AnchorCast — Timer',
    backgroundColor: '#07080f',
    autoHideMenuBar: true,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });
  timerWin.loadFile(path.join(__dirname, 'renderer', 'countdown-window.html'));
  timerWin.once('ready-to-show', () => timerWin.show());
  timerWin.on('closed', () => {
    timerWin = null;
    if (projWin && !projWin.isDestroyed()) { projWin.close(); projWin = null; }
    app.quit();
  });
  buildMenu();
}

// ── Create standalone projection window ───────────────────────────────────────
function createStandaloneProjection() {
  if (projWin && !projWin.isDestroyed()) { projWin.focus(); return; }
  const displays = screen.getAllDisplays();
  const primary  = screen.getPrimaryDisplay();
  const target   = displays.find(d => d.id !== primary.id) || primary;
  const { x, y, width, height } = target.bounds;
  const isSame = target.id === primary.id;

  projWin = new BrowserWindow({
    x, y, width, height,
    fullscreen: true, frame: false,
    alwaysOnTop: !isSame,
    skipTaskbar: !isSame,
    backgroundColor: '#000000',
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      backgroundThrottling: false,
    },
  });
  projWin.loadFile(path.join(__dirname, 'renderer', 'projection.html'));
  projWin.once('ready-to-show', () => {
    projWin.show();
    if (!isSame && IS_WIN) projWin.setAlwaysOnTop(true, 'screen-saver');
  });
  projWin.webContents.on('did-finish-load', () => {
    // Sync current timer state to new projection window
    if (_activeTimer) {
      const elapsed = (Date.now() - _activeTimer.startedAt) / 1000;
      const remaining = _activeTimer.mode === 'countdown'
        ? Math.max(0, _activeTimer.seconds - elapsed)
        : elapsed;
      if (remaining > 0) {
        projWin.webContents.send('show-timer', {
          ..._activeTimer.data,
          seconds: Math.ceil(remaining),
        });
      }
    }
    // Sync clock state if active
    if (_activeClockData) {
      projWin.webContents.send('show-clock', _activeClockData);
    }
  });
  projWin.on('closed', () => { projWin = null; });
}

// ── IPC handlers ──────────────────────────────────────────────────────────────
ipcMain.handle('show-timer', (_, data) => {
  _activeTimer = {
    data,
    startedAt: Date.now(),
    seconds:   data.seconds || 0,
    mode:      data.mode    || 'countdown',
    label:     data.label   || '',
    scale:     data.scale   || 1,
    position:  data.position || 'edge',
  };
  if (!projWin || projWin.isDestroyed()) createStandaloneProjection();
  else projWin.webContents.send('show-timer', data);

  // Sync to timer control window
  const syncPayload = { ..._activeTimer, running: true, remaining: _activeTimer.seconds };
  timerWin?.webContents.send('timer-state-sync', syncPayload);
  return { success: true };
});

ipcMain.handle('stop-timer', () => {
  _activeTimer = null;
  projWin?.webContents.send('stop-timer');
  timerWin?.webContents.send('timer-stopped');
  return { success: true };
});

ipcMain.handle('get-timer-state', () => {
  if (!_activeTimer) return { running: false };
  const elapsed = (Date.now() - _activeTimer.startedAt) / 1000;
  const remaining = _activeTimer.mode === 'countdown'
    ? Math.max(0, _activeTimer.seconds - elapsed)
    : elapsed;
  if (_activeTimer.mode === 'countdown' && remaining <= 0) {
    _activeTimer = null;
    return { running: false };
  }
  return {
    running: true, data: _activeTimer.data,
    mode: _activeTimer.mode, label: _activeTimer.label,
    scale: _activeTimer.scale, position: _activeTimer.position,
    seconds: _activeTimer.seconds, remaining: Math.ceil(remaining),
  };
});

ipcMain.handle('timer-scale', (_, data) => {
  if (_activeTimer) { _activeTimer.scale = data.scale; _activeTimer.position = data.position; }
  projWin?.webContents.send('timer-scale', data);
  return { success: true };
});

ipcMain.handle('timer-flash-speed', (_, data) => {
  projWin?.webContents.send('timer-flash-speed', data);
  return { success: true };
});

ipcMain.handle('set-projection-bg', (_, data) => {
  projWin?.webContents.send('set-projection-bg', data);
  return { success: true };
});

// File picker for image/video background
ipcMain.handle('pick-bg-file', async (_, type) => {
  const filters = type === 'image'
    ? [{ name: 'Images', extensions: ['jpg','jpeg','png','webp','gif'] }]
    : [{ name: 'Videos', extensions: ['mp4','webm','mov','mkv'] }];
  const { dialog } = require('electron');
  const res = await dialog.showOpenDialog(timerWin, { filters, properties: ['openFile'] });
  if (res.canceled || !res.filePaths.length) return null;
  return res.filePaths[0];
});

ipcMain.handle('open-projection', () => { createStandaloneProjection(); return { success: true }; });
ipcMain.handle('close-projection', () => {
  if (projWin && !projWin.isDestroyed()) { projWin.close(); projWin = null; }
  return { success: true };
});

// Tell countdown-window.html it's running in standalone mode
ipcMain.handle('is-timer-standalone', () => true);

// Clock / Date display on projection
let _activeClockData = null;
ipcMain.handle('show-clock', (_, data) => {
  _activeClockData = data;
  projWin?.webContents.send('show-clock', data);
  return { success: true };
});
ipcMain.handle('hide-clock', () => {
  _activeClockData = null;
  projWin?.webContents.send('hide-clock');
  return { success: true };
});
ipcMain.handle('get-clock-state', () => _activeClockData || null);

// Stub out handlers the countdown-window.html may call but aren't needed standalone
const _noop = () => ({});
[
  'get-settings','save-settings','get-current-render-state',
  'get-remote-info','get-themes','get-presets',
].forEach(ch => ipcMain.handle(ch, _noop));

// ── Menu ──────────────────────────────────────────────────────────────────────
function buildMenu() {
  const tpl = [
    { label: 'File', submenu: [
      { label: 'Open AnchorCast', click: () => {
        const exe = app.getPath('exe');
        const dir = path.dirname(exe);
        const main = path.join(dir, 'AnchorCast.exe');
        if (fs.existsSync(main)) shell.openPath(main);
      }},
      { type: 'separator' },
      { label: 'Quit Timer', accelerator: 'CmdOrCtrl+Q', click: () => app.quit() },
    ]},
    { label: 'Display', submenu: [
      { label: 'Open Projection Screen', click: createStandaloneProjection },
      { label: 'Close Projection Screen', click: () => {
        if (projWin && !projWin.isDestroyed()) { projWin.close(); projWin = null; }
      }},
    ]},
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(tpl));
}

// ── App ready ─────────────────────────────────────────────────────────────────
app.whenReady().then(createTimerWindow);
app.on('window-all-closed', () => app.quit());

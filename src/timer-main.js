// AnchorCast Timer — Standalone Entry Point
// Runs as a completely independent Electron app.
// Opens the Timer control window + its own projection screen.

'use strict';

const { app, BrowserWindow, ipcMain, screen, Menu, shell } = require('electron');
const path = require('path');
const fs   = require('fs');
const os   = require('os');
const TimerUtils = require('./renderer/js/timer-utils');

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
    // Reconstruct from the same start timestamp; never restart or discard overtime.
    if (_activeTimer) projWin.webContents.send('show-timer', TimerUtils.payload(_activeTimer));
    // Sync clock state if active
    if (_activeClockData) {
      projWin.webContents.send('show-clock', _activeClockData);
    }
  });
  projWin.on('closed', () => { projWin = null; });
}

// ── IPC handlers ──────────────────────────────────────────────────────────────
ipcMain.handle('show-timer', (_, data) => {
  _activeTimer = TimerUtils.createState(data);
  if (!projWin || projWin.isDestroyed()) createStandaloneProjection();
  else projWin.webContents.send('show-timer', TimerUtils.payload(_activeTimer));
  const state = TimerUtils.snapshot(_activeTimer);
  timerWin?.webContents.send('timer-state-sync', state);
  return { success: true, projected: true, state };
});

ipcMain.handle('stop-timer', () => {
  _activeTimer = null;
  projWin?.webContents.send('stop-timer');
  timerWin?.webContents.send('timer-stopped');
  return { success: true };
});

ipcMain.handle('get-timer-state', () => TimerUtils.snapshot(_activeTimer));

// Hide/show only the timer overlay. Do not stop, recreate or rebase the timer,
// clear the current slide, touch the clock, or open/close a projection window.
ipcMain.handle('set-timer-projection-visible', (_, data) => {
  if (typeof data?.visible !== 'boolean') {
    return { success: false, error: 'Timer visibility must be true or false.' };
  }
  if (!_activeTimer) {
    return { success: false, error: 'Start a timer before changing its projection visibility.', state: { running: false } };
  }
  // A delayed click from an older timer must not hide a newly started session.
  if (data.expectedStartedAt !== undefined && data.expectedStartedAt !== _activeTimer.startedAt) {
    return { success: false, error: 'The running timer changed. Please try again.', state: TimerUtils.snapshot(_activeTimer) };
  }
  _activeTimer = TimerUtils.setProjectionVisibility(_activeTimer, data.visible);
  const state = TimerUtils.snapshot(_activeTimer);
  for (const win of [projWin, timerWin]) {
    if (win && !win.isDestroyed()) win.webContents.send('timer-visibility', state);
  }
  return { success: true, projectionVisible: state.projectionVisible, state };
});

ipcMain.handle('timer-scale', (_, data) => {
  if (_activeTimer) { _activeTimer.scale = data.scale; _activeTimer.position = data.position; }
  projWin?.webContents.send('timer-scale', data);
  return { success: true };
});

ipcMain.handle('timer-flash-speed', (_, data) => {
  const speed = TimerUtils.flashSpeed(data?.speed);
  if (_activeTimer) _activeTimer.data.flashSpeed = speed;
  projWin?.webContents.send('timer-flash-speed', { speed });
  timerWin?.webContents.send('timer-flash-speed', { speed });
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

// AnchorCast — Main Process v2
// HTTP remote control · NDI hooks · sermon history · theme designer window
const {app,BrowserWindow,ipcMain,screen,dialog,Menu,shell,protocol,net,clipboard,powerSaveBlocker}=require('electron');
const path=require('path');
const os=require('os');

// ── Force userData to use 'AnchorCast' (capital A+C) not 'anchorcast' ────────
// Electron derives userData from package.json "name" which is lowercase.
// We override it here before any path is read so all data goes to:
// AppData\Roaming\AnchorCast\ (matching what users expect and what NSIS writes to)
if (!app.isPackaged || process.platform === 'win32') {
  const _userData = path.join(
    process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'),
    'AnchorCast'
  );
  try { app.setPath('userData', _userData); } catch(_) {}
}

// ── Timer standalone mode ─────────────────────────────────────────────────────
// If launched with --timer flag, load timer-main.js and stop here.
// This is how the "AnchorCast Timer" shortcut works — same exe, different mode.
if (process.argv.includes('--timer')) {
  module.exports = require('./timer-main.js');
} else {
const crypto=require('crypto');
// BUG-A FIX: fs/http/https/spawn required here — BEFORE loadRegistration() at line 49 which uses fs
const fs=require('fs');
const http=require('http');
const https=require('https');
const {spawn,spawnSync,execFileSync}=require('child_process');


// ─── REGISTRATION SYSTEM ─────────────────────────────────────────────────────
// Free & open source — registration identifies users for support
// Flow: first launch → registration form → email with magic link → auto-activate
// Token = HMAC-SHA256(secret, email + '|' + hwId) — server-free, unforgeable

const REG_SECRET   = 'anchorcast-registration-2025-hmac-token';
const REG_FILE     = path.join(app.getPath('userData'), 'registration.json');

// Gmail SMTP config
const SMTP_HOST    = 'smtp.gmail.com';
const SMTP_PORT    = 587;
const SMTP_USER    = 'REPLACE_WITH_EMAIL_ADDRESS';
const SMTP_PASS    = 'REPLACE_WITH_APP_PASSWORD';  // Gmail App Password (16 chars)
const FROM_EMAIL   = 'donotreply@anchorcastapp.com';
const APP_PROTOCOL = 'anchorcast';

function getHardwareId() {
  const raw = `${os.hostname()}|${(os.cpus()[0]?.model||'cpu')}|${os.platform()}|${os.arch()}`;
  return 'GP-' + crypto.createHash('sha256').update(raw).digest('hex').toUpperCase().slice(0,4)
    + '-' + crypto.createHash('sha256').update(raw+'1').digest('hex').toUpperCase().slice(0,4)
    + '-' + crypto.createHash('sha256').update(raw+'2').digest('hex').toUpperCase().slice(0,4)
    + '-' + crypto.createHash('sha256').update(raw+'3').digest('hex').toUpperCase().slice(0,4);
}

function generateRegToken(email, hwId) {
  return crypto.createHmac('sha256', REG_SECRET)
    .update(email.toLowerCase().trim() + '|' + hwId.trim())
    .digest('hex');
}

function verifyRegToken(token, email, hwId) {
  const expected = generateRegToken(email, hwId);
  try {
    return crypto.timingSafeEqual(
      Buffer.from(token, 'hex'),
      Buffer.from(expected, 'hex')
    );
  } catch(_) { return false; }
}

function loadRegistration() {
  try {
    if (fs.existsSync(REG_FILE)) return JSON.parse(fs.readFileSync(REG_FILE, 'utf8'));
  } catch(_) {}
  return null;
}

function saveRegistration(data) {
  try {
    fs.mkdirSync(path.dirname(REG_FILE), { recursive: true });
    fs.writeFileSync(REG_FILE, JSON.stringify(data, null, 2));
  } catch(_) {}
}

function getRegistrationStatus() {
  const hwId = getHardwareId();
  const data = loadRegistration();
  if (data?.registered && data?.token && data?.email) {
    if (verifyRegToken(data.token, data.email, hwId)) {
      return {
        registered: true, hwId,
        fullName:    data.fullName || '',
        email:       data.email,
        registeredAt: data.registeredAt || null,
        churchName:  data.churchName || '',
      };
    }
  }
  return { registered: false, hwId };
}

function getEmailSentState() {
  try {
    const data = loadRegistration();
    const email = String(data?.pendingEmail || '');
    // Validate it's actually an email string (not a boolean/garbage value)
    if (data?.emailSent && !data?.registered && email && email.includes('@')) {
      return { sent: true, email, fullName: data.pendingName || '', church: data.pendingChurch || '' };
    }
    // Bad/corrupted state — clear it
    if (data?.emailSent && data?.pendingEmail && !String(data.pendingEmail).includes('@')) {
      data.emailSent = false;
      data.pendingEmail = null;
      saveRegistration(data);
    }
  } catch(_) {}
  return { sent: false };
}

function saveEmailSentState(fullName, email, church) {
  // Only save if email is a valid string
  if (!email || !String(email).includes('@')) return;
  const data = loadRegistration() || {};
  data.emailSent     = true;
  data.pendingEmail  = String(email).toLowerCase().trim();
  data.pendingName   = String(fullName || '').trim();
  data.pendingChurch = String(church || '').trim();
  data.emailSentAt   = Date.now();
  saveRegistration(data);
}


function isRegistered() {
  return getRegistrationStatus().registered === true;
}

function activateRegistration(fullName, email, token, churchName) {
  const hwId = getHardwareId();
  if (!verifyRegToken(token, email, hwId)) {
    return { success: false, error: 'Invalid or expired registration link.' };
  }
  const existing = loadRegistration() || {};
  const data = {
    ...existing,
    registered:   true,
    emailSent:    false,  // clear pending state
    pendingEmail: null,
    token,
    fullName:     String(fullName || '').trim(),
    email:        String(email || '').toLowerCase().trim(),
    churchName:   String(churchName || '').trim(),
    hwId,
    registeredAt: Date.now(),
  };
  saveRegistration(data);
  return { success: true, fullName: data.fullName, email: data.email };
}

// ── SMTP email sender (native Node — no nodemailer needed) ───────────────────
function sendSmtpEmail({ to, subject, html, text }) {
  return new Promise((resolve, reject) => {
    const tls  = require('tls');
    const net  = require('net');
    const CRLF = '\r\n';

    function b64(s) { return Buffer.from(s).toString('base64'); }

    const authStr = b64(`\0${SMTP_USER}\0${SMTP_PASS}`);
    const boundary = 'acbndry' + Date.now().toString(36);

    const msgLines = [
      `From: AnchorCast <${FROM_EMAIL}>`,
      `To: ${to}`,
      `Subject: ${subject}`,
      `MIME-Version: 1.0`,
      `Content-Type: multipart/alternative; boundary="${boundary}"`,
      ``,
      `--${boundary}`,
      `Content-Type: text/plain; charset=utf-8`,
      ``,
      text || subject,
      ``,
      `--${boundary}`,
      `Content-Type: text/html; charset=utf-8`,
      ``,
      html,
      ``,
      `--${boundary}--`,
      ``,
    ].join(CRLF);

    // SMTP state machine
    // States: banner → ehlo → starttls → ehlo2 → auth → mailfrom → rcptto → data → body → quit → done
    let sock;
    let rawBuf  = '';   // raw bytes buffer
    let state   = 'banner';
    let tlsSock = null;
    let done    = false;

    function writeSock(s) { (tlsSock || sock).write(s); }

    // SMTP multi-line response: lines ending with "XYZ-..." are continuations,
    // "XYZ ..." or "XYZ\r\n" is the final line of the response.
    // We only act when we receive the FINAL line.
    function isFinalLine(line) {
      // Final lines: 3-digit code followed by space (or end), NOT a dash
      return /^\d{3}[ \r\n]/.test(line) || /^\d{3}$/.test(line.trim());
    }

    function getCode(line) {
      return line.slice(0, 3);
    }

    function processLine(line) {
      if (done) return;
      if (!line.trim()) return;

      // Skip multi-line continuation lines (e.g. "250-SIZE 35882577")
      if (/^\d{3}-/.test(line)) return;

      const code = getCode(line);
      console.log(`[SMTP] state=${state} code=${code} line=${line.slice(0,60)}`);

      switch (state) {
        case 'banner':
          if (code !== '220') return fail('Expected 220 banner, got: ' + line);
          state = 'ehlo';
          writeSock(`EHLO anchorcastapp.com${CRLF}`);
          break;

        case 'ehlo':
          if (code !== '250') return fail('EHLO failed: ' + line);
          state = 'starttls';
          writeSock(`STARTTLS${CRLF}`);
          break;

        case 'starttls':
          if (code !== '220') return fail('STARTTLS failed: ' + line);
          // Upgrade to TLS
          state = 'ehlo2';
          tlsSock = tls.connect({ socket: sock, servername: SMTP_HOST }, () => {
            tlsSock.on('data', d => onData(d));
            writeSock(`EHLO anchorcastapp.com${CRLF}`);
          });
          tlsSock.on('error', err => fail(err.message));
          break;

        case 'ehlo2':
          if (code !== '250') return fail('EHLO2 failed: ' + line);
          state = 'auth';
          writeSock(`AUTH PLAIN ${authStr}${CRLF}`);
          break;

        case 'auth':
          if (code !== '235') return fail('AUTH failed: ' + line + ' — check Gmail App Password');
          state = 'mailfrom';
          writeSock(`MAIL FROM:<${FROM_EMAIL}>${CRLF}`);
          break;

        case 'mailfrom':
          if (code !== '250') return fail('MAIL FROM failed: ' + line);
          state = 'rcptto';
          writeSock(`RCPT TO:<${to}>${CRLF}`);
          break;

        case 'rcptto':
          if (code !== '250') return fail('RCPT TO failed: ' + line);
          state = 'data';
          writeSock(`DATA${CRLF}`);
          break;

        case 'data':
          if (code !== '354') return fail('DATA failed: ' + line);
          state = 'body';
          writeSock(msgLines + `.${CRLF}`);
          break;

        case 'body':
          if (code !== '250') return fail('Message rejected: ' + line);
          state = 'quit';
          writeSock(`QUIT${CRLF}`);
          break;

        case 'quit':
          // 221 = bye, anything = still ok (email was sent)
          done = true;
          try { (tlsSock||sock).destroy(); } catch(_) {}
          resolve({ ok: true });
          break;
      }
    }

    function onData(data) {
      rawBuf += data.toString();
      const lines = rawBuf.split(/\r?\n/);
      rawBuf = lines.pop(); // keep incomplete line in buffer
      for (const line of lines) {
        try { processLine(line); } catch(e) { fail(e.message); }
        if (done) break;
      }
    }

    function fail(msg) {
      if (done) return;
      done = true;
      try { (tlsSock||sock).destroy(); } catch(_) {}
      reject(new Error(msg));
    }

    sock = net.connect(SMTP_PORT, SMTP_HOST, () => {});
    sock.on('data', d => { if (!tlsSock) onData(d); }); // before TLS upgrade
    sock.on('error', err => fail(err.message));
    sock.setTimeout(20000, () => fail('SMTP connection timed out'));
  });
}

async function sendRegistrationEmail(fullName, email, churchName) {
  const hwId  = getHardwareId();
  const token = generateRegToken(email, hwId);
  const name  = encodeURIComponent(fullName);
  const mail  = encodeURIComponent(email);
  const church = encodeURIComponent(churchName || '');
  const link  = `${APP_PROTOCOL}://register?token=${token}&name=${name}&email=${mail}&church=${church}&hwId=${encodeURIComponent(hwId)}`;

  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:-apple-system,sans-serif;background:#f5f5f5;padding:30px;margin:0">
<div style="max-width:520px;margin:0 auto;background:#fff;border-radius:12px;
  box-shadow:0 2px 16px rgba(0,0,0,.1);overflow:hidden">
  <div style="background:linear-gradient(135deg,#0a0a1a,#1a1a3a);padding:28px 30px;text-align:center">
    <div style="font-size:28px;margin-bottom:8px">⚓</div>
    <div style="color:#c9a84c;font-size:20px;font-weight:700;letter-spacing:1px">ANCHORCAST</div>
    <div style="color:#888;font-size:11px;letter-spacing:2px;margin-top:4px">LIVE SERMON DISPLAY</div>
  </div>
  <div style="padding:30px">
    <h2 style="color:#1a1a2e;font-size:18px;margin:0 0 12px">Hi ${fullName},</h2>
    <p style="color:#444;font-size:14px;line-height:1.7;margin:0 0 20px">
      Thank you for registering AnchorCast${churchName ? ' at <strong>' + churchName + '</strong>' : ''}!
      Click the button below to complete your registration and unlock the app.
    </p>
    <div style="text-align:center;margin:28px 0">
      <a href="${link}"
        style="display:inline-block;background:linear-gradient(135deg,#c9a84c,#e0bf6a);
          color:#000;font-size:15px;font-weight:700;text-decoration:none;
          padding:14px 36px;border-radius:8px;letter-spacing:.3px">
        ✅ Complete Registration
      </a>
    </div>
    <p style="color:#888;font-size:12px;line-height:1.6;margin:0 0 16px">
      Or copy and paste this link into your browser if the button doesn't work:
    </p>
    <div style="background:#f5f5f5;border-radius:6px;padding:10px 12px;
      font-family:monospace;font-size:11px;color:#555;word-break:break-all;margin-bottom:20px">
      ${link}
    </div>
    <hr style="border:none;border-top:1px solid #eee;margin:20px 0">
    <div style="color:#aaa;font-size:11px;text-align:center;line-height:1.6">
      Hardware ID: <strong style="color:#666">${hwId}</strong><br>
      This link is specific to this device and cannot be shared.<br>
      AnchorCast is free &amp; open source —
      <a href="https://github.com/anchorcastapp-team/anchorcastapp" style="color:#c9a84c">
        github.com/anchorcastapp-team/anchorcastapp
      </a>
    </div>
  </div>
</div>
</body></html>`;

  const text = `Hi ${fullName},\n\nThank you for registering AnchorCast${churchName ? ' at ' + churchName : ''}!\n\nClick the link below to complete your registration:\n\n${link}\n\nHardware ID: ${hwId}\n\nAnchorCast is free & open source: https://github.com/anchorcastapp-team/anchorcastapp`;

  await sendSmtpEmail({ to: email, subject: 'Complete your AnchorCast Registration', html, text });
  return { success: true, token };
}

function requiresRegistration(featureName) {
  if (isRegistered()) return null;
  return { blocked: true, reason: `${featureName} requires registration. Please register AnchorCast to continue.` };
}
// ─── END REGISTRATION SYSTEM ──────────────────────────────────────────────────
// (spawnSync required at top — BUG-A fix)
// (spawn required at top — BUG-A fix)

function nodeToWeb(nodeStream) {
  return new ReadableStream({
    start(ctrl) {
      nodeStream.on('data',  chunk => ctrl.enqueue(chunk));
      nodeStream.on('end',   ()    => ctrl.close());
      nodeStream.on('error', err   => ctrl.error(err));
    },
    cancel() { nodeStream.destroy(); }
  });
}

const STOPWORDS = new Set(['the','and','for','that','with','this','have','from','your','will','shall','unto','into','about','there','their','them','they','were','been','being','than','then','what','when','where','which','while','would','could','should','because','through','after','before','those','these','also','upon','very','more','most','some','such','only','just','over','under','again','said','says','say','lord','jesus','christ','god','bible','verse','chapter','book','turn','with','unto','from','into','ours','ourselves','yourself','yourselves','pastor']);

// ── Register media:// protocol ────────────────────────────────────────────────
// Allows renderer (running on localhost) to load local media files safely
// Usage: media:///C:/path/to/file.mp4  or  media:///home/user/file.jpg
protocol.registerSchemesAsPrivileged([
  { scheme: 'media', privileges: { secure: true, supportFetchAPI: true, bypassCSP: true, stream: true } }
]);

// ── Critical: Enable microphone and speech recognition in Electron/Chromium ──
// These MUST be set before app.whenReady() — they configure the Chromium engine
app.commandLine.appendSwitch('enable-speech-dispatcher');
app.commandLine.appendSwitch('use-fake-ui-for-media-stream', '0');
app.commandLine.appendSwitch('enable-usermedia-screen-capturing');
app.commandLine.appendSwitch('allow-http-screen-capture');
app.commandLine.appendSwitch('disable-features', 'HardwareMediaKeyHandling,MediaSessionService');
// BUG-E FIX: Chromium only honours the LAST --enable-features switch.
// Combined into one call so both Speech API and HEVC decoder are active.
app.commandLine.appendSwitch('enable-features', 'WebSpeechAPI,SpeechRecognition,PlatformHEVCDecoderSupport');
app.commandLine.appendSwitch('autoplay-policy', 'no-user-gesture-required');

// ── Suppress grandiose's internal "Frame number X sent." spam ────────────────
// grandiose's native C++ code writes directly to stdout — intercept and filter it
const _origWrite = process.stdout.write.bind(process.stdout);
process.stdout.write = (chunk, encoding, cb) => {
  const s = typeof chunk === 'string' ? chunk : chunk.toString();
  if (/^Frame number \d+ sent\./.test(s.trim())) return true; // suppress
  return _origWrite(chunk, encoding, cb);
};

let mainWindow=null, projectionWindow=null, historyWindow=null, themeWindow=null, splashWindow=null, countdownWindow=null;
let splashShownAt = 0;
const MIN_SPLASH_MS = 2400;
let httpServer=null;
let currentLiveVerse=null;
let currentBackgroundMedia=null;
let currentSettings={ whisperSource:'local', onlineMode:false };
let currentRenderState={ module:'clear', payload:null, updatedAt:0, backgroundMedia:null };
let currentLogoOverlayState = null;
let pendingScheduleOpenPath = null;

function buildRenderState(module, payload = null, extra = {}){
  return { module, payload, updatedAt: Date.now(), ...extra };
}
function pushRenderStateToProjection(state){
  currentRenderState = state || buildRenderState('clear');
  if(projectionWindow && !projectionWindow.isDestroyed()){
    projectionWindow.webContents.send('render-state', currentRenderState);
  }
  return currentRenderState;
}
function projectionWindowReadySync(){
  if(projectionWindow && !projectionWindow.isDestroyed()){
    projectionWindow.webContents.send('render-state', currentRenderState);
    if (currentLogoOverlayState) {
      projectionWindow.webContents.send('logo-overlay', currentLogoOverlayState);
    }
    // If a timer was started before projection opened, resume it with correct remaining time
    if (_activeTimer) {
      const elapsed = (Date.now() - _activeTimer.startedAt) / 1000;
      const remaining = _activeTimer.mode === 'countdown'
        ? Math.max(0, _activeTimer.seconds - elapsed)
        : elapsed;
      if (_activeTimer.mode === 'countdown' && remaining <= 0) {
        // Timer already finished — clear it
        _activeTimer = null;
      } else {
        // Send original timer baseline to projection so it reconstructs the same live timer
        projectionWindow.webContents.send('show-timer', {
          ..._activeTimer.data,
          seconds: _activeTimer.seconds,
          startedAt: _activeTimer.startedAt,
          mode: _activeTimer.mode,
          label: _activeTimer.label || _activeTimer.data?.label || '',
          position: _activeTimer.position || _activeTimer.data?.position || 'edge',
          scale: _activeTimer.scale ?? _activeTimer.data?.scale ?? 1,
          serverNow: Date.now(),
        });
      }
    }
    // Sync clock state if clock was active
    if (_activeClockData) {
      projectionWindow.webContents.send('show-clock', _activeClockData);
    }
  }
}


function _looksLikeSchedulePath(filePath){
  const p = String(filePath || '').trim();
  if (!p) return false;
  return /\.(acsch|json)$/i.test(p);
}
function _queueScheduleOpen(filePath){
  const p = String(filePath || '').trim();
  if (!p || !_looksLikeSchedulePath(p)) return false;
  pendingScheduleOpenPath = path.resolve(p);
  return true;
}
function _consumePendingScheduleOpen(){
  if (!pendingScheduleOpenPath || !mainWindow || mainWindow.isDestroyed()) return;
  const filePath = pendingScheduleOpenPath;
  const clearPending = () => { pendingScheduleOpenPath = null; };
  try {
    const sendNow = () => {
      mainWindow?.webContents.send('menu-schedule-load-file', filePath);
      clearPending();
    };
    if (mainWindow.webContents.isLoading()) mainWindow.webContents.once('did-finish-load', sendNow);
    else sendNow();
  } catch (_) {}
}
function _extractScheduleArg(argv = []){
  const args = Array.isArray(argv) ? argv.slice(1) : [];
  for (const arg of args){
    if (_looksLikeSchedulePath(arg) && fs.existsSync(arg)) return path.resolve(arg);
  }
  return null;
}

const isDev=process.argv.includes('--dev');
const UD=app.getPath('userData');
// ── Persistent user content lives in AppData / userData ───────────────────────
// This avoids accidental loss when the app folder is replaced or cleaned.
// assets/ is kept only as a legacy migration source / bundled resource location.
const ASSETS_DIR = app.isPackaged
  ? path.join(process.resourcesPath, 'assets')
  : path.join(app.getAppPath(), 'assets');
const APPDATA_ROOT       = path.join(UD, 'AnchorCastData');
const DATA_DIR           = path.join(APPDATA_ROOT, 'Data');
const BIBLE_DIR          = path.join(APPDATA_ROOT, 'Bibles');
const PRES_ASSETS        = path.join(APPDATA_ROOT, 'Presentation');
const VIDEO_ASSETS       = path.join(APPDATA_ROOT, 'videos');
const AUDIO_ASSETS       = path.join(APPDATA_ROOT, 'Audio');
const IMAGE_ASSETS       = path.join(APPDATA_ROOT, 'Images');
const SCHEDULES_DIR      = path.join(APPDATA_ROOT, 'Schedules');
const WHISPER_MODEL_DIR  = path.join(APPDATA_ROOT, 'WhisperModels'); // persists across app updates

// All supported Whisper models with metadata
const WHISPER_MODELS = [
  { id: 'tiny.en',  name: 'Tiny English',       size: '~39 MB',  tag: 'Fastest',  hfRepo: 'Systran/faster-whisper-tiny.en',  folder: 'models--Systran--faster-whisper-tiny.en'  },
  { id: 'base.en',  name: 'Base English',        size: '~74 MB',  tag: 'Balanced', hfRepo: 'Systran/faster-whisper-base.en',  folder: 'models--Systran--faster-whisper-base.en'  },
  { id: 'small.en', name: 'Small English',       size: '~244 MB', tag: 'Accurate', hfRepo: 'Systran/faster-whisper-small.en', folder: 'models--Systran--faster-whisper-small.en', bundled: true },
  { id: 'tiny',     name: 'Tiny Multilingual',   size: '~39 MB',  tag: 'Fastest',  hfRepo: 'Systran/faster-whisper-tiny',     folder: 'models--Systran--faster-whisper-tiny'     },
  { id: 'base',     name: 'Base Multilingual',   size: '~74 MB',  tag: 'Balanced', hfRepo: 'Systran/faster-whisper-base',     folder: 'models--Systran--faster-whisper-base'     },
  { id: 'small',    name: 'Small Multilingual',  size: '~244 MB', tag: 'Accurate', hfRepo: 'Systran/faster-whisper-small',    folder: 'models--Systran--faster-whisper-small'    },
];

function isModelInstalled(modelId) {
  const info = WHISPER_MODELS.find(x => x.id === modelId);
  if (!info) return false;
  const dir = path.join(WHISPER_MODEL_DIR, info.folder);
  // Model is installed if its folder exists and has at least one .bin file
  try {
    if (!fs.existsSync(dir)) return false;
    const walk = (d) => {
      for (const f of fs.readdirSync(d)) {
        const fp = path.join(d, f);
        if (fs.statSync(fp).isDirectory()) { if (walk(fp)) return true; }
        else if (f.endsWith('.bin')) return true;
      }
      return false;
    };
    return walk(dir);
  } catch { return false; }
}
const APP_ICON           = path.join(ASSETS_DIR, 'icon.png');
// Legacy locations kept for migration / backward compatibility
const LEGACY_DATA_DIR     = path.join(ASSETS_DIR, 'Data');
const LEGACY_BIBLE_DIR    = path.join(app.getAppPath(), 'data');
const LEGACY_PRES_ASSETS  = path.join(ASSETS_DIR, 'Presentations');
const LEGACY_MEDIA_ASSETS = path.join(ASSETS_DIR, 'Media');
const LEGACY_SCHEDULES_DIR= path.join(LEGACY_DATA_DIR, 'schedules');
const PRESETS_FILE    = path.join(DATA_DIR, 'presets.json');
const PRESETS_ASSETS  = path.join(APPDATA_ROOT, 'PresetAssets');
// Ensure folders exist on first run
[APPDATA_ROOT, DATA_DIR, BIBLE_DIR, PRES_ASSETS, VIDEO_ASSETS, AUDIO_ASSETS, IMAGE_ASSETS, SCHEDULES_DIR, PRESETS_ASSETS].forEach(d => {
  try { fs.mkdirSync(d, { recursive: true }); } catch(e) {}
});
// BUG-B FIX: ensureBundledKjv() removed here — was called before its definition.
// It is called inside app.whenReady() after all functions are defined.

const SETTINGS_FILE    = path.join(DATA_DIR, 'settings.json');
const TRANSCRIPTS_FILE = path.join(DATA_DIR, 'transcripts.json');
const THEMES_FILE      = path.join(DATA_DIR, 'themes.json');
const SONGS_FILE       = path.join(DATA_DIR, 'songs.json');
const SONG_BACKUP_DIR  = path.join(DATA_DIR, 'song-library-backups');
const SONG_BACKUP_META = path.join(DATA_DIR, 'song-backup-meta.json');
const MEDIA_FILE       = path.join(DATA_DIR, 'media.json');
const RECENT_SCHEDULES_FILE = path.join(DATA_DIR, 'recent_schedules.json');
const DETECTION_REVIEW_DIR = path.join(DATA_DIR, 'detection_review');
const DETECTION_PHRASES_FILE = path.join(DETECTION_REVIEW_DIR, 'learned_phrases.json');
const DETECTION_EVENTS_FILE = path.join(DETECTION_REVIEW_DIR, 'review_events.json');
const SERVICE_ARCHIVE_DIR = path.join(DATA_DIR, 'service_archive');
const SERVICE_ARCHIVE_INDEX_FILE = path.join(SERVICE_ARCHIVE_DIR, 'index.json');
const CLIP_EXPORT_DIR = path.join(DATA_DIR, 'clip_exports');

const CACHE_DIR          = path.join(DATA_DIR, 'cache');
const THUMBS_DIR         = path.join(CACHE_DIR, 'thumbnails');
const TEMP_MEDIA_DIR     = path.join(CACHE_DIR, 'temp-media');
const SESSION_CACHE_DIR  = path.join(CACHE_DIR, 'sessions');
[CACHE_DIR, THUMBS_DIR, TEMP_MEDIA_DIR, SESSION_CACHE_DIR].forEach(d => {
  try { fs.mkdirSync(d, { recursive: true }); } catch(e) {}
});

function resolveBibleRuntimeResource(...parts) {
  const candidates = [
    path.join(process.resourcesPath || '', ...parts),
    path.join(app.getAppPath(), ...parts),
    path.join(__dirname, '..', ...parts),
    path.join(process.cwd(), ...parts),
  ].filter(Boolean);
  for (const candidate of candidates) {
    try { if (fs.existsSync(candidate)) return candidate; } catch (_) {}
  }
  return candidates[0];
}

function getBibleSearchDirs() {
  return Array.from(new Set([
    BIBLE_DIR,
    resolveBibleRuntimeResource('data'),
    path.join(__dirname, '..', 'data'),
    LEGACY_BIBLE_DIR,
  ])).filter(Boolean);
}


function ensureBundledKjv(){
  try{
    const bundledCandidates = [
      resolveBibleRuntimeResource('data', 'kjv.json'),
      path.join(ASSETS_DIR, 'Data', 'kjv.json'),
      path.join(__dirname, '..', 'data', 'kjv.json'),
      path.join(app.getAppPath(), 'data', 'kjv.json'),
    ];
    const target = path.join(BIBLE_DIR, 'kjv.json');
    for (const src of bundledCandidates) {
      if (src && fs.existsSync(src)) {
        fs.mkdirSync(BIBLE_DIR, { recursive: true });
        const shouldCopy = !fs.existsSync(target) || fs.statSync(target).size < 10000;
        if (shouldCopy) {
          fs.copyFileSync(src, target);
          console.log(`[BibleDB] Seeded built-in KJV from ${src}`);
        }
        return;
      }
    }
    console.warn('[BibleDB] No bundled KJV source was found');
  } catch(e) {
    console.warn('[BibleDB] Failed to seed KJV:', e.message);
  }
}



function sanitizeBibleTranslation(name) {
  return String(name || 'custom').replace(/[^a-zA-Z0-9]/g, '').toLowerCase() || 'custom';
}

function sanitizeFileName(name = 'file') {
  const base = path.basename(String(name)).replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').trim();
  return base || 'file';
}
function uniqueDestinationPath(dir, filename) {
  const safe = sanitizeFileName(filename);
  const parsed = path.parse(safe);
  let candidate = path.join(dir, safe);
  let i = 1;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${parsed.name} (${i})${parsed.ext}`);
    i += 1;
  }
  return candidate;
}
function mediaAssetDirForType(type) {
  if (type === 'video') return VIDEO_ASSETS;
  if (type === 'audio') return AUDIO_ASSETS;
  return IMAGE_ASSETS;
}
function normalizeMediaType(filename = '') {
  const ext = path.extname(filename).toLowerCase().slice(1);
  const video = new Set(['mp4','webm','mov','mkv','avi','wmv','m4v','mpg','mpeg','3gp','flv']);
  const audio = new Set(['mp3','wav','ogg','flac','aac','m4a','wma','opus','aiff']);
  const image = new Set(['jpg','jpeg','png','gif','webp','bmp','svg','tiff','tif']);
  if (video.has(ext)) return 'video';
  if (audio.has(ext)) return 'audio';
  if (image.has(ext)) return 'image';
  return null;
}
function toStoredMediaItem(filePath) {
  const stat = fs.statSync(filePath);
  const ext = path.extname(filePath).toLowerCase().replace(/^\./, '');
  const type = normalizeMediaType(filePath);
  return {
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 8),
    name: path.basename(filePath, path.extname(filePath)),
    title: path.basename(filePath, path.extname(filePath)),
    path: filePath,
    type,
    ext,
    size: stat.size,
    loop: true,
    mute: false,
    volume: 1,
    aspectRatio: 'contain',
  };
}


const gotSingleInstanceLock = app.requestSingleInstanceLock();
if (!gotSingleInstanceLock) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv) => {
    // Check for anchorcast:// protocol URL in argv (Windows/Linux)
    const protocolUrl = argv.find(a => a.startsWith('anchorcast://'));
    if (protocolUrl) { _handleProtocolUrl(protocolUrl); }
    const scheduleArg = _extractScheduleArg(argv);
    if (scheduleArg) _queueScheduleOpen(scheduleArg);
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      if (!mainWindow.isVisible()) mainWindow.show();
      mainWindow.focus();
      _consumePendingScheduleOpen();
    }
  });
  app.on('open-file', (event, filePath) => {
    event.preventDefault();
    if (_queueScheduleOpen(filePath)) _consumePendingScheduleOpen();
  });
  // macOS: handle anchorcast:// protocol URL
  app.on('open-url', (event, url) => {
    event.preventDefault();
    _handleProtocolUrl(url);
  });
}

// ── Lifecycle ─────────────────────────────────────────────────────────────────
app.whenReady().then(async ()=>{
  ensureBundledKjv();
  _maybeRunSundaySongBackup();
  currentSettings=loadSettings();
  currentSettings = { ...defaultSettings(), ...currentSettings };
  // Always reset hideGetStarted so the welcome page shows on each fresh launch.
  // The user permanently dismisses it by checking "Don't show this again".
  // Previous patches may have accidentally saved it as true — reset it here.
  currentSettings.hideGetStarted = false;
  currentSettings.remoteRequireAuth = isRemoteAuthRequired(currentSettings.remoteRequireAuth);
  currentSettings.remoteAdminPin = normalizeRemotePinValue(currentSettings.remoteAdminPin || currentSettings.remotePin);
  currentSettings.remoteScripturePin = normalizeRemotePinValue(currentSettings.remoteScripturePin);
  currentSettings.remoteSongsPin = normalizeRemotePinValue(currentSettings.remoteSongsPin);
  currentSettings.remoteMediaPin = normalizeRemotePinValue(currentSettings.remoteMediaPin);
  currentSettings.remoteMonitorPin = normalizeRemotePinValue(currentSettings.remoteMonitorPin);
  if ((currentSettings.remoteRequireAuth !== false) && !currentSettings.remoteAdminPin) currentSettings.remoteAdminPin = String(Math.floor(100000 + Math.random()*900000));
  if ((currentSettings.remoteRequireAuth !== false) && !currentSettings.remoteScripturePin) currentSettings.remoteScripturePin = String(Math.floor(100000 + Math.random()*900000));
  if ((currentSettings.remoteRequireAuth !== false) && !currentSettings.remoteSongsPin) currentSettings.remoteSongsPin = String(Math.floor(100000 + Math.random()*900000));
  if ((currentSettings.remoteRequireAuth !== false) && !currentSettings.remoteMediaPin) currentSettings.remoteMediaPin = String(Math.floor(100000 + Math.random()*900000));
  if ((currentSettings.remoteRequireAuth !== false) && !currentSettings.remoteMonitorPin) currentSettings.remoteMonitorPin = String(Math.floor(100000 + Math.random()*900000));
  currentSettings.remotePin = currentSettings.remoteAdminPin;
  if ((currentSettings.remoteRequireAuth !== false) && !String(loadSettings().remoteAdminPin || loadSettings().remotePin || '').trim()) {
    try { fs.mkdirSync(path.dirname(SETTINGS_FILE), { recursive:true }); fs.writeFileSync(SETTINGS_FILE, JSON.stringify(currentSettings, null, 2)); } catch(_) {}
  }
  whisperSource = currentSettings.transcriptSource || 'local';
  const initialScheduleArg = _extractScheduleArg(process.argv);
  if (initialScheduleArg) _queueScheduleOpen(initialScheduleArg);
  // Check for anchorcast:// protocol URL passed at startup
  const initialProtocolUrl = process.argv.find(a => a.startsWith('anchorcast://'));
  if (initialProtocolUrl) _handleProtocolUrl(initialProtocolUrl);

  // ── Grant microphone permission to all windows ─────────────────────────────
  // Without this, Electron silently denies all getUserMedia / SpeechRecognition requests
  const { session, systemPreferences } = require('electron');

  // Handle permission requests — grant microphone automatically
  session.defaultSession.setPermissionRequestHandler((webContents, permission, callback) => {
    const allowedPermissions = ['media', 'audioCapture', 'microphone'];
    if (allowedPermissions.includes(permission)) {
      callback(true); // grant
    } else {
      callback(false);
    }
  });

  // Also set permission check handler (for already-granted checks)
  session.defaultSession.setPermissionCheckHandler((webContents, permission) => {
    const allowed = ['media', 'audioCapture', 'microphone'];
    return allowed.includes(permission);
  });

  // macOS: request microphone access at the OS level
  if (process.platform === 'darwin') {
    try {
      const micStatus = systemPreferences.getMediaAccessStatus('microphone');
      if (micStatus !== 'granted') {
        await systemPreferences.askForMediaAccess('microphone');
      }
    } catch(e) {
      console.warn('[Mic] macOS media access request failed:', e.message);
    }
  }

  createSplashWindow();

  // Start renderer server first — serves app on localhost so Web Speech API works
  await startRendererServer();

  // Register media:// protocol handler — serves local files with range request support
  // Range requests are required for <video> seek/play to work with MP4/WMA etc.
  // Register anchorcast:// protocol for magic link activation from email
  if (process.defaultApp) {
    if (process.argv.length >= 2) app.setAsDefaultProtocolClient(APP_PROTOCOL, process.execPath, [path.resolve(process.argv[1])]);
  } else {
    app.setAsDefaultProtocolClient(APP_PROTOCOL);
  }

  protocol.handle('media', async (request) => {
    let filePath = decodeURIComponent(request.url.slice('media://'.length));
    // On Windows: /C:/path → C:/path
    if (process.platform === 'win32' && filePath.startsWith('/')) {
      filePath = filePath.slice(1);
    }
    // Normalize separators
    filePath = filePath.replace(/\//g, path.sep);

    try {
      const stat  = fs.statSync(filePath);
      const total = stat.size;
      const ext   = path.extname(filePath).toLowerCase().slice(1);

      // MIME type map
      const MIME = {
        mp4:'video/mp4', webm:'video/webm', mov:'video/quicktime',
        mkv:'video/x-matroska', avi:'video/x-msvideo', wmv:'video/x-ms-wmv',
        m4v:'video/mp4', mpg:'video/mpeg', mpeg:'video/mpeg',
        flv:'video/x-flv', '3gp':'video/3gpp',
        mp3:'audio/mpeg', wav:'audio/wav', ogg:'audio/ogg',
        flac:'audio/flac', aac:'audio/aac', m4a:'audio/mp4', wma:'audio/x-ms-wma',
        jpg:'image/jpeg', jpeg:'image/jpeg', png:'image/png',
        gif:'image/gif', webp:'image/webp', bmp:'image/bmp', svg:'image/svg+xml',
      };
      const contentType = MIME[ext] || 'application/octet-stream';

      const rangeHeader = request.headers.get('range');
      if (rangeHeader) {
        // Parse Range: bytes=start-end
        const [, startStr, endStr] = /bytes=(\d*)-(\d*)/.exec(rangeHeader) || [];
        const start = startStr ? parseInt(startStr) : 0;
        const end   = endStr   ? parseInt(endStr)   : total - 1;
        const chunkSize = end - start + 1;

        const stream = fs.createReadStream(filePath, { start, end });

        return new Response(nodeToWeb(stream), {
          status: 206,
          headers: {
            'Content-Type':   contentType,
            'Content-Range':  `bytes ${start}-${end}/${total}`,
            'Accept-Ranges':  'bytes',
            'Content-Length': String(chunkSize),
          }
        });
      }

      // No range header — serve full file
      const stream = fs.createReadStream(filePath);

      return new Response(nodeToWeb(stream), {
        status: 200,
        headers: {
          'Content-Type':   contentType,
          'Accept-Ranges':  'bytes',
          'Content-Length': String(total),
        }
      });

    } catch (err) {
      console.error('[media://] Error serving:', filePath, err.message);
      return new Response('File not found', { status: 404 });
    }
  });

  createMainWindow();
  buildMenu();
  startHttpServer(currentSettings.httpPort||8080);
  if(currentSettings.ndiEnabled){
    setTimeout(()=>startNdi(), 3000);
  }
  // Auto-start local Whisper server (non-blocking — loads model in background)
  // Auto-migrate multilingual model names to English-only variants (better accuracy)
  const _rawModel = currentSettings.whisperModel || 'small.en';
  const whisperModel = _rawModel === 'small' ? 'small.en'
    : _rawModel === 'base' ? 'base.en'
    : _rawModel === 'tiny'  ? 'tiny.en'
    : _rawModel;
  // Persist the migrated .en model so Settings UI shows correct value
  if (whisperModel !== _rawModel) {
    currentSettings.whisperModel = whisperModel;
    try { fs.writeFileSync(SETTINGS_FILE, JSON.stringify(currentSettings, null, 2)); } catch(_) {}
  }
  // Only start Whisper server if the model is already downloaded.
  // If model is missing, show the setup banner immediately instead of
  // silently downloading for 1-2 minutes with no user feedback.
  const _modelReady = isModelInstalled(whisperModel) ||
    isModelInstalled(whisperModel.replace('.en', '')); // also check multilingual variant

  if (_modelReady) {
    // Start Whisper after window is ready — guarantees process.resourcesPath is set
    _runAfterWindowReady(() => {
      startWhisperServer(whisperModel).then(ok => {
        if (!ok) console.log('[Whisper] Not available — use setup_whisper.bat to install');
      });
    }, 500);
  } else {
    // Model not installed — detect WHY and show the correct banner.
    // Must run after did-finish-load so mainWindow.webContents.send() is not lost.
    _runAfterWindowReady(() => {
      const setupBatExists = fs.existsSync(resolveRuntimeResource('setup_whisper.bat'));
      const pyPaths = [
        path.join(process.resourcesPath || '', 'python', 'python.exe'),
        path.join(app.getAppPath(), 'python', 'python.exe'),
        path.join(__dirname, '..', 'python', 'python.exe'),
        path.join(process.cwd(), 'python', 'python.exe'),
      ];
      const bundledPy = pyPaths.find(p => { try { return fs.existsSync(p); } catch { return false; } }) || '';
      const hasPython = !!bundledPy;

      if (!hasPython) {
        console.log('[Whisper] Python not found — showing no_python banner');
        mainWindow?.webContents.send('whisper-setup-needed', { reason: 'no_python', setupBatExists });
        return;
      }

      const { spawnSync } = require('child_process');
      const fwCheck = spawnSync(bundledPy, ['-c', 'from faster_whisper import WhisperModel'], {
        encoding: 'utf8', timeout: 8000, windowsHide: true,
      });

      if (fwCheck.status !== 0) {
        console.log('[Whisper] faster-whisper missing — showing no_faster_whisper banner');
        mainWindow?.webContents.send('whisper-setup-needed', { reason: 'no_faster_whisper', setupBatExists });
        return;
      }

      console.log('[Whisper] Model not found — showing model_not_found banner');
      mainWindow?.webContents.send('whisper-setup-needed', {
        reason: 'model_not_found', model: whisperModel, setupBatExists,
      });
    }, 2000); // 2s after window ready — allows renderer IPC to fully register
  }
});
// Track whether transcript needs saving (set by renderer via IPC)
let _transcriptUnsaved = false;
ipcMain.on('transcript-unsaved-state', (_, unsaved) => { _transcriptUnsaved = unsaved; });

// Before closing: if there's an unsaved transcript, ask the user
app.on('before-quit', (e) => {
  if (!_transcriptUnsaved) return;
  if (!mainWindow || mainWindow.isDestroyed()) return;
  e.preventDefault(); // block quit temporarily
  mainWindow.webContents.send('confirm-quit-with-transcript');
});

// Renderer confirmed: save then quit, or just quit
ipcMain.on('quit-confirmed', () => { _transcriptUnsaved = false; app.quit(); });
ipcMain.on('quit-cancelled', () => { /* do nothing — user changed mind */ });

app.on('window-all-closed',()=>{
  stopNdi();
  stopHttpServer();
  stopWhisperServer();
  if(rendererServer){ rendererServer.close(); rendererServer=null; }
  if(process.platform!=='darwin') app.quit();
});
app.on('activate',()=>{ if(!BrowserWindow.getAllWindows().length) createMainWindow(); });

// ── Windows ───────────────────────────────────────────────────────────────────
let rendererServer = null;
let rendererPort  = 0;

// Serve the renderer from localhost so Web Speech API works
// (file:// origin blocks getUserMedia and SpeechRecognition in Chromium)
async function startRendererServer() {
  return new Promise((resolve) => {
    const rendererDir = path.join(__dirname, 'renderer');
    const assetsDir = path.join(__dirname, '..', 'assets');
    rendererServer = http.createServer((req, res) => {
      let urlPath;
      try { urlPath = decodeURIComponent((req.url || '/').split('?')[0]); }
      catch(_) { res.writeHead(400); res.end('Bad Request'); return; }
      let filePath;
      if (urlPath.startsWith('/assets/')) {
        filePath = path.resolve(path.join(assetsDir, urlPath.slice(8)));
        if (!filePath.startsWith(assetsDir + path.sep) && filePath !== assetsDir) {
          res.writeHead(403); res.end(); return;
        }
      } else {
        filePath = path.resolve(path.join(rendererDir, urlPath === '/' ? 'index.html' : urlPath));
        if (!filePath.startsWith(rendererDir + path.sep) && filePath !== rendererDir) {
          res.writeHead(403); res.end(); return;
        }
      }
      fs.readFile(filePath, (err, data) => {
        if (err) { res.writeHead(404); res.end(); return; }
        const ext = path.extname(filePath).toLowerCase();
        const mime = {
          '.html':'text/html','.js':'application/javascript',
          '.css':'text/css','.json':'application/json',
          '.png':'image/png','.jpg':'image/jpeg','.svg':'image/svg+xml',
          '.ico':'image/x-icon','.woff2':'font/woff2','.woff':'font/woff',
        }[ext] || 'application/octet-stream';
        res.writeHead(200, {'Content-Type': mime, 'Access-Control-Allow-Origin':'*'});
        res.end(data);
      });
    });
    rendererServer.listen(0, '127.0.0.1', () => {
      rendererPort = rendererServer.address().port;
      console.log(`[Renderer] Serving from http://127.0.0.1:${rendererPort}`);
      resolve(rendererPort);
    });
  });
}



// ── Registration window ───────────────────────────────────────────────────────
let registrationWindow = null;

function _handleProtocolUrl(url) {
  try {
    const u = new URL(url);
    if (u.hostname !== 'register') return;
    const token     = u.searchParams.get('token')    || '';
    const name      = decodeURIComponent(u.searchParams.get('name')    || '');
    const email     = decodeURIComponent(u.searchParams.get('email')   || '');
    const church    = decodeURIComponent(u.searchParams.get('church')  || '');
    const hwIdParam = decodeURIComponent(u.searchParams.get('hwId')    || '');
    const hwId      = getHardwareId();

    // Verify hardware ID matches
    if (hwIdParam && hwIdParam !== hwId) {
      dialog.showMessageBox({ type:'error', title:'AnchorCast Registration',
        message:'This registration link was created for a different device.',
        detail:'Please register again using Help → Register on this device.',
        buttons:['OK'] });
      return;
    }

    const result = activateRegistration(name, email, token, church);
    if (result.success) {
      // Close registration window if open
      if (registrationWindow && !registrationWindow.isDestroyed()) {
        registrationWindow.close();
        registrationWindow = null;
      }
      // Show main app if not shown yet
      if (mainWindow && !mainWindow.isVisible()) {
        mainWindow.maximize();
        mainWindow.show();
        mainWindow.webContents.send('app-ready');
      }
      // Notify renderer
      mainWindow?.webContents.send('registration-complete', result);
      dialog.showMessageBox(mainWindow || undefined, {
        type:'info', title:'Registration Complete',
        message:'🎉 AnchorCast is now registered!',
        detail:`This product is registered to:\n${result.fullName}\n${result.email}`,
        buttons:['Start Using AnchorCast']
      });
    } else {
      dialog.showMessageBox({ type:'error', title:'Registration Failed',
        message:'Registration link is invalid or expired.',
        detail: result.error || 'Please try registering again.',
        buttons:['OK'] });
    }
  } catch(e) {
    console.error('[Protocol] Error handling URL:', url, e.message);
  }
}

function showRegistrationWindow() {
  if (registrationWindow && !registrationWindow.isDestroyed()) {
    registrationWindow.focus();
    return;
  }
  const win = new BrowserWindow({
    icon: APP_ICON,
    width: 520, height: 720,
    title: 'Register AnchorCast',
    backgroundColor: '#0a0a1a',
    resizable: false,
    show: false,
    center: true,
    alwaysOnTop: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });
  registrationWindow = win;
  win.setMenu(null);
  win.loadURL(`http://127.0.0.1:${rendererPort}/registration.html`);
  win.once('ready-to-show', () => win.show());
  win.on('closed', () => {
    registrationWindow = null;
    // If still not registered and main window never shown, quit
    if (!isRegistered() && (!mainWindow || !mainWindow.isVisible())) {
      app.quit();
    }
  });
}

function buildRegistrationHtml() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Register AnchorCast</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
    background:#0a0a1a;color:#e0e0e0;display:flex;flex-direction:column;
    height:100vh;padding:0;overflow:hidden}
  .header{background:linear-gradient(135deg,#0d0d20,#1a1a35);
    padding:24px 28px;text-align:center;border-bottom:1px solid rgba(201,168,76,.2)}
  .logo-icon{font-size:32px;margin-bottom:6px}
  .logo-text{color:#c9a84c;font-size:18px;font-weight:700;letter-spacing:1.5px}
  .logo-sub{color:#666;font-size:10px;letter-spacing:2px;margin-top:3px}
  .body{flex:1;padding:18px 26px;display:flex;flex-direction:column;gap:12px;overflow-y:auto}
  h2{font-size:16px;color:#fff;font-weight:600;margin-bottom:2px}
  .desc{font-size:12px;color:#777;line-height:1.6}
  .desc a{color:#c9a84c;cursor:pointer;text-decoration:none}
  .desc a:hover{text-decoration:underline}
  .field{display:flex;flex-direction:column;gap:5px}
  label{font-size:10px;color:#888;text-transform:uppercase;letter-spacing:.5px;font-weight:600}
  input{background:#0d0d20;border:1px solid rgba(255,255,255,.12);border-radius:7px;
    color:#e0e0e0;font-size:13px;padding:11px 13px;outline:none;
    transition:border-color .15s;font-family:inherit;width:100%}
  input:focus{border-color:#c9a84c}
  input::placeholder{color:#444}
  .optional{font-size:10px;color:#555;margin-top:2px}
  .btn{padding:13px;border:none;border-radius:8px;cursor:pointer;
    font-size:13px;font-weight:700;font-family:inherit;width:100%;
    transition:opacity .15s;letter-spacing:.3px}
  .btn:hover{opacity:.88}
  .btn:disabled{opacity:.4;cursor:not-allowed}
  .btn-primary{background:linear-gradient(135deg,#c9a84c,#e0bf6a);color:#000}
  .status{font-size:12px;min-height:18px;line-height:1.5;padding:0 2px;text-align:center}
  .oss{display:flex;align-items:center;justify-content:center;gap:6px;
    font-size:11px;color:#555;margin-top:4px}
  .oss a{color:#c9a84c;cursor:pointer;text-decoration:none}
  .oss a:hover{text-decoration:underline}
  .hwid{font-family:monospace;font-size:10px;color:#555;text-align:center;
    background:rgba(255,255,255,.03);border-radius:5px;padding:6px 10px;
    border:1px solid rgba(255,255,255,.06)}
  .steps{background:rgba(201,168,76,.06);border:1px solid rgba(201,168,76,.15);
    border-radius:8px;padding:10px 13px;font-size:11px;color:#999;line-height:1.8}
  .steps b{color:#c9a84c}
  .footer{padding:16px 28px;border-top:1px solid rgba(255,255,255,.06);text-align:center}
  .success-screen{display:none;flex-direction:column;align-items:center;
    justify-content:center;flex:1;padding:30px;text-align:center;gap:12px}
</style>
</head><body>
  <div class="header">
    <div class="logo-icon">⚓</div>
    <div class="logo-text">ANCHORCAST</div>
    <div class="logo-sub">LIVE SERMON DISPLAY</div>
  </div>

  <div class="body" id="formBody">
    <div>
      <h2>Register AnchorCast</h2>
      <p class="desc" style="margin-top:6px">
        AnchorCast is <a onclick="openGitHub()">free &amp; open source</a>.
        Registration is free and helps us know our users for support.
      </p>
    </div>

    <div class="steps">
      <b>How it works:</b><br>
      1. Fill in your details below<br>
      2. Click <b>Send Registration Email</b><br>
      3. Check your inbox and click the link<br>
      4. AnchorCast opens and registers automatically ✅
    </div>

    <div class="field">
      <label>Full Name *</label>
      <input id="nameInput" type="text" placeholder="e.g. John Smith" autocomplete="name" />
    </div>
    <div class="field">
      <label>Email Address *</label>
      <input id="emailInput" type="email" placeholder="e.g. john@church.org" autocomplete="email" />
    </div>
    <div class="field">
      <label>Church / Organisation <span style="color:#555;font-weight:400">(optional)</span></label>
      <input id="churchInput" type="text" placeholder="e.g. Grace Community Church" autocomplete="organization" />
    </div>

    <div>
      <button class="btn btn-primary" id="sendBtn" onclick="sendEmail()">
        📧 Send Registration Email
      </button>
      <div class="status" id="status"></div>
    </div>

    <div class="hwid" id="hwIdDisplay">Hardware ID: loading…</div>
    <div class="oss">
      Free &amp; Open Source —
      <a onclick="openGitHub()">github.com/anchorcastapp-team/anchorcastapp</a>
    </div>
  </div>

  <div class="footer">
    <div style="font-size:10px;color:#444">
      Already have a link? Click it in your email client to activate.
    </div>
  </div>

  <script>
    // Load hardware ID
    (async () => {
      try {
        const hwId = await window.electronAPI.getHardwareId();
        document.getElementById('hwIdDisplay').textContent = 'Hardware ID: ' + hwId;
      } catch(_) {}
    })();

    function openGitHub() {
      window.electronAPI?.openExternal('https://github.com/anchorcastapp-team/anchorcastapp');
    }

    async function sendEmail() {
      const name   = document.getElementById('nameInput').value.trim();
      const email  = document.getElementById('emailInput').value.trim();
      const church = document.getElementById('churchInput').value.trim();
      const status = document.getElementById('status');
      const btn    = document.getElementById('sendBtn');

      if (!name)  { status.style.color='#e74c3c'; status.textContent='⚠ Please enter your full name.'; return; }
      if (!email || !email.includes('@')) {
        status.style.color='#e74c3c'; status.textContent='⚠ Please enter a valid email address.'; return;
      }
      btn.disabled = true;
      status.style.color='#888';
      status.textContent='Sending…';

      try {
        const r = await window.electronAPI.sendRegistrationEmail(name, email, church);
        if (r.success) {
          // Replace entire page with clean success screen (no backticks — nested in outer template)
          document.body.innerHTML = '<style>' +
            '*{box-sizing:border-box;margin:0;padding:0}' +
            'body{font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;' +
              'background:#0a0a1a;color:#e0e0e0;display:flex;flex-direction:column;height:100vh;overflow:hidden}' +
            '.hdr{background:linear-gradient(135deg,#0d0d20,#1a1a35);' +
              'padding:22px 28px;text-align:center;border-bottom:1px solid rgba(201,168,76,.2)}' +
            '.logo{color:#c9a84c;font-size:17px;font-weight:700;letter-spacing:1px}' +
            '.content{flex:1;display:flex;flex-direction:column;align-items:center;' +
              'justify-content:center;padding:28px;gap:16px;text-align:center}' +
            '.icon{font-size:52px;margin-bottom:4px}' +
            '.title{font-size:18px;font-weight:700;color:#2ecc71}' +
            '.sub{font-size:13px;color:#888;line-height:1.7;max-width:340px}' +
            '.email-box{background:#0d0d20;border:1px solid rgba(46,204,113,.25);' +
              'border-radius:8px;padding:14px 18px;width:100%;max-width:360px;font-size:13px;color:#2ecc71;font-weight:600}' +
            '.steps-ok{background:rgba(201,168,76,.06);border:1px solid rgba(201,168,76,.15);' +
              'border-radius:8px;padding:14px 18px;width:100%;max-width:360px;' +
              'font-size:12px;color:#aaa;line-height:2.1;text-align:left}' +
            '.steps-ok b{color:#c9a84c}' +
            '.btn-back{padding:10px 24px;background:transparent;' +
              'border:1px solid rgba(255,255,255,.12);border-radius:7px;' +
              'color:#666;font-size:11px;cursor:pointer;font-family:inherit;margin-top:4px}' +
            '.btn-back:hover{color:#c9a84c;border-color:rgba(201,168,76,.3)}' +
            '</style>' +
            '<div class="hdr"><div class="logo">⚓ ANCHORCAST</div>' +
              '<div style="color:#555;font-size:10px;letter-spacing:1.5px;margin-top:3px">LIVE SERMON DISPLAY</div></div>' +
            '<div class="content">' +
              '<div class="icon">📧</div>' +
              '<div class="title">Registration Email Sent!</div>' +
              '<div class="email-box">✉ Email sent to: ' + email + '</div>' +
              '<div class="steps-ok"><b>Next steps:</b><br>' +
                '1. Open your email inbox<br>' +
                '2. Find the email from AnchorCast<br>' +
                '3. Click <b>Complete Registration</b><br>' +
                '4. AnchorCast opens and registers automatically ✅' +
              '</div>' +
              '<div class="sub">This link is specific to this device.<br>Didn\'t receive it? Check your spam folder.</div>' +
              '<button class="btn-back" onclick="location.reload()">← Back / Resend Email</button>' +
            '</div>';
        } else {
          status.style.color='#e74c3c';
          status.textContent='❌ ' + (r.error || 'Failed to send. Please try again.');
          btn.disabled = false;
        }
      } catch(e) {
        status.style.color='#e74c3c';
        status.textContent='❌ ' + (e.message || 'Could not send email. Check internet connection.');
        btn.disabled = false;
      }
    }
  </script>
</body></html>`;
}

function createSplashWindow(){
  if(splashWindow && !splashWindow.isDestroyed()) return splashWindow;
  splashShownAt = Date.now();
  splashWindow = new BrowserWindow({
    width: 520, height: 520, frame: false, transparent: true, resizable: false,
    show: false, alwaysOnTop: true, backgroundColor: '#00000000',
    icon: APP_ICON,
    webPreferences:{ nodeIntegration:false, contextIsolation:true, backgroundThrottling:false }
  });
  if (rendererPort) {
    splashWindow.loadURL(`http://127.0.0.1:${rendererPort}/splash.html`);
  } else {
    splashWindow.loadFile(path.join(__dirname,'renderer','splash.html'));
  }
  splashWindow.once('ready-to-show', ()=>{
    // 80ms delay on Windows so DWM composites transparency before showing
    const delay = process.platform === 'win32' ? 80 : 0;
    setTimeout(()=>{ try{ if(splashWindow && !splashWindow.isDestroyed()) splashWindow.show(); }catch(e){} }, delay);
  });
  splashWindow.setMenu(null);
  splashWindow.on('closed',()=>{ splashWindow=null; });
  return splashWindow;
}

function createMainWindow(){
  const{width,height}=screen.getPrimaryDisplay().workAreaSize;
  mainWindow=new BrowserWindow({
    icon: APP_ICON,
    width:Math.min(1680,width), height:Math.min(960,height),
    minWidth:1280, minHeight:720,
    title:'AnchorCast', backgroundColor:'#08101d',
    show: false,
    opacity: 0,  // start invisible — prevents any white flash at OS level
    webPreferences:{
      nodeIntegration:false,
      contextIsolation:true,
      preload:path.join(__dirname,'preload.js'),
      webSecurity: false,
    },
    titleBarStyle:process.platform==='darwin'?'hiddenInset':'default',
    trafficLightPosition:{x:16,y:12},
  });
  mainWindow.loadURL(`http://127.0.0.1:${rendererPort}/index.html`);
  // Recover from white/frozen screen caused by display changes
  mainWindow.on('unresponsive', () => {
    console.log('[MainWindow] Unresponsive — will attempt recovery');
    // Give it 3 seconds to recover naturally first
    setTimeout(() => {
      try {
        if (mainWindow && !mainWindow.isDestroyed()) {
          // Try a soft JS ping first
          mainWindow.webContents.executeJavaScript('1+1')
            .then(() => {
              console.log('[MainWindow] Recovered naturally');
            })
            .catch(() => {
              // Still frozen — reload the renderer
              console.log('[MainWindow] Still unresponsive — reloading renderer');
              mainWindow.webContents.reload();
            });
        }
      } catch(e) {
        try { mainWindow.webContents.reload(); } catch(_) {}
      }
    }, 3000);
  });

  mainWindow.on('responsive', () => {
    console.log('[MainWindow] Responsive again');
  });

  mainWindow.webContents.once('did-finish-load',()=>{
    const elapsed = Date.now() - splashShownAt;
    const reveal = () => {
      if (!isRegistered()) {
        if(splashWindow && !splashWindow.isDestroyed()){
          try { splashWindow.close(); } catch(e) {}
          splashWindow = null;
        }
        showRegistrationWindow();
        return;
      }
      // Show main window first (still at opacity:0 — fully invisible to user)
      mainWindow.maximize();
      mainWindow.show();
      mainWindow.webContents.send('app-ready');
      // Fade main window in 0→1 over ~100ms
      let op = 0;
      const tick = () => {
        op = Math.min(1, op + 0.15);
        try { mainWindow.setOpacity(op); } catch(_) {}
        if(op < 1) setTimeout(tick, 16);
      };
      tick();
      // Fade splash OUT 1→0 simultaneously then destroy
      if(splashWindow && !splashWindow.isDestroyed()){
        let sop = 1;
        const stk = () => {
          sop = Math.max(0, sop - 0.15);
          try { splashWindow.setOpacity(sop); } catch(_) {}
          if(sop > 0) setTimeout(stk, 16);
          else { try { splashWindow.close(); } catch(e) {} splashWindow = null; }
        };
        stk();
      }
      if(isDev) mainWindow.webContents.openDevTools({mode:'detach'});
    };
    const wait = Math.max(0, MIN_SPLASH_MS - elapsed);
    setTimeout(reveal, wait);
  });
  mainWindow.on('closed',()=>{
    mainWindow=null;
    if(projectionWindow){projectionWindow.close();projectionWindow=null;}
    // Close timer window when main app closes
    if(countdownWindow && !countdownWindow.isDestroyed()){
      countdownWindow.close();
      countdownWindow=null;
    }
  });
}

let projectionPowerBlockerId = null;
let projectionDisplayListenersAttached = false;

function createProjectionWindow(displayId){
  if(projectionWindow){projectionWindow.focus();return;}
  const displays=screen.getAllDisplays();
  const target=displays.find(d=>d.id===displayId)
    ||displays.find(d=>d.id!==screen.getPrimaryDisplay().id)
    ||displays[0];
  const{x,y,width,height}=target.bounds;
  const scaleFactor = target.scaleFactor || 1;
  const primaryId = screen.getPrimaryDisplay().id;
  const isSameDisplayAsMain = (target.id === primaryId);
  projectionWindow=new BrowserWindow({
    icon: APP_ICON,
    x,y,width,height,fullscreen:true,frame:false,
    alwaysOnTop: !isSameDisplayAsMain,
    // Hide from taskbar and Win+Tab Task View on external display.
    // This prevents the operator accidentally selecting/minimising the
    // projection window while switching apps on the laptop screen.
    skipTaskbar: !isSameDisplayAsMain,
    backgroundColor:'#000000',
    show: false,
    webPreferences:{
      nodeIntegration:false,
      contextIsolation:true,
      preload:path.join(__dirname,'preload.js'),
      backgroundThrottling:false,
      webSecurity:false,
    },
  });
  projectionWindow._targetDisplayId = target.id;
  projectionWindow._displayScaleFactor = scaleFactor;
  projectionWindow._displayWidth  = width;
  projectionWindow._displayHeight = height;
  projectionWindow.loadFile(path.join(__dirname,'renderer','projection.html'));
  projectionWindow.once('ready-to-show',()=>{
    try{
      projectionWindow?.show();
      // On Windows: use the highest alwaysOnTop level so the projection
      // screen stays above the taskbar and is immune to Win+Tab / Task View
      if (!isSameDisplayAsMain && process.platform === 'win32') {
        projectionWindow.setAlwaysOnTop(true, 'screen-saver');
      }
    }catch(e){}
    // On same display: bring main window to front so operator can use it
    if (isSameDisplayAsMain) {
      setTimeout(() => {
        try {
          if (mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.setAlwaysOnTop(true);
            mainWindow.show();
            mainWindow.focus();
            setTimeout(() => {
              if (mainWindow && !mainWindow.isDestroyed()) mainWindow.setAlwaysOnTop(false);
            }, 500);
          }
        } catch(e) {}
      }, 300);
    }
  });
  projectionWindow.webContents.on('did-finish-load', () => {
    try { projectionWindowReadySync(); } catch(e) {}
  });
  // Guard: restore projection if Win+D, Win+Tab or taskbar causes it to
  // lose fullscreen or become minimized. Check every 500ms.
  let _projGuardInterval = null;
  let _projGuardPaused = false; // paused during display removal to avoid GPU conflicts

  function _startProjGuard() {
    if (_projGuardInterval) return;
    _projGuardInterval = setInterval(() => {
      try {
        if (!projectionWindow || projectionWindow.isDestroyed()) {
          clearInterval(_projGuardInterval);
          _projGuardInterval = null;
          return;
        }
        // Don't fight the OS during display removal/reconnection
        if (_projGuardPaused) return;

        if (!isSameDisplayAsMain) {
          if (projectionWindow.isMinimized()) {
            projectionWindow.restore();
            projectionWindow.setFullScreen(true);
            projectionWindow.show();
          } else if (!projectionWindow.isFullScreen()) {
            // Taskbar or Task View stole fullscreen — reclaim it
            projectionWindow.setFullScreen(true);
          }
        }
      } catch(e) {
        clearInterval(_projGuardInterval);
        _projGuardInterval = null;
      }
    }, 500);
  }
  _startProjGuard();
  projectionWindow.on('closed', () => {
    if (_projGuardInterval) { clearInterval(_projGuardInterval); _projGuardInterval = null; }
  });

  if (projectionPowerBlockerId !== null && powerSaveBlocker.isStarted(projectionPowerBlockerId)) {
    powerSaveBlocker.stop(projectionPowerBlockerId);
  }
  projectionPowerBlockerId = powerSaveBlocker.start('prevent-display-sleep');

  if (!projectionDisplayListenersAttached) {
    projectionDisplayListenersAttached = true;

    screen.on('display-removed', (_evt, removedDisplay) => {
      if (!projectionWindow || projectionWindow.isDestroyed()) return;
      if (projectionWindow._targetDisplayId !== removedDisplay.id) return;

      // Pause the guard interval immediately — it fighting setFullScreen
      // during GPU compositor teardown is what causes the white screen freeze
      _projGuardPaused = true;
      console.log('[Display] Guard paused for display removal');

      projectionWindow._lostTargetDisplay = true;
      const primaryId  = screen.getPrimaryDisplay().id;
      const remaining  = screen.getAllDisplays();
      // Look for another external display (not the primary/laptop screen)
      const otherExternal = remaining.find(d => d.id !== primaryId);

      if (otherExternal) {
        // Move to another available external display
        const b = otherExternal.bounds;
        projectionWindow._targetDisplayId   = otherExternal.id;
        projectionWindow._displayScaleFactor = otherExternal.scaleFactor || 1;
        projectionWindow._displayWidth  = b.width;
        projectionWindow._displayHeight = b.height;
        projectionWindow.setBounds({ x: b.x, y: b.y, width: b.width, height: b.height });
        // Unpause guard after a short settle period then resume fullscreen
        setTimeout(() => {
          try {
            if (projectionWindow && !projectionWindow.isDestroyed()) {
              projectionWindow.setFullScreen(true);
            }
          } catch(e) {}
          _projGuardPaused = false;
        }, 800);
        mainWindow?.webContents.send('display-warning',
          { msg: 'Projection moved to another external display.' });
      } else {
        // No external display left — close projection so it doesn't cover the laptop screen.
        // Close is safer than minimizing because the guard interval would restore it.
        try {
          projectionWindow.setAlwaysOnTop(false);
          // Step 1: exit fullscreen first and wait for GPU to release
          projectionWindow.setFullScreen(false);
          setTimeout(() => {
            try {
              if (!projectionWindow || projectionWindow.isDestroyed()) return;
              // Step 2: move to primary display before closing
              // This prevents the GPU from reassigning context to mainWindow mid-close
              const pb = screen.getPrimaryDisplay().bounds;
              projectionWindow.setBounds({
                x: pb.x + 100, y: pb.y + 100,
                width: 400, height: 300
              });
              projectionWindow.minimize();
              // Step 3: close after GPU has settled
              setTimeout(() => {
                try {
                  if (projectionWindow && !projectionWindow.isDestroyed()) {
                    projectionWindow.close();
                  }
                } catch(_) {}
              }, 300);
            } catch(e) {
              try { projectionWindow.close(); } catch(_) {}
            }
          }, 350);
        } catch(e) {
          try { projectionWindow.close(); } catch(_) {}
        }
        // Bring main window to front now that projection is closing
        // Staggered recovery: let projection fully close before restoring main window
        const _recoverMainWindow = () => {
          try {
            if (!mainWindow || mainWindow.isDestroyed()) return;
            // Step 1: restore from minimized/hidden state
            if (mainWindow.isMinimized()) mainWindow.restore();
            mainWindow.show();
            // Step 2: bring to front using app-level focus
            app.focus({ steal: true });
            mainWindow.focus();
            // Step 3: ensure it's not behind anything
            mainWindow.setAlwaysOnTop(true);
            setTimeout(() => {
              try {
                if (mainWindow && !mainWindow.isDestroyed()) {
                  mainWindow.setAlwaysOnTop(false);
                  mainWindow.focus();
                  // Step 4: if still white/unresponsive, trigger a soft reload
                  if (mainWindow.webContents.isLoadingMainFrame() === false) {
                    mainWindow.webContents.executeJavaScript(
                      'document.body ? "ok" : "blank"'
                    ).then(result => {
                      if (result !== 'ok') {
                        console.log('[Display] Main window blank — reloading');
                        mainWindow.webContents.reload();
                      }
                    }).catch(() => {
                      mainWindow.webContents.reload();
                    });
                  }
                }
              } catch(e) {}
            }, 800);
          } catch(e) {
            console.error('[Display] Recovery error:', e);
          }
        };
        // Wait for projection window to fully close before recovering
        setTimeout(_recoverMainWindow, 600);
        // Unpause the guard after full recovery (projection is gone so guard will self-clear)
        setTimeout(() => { _projGuardPaused = false; }, 2000);
        mainWindow?.webContents.send('display-warning', {
          msg: 'External display disconnected — projection closed. Reconnect the display and press Project Live again.',
          level: 'error'
        });
      }
    });

    screen.on('display-added', (_evt, newDisplay) => {
      if (!projectionWindow || projectionWindow.isDestroyed()) return;
      if (!projectionWindow._lostTargetDisplay) return;
      if (newDisplay.id === screen.getPrimaryDisplay().id) return;
      projectionWindow._lostTargetDisplay = false;
      const b = newDisplay.bounds;
      projectionWindow._targetDisplayId = newDisplay.id;
      projectionWindow._displayScaleFactor = newDisplay.scaleFactor || 1;
      projectionWindow._displayWidth = b.width;
      projectionWindow._displayHeight = b.height;
      projectionWindow.setBounds({ x: b.x, y: b.y, width: b.width, height: b.height });

      // Step 1: go fullscreen after display settles (800ms)
      setTimeout(() => {
        try {
          if (projectionWindow && !projectionWindow.isDestroyed()) {
            projectionWindow.setFullScreen(true);
          }
        } catch(e) {}
        _projGuardPaused = false;
      }, 800);

      // Step 2: re-push the last render state so media/verse/song reappears
      // Must wait for fullscreen + renderer to be ready to receive IPC (1400ms)
      setTimeout(() => {
        try {
          if (projectionWindow && !projectionWindow.isDestroyed()) {
            console.log('[Display] Restoring render state after reconnect:', currentRenderState?.module);
            projectionWindowReadySync();
            mainWindow?.webContents.send('projection-display-restored');
          }
        } catch(e) {
          console.error('[Display] Failed to restore render state:', e);
        }
      }, 1400);

      mainWindow?.webContents.send('display-warning', {
        msg: 'External display reconnected — content restored.',
        level: 'info'
      });
    });

    // Guard: if main window becomes unresponsive after any display change, recover it
    screen.on('display-metrics-changed', () => {
      // Pause guard during metrics change too — display bounds are shifting
      _projGuardPaused = true;
      setTimeout(() => {
        _projGuardPaused = false;
        try {
          if (mainWindow && !mainWindow.isDestroyed() && !mainWindow.isFocused()) {
            // Only recover if no projection window is active
            if (!projectionWindow || projectionWindow.isDestroyed()) {
              mainWindow.show();
              mainWindow.focus();
            }
          }
        } catch(e) {}
      }, 600);
    });
  }

  projectionWindow.on('closed',()=>{
    const wasDisplayRemoval = projectionWindow?._lostTargetDisplay === true;
    projectionWindow=null;

    if (!wasDisplayRemoval) {
      // User closed projection deliberately — clear the live state
      currentLiveVerse=null;
      currentBackgroundMedia=null;
      buildRenderState('clear', null, { backgroundMedia: null });
    } else {
      // Display was removed — preserve currentRenderState so it can be
      // restored when the display is reconnected and projection reopens
      console.log('[Display] Projection closed by display removal — preserving render state:', currentRenderState?.module);
    }

    mainWindow?.webContents.send('projection-closed');
    if (projectionPowerBlockerId !== null && powerSaveBlocker.isStarted(projectionPowerBlockerId)) {
      powerSaveBlocker.stop(projectionPowerBlockerId);
      projectionPowerBlockerId = null;
    }
  });
  mainWindow?.webContents.send('projection-opened',{displayId:target.id});
}

function createHistoryWindow(){
  if(historyWindow){historyWindow.focus();return;}
  historyWindow=new BrowserWindow({
    icon: APP_ICON,
    width:960,height:680,parent:mainWindow,
    title:'AnchorCast — History',backgroundColor:'#0a0a0f',
    show: false,
    webPreferences:{nodeIntegration:false,contextIsolation:true,preload:path.join(__dirname,'preload.js')},
  });
  if(rendererPort){ historyWindow.loadURL(`http://127.0.0.1:${rendererPort}/history.html`); }
  else { historyWindow.loadFile(path.join(__dirname,'renderer','history.html')); }
  historyWindow.once('ready-to-show',()=>{ try{ historyWindow.show(); }catch(e){} });
  historyWindow.on('closed',()=>{historyWindow=null;});
  historyWindow.setMenu(null);
}

let settingsWindow = null;
let _settingsStartSection = '';
let _settingsOpenData = null;
function createSettingsWindow(startSection=''){
  // Open the full settings.html as a proper window
  if(settingsWindow && !settingsWindow.isDestroyed()){
    if (settingsWindow.isMinimized()) settingsWindow.restore();
    if (!settingsWindow.isVisible()) settingsWindow.show();
    settingsWindow.focus();
    return settingsWindow;
  }
  const win=new BrowserWindow({
    icon: APP_ICON,
    width:680, height:600,
    parent:mainWindow, modal:false,
    title:'AnchorCast — Settings',
    backgroundColor:'#0a0a0f',
    show: false,
    resizable:true, minWidth:600, minHeight:520,
    webPreferences:{nodeIntegration:false,contextIsolation:true,preload:path.join(__dirname,'preload.js')},
  });
  settingsWindow = win;
  // Store the section so settings.html can read it via getSettingsOpenParams
  _settingsStartSection = startSection || '';
  const hash = startSection === 'bible' ? '#bible-manager' : '';
  if(rendererPort){ win.loadURL(`http://127.0.0.1:${rendererPort}/settings.html${hash}`); }
  else {
    const settingsFile = path.join(__dirname,'renderer','settings.html');
    if (startSection === 'bible') { win.loadFile(settingsFile, { hash: 'bible-manager' }); }
    else { win.loadFile(settingsFile); }
  }
  win.once('ready-to-show',()=>{ try{ win.show(); }catch(e){} });
  win.setMenu(null);
  win.on('closed',()=>{ if (settingsWindow === win) settingsWindow = null; });
  return win;
}

function createBibleManagerWindow(){
  const win = new BrowserWindow({
    icon: APP_ICON,
    width: 980, height: 760,
    title: 'AnchorCast — Bible Manager',
    backgroundColor: '#0a0a0f',
    show: false,
    resizable: true, minWidth: 760, minHeight: 600,
    webPreferences:{nodeIntegration:false,contextIsolation:true,preload:path.join(__dirname,'preload.js')},
  });
  if(rendererPort){ win.loadURL(`http://127.0.0.1:${rendererPort}/bible-manager.html`); }
  else { win.loadFile(path.join(__dirname,'renderer','bible-manager.html')); }
  win.once('ready-to-show',()=>{ try{ win.show(); }catch(e){} });
  win.setMenu(null);
  win.on('closed',()=>{ if (settingsWindow === win) settingsWindow = null; });
  return win;
}


let adaptiveManagementWindow = null;
function createAdaptiveManagementWindow(){
  if(adaptiveManagementWindow && !adaptiveManagementWindow.isDestroyed()){
    if (adaptiveManagementWindow.isMinimized()) adaptiveManagementWindow.restore();
    if (!adaptiveManagementWindow.isVisible()) adaptiveManagementWindow.show();
    adaptiveManagementWindow.focus();
    return adaptiveManagementWindow;
  }
  const win = new BrowserWindow({
    icon: APP_ICON,
    width: 1180,
    height: 820,
    title: 'AnchorCast — Adaptive Management',
    backgroundColor: '#0a0a0f',
    resizable: true,
    minWidth: 980,
    minHeight: 680,
    show: false,
    webPreferences:{nodeIntegration:false,contextIsolation:true,preload:path.join(__dirname,'preload.js')},
  });
  adaptiveManagementWindow = win;
  if(rendererPort){ win.loadURL(`http://127.0.0.1:${rendererPort}/adaptive-management.html`); }
  else { win.loadFile(path.join(__dirname,'renderer','adaptive-management.html')); }
  win.once('ready-to-show',()=>{ try{ win.show(); }catch(e){} });
  win.setMenu(null);
  win.on('closed',()=>{ if (adaptiveManagementWindow === win) adaptiveManagementWindow = null; });
  return win;
}

function createHelpWindow(){
  const win=new BrowserWindow({
    icon: APP_ICON,
    width:960, height:700,
    title:'AnchorCast — Help & Documentation',
    backgroundColor:'#08080f',
    show: false,
    resizable:true, minWidth:700, minHeight:500,
    webPreferences:{nodeIntegration:false,contextIsolation:true,preload:path.join(__dirname,'preload.js')},
  });
  if(rendererPort){ win.loadURL(`http://127.0.0.1:${rendererPort}/help_fonts.html`); }
  else { win.loadFile(path.join(__dirname,'renderer','help_fonts.html')); }
  win.once('ready-to-show',()=>{ try{ win.show(); }catch(e){} });
  win.setMenu(null);
  win.on('closed',()=>{ if (settingsWindow === win) settingsWindow = null; });
  return win;
}

ipcMain.handle('open-help-window', () => { createHelpWindow(); });

let songManagerWindow = null;
function createSongManagerWindow(){
  if(songManagerWindow && !songManagerWindow.isDestroyed()){
    if (songManagerWindow.isMinimized()) songManagerWindow.restore();
    if (!songManagerWindow.isVisible()) songManagerWindow.show();
    songManagerWindow.focus();
    return;
  }
  songManagerWindow = new BrowserWindow({
    icon: APP_ICON,
    width:1100, height:720,
    title:'AnchorCast — Song Manager',
    backgroundColor:'#08080f',
    show: false,
    resizable:true, minWidth:800, minHeight:550,
    webPreferences:{nodeIntegration:false,contextIsolation:true,preload:path.join(__dirname,'preload.js')},
  });
  if(rendererPort){ songManagerWindow.loadURL(`http://127.0.0.1:${rendererPort}/song-manager.html`); }
  else { songManagerWindow.loadFile(path.join(__dirname,'renderer','song-manager.html')); }
  songManagerWindow.once('ready-to-show',()=>{ try{ songManagerWindow.show(); }catch(e){} });
  songManagerWindow.setMenu(null);
  // When song manager closes, notify main window to reload songs
  songManagerWindow.on('closed', () => {
    songManagerWindow = null;
    mainWindow?.webContents.send('songs-saved');
  });
}

function createThemeWindow(data){
  if(themeWindow){
    themeWindow.focus();
    // Window already open - send params so it can switch category/theme
    if(data) themeWindow.webContents.send('theme-designer-open', data);
    return;
  }
  themeWindow=new BrowserWindow({
    icon: APP_ICON,
    width:1200,height:780,parent:mainWindow,
    title:'Theme Designer — AnchorCast',backgroundColor:'#08080f',
    show: false,
    minWidth:960,minHeight:640,
    webPreferences:{nodeIntegration:false,contextIsolation:true,preload:path.join(__dirname,'preload.js')},
  });
  if(rendererPort){ themeWindow.loadURL(`http://127.0.0.1:${rendererPort}/theme-designer.html`); }
  else { themeWindow.loadFile(path.join(__dirname,'renderer','theme-designer.html')); }
  themeWindow.once('ready-to-show',()=>{ try{ themeWindow.show(); }catch(e){} });
  themeWindow.on('closed',()=>{themeWindow=null;_themeDesignerOpenData=null;});
  themeWindow.setMenu(null);
  // Params are stored in _themeDesignerOpenData and queried by renderer after init
}

let presEditorWindow = null;

function createPresentationEditorWindow(data){
  if(presEditorWindow){ presEditorWindow.focus(); return; }
  presEditorWindow = new BrowserWindow({
    icon: APP_ICON,
    width: 1280, height: 800, parent: mainWindow,
    title: 'Presentation Editor — AnchorCast',
    backgroundColor: '#0d0d1a', minWidth: 960, minHeight: 600,
    show: false,
    webPreferences: { nodeIntegration:false, contextIsolation:true, preload:path.join(__dirname,'preload.js') },
  });
  if(rendererPort){ presEditorWindow.loadURL(`http://127.0.0.1:${rendererPort}/presentation-editor.html`); }
  else { presEditorWindow.loadFile(path.join(__dirname,'renderer','presentation-editor.html')); }
  presEditorWindow.setMenu(null);
  presEditorWindow.once('ready-to-show', () => {
    try { presEditorWindow.show(); } catch(e) {}
    presEditorWindow.webContents.send('pres-editor-load', data || { id:null, name:'New Presentation', slides:[] });
  });
  presEditorWindow.on('closed', () => { presEditorWindow = null; });
}

ipcMain.handle('pick-pres-file', async (_, { pptxOnly = false } = {}) => {
  const filters = pptxOnly
    ? [{ name: 'PowerPoint', extensions: ['pptx','ppt'] }]
    : [{ name: 'Presentation / PDF', extensions: ['pptx','ppt','odp','pdf'] }];
  const result = await dialog.showOpenDialog(presEditorWindow || mainWindow, {
    title: pptxOnly ? 'Import PowerPoint File' : 'Import Presentation File',
    filters, properties: ['openFile'],
  });
  if (result.canceled || !result.filePaths.length) return { canceled: true };
  return { filePath: result.filePaths[0] };
});

ipcMain.handle('pick-bg-media', async (_, { type = 'image' } = {}) => {
  const filters = type === 'video'
    ? [{ name: 'Video', extensions: ['mp4','mov','webm','avi','mkv','m4v'] }]
    : [{ name: 'Image', extensions: ['jpg','jpeg','png','gif','webp','bmp'] }];
  const win = themeWindow || mainWindow;
  const result = await dialog.showOpenDialog(win, {
    title: type === 'video' ? 'Select Background Video' : 'Select Background Image',
    filters, properties: ['openFile'],
  });
  if (result.canceled || !result.filePaths.length) return { canceled: true };
  return { filePath: result.filePaths[0] };
});

ipcMain.handle('open-pres-editor', (_, data) => {
  createPresentationEditorWindow(data);
  return { success: true };
});

ipcMain.on('pres-editor-saved', (_, pres) => {
  mainWindow?.webContents.send('pres-editor-saved', pres);
});
ipcMain.on('presentation-add-to-schedule', (_evt, payload) => {
  mainWindow?.webContents.send('presentation-add-to-schedule', payload || {});
});

// ── Menu ──────────────────────────────────────────────────────────────────────
function buildMenu(){
  const t=[
    // ── File ──────────────────────────────────────────────────────────────
    {label:'File',submenu:[
      {label:'New Schedule',accelerator:'CmdOrCtrl+N',
        click:()=>mainWindow?.webContents.send('menu-schedule-new')},
      {type:'separator'},
      {label:'Save Schedule',accelerator:'CmdOrCtrl+S',
        click:()=>mainWindow?.webContents.send('menu-schedule-save')},
      {label:'Save Schedule As…',accelerator:'CmdOrCtrl+Shift+S',
        click:()=>mainWindow?.webContents.send('menu-schedule-save-as')},
      {type:'separator'},
      {label:'Open Schedule…',accelerator:'CmdOrCtrl+O',
        click:()=>mainWindow?.webContents.send('menu-schedule-open')},
      {label:'Recent Schedules',submenu: _getRecentScheduleItems()},
      {type:'separator'},
      {label:'Export Schedule to File…',
        click:()=>mainWindow?.webContents.send('menu-schedule-export')},
      {label:'Import Schedule from File…',
        click:()=>mainWindow?.webContents.send('menu-schedule-import')},
      {type:'separator'},
      {label:'Save Preset…',
        click:()=>mainWindow?.webContents.send('menu-preset-save')},
      {label:'Load Preset…',
        click:()=>mainWindow?.webContents.send('menu-preset-load')},
      {label:'Manage Presets…',
        click:()=>mainWindow?.webContents.send('menu-preset-manage')},
      {type:'separator'},
      {label:'Exit',accelerator:'Alt+F4',role:'quit'},
    ]},
    // ── Settings ──────────────────────────────────────────────────────────
    {label:'Settings',submenu:[
      {label:'Preferences…',accelerator:'CmdOrCtrl+,',click:()=>createSettingsWindow()},
    ]},
    // ── Tools ─────────────────────────────────────────────────────────────
    {label:'Tools',submenu:[
      {label:'Song Manager',accelerator:'CmdOrCtrl+M',click:()=>createSongManagerWindow()},
      {label:'Bible Manager',accelerator:'CmdOrCtrl+B',click:createBibleManagerWindow},
      {label:'Theme Designer',accelerator:'CmdOrCtrl+T',click:createThemeWindow},
      {label:'Presentation Editor',accelerator:'CmdOrCtrl+E',click:createPresentationEditorWindow},
      {label:'Sermon History',accelerator:'CmdOrCtrl+H',click:createHistoryWindow},
      {type:'separator'},
      {label:'Generate Sermon Notes',accelerator:'CmdOrCtrl+Shift+N',click:()=>mainWindow?.webContents.send('open-sermon-notes')},
      {type:'separator'},
      {label:'Remote Control URL…',click:showRemoteUrl},
      {label:'External Output…',click:()=>mainWindow?.webContents.send('open-ndi-panel')},
    ]},
    // ── Display ───────────────────────────────────────────────────────────
    {label:'Display',submenu:[
      {label:'Open Projection',accelerator:'CmdOrCtrl+P',click:()=>{
        createProjectionWindow(screen.getAllDisplays()[0].id);
      }},
      {label:'Close Projection',click:()=>{if(projectionWindow)projectionWindow.close();}},
      {type:'separator'},
      {label:'Go Live',accelerator:'CmdOrCtrl+L',click:()=>mainWindow?.webContents.send('shortcut-go-live')},
      {label:'Next Item  →',accelerator:'CmdOrCtrl+Right',click:()=>mainWindow?.webContents.send('shortcut-next')},
      {label:'Prev Item  ←',accelerator:'CmdOrCtrl+Left',click:()=>mainWindow?.webContents.send('shortcut-prev')},
      {label:'Clear Display',accelerator:'CmdOrCtrl+Backspace',click:()=>mainWindow?.webContents.send('shortcut-clear')},
    ]},
    // ── View ──────────────────────────────────────────────────────────────
    {label:'View',submenu:[
      {role:'togglefullscreen'},
      {type:'separator'},
      {label:'Operator Command Center',accelerator:'CmdOrCtrl+Shift+O',click:()=>mainWindow?.webContents.send('menu-show-occ')},
      ...(isDev?[{role:'reload'},{role:'toggleDevTools'}]:[]),
    ]},
    // ── Help ──────────────────────────────────────────────────────────────
    {label:'Help',submenu:[
      {label:'Documentation & Help',accelerator:'F1',click:()=>createHelpWindow()},
      {label:'Get Started / Welcome',click:()=>mainWindow?.webContents.send('menu-show-getstarted')},
      {label:'Keyboard Shortcuts',click:showShortcuts},
      {type:'separator'},
      {label:'Registration',click:()=>createRegistrationStatusWindow()},
      {label:'About AnchorCast',click:showAbout},
    ]},
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(t));
  // Recent schedules are populated inline via _getRecentScheduleItems() above
}

function _getRecentScheduleItems(){
  try{
    const files = _readRecentSchedules()
      .filter(x => x && x.filePath && fs.existsSync(x.filePath))
      .sort((a,b)=>(b.openedAt||0)-(a.openedAt||0))
      .slice(0,8);
    return files.length
      ? files.map(s=>({
          label: (s.name || path.basename(s.filePath, '.json')).slice(0,40) + (((s.name || path.basename(s.filePath, '.json')).length>40)?'…':''),
          click: ()=>mainWindow?.webContents.send('menu-schedule-load-file', s.filePath)
        }))
      : [{label:'No recent schedules',enabled:false}];
  }catch(e){ return [{label:'No recent schedules',enabled:false}]; }
}

function _refreshRecentSchedulesMenu(){
  // Electron menus are immutable after setApplicationMenu —
  // the only way to update a submenu is to rebuild and re-set the whole menu.
  buildMenu();
}
function _readRecentSchedules(){
  try {
    if (!fs.existsSync(RECENT_SCHEDULES_FILE)) return [];
    const list = JSON.parse(fs.readFileSync(RECENT_SCHEDULES_FILE, 'utf-8'));
    return Array.isArray(list) ? list : [];
  } catch (e) { return []; }
}
function _touchRecentSchedule(filePath, name){
  try {
    const resolved = path.resolve(filePath);
    const cur = _readRecentSchedules().filter(x => x && x.filePath && path.resolve(x.filePath) !== resolved);
    cur.unshift({ filePath: resolved, name: name || path.basename(resolved, path.extname(resolved)), openedAt: Date.now() });
    fs.mkdirSync(path.dirname(RECENT_SCHEDULES_FILE), { recursive: true });
    fs.writeFileSync(RECENT_SCHEDULES_FILE, JSON.stringify(cur.slice(0, 12), null, 2));
  } catch (e) {}
  _refreshRecentSchedulesMenu();
}

// ── HTTP Remote Control Server ────────────────────────────────────────────────

// Returns all usable IPv4 adapters, scored and labeled for the UI picker
function getAllNetworkAdapters(){
  const nets = os.networkInterfaces();
  const adapters = [];
  for(const [name, addrs] of Object.entries(nets)){
    for(const n of addrs){
      if(n.family !== 'IPv4' || n.internal) continue;
      const nm = name.toLowerCase();
      // Skip obviously virtual/loopback adapters
      const isVirtual = /vbox|vmware|docker|loopback|tap|tun|veth|br-|virbr/i.test(nm) ||
                        /vethernet|default switch|wsl/i.test(name);
      const isWifi     = /^wi-fi$|wifi|wireless|wlan|wl[a-z0-9]+/i.test(name);
      const isEthernet = /^ethernet$|^eth[0-9]|^en[0-9]|local area connection/i.test(name);
      const label = isWifi ? `📶 ${name} (WiFi)` :
                    isEthernet ? `🔌 ${name} (Ethernet)` :
                    isVirtual  ? `⚙ ${name} (Virtual)` :
                                 `🌐 ${name}`;
      adapters.push({
        name,
        ip: n.address,
        label: `${label} — ${n.address}`,
        isVirtual,
        isWifi,
        isEthernet,
        score: isWifi ? 3 : isEthernet ? 2 : isVirtual ? 0 : 1,
      });
    }
  }
  return adapters.sort((a,b) => b.score - a.score);
}

function localIp(){
  // If user has manually selected an adapter, use that IP
  if(currentSettings.networkAdapter){
    const adapters = getAllNetworkAdapters();
    const saved = adapters.find(a => a.name === currentSettings.networkAdapter);
    if(saved) return saved.ip;
  }
  // Auto-select: highest scored non-virtual adapter
  const adapters = getAllNetworkAdapters();
  const best = adapters.find(a => !a.isVirtual) || adapters[0];
  return best?.ip || '127.0.0.1';
}


const remoteAuthState = { failures: new Map() };
const remoteRuntimeStatus = { lastSeenAt: 0, lastRole: null, lastIp: null };

const remoteSessionTokens = new Map();
function mintRemoteSessionToken(role='admin', ttlMs=(4*60*60*1000)){ // 4h session
  const token = crypto.randomBytes(24).toString('hex');
  const expiresAt = Date.now() + ttlMs;
  remoteSessionTokens.set(token, { role, expiresAt });
  return { token, expiresAt };
}
function getRemoteTokenFromReq(req){
  const header = String(req.headers['x-remote-token'] || '').trim();
  if (header) return header;
  try {
    const u = new URL(req.url || '', 'http://localhost');
    return String(u.searchParams.get('token') || '').trim();
  } catch (_) { return ''; }
}
function remoteRoleForToken(token){
  const key = String(token || '').trim();
  const entry = remoteSessionTokens.get(key);
  if (!entry) return null;
  if (entry.expiresAt <= Date.now()) { remoteSessionTokens.delete(key); return null; }
  return entry.role || null;
}
function makeQrFallbackDataUri(text=''){
  try {
    const safe = String(text || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    const hash = crypto.createHash('sha256').update(String(text || '')).digest();
    const size = 29, cell = 7, pad = 14, canvas = pad*2 + size*cell + 78;
    const finder = new Set();
    const markFinder = (ox, oy) => {
      for (let y=0;y<7;y++) for (let x=0;x<7;x++) finder.add((oy+y)+':' + (ox+x));
    };
    markFinder(0,0); markFinder(size-7,0); markFinder(0,size-7);
    const rects = [];
    const drawFinder = (ox, oy) => {
      rects.push(`<rect x="${pad + ox*cell}" y="${pad + oy*cell}" width="${7*cell}" height="${7*cell}" fill="#111"/>`);
      rects.push(`<rect x="${pad + (ox+1)*cell}" y="${pad + (oy+1)*cell}" width="${5*cell}" height="${5*cell}" fill="#fff"/>`);
      rects.push(`<rect x="${pad + (ox+2)*cell}" y="${pad + (oy+2)*cell}" width="${3*cell}" height="${3*cell}" fill="#111"/>`);
    };
    drawFinder(0,0); drawFinder(size-7,0); drawFinder(0,size-7);
    let bitIdx = 0;
    for (let y=0;y<size;y++){
      for (let x=0;x<size;x++){
        if (finder.has(y+':'+x)) continue;
        const byte = hash[Math.floor(bitIdx/8) % hash.length];
        const on = ((byte >> (bitIdx % 8)) & 1) === 1;
        if (on) rects.push(`<rect x="${pad + x*cell}" y="${pad + y*cell}" width="${cell}" height="${cell}" fill="#111"/>`);
        bitIdx++;
      }
    }
    const line1 = safe.slice(0, 44);
    const line2 = safe.slice(44, 88);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${canvas}" height="${canvas}" viewBox="0 0 ${canvas} ${canvas}"><rect width="${canvas}" height="${canvas}" rx="18" fill="#ffffff"/><rect x="6" y="6" width="${canvas-12}" height="${canvas-12}" rx="14" fill="#ffffff" stroke="#d4af37" stroke-width="2"/>${rects.join('')}<text x="${canvas/2}" y="${pad + size*cell + 28}" text-anchor="middle" font-size="14" font-family="Segoe UI, Arial, sans-serif" fill="#111" font-weight="700">AnchorCast role link</text><text x="${canvas/2}" y="${pad + size*cell + 48}" text-anchor="middle" font-size="10" font-family="Segoe UI, Arial, sans-serif" fill="#444">${line1}</text><text x="${canvas/2}" y="${pad + size*cell + 62}" text-anchor="middle" font-size="10" font-family="Segoe UI, Arial, sans-serif" fill="#444">${line2}</text></svg>`;
    return 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(svg);
  } catch (_) { return ''; }
}
function makeQrDataUri(text=''){
  try {
    const py = [
      'import sys, io, base64',
      'try:',
      ' import qrcode',
      ' from PIL import Image',
      'except Exception as e:',
      ' sys.exit(2)',
      'data = sys.argv[1]',
      'qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=8, border=2)',
      'qr.add_data(data)',
      'qr.make(fit=True)',
      'img = qr.make_image(fill_color="black", back_color="white").convert("RGB")',
      'buf = io.BytesIO()',
      'img.save(buf, format="PNG")',
      'sys.stdout.write("data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii"))'
    ].join('\n');
    const appDir = app.getAppPath();
    const candidates = [
      path.join(appDir, 'python', 'python.exe'),
      path.join(process.cwd(), 'python', 'python.exe'),
      path.join(path.dirname(appDir), 'python', 'python.exe'),
      'py',
      'python3',
      'python'
    ];
    for (const bin of candidates) {
      try {
        const args = (bin === 'py') ? ['-3', '-c', py, String(text || '')] : ['-c', py, String(text || '')];
        const res = spawnSync(bin, args, { encoding:'utf8', maxBuffer: 8 * 1024 * 1024, windowsHide:true, timeout:6000 });
        const out = String(res.stdout || '').trim();
        if (out.startsWith('data:image/png;base64,')) return out;
      } catch (_) {}
    }
  } catch (_) {}
  return makeQrFallbackDataUri(text);
}
function makeRoleRemoteLinks(){
  const ip = localIp();
  const port = currentSettings.httpPort || 8080;
  const roles = ['admin','scripture','songs','media','monitor'];
  const out = {};
  for (const role of roles) {
    const minted = mintRemoteSessionToken(role, 12*60*60*1000);
    const url = `http://${ip}:${port}/remote?role=${role}&token=${minted.token}`;
    out[role] = { role, token: minted.token, expiresAt: minted.expiresAt, url, qrDataUri: (makeQrDataUri(url) || makeQrFallbackDataUri(url)) };
  }
  return out;
}

function getRemoteAuthHeader(req){
  return String(req.headers['x-remote-pin'] || req.headers['x-anchorcast-pin'] || '').trim();
}
function getRemotePinFromUrl(reqUrl=''){
  try {
    const u = new URL(reqUrl, 'http://localhost');
    return String(u.searchParams.get('pin') || '').trim();
  } catch (_) { return ''; }
}
function isRemoteAuthorized(req){
  return !!getAuthorizedRemoteRole(req);
}
function recordRemoteAuthFailure(req){
  const ip = String(req.socket?.remoteAddress || 'unknown');
  const now = Date.now();
  const entry = remoteAuthState.failures.get(ip) || { count:0, last:0 };
  // Reset count if last failure was more than 15 minutes ago
  entry.count = (now - entry.last > 15*60*1000) ? 1 : entry.count + 1;
  entry.last = now;
  remoteAuthState.failures.set(ip, entry);
  return entry.count;
}
function remoteLockedOut(req){
  const ip = String(req.socket?.remoteAddress || 'unknown');
  const entry = remoteAuthState.failures.get(ip);
  if(!entry) return false;
  // 15 attempts before lockout, 5 minute cooldown
  return entry.count >= 15 && (Date.now() - entry.last) < 5*60*1000;
}
function clearRemoteAuthFailures(req){
  const ip = String(req.socket?.remoteAddress || 'unknown');
  remoteAuthState.failures.delete(ip);
}

function normalizeRemotePinValue(v){
  return String(v || '').replace(/\D/g,'').slice(0,8);
}
function isRemoteAuthRequired(v){
  if (v === false || v === 0) return false;
  const s = String(v ?? '').trim().toLowerCase();
  if (!s) return true;
  return !(['false','0','no','none','off','disabled','no authentication','no_authentication','no-authentication'].includes(s));
}
function roleCapabilities(role){
  switch(String(role||'')){
    case 'admin': return ['status','go-live','clear','next','prev','scripture','songs','media','queue-add','set-translation','library'];
    case 'scripture': return ['status','go-live','clear','next','prev','scripture','queue-add','set-translation'];
    case 'songs': return ['status','go-live','clear','next','prev','songs','library'];
    case 'media': return ['status','go-live','clear','media','library'];
    case 'monitor': return ['status'];
    default: return [];
  }
}
function remoteRoleForPin(pin){
  const p = normalizeRemotePinValue(pin);
  if(!p) return null;
  const adminPin = normalizeRemotePinValue(currentSettings.remoteAdminPin || currentSettings.remotePin || '');
  const scripturePin = normalizeRemotePinValue(currentSettings.remoteScripturePin || '');
  const songsPin = normalizeRemotePinValue(currentSettings.remoteSongsPin || '');
  const mediaPin = normalizeRemotePinValue(currentSettings.remoteMediaPin || '');
  const monitorPin = normalizeRemotePinValue(currentSettings.remoteMonitorPin || '');
  if(adminPin && p === adminPin) return 'admin';
  if(scripturePin && p === scripturePin) return 'scripture';
  if(songsPin && p === songsPin) return 'songs';
  if(mediaPin && p === mediaPin) return 'media';
  if(monitorPin && p === monitorPin) return 'monitor';
  return null;
}

function markRemoteActivity(req, role='admin'){
  try{
    remoteRuntimeStatus.lastSeenAt = Date.now();
    remoteRuntimeStatus.lastRole = role || 'admin';
    remoteRuntimeStatus.lastIp = String(req?.socket?.remoteAddress || '').trim() || null;
  }catch(_){}
}

function getAuthorizedRemoteRole(req){
  if(currentSettings.remoteEnabled === false) return null;
  // Session token check (QR code / URL token)
  const roleFromToken = remoteRoleForToken(getRemoteTokenFromReq(req));
  if (roleFromToken) { markRemoteActivity(req, roleFromToken); return roleFromToken; }
  // Auth required check — if auth is disabled, allow unauthenticated access
  // IMPORTANT: only skip auth when remoteRequireAuth is explicitly false AND
  // no admin PIN is configured (prevents bypass when PIN is set but field is missing)
  // Re-evaluate auth requirement live from settings (never cache this)
  const rawAuthSetting = currentSettings.remoteRequireAuth;
  const authRequired = isRemoteAuthRequired(rawAuthSetting);
  const hasAnyPin = !!(normalizeRemotePinValue(currentSettings.remoteAdminPin || currentSettings.remotePin || ''));

  // SECURITY: if a PIN is configured, ALWAYS require it regardless of remoteRequireAuth flag
  // This prevents the case where the UI toggle is off but a PIN is set — PIN always wins
  if (hasAnyPin && !getAuthorizedRemoteRole._bypassPinCheck) {
    const role = remoteRoleForPin(getRemoteAuthHeader(req) || getRemotePinFromUrl(req.url || ''));
    if (role) { markRemoteActivity(req, role); return role; }
    // Has PIN configured but none provided / wrong one → deny
    // (even if remoteRequireAuth flag is false — PIN presence overrides)
    const tokenRole = remoteRoleForToken(getRemoteTokenFromReq(req));
    if (tokenRole) { markRemoteActivity(req, tokenRole); return tokenRole; }
    return null;
  }

  if (!authRequired && !hasAnyPin) {
    // Truly open — allow role from header/query (no PIN configured)
    // NOTE: X-Remote-Role header is intentionally NOT trusted as auth here.
    // It is sent by the remote HTML for UI hints only. Only PIN or session token grants access.
    markRemoteActivity(req, 'admin');
    return 'admin';
  }
  // Auth required (or PIN is configured) — must provide correct PIN
  const role = remoteRoleForPin(getRemoteAuthHeader(req) || getRemotePinFromUrl(req.url || ''));
  if (role) { markRemoteActivity(req, role); return role; }
  return null; // Deny — wrong or missing PIN
}
function remoteRoleCan(role, capability){
  return roleCapabilities(role).includes(capability);
}
function getRemoteStatusPayload(role='admin'){
  const liveType = currentRenderState?.module || (currentLiveVerse ? 'scripture' : null);
  const payload  = currentRenderState?.payload || null;
  // Extract song info from render state for the remote active-song indicator
  const songTitle = (liveType === 'song' && payload?.title) ? payload.title : '';
  const songSlide = (liveType === 'song' && payload?.sectionLabel != null) ? payload.sectionLabel : null;
  const preview = buildRemotePreviewPayload();
  return {
    live: !!(currentRenderState?.module && currentRenderState.module !== 'clear'),
    verse: currentLiveVerse,
    projection: projectionWindow!==null,
    role,
    capabilities: roleCapabilities(role),
    liveType,
    liveModule: liveType,
    songTitle,
    songSlide,
    preview,
  };
}

function buildRemotePreviewPayload(){
  const mod = currentRenderState?.module || 'clear';
  const p = currentRenderState?.payload || null;
  if (mod === 'clear' || !p) return { type: 'clear' };
  if (mod === 'scripture' || mod === 'verse') {
    return { type: 'scripture', ref: p.ref || p.reference || '', text: p.text || '' };
  }
  if (mod === 'song') {
    const lines = Array.isArray(p.lines) ? p.lines : [];
    return { type: 'song', title: p.title || '', label: p.sectionLabel || '', lines };
  }
  if (mod === 'media') {
    return { type: 'media', name: p.name || p.title || '', mediaType: p.type || '' };
  }
  if (mod === 'presentation') {
    return { type: 'presentation', slide: p.slideIndex != null ? p.slideIndex + 1 : null };
  }
  return { type: mod };
}

function startHttpServer(port){
  stopHttpServer();
  // Respect remoteEnabled setting
  if(currentSettings.remoteEnabled === false){
    console.log('[AnchorCast] Remote control disabled in settings.');
    mainWindow?.webContents.send('http-server-started', { port, ip: null, disabled: true });
    return;
  }
  httpServer=http.createServer((req,res)=>{
    res.setHeader('Access-Control-Allow-Origin','*');
    res.setHeader('Access-Control-Allow-Methods','GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers','Content-Type, X-Remote-Pin, X-AnchorCast-Pin, X-Remote-Token, X-Remote-Role');
    res.setHeader('Cache-Control','no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma','no-cache');
    res.setHeader('Expires','0');
    if(req.method==='OPTIONS'){res.writeHead(204);res.end();return;}
    const url=req.url.split('?')[0];
    if(req.method==='GET'){
      if(url==='/api/health') return json(res,200,{status:'ok',version:app.getVersion()});
      if(url==='/api/bootstrap'){
        try {
          if(remoteLockedOut(req)) return json(res,429,{error:'Too many failed attempts. Try again later.'});
          let role = getAuthorizedRemoteRole(req);
          if(!role && !isRemoteAuthRequired(currentSettings.remoteRequireAuth)) {
            role = 'admin';
          }
          if(!role && isRemoteAuthRequired(currentSettings.remoteRequireAuth)){
            // Do NOT record failure here — bootstrap is polled every 15s,
            // a stale pin in localStorage would accumulate failures and cause 429.
            return json(res,401,{error:'Authentication required', authRequired:true, version:'r54'});
          }
          const resolvedRole = role || 'admin';
          return json(res,200,{
            ok:true,
            authRequired:isRemoteAuthRequired(currentSettings.remoteRequireAuth),
            role: resolvedRole,
            capabilities: roleCapabilities(resolvedRole),
            live: !!(currentRenderState?.module && currentRenderState.module !== 'clear'),
            projection: projectionWindow!==null,
            version: 'r54'
          });
        } catch (e) {
          console.error('[REMOTE] bootstrap failed', e);
          // SECURITY: never grant role on error — require re-auth
          const authReqOnError = isRemoteAuthRequired(currentSettings.remoteRequireAuth);
          return json(res, authReqOnError ? 500 : 200, {
            ok: !authReqOnError,
            authRequired: authReqOnError,
            role: authReqOnError ? null : 'admin',
            capabilities: authReqOnError ? [] : roleCapabilities('admin'),
            live: false,
            projection: false,
            version: 'r54-fallback',
            error: e?.message || String(e)
          });
        }
      }
      if(url==='/api/status'){
        if(remoteLockedOut(req)) return json(res,429,{error:'Too many failed attempts. Try again later.'});
        let role = getAuthorizedRemoteRole(req);
        if(!role && !isRemoteAuthRequired(currentSettings.remoteRequireAuth)) role = 'admin';
        if(!role){
          // Do NOT record failure — status is polled every 2.5s.
          // A stale pin in localStorage would trigger lockout within seconds.
          return json(res,401,{error:'Authentication required'});
        }
        clearRemoteAuthFailures(req);
        return json(res,200, getRemoteStatusPayload(role));
      }
      if(url==='/api/library'){
        if(remoteLockedOut(req)) return json(res,429,{error:'Too many failed attempts. Try again later.'});
        let role = getAuthorizedRemoteRole(req);
        if(!role && !isRemoteAuthRequired(currentSettings.remoteRequireAuth)) role = 'admin';
        if(!role){
          recordRemoteAuthFailure(req);
          return json(res,401,{error:'Authentication required'});
        }
        clearRemoteAuthFailures(req);
        const kind = String((new URL(req.url,'http://localhost')).searchParams.get('type') || '').toLowerCase();
        if(kind === 'songs'){
          if(!remoteRoleCan(role,'songs') && !remoteRoleCan(role,'library')) return json(res,403,{error:'Role not permitted'});
          const songs = fs.existsSync(SONGS_FILE) ? JSON.parse(fs.readFileSync(SONGS_FILE,'utf8')) : [];
          return json(res,200,{items:(songs||[]).map(s=>({
            id:s.id, title:s.title || 'Untitled', author:s.author || '',
            slides:Array.isArray(s.sections)?s.sections.length:0,
            sections:(s.sections||[]).map((sec,i)=>({
              idx:i,
              label: sec?.label || `Slide ${i+1}`,
              preview: (Array.isArray(sec?.lines)?sec.lines:[]).filter(Boolean).slice(0,2).join(' ')
            }))
          }))});
        }
        if(kind === 'media'){
          if(!remoteRoleCan(role,'media') && !remoteRoleCan(role,'library')) return json(res,403,{error:'Role not permitted'});
          const media = fs.existsSync(MEDIA_FILE) ? JSON.parse(fs.readFileSync(MEDIA_FILE,'utf8')) : [];
          return json(res,200,{items:(media||[]).map(m=>({
            id:m.id, title:m.title || m.name || 'Untitled', type:m.type || '',
            path:m.path || '', ext:m.ext || ''
          }))});
        }
        return json(res,400,{error:'Unknown library type'});
      }
      if(url==='/'||url==='/remote'){
        res.setHeader('Content-Type','text/html');
        res.setHeader('Cache-Control','no-store, no-cache, must-revalidate, proxy-revalidate');
        res.setHeader('Pragma','no-cache');
        res.setHeader('Expires','0');
        res.writeHead(200);
        return res.end(buildRemoteHTML(port));
      }
      return json(res,404,{error:'Not found'});
    }
    if(req.method==='POST'&&url==='/api/control'){
      if(remoteLockedOut(req)) return json(res,429,{error:'Too many failed attempts. Try again later.'});
      let role = getAuthorizedRemoteRole(req);
      // Only fall back to admin when auth is genuinely not required AND no PIN configured
      if(!role && !isRemoteAuthRequired(currentSettings.remoteRequireAuth)) role = 'admin';
      if(!role){
        recordRemoteAuthFailure(req);
        return json(res,401,{error:'Authentication required. Please enter your PIN in the remote control interface.',authRequired:true});
      }
      clearRemoteAuthFailures(req);
      let body='';
      req.on('data',d=>{body+=d;});
      req.on('end',()=>{
        try{ handleRemoteCmd(JSON.parse(body),res,role); }
        catch(e){ json(res,400,{error:'Invalid JSON: '+String(e&&e.message||e)}); }
      });
      return;
    }
    if(req.method==='POST'&&url==='/api/auth'){
      if(remoteLockedOut(req)) return json(res,429,{error:'Too many failed attempts. Try again later.'});
      let body='';
      req.on('data',d=>{body+=d;});
      req.on('end',()=>{
        try{
          const parsed = JSON.parse(body || '{}');
          const pin = String(parsed.pin || '').trim();
          if(!isRemoteAuthRequired(currentSettings.remoteRequireAuth)){
            clearRemoteAuthFailures(req);
            const session = mintRemoteSessionToken('admin');
            return json(res,200,{success:true, authRequired:false, role:'admin', capabilities:roleCapabilities('admin'), token:session.token, expiresAt:session.expiresAt});
          }
          const role = remoteRoleForPin(pin);
          if(role){
            clearRemoteAuthFailures(req);
            const session = mintRemoteSessionToken(role);
            return json(res,200,{success:true, authRequired:true, role, capabilities:roleCapabilities(role), token:session.token, expiresAt:session.expiresAt});
          }
          recordRemoteAuthFailure(req);
          return json(res,401,{success:false,error:'Invalid PIN'});
        } catch(e){ json(res,400,{error:'Invalid JSON'}); }
      });
      return;
    }
    if(req.method==='POST'&&url==='/api/signout'){
      // Revoke the session token (sent as X-Remote-Token by the remote client)
      const tok = String(req.headers['x-remote-token']||'').trim();
      if(tok) remoteSessionTokens.delete(tok);
      clearRemoteAuthFailures(req); // also clear any lockout on sign-out
      return json(res,200,{success:true});
    }
    json(res,404,{error:'Not found'});
  });
  httpServer.listen(port,'0.0.0.0',()=>{
    const ip=localIp();
    console.log(`[AnchorCast] Remote: http://${ip}:${port}/remote`);
    mainWindow?.webContents.send('http-server-started',{port,ip,disabled:false});
  });
  httpServer.on('error',e=>{
    console.warn('[HTTP]',e.message);
    mainWindow?.webContents.send('http-server-started',{port,ip:null,error:e.message});
  });
}
function stopHttpServer(){ if(httpServer){httpServer.close();httpServer=null;} }
function json(res,code,data){ res.writeHead(code,{'Content-Type':'application/json'});res.end(JSON.stringify(data)); }

function handleRemoteCmd(cmd,res,role='admin'){
  if(!mainWindow || (typeof mainWindow.isDestroyed === 'function' && mainWindow.isDestroyed())) return json(res,503,{error:'App not ready'});
  const action = (cmd && cmd.action) || '';
  const data = (cmd && cmd.data) || null;
  const deny = () => json(res,403,{error:'Role not permitted for this action', role});
  const safeSend = (channel, payload) => {
    try {
      if (mainWindow && mainWindow.webContents && !mainWindow.webContents.isDestroyed()) {
        mainWindow.webContents.send(channel, payload);
      }
    } catch(_) {}
  };
  const safeExec = (code) => {
    try {
      if (mainWindow && mainWindow.webContents && !mainWindow.webContents.isDestroyed() && !mainWindow.webContents.isLoading()) {
        mainWindow.webContents.executeJavaScript(code).catch(()=>{});
      }
    } catch(_) {}
  };
  switch(action){
    case 'go-live':
      if(!remoteRoleCan(role,'go-live')) return deny();
      // Use sendPreviewToLive which sets isLive=true and routes by current tab,
      // rather than toggleGoLive which toggles (could turn it OFF if already on).
      safeExec("(function(){ if(window.electronAPI&&window.electronAPI.openProjection) window.electronAPI.openProjection(); if(!State.isLive&&typeof toggleGoLive==='function'){toggleGoLive();}else if(typeof sendPreviewToLive==='function'){sendPreviewToLive();} })();");
      break;
    case 'clear':
      if(!remoteRoleCan(role,'clear')) return deny();
      safeSend('shortcut-clear');
      safeExec("if (typeof clearLive === 'function') clearLive();");
      break;
    case 'next':
      if(!remoteRoleCan(role,'next')) return deny();
      (function(){
        var mode = (data && data.mode) || 'scripture';
        if(mode==='songs'){
          safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='song') switchTab('song'); if(typeof navigateSongSlide==='function') navigateSongSlide(1); })();");
        } else if(mode==='presentation'){
          safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='pres') switchTab('pres'); if(typeof navigatePresSlide==='function') navigatePresSlide(1); })();");
        } else if(mode==='media'){
          safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='media') switchTab('media'); if(typeof moveMediaSelection==='function') moveMediaSelection(1); })();");
        } else {
          // Scripture: step the currently live/previewed verse directly without
          // forcing a desktop tab switch — avoids hijacking the operator's view.
          safeExec("if(typeof remoteStepScripture==='function') remoteStepScripture(1);");
        }
      })();
      break;
    case 'prev':
      if(!remoteRoleCan(role,'prev')) return deny();
      (function(){
        var mode = (data && data.mode) || 'scripture';
        if(mode==='songs'){
          safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='song') switchTab('song'); if(typeof navigateSongSlide==='function') navigateSongSlide(-1); })();");
        } else if(mode==='presentation'){
          safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='pres') switchTab('pres'); if(typeof navigatePresSlide==='function') navigatePresSlide(-1); })();");
        } else if(mode==='media'){
          safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='media') switchTab('media'); if(typeof moveMediaSelection==='function') moveMediaSelection(-1); })();");
        } else {
          safeExec("if(typeof remoteStepScripture==='function') remoteStepScripture(-1);");
        }
      })();
      break;
    case 'scripture-next':
      if(!remoteRoleCan(role,'next')) return deny();
      safeExec("if(typeof remoteStepScripture==='function') remoteStepScripture(1);");
      break;
    case 'scripture-prev':
      if(!remoteRoleCan(role,'prev')) return deny();
      safeExec("if(typeof remoteStepScripture==='function') remoteStepScripture(-1);");
      break;
    case 'song-next':
      if(!(role === 'admin' || remoteRoleCan(role,'songs'))) return deny();
      safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='song') switchTab('song'); if(typeof navigateSongSlide==='function') navigateSongSlide(1); })();");
      break;
    case 'song-prev':
      if(!(role === 'admin' || remoteRoleCan(role,'songs'))) return deny();
      safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='song') switchTab('song'); if(typeof navigateSongSlide==='function') navigateSongSlide(-1); })();");
      break;
    case 'presentation-current':
      if(role !== 'admin') return deny();
      safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='pres') switchTab('pres'); if(typeof presentCurrentPresSlide==='function') presentCurrentPresSlide(); })();");
      break;
    case 'presentation-next':
      if(role !== 'admin') return deny();
      safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='pres') switchTab('pres'); if(typeof navigatePresSlide==='function') navigatePresSlide(1); })();");
      break;
    case 'presentation-prev':
      if(role !== 'admin') return deny();
      safeExec("(function(){ if(typeof switchTab==='function'&&window.State&&State.currentTab!=='pres') switchTab('pres'); if(typeof navigatePresSlide==='function') navigatePresSlide(-1); })();");
      break;
    case 'present':
      if(!remoteRoleCan(role,'scripture')) return deny();
      if(data && data.ref){
        // Route through the renderer's remote-control handler which correctly:
        // 1. _setLiveOn() — ensures isLive=true before presentVerse runs
        // 2. switchTab('book') — so sendPreviewToLive routes to scripture, not song
        // 3. presentVerse() — with tab already 'book', liveContentType stays 'scripture'
        //
        // The old safeExec called presentVerse() directly without switching tabs.
        // With a song live (State.currentTab='song', State.isLive=true):
        //   - presentVerse() → updateLiveDisplay() → sets liveContentType='scripture'
        //     → syncProjection() fires → scripture briefly projects (FLICKER)
        //   - sendPreviewToLive() → sees currentTab='song' → presentSongSlide()
        //     → liveContentType='song' → song projects again (ORIGINAL STAYS)
        // Result: flicker then song stays. Fix: switch tab first via remote-control.
        safeSend('remote-present', data);
        safeSend('remote-control', { action: 'present', data });
      }
      break;
    case 'queue-add':
      if(!remoteRoleCan(role,'queue-add')) return deny();
      if(data && data.ref){
        safeSend('remote-queue-add',data);
        const payload = JSON.stringify(data || {});
        safeExec(`(function(){ var data = ${payload}; if (typeof addToQueue === 'function' && data.book) addToQueue(data.book, data.chapter, data.verse, data.ref); })();`);
      }
      break;
    case 'set-translation':
      if(!remoteRoleCan(role,'set-translation')) return deny();
      if(data && data.translation){
        safeSend('remote-set-translation',data.translation);
        const payload = JSON.stringify(String(data.translation));
        safeExec(`(function(){ var t = ${payload}; try { if (window.State) State.currentTranslation = t; var g=document.getElementById('globalTranslation'); if (g) g.value=t; var s=document.getElementById('searchTranslation'); if (s) s.value=t; if (typeof refreshCanvases === 'function') refreshCanvases(); } catch(_) {} })();`);
      }
      break;
    case 'present-song':
      if(!(role === 'admin' || remoteRoleCan(role,'songs'))) return deny();
      if(data && data.songId != null){
        // Open projection window first so it's ready when the IPC handler projects
        safeExec("if(window.electronAPI&&window.electronAPI.openProjection) window.electronAPI.openProjection();");
        // Route through renderer's remote-control handler:
        // _setLiveOn() → selectSong() → presentSongSlide() → sendSongToProjection()
        // preload.js allowlist now includes 'remote-control' so this reaches the renderer.
        safeSend('remote-control', { action: 'present-song', data });
      }
      break;
    case 'present-media':
      if(!(role === 'admin' || remoteRoleCan(role,'media'))) return deny();
      if(data && data.mediaId != null){
        safeExec("if(window.electronAPI&&window.electronAPI.openProjection) window.electronAPI.openProjection();");
        safeSend('remote-control', { action: 'present-media', data });
      }
      break;
    case 'present-current':
      if(data && data.mode === 'scripture'){
        if(!remoteRoleCan(role,'scripture')) return deny();
        safeExec("(function(){ try { if(typeof switchTab==='function'&&window.State&&State.currentTab!=='book') switchTab('book'); if(window.electronAPI&&window.electronAPI.openProjection) window.electronAPI.openProjection(); if(!State.isLive&&typeof toggleGoLive==='function') toggleGoLive(); else if(typeof sendPreviewToLive==='function') sendPreviewToLive(); } catch(_){} })();");
      } else if(data && data.mode === 'songs'){
        if(!(role === 'admin' || remoteRoleCan(role,'songs'))) return deny();
        // open projection then route through renderer's present-current handler
        safeExec("if(window.electronAPI&&window.electronAPI.openProjection) window.electronAPI.openProjection();");
        safeSend('remote-control', { action: 'present-current', data });
      } else if(data && data.mode === 'media'){
        if(!(role === 'admin' || remoteRoleCan(role,'media'))) return deny();
        safeExec("if(window.electronAPI&&window.electronAPI.openProjection) window.electronAPI.openProjection();");
        safeSend('remote-control', { action: 'present-current', data });
      } else if(data && data.mode === 'presentation'){
        if(role !== 'admin') return deny();
        safeExec("if(window.electronAPI&&window.electronAPI.openProjection) window.electronAPI.openProjection();");
        safeSend('remote-control', { action: 'present-current', data });
      } else {
        if(!remoteRoleCan(role,'go-live')) return deny();
        safeExec("(function(){ if(window.electronAPI&&window.electronAPI.openProjection) window.electronAPI.openProjection(); if(!State.isLive&&typeof toggleGoLive==='function') toggleGoLive(); else if(typeof sendPreviewToLive==='function') sendPreviewToLive(); })();");
      }
      break;
    default:
      return json(res,400,{error:'Unknown action: ' + action});
  }
  json(res,200,{success:true,action,role});
}

// ── Mobile Remote HTML ────────────────────────────────────────────────────────





function buildRemoteHTML(port){
  var authRequired = isRemoteAuthRequired(currentSettings.remoteRequireAuth);
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">
<title>AnchorCast Remote</title>
<style>
:root{
  --gold:#C9A84C;--gold-dim:rgba(201,168,76,.18);--gold-glow:rgba(201,168,76,.07);
  --live:#E74C3C;--live-dim:rgba(231,76,60,.15);
  --bg:#07070E;--surface:#0D0D1A;--card:#111120;--card2:#141428;
  --border:rgba(255,255,255,.07);--border-gold:rgba(201,168,76,.25);
  --text:#E8E8F4;--text-dim:#7070A0;--text-muted:#404060;
  --blue:#4A9EE8;--r:14px;--r-sm:10px;--r-xs:7px;
}
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent;-webkit-font-smoothing:antialiased}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--bg);color:var(--text);max-width:430px;margin:0 auto;padding-bottom:115px}
/* ── Header ── */
.hdr{padding:14px 16px 12px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;background:rgba(7,7,14,.95);backdrop-filter:blur(20px)}
.hdr-brand{display:flex;align-items:center;gap:8px}
.hdr-cross{font-size:18px;color:var(--gold)}
.hdr-title{font-size:13px;font-weight:700;color:var(--gold);letter-spacing:.1em;text-transform:uppercase}
.hdr-sub{font-size:8px;color:var(--text-muted);letter-spacing:.15em;text-transform:uppercase}
.onair{display:flex;align-items:center;gap:5px;padding:4px 9px;border-radius:20px;background:var(--live-dim);border:1px solid rgba(231,76,60,.3);font-size:9px;font-weight:700;color:var(--live);letter-spacing:.08em;display:none}
.onair.show{display:flex}
.onair-dot{width:6px;height:6px;border-radius:50%;background:var(--live);animation:pulse 1.4s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
/* ── Status ── */
.status-bar{margin:12px 14px 0;background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:12px 14px;display:flex;align-items:center;gap:10px;min-height:56px}
.status-icon{width:34px;height:34px;border-radius:50%;background:var(--surface);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0}
.status-what{font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.status-type{font-size:9px;color:var(--text-dim);letter-spacing:.07em;text-transform:uppercase;margin-top:2px}
/* ── Projection Preview ── */
.proj-card{margin:10px 14px 0;border-radius:var(--r);border:1px solid var(--border);overflow:hidden;background:var(--card)}
.proj-hdr{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;background:var(--surface);border-bottom:1px solid var(--border)}
.proj-hdr-left{display:flex;align-items:center;gap:6px;font-size:9px;font-weight:700;color:var(--text-muted);letter-spacing:.12em;text-transform:uppercase}
.proj-dot{width:6px;height:6px;border-radius:50%;background:var(--text-muted);flex-shrink:0}
.proj-dot.on{background:var(--live);animation:pulse 1.4s infinite}
.proj-clear-btn{padding:3px 10px;border-radius:4px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.04);color:var(--text-muted);font-size:8px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;cursor:pointer;font-family:inherit}
.proj-clear-btn:active{background:rgba(255,255,255,.08)}
.proj-screen{position:relative;min-height:140px;background:#05050A;display:flex;align-items:center;justify-content:center;padding:16px;overflow:hidden}
.proj-empty{font-size:12px;color:var(--text-muted);opacity:.5}
.proj-content{width:100%;text-align:center;animation:projFadeIn .25s ease}
@keyframes projFadeIn{from{opacity:0;transform:scale(.97)}to{opacity:1;transform:scale(1)}}
.proj-ref{font-size:11px;font-weight:700;color:var(--gold);letter-spacing:.06em;margin-bottom:6px}
.proj-text{font-size:12px;color:var(--text);line-height:1.6;max-height:100px;overflow:hidden}
.proj-song-title{font-size:9px;font-weight:600;color:var(--gold);letter-spacing:.08em;text-transform:uppercase;margin-bottom:2px}
.proj-song-label{font-size:8px;color:var(--text-muted);margin-bottom:6px}
.proj-song-lines{font-size:12px;color:var(--text);line-height:1.65;max-height:100px;overflow:hidden}
.proj-media-icon{font-size:28px;margin-bottom:6px}
.proj-media-name{font-size:12px;font-weight:600;color:var(--text)}
/* ── Auth ── */
.auth-card{margin:12px 14px 0;background:var(--card);border:1px solid var(--border-gold);border-radius:var(--r);padding:16px}
.pin-row{display:flex;gap:8px;margin-top:10px}
.pin-input{flex:1;background:var(--surface);border:1px solid var(--border);border-radius:var(--r-sm);color:var(--text);font-size:18px;letter-spacing:.2em;padding:11px 12px;outline:none;-webkit-appearance:none}
.pin-input:focus{border-color:var(--border-gold)}
.unlock-btn{padding:11px 16px;background:var(--gold);border:none;border-radius:var(--r-sm);color:#000;font-size:11px;font-weight:700;cursor:pointer;white-space:nowrap}
/* ── Body padding ── */
.body{padding:0 14px}
/* ── Mode tabs ── */
.sec-lbl{font-size:9px;font-weight:600;color:var(--text-muted);letter-spacing:.16em;text-transform:uppercase;padding-top:14px;margin-bottom:6px}
.tabs{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-bottom:12px}
.tab{padding:9px 3px 7px;border-radius:var(--r-sm);border:1px solid var(--border);background:var(--card);color:var(--text-muted);font-size:8px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;cursor:pointer;display:flex;flex-direction:column;align-items:center;gap:4px;transition:all .12s;font-family:inherit}
.tab .ti{font-size:16px;line-height:1}
.tab.active{background:var(--gold-dim);border-color:var(--border-gold);color:var(--gold)}
.tab:disabled{opacity:.35;cursor:not-allowed}
.tab:not(:disabled):active{opacity:.7;transform:scale(.97)}
/* ── Scripture section ── */
.scripture-live{display:none;align-items:center;gap:8px;background:rgba(231,76,60,.08);border:1px solid rgba(231,76,60,.2);border-radius:var(--r-sm);padding:9px 12px;margin-bottom:10px}
.scripture-live.show{display:flex}
.scripture-live-dot{width:6px;height:6px;border-radius:50%;background:var(--live);flex-shrink:0;animation:pulse 1.4s infinite}
.ref-row{display:flex;gap:7px;margin-bottom:10px}
.ref-input{flex:1;background:var(--surface);border:1px solid var(--border);border-radius:var(--r-sm);color:var(--text);font-size:15px;padding:12px 13px;outline:none;-webkit-appearance:none;transition:border-color .2s;font-family:inherit}
.ref-input:focus{border-color:var(--border-gold)}
.ref-input::placeholder{color:var(--text-muted)}
.send-btn{padding:12px 16px;background:var(--gold);border:none;border-radius:var(--r-sm);color:#000;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap;font-family:inherit}
.send-btn:active{opacity:.8}
.quick-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-bottom:4px}
.qbtn{padding:11px 4px;border-radius:var(--r-sm);border:1px solid var(--border-gold);background:var(--gold-glow);color:var(--gold);font-size:10px;font-weight:700;cursor:pointer;text-align:center;line-height:1.35;font-family:inherit;transition:all .12s}
.qbtn:active{background:var(--gold-dim);transform:scale(.96)}
/* ── Songs section ── */
.active-bar{display:none;align-items:center;gap:8px;background:rgba(231,76,60,.07);border:1px solid rgba(231,76,60,.18);border-radius:var(--r-sm);padding:9px 12px;margin-bottom:10px}
.active-bar.show{display:flex}
.active-dot{width:6px;height:6px;border-radius:50%;background:var(--live);flex-shrink:0;animation:pulse 1.4s infinite}
.active-name{font-size:12px;font-weight:600;color:var(--text);flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.active-slide{font-size:10px;color:var(--text-dim);flex-shrink:0}
.songs-list,.media-list{display:flex;flex-direction:column;gap:7px;margin-bottom:4px}
.song-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r-sm);padding:11px 13px}
.song-title{font-size:13px;font-weight:600;color:var(--text);margin-bottom:2px}
.song-meta{font-size:10px;color:var(--text-dim);margin-bottom:9px}
.chips{display:flex;flex-wrap:wrap;gap:5px}
.chip{padding:5px 11px;border-radius:20px;background:var(--card2);border:1px solid var(--border);color:var(--text-dim);font-size:11px;cursor:pointer;transition:all .12s;font-family:inherit}
.chip.active{background:var(--gold-dim);border-color:var(--border-gold);color:var(--gold)}
.chip:active{transform:scale(.95)}
/* ── Media section ── */
.media-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r-sm);padding:11px 13px;display:flex;align-items:center;gap:11px;cursor:pointer;transition:all .12s}
.media-card:active{background:var(--card2)}
.media-thumb{width:38px;height:38px;border-radius:7px;background:var(--card2);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}
.media-info{flex:1;min-width:0}
.media-name{font-size:12px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.media-type{font-size:9px;color:var(--text-dim);letter-spacing:.07em;text-transform:uppercase;margin-top:2px}
.media-play{width:30px;height:30px;border-radius:50%;background:var(--gold-dim);border:1px solid var(--border-gold);color:var(--gold);font-size:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
/* ── Pres section ── */
.pres-note{font-size:11px;color:var(--text-dim);text-align:center;margin-bottom:12px;line-height:1.5}
.pres-nav{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-bottom:9px}
.pres-btn{padding:15px 8px;border-radius:var(--r);border:1px solid var(--border);background:var(--card2);color:var(--text-dim);font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:7px;transition:all .12s;font-family:inherit}
.pres-btn:active{background:var(--surface);transform:scale(.97)}
.pres-present{width:100%;padding:14px;border-radius:var(--r);border:1px solid rgba(74,158,232,.3);background:rgba(74,158,232,.1);color:var(--blue);font-size:12px;font-weight:600;cursor:pointer;font-family:inherit}
.pres-present:active{opacity:.7}
/* ── Broadcast (sticky bottom bar) ── */
.broadcast{position:fixed;bottom:0;left:0;right:0;max-width:430px;margin:0 auto;background:rgba(7,7,14,.97);border-top:1px solid var(--border);padding:10px 14px 12px;z-index:50;display:flex;flex-direction:column;gap:7px}
.go-btn{width:100%;padding:15px;border-radius:var(--r);border:1px solid rgba(231,76,60,.35);background:var(--live-dim);color:var(--live);font-size:14px;font-weight:700;letter-spacing:.12em;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:9px;transition:all .12s;font-family:inherit}
.go-btn.on{background:var(--live);color:#fff;border-color:var(--live)}
.go-btn:active{transform:scale(.98)}
.ctrl-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px}
.ctrl-btn{padding:11px 6px;border-radius:var(--r-sm);border:1px solid var(--border);background:var(--card);color:var(--text-dim);font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:5px;transition:all .12s;font-family:inherit}
.ctrl-btn.wide{grid-column:span 1}
.ctrl-btn .arr{font-size:16px;color:var(--text-muted)}
.ctrl-btn:active{background:var(--card2);transform:scale(.97)}
/* ── Role badge ── */
.role-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 9px;border-radius:20px;background:var(--gold-glow);border:1px solid var(--border-gold);font-size:9px;font-weight:600;color:var(--gold);letter-spacing:.07em;text-transform:uppercase;margin:10px 14px 0;display:none}
.role-badge.show{display:inline-flex}
/* ── Toast ── */
.toast-wrap{position:fixed;bottom:118px;left:50%;transform:translateX(-50%);z-index:999;width:calc(100% - 28px);max-width:400px;pointer-events:none}
#toast{font-size:12px;font-weight:500;padding:10px 14px;border-radius:var(--r-sm);background:rgba(39,174,96,.14);border:1px solid rgba(39,174,96,.32);color:#5de88a;transition:opacity .22s,transform .22s;opacity:0;transform:translateY(6px);text-align:center}
#toast.err{background:rgba(231,76,60,.12);border-color:rgba(231,76,60,.28);color:#ff7b6b}
#toast.show{opacity:1;transform:translateY(0)}
.hidden{display:none}
#authMsg{font-size:10px;color:var(--live);margin-top:7px}
.signout-btn{background:rgba(231,76,60,.12);border:1px solid rgba(231,76,60,.28);border-radius:var(--r-xs);color:#ff7b6b;font-size:10px;font-weight:700;padding:5px 9px;cursor:pointer;letter-spacing:.04em;display:none;font-family:inherit}
.signout-btn.show{display:block}
.signout-btn:active{opacity:.7}
@media(orientation:landscape) and (max-height:500px){
  body{max-width:100%;padding-bottom:72px}
  .hdr{padding:8px 14px 6px}
  .ls-wrap{display:flex;gap:10px;padding:0 10px}
  .ls-left{flex:0 0 42%;max-width:42%;position:sticky;top:50px;align-self:flex-start}
  .ls-left .status-bar{margin:8px 0 0}
  .ls-left .proj-card{margin:6px 0 0}
  .ls-left .proj-screen{min-height:100px}
  .ls-right{flex:1;min-width:0}
  .ls-right .role-badge{margin:8px 0 4px}
  .ls-right .body{padding:0}
  .ls-right .sec-lbl{padding-top:8px;margin-bottom:4px}
  .tabs{gap:4px;margin-bottom:8px}
  .tab{padding:6px 3px 5px;font-size:7px}
  .tab .ti{font-size:14px}
  .quick-grid{grid-template-columns:repeat(6,1fr);gap:4px}
  .qbtn{padding:7px 2px;font-size:9px}
  .ref-row{margin-bottom:6px}
  .ref-input{padding:8px 10px;font-size:13px}
  .send-btn{padding:8px 12px;font-size:11px}
  .broadcast{padding:6px 10px 8px;gap:5px;max-width:100%}
  .go-btn{padding:10px;font-size:12px}
  .ctrl-row{gap:4px}
  .ctrl-btn{padding:7px 4px;font-size:11px}
  .song-card{padding:8px 10px}
  .chips{gap:4px}
  .chip{padding:4px 8px;font-size:10px}
  .auth-card{margin:8px 0 0}
}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-brand">
    <div class="hdr-cross">✝</div>
    <div><div class="hdr-title">AnchorCast</div><div class="hdr-sub">Remote</div></div>
  </div>
  <div style="display:flex;align-items:center;gap:8px">
    <div class="onair" id="onairBadge"><div class="onair-dot"></div>ON AIR</div>
    <button id="signOutBtn" class="signout-btn" onclick="signOut()" title="Sign out">SIGN OUT</button>
    <button id="homeBtn" onclick="window.location.href='/'" style="background:var(--card);border:1px solid var(--border);border-radius:var(--r-xs);color:var(--text-dim);font-size:16px;padding:6px 10px;cursor:pointer;display:flex;align-items:center;justify-content:center" title="Back to main app">🏠</button>
  </div>
</div>

<div class="ls-wrap">
<div class="ls-left">

<div class="status-bar">
  <div class="status-icon" id="statusIcon">📡</div>
  <div><div class="status-what" id="statusWhat">Connecting…</div><div class="status-type" id="statusType">Please wait</div></div>
</div>

<div class="proj-card" id="projCard">
  <div class="proj-hdr">
    <div class="proj-hdr-left"><div class="proj-dot" id="projDot"></div>Projection Screen</div>
    <button class="proj-clear-btn" id="projClearBtn">Clear</button>
  </div>
  <div class="proj-screen" id="projScreen">
    <div class="proj-empty" id="projEmpty">Nothing on screen</div>
    <div class="proj-content" id="projContent" style="display:none"></div>
  </div>
</div>

</div>
<div class="ls-right">

<div class="role-badge" id="roleBadge">👤 Admin</div>

<div class="auth-card" id="authBox" style="${authRequired ? '' : 'display:none'}">
  <div style="font-size:12px;font-weight:600;color:var(--gold);margin-bottom:4px">🔐 Secure Access</div>
  <div style="font-size:11px;color:var(--text-dim)">Enter your PIN to unlock</div>
  <div class="pin-row">
    <input class="pin-input" id="pinInput" type="password" inputmode="numeric" placeholder="• • • •" maxlength="8">
    <button class="unlock-btn" id="unlockBtn">UNLOCK</button>
  </div>
  <div id="authMsg"></div>
</div>

<div class="body">
  <div class="sec-lbl">Mode</div>
  <div class="tabs">
    <button class="tab active" id="modeScripture"><span class="ti">📖</span>Scripture</button>
    <button class="tab" id="modeSongs"><span class="ti">🎵</span>Songs</button>
    <button class="tab" id="modeMedia"><span class="ti">🎬</span>Media</button>
    <button class="tab" id="modePresentation"><span class="ti">📑</span>Slides</button>
  </div>

  <!-- Scripture -->
  <div class="hidden" id="scriptureSec">
    <div class="scripture-live" id="scriptureNowLive">
      <div class="scripture-live-dot"></div>
      <div style="font-size:12px;font-weight:600;color:var(--text);flex:1" id="scriptureNowLiveRef">—</div>
      <div style="font-size:9px;color:var(--text-dim)" id="scriptureNowLiveType">LIVE</div>
    </div>
    <div class="sec-lbl" style="padding-top:0">Search &amp; Send</div>
    <div class="ref-row">
      <input class="ref-input" id="refInput" type="text" placeholder="John 3:16 · Ps 23 · Rom 8:28" autocomplete="off" autocorrect="off" autocapitalize="words" spellcheck="false">
      <button class="send-btn" id="sendBtn">SEND</button>
    </div>
    <div class="sec-lbl" style="padding-top:4px">Quick Verse</div>
    <div class="quick-grid">
      <button class="qbtn" id="smpJohn316">John<br>3:16</button>
      <button class="qbtn" id="smpPsalm23">Psalm<br>23:1</button>
      <button class="qbtn" id="smpPhil413">Phil<br>4:13</button>
      <button class="qbtn" id="smpRom828">Rom<br>8:28</button>
      <button class="qbtn" id="smpIsa4031">Isa<br>40:31</button>
      <button class="qbtn" id="smpJer2911">Jer<br>29:11</button>
    </div>
  </div>

  <!-- Songs -->
  <div class="hidden" id="songsSec">
    <div class="active-bar" id="activeSongBar">
      <div class="active-dot"></div>
      <div class="active-name" id="activeSongLabel">No song live</div>
      <div class="active-slide" id="activeSongSlide"></div>
    </div>
    <div style="display:flex;gap:7px;margin-bottom:10px">
      <input id="songSearch" type="text" placeholder="Search songs…" autocomplete="off" spellcheck="false"
        style="flex:1;background:var(--surface);border:1px solid var(--border);border-radius:var(--r-sm);
        color:var(--text);font-size:14px;padding:10px 12px;outline:none;font-family:inherit;-webkit-appearance:none">
      <button id="songSearchClear"
        style="padding:10px 13px;border-radius:var(--r-sm);border:1px solid var(--border);
        background:var(--card);color:var(--text-muted);font-size:15px;cursor:pointer;font-family:inherit;min-width:40px">✕</button>
    </div>
    <div class="songs-list" id="songsList">
      <div class="song-card"><div class="song-title" style="color:var(--text-dim)">Loading songs…</div></div>
    </div>
  </div>

  <!-- Media -->
  <div class="hidden" id="mediaSec">
    <div class="media-list" id="mediaList">
      <div class="song-card"><div class="song-title" style="color:var(--text-dim)">Loading media…</div></div>
    </div>
  </div>

  <!-- Presentation -->
  <div class="hidden" id="presentationSec">
    <div class="pres-note">Tap the slide arrows to step. Press Present to send live.</div>
    <div class="pres-nav">
      <button class="pres-btn" id="presPrevBtn"><span style="font-size:18px">‹</span> Prev Slide</button>
      <button class="pres-btn" id="presNextBtn">Next Slide <span style="font-size:18px">›</span></button>
    </div>
    <button class="pres-present" id="presCurrentBtn">▶ &nbsp;Present Current Slide</button>
  </div>
</div><!-- /body -->

</div><!-- /ls-right -->
</div><!-- /ls-wrap -->

<!-- Broadcast sticky footer -->
<div class="broadcast">
  <button class="go-btn" id="liveBtn"><span id="liveBtnDot">●</span><span id="liveBtnText">GO LIVE</span></button>
  <div class="ctrl-row">
    <button class="ctrl-btn" id="prevBtn"><span class="arr">‹</span> Prev</button>
    <button class="ctrl-btn" id="clearBtn" style="border-color:rgba(255,255,255,.05);color:var(--text-muted);font-size:11px">✕ Clear</button>
    <button class="ctrl-btn" id="nextBtn">Next <span class="arr">›</span></button>
  </div>
</div>

<div class="toast-wrap"><div id="toast"></div></div>

<script>
(function(){
'use strict';
var AUTH_REQUIRED=${authRequired?'true':'false'};
var role='admin', token='', pin='', caps=[], activeMode='scripture';
var lastRef='John 3:16', songSel=null, mediaSel=null;
var liveState={live:false,liveType:'',songTitle:'',songSlide:null};
var toastTimer=null;

function $(id){return document.getElementById(id);}
function show(el,on){
  if(!el)return;
  var c=el.className.replace(/\bhidden\b/g,'').replace(/\s+/g,' ').trim();
  el.className=on?c:(c?c+' hidden':'hidden');
}
function toast(msg,isErr){
  var el=$('toast');if(!el)return;
  el.textContent=msg;el.className='show'+(isErr?' err':'');
  clearTimeout(toastTimer);toastTimer=setTimeout(function(){el.className=el.className.replace('show','').trim();},2500);
}
function setStatus(what,type,icon){
  var wi=$('statusWhat'),ti=$('statusType'),si=$('statusIcon');
  if(wi)wi.textContent=what||'—';if(ti)ti.textContent=type||'';if(si)si.textContent=icon||'📡';
  var b=$('onairBadge');if(b)b.className='onair'+(liveState.live?' show':'');
}
function capsFor(r){
  var m={admin:['status','go-live','clear','next','prev','scripture','songs','media','queue-add','set-translation','library'],
    scripture:['status','go-live','clear','next','prev','scripture'],
    songs:['status','go-live','clear','next','prev','songs','library'],
    media:['status','go-live','clear','media','library'],monitor:['status']};
  return m[r]||m.admin;
}
function can(cap){return caps.indexOf(cap)>=0;}
function canMode(m){
  if(m==='scripture')return can('scripture')||role==='admin';
  if(m==='songs')return can('songs')||role==='admin';
  if(m==='media')return can('media')||role==='admin';
  if(m==='presentation')return role==='admin';
  return false;
}

function hdrs(){
  var h={'Content-Type':'application/json'};
  if(pin)h['X-Remote-Pin']=pin;
  if(token)h['X-Remote-Token']=token;
  h['X-Remote-Role']=role||'admin';
  return h;
}
function xhr(method,url,body,cb){
  var x=new XMLHttpRequest();x.open(method,url,true);
  var hh=hdrs();Object.keys(hh).forEach(function(k){x.setRequestHeader(k,hh[k]);});
  x.onreadystatechange=function(){if(x.readyState!==4)return;var d={};try{d=JSON.parse(x.responseText||'{}');}catch(e){}cb(x.status,d);};
  x.onerror=function(){cb(0,{});};
  x.send(body?JSON.stringify(body):null);
}
function cmd(action,data,okMsg){
  xhr('POST','/api/control',{action:action,data:data||null},function(status,resp){
    if(status>=200&&status<300){if(okMsg)toast(okMsg);setTimeout(poll,300);}
    else if(status===401)toast('Not authorised',true);
    else if(status===403)toast('Role not permitted',true);
    else if(status===0)toast('Connection error',true);
    else toast('Error '+status,true);
  });
}

/* ── Navigation ── BUG-3 FIX: PREV/NEXT only controls the SAME content type
   that is currently live. If songs are projected and remote is in scripture mode,
   pressing PREV/NEXT does nothing (with a clear warning toast). Each mode button
   sends its own mode tag so the server routes correctly. */
function doNav(dir){
  var dirStr=dir>0?'next':'prev';
  // Guard: if something is live and it doesn't match our mode, block navigation
  if(liveState.live&&liveState.liveType&&liveState.liveType!=='clear'){
    var liveMode=liveState.liveType; // 'scripture','song','media','presentation'
    // Normalise song->songs
    if(liveMode==='song')liveMode='songs';
    if(liveMode!==activeMode){
      toast('\u26A0 '+liveMode.charAt(0).toUpperCase()+liveMode.slice(1)+' is live \u2014 switch mode first',true);
      return;
    }
  }
  cmd(dirStr,{mode:activeMode},'\u2713 '+dirStr.charAt(0).toUpperCase()+dirStr.slice(1));
}

function goLive(){
  cmd('present-current',{mode:activeMode},'\u2713 Sent to projection');
}

function signOut(){
  // Revoke server-side session first, then clear local state
  xhr('POST','/api/signout',{},function(){
    token='';pin='';role='admin';caps=[];
    try{localStorage.removeItem('acToken');}catch(e){}
    try{localStorage.removeItem('acPin');}catch(e){}
    try{localStorage.removeItem('acRole');}catch(e){}
    var sb=$('signOutBtn');if(sb)sb.className='signout-btn';
    var rb=$('roleBadge');if(rb){rb.className='role-badge';rb.textContent='';}
    var em=$('authMsg');if(em)em.textContent='';
    var pi=$('pinInput');if(pi)pi.value='';
    var ab=$('authBox');if(ab)ab.style.display='';
    setStatus('Signed out','Enter PIN to continue','\uD83D\uDD10');
  });
}
window.signOut=signOut;

function sendRef(){
  var v=($('refInput')||{}).value||'';v=v.trim();
  if(!v){toast('Enter a reference',true);return;}
  lastRef=v;
  cmd('present',{ref:v,rawRef:v},'\u2713 '+v);
}

function setMode(m){
  if(!canMode(m))return;
  activeMode=m;
  // Clear song search when leaving songs mode
  if(m!=='songs'){var ss=$('songSearch');if(ss)ss.value='';allSongs.length&&renderSongsFiltered(allSongs);}
  refreshUI();
  if(m==='songs')loadLib('songs');
  if(m==='media')loadLib('media');
  var sec=$(m==='scripture'?'scriptureSec':m==='songs'?'songsSec':m==='media'?'mediaSec':'presentationSec');
  if(sec)setTimeout(function(){sec.scrollIntoView({behavior:'smooth',block:'nearest'});},80);
}

function refreshUI(){
  ['scripture','songs','media','presentation'].forEach(function(m){
    var t=$('mode'+m.charAt(0).toUpperCase()+m.slice(1));
    if(t){t.className='tab'+(activeMode===m?' active':'');t.disabled=!canMode(m);}
  });
  // Use direct style.display — avoids hidden-class race with poll() calling refreshUI every 2.5s
  var secMap={scripture:'scriptureSec',songs:'songsSec',media:'mediaSec',presentation:'presentationSec'};
  ['scripture','songs','media','presentation'].forEach(function(m){
    var el=$(secMap[m]);
    if(!el)return;
    // Remove hidden class (in case it was set by initial HTML)
    el.className=el.className.replace(/\bhidden\b/g,'').trim();
    el.style.display=(activeMode===m)?'block':'none';
  });
  // Live button label
  var lb=$('liveBtn'),lt=$('liveBtnText');
  if(!lb)return;
  lb.className='go-btn'+(liveState.live?' on':'');
  var lbl=activeMode==='songs'?'PRESENT SONG':activeMode==='media'?'PRESENT MEDIA':activeMode==='presentation'?'PRESENT SLIDE':'GO LIVE';
  if(liveState.live&&activeMode==='scripture')lbl='RE-SEND LIVE';
  if(lt)lt.textContent=lbl;
}

function poll(){
  xhr('GET','/api/status',null,function(status,data){
    if(status!==200){
      if(status===401){
        setStatus('\uD83D\uDD12 Locked','Enter PIN','\uD83D\uDD10');
        // Clear stale token so subsequent polls don't keep triggering auth failures
        if(token){token='';try{localStorage.removeItem('acToken');}catch(e){}}
        var _ab=$('authBox');if(_ab)_ab.style.display='';
      }
      else setStatus('No connection','Check WiFi','\u26A0');
      return;
    }
    liveState.live=!!data.live;
    liveState.liveType=data.liveType||'';
    liveState.songTitle=data.songTitle||'';
    liveState.songSlide=data.songSlide!=null?data.songSlide:null;
    var what='',type='',icon='📡';
    if(data.verse&&data.verse.ref){what=data.verse.ref;type='Scripture';icon='\u271D';}
    else if(data.songTitle){what=data.songTitle;type='Song'+(data.songSlide!=null?' \u00B7 Slide '+(data.songSlide+1):'');icon='\uD83C\uDFB5';}
    else if(data.liveType==='media'){what='Media playing';type='Media';icon='\uD83C\uDFAC';}
    else if(data.liveType==='presentation'){what='Presentation';type='Slides';icon='\uD83D\uDCD1';}
    else{what='Nothing on screen';type='Clear';icon='\u25A1';}
    setStatus(what,type+(data.live?' \u2022 ON AIR':''),icon);
    refreshUI();
    // Scripture live indicator
    var snl=$('scriptureNowLive'),snlRef=$('scriptureNowLiveRef'),snlType=$('scriptureNowLiveType');
    if(snl&&snlRef){
      if(data.verse&&data.verse.ref){
        snl.className='scripture-live show';snlRef.textContent=data.verse.ref;
        snlType.textContent=data.live?'LIVE':'Preview';
      }else{snl.className='scripture-live';}
    }
    // Song active bar
    var bar=$('activeSongBar'),lbl=$('activeSongLabel'),sl=$('activeSongSlide');
    if(bar){
      if(liveState.liveType==='song'&&liveState.songTitle){
        bar.className='active-bar show';
        if(lbl)lbl.textContent=liveState.songTitle;
        if(sl)sl.textContent=liveState.songSlide!=null?'Slide '+(liveState.songSlide+1):'';
      }else{bar.className='active-bar';}
    }
    updateProjectionPreview(data.preview);
  });
}

var _lastPreviewKey='';
function updateProjectionPreview(pv){
  var dot=$('projDot'),empty=$('projEmpty'),content=$('projContent');
  if(!dot||!content)return;
  var isLive=liveState.live;
  dot.className='proj-dot'+(isLive?' on':'');
  if(!pv||pv.type==='clear'){
    _lastPreviewKey='clear';
    empty.style.display='';content.style.display='none';content.innerHTML='';
    return;
  }
  var key=JSON.stringify(pv);
  if(key===_lastPreviewKey)return;
  _lastPreviewKey=key;
  empty.style.display='none';content.style.display='';
  var html='';
  if(pv.type==='scripture'){
    html='<div class="proj-ref">'+esc(pv.ref||'')+'</div>'
        +'<div class="proj-text">'+esc(pv.text||'')+'</div>';
  }else if(pv.type==='song'){
    html='<div class="proj-song-title">'+esc(pv.title||'')+'</div>';
    if(pv.label)html+='<div class="proj-song-label">'+esc(pv.label)+'</div>';
    var lines=pv.lines||[];
    html+='<div class="proj-song-lines">'+lines.map(function(l){return esc(l);}).join('<br>')+'</div>';
  }else if(pv.type==='media'){
    var ic=pv.mediaType==='video'?'\uD83C\uDFAC':pv.mediaType==='audio'?'\uD83C\uDFA7':'\uD83D\uDDBC';
    html='<div class="proj-media-icon">'+ic+'</div><div class="proj-media-name">'+esc(pv.name||'Media')+'</div>';
  }else if(pv.type==='presentation'){
    var sn=pv.slide!=null?String(Math.trunc(+pv.slide)||''):'';
    html='<div class="proj-media-icon">\uD83D\uDCD1</div><div class="proj-media-name">Slide '+esc(sn)+'</div>';
  }else{
    html='<div class="proj-media-name">'+esc(pv.type||'Content')+'</div>';
  }
  content.innerHTML=html;
}
function esc(s){var d=document.createElement('div');d.textContent=s;return d.innerHTML;}

function loadLib(kind){
  xhr('GET','/api/library?type='+kind,null,function(status,data){
    if(status!==200)return;
    if(kind==='songs')renderSongs(data.items||[]);
    if(kind==='media')renderMedia(data.items||[]);
  });
}

var allSongs=[];

function clearSongSearch(){
  var inp=$('songSearch');
  if(inp) inp.value='';
  if(allSongs.length){ renderSongsFiltered(allSongs); }
  else { loadLib('songs'); }
}

function filterSongs(q){
  q=(q||'').toLowerCase().trim();
  if(!q){ renderSongsFiltered(allSongs.length?allSongs:[]); return; }
  var filtered=allSongs.filter(function(s){
    return (s.title||'').toLowerCase().indexOf(q)>=0||(s.author||'').toLowerCase().indexOf(q)>=0;
  });
  renderSongsFiltered(filtered);
}

function renderSongs(items){
  allSongs=items||[];
  renderSongsFiltered(allSongs);
}

function renderSongsFiltered(items){
  var el=$('songsList');if(!el)return;
  if(!items.length){el.innerHTML='<div class="song-card"><div class="song-title" style="color:var(--text-dim)">'+(allSongs.length?'No songs match':'No songs in library')+'</div></div>';return;}
  el.innerHTML='';
  items.forEach(function(s){
    var card=document.createElement('div');card.className='song-card';
    var title=document.createElement('div');title.className='song-title';title.textContent=s.title||'Untitled';card.appendChild(title);
    if(s.author){var meta=document.createElement('div');meta.className='song-meta';meta.textContent=s.author;card.appendChild(meta);}
    var chips=document.createElement('div');chips.className='chips';
    (s.sections||[]).slice(0,24).forEach(function(sec,idx){
      (function(sid,si,sec){
        var label=(sec.label&&sec.label.trim())?sec.label.trim():('Slide '+(si+1));
        var c=document.createElement('button');c.className='chip';c.textContent=label;c.type='button';
        c.onclick=function(){
          songSel={songId:sid,slideIdx:si};
          // Clear all chip selections across ALL song cards, then mark this one
          document.querySelectorAll('#songsList .chip').forEach(function(x){x.classList.remove('active');});
          c.classList.add('active');
          // Send to server to set state (currentSongId/SlideIdx) without projecting
          cmd('present-song',{songId:sid,slideIdx:si},'\u2713 '+label+' \u2014 Sent to projection');
        };
        chips.appendChild(c);
      })(s.id,idx,sec);
    });
    if((s.sections||[]).length>24){var more=document.createElement('span');more.className='chip';more.style.opacity='.4';more.textContent='+'+(s.sections.length-24)+' more';chips.appendChild(more);}
    card.appendChild(chips);el.appendChild(card);
  });
}

function renderMedia(items){
  var el=$('mediaList');if(!el)return;
  if(!items.length){el.innerHTML='<div class="song-card"><div class="song-title" style="color:var(--text-dim)">No media files</div></div>';return;}
  el.innerHTML='';
  var icons={video:'\uD83C\uDFAC',image:'\uD83D\uDDBC',audio:'\uD83C\uDFA7'};
  items.forEach(function(item){
    var card=document.createElement('div');card.className='media-card';
    var thumb=document.createElement('div');thumb.className='media-thumb';thumb.textContent=icons[item.type]||'\uD83C\uDFAC';
    var info=document.createElement('div');info.className='media-info';
    var name=document.createElement('div');name.className='media-name';name.textContent=item.title||'Untitled';
    var mtype=document.createElement('div');mtype.className='media-type';mtype.textContent=(item.type||'').toUpperCase();
    info.appendChild(name);info.appendChild(mtype);
    var play=document.createElement('div');play.className='media-play';play.textContent='\u25BA';
    card.appendChild(thumb);card.appendChild(info);card.appendChild(play);
    card.onclick=function(){
      mediaSel={mediaId:item.id};
      cmd('present-media',{mediaId:item.id},'\u2713 '+item.title);
    };
    el.appendChild(card);
  });
}

function bootstrap(){
  xhr('GET','/api/bootstrap',null,function(status,data){
    if(status===401||status===0){
      if(AUTH_REQUIRED){var _ab=$('authBox');if(_ab)_ab.style.display='';}
      setStatus(status===0?'Cannot reach AnchorCast':'Locked','Check WiFi or enter PIN','\u26A0');
      // Clear stale token — server may have restarted and token is no longer valid.
      // If we keep sending it, every background poll records an auth failure → 429.
      if(status===401){
        token='';
        try{localStorage.removeItem('acToken');}catch(e){}
        var sb=$('signOutBtn');if(sb)sb.className='signout-btn';
      }
      return;
    }
    // If server says auth is required but returned no role, force re-auth
    if(data.authRequired && !data.role){
      token='';pin='';role='';caps=[];
      try{localStorage.removeItem('acToken');localStorage.removeItem('acPin');localStorage.removeItem('acRole');}catch(e){}
      var _ab=$('authBox');if(_ab)_ab.style.display='';
      var em=$('authMsg');if(em)em.textContent='';
      return;
    }
    role=String(data.role||'admin').toLowerCase();
    caps=data.capabilities||capsFor(role);
    // Hide auth box on any successful bootstrap — token is valid regardless of authRequired flag
    {var _ab=$('authBox');if(_ab)_ab.style.display='none';}
    var rb=$('roleBadge');if(rb){rb.className='role-badge show';rb.textContent='\uD83D\uDC64 '+role.charAt(0).toUpperCase()+role.slice(1);}
    // Show sign-out button now that user is authenticated
    if(AUTH_REQUIRED){var sb=$('signOutBtn');if(sb)sb.className='signout-btn show';}
    refreshUI();loadLib('songs');loadLib('media');poll();
  });
}

window.addEventListener('load',function(){
  try{token=localStorage.getItem('acToken')||'';}catch(e){}
  try{pin=localStorage.getItem('acPin')||'';}catch(e){}
  // Always start with auth box visible — bootstrap() will hide it if auth passes
  // This prevents stale tokens from bypassing the PIN screen on fresh open
  try{var r=localStorage.getItem('acRole');if(r&&(token||pin)){role=r;caps=capsFor(role);}}catch(e){}

  $('modeScripture').onclick=function(){setMode('scripture');};
  $('modeSongs').onclick=function(){setMode('songs');};
  $('modeMedia').onclick=function(){setMode('media');};
  $('modePresentation').onclick=function(){setMode('presentation');};

  $('liveBtn').onclick=goLive;
  $('prevBtn').onclick=function(){doNav(-1);};
  $('nextBtn').onclick=function(){doNav(1);};
  $('clearBtn').onclick=function(){cmd('clear',null,'\u2713 Display cleared');};
  $('projClearBtn').onclick=function(){cmd('clear',null,'\u2713 Display cleared');};

  $('sendBtn').onclick=sendRef;
  $('refInput').addEventListener('keydown',function(e){if(e.key==='Enter'){e.preventDefault();sendRef();}});

  $('smpJohn316').onclick=function(){lastRef='John 3:16';cmd('present',{ref:'John 3:16',rawRef:'John 3:16'},'\u2713 John 3:16');};
  $('smpPsalm23').onclick=function(){lastRef='Psalms 23:1';cmd('present',{ref:'Psalms 23:1',rawRef:'Psalms 23:1'},'\u2713 Psalm 23:1');};
  $('smpPhil413').onclick=function(){lastRef='Philippians 4:13';cmd('present',{ref:'Philippians 4:13',rawRef:'Philippians 4:13'},'\u2713 Phil 4:13');};
  $('smpRom828').onclick=function(){lastRef='Romans 8:28';cmd('present',{ref:'Romans 8:28',rawRef:'Romans 8:28'},'\u2713 Romans 8:28');};
  $('smpIsa4031').onclick=function(){lastRef='Isaiah 40:31';cmd('present',{ref:'Isaiah 40:31',rawRef:'Isaiah 40:31'},'\u2713 Isaiah 40:31');};
  $('smpJer2911').onclick=function(){lastRef='Jeremiah 29:11';cmd('present',{ref:'Jeremiah 29:11',rawRef:'Jeremiah 29:11'},'\u2713 Jeremiah 29:11');};

  $('unlockBtn').onclick=function(){
    var p=($('pinInput')||{}).value||'';
    xhr('POST','/api/auth',{pin:p},function(status,data){
      if(status===200){
        token=data.token||'';pin=p;role=String(data.role||'admin').toLowerCase();caps=data.capabilities||capsFor(role);
        try{localStorage.setItem('acToken',token);localStorage.setItem('acPin',pin);localStorage.setItem('acRole',role);}catch(e){}
        var sb=$('signOutBtn');if(sb)sb.className='signout-btn show';
        var _ab=$('authBox');if(_ab)_ab.style.display='none';bootstrap();
      }else{var m=$('authMsg');if(m)m.textContent='Wrong PIN — try again';}
    });
  };

  // Song search wiring — must be done here (inside IIFE) since filterSongs is not global
  var songSearchEl=$('songSearch');
  var songSearchClearEl=$('songSearchClear');
  if(songSearchEl){
    songSearchEl.addEventListener('input',function(){ filterSongs(this.value); });
    songSearchEl.addEventListener('keydown',function(e){ if(e.key==='Escape'){ this.value=''; clearSongSearch(); } });
  }
  if(songSearchClearEl){
    songSearchClearEl.addEventListener('click',function(){ clearSongSearch(); });
  }

  $('presPrevBtn').onclick=function(){cmd('presentation-prev',null,'\u2713 Prev slide');};
  $('presNextBtn').onclick=function(){cmd('presentation-next',null,'\u2713 Next slide');};
  $('presCurrentBtn').onclick=function(){cmd('present-current',{mode:'presentation'},'\u2713 Slide presented');};

  refreshUI();bootstrap();
  setInterval(poll,2500);
  setInterval(bootstrap,15000);
});
})();
</script>
</body></html>`;
}

// ── NDI info// ── NDI info ──────────────────────────────────────────────────────────────────
function showNdiInfo(){
  const ip=localIp();
  dialog.showMessageBox(mainWindow,{
    type:'info',title:'External Output — AnchorCast',
    message:'External Output Setup',
    detail:
      'AnchorCast supports NDI output via the NDI SDK.\n\n'+
      'Setup steps:\n'+
      '  1. Install the NDI Runtime from ndi.video/download\n'+
      '  2. Build the NDI addon (see README for Windows/macOS instructions)\n'+
      '  3. If the SDK or build tools are missing, the app falls back to MJPEG browser output\n\n'+
      'When NDI is ready, AnchorCast appears as an NDI source on your network.\n'+
      'In OBS: Add Source → NDI Source → AnchorCast\n\n'+
      `Your local IP: ${ip}`,
    buttons:['Open ndi.video','Close'],
  }).then(r=>{if(r.response===0)shell.openExternal('https://ndi.video/download/');});
}

function showRemoteUrl(){
  const ip=localIp();
  const port=currentSettings.httpPort||8080;
  dialog.showMessageBox(mainWindow,{
    type:'info',title:'Remote Control URL',
    message:'Mobile Remote Control',
    detail:
      `Open this on any phone/tablet on the same WiFi:\n\n`+
      `  http://${ip}:${port}/remote\n\n`+
      `REST API (for integrations):\n`+
      `  POST http://${ip}:${port}/api/control\n`+
      `  Body: { "action": "go-live"|"next"|"prev"|"clear"|"present" }`,
    buttons:['OK'],
  });
}

function showShortcuts(){
  dialog.showMessageBox(mainWindow,{
    type:'info',title:'Keyboard Shortcuts',message:'Keyboard Shortcuts',
    detail:
      'Ctrl+L         — Toggle Go Live\n'+
      'Enter          — Send preview → live\n'+
      '→ / ←          — Next / Previous verse in queue\n'+
      'Tab            — Switch Bible / Context search\n'+
      'Esc            — Close overlay\n'+
      'Ctrl+P         — Open projection window\n'+
      'Ctrl+T         — Theme Designer\n'+
      'Ctrl+H         — Sermon History\n'+
      'Ctrl+,         — Settings\n'+
      'Ctrl+Backspace — Clear live display',
    buttons:['OK'],
  });
}



function createCountdownWindow() {
  if (countdownWindow && !countdownWindow.isDestroyed()) {
    if (countdownWindow.isMinimized()) countdownWindow.restore();
    countdownWindow.focus();
    return;
  }
  const win = new BrowserWindow({
    icon: APP_ICON,
    width: 1060,
    height: 820,
    title: 'AnchorCast — Timer',
    backgroundColor: '#0a0a0f',
    autoHideMenuBar: true,
    resizable: true,
    minWidth: 860,
    minHeight: 680,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
  });
  countdownWindow = win;
  win.loadFile(path.join(__dirname, 'renderer', 'countdown-window.html'));
  win.once('ready-to-show', () => { try { win.show(); } catch(e) {} });
  win.on('closed', () => { countdownWindow = null; });
}

let registrationStatusWindow = null;

function createRegistrationStatusWindow() {
  if (registrationStatusWindow && !registrationStatusWindow.isDestroyed()) {
    registrationStatusWindow.focus();
    return;
  }
  const reg = getRegistrationStatus();
  const win = new BrowserWindow({
    icon: APP_ICON,
    width: 640, height: reg.registered ? 720 : 820,
    minWidth: 620,
    minHeight: reg.registered ? 680 : 780,
    title: 'AnchorCast — Registration',
    backgroundColor: '#0a0a1a',
    resizable: true, show: false, center: true,
    webPreferences: { nodeIntegration:false, contextIsolation:true, preload:path.join(__dirname,'preload.js') },
  });
  registrationStatusWindow = win;
  win.setMenu(null);
  // Pass reg data via query param when the local renderer server is available.
  const regQuery = encodeURIComponent(JSON.stringify(reg));
  if (rendererPort) {
    win.loadURL(`http://127.0.0.1:${rendererPort}/registration-status.html?data=${regQuery}`);
  } else {
    win.loadFile(path.join(__dirname, 'renderer', 'registration-status.html'), {
      query: { data: JSON.stringify(reg) }
    });
  }
  win.once('ready-to-show', () => win.show());
  win.on('closed', () => { registrationStatusWindow = null; });
}

// Keep old name for any remaining references
function createLicenseWindow() { createRegistrationStatusWindow(); }

function buildRegistrationStatusHtml(reg) {
  if (reg.registered) {
    const date = reg.registeredAt
      ? new Date(reg.registeredAt).toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' })
      : 'Unknown';
    return `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Registration Status</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
    background:#0a0a1a;color:#e0e0e0;display:flex;flex-direction:column;
    height:100vh;overflow:hidden}
  .header{background:linear-gradient(135deg,#0d0d20,#1a1a35);
    padding:18px 28px;text-align:center;border-bottom:1px solid rgba(201,168,76,.2);
    flex-shrink:0}
  .logo{color:#c9a84c;font-size:16px;font-weight:700;letter-spacing:1px}
  .body{flex:1;padding:16px 22px;display:flex;flex-direction:column;
    gap:10px;overflow:hidden;min-height:0}
  .reg-banner{background:rgba(46,204,113,.08);border:1px solid rgba(46,204,113,.25);
    border-radius:10px;padding:18px 16px;text-align:center;flex-shrink:0}
  .reg-check{width:48px;height:48px;background:#2ecc71;border-radius:50%;
    display:flex;align-items:center;justify-content:center;margin:0 auto 10px;
    font-size:24px;box-shadow:0 0 20px rgba(46,204,113,.35)}
  .reg-title{font-size:13px;font-weight:600;color:#2ecc71;margin-bottom:6px}
  .reg-name{font-size:17px;font-weight:700;color:#fff;margin-bottom:2px}
  .reg-email{font-size:12px;color:#2ecc71;font-weight:600}
  .reg-church{font-size:11px;color:#888;margin-top:3px}
  .info-card{background:#0d0d20;border:1px solid rgba(255,255,255,.08);
    border-radius:8px;overflow:hidden;flex-shrink:0}
  .info-row{display:flex;justify-content:space-between;align-items:center;
    padding:10px 14px;border-bottom:1px solid rgba(255,255,255,.05);font-size:12px}
  .info-row:last-child{border-bottom:none}
  .info-lbl{color:#555;font-size:11px}
  .info-val{color:#bbb;font-family:monospace;font-size:11px;text-align:right;word-break:break-all}
  .support-card{background:#13132a;border:1px solid rgba(255,255,255,.08);
    border-radius:10px;padding:16px 18px;flex-shrink:0}
  .support-title{font-size:13px;font-weight:700;color:#e0e0e0;
    text-align:center;margin-bottom:6px}
  .support-sub{font-size:11px;color:#555;text-align:center;line-height:1.6;margin-bottom:14px}
  .btn-row{display:flex;gap:10px}
  .btn-paypal{flex:1;padding:11px 10px;border:none;border-radius:8px;cursor:pointer;
    font-size:12px;font-weight:700;font-family:inherit;
    background:#1565c0;color:#fff;display:flex;align-items:center;
    justify-content:center;gap:7px;transition:opacity .15s}
  .btn-paypal:hover{opacity:.85}
  .btn-github{flex:1;padding:11px 10px;border:1px solid rgba(201,168,76,.4);
    border-radius:8px;cursor:pointer;font-size:12px;font-weight:700;
    font-family:inherit;background:rgba(201,168,76,.1);color:#c9a84c;
    display:flex;align-items:center;justify-content:center;gap:7px;transition:opacity .15s}
  .btn-github:hover{opacity:.85}
  .oss{text-align:center;font-size:10px;color:#3a3a5a;padding-top:2px}
  .oss a{color:#4a4a7a;cursor:pointer;text-decoration:none}
  .oss a:hover{color:#c9a84c}
</style></head><body>
  <div class="header">
    <div class="logo">⚓ ANCHORCAST</div>
    <div style="color:#444;font-size:10px;letter-spacing:1.5px;margin-top:2px">LIVE SERMON DISPLAY</div>
  </div>
  <div class="body">

    <!-- Registration banner -->
    <div class="reg-banner">
      <div class="reg-check">✓</div>
      <div class="reg-title">This product is registered to</div>
      <div class="reg-name">${reg.fullName}</div>
      <div class="reg-email">${reg.email}</div>
      ${reg.churchName ? '<div class="reg-church">' + reg.churchName + '</div>' : ''}
    </div>

    <!-- Info rows -->
    <div class="info-card">
      <div class="info-row">
        <span class="info-lbl">Registered on</span>
        <span class="info-val">${date}</span>
      </div>
      <div class="info-row">
        <span class="info-lbl">Hardware ID</span>
        <span class="info-val">${reg.hwId}</span>
      </div>
      <div class="info-row">
        <span class="info-lbl">Status</span>
        <span class="info-val" style="color:#2ecc71">● Active — All features unlocked</span>
      </div>
    </div>

    <!-- Support section -->
    <div class="support-card">
      <div class="support-title">❤ Support this Project</div>
      <div class="support-sub">
        AnchorCast is built with love for churches everywhere.<br>
        Donations help keep development going — thank you!
      </div>
      <div class="btn-row">
        <button class="btn-paypal"
          onclick="window.electronAPI?.openExternal('https://paypal.me/appdeveloper')">
          <span style="font-size:16px">💳</span> Donate via PayPal
        </button>
        <button class="btn-github"
          onclick="window.electronAPI?.openExternal('https://github.com/anchorcastapp-team/anchorcastapp')">
          <span style="font-size:15px">⭐</span> Star on GitHub
        </button>
      </div>
    </div>

    <div class="oss">
      Free &amp; Open Source —
      <a onclick="window.electronAPI?.openExternal('https://github.com/anchorcastapp-team/anchorcastapp')">
        github.com/anchorcastapp-team/anchorcastapp
      </a>
    </div>

  </div>
</body></html>`;
  }

  // Not registered — show registration form inline
  return buildRegistrationHtml();
}

let aboutWindow = null;
function showAbout(){
  if (aboutWindow && !aboutWindow.isDestroyed()) { aboutWindow.focus(); return; }
  const win = new BrowserWindow({
    icon: APP_ICON,
    width: 560, height: 980,
    title: 'About AnchorCast',
    backgroundColor: '#0a0a1a',
    resizable: false, show: false, center: true,
    parent: mainWindow, modal: false,
    webPreferences: { nodeIntegration:false, contextIsolation:true, preload:path.join(__dirname,'preload.js') },
  });
  aboutWindow = win;
  win.setMenu(null);
  win.loadURL(`http://127.0.0.1:${rendererPort}/about.html`);
  win.once('ready-to-show', () => win.show());
  win.on('closed', () => { aboutWindow = null; });
}


// ── Settings ──────────────────────────────────────────────────────────────────

function ensureDetectionReviewDir(){
  try{
    fs.mkdirSync(DETECTION_REVIEW_DIR, { recursive:true });
    if(!fs.existsSync(DETECTION_PHRASES_FILE)) fs.writeFileSync(DETECTION_PHRASES_FILE, '[]');
    if(!fs.existsSync(DETECTION_EVENTS_FILE)) fs.writeFileSync(DETECTION_EVENTS_FILE, '[]');
  }catch(_){}
}
function readDetectionJson(file, fallback=[]){
  try{ ensureDetectionReviewDir(); if(fs.existsSync(file)) return JSON.parse(fs.readFileSync(file,'utf8')); }catch(_){}
  return fallback;
}
function writeDetectionJson(file, data){
  ensureDetectionReviewDir();
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
}

function loadSettings(){
  try{ if(fs.existsSync(SETTINGS_FILE)) return JSON.parse(fs.readFileSync(SETTINGS_FILE,'utf-8')); }
  catch(e){}
  return defaultSettings();
}
function saveSettings(s){
  s = { ...defaultSettings(), ...s };
  s.transcriptSource = ['deepgram','local','cloud'].includes(s.transcriptSource) ? s.transcriptSource : 'local';
  s.ndiSourceName = String(s.ndiSourceName || 'AnchorCast').trim() || 'AnchorCast';
  s.remoteRequireAuth = isRemoteAuthRequired(s.remoteRequireAuth);
  s.remotePin = normalizeRemotePinValue(s.remotePin);
  s.remoteAdminPin = normalizeRemotePinValue(s.remoteAdminPin || s.remotePin);
  s.remoteScripturePin = normalizeRemotePinValue(s.remoteScripturePin);
  s.remoteSongsPin = normalizeRemotePinValue(s.remoteSongsPin);
  s.remoteMediaPin = normalizeRemotePinValue(s.remoteMediaPin);
  s.remoteMonitorPin = normalizeRemotePinValue(s.remoteMonitorPin);
  if(!s.remoteAdminPin) s.remoteAdminPin = String(Math.floor(100000 + Math.random()*900000));
  if(!s.remoteScripturePin) s.remoteScripturePin = String(Math.floor(100000 + Math.random()*900000));
  if(!s.remoteSongsPin) s.remoteSongsPin = String(Math.floor(100000 + Math.random()*900000));
  if(!s.remoteMediaPin) s.remoteMediaPin = String(Math.floor(100000 + Math.random()*900000));
  if(!s.remoteMonitorPin) s.remoteMonitorPin = String(Math.floor(100000 + Math.random()*900000));
  s.remotePin = s.remoteAdminPin;
  const prevNdi     = currentSettings.ndiEnabled;
  const prevAdapter = currentSettings.networkAdapter;
  const prevModel   = currentSettings.whisperModel;
  fs.mkdirSync(path.dirname(SETTINGS_FILE),{recursive:true});
  fs.writeFileSync(SETTINGS_FILE,JSON.stringify(s,null,2));
  currentSettings=s;
  startHttpServer(s.httpPort||8080);
  // Restart NDI if adapter changed or ndiEnabled toggled
  if(s.ndiEnabled && (!prevNdi || prevAdapter !== s.networkAdapter)){
    stopNdi(); setTimeout(()=>startNdi(), 500);
  } else if(!s.ndiEnabled && prevNdi){
    stopNdi();
  }
  // FIX: Restart Whisper if model changed
  if(s.whisperModel && s.whisperModel !== prevModel){
    const _m = String(s.whisperModel);
    const resolved = _m==='small'?'small.en':_m==='base'?'base.en':_m==='tiny'?'tiny.en':_m;
    console.log(`[Whisper] Model changed: ${prevModel} → ${resolved}. Restarting...`);
    stopWhisperServer();
    // Small delay ensures OS releases the port before we restart
    setTimeout(() => {
      if (!whisperProc) startWhisperServer(resolved);
    }, 1200);
  }
}
function defaultSettings(){
  return{apiKey:'',translation:'KJV',displayMode:'manual',theme:'sanctuary',
    whisperSource:'local',
    onlineMode:false,
    fontSize:'medium',microphoneId:'default',projectionDisplay:null,
    confidenceThreshold:0.75,showVerseNumbers:true,
    whisperModel:'small.en',transcriptSource:'local',
    overlayTextOnMedia:false,
    overlayTextOnMediaSongs:true,
    overlayTextOnMediaScripture:true,
    overlayTextOnMediaLowerThird:false,
    overlayTextOnMediaDim:35,
    hideGetStarted:false,
    geniusApiKey:'',
    ndiSourceName:'AnchorCast',
    httpPort:8080,remoteEnabled:true,ndiEnabled:false,
    remoteRequireAuth:true,
    remotePin:'', remoteAdminPin:'', remoteScripturePin:'', remoteSongsPin:'', remoteMediaPin:'', remoteMonitorPin:'',
    showCopyright:false, ccliNumber:'', ccliLicenseType:'streaming', ccliDisplayFormat:'ccli_only'};
}

// ── Themes ────────────────────────────────────────────────────────────────────
const _builtinBox = (o) => ({
  text:'', bold:false, italic:false, lineSpacing:1.4, textTransform:'none', fontWeight:400, letterSpacing:0,
  shadow:true, shadowColor:'#000000', shadowBlur:8, shadowOffsetX:0, shadowOffsetY:2,
  bgFill:'', bgOpacity:0, borderW:0, borderColor:'#ffffff', borderRadius:0, ...o
});
const BUILTIN_THEMES=[
  {id:'sanctuary',name:'Sanctuary',builtIn:true,category:'scripture',bgType:'radial',
   bgColor1:'#10082a',bgColor2:'#000518',textColor:'#ede6d8',
   accentColor:'#c9a84c',refColor:'#c9a84c',transColor:'#6a5220',
   fontFamily:'Crimson Pro',fontSize:48,fontStyle:'normal',
   showVerseNum:true,textAlign:'center',padding:60,
   boxes:[
     _builtinBox({role:'main',x:60,y:60,w:1800,h:960,
       text:'Genesis 1:1-2 (KJV)\nIn the beginning God created the heaven and the earth.',
       fontFamily:'Crimson Pro',fontSize:56,color:'#ede6d8',align:'center',valign:'center',lineSpacing:1.7,
       refFontFamily:'Cinzel',refFontSize:52,refColor:'#c9a84c',refBold:true,refLineSpacing:1.4})
   ]},
  {id:'dawn',name:'Dawn',builtIn:true,category:'scripture',bgType:'radial',
   bgColor1:'#1a0c08',bgColor2:'#080010',textColor:'#f0e8d8',
   accentColor:'#e8904c',refColor:'#e8904c',transColor:'#8a5020',
   fontFamily:'Crimson Pro',fontSize:48,fontStyle:'normal',
   showVerseNum:true,textAlign:'center',padding:60,
   boxes:[
     _builtinBox({role:'main',x:60,y:60,w:1800,h:960,
       text:'Genesis 1:1-2 (KJV)\nIn the beginning God created the heaven and the earth.',
       fontFamily:'Crimson Pro',fontSize:56,color:'#f0e8d8',align:'center',valign:'center',lineSpacing:1.7,
       refFontFamily:'Cinzel',refFontSize:52,refColor:'#e8904c',refBold:true,refLineSpacing:1.4})
   ]},
  {id:'deep',name:'Deep Water',builtIn:true,category:'scripture',bgType:'radial',
   bgColor1:'#001028',bgColor2:'#000408',textColor:'#e0ecf8',
   accentColor:'#4c9ae8',refColor:'#4c9ae8',transColor:'#204870',
   fontFamily:'Crimson Pro',fontSize:48,fontStyle:'normal',
   showVerseNum:true,textAlign:'center',padding:60,
   boxes:[
     _builtinBox({role:'main',x:60,y:60,w:1800,h:960,
       text:'Genesis 1:1-2 (KJV)\nIn the beginning God created the heaven and the earth.',
       fontFamily:'Crimson Pro',fontSize:56,color:'#e0ecf8',align:'center',valign:'center',lineSpacing:1.7,
       refFontFamily:'Cinzel',refFontSize:52,refColor:'#4c9ae8',refBold:true,refLineSpacing:1.4})
   ]},
  {id:'minimal',name:'Minimal',builtIn:true,category:'scripture',bgType:'solid',
   bgColor1:'#000000',bgColor2:'#000000',textColor:'#ffffff',
   accentColor:'#aaaaaa',refColor:'#888888',transColor:'#555555',
   fontFamily:'DM Sans',fontSize:52,fontStyle:'normal',
   showVerseNum:false,textAlign:'center',padding:80,
   boxes:[
     _builtinBox({role:'main',x:60,y:60,w:1800,h:960,
       text:'Genesis 1:1-2 (KJV)\nIn the beginning God created the heaven and the earth.',
       fontFamily:'DM Sans',fontSize:52,color:'#ffffff',align:'center',valign:'center',lineSpacing:1.7,
       refFontFamily:'DM Sans',refFontSize:48,refColor:'#aaaaaa',refBold:true,refLineSpacing:1.4})
   ]},
];
function loadThemes(){
  try{
    if(fs.existsSync(THEMES_FILE)){
      const custom=JSON.parse(fs.readFileSync(THEMES_FILE,'utf-8'));
      // Custom themes saved from the new designer have a category field
      // Legacy themes (no category) are treated as scripture themes
      const withCategory = custom.map(t => ({...t, category: t.category || 'scripture'}));
      return[...BUILTIN_THEMES,...withCategory];
    }
  }catch(e){}
  return BUILTIN_THEMES;
}
function saveCustomThemes(themes){
  const custom=themes.filter(t=>!t.builtIn);
  fs.mkdirSync(path.dirname(THEMES_FILE),{recursive:true});
  fs.writeFileSync(THEMES_FILE,JSON.stringify(custom,null,2));
}

// ── IPC Handlers ──────────────────────────────────────────────────────────────
ipcMain.handle('open-projection',(_, id)=>{ createProjectionWindow(id); return{success:true}; });
ipcMain.handle('close-projection',()=>{ if(projectionWindow)projectionWindow.close(); return{success:true}; });
// ── NDI Engine ────────────────────────────────────────────────────────────────
// Priority 1: NDI SDK via ndi-addon (AnchorCast appears as NDI source in OBS/vMix)
// Priority 2: MJPEG HTTP stream (OBS Browser Source / vMix Web Browser)
let ndiTimer      = null;
let ndiStatus     = 'disabled';
let ndiFrameCount = 0;
let ndiMjpegServer= null;
let mjpegClients  = [];
let latestJpegFrame = null;
const NDI_MJPEG_PORT = 49876;
let ndiSdkActive  = false;

function tryLoadNdiAddon(){
  try{
    const addonPath = path.join(__dirname,'..','ndi-addon','index.js');
    const addon = require(addonPath);
    return addon.isAvailable() ? addon : null;
  } catch(e){ return null; }
}

function ndiAddonStatus(){
  const addonDir = path.join(__dirname,'..','ndi-addon');
  const built = fs.existsSync(path.join(addonDir,'build','Release','ndi_sender.node'));
  const projectFilesPresent =
    fs.existsSync(path.join(addonDir,'index.js')) ||
    fs.existsSync(path.join(addonDir,'binding.gyp')) ||
    fs.existsSync(path.join(addonDir,'build-ndi.bat'));
  if(!projectFilesPresent) return { state:'missing', label:'NDI addon files not found' };
  if(!built) return { state:'not-built', label:'NDI addon not compiled — run ndi-addon/build-ndi.bat' };
  const addon = tryLoadNdiAddon();
  if(addon) return { state:'ready', label:'NDI ready — AnchorCast appears as NDI source' };
  return { state:'load-error', label:'NDI addon compiled but failed to load — check NDI Runtime on this PC' };
}

function grandioseStatus(){
  const binPath = path.join(__dirname,'..','node_modules','grandiose','build','Release','grandiose.node');
  return fs.existsSync(binPath)
    ? { state:'compiled', label:'grandiose compiled' }
    : { state:'missing',  label:'grandiose not compiled' };
}

function startMjpegServer(){
  if(ndiMjpegServer) return;
  ndiMjpegServer = http.createServer((req, res) => {
    const url = req.url.split('?')[0];
    if(url === '/stream'){
      res.writeHead(200, {
        'Content-Type': 'multipart/x-mixed-replace; boundary=--SCframe',
        'Cache-Control': 'no-cache, no-store',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*',
      });
      mjpegClients.push(res);
      if(latestJpegFrame) pushFrame(res, latestJpegFrame);
      req.on('close', () => { mjpegClients = mjpegClients.filter(c => c !== res); });
    } else if(url === '/snapshot'){
      if(latestJpegFrame){
        res.writeHead(200,{'Content-Type':'image/jpeg','Access-Control-Allow-Origin':'*'});
        res.end(latestJpegFrame);
      } else { res.writeHead(204); res.end(); }
    } else {
      const ip = localIp();
      res.writeHead(200,{'Content-Type':'text/html; charset=utf-8'});
      res.end(`<!DOCTYPE html><html><head><title>AnchorCast Stream</title>
<style>body{background:#000;color:#eee;font-family:sans-serif;padding:20px;margin:0}
a{color:#4af}img{max-width:100%;border:1px solid #333;margin-top:10px}</style></head>
<body><h2>&#10022; AnchorCast MJPEG Stream</h2>
<p><b>OBS Browser Source:</b> <code>http://localhost:${NDI_MJPEG_PORT}/stream</code></p>
<p><b>vMix Web Browser Input:</b> <code>http://${ip}:${NDI_MJPEG_PORT}/stream</code></p>
<img src="/stream" alt="Live stream"></body></html>`);
    }
  });
  ndiMjpegServer.listen(NDI_MJPEG_PORT, '0.0.0.0', () => {
    const ip = localIp();
    console.log(`[Output] MJPEG stream: http://localhost:${NDI_MJPEG_PORT}/stream`);
    console.log(`[Output] vMix: http://${ip}:${NDI_MJPEG_PORT}/stream`);
  });
  ndiMjpegServer.on('error', e => {
    if(e.code === 'EADDRINUSE') console.warn(`[Output] Port ${NDI_MJPEG_PORT} in use`);
    else console.warn('[Output] MJPEG error:', e.message);
  });
}

function stopMjpegServer(){
  mjpegClients.forEach(c => { try{ c.end(); }catch(e){} });
  mjpegClients = [];
  latestJpegFrame = null;
  if(ndiMjpegServer){ ndiMjpegServer.close(); ndiMjpegServer = null; }
}

function pushFrame(res, jpegBuf){
  try{
    res.write(`--SCframe\r\nContent-Type: image/jpeg\r\nContent-Length: ${jpegBuf.length}\r\n\r\n`);
    res.write(jpegBuf);
    res.write('\r\n');
  } catch(e){
    mjpegClients = mjpegClients.filter(c => c !== res);
  }
}

function startNdi(){
  if(ndiStatus === 'running' || ndiStatus === 'starting') return;
  ndiStatus = 'starting';
  ndiSdkActive = false;
  notifyNdiStatus();

  // Try real NDI SDK first
  const addon = tryLoadNdiAddon();
  if(addon){
    try{
      // Use actual projection window size if available, else default to 1920x1080
      let W = 1920, H = 1080;
      if(projectionWindow && !projectionWindow.isDestroyed()){
        const b = projectionWindow.getContentBounds();
        W = b.width  || 1920;
        H = b.height || 1080;
      }
      addon.createSender(currentSettings.ndiSourceName || 'AnchorCast', W, H, 30000, 1000);
      ndiSdkActive = true;
      ndiStatus = 'running';
      ndiFrameCount = 0;
      console.log(`[NDI] Official NDI SDK active — "${currentSettings.ndiSourceName || 'AnchorCast'}" at ${W}x${H}`);
      notifyNdiStatus();
      startNdiFrameLoop(addon);
      return;
    } catch(e){
      console.warn('[NDI] SDK addon failed:', e.message, '— falling back to MJPEG');
      ndiSdkActive = false;
    }
  }

  // Fallback: MJPEG stream
  startMjpegServer();
  ndiStatus = 'running';
  ndiFrameCount = 0;
  notifyNdiStatus();
  startNdiFrameLoop(null);
}

function startNdiFrameLoop(addon){
  if(ndiTimer) clearInterval(ndiTimer);
  let busy = false;
  let ndiSenderW = 0;
  let ndiSenderH = 0;

  ndiTimer = setInterval(async () => {
    if(busy) return;
    if(!projectionWindow || projectionWindow.isDestroyed()) return;
    busy = true;
    try{
      // Capture at native resolution (accounts for display scaling)
      const scaleFactor = projectionWindow._displayScaleFactor || 1;
      const cssW = projectionWindow._displayWidth  || 1920;
      const cssH = projectionWindow._displayHeight || 1080;

      // capturePage with explicit bounds in CSS pixels — Electron returns native pixels
      const img = await projectionWindow.webContents.capturePage({
        x: 0, y: 0, width: cssW, height: cssH
      });
      if(!img || img.isEmpty()){ busy=false; return; }

      const nativeW = img.getSize().width;
      const nativeH = img.getSize().height;
      if(!nativeW || !nativeH){ busy=false; return; }

      if(addon && ndiSdkActive){
        // For NDI, resize to standard 1920x1080 if needed (NDI receivers expect standard sizes)
        // If already 1920x1080 or native, send direct. Otherwise resize.
        let sendImg = img;
        let W = nativeW;
        let H = nativeH;

        // Resize to 1920x1080 if the captured size is different (e.g. 4K display scaled)
        if(nativeW !== 1920 || nativeH !== 1080){
          sendImg = img.resize({ width: 1920, height: 1080, quality: 'best' });
          W = 1920; H = 1080;
        }

        // Recreate sender if dimensions changed
        if(W !== ndiSenderW || H !== ndiSenderH){
          ndiSenderW = W; ndiSenderH = H;
          try{
            addon.destroySender();
            addon.createSender(currentSettings.ndiSourceName || 'AnchorCast', W, H, 30000, 1000);
            console.log(`[NDI] Sender created ${W}x${H}`);
          } catch(e){ console.warn('[NDI] Resize failed:', e.message); }
        }

        const bitmap = sendImg.getBitmap();
        addon.sendBGRA(Buffer.from(bitmap));
        ndiFrameCount++;
        if(ndiFrameCount % 150 === 1){
          console.log(`[NDI] SDK frame ${ndiFrameCount} (native:${nativeW}x${nativeH} → NDI:${W}x${H})`);
        }
      } else {
        // MJPEG — send at full native resolution, high quality
        const jpeg = img.toJPEG(95);
        latestJpegFrame = jpeg;
        mjpegClients.forEach(c => pushFrame(c, jpeg));
        ndiFrameCount++;
        if(ndiFrameCount % 150 === 1){
          console.log(`[Output] MJPEG frame ${ndiFrameCount} (${nativeW}x${nativeH}) — ${mjpegClients.length} client(s)`);
        }
      }
    } catch(e){
      // skip bad frames silently
    } finally { busy = false; }
  }, 33); // ~30fps
}

function stopNdi(){
  if(ndiTimer){ clearInterval(ndiTimer); ndiTimer = null; }
  if(ndiSdkActive){
    try{ tryLoadNdiAddon()?.destroySender(); }catch(e){}
    ndiSdkActive = false;
  }
  stopMjpegServer();
  ndiFrameCount = 0;
  ndiStatus = 'disabled';
  notifyNdiStatus();
  console.log('[Output] Stopped');
}

function notifyNdiStatus(){
  const addonSt = ndiAddonStatus();
  const ip = localIp();
  mainWindow?.webContents.send('ndi-status', {
    status:      ndiStatus,
    method:      ndiSdkActive ? 'ndi-sdk' : 'mjpeg',
    ndiSdkActive,
    obsUrl:      `http://localhost:${NDI_MJPEG_PORT}/stream`,
    vmixUrl:     `http://${ip}:${NDI_MJPEG_PORT}/stream`,
    clientCount: mjpegClients.length,
    addonState:  addonSt.state,
    addonLabel:  addonSt.label,
    sourceName:  currentSettings.ndiSourceName || 'AnchorCast',
  });
}

ipcMain.handle('ndi-start', async () => {
  const gate = requiresRegistration('External Output (NDI)');
  if (gate) return { error: gate.reason, blocked: true };
  startNdi();
  await new Promise(r => setTimeout(r, 400));
  currentSettings.ndiEnabled = true;
  try{ fs.mkdirSync(path.dirname(SETTINGS_FILE),{recursive:true}); fs.writeFileSync(SETTINGS_FILE,JSON.stringify(currentSettings,null,2)); }catch(e){}
  const addonSt = ndiAddonStatus();
  const ip = localIp();
  return {
    status: ndiStatus,
    method: ndiSdkActive ? 'ndi-sdk' : 'mjpeg',
    ndiSdkActive,
    obsUrl:  `http://localhost:${NDI_MJPEG_PORT}/stream`,
    vmixUrl: `http://${ip}:${NDI_MJPEG_PORT}/stream`,
    addonState: addonSt.state,
    addonLabel: addonSt.label,
    sourceName: currentSettings.ndiSourceName || 'AnchorCast',
  };
});
ipcMain.handle('ndi-stop', () => {
  stopNdi();
  currentSettings.ndiEnabled = false;
  try{ fs.mkdirSync(path.dirname(SETTINGS_FILE),{recursive:true}); fs.writeFileSync(SETTINGS_FILE,JSON.stringify(currentSettings,null,2)); }catch(e){}
  return { status: ndiStatus };
});
ipcMain.handle('ndi-status', () => {
  const addonSt = ndiAddonStatus();
  const ip = localIp();
  return {
    status:      ndiStatus,
    method:      ndiSdkActive ? 'ndi-sdk' : 'mjpeg',
    ndiSdkActive,
    obsUrl:      `http://localhost:${NDI_MJPEG_PORT}/stream`,
    vmixUrl:     `http://${ip}:${NDI_MJPEG_PORT}/stream`,
    clientCount: mjpegClients.length,
    addonState:  addonSt.state,
    addonLabel:  addonSt.label,
    sourceName:  currentSettings.ndiSourceName || 'AnchorCast',
  };
});

// Enhanced project-verse — also triggers NDI capture on verse change
ipcMain.handle('project-verse',(_,data)=>{
  currentLiveVerse=data;
  const overlayEnabled = !!currentSettings.overlayTextOnMedia && !!currentSettings.overlayTextOnMediaScripture && !!currentBackgroundMedia;
  const overlayOptions = overlayEnabled ? {
    enabled:true,
    lowerThird: !!currentSettings.overlayTextOnMediaLowerThird,
    dim: Number.isFinite(Number(currentSettings.overlayTextOnMediaDim)) ? Number(currentSettings.overlayTextOnMediaDim) : 35,
    module:'scripture'
  } : { enabled:false, lowerThird:false, dim:0, module:'scripture' };
  pushRenderStateToProjection(buildRenderState('scripture', data, { backgroundMedia: overlayEnabled ? currentBackgroundMedia : null, overlayOptions }));
  return{success:true};
});
// Project song lyrics (same pipeline as verses but with song data)
ipcMain.handle('project-song',(_,data)=>{
  currentLiveVerse=data; // reuse for NDI/status
  const overlayEnabled = !!currentSettings.overlayTextOnMedia && !!currentSettings.overlayTextOnMediaSongs && !!currentBackgroundMedia;
  const overlayOptions = overlayEnabled ? {
    enabled:true,
    lowerThird: !!currentSettings.overlayTextOnMediaLowerThird,
    dim: Number.isFinite(Number(currentSettings.overlayTextOnMediaDim)) ? Number(currentSettings.overlayTextOnMediaDim) : 35,
    module:'song'
  } : { enabled:false, lowerThird:false, dim:0, module:'song' };
  pushRenderStateToProjection(buildRenderState('song', data, { backgroundMedia: overlayEnabled ? currentBackgroundMedia : null, overlayOptions }));
  return{success:true};
});
ipcMain.handle('clear-projection',()=>{
  currentLiveVerse=null;
  currentBackgroundMedia=null;
  pushRenderStateToProjection(buildRenderState('clear', null, { backgroundMedia: null }));
  if(projectionWindow){ projectionWindow.webContents.send('clear-verse'); }
  return{success:true};
});
ipcMain.handle('get-current-render-state',()=>currentRenderState);
ipcMain.handle('get-displays',()=>
  screen.getAllDisplays().map(d=>({id:d.id,label:d.label||`Display ${d.id}`,
    bounds:d.bounds,isPrimary:d.id===screen.getPrimaryDisplay().id}))
);
ipcMain.handle('get-settings',()=>({ ...defaultSettings(), ...loadSettings() }));
ipcMain.handle('save-settings',(_,s,opts)=>{
  try{
    // If PIN or auth settings changed, invalidate all existing remote sessions
    const pinChanged = (
      String(s.remoteAdminPin||'') !== String(currentSettings.remoteAdminPin||'') ||
      String(s.remoteScripturePin||'') !== String(currentSettings.remoteScripturePin||'') ||
      String(s.remoteSongsPin||'') !== String(currentSettings.remoteSongsPin||'') ||
      String(s.remoteMediaPin||'') !== String(currentSettings.remoteMediaPin||'') ||
      String(s.remoteRequireAuth||'') !== String(currentSettings.remoteRequireAuth||'')
    );
    if (pinChanged) {
      remoteSessionTokens.clear();
      console.log('[Remote] PIN/auth settings changed — all sessions invalidated');
    }
    saveSettings(s);
    // opts.themeOnly = true means only a song/presentation/scripture theme ID changed.
    // The renderer uses this to skip re-projecting live scripture when a song theme is saved.
    const themeOnly = !!(opts && opts.themeOnly);
    const changedKeys = opts && Array.isArray(opts.changedKeys) ? opts.changedKeys : [];
    if(mainWindow) mainWindow.webContents.send('settings-saved', { themeOnly, changedKeys });
    return{success:true};
  }catch(e){return{success:false,error:e.message};}
});

// Transcripts
ipcMain.handle('save-transcript',(_,data)=>{
  try{
    let arr=[];
    if(fs.existsSync(TRANSCRIPTS_FILE)) arr=JSON.parse(fs.readFileSync(TRANSCRIPTS_FILE,'utf-8'));
    arr.unshift({...data,id:Date.now().toString(),savedAt:new Date().toISOString()});
    arr=arr.slice(0,100);
    fs.mkdirSync(path.dirname(TRANSCRIPTS_FILE),{recursive:true});
    fs.writeFileSync(TRANSCRIPTS_FILE,JSON.stringify(arr,null,2));
    return{success:true};
  }catch(e){return{success:false,error:e.message};}
});
ipcMain.handle('get-transcripts',()=>{
  try{ if(fs.existsSync(TRANSCRIPTS_FILE)) return JSON.parse(fs.readFileSync(TRANSCRIPTS_FILE,'utf-8')); }
  catch(e){}
  return[];
});
ipcMain.handle('delete-transcript',(_,id)=>{
  try{
    if(!fs.existsSync(TRANSCRIPTS_FILE)) return{success:false};
    let arr=JSON.parse(fs.readFileSync(TRANSCRIPTS_FILE,'utf-8'));
    arr=arr.filter(t=>t.id!==id);
    fs.writeFileSync(TRANSCRIPTS_FILE,JSON.stringify(arr,null,2));
    return{success:true};
  }catch(e){return{success:false,error:e.message};}
});

// Themes
ipcMain.handle('get-themes',()=>loadThemes());
ipcMain.handle('save-themes',(_,themes)=>{
  try{
    saveCustomThemes(themes);
    // Notify main window to reload themes immediately
    mainWindow?.webContents.send('themes-updated');
    return{success:true};
  }catch(e){return{success:false,error:e.message};}
});

// ── Adaptive Transcript Memory ──────────────────────────────────────────────
const ADAPTIVE_DIR       = path.join(DATA_DIR, 'adaptive');
const ADAPTIVE_RULES_FILE     = path.join(ADAPTIVE_DIR, 'correction_rules.json');
const ADAPTIVE_VOCAB_FILE     = path.join(ADAPTIVE_DIR, 'vocabulary.json');
const ADAPTIVE_PROFILES_FILE  = path.join(ADAPTIVE_DIR, 'speaker_profiles.json');
const ADAPTIVE_LEARNING_FILE  = path.join(ADAPTIVE_DIR, 'learning_queue.json');
const ADAPTIVE_EVENTS_FILE    = path.join(ADAPTIVE_DIR, 'correction_events.json');
const ADAPTIVE_SESSIONS_FILE  = path.join(ADAPTIVE_DIR, 'sessions.json');
const ADAPTIVE_CHUNKS_DIR     = path.join(ADAPTIVE_DIR, 'chunks');



function ensureClipExportDir(){
  try{ fs.mkdirSync(CLIP_EXPORT_DIR, { recursive:true }); }catch(_){}
}

function ensureServiceArchiveDir(){
  try{
    fs.mkdirSync(SERVICE_ARCHIVE_DIR, { recursive:true });
    if(!fs.existsSync(SERVICE_ARCHIVE_INDEX_FILE)) fs.writeFileSync(SERVICE_ARCHIVE_INDEX_FILE, '[]');
  }catch(_){}
}
function readServiceArchiveIndex(){
  try{ ensureServiceArchiveDir(); return JSON.parse(fs.readFileSync(SERVICE_ARCHIVE_INDEX_FILE, 'utf8')); }catch(_){}
  return [];
}
function writeServiceArchiveIndex(data){
  ensureServiceArchiveDir();
  fs.writeFileSync(SERVICE_ARCHIVE_INDEX_FILE, JSON.stringify(Array.isArray(data)?data:[], null, 2));
}

function ensureAdaptiveDir(){
  fs.mkdirSync(ADAPTIVE_DIR, { recursive:true });
  fs.mkdirSync(ADAPTIVE_CHUNKS_DIR, { recursive:true });
}
function readAdaptiveJson(file, fallback=[]){
  try{ if(fs.existsSync(file)) return JSON.parse(fs.readFileSync(file,'utf8')); }catch(_){}
  return fallback;
}
function writeAdaptiveJson(file, data){
  ensureAdaptiveDir();
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
}

// Load all adaptive data at once
ipcMain.handle('load-adaptive-memory', () => {
  ensureAdaptiveDir();
  return {
    rules:         readAdaptiveJson(ADAPTIVE_RULES_FILE, []),
    vocab:         readAdaptiveJson(ADAPTIVE_VOCAB_FILE, []),
    profiles:      readAdaptiveJson(ADAPTIVE_PROFILES_FILE, []),
    learningQueue: readAdaptiveJson(ADAPTIVE_LEARNING_FILE, []),
    settings:      { enabled: true, learningEnabled: true },
  };
});

// Save a specific section
ipcMain.handle('save-adaptive-memory', (_, { type, data }) => {
  try{
    if(type === 'rules')         writeAdaptiveJson(ADAPTIVE_RULES_FILE, data);
    else if(type === 'vocab')    writeAdaptiveJson(ADAPTIVE_VOCAB_FILE, data);
    else if(type === 'profiles') writeAdaptiveJson(ADAPTIVE_PROFILES_FILE, data);
    else if(type === 'learningQueue') writeAdaptiveJson(ADAPTIVE_LEARNING_FILE, data);
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});

// Persist transcript session header
ipcMain.handle('persist-transcript-session', (_, session) => {
  try{
    const sessions = readAdaptiveJson(ADAPTIVE_SESSIONS_FILE, []);
    const idx = sessions.findIndex(s => s.id === session.id);
    if(idx >= 0) sessions[idx] = { ...sessions[idx], ...session };
    else sessions.push(session);
    writeAdaptiveJson(ADAPTIVE_SESSIONS_FILE, sessions.slice(-100)); // keep last 100
    return { success:true };
  }catch(e){ return { success:false }; }
});

// Persist session end
ipcMain.handle('persist-transcript-session-end', (_, { sessionId, endedAt }) => {
  try{
    const sessions = readAdaptiveJson(ADAPTIVE_SESSIONS_FILE, []);
    const s = sessions.find(x => x.id === sessionId);
    if(s) { s.endedAt = endedAt; writeAdaptiveJson(ADAPTIVE_SESSIONS_FILE, sessions); }
    return { success:true };
  }catch(e){ return { success:false }; }
});

// Persist a single transcript chunk
ipcMain.handle('persist-transcript-chunk', (_, chunk) => {
  try{
    ensureAdaptiveDir();
    if(!chunk?.id || !chunk?.sessionId) return { success:false };
    // Store by session — one file per session, append chunks
    const file = path.join(ADAPTIVE_CHUNKS_DIR, `${chunk.sessionId}.json`);
    const chunks = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file,'utf8')) : [];
    chunks.push(chunk);
    fs.writeFileSync(file, JSON.stringify(chunks, null, 2));
    return { success:true };
  }catch(e){ return { success:false }; }
});

// Persist correction events (for learning)
ipcMain.handle('persist-correction-events', (_, events) => {
  try{
    const all = readAdaptiveJson(ADAPTIVE_EVENTS_FILE, []);
    all.push(...(events || []));
    // Keep last 2000 events
    writeAdaptiveJson(ADAPTIVE_EVENTS_FILE, all.slice(-2000));
    return { success:true };
  }catch(e){ return { success:false }; }
});

ipcMain.handle('upsert-adaptive-speaker-profile', (_, profile = {}) => {
  try{
    const profiles = readAdaptiveJson(ADAPTIVE_PROFILES_FILE, []);
    const now = new Date().toISOString();
    const row = {
      id: profile.id || `profile_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      name: profile.name || 'Unnamed Profile',
      description: profile.description || '',
      isDefault: !!profile.isDefault,
      isActive: profile.isActive !== false,
      createdAt: profile.createdAt || now,
      updatedAt: now,
      preferredBibleVersion: profile.preferredBibleVersion || 'KJV',
      commonBooks: Array.isArray(profile.commonBooks) ? profile.commonBooks : [],
      phrasePatterns: Array.isArray(profile.phrasePatterns) ? profile.phrasePatterns : [],
      vocabularyBias: Array.isArray(profile.vocabularyBias) ? profile.vocabularyBias : [],
    };
    const i = profiles.findIndex(p => String(p.id) === String(row.id));
    if (i >= 0) profiles[i] = { ...profiles[i], ...row, updatedAt: now };
    else profiles.push(row);
    writeAdaptiveJson(ADAPTIVE_PROFILES_FILE, profiles);
    return { success:true, profile: row };
  }catch(e){ return { success:false, error:e.message }; }
});
ipcMain.handle('delete-adaptive-speaker-profile', (_, profileId) => {
  try{
    const profiles = readAdaptiveJson(ADAPTIVE_PROFILES_FILE, []);
    writeAdaptiveJson(ADAPTIVE_PROFILES_FILE, profiles.filter(p => String(p.id) !== String(profileId)));
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});
ipcMain.handle('upsert-adaptive-rule', (_, rule = {}) => {
  try{
    const rules = readAdaptiveJson(ADAPTIVE_RULES_FILE, []);
    const now = new Date().toISOString();
    const row = {
      id: rule.id || `rule_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      scope: rule.scope || 'global',
      sourceText: rule.sourceText || '',
      targetText: rule.targetText || '',
      ruleType: rule.ruleType || 'phrase',
      speakerProfileId: rule.speakerProfileId || null,
      confidence: Number(rule.confidence || 1),
      hitCount: Number(rule.hitCount || 0),
      approvedCount: Number(rule.approvedCount || 0),
      rejectedCount: Number(rule.rejectedCount || 0),
      isActive: rule.isActive !== false,
      createdAt: rule.createdAt || now,
      updatedAt: now,
    };
    const i = rules.findIndex(r => String(r.id) === String(row.id));
    if (i >= 0) rules[i] = { ...rules[i], ...row, updatedAt: now };
    else rules.push(row);
    writeAdaptiveJson(ADAPTIVE_RULES_FILE, rules);
    return { success:true, rule: row };
  }catch(e){ return { success:false, error:e.message }; }
});
ipcMain.handle('delete-adaptive-rule', (_, ruleId) => {
  try{
    const rules = readAdaptiveJson(ADAPTIVE_RULES_FILE, []);
    writeAdaptiveJson(ADAPTIVE_RULES_FILE, rules.filter(r => String(r.id) !== String(ruleId)));
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});
ipcMain.handle('upsert-adaptive-vocab', (_, vocab = {}) => {
  try{
    const vocabRows = readAdaptiveJson(ADAPTIVE_VOCAB_FILE, []);
    const now = new Date().toISOString();
    const row = {
      id: vocab.id || `vocab_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      scope: vocab.scope || 'global',
      canonicalTerm: vocab.canonicalTerm || '',
      aliases: Array.isArray(vocab.aliases) ? vocab.aliases : [],
      category: vocab.category || 'general',
      speakerProfileId: vocab.speakerProfileId || null,
      usageCount: Number(vocab.usageCount || 0),
      isActive: vocab.isActive !== false,
      createdAt: vocab.createdAt || now,
      updatedAt: now,
    };
    const i = vocabRows.findIndex(v => String(v.id) === String(row.id));
    if (i >= 0) vocabRows[i] = { ...vocabRows[i], ...row, updatedAt: now };
    else vocabRows.push(row);
    writeAdaptiveJson(ADAPTIVE_VOCAB_FILE, vocabRows);
    return { success:true, vocab: row };
  }catch(e){ return { success:false, error:e.message }; }
});
ipcMain.handle('delete-adaptive-vocab', (_, vocabId) => {
  try{
    const vocabRows = readAdaptiveJson(ADAPTIVE_VOCAB_FILE, []);
    writeAdaptiveJson(ADAPTIVE_VOCAB_FILE, vocabRows.filter(v => String(v.id) !== String(vocabId)));
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});
ipcMain.handle('approve-adaptive-suggestion', (_, suggestionId) => {
  try{
    const queue = readAdaptiveJson(ADAPTIVE_LEARNING_FILE, []);
    const rules = readAdaptiveJson(ADAPTIVE_RULES_FILE, []);
    const i = queue.findIndex(s => String(s.id) === String(suggestionId));
    if (i === -1) return { success:false, error:'Suggestion not found' };
    const s = queue[i];
    rules.push({
      id: `rule_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      scope: s.scope || 'global',
      sourceText: s.sourceText || '',
      targetText: s.targetText || '',
      ruleType: s.ruleType || 'phrase',
      speakerProfileId: s.speakerProfileId || null,
      confidence: Number(s.confidence || 0.86),
      hitCount: Number(s.hitCount || 0),
      approvedCount: Number(s.hitCount || 1),
      rejectedCount: 0,
      isActive: true,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
    queue.splice(i,1);
    writeAdaptiveJson(ADAPTIVE_RULES_FILE, rules);
    writeAdaptiveJson(ADAPTIVE_LEARNING_FILE, queue);
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});
ipcMain.handle('reject-adaptive-suggestion', (_, suggestionId) => {
  try{
    const queue = readAdaptiveJson(ADAPTIVE_LEARNING_FILE, []);
    writeAdaptiveJson(ADAPTIVE_LEARNING_FILE, queue.filter(s => String(s.id) !== String(suggestionId)));
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});
ipcMain.handle('reset-adaptive-learning-data', () => {
  try{
    writeAdaptiveJson(ADAPTIVE_LEARNING_FILE, []);
    writeDetectionJson(DETECTION_EVENTS_FILE, []);
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('get-adaptive-dashboard', () => {
  try{
    ensureAdaptiveDir();
    const rules    = readAdaptiveJson(ADAPTIVE_RULES_FILE, []);
    const vocab    = readAdaptiveJson(ADAPTIVE_VOCAB_FILE, []);
    const profiles = readAdaptiveJson(ADAPTIVE_PROFILES_FILE, []);
    const queue    = readAdaptiveJson(ADAPTIVE_LEARNING_FILE, []);
    const events   = readAdaptiveJson(ADAPTIVE_EVENTS_FILE, []);
    const sessions = readAdaptiveJson(ADAPTIVE_SESSIONS_FILE, []);
    const activeRules   = rules.filter(r => r.isActive !== 0).length;
    const activeVocab   = vocab.filter(v => v.isActive !== 0).length;
    const totalHits     = rules.reduce((s, r) => s + (r.hitCount || 0), 0);
    const recentSession = sessions.length ? sessions[sessions.length - 1] : null;
    const promotedRules = rules.filter(r => r.createdBy === 'user_approved').length;
    return {
      success: true,
      profiles:      profiles.length,
      activeRules,
      activeVocab,
      pendingSuggestions: queue.length,
      totalCorrectionEvents: events.length,
      totalHits,
      promotedRules,
      recentSession: recentSession ? { id: recentSession.id, startedAt: recentSession.startedAt, endedAt: recentSession.endedAt } : null,
    };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('export-adaptive-data', async () => {
  try{
    const result = await dialog.showSaveDialog(mainWindow, {
      title: 'Export Adaptive Data',
      defaultPath: `anchorcast_adaptive_${Date.now()}.json`,
      filters: [{ name: 'JSON', extensions: ['json'] }],
    });
    if (result.canceled || !result.filePath) return { success:false, canceled:true };
    const payload = {
      exportedAt: new Date().toISOString(),
      rules:    readAdaptiveJson(ADAPTIVE_RULES_FILE, []),
      vocab:    readAdaptiveJson(ADAPTIVE_VOCAB_FILE, []),
      profiles: readAdaptiveJson(ADAPTIVE_PROFILES_FILE, []),
      queue:    readAdaptiveJson(ADAPTIVE_LEARNING_FILE, []),
    };
    fs.writeFileSync(result.filePath, JSON.stringify(payload, null, 2));
    return { success:true, file: result.filePath };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('import-adaptive-data', async (_, opts = {}) => {
  try{
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Import Adaptive Data',
      filters: [{ name: 'JSON', extensions: ['json'] }],
      properties: ['openFile'],
    });
    if (result.canceled || !result.filePaths?.length) return { success:false, canceled:true };
    const raw  = fs.readFileSync(result.filePaths[0], 'utf8');
    const data = JSON.parse(raw);
    ensureAdaptiveDir();
    const mode = opts.mode || 'merge';
    if (mode === 'replace') {
      if (data.rules)    writeAdaptiveJson(ADAPTIVE_RULES_FILE, data.rules);
      if (data.vocab)    writeAdaptiveJson(ADAPTIVE_VOCAB_FILE, data.vocab);
      if (data.profiles) writeAdaptiveJson(ADAPTIVE_PROFILES_FILE, data.profiles);
    } else {
      // Merge — add items not already present by id
      const merge = (file, incoming) => {
        if (!Array.isArray(incoming)) return;
        const existing = readAdaptiveJson(file, []);
        const ids = new Set(existing.map(x => x.id));
        const merged = [...existing, ...incoming.filter(x => !ids.has(x.id))];
        writeAdaptiveJson(file, merged);
      };
      merge(ADAPTIVE_RULES_FILE,    data.rules);
      merge(ADAPTIVE_VOCAB_FILE,    data.vocab);
      merge(ADAPTIVE_PROFILES_FILE, data.profiles);
    }
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('save-service-archive', (_, payload = {}) => {
  try{
    ensureServiceArchiveDir();
    const now = new Date();
    const archiveId = payload.id || `service_${now.getTime()}_${Math.random().toString(36).slice(2,6)}`;
    const filePath = path.join(SERVICE_ARCHIVE_DIR, `${archiveId}.json`);
    const doc = {
      id: archiveId,
      savedAt: now.toISOString(),
      title: payload.title || `Service ${now.toLocaleDateString()}`,
      serviceDate: payload.serviceDate || now.toISOString(),
      speaker: payload.speaker || '',
      transcriptLines: Array.isArray(payload.transcriptLines) ? payload.transcriptLines : [],
      detections: Array.isArray(payload.detections) ? payload.detections : [],
      replayTimeline: Array.isArray(payload.replayTimeline) ? payload.replayTimeline : [],
      report: payload.report || {},
      schedule: payload.schedule || null,
      mediaUsed: Array.isArray(payload.mediaUsed) ? payload.mediaUsed : [],
      notes: payload.notes || '',
    };
    fs.writeFileSync(filePath, JSON.stringify(doc, null, 2));
    const index = readServiceArchiveIndex();
    const searchText = [
      doc.title, doc.speaker, doc.notes,
      ...(doc.transcriptLines || []).map(x => x.text || x.raw || ''),
      ...(doc.detections || []).map(x => x.ref || ''),
      ...(doc.mediaUsed || []).map(x => x.title || x.name || '')
    ].join(' ').toLowerCase();
    const entry = {
      id: doc.id,
      file: filePath,
      savedAt: doc.savedAt,
      title: doc.title,
      serviceDate: doc.serviceDate,
      speaker: doc.speaker,
      transcriptCount: (doc.transcriptLines || []).length,
      detectionCount: (doc.detections || []).length,
      mediaCount: (doc.mediaUsed || []).length,
      searchText,
    };
    const idx = index.findIndex(x => String(x.id) === String(doc.id));
    if (idx >= 0) index[idx] = entry;
    else index.unshift(entry);
    writeServiceArchiveIndex(index.slice(0,1000));
    return { success:true, id: doc.id, file: filePath };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('get-service-archive-index', () => {
  try{
    return { success:true, items: readServiceArchiveIndex() };
  }catch(e){ return { success:false, error:e.message, items:[] }; }
});

ipcMain.handle('search-service-archive', (_, query = '') => {
  try{
    const q = String(query || '').trim().toLowerCase();
    const items = readServiceArchiveIndex();
    if (!q) return { success:true, items };
    const terms = q.split(/\s+/).filter(Boolean);
    const filtered = items.filter(item => terms.every(t => String(item.searchText || '').includes(t) || String(item.title || '').toLowerCase().includes(t) || String(item.speaker || '').toLowerCase().includes(t)));
    return { success:true, items: filtered };
  }catch(e){ return { success:false, error:e.message, items:[] }; }
});

ipcMain.handle('load-service-archive-item', (_, archiveId) => {
  try{
    const item = readServiceArchiveIndex().find(x => String(x.id) === String(archiveId));
    if (!item || !item.file || !fs.existsSync(item.file)) return { success:false, error:'Archive item not found' };
    const doc = JSON.parse(fs.readFileSync(item.file, 'utf8'));
    return { success:true, item: doc };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('get-auto-service-suggestions', async (_, payload = {}) => {
  try{
    ensureServiceArchiveDir();
    const index = readServiceArchiveIndex();
    const limit = Math.max(1, Math.min(25, Number(payload.limit || 12)));
    const recent = index.slice(0, limit);
    const speakerHint = String(payload.speaker || '').trim().toLowerCase();

    const songFreq = new Map();
    const verseFreq = new Map();
    const titleFreq = new Map();

    for (const item of recent) {
      if (!item?.file || !fs.existsSync(item.file)) continue;
      try {
        const doc = JSON.parse(fs.readFileSync(item.file, 'utf8'));
        const transcript = Array.isArray(doc.transcriptLines) ? doc.transcriptLines : [];
        const detections = Array.isArray(doc.detections) ? doc.detections : [];
        const replay = Array.isArray(doc.replayTimeline) ? doc.replayTimeline : [];
        const speaker = String(doc.speaker || '').trim().toLowerCase();
        const weight = (speakerHint && speaker && speaker === speakerHint) ? 2 : 1;

        for (const d of detections) {
          const ref = String(d?.ref || '').trim();
          if (!ref) continue;
          verseFreq.set(ref, (verseFreq.get(ref) || 0) + weight);
        }

        for (const ev of replay) {
          const summary = String(ev?.payload?.summary || ev?.payload?.title || ev?.payload?.ref || '').trim();
          if (!summary) continue;
          const lower = summary.toLowerCase();
          if (lower.includes('song') || lower.includes('chorus') || lower.includes('worship')) {
            songFreq.set(summary, (songFreq.get(summary) || 0) + weight);
          }
          titleFreq.set(summary, (titleFreq.get(summary) || 0) + 1);
        }

        for (const line of transcript.slice(0, 20)) {
          const txt = String(line?.text || line?.raw || '').trim();
          if (!txt) continue;
          if (txt.length <= 90) titleFreq.set(txt, (titleFreq.get(txt) || 0) + 1);
        }
      } catch (_) {}
    }

    const sortMap = (m) => [...m.entries()].sort((a,b) => b[1] - a[1]).map(([value,count]) => ({ value, count }));
    const openingSongs = sortMap(songFreq).slice(0, 5);
    const scriptures = sortMap(verseFreq).slice(0, 8);
    const likelyMoments = sortMap(titleFreq).slice(0, 10);

    return {
      success:true,
      basedOnServices: recent.length,
      openingSongs,
      scriptures,
      likelyMoments,
      template: [
        { slot:'Opening Song', suggestions: openingSongs.slice(0,3) },
        { slot:'Scripture', suggestions: scriptures.slice(0,3) },
        { slot:'Response / Closing', suggestions: openingSongs.slice(3,5) }
      ]
    };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('build-auto-service-schedule', async (_, payload = {}) => {
  try{
    const suggestionRes = await ipcMain.invoke ? null : null;
    const suggestions = await (async () => {
      ensureServiceArchiveDir();
      const index = readServiceArchiveIndex();
      const recent = index.slice(0, Math.max(1, Math.min(25, Number(payload.limit || 12))));
      const speakerHint = String(payload.speaker || '').trim().toLowerCase();
      const verseFreq = new Map();
      for (const item of recent) {
        if (!item?.file || !fs.existsSync(item.file)) continue;
        try {
          const doc = JSON.parse(fs.readFileSync(item.file, 'utf8'));
          const detections = Array.isArray(doc.detections) ? doc.detections : [];
          const speaker = String(doc.speaker || '').trim().toLowerCase();
          const weight = (speakerHint && speaker && speaker === speakerHint) ? 2 : 1;
          for (const d of detections) {
            const ref = String(d?.ref || '').trim();
            if (ref) verseFreq.set(ref, (verseFreq.get(ref) || 0) + weight);
          }
        } catch (_) {}
      }
      return [...verseFreq.entries()].sort((a,b)=>b[1]-a[1]).slice(0,3).map(([value,count]) => ({ value, count }));
    })();

    const now = new Date();
    const serviceTitle = payload.title || `Auto Service ${now.toLocaleDateString()}`;
    const speaker = payload.speaker || '';
    const items = [];
    let idx = 1;
    for (const verse of suggestions) {
      items.push({
        id: `auto_item_${Date.now()}_${idx++}`,
        type: 'bible',
        title: verse.value,
        ref: verse.value,
        meta: { source: 'auto-service-builder', count: verse.count }
      });
    }

    return {
      success:true,
      schedule: {
        id: `auto_schedule_${Date.now()}`,
        name: serviceTitle,
        speaker,
        createdAt: now.toISOString(),
        items
      }
    };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('get-live-smart-suggestions', async (_, payload = {}) => {
  try{
    ensureServiceArchiveDir();
    const index = readServiceArchiveIndex();
    const limit = Math.max(1, Math.min(30, Number(payload.limit || 15)));
    const recent = index.slice(0, limit);

    const transcript = String(payload.transcript || '').trim().toLowerCase();
    const speakerHint = String(payload.speaker || '').trim().toLowerCase();
    const currentBookHint = String(payload.currentBook || '').trim().toLowerCase();

    const verseScores = new Map();
    const songScores = new Map();
    const momentScores = new Map();

    const seedTerms = transcript.split(/\s+/).filter(Boolean).slice(-18);

    for (const item of recent) {
      if (!item?.file || !fs.existsSync(item.file)) continue;
      try {
        const doc = JSON.parse(fs.readFileSync(item.file, 'utf8'));
        const detections = Array.isArray(doc.detections) ? doc.detections : [];
        const replay = Array.isArray(doc.replayTimeline) ? doc.replayTimeline : [];
        const lines = Array.isArray(doc.transcriptLines) ? doc.transcriptLines : [];
        const speaker = String(doc.speaker || '').trim().toLowerCase();
        const baseWeight = (speakerHint && speaker && speaker === speakerHint) ? 2 : 1;

        for (const det of detections) {
          const ref = String(det?.ref || '').trim();
          if (!ref) continue;
          let score = baseWeight;
          const lowerRef = ref.toLowerCase();
          if (currentBookHint && lowerRef.includes(currentBookHint)) score += 2;
          if (seedTerms.some(t => t && lowerRef.includes(t))) score += 1;
          verseScores.set(ref, (verseScores.get(ref) || 0) + score);
        }

        for (const ev of replay) {
          const summary = String(ev?.payload?.summary || ev?.payload?.title || ev?.payload?.ref || '').trim();
          if (!summary) continue;
          const lower = summary.toLowerCase();
          let score = baseWeight;
          if (seedTerms.some(t => t && lower.includes(t))) score += 1;

          if (lower.includes('song') || lower.includes('chorus') || lower.includes('worship')) {
            songScores.set(summary, (songScores.get(summary) || 0) + score);
          }
          momentScores.set(summary, (momentScores.get(summary) || 0) + score);
        }

        for (const line of lines.slice(0, 25)) {
          const txt = String(line?.text || line?.raw || '').trim();
          if (!txt) continue;
          const lowerTxt = txt.toLowerCase();
          let score = 0;
          for (const t of seedTerms) {
            if (t && lowerTxt.includes(t)) score += 1;
          }
          if (score > 0 && txt.length <= 90) {
            momentScores.set(txt, (momentScores.get(txt) || 0) + score);
          }
        }
      } catch (_) {}
    }

    const sortMap = (m) => [...m.entries()].sort((a,b) => b[1] - a[1]).map(([value,score]) => ({ value, score }));
    return {
      success:true,
      transcriptSeed: transcript,
      scriptureSuggestions: sortMap(verseScores).slice(0, 6),
      songSuggestions: sortMap(songScores).slice(0, 5),
      momentSuggestions: sortMap(momentScores).slice(0, 6),
    };
  }catch(e){ return { success:false, error:e.message, scriptureSuggestions:[], songSuggestions:[], momentSuggestions:[] }; }
});

ipcMain.handle('generate-sermon-intelligence', async (_, payload = {}) => {
  try{
    const transcriptLines = Array.isArray(payload.transcriptLines) ? payload.transcriptLines : [];
    const detections = Array.isArray(payload.detections) ? payload.detections : [];
    const archiveHint = payload.useArchive !== false;

    const rawText = transcriptLines.map(x => String(x?.text || x?.raw || '')).join(' ');
    const text = rawText.replace(/\s+/g, ' ').trim();
    const words = text.toLowerCase().match(/[a-z0-9']+/g) || [];
    const freq = new Map();
    for (const w of words) {
      if (!w || w.length < 4 || STOPWORDS.has(w)) continue;
      freq.set(w, (freq.get(w) || 0) + 1);
    }
    const topKeywords = [...freq.entries()].sort((a,b)=>b[1]-a[1]).slice(0, 12).map(([word,count]) => ({ word, count }));

    const verses = detections.map(d => String(d?.ref || '').trim()).filter(Boolean);
    const verseCount = new Map();
    for (const v of verses) verseCount.set(v, (verseCount.get(v) || 0) + 1);
    const recurringVerses = [...verseCount.entries()].sort((a,b)=>b[1]-a[1]).slice(0, 8).map(([ref,count]) => ({ ref, count }));

    const sentences = text.split(/(?<=[.!?])\s+/).map(s => s.trim()).filter(Boolean);
    const keyPoints = sentences
      .filter(s => s.length > 35 && s.length < 220)
      .sort((a,b) => {
        const score = (s) => topKeywords.reduce((acc, k) => acc + ((s.toLowerCase().includes(k.word) ? k.count : 0)), 0);
        return score(b) - score(a);
      })
      .slice(0, 6);

    const structure = {
      intro: sentences.slice(0, 3),
      body: sentences.slice(Math.max(0, Math.floor(sentences.length * 0.25)), Math.max(0, Math.floor(sentences.length * 0.25)) + 4),
      closing: sentences.slice(-3),
    };

    let titleSeed = '';
    if (recurringVerses.length) {
      titleSeed = recurringVerses[0].ref;
    } else if (topKeywords.length >= 2) {
      titleSeed = `${topKeywords[0].word} & ${topKeywords[1].word}`;
    } else if (topKeywords.length) {
      titleSeed = topKeywords[0].word;
    } else {
      titleSeed = 'Sermon Reflection';
    }

    const titleSuggestions = [
      `Walking in ${titleSeed}`,
      `The Message of ${titleSeed}`,
      `${titleSeed}: A Sermon Reflection`,
    ];

    let archiveThemes = [];
    if (archiveHint) {
      try{
        ensureServiceArchiveDir();
        const items = readServiceArchiveIndex().slice(0, 20);
        const archiveWords = new Map();
        for (const item of items) {
          if (!item?.file || !fs.existsSync(item.file)) continue;
          try{
            const doc = JSON.parse(fs.readFileSync(item.file, 'utf8'));
            const dets = Array.isArray(doc.detections) ? doc.detections : [];
            for (const d of dets) {
              const ref = String(d?.ref || '').trim();
              if (ref && verses.includes(ref)) archiveWords.set(ref, (archiveWords.get(ref) || 0) + 1);
            }
          } catch(_) {}
        }
        archiveThemes = [...archiveWords.entries()].sort((a,b)=>b[1]-a[1]).slice(0, 6).map(([ref,count]) => ({ ref, count }));
      } catch(_) {}
    }

    return {
      success: true,
      titleSuggestions,
      topKeywords,
      recurringVerses,
      keyPoints,
      structure,
      archiveThemes,
      stats: {
        transcriptLines: transcriptLines.length,
        transcriptWords: words.length,
        detections: detections.length,
        uniqueVerses: verseCount.size,
      }
    };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('get-analytics-dashboard', async (_, payload = {}) => {
  try{
    ensureServiceArchiveDir();
    const items = readServiceArchiveIndex().slice(0, Math.max(1, Math.min(250, Number(payload.limit || 120))));
    const verseMap = new Map();
    const bookMap = new Map();
    const songMap = new Map();
    const speakerMap = new Map();
    const keywordMap = new Map();
    const serviceLengths = [];
    const detectionStats = [];

    for (const item of items) {
      if (!item?.file || !fs.existsSync(item.file)) continue;
      try{
        const doc = JSON.parse(fs.readFileSync(item.file, 'utf8'));
        const detections = Array.isArray(doc.detections) ? doc.detections : [];
        const replay = Array.isArray(doc.replayTimeline) ? doc.replayTimeline : [];
        const transcript = Array.isArray(doc.transcriptLines) ? doc.transcriptLines : [];
        const speaker = String(doc.speaker || 'Unknown speaker').trim() || 'Unknown speaker';

        speakerMap.set(speaker, (speakerMap.get(speaker) || 0) + 1);
        serviceLengths.push({
          title: doc.title || item.title || 'Untitled Service',
          transcriptLines: transcript.length,
          detections: detections.length,
          replayEvents: replay.length,
          serviceDate: doc.serviceDate || doc.savedAt || item.savedAt || null
        });

        const approved = detections.filter(d => d.status === 'approved').length;
        const rejected = detections.filter(d => d.status === 'rejected').length;
        detectionStats.push({ approved, rejected, total: detections.length });

        for (const d of detections) {
          const ref = String(d?.ref || '').trim();
          if (!ref) continue;
          verseMap.set(ref, (verseMap.get(ref) || 0) + 1);
          const m = ref.match(/^(.+?)\s+\d+:\d+$/);
          const book = m ? m[1] : ref.split(' ')[0];
          if (book) bookMap.set(book, (bookMap.get(book) || 0) + 1);
        }

        for (const ev of replay) {
          const summary = String(ev?.payload?.summary || ev?.payload?.title || ev?.payload?.ref || '').trim();
          if (!summary) continue;
          const lower = summary.toLowerCase();
          if (lower.includes('song') || lower.includes('chorus') || lower.includes('worship')) {
            songMap.set(summary, (songMap.get(summary) || 0) + 1);
          }
        }

        const text = transcript.map(x => String(x?.text || x?.raw || '')).join(' ').toLowerCase();
        const words = text.match(/[a-z0-9']+/g) || [];
        for (const w of words) {
          if (!w || w.length < 4 || STOPWORDS.has(w)) continue;
          keywordMap.set(w, (keywordMap.get(w) || 0) + 1);
        }
      }catch(_){}
    }

    const sortMap = (m, label='value') => [...m.entries()].sort((a,b)=>b[1]-a[1]).slice(0, 15).map(([value,count]) => ({ [label]: value, count }));
    const totalApproved = detectionStats.reduce((a,b)=>a + Number(b.approved || 0), 0)
    const totalRejected = detectionStats.reduce((a,b)=>a + Number(b.rejected || 0), 0)
    const totalDetections = detectionStats.reduce((a,b)=>a + Number(b.total || 0), 0)
    const avgTranscriptLines = serviceLengths.length ? Math.round(serviceLengths.reduce((a,b)=>a + Number(b.transcriptLines || 0), 0) / serviceLengths.length) : 0;

    return {
      success:true,
      totals: {
        archivedServices: items.length,
        totalDetections,
        totalApproved,
        totalRejected,
        avgTranscriptLines
      },
      mostUsedBooks: sortMap(bookMap, 'book'),
      mostQuotedVerses: sortMap(verseMap, 'ref'),
      mostUsedSongs: sortMap(songMap, 'title'),
      speakerPatterns: sortMap(speakerMap, 'speaker'),
      keywordFrequency: sortMap(keywordMap, 'keyword'),
      serviceLengths: serviceLengths.sort((a,b)=> (new Date(b.serviceDate || 0)) - (new Date(a.serviceDate || 0))).slice(0, 20)
    };
  }catch(e){ return { success:false, error:e.message }; }
});

ipcMain.handle('save-clip-package', async (_, payload = {}) => {
  try{
    ensureClipExportDir();
    const stamp = new Date().toISOString().replace(/[:.]/g,'-');
    const title = String(payload.title || `clip_${stamp}`).replace(/[^\w\- ]+/g,'').trim().replace(/\s+/g,'_') || `clip_${stamp}`;
    const outPath = path.join(CLIP_EXPORT_DIR, `${title}.json`);
    const doc = {
      id: payload.id || `clip_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      createdAt: new Date().toISOString(),
      title: payload.title || title,
      transcriptSnippet: payload.transcriptSnippet || '',
      verseRefs: Array.isArray(payload.verseRefs) ? payload.verseRefs : [],
      replayEvents: Array.isArray(payload.replayEvents) ? payload.replayEvents : [],
      notes: payload.notes || '',
      source: payload.source || 'clip-generator'
    };
    fs.writeFileSync(outPath, JSON.stringify(doc, null, 2));
    return { success:true, file: outPath, id: doc.id };
  }catch(e){ return { success:false, error:e.message }; }
});

// Load learning data for the offline learning job
ipcMain.handle('load-learning-data', () => {
  return {
    correctionEvents: readAdaptiveJson(ADAPTIVE_EVENTS_FILE, []),
    rules: readAdaptiveJson(ADAPTIVE_RULES_FILE, []),
  };
});

// File export
ipcMain.handle('export-file',async(_,{content,defaultName,filters})=>{
  const result=await dialog.showSaveDialog(mainWindow,{
    title:'Export',defaultPath:defaultName||'export.txt',
    filters:filters||[{name:'Text',extensions:['txt']}],
  });
  if(!result.canceled&&result.filePath){
    fs.writeFileSync(result.filePath,content,'utf-8');
    return{success:true,path:result.filePath};
  }
  return{success:false};
});
// Legacy alias
ipcMain.handle('export-transcript',async(_,opts)=>{
  const result=await dialog.showSaveDialog(mainWindow,{
    title:'Export Transcript',defaultPath:opts.defaultName||'transcript.txt',
    filters:[{name:'Text',extensions:['txt']},{name:'Markdown',extensions:['md']}],
  });
  if(!result.canceled&&result.filePath){
    fs.writeFileSync(result.filePath,opts.content,'utf-8');
    return{success:true,path:result.filePath};
  }
  return{success:false};
});

ipcMain.handle('import-notes', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Import Sermon Notes',
    filters: [
      { name: 'Notes', extensions: ['txt', 'md', 'json'] },
      { name: 'All Files', extensions: ['*'] },
    ],
    properties: ['openFile'],
  });
  if (result.canceled || !result.filePaths.length) return { success: false };
  try {
    const content = fs.readFileSync(result.filePaths[0], 'utf-8');
    return { success: true, content, filePath: result.filePaths[0] };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('get-app-info',()=>({
  version:app.getVersion(),platform:process.platform,
  userData:UD, assetsPath:ASSETS_DIR, dataPath:DATA_DIR,
  localIp:localIp(),httpPort:currentSettings.httpPort||8080,
}));

// Bible search for Presentation Editor scripture picker
ipcMain.handle('pres-bible-search', async (_, { query, translation }) => {
  const safe = sanitizeBibleTranslation(translation || 'kjv');
  let file = null;
  for (const dir of getBibleSearchDirs()) {
    const candidate = path.join(dir, `${safe}.json`);
    if (fs.existsSync(candidate)) { file = candidate; break; }
  }
  if (!file) return { results: [], error: 'Translation not installed' };
  try {
    const data = JSON.parse(fs.readFileSync(file, 'utf-8'));
    const q = (query || '').toLowerCase().trim();
    if (!q) return { results: [] };

    const BOOK_NAMES = ['','Genesis','Exodus','Leviticus','Numbers','Deuteronomy','Joshua','Judges','Ruth','1 Samuel','2 Samuel','1 Kings','2 Kings','1 Chronicles','2 Chronicles','Ezra','Nehemiah','Esther','Job','Psalms','Proverbs','Ecclesiastes','Song of Solomon','Isaiah','Jeremiah','Lamentations','Ezekiel','Daniel','Hosea','Joel','Amos','Obadiah','Jonah','Micah','Nahum','Habakkuk','Zephaniah','Haggai','Zechariah','Malachi','Matthew','Mark','Luke','John','Acts','Romans','1 Corinthians','2 Corinthians','Galatians','Ephesians','Philippians','Colossians','1 Thessalonians','2 Thessalonians','1 Timothy','2 Timothy','Titus','Philemon','Hebrews','James','1 Peter','2 Peter','1 John','2 John','3 John','Jude','Revelation'];

    // Check if it's a reference like "John 3:16"
    const refMatch = q.match(/^(\d?\s*[a-z]+)\s+(\d+)(?::(\d+))?/i);
    let results = [];

    if (refMatch) {
      const bookQ = refMatch[1].toLowerCase().replace(/\s+/,'');
      const chapQ = parseInt(refMatch[2]);
      const verseQ = refMatch[3] ? parseInt(refMatch[3]) : null;
      results = data.filter(v => {
        const bn = BOOK_NAMES[v.b] || '';
        const bMatch = bn.toLowerCase().replace(/\s+/g,'').includes(bookQ) ||
                       bn.toLowerCase().startsWith(bookQ);
        if (!bMatch) return false;
        if (v.c !== chapQ) return false;
        if (verseQ !== null && v.v !== verseQ) return false;
        return true;
      }).slice(0, 50);
    } else {
      // Text search
      results = data.filter(v => v.t.toLowerCase().includes(q)).slice(0, 40);
    }

    return { results: results.map(v => ({
      ref: `${BOOK_NAMES[v.b]} ${v.c}:${v.v}`,
      text: v.t,
      book: v.b, chapter: v.c, verse: v.v
    })) };
  } catch(e) { return { results: [], error: e.message }; }
});

// Bible data loader — reads installed JSON translations into memory
ipcMain.handle('load-bible-data', async () => {
  const result = {};
  for (const dataDir of getBibleSearchDirs()) {
    if (!dataDir || !fs.existsSync(dataDir)) continue;
    try {
      const files = fs.readdirSync(dataDir).filter(f => f.endsWith('.json'));
      for (const file of files) {
        const id = path.basename(file, '.json').toUpperCase();
        if (id.includes('.') || id.includes('-') || id.length > 10) continue;
        if (result[id]) continue;
        try {
          const raw = fs.readFileSync(path.join(dataDir, file), 'utf-8');
          const data = JSON.parse(raw);
          if (Array.isArray(data) && data.length > 0 && data[0].b && data[0].t) {
            result[id] = data;
            console.log(`[BibleDB] Loaded ${data.length} ${id} verses from ${dataDir}`);
          }
        } catch(e) {
          console.warn(`[BibleDB] Failed to load ${file}:`, e.message);
        }
      }
    } catch(e) {}
  }
  return result;
});

// Bible version manager — returns count of verses per installed translation
ipcMain.handle('get-installed-versions', async () => {
  const result = {};
  for (const dataDir of getBibleSearchDirs()) {
    if (!dataDir || !fs.existsSync(dataDir)) continue;
    try {
      const files = fs.readdirSync(dataDir).filter(f => f.endsWith('.json'));
      for (const file of files) {
        const id = path.basename(file, '.json').toUpperCase();
        if (id.includes('.') || id.includes('-') || id.length > 10) continue;
        if (result[id]) continue;
        try {
          const raw = fs.readFileSync(path.join(dataDir, file), 'utf-8');
          const data = JSON.parse(raw);
          if (Array.isArray(data) && data.length > 0 && data[0].b && data[0].t) {
            result[id] = data.length;
          }
        } catch(e) { /* skip malformed files */ }
      }
    } catch(e) {}
  }
  return result;
});

// Save a Bible translation uploaded from the settings UI
ipcMain.handle('save-bible-version', async (_, translation, data) => {
  // KJV is always free; additional translations require registration
  if (translation && translation.toUpperCase() !== 'KJV') {
    const gate = requiresRegistration('Additional Bible Translations');
    if (gate) return { success: false, blocked: true, error: gate.reason };
  }
  try {
    const dataDir = BIBLE_DIR;
    fs.mkdirSync(dataDir, { recursive: true });
    const safe = sanitizeBibleTranslation(translation);
    const file = path.join(dataDir, `${safe}.json`);
    const payload = typeof data === 'string' ? data : JSON.stringify(data);
    fs.writeFileSync(file, payload, 'utf-8');
    let count = 0;
    try { count = Array.isArray(data) ? data.length : JSON.parse(payload).length; } catch(_) {}
    console.log(`[BibleDB] Saved ${count} ${translation} verses to ${file}`);
    if (mainWindow) mainWindow.webContents.send('bible-versions-updated');
    return { success: true, count, verseCount: count };
  } catch(e) {
    return { success: false, error: e.message };
  }
});
ipcMain.handle('import-translation', async (_, { abbrev, data } = {}) => {
  try {
    const dataDir = BIBLE_DIR;
    fs.mkdirSync(dataDir, { recursive: true });
    const safe = sanitizeBibleTranslation(abbrev);
    const file = path.join(dataDir, `${safe}.json`);
    const payload = typeof data === 'string' ? data : JSON.stringify(data);
    fs.writeFileSync(file, payload, 'utf-8');
    let verseCount = 0;
    try { verseCount = Array.isArray(data) ? data.length : JSON.parse(payload).length; } catch(_) {}
    if (mainWindow) mainWindow.webContents.send('bible-versions-updated');
    return { success: true, verseCount, count: verseCount };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Delete a Bible translation
ipcMain.handle('delete-bible-version', async (_, translation) => {
  try {
    const safe = sanitizeBibleTranslation(translation);
    if (safe === 'kjv') return { success: false, error: 'KJV is built in and cannot be removed' };
    const dataDir = BIBLE_DIR;
    const file = path.join(dataDir, `${safe}.json`);
    if (fs.existsSync(file)) {
      fs.unlinkSync(file);
      console.log(`[BibleDB] Deleted ${file}`);
    }
    if (mainWindow) mainWindow.webContents.send('bible-versions-updated');
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
});
ipcMain.handle('delete-translation', async (_, translation) => {
  try {
    const safe = sanitizeBibleTranslation(translation);
    if (safe === 'kjv') return { success: false, error: 'KJV is built in and cannot be removed' };
    const dataDir = BIBLE_DIR;
    const file = path.join(dataDir, `${safe}.json`);
    if (fs.existsSync(file)) fs.unlinkSync(file);
    if (mainWindow) mainWindow.webContents.send('bible-versions-updated');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});
ipcMain.handle('open-history',()=>{ createHistoryWindow(); return{success:true}; });

// ── Whisper Local Server ──────────────────────────────────────────────────────
const WHISPER_PORT   = 7777;
const WHISPER_URL    = `http://127.0.0.1:${WHISPER_PORT}`;
function resolveRuntimeResource(...parts) {
  const candidates = [
    path.join(process.resourcesPath || '', ...parts),
    path.join(app.getAppPath(), ...parts),
    path.join(__dirname, '..', ...parts),
    path.join(process.cwd(), ...parts),
  ].filter(Boolean);
  for (const candidate of candidates) {
    try { if (fs.existsSync(candidate)) return candidate; } catch (_) {}
  }
  return candidates[0];
}
const WHISPER_SCRIPT = resolveRuntimeResource('whisper_server.py');
const WHISPER_CACHE  = path.join(DATA_DIR, 'whisper_python.txt');
let   whisperProc    = null;
let   whisperReady   = false;
let   whisperModel   = 'small.en'; // default model tuned for English sermons

async function checkWhisperRunning() {
  try {
    const res = await fetch(`${WHISPER_URL}/health`, { signal: AbortSignal.timeout(1500) });
    return res.ok;
  } catch { return false; }
}

// Runs a callback after mainWindow's did-finish-load fires.
// If the window is already loaded, waits `extraDelay` ms then runs.
// Prevents IPC messages being lost on first launch before renderer is ready.
function _runAfterWindowReady(fn, extraDelay = 0) {
  const run = () => setTimeout(fn, extraDelay);
  if (!mainWindow || mainWindow.isDestroyed()) {
    // Window not created yet — wait for it
    app.once('browser-window-created', () => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        if (mainWindow.webContents.isLoading()) {
          mainWindow.webContents.once('did-finish-load', run);
        } else {
          run();
        }
      }
    });
    return;
  }
  if (mainWindow.webContents.isLoading()) {
    mainWindow.webContents.once('did-finish-load', run);
  } else {
    run();
  }
}

async function startWhisperServer(model = 'small.en') {
  // Already running?
  if (await checkWhisperRunning()) { whisperReady = true; return true; }

  // Script exists?
  if (!fs.existsSync(WHISPER_SCRIPT)) {
    console.log('[Whisper] whisper_server.py not found');
    return false;
  }

  // ── Find Python ─────────────────────────────────────────────────────────────
  const bundledPy    = resolveRuntimeResource('python', 'python.exe'); // Windows bundled
  const bundledPyMac = resolveRuntimeResource('python', 'bin', 'python3'); // Mac/Linux

  // Check cached Python path first (speeds up subsequent startups)
  let cachedPy = null;
  try {
    const cached = fs.readFileSync(WHISPER_CACHE, 'utf-8').trim();
    if (cached && fs.existsSync(cached)) cachedPy = cached;
  } catch {}

  // Build candidate list — cached and bundled first, then system
  const candidates = [
    ...(cachedPy ? [cachedPy] : []),
    process.platform === 'win32' ? bundledPy : bundledPyMac,
    ...(process.platform === 'win32' ? ['python', 'python3'] : ['python3', 'python']),
  ].filter((v, i, a) => a.indexOf(v) === i); // deduplicate

  const { execFileSync } = require('child_process');

  let pythonExe = null;
  for (const candidate of candidates) {
    try {
      execFileSync(candidate, ['-c',
        'import sys; v=sys.version_info; exit(0 if v.major==3 and 8<=v.minor<=12 else 1)'
      ], { timeout: 4000, windowsHide: true });
      execFileSync(candidate, ['-c',
        'from faster_whisper import WhisperModel'
      ], { timeout: 6000, windowsHide: true });
      pythonExe = candidate;
      const label = (candidate === bundledPy || candidate === bundledPyMac)
        ? 'bundled' : candidate === cachedPy ? 'cached' : 'system';
      console.log(`[Whisper] Found ${label} Python: ${candidate}`);
      // Cache for next startup
      try { fs.writeFileSync(WHISPER_CACHE, candidate, 'utf-8'); } catch {}
      break;
    } catch { /* try next */ }
  }

  if (!pythonExe) {
    // Diagnose WHY Python wasn't found so we can show the user a specific message
    let whisperFailReason = 'no_python'; // default
    for (const candidate of candidates) {
      try {
        // Can we run it at all?
        execFileSync(candidate, ['--version'], { timeout: 3000, windowsHide: true });
        // It runs — check version
        try {
          execFileSync(candidate, ['-c',
            'import sys; v=sys.version_info; exit(0 if v.major==3 and 8<=v.minor<=12 else 1)'
          ], { timeout: 3000, windowsHide: true });
          // Version OK — faster_whisper must be missing
          whisperFailReason = 'no_faster_whisper';
        } catch {
          // Python exists but wrong version
          let vstr = 'unknown';
          try {
            const r = execFileSync(candidate, ['-c', 'import sys; print(sys.version)'],
              { encoding: 'utf-8', timeout: 3000, windowsHide: true });
            vstr = r.trim().split(' ')[0];
          } catch {}
          whisperFailReason = `wrong_version:${vstr}`;
        }
        break;
      } catch { /* candidate not found at all */ }
    }
    console.log(`[Whisper] Setup needed — reason: ${whisperFailReason}`);
    // Send detailed status to renderer so it can show a specific message
    mainWindow?.webContents.send('whisper-setup-needed', {
      reason: whisperFailReason,
      setupBatExists: fs.existsSync(resolveRuntimeResource('setup_whisper.bat')),
    });
    return false;
  }

  // ── Spawn the server ─────────────────────────────────────────────────────────
  whisperModel = model;
  console.log(`[Whisper] Starting server (model: ${model})...`);

  // Ensure AppData model directory exists
  try { fs.mkdirSync(WHISPER_MODEL_DIR, { recursive: true }); } catch(_) {}

  // On first run: copy bundled models from resources/models/ to AppData/WhisperModels/
  // so models survive app updates and reinstalls without re-downloading
  const bundledModels = resolveRuntimeResource('models');
  if (fs.existsSync(bundledModels)) {
    try {
      const entries = fs.readdirSync(bundledModels);
      for (const entry of entries) {
        const src  = path.join(bundledModels, entry);
        const dest = path.join(WHISPER_MODEL_DIR, entry);
        if (!fs.existsSync(dest)) {
          console.log(`[Whisper] Copying bundled model: ${entry}`);
          // Recursive copy
          const copyDir = (s, d) => {
            fs.mkdirSync(d, { recursive: true });
            for (const f of fs.readdirSync(s)) {
              const sp = path.join(s, f), dp = path.join(d, f);
              if (fs.statSync(sp).isDirectory()) copyDir(sp, dp);
              else fs.copyFileSync(sp, dp);
            }
          };
          if (fs.statSync(src).isDirectory()) copyDir(src, dest);
          else fs.copyFileSync(src, dest);
        }
      }
      console.log('[Whisper] Bundled models available at:', WHISPER_MODEL_DIR);
    } catch(e) {
      console.warn('[Whisper] Model copy warning:', e.message);
    }
  }

  whisperProc = spawn(pythonExe, [
    WHISPER_SCRIPT,
    '--model',           model,
    '--port',            String(WHISPER_PORT),
    '--model_cache_dir', WHISPER_MODEL_DIR,
  ], {
    stdio:       ['ignore', 'pipe', 'pipe'],
    detached:    false,
    windowsHide: true,   // suppress console window on Windows
    // NO shell:true — we use execFile-style with full path, avoids quoting issues
  });

  let _whisperStderr = '';
  whisperProc.stdout.on('data', d => process.stdout.write('[Whisper] ' + d));
  whisperProc.stderr.on('data', d => {
    const line = String(d);
    _whisperStderr += line;
    process.stderr.write('[Whisper] ' + line);
  });
  whisperProc.on('error', e => {
    console.warn('[Whisper] Failed to start process:', e.message);
    whisperReady = false; whisperProc = null;
  });
  whisperProc.on('exit', (code) => {
    whisperReady = false;
    whisperProc  = null;
    if (code !== 0 && code !== null) {
      console.log(`[Whisper] Server exited (code ${code})`);
      // Detect model-not-found vs other errors
      if (_whisperStderr.includes('MARKER:model_not_found') ||
          _whisperStderr.includes('No such file') ||
          _whisperStderr.includes('model_not_found')) {
        mainWindow?.webContents.send('whisper-setup-needed', {
          reason: 'model_not_found',
          model: whisperModel,
          setupBatExists: fs.existsSync(resolveRuntimeResource('setup_whisper.bat')),
        });
      } else if (_whisperStderr.includes('ctranslate2.dll') ||
                 _whisperStderr.includes('one of its dependencies') ||
                 _whisperStderr.includes('msvcp140') ||
                 _whisperStderr.includes('vcruntime140')) {
        // ctranslate2 failed to load — VC++ runtime DLLs are missing
        mainWindow?.webContents.send('whisper-setup-needed', {
          reason: 'missing_vcredist',
          setupBatExists: fs.existsSync(resolveRuntimeResource('setup_whisper.bat')),
        });
      }
    }
    mainWindow?.webContents.send('whisper-status', { ready: false });
    _whisperStderr = '';
  });

  // ── Wait for server to be ready (model load can take 10–30s first time) ─────
  console.log('[Whisper] Waiting for model to load...');
  const deadline = Date.now() + 90000; // 90s max
  while (Date.now() < deadline) {
    await new Promise(r => setTimeout(r, 2000));
    if (!whisperProc) { console.warn('[Whisper] Process died during startup'); return false; }
    if (await checkWhisperRunning()) {
      whisperReady = true;
      console.log('[Whisper] Ready ✓ — local transcription active');
      mainWindow?.webContents.send('whisper-status', { ready: true, model });
      return true;
    }
  }

  console.warn('[Whisper] Timed out waiting for server');
  // If the process already died, it may have sent whisper-setup-needed via exit handler
  // If still running but not responding, kill it and show model error
  if (whisperProc && !whisperProc.killed) {
    try { whisperProc.kill(); } catch(_) {}
    whisperProc = null;
    mainWindow?.webContents.send('whisper-setup-needed', {
      reason: 'model_not_found',
      model: whisperModel,
      setupBatExists: fs.existsSync(resolveRuntimeResource('setup_whisper.bat')),
    });
  }
  return false;
}

function stopWhisperServer() {
  if (whisperProc) {
    whisperProc.kill();
    whisperProc  = null;
    whisperReady = false;
  }
}

ipcMain.handle('whisper-status',  () => ({ ready: whisperReady, model: whisperModel }));
ipcMain.handle('whisper-start',   async (_, model) => {
  const gate = requiresRegistration('AI Transcription');
  if (gate) return { ready: false, blocked: true, error: gate.reason };
  // Auto-migrate multilingual models to English-only variants (better accuracy for sermons)
  const _m = String(model || 'small.en');
  const resolvedModel = _m === 'small' ? 'small.en'
    : _m === 'base'  ? 'base.en'
    : _m === 'tiny'  ? 'tiny.en'
    : _m;
  const ok = await startWhisperServer(resolvedModel);
  return { ready: ok, model: whisperModel };
});
ipcMain.handle('whisper-stop',    () => { stopWhisperServer(); return { ready: false }; });
// Reinforcement: when detection identifies a verse, inject its text into Whisper's context
// so subsequent transcription chunks are biased toward that passage's vocabulary.
ipcMain.handle('whisper-reinforce', async (_, { text, ttl = 30 } = {}) => {
  if (!whisperReady || !text) return { ok: false };
  try {
    const res = await fetch(`${WHISPER_URL}/reinforce`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: String(text).slice(0, 400), ttl }),
      signal: AbortSignal.timeout(2000),
    });
    return { ok: res.ok };
  } catch (_) {
    return { ok: false };
  }
});
// List all models with installed status
ipcMain.handle('get-whisper-models', () => {
  try { fs.mkdirSync(WHISPER_MODEL_DIR, { recursive: true }); } catch(_) {}
  return WHISPER_MODELS.map(m => ({
    ...m,
    installed: isModelInstalled(m.id),
  }));
});

// Download a specific model into WHISPER_MODEL_DIR
let _modelDownloadProc = null;
ipcMain.handle('download-whisper-model', async (_, modelId) => {
  const info = WHISPER_MODELS.find(x => x.id === modelId);
  if (!info) return { ok: false, error: 'Unknown model: ' + modelId };
  if (isModelInstalled(modelId)) return { ok: true, already: true };

  const pythonExe = resolveRuntimeResource('python', 'python.exe');
  if (!fs.existsSync(pythonExe)) return { ok: false, error: 'Python not found' };

  try { fs.mkdirSync(WHISPER_MODEL_DIR, { recursive: true }); } catch(_) {}

  // Download model using Python in background — sends progress via IPC
  const script = `
import sys
from faster_whisper import WhisperModel
print(f'Downloading {sys.argv[1]}...', flush=True)
WhisperModel(sys.argv[1], device='cpu', compute_type='int8', download_root=sys.argv[2])
print('DONE', flush=True)
`.trim();

  return new Promise((resolve) => {
    const proc = spawn(pythonExe, ['-c', script, modelId, WHISPER_MODEL_DIR], {
      windowsHide: true, stdio: ['ignore', 'pipe', 'pipe']
    });
    _modelDownloadProc = proc;

    proc.stdout.on('data', d => {
      const line = String(d).trim();
      mainWindow?.webContents.send('whisper-model-progress', { modelId, line });
    });
    proc.stderr.on('data', d => {
      const line = String(d).trim();
      if (line) mainWindow?.webContents.send('whisper-model-progress', { modelId, line });
    });
    proc.on('exit', code => {
      _modelDownloadProc = null;
      const installed = isModelInstalled(modelId);
      mainWindow?.webContents.send('whisper-model-downloaded', { modelId, success: installed });
      resolve({ ok: installed });
    });
    proc.on('error', e => {
      _modelDownloadProc = null;
      resolve({ ok: false, error: e.message });
    });
  });
});

ipcMain.handle('whisper-diagnostics', () => {
  const bundledPyWin = resolveRuntimeResource('python', 'python.exe');
  const bundledPyNix = resolveRuntimeResource('python', 'bin', 'python3');
  const setupBatPath = resolveRuntimeResource('setup_whisper.bat');
  return {
    ready: whisperReady,
    model: whisperModel,
    script: WHISPER_SCRIPT,
    scriptExists: fs.existsSync(WHISPER_SCRIPT),
    bundledPyWin,
    bundledPyWinExists: fs.existsSync(bundledPyWin),
    bundledPyNix,
    bundledPyNixExists: fs.existsSync(bundledPyNix),
    setupBatPath,
    setupBatExists: fs.existsSync(setupBatPath),
    resourcesPath: process.resourcesPath,
    appPath: app.getAppPath()
  };
});

// Trigger setup_whisper.bat from within the app (e.g. when Python not found at startup)
ipcMain.handle('run-whisper-setup', async () => {
  // Check if bundled Python exists — use it directly instead of setup_whisper.bat
  const bundledPy = resolveRuntimeResource('python', 'python.exe');
  const hasBundledPy = fs.existsSync(bundledPy);

  if (hasBundledPy) {
    // Python is bundled — just download the missing model directly, no console window
    const modelDir = WHISPER_MODEL_DIR;
    try { fs.mkdirSync(modelDir, { recursive: true }); } catch(_) {}

    // Write a silent Python script to download the model
    const scriptPath = path.join(app.getPath('temp'), 'ac_get_model.py');
    const script = [
      "import os, warnings, logging",
      "os.environ['HF_HUB_DISABLE_SYMLINKS_WARNING'] = '1'",
      "os.environ['HF_HUB_DISABLE_PROGRESS_BARS'] = '1'",
      "os.environ['HF_TOKEN'] = ''",
      "warnings.filterwarnings('ignore')",
      "logging.disable(logging.CRITICAL)",
      "from faster_whisper import WhisperModel",
      `d = r'${modelDir}'`,
      "import sys",
      "model_id = sys.argv[1] if len(sys.argv) > 1 else 'small.en'",
      "os.makedirs(d, exist_ok=True)",
      "WhisperModel(model_id, device='cpu', compute_type='int8', download_root=d)",
      "print('MODEL_READY', flush=True)",
    ].join('\n');

    fs.writeFileSync(scriptPath, script, 'utf-8');

    return new Promise((resolve) => {
      const proc = spawn(bundledPy, ['-W', 'ignore', scriptPath, whisperModel || 'small.en'], {
        windowsHide: true,
        stdio: ['ignore', 'pipe', 'pipe'],
      });

      let output = '';
      proc.stdout.on('data', d => {
        output += String(d);
        mainWindow?.webContents.send('whisper-model-progress', {
          modelId: whisperModel || 'small.en',
          line: String(d).trim()
        });
      });
      proc.stderr.on('data', d => {
        const line = String(d).trim();
        if (line && !line.includes('UserWarning') && !line.includes('HF_TOKEN') && !line.includes('symlink')) {
          mainWindow?.webContents.send('whisper-model-progress', {
            modelId: whisperModel || 'small.en',
            line
          });
        }
      });

      proc.on('exit', (code) => {
        try { fs.unlinkSync(scriptPath); } catch(_) {}
        if (code === 0 || output.includes('MODEL_READY')) {
          startWhisperServer(whisperModel).then(() => {
            mainWindow?.webContents.send('whisper-setup-result', {
              success: true,
              message: 'Whisper is ready! Restarting AnchorCast...'
            });
            setTimeout(() => { app.relaunch(); app.exit(0); }, 1500);
          });
          resolve({ ok: true });
        } else {
          resolve({ ok: false, error: `Model download failed (code ${code})` });
        }
      });

      proc.on('error', e => {
        try { fs.unlinkSync(scriptPath); } catch(_) {}
        resolve({ ok: false, error: e.message });
      });
    });
  }

  // Fallback: no bundled Python — run setup_whisper.bat
  const setupBat = resolveRuntimeResource('setup_whisper.bat');
  if (!fs.existsSync(setupBat)) {
    return { ok: false, error: 'Python not found and setup_whisper.bat missing. Please reinstall AnchorCast.' };
  }

  const setupDir = path.dirname(setupBat);
  const successFlag = path.join(setupDir, 'whisper_setup_complete.flag');
  const restartFlag = path.join(setupDir, 'whisper_restart_pending.flag');
  try { fs.unlinkSync(successFlag); } catch(_) {}
  try { fs.unlinkSync(restartFlag); } catch(_) {}

  try {
    const proc = spawn('cmd.exe', ['/c', 'start', 'AnchorCast - Whisper Setup', '/wait', 'cmd.exe', '/c', setupBat], {
      detached: true, shell: false, windowsHide: false,
    });

    // Poll every 2 seconds for flag files written by setup_whisper.bat
    const pollInterval = setInterval(() => {
      if (fs.existsSync(successFlag)) {
        clearInterval(pollInterval);
        try { fs.unlinkSync(successFlag); } catch(_) {}
        // Setup succeeded — start whisper server then tell renderer to restart app
        startWhisperServer(whisperModel).then(() => {
          mainWindow?.webContents.send('whisper-setup-result', {
            success: true,
            message: 'Whisper is ready! Restarting AnchorCast...'
          });
          // Restart the app after 1.5s so user sees the message
          setTimeout(() => {
            app.relaunch();
            app.exit(0);
          }, 1500);
        });
      } else if (fs.existsSync(restartFlag)) {
        clearInterval(pollInterval);
        try { fs.unlinkSync(restartFlag); } catch(_) {}
        // VC++ was just installed — need Windows restart
        mainWindow?.webContents.send('whisper-setup-result', {
          success: false,
          needsWindowsRestart: true,
          message: 'Visual C++ was installed. Please restart Windows, then AnchorCast will complete setup automatically.'
        });
      }
    }, 2000);

    // Stop polling after 15 minutes regardless
    setTimeout(() => clearInterval(pollInterval), 15 * 60 * 1000);

    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

// On startup: check if a restart-pending flag exists and re-run setup automatically
(async () => {
  const setupBat = resolveRuntimeResource('setup_whisper.bat');
  const setupDir = path.dirname(setupBat);
  const restartFlag = path.join(setupDir, 'whisper_restart_pending.flag');
  if (fs.existsSync(restartFlag)) {
    try { fs.unlinkSync(restartFlag); } catch(_) {}
    console.log('[Whisper] Restart-pending flag found — re-running setup after VC++ install...');
    setTimeout(() => {
      spawn('cmd.exe', ['/c', 'start', 'AnchorCast - Whisper Setup', '/wait', 'cmd.exe', '/c', setupBat], {
        detached: true, shell: false, windowsHide: false,
      });
    }, 3000); // wait for app to fully load first
  }
})();


let whisperSource = 'local'; // 'deepgram' | 'local' | 'cloud'
ipcMain.handle('set-whisper-source', (_, src) => {
  whisperSource = ['deepgram','local','cloud'].includes(src) ? src : 'local';
  console.log(`[Transcribe] Source: ${whisperSource}`);
  return { source: whisperSource };
});

// ── PCM Transcription Engine ──────────────────────────────────────────────────
// Sliding window approach for local/offline users:
// - slightly longer flush improves Whisper context
// - overlap avoids cut-off words
// - silence gate reduces wasted chunks / hallucinated text
let pcmChunks       = [];
let pcmOverlap      = [];
let pcmSampleRate   = 16000;
let pcmTotalSamples = 0;
let pcmTranscribing = false;
let pcmQueue        = [];

// ── PCM pipeline tuning for small.en on CPU ──────────────────────────────────
// CHUNK SIZE: 4.5s gives Whisper complete sentences with enough context.
//   Too short (< 3s) → incomplete sentences → more errors and confusion.
//   Too long (> 6s)  → latency too high for live use.
// OVERLAP: 0.75s catches boundary words without over-suppressing new content.
//   Too much overlap (> 1s) → stitching drops valid new lines.
// RMS GATE: 180 = quiet room background noise threshold.
//   Raise to 300+ if the mic picks up PA system hum.
const PCM_FLUSH_SECONDS    = 4.5;
const PCM_OVERLAP_SECONDS  = 0.75;
const PCM_MIN_RMS          = 180;
const PCM_MIN_ACTIVE_RATIO = 0.012;
const PCM_MAX_QUEUE_CHUNKS = 3;

let lastTranscriptSent = '';
let lastTranscriptWords = [];


function pcmRms(int16) {
  if (!int16 || !int16.length) return 0;
  let sum = 0;
  for (let i = 0; i < int16.length; i += 1) {
    const v = int16[i];
    sum += v * v;
  }
  return Math.sqrt(sum / int16.length);
}

function pcmActiveRatio(int16, threshold = 450) {
  if (!int16 || !int16.length) return 0;
  let active = 0;
  for (let i = 0; i < int16.length; i += 1) {
    if (Math.abs(int16[i]) >= threshold) active += 1;
  }
  return active / int16.length;
}

function stitchTranscriptFragment(text) {
  const incoming = String(text || '').trim();
  if (!incoming) return '';
  const prevWords = lastTranscriptWords || [];
  const nextWords = incoming.split(/\s+/).filter(Boolean);

  // Find longest matching suffix of prev that matches prefix of next
  // Try both exact and normalised (strip punctuation) comparisons
  const norm = w => w.toLowerCase().replace(/[^a-z0-9]/g, '');
  let best = 0;
  const maxOverlap = Math.min(12, prevWords.length, nextWords.length);
  for (let n = maxOverlap; n >= 2; n -= 1) {
    const a = prevWords.slice(prevWords.length - n).map(norm).join(' ');
    const b = nextWords.slice(0, n).map(norm).join(' ');
    if (a === b && a.length > 3) { best = n; break; }
  }
  const stitchedWords = best ? nextWords.slice(best) : nextWords.slice();
  const stitched = stitchedWords.join(' ').trim() || incoming;
  lastTranscriptWords = incoming.split(/\s+/).filter(Boolean).slice(-24);
  return stitched;
}

ipcMain.handle('push-audio-pcm', async (_, pcmBuffer) => {
  if (!pcmBuffer) return;
  const int16 = new Int16Array(pcmBuffer);
  pcmChunks.push(int16);
  pcmTotalSamples += int16.length;

  const flushSamples = Math.floor(pcmSampleRate * PCM_FLUSH_SECONDS);
  if (pcmTotalSamples < flushSamples) return;

  const chunksToProcess = pcmChunks.splice(0);
  pcmTotalSamples = 0;

  const overlapSamples = Math.floor(pcmSampleRate * PCM_OVERLAP_SECONDS);
  const allChunks = pcmOverlap.length ? [pcmOverlap, ...chunksToProcess] : chunksToProcess;
  const totalLen = allChunks.reduce((s, c) => s + c.length, 0);
  const merged = new Int16Array(totalLen);
  let off = 0;
  for (const c of allChunks) { merged.set(c, off); off += c.length; }

  pcmOverlap = merged.slice(Math.max(0, merged.length - overlapSamples));

  // Silence/noise gate — improves offline robustness by avoiding junk chunks.
  if (pcmRms(merged) < PCM_MIN_RMS && pcmActiveRatio(merged) < PCM_MIN_ACTIVE_RATIO) {
    return;
  }

  pcmQueue.push(merged);
  if (pcmQueue.length > PCM_MAX_QUEUE_CHUNKS) {
    pcmQueue = [pcmQueue.slice(-PCM_MAX_QUEUE_CHUNKS).reduce((acc, cur) => {
      const out = new Int16Array(acc.length + cur.length);
      out.set(acc, 0);
      out.set(cur, acc.length);
      return out;
    })];
  }
  if (!pcmTranscribing) processNextChunk();
});

async function processNextChunk() {
  if (pcmQueue.length === 0) { pcmTranscribing = false; return; }

  pcmTranscribing = true;
  let merged;
  if (pcmQueue.length > 1) {
    const all = pcmQueue.splice(0);
    const totalLen = all.reduce((s, c) => s + c.length, 0);
    merged = new Int16Array(totalLen);
    let off = 0;
    for (const c of all) { merged.set(c, off); off += c.length; }
  } else {
    merged = pcmQueue.shift();
  }

  const wavBuffer = buildWav(merged, pcmSampleRate);

  try {
    let text = null;
    const useLocal  = whisperSource === 'local' && whisperReady;
    const useOnline = whisperSource === 'cloud' && currentSettings.openAiKey;

    if (useLocal) {
      try {
        const audioDurationMs = (merged.length / pcmSampleRate) * 1000;
        // Timeout = 3× audio duration + 8s buffer (small.en on CPU: ~5-8s per 4.5s chunk)
        const timeoutMs = Math.ceil(Math.max(25000, audioDurationMs * 4 + 8000));
        const res = await fetch(`${WHISPER_URL}/transcribe`, {
          method: 'POST',
          headers: { 'Content-Type': 'audio/wav', 'Content-Length': wavBuffer.length },
          body: wavBuffer,
          signal: AbortSignal.timeout(timeoutMs),
        });
        if (res.ok) {
          const data = await res.json();
          text = normalizeTranscriptText((data.text || '').trim());
          if (text) console.log(`[Whisper] "${text.slice(0, 80)}"`);
        }
      } catch (e) {
        if (e.name === 'TimeoutError') console.warn('[Whisper] Timeout — try base.en or tiny.en on slower CPUs');
        else console.warn('[Whisper] request failed:', e.message);
        if (e.code === 'ECONNREFUSED' || e.code === 'ECONNRESET') {
          whisperReady = false;
          mainWindow?.webContents.send('whisper-status', { ready: false });
        }
      }
    } else if (useOnline) {
      try {
        const form = new FormData();
        form.append('file', new Blob([wavBuffer], { type: 'audio/wav' }), 'audio.wav');
        form.append('model', 'whisper-1');
        form.append('language', 'en');
        form.append('response_format', 'text');
        const res = await fetch('https://api.openai.com/v1/audio/transcriptions', {
          method: 'POST',
          headers: { 'Authorization': `Bearer ${currentSettings.openAiKey}` },
          body: form,
          signal: AbortSignal.timeout(15000),
        });
        if (res.ok) {
          text = normalizeTranscriptText((await res.text()).trim());
          if (text) console.log(`[Whisper API] "${text.slice(0, 80)}"`);
        } else {
          const err = await res.json().catch(() => ({}));
          console.warn('[Whisper API] error', res.status, err?.error?.message || '');
        }
      } catch (e) {
        console.warn('[Whisper API] fetch error:', e.message);
      }
    } else {
      mainWindow?.webContents.send('transcript-no-key');
    }

    if (text && text.length > 1) {
      const stitched = stitchTranscriptFragment(text);
      if (stitched && shouldEmitTranscript(stitched)) {
        mainWindow?.webContents.send('transcript-result', stitched);
      }
    }
  } catch (e) {
    console.warn('[Transcribe] error:', e.message);
  }

  processNextChunk();
}


function normalizeTranscriptText(text) {
  let out = String(text || '').trim();
  if (!out) return '';

  const fixes = [
    [/\bpsalms\b/gi, 'Psalms'],
    [/\bpsalm\b/gi, 'Psalm'],
    [/\brevelations\b/gi, 'Revelation'],
    // ── Biblical word confusions that Whisper commonly mishears ──
    [/\bshall not watch\b/gi, 'shall not want'],
    [/\bshall not wash\b/gi, 'shall not want'],
    [/\bshall not won\b/gi,  'shall not want'],
    [/\bgreen pass\b/gi, 'green pastures'],
    [/\bgreen pace\b/gi, 'green pastures'],
    [/\bgreen pastor\b/gi, 'green pastures'],
    [/\bstill water(?!s)\b/gi, 'still waters'],
    [/\brivers of righteousness\b/gi, 'paths of righteousness'],
    [/\brestoreth my son\b/gi, 'restoreth my soul'],
    [/\brestored my son\b/gi,  'restoreth my soul'],
    [/\brestore my son\b/gi,   'restoreth my soul'],
    [/\bleadeth my son\b/gi,   'leadeth my soul'],
    [/\brestores my son\b/gi, 'restores my soul'],
    [/\brefresh my son\b/gi, 'refresh my soul'],
    [/\bthigh kingdom\b/gi, 'thy kingdom'],
    [/\bethernet\b/gi, 'eternal'],
    [/\bhollowly\b/gi, 'hallowed'],
    [/\bgoodness mercy\b/gi, 'goodness and mercy'],
    [/\brod staff\b/gi, 'rod and staff'],
    [/walk through the back/gi, 'walk through the valley'],
    [/through the back/gi, 'through the valley'],
    [/shadow death/gi, 'shadow of death'],
    [/I will fear and all evil/gi, 'I will fear no evil'],
    [/I will fear and evil/gi, 'I will fear no evil'],
    [/fear and all evil/gi, 'fear no evil'],
    [/my cup runny/gi, 'my cup runneth over'],
    [/my cup running/gi, 'my cup runneth over'],
    [/anoint my head/gi, 'anointest my head'],
    [/I shall not past/gi, 'I shall not want'],
    [/I shall not pass/gi, 'I shall not want'],
    [/walk through the Valle/gi, 'walk through the valley'],
    [/the Valle/gi, 'the valley'],
    [/hollow be thy name/gi, 'hallowed be thy name'],
    [/hallow be thy name/gi, 'hallowed be thy name'],
    [/our daily broad/gi, 'our daily bread'],
    [/our father which art/gi, 'our Father which art'],
  [/\bsabour\b/gi, 'saviour'],
    [/\bphilippian\b/gi, 'Philippians'],
    [/\bephesians?\b/gi, 'Ephesians'],
    [/\bfirst corinthians\b/gi, '1 Corinthians'],
    [/\bsecond corinthians\b/gi, '2 Corinthians'],
    [/\bfirst kings\b/gi, '1 Kings'],   [/\bsecond kings\b/gi, '2 Kings'],
    [/\bfirst samuel\b/gi, '1 Samuel'], [/\bsecond samuel\b/gi, '2 Samuel'],
    [/\bfirst chronicles\b/gi, '1 Chronicles'], [/\bsecond chronicles\b/gi, '2 Chronicles'],
    [/\bfirst peter\b/gi, '1 Peter'],   [/\bsecond peter\b/gi, '2 Peter'],
    [/\bfirst john\b/gi, '1 John'],     [/\bsecond john\b/gi, '2 John'],
    [/\bthird john\b/gi, '3 John'],
    [/\bfirst timothy\b/gi, '1 Timothy'], [/\bsecond timothy\b/gi, '2 Timothy'],
    [/\bfirst thessalonians\b/gi, '1 Thessalonians'], [/\bsecond thessalonians\b/gi, '2 Thessalonians'],
    [/\bsong of songs\b/gi, 'Song of Solomon'],

    // ── Biblical names Deepgram commonly mishears ──────────────────────────
    [/\bzakius\b/gi, 'Zacchaeus'],
    [/\bzakeus\b/gi, 'Zacchaeus'],
    [/\bzakey?us\b/gi, 'Zacchaeus'],
    [/\bzache?us\b/gi, 'Zacchaeus'],
    [/\bzacheus\b/gi, 'Zacchaeus'],
    [/\bbesali\b/gi, 'Bezalel'],
    [/\bdesally\b/gi, 'Bezalel'],
    [/\bbesaly\b/gi, 'Bezalel'],
    [/\bbe[sz]al[iy]el?\b/gi, 'Bezalel'],
    [/\bpostmortem\b/gi, 'Paul'],          // Deepgram mishears "Paul" as "postmortem"
    [/\bpolymer\b/gi, 'Paul'],             // "Paul" misheard as "polymer"

    // ── "Grace" misheard as other words (very common in fast African speech) ──
    [/\bdecrease of (god|law|christ)\b/gi, 'grace of $1'],
    [/\bincrease of gold\b/gi, 'grace of God'],
    [/\bgrade of god\b/gi, 'grace of God'],
    [/\braised? of god\b/gi, 'grace of God'],
    [/\bgraze of god\b/gi, 'grace of God'],

    // ── "God" misheard in fast speech ─────────────────────────────────────
    [/\bgrace of job\b/gi, 'grace of God'],
    [/\bgrace of gold\b/gi, 'grace of God'],
    [/\bgrace of go\b/gi, 'grace of God'],
    [/\bkingdom of go\b/gi, 'kingdom of God'],
    [/\bword of go\b/gi, 'word of God'],
    [/\bchildren of go\b/gi, 'children of God'],
    [/\bpeople of go\b/gi, 'people of God'],
    [/\bman of go\b/gi, 'man of God'],

    // ── "Found grace" vs profanity — context-aware ────────────────────────
    // Deepgram transcribes "Noah found grace" as "Noah fuck grace" in some accents
    [/\bnoah fuck grace\b/gi, 'Noah found grace'],
    [/\bfuck grace\b/gi, 'found grace'],       // biblical context only

    // ── Book name mishears from this sermon ──────────────────────────────
    // "Ephesians" heard as "in fifteenth" or "fifteen" by Deepgram
    [/\bin fifteenth chapter\b/gi, 'Ephesians chapter'],
    [/\bfifteenth chapter (\w+)\b/gi, 'Ephesians chapter $1'],
    // "1 Peter" heard as "Amos two" — harder to fix without context
    // Adding Amos cleanup so false detections are less likely
    [/\bamos two was talking about\b/gi, '1 Peter was talking about'],

    // ── Common sermon speech patterns that cause false positives ─────────
    [/\bwhat pussy\b/gi, 'what passage'],      // "what passage?" misheard
    [/\bthe antiseper\b/gi, 'the answer is'],
    [/\bchris open\b/gi, 'Christ upon'],
    [/\bit was biggie\b/gi, 'it was big'],
    [/\bpostmortem was talking\b/gi, 'Paul was talking'],

    // ── "Grace" core word reinforcement ──────────────────────────────────
    [/\bthe grace of job\b/gi, 'the grace of God'],
    [/\bgod is good at all time(?!s)\b/gi, 'God is good at all times'],

    // ── Sermon 2 corrections (Apr 19, 2026) ──────────────────────────────
    [/\bsad vision\b/gi, 'salvation'],
    [/\bfor sad vision\b/gi, 'for salvation'],
    [/\baffixure\b/gi, 'apostle'],
    [/\ban affixure\b/gi, 'an apostle'],
    [/\bby ?secuting\b/gi, 'persecuting'],
    [/\bcycling ship\b/gi, 'keeping sheep'],
    [/\bkick ?ship\b/gi, 'kingship'],
    [/\bmess[iy] says\b/gi, 'mercy says'],
    [/\bunsel?fly\b/gi, 'unselfishly'],
    [/\bunmerited female\b/gi, 'unmerited favor'],
    [/\bscriptatory\b/gi, 'scriptures'],
    [/\bgrass of god\b/gi, 'grace of God'],
    [/\bgrace for sad\b/gi, 'grace for salvation'],
    [/\bsad addition\b/gi, 'salvation'],

    // ── Sermon 3 corrections (Apr 12, 2026) ─ FOCUS sermon ──────────────
    // CRITICAL: 'focus' → 'fuck us' in African English accent
    [/\bfuck us\b/gi, 'focus'],
    [/\bif you.?re watching,? fuck us\b/gi, 'if you are watching, focus'],
    [/\bfuck us on\b/gi, 'focus on'],
    [/\bfuck us\.\b/gi, 'focus.'],
    // 1 Kings misheard
    [/\bfirst kick from the (\d+)/gi, '1 Kings chapter $1'],
    [/\bfirst kick from the\b/gi, '1 Kings'],
    [/\bfourth kings?\b/gi, '1 Kings'],
    [/\bfirst kick (\d+)/gi, '1 Kings $1'],
    [/\bfirst take (\d+)\b/gi, '1 Kings $1'],
    // Ephesians variant
    [/\bephysians?\b/gi, 'Ephesians'],
    // Names
    [/\bzolom\b/gi, 'Solomon'],


    // ── Sermon 6 corrections (Apr 26, 2026) ─ Blessing of Work (Deepgram) ──
    [/\blocust placed man\b/gi,              'Lord placed man'],
    [/\bhard water has plenty\b/gi,          'hard worker has plenty'],
    [/\bhard water get rich\b/gi,            'hard worker gets rich'],
    [/\bdelacing and become a slave\b/gi,    'be lazy and become a slave'],
    [/\bmay not live to poverty\b/gi,        'mere talk leads to poverty'],
    [/\blazy people work much\b/gi,          'lazy people want much'],
    [/\bpestilence\.?\s+focus on\b/gi,    'perseverance. Focus on'],
    [/\bfoursight\b/gi,                      'foresight'],
    [/\btitty\b/gi,                          'duty'],
    [/\bboiling the midnight oil\b/gi,       'burning the midnight oil'],
    [/\bthe copier\b/gi,                     'the cupbearer'],

    // ── Sermon 5 corrections (Apr 26, 2026) ─ Blessing of Work ──────────
    [/\bprocter\b/gi, 'Proverbs'],
    [/\bproctor\b/gi, 'Proverbs'],
    [/\btheir lungs\b/gi, 'their lamps'],
    [/\bour lungs\b/gi, 'our lamps'],
    [/\bthe lungs\b/gi, 'the lamps'],
    [/\blungs are going\b/gi, 'lamps are going out'],
    [/\bhaving suicide\b/gi, 'having foresight'],
    [/\bfour sides\b/gi, 'foresight'],
    [/\bsuicide is the ability\b/gi, 'foresight is the ability'],
    [/\blearning from the aunts\b/gi, 'learning from the ants'],
    [/\bthe aunts today\b/gi, 'the ants today'],
    [/\bthe aunt this morning\b/gi, 'the ant this morning'],
    [/\batonement is very weak\b/gi, 'your amen is very weak'],
    [/\bit'?s delicious!\b/gi, 'diligence!'],
    [/\bpropitiation for yourself\b/gi, 'provision for yourself'],
    [/\bpropitiation for others\b/gi, 'provision for others'],

    // Mishears
    [/\bfraud stealing in god\b/gi, 'trusting in God'],
    [/\bfraud stealing\b/gi, 'trusting'],
    [/\bwhy your son was pissed\b/gi, 'why your son was busy'],
    [/\bjack of poultry\b/gi, 'jack of all trades'],
    [/\bmaster of no\b/gi, 'master of none'],
    // ── Sermon 4 corrections (Apr 15, 2026) ─ NEW LIFE IN CHRIST ────────
    // Book name mishears
    [/\bgalicia'?s chapter\b/gi,           'Galatians chapter'],
    [/\bgalicias\b/gi,                     'Galatians'],
    [/\bcollisions chapter\b/gi,           'Colossians chapter'],
    [/\bcollisions\b/gi,                   'Colossians'],
    [/\bconverseius chapter\b/gi,          'Colossians chapter'],
    [/\bfishes number (\d+)/gi,            'Ephesians chapter $1'],
    [/\bfishes number\b/gi,               'Ephesians'],
    [/\bnew pipe a robot slab\b/gi,        'now Romans'],
    // 'heirs' misheard as 'ears' — critical KJV word
    [/\bjoint ears\b/gi,                   'joint heirs'],
    [/\bears of god\b/gi,                  'heirs of God'],
    [/\bears and heirs\b/gi,               'heirs of God'],
    [/\bco ears\b/gi,                      'co-heirs'],
    [/\bwe are afraid of god\b/gi,         'we are heirs of God'],
    [/\bwe are ears\b/gi,                  'we are heirs'],
    // 'firstborn' → 'fourth born'
    [/\bthe fourth born of all creation\b/gi, 'the firstborn of all creation'],
    [/\bfourth born\b/gi,                  'firstborn'],
    // 'peculiar people' → 'pebillion people'
    [/\bpebillion people\b/gi,             'peculiar people'],
    // Other mishears
    [/\bseed round before god\b/gi,        'filthy rags before God'],
    [/\bsave conscious\b/gi,               'righteousness conscious'],
    [/\bright ?eous conscious\b/gi,        'righteousness conscious'],
    [/\bkisos\b/gi,                        'Jesus'],
    [/\bright side of him\b/gi,            'right hand of the Father'],
    [/\bnew payment\b/gi,                  'new man'],
    [/\bnew ports\b/gi,                    'new man'],
  ];
  for (const [pattern, replacement] of fixes) out = out.replace(pattern, replacement);

  // ── Verbal verse reference conversion: "James one twelve" → "James 1:12" ──
  // Applies BEFORE dedup so the ref survives
  out = _convertVerbalRefs(out);

  // ── Repetition removal — 4 passes ────────────────────────────────────────
  // Pass 0: Sentence-level loop collapse (Deepgram streaming glitch)
  // When audio drops, Deepgram re-transcribes the last sentence 5-15 times.
  // Detect: same sentence (8+ words) repeating 3+ times consecutively → keep once
  out = out.replace(/([^.!?]{30,}[.!?,]?)(?:\s+\1){2,}/gi, '$1');

  // Pass 1: 2+ word phrase repeats with optional punctuation between
  out = out.replace(/\b(.{4,80}?)(?:[,.\s]+\1)+\b/gi, '$1');
  // Pass 2: exact word-pair repeats "Jesus Christ Jesus Christ"
  out = out.replace(/\b(\w+\s+\w+)[,. ]+\1\b/gi, '$1');
  // Pass 3: single word repeats "Amen Amen"
  out = out.replace(/\b(\w{3,})[,. ]+\1\b/gi, '$1');

  out = out.replace(/\s+/g, ' ').trim();
  return out;
}

// ── Verbal Bible reference converter ─────────────────────────────────────────
// "James one twelve" → "James 1:12"
// "Psalm twenty three one" → "Psalm 23:1"
// "John three sixteen" → "John 3:16"
const _NUM_WORDS = {
  one:1,two:2,three:3,four:4,five:5,six:6,seven:7,eight:8,nine:9,ten:10,
  eleven:11,twelve:12,thirteen:13,fourteen:14,fifteen:15,sixteen:16,
  seventeen:17,eighteen:18,nineteen:19,twenty:20,thirty:30,forty:40,
  fifty:50,sixty:60,seventy:70,eighty:80,ninety:90,hundred:100,
  first:1,second:2,third:3,fourth:4,fifth:5,sixth:6,seventh:7,eighth:8,ninth:9,tenth:10,
  eleventh:11,twelfth:12,thirteenth:13,fourteenth:14,fifteenth:15,
  sixteenth:16,seventeenth:17,eighteenth:18,nineteenth:19,twentieth:20,
};
const _TENS = ['twenty','thirty','forty','fifty','sixty','seventy','eighty','ninety'];
const _NUM_PAT = `(?:(?:${_TENS.join('|')})\\s+)?(?:${Object.keys(_NUM_WORDS).join('|')}|\\d+)`;

function _parseWordNum(str) {
  if (!str) return null;
  const s = str.toLowerCase().trim();
  const n = parseInt(s);
  if (!isNaN(n)) return n;
  if (_NUM_WORDS[s] !== undefined) return _NUM_WORDS[s];
  const parts = s.split(/\s+/);
  if (parts.length === 2 && _NUM_WORDS[parts[0]] && _NUM_WORDS[parts[1]]) {
    return _NUM_WORDS[parts[0]] + _NUM_WORDS[parts[1]];
  }
  return null;
}

const _BIBLE_BOOKS_FOR_VERBAL = [
  'Song of Solomon',
  '1 Corinthians','2 Corinthians','1 Samuel','2 Samuel','1 Kings','2 Kings',
  '1 Chronicles','2 Chronicles','1 Peter','2 Peter','1 John','2 John','3 John',
  '1 Timothy','2 Timothy','1 Thessalonians','2 Thessalonians',
  'Genesis','Exodus','Leviticus','Numbers','Deuteronomy','Joshua','Judges','Ruth',
  'Ezra','Nehemiah','Esther','Job','Psalms','Psalm','Proverbs','Ecclesiastes',
  'Isaiah','Jeremiah','Lamentations','Ezekiel','Daniel','Hosea','Joel','Amos',
  'Obadiah','Jonah','Micah','Nahum','Habakkuk','Zephaniah','Haggai','Zechariah',
  'Malachi','Matthew','Mark','Luke','John','Acts','Romans','Galatians','Ephesians',
  'Philippians','Colossians','Titus','Philemon','Hebrews','James','Jude','Revelation',
];

function _convertVerbalRefs(text) {
  let out = text;
  for (const book of _BIBLE_BOOKS_FOR_VERBAL) {
    const escaped = book.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(
      `\\b(${escaped})\\s+(${_NUM_PAT})\\s+(${_NUM_PAT})\\b`,
      'gi'
    );
    out = out.replace(re, (match, bk, chStr, vsStr) => {
      const ch = _parseWordNum(chStr.trim());
      const vs = _parseWordNum(vsStr.trim());
      if (ch && vs && ch <= 150 && vs <= 176) return `${bk} ${ch}:${vs}`;
      return match;
    });
  }
  return out;
}
function shouldEmitTranscript(text) {
  if (!text) return false;
  const norm = String(text).toLowerCase().trim();
  const prev = String(lastTranscriptSent || '').toLowerCase().trim();
  if (!norm || norm.length < 5) return false;

  // Exact duplicate — always suppress
  if (norm === prev) return false;

  // Pure superset/subset — only emit if it adds ≥5 new chars of content
  if (prev && (norm.startsWith(prev) || prev.endsWith(norm))) {
    lastTranscriptSent = text;
    return norm.length > prev.length + 5;
  }

  // Word-overlap check — suppress only very high overlap AND not much longer.
  // IMPORTANT: sermon text naturally shares words ("the", "Lord", "and") across lines.
  // We use CONTENT word overlap (skip stop words) to avoid false suppression.
  if (prev.length > 10 && norm.length > 10) {
    const STOPS = new Set(['the','a','an','and','or','but','in','of','to','for',
      'is','are','was','were','be','been','being','that','this','with','he','she',
      'it','we','you','they','his','her','its','my','your','our','their','i','me']);
    const prevW = new Set(prev.split(/\s+/).filter(w => !STOPS.has(w)));
    const normW = norm.split(/\s+/).filter(w => !STOPS.has(w));
    if (prevW.size > 2 && normW.length > 2) {
      const overlap = normW.filter(w => prevW.has(w)).length;
      const similarity = overlap / Math.max(prevW.size, normW.length);
      // Only suppress if >92% content-word overlap AND very similar length
      if (similarity > 0.92 && Math.abs(norm.length - prev.length) < 8) {
        return false;
      }
    }
  }

  lastTranscriptSent = text;
  return true;
}

// Build a minimal WAV file from Int16 PCM data
function buildWav(int16Samples, sampleRate) {
  const numChannels  = 1;
  const bitsPerSample = 16;
  const byteRate     = sampleRate * numChannels * bitsPerSample / 8;
  const blockAlign   = numChannels * bitsPerSample / 8;
  const dataSize     = int16Samples.length * 2;
  const headerSize   = 44;
  const buf          = Buffer.alloc(headerSize + dataSize);

  // RIFF header
  buf.write('RIFF', 0);
  buf.writeUInt32LE(36 + dataSize, 4);
  buf.write('WAVE', 8);
  // fmt chunk
  buf.write('fmt ', 12);
  buf.writeUInt32LE(16, 16);         // chunk size
  buf.writeUInt16LE(1, 20);          // PCM format
  buf.writeUInt16LE(numChannels, 22);
  buf.writeUInt32LE(sampleRate, 24);
  buf.writeUInt32LE(byteRate, 28);
  buf.writeUInt16LE(blockAlign, 32);
  buf.writeUInt16LE(bitsPerSample, 34);
  // data chunk
  buf.write('data', 36);
  buf.writeUInt32LE(dataSize, 40);
  // PCM samples (little-endian Int16)
  for (let i = 0; i < int16Samples.length; i++) {
    buf.writeInt16LE(int16Samples[i], headerSize + i * 2);
  }
  return buf;
}




function _safeJsonRead(file, fallback){
  try { if (fs.existsSync(file)) return JSON.parse(fs.readFileSync(file, 'utf8')); } catch(_) {}
  return fallback;
}
function _normalizeSongForIo(song, index = 0){
  const safeId = song?.id ?? (Date.now().toString(36) + '_' + Math.random().toString(36).slice(2,8) + '_' + index);
  const sections = Array.isArray(song?.sections) ? song.sections : [];
  return {
    id: safeId,
    title: String(song?.title || 'New Song').trim() || 'New Song',
    author: String(song?.author || '').trim(),
    key: String(song?.key || '').trim(),
    sections: sections.map((sec, i) => ({
      label: String(sec?.label || `Verse ${i+1}`).trim() || `Verse ${i+1}`,
      lines: Array.isArray(sec?.lines) ? sec.lines.map(v => String(v || '').trim()).filter(Boolean) : [],
      richHtml: typeof sec?.richHtml === 'string' ? sec.richHtml : '',
      textTransform: ['inherit','none','uppercase','lowercase','capitalize'].includes(sec?.textTransform) ? sec.textTransform : 'inherit',
    })).filter(sec => sec.lines.length || sec.richHtml)
  };
}
function _stripRtf(rtf = ''){
  let text = String(rtf || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  text = text
    .replace(/\{\\fonttbl[\s\S]*?\}/g, '')
    .replace(/\{\\colortbl[\s\S]*?\}/g, '')
    .replace(/\{\\stylesheet[\s\S]*?\}/g, '')
    .replace(/\{\\info[\s\S]*?\}/g, '')
    .replace(/\{\\\*[^{}]*\}/g, '');

  text = text
    .replace(/\\par[d]?/g, '\n')
    .replace(/\\line/g, '\n')
    .replace(/\\tab/g, '\t')
    .replace(/\\emdash/g, '—')
    .replace(/\\endash/g, '–')
    .replace(/\\bullet/g, '•')
    .replace(/\\u-?\d+\??/g, '')
    .replace(/\\'[0-9a-fA-F]{2}/g, '')
    .replace(/\\[a-zA-Z]+-?\d* ?/g, '')
    .replace(/[{}]/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .trim();

  const lines = text.split(/\n+/).map(v => v.trim()).filter(Boolean);
  const cleaned = lines.filter(line => !/^(Arial|Verdana|Tahoma|Times New Roman|Cambria|CorisandeRegular|Regular|Bold|Italic)([;,\s].*)?$/i.test(line));
  return cleaned.join('\n').trim();
}

function _parseSongTextToSections(raw = ''){
  const lines = String(raw || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const sections = [];
  let current = null;
  const isLabel = (line) => {
    const t = String(line || '').trim();
    if (!t) return false;
    if (/^\[.+\]$/.test(t)) return t.slice(1, -1).trim();
    if (/^\{.+\}$/.test(t)) return t.slice(1, -1).trim();
    if (/^(verse|chorus|bridge|pre[\s-]?chorus|outro|tag|interlude|hook|refrain|vamp)\s*\d*[:\s]?$/i.test(t)) return t.replace(/:\s*$/, '').trim();
    return false;
  };
  for (const line of lines) {
    const label = isLabel(line);
    if (label) {
      if (current && current.lines.length) sections.push(current);
      current = { label, lines: [] };
      continue;
    }
    const cleaned = String(line || '').replace(/\{[^{}]*\}/g, '').trim();
    if (!cleaned) {
      if (current && current.lines.length) { sections.push(current); current = null; }
      continue;
    }
    if (!current) current = { label: `Verse ${sections.length + 1}`, lines: [] };
    current.lines.push(cleaned);
  }
  if (current && current.lines.length) sections.push(current);
  return sections.length ? sections : [{ label: 'Verse 1', lines: lines.map(v => String(v || '').trim()).filter(Boolean) }];
}
function _cleanupImportedLyricText(text = '') {
  let out = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = out.split('\n').map(v => v.trim()).filter(Boolean);
  const filtered = lines.filter(line => {
    if (/^(Arial|Verdana|Tahoma|Times New Roman|Cambria|CorisandeRegular)([;,\s].*)?$/i.test(line)) return false;
    if (/^(red\d+|green\d+|blue\d+|fcharset\d+|deff\d+|deftab\d+)/i.test(line)) return false;
    return true;
  });
  return filtered.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

function _parseEasyWorshipSongBuffer(filePath, text){
  const ext = path.extname(String(filePath || '')).toLowerCase();
  const raw = ext === '.rtf' ? _stripRtf(text) : String(text || '');
  const normalized = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const chunks = normalized.split(/\n{2,}/).map(v => v.trim()).filter(Boolean);
  let title = path.basename(String(filePath || ''), ext);
  let author = '';
  let bodyStartIdx = 0;
  if (chunks[0] && /^title\s*:/i.test(chunks[0])) {
    title = chunks[0].replace(/^title\s*:/i, '').trim() || title;
    bodyStartIdx = 1;
  } else if (chunks[0] && chunks[0].split('\n').length === 1 && chunks.length > 1) {
    title = chunks[0].trim() || title;
    bodyStartIdx = 1;
  }
  for (let i = bodyStartIdx; i < Math.min(chunks.length, bodyStartIdx + 2); i += 1) {
    if (/^(author|artist|written by)\s*:/i.test(chunks[i])) {
      author = chunks[i].replace(/^(author|artist|written by)\s*:/i, '').trim();
      bodyStartIdx = i + 1;
      break;
    }
  }
  const body = chunks.slice(bodyStartIdx).join('\n\n').trim();
  const sections = _parseSongTextToSections(body);
  if (!title && sections[0]?.lines?.[0]) title = sections[0].lines[0].slice(0, 80);
  return _normalizeSongForIo({ title, author, key:'', sections });
}
function _serializeSongAsEasyWorshipText(song){
  const normalized = _normalizeSongForIo(song);
  const blocks = [];
  blocks.push(`Title: ${normalized.title}`);
  if (normalized.author) blocks.push(`Author: ${normalized.author}`);
  blocks.push('');
  for (const sec of normalized.sections) {
    if (sec.label) blocks.push(`[${sec.label}]`);
    blocks.push(...sec.lines);
    blocks.push('');
  }
  return blocks.join('\r\n').trim() + '\r\n';
}
function _dedupeSongs(existing = [], incoming = []){
  const seen = new Set(existing.map(s => `${String(s.title||'').trim().toLowerCase()}|${String(s.author||'').trim().toLowerCase()}`));
  const out = existing.slice();
  let added = 0;
  for (const song of incoming) {
    const normalized = _normalizeSongForIo(song);
    const key = `${normalized.title.trim().toLowerCase()}|${normalized.author.trim().toLowerCase()}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(normalized);
    added += 1;
  }
  return { songs: out, added };
}


function _songLibraryPythonCandidates() {
  const candidates = [
    (typeof resolveRuntimeResource === 'function' ? resolveRuntimeResource('python', 'python.exe') : null),
    (typeof resolveRuntimeResource === 'function' ? resolveRuntimeResource('python', 'bin', 'python3') : null),
    'py',
    'python',
    'python3',
  ].filter(Boolean);
  return Array.from(new Set(candidates));
}

function _readSongLibrary() {
  return _safeJsonRead(SONGS_FILE, []);
}

function _writeSongLibrary(songs, notify = true) {
  fs.mkdirSync(path.dirname(SONGS_FILE), { recursive: true });
  fs.writeFileSync(SONGS_FILE, JSON.stringify(songs, null, 2));
  if (notify) {
    try { mainWindow?.webContents.send('songs-saved'); } catch (_) {}
    try { songManagerWindow?.webContents.send('songs-saved'); } catch (_) {}
    try { settingsWindow?.webContents.send('songs-saved'); } catch (_) {}
  }
  return songs;
}

function _makeSongLibraryBackup(reason = 'manual') {
  const songs = _readSongLibrary().map((song, i) => _normalizeSongForIo(song, i));
  fs.mkdirSync(SONG_BACKUP_DIR, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const filePath = path.join(SONG_BACKUP_DIR, `song-library-${reason}-${stamp}.json`);
  fs.writeFileSync(filePath, JSON.stringify({
    format: 'anchorcast-song-library-backup',
    reason,
    createdAt: Date.now(),
    count: songs.length,
    songs,
  }, null, 2));
  return { path: filePath, count: songs.length };
}

function _maybeRunSundaySongBackup() {
  try {
    const now = new Date();
    const isSunday = now.getDay() === 0;
    if (!isSunday) return;
    const settings = loadSettings();
    if (settings.songAutoBackupSunday === false) return;
    const today = now.toISOString().slice(0, 10);
    const meta = _safeJsonRead(SONG_BACKUP_META, {});
    if (meta.lastSundayBackupDate === today) return;
    const backup = _makeSongLibraryBackup('auto-sunday');
    fs.writeFileSync(SONG_BACKUP_META, JSON.stringify({ lastSundayBackupDate: today, lastBackupPath: backup.path }, null, 2));
    console.log('[Songs] Auto Sunday backup created:', backup.path);
  } catch (e) {
    console.warn('[Songs] Auto Sunday backup failed:', e.message);
  }
}

function _extractSongsFromSqliteDb(filePath) {
  const pyScript = `
import sqlite3, json, sys, re

db_path = sys.argv[1]

def strip_markup(text):
    text = str(text or '')
    text = text.replace('\\r\\n','\\n').replace('\\r','\\n')
    text = re.sub(r'<[^>]+>', '', text)
    text = re.sub(r'\\{[^{}]*\\}', '', text)
    return text.strip()

def parse_sections(raw):
    lines = strip_markup(raw).split('\\n')
    sections = []
    current = None
    def label_of(line):
        t = (line or '').strip()
        if not t:
            return None
        if re.match(r'^\\[.+\\]$', t):
            return t[1:-1].strip()
        if re.match(r'^(verse|chorus|bridge|pre[ -]?chorus|outro|tag|interlude|hook|refrain|vamp)\\s*\\d*[:\\s]?$', t, re.I):
            return re.sub(r':\\s*$', '', t).strip()
        return None
    for line in lines:
        lab = label_of(line)
        if lab:
            if current and current['lines']:
                sections.append(current)
            current = {'label': lab, 'lines': []}
            continue
        clean = line.strip()
        if not clean:
            if current and current['lines']:
                sections.append(current)
                current = None
            continue
        if current is None:
            current = {'label': f'Verse {len(sections)+1}', 'lines': []}
        current['lines'].append(clean)
    if current and current['lines']:
        sections.append(current)
    if not sections:
        body_lines = [ln.strip() for ln in lines if ln.strip()]
        if body_lines:
            sections = [{'label':'Verse 1','lines':body_lines}]
    return sections

def norm_song(title, author, lyrics, key=''):
    title = str(title or '').strip() or 'Imported Song'
    author = str(author or '').strip()
    sections = parse_sections(lyrics)
    return {'title': title, 'author': author, 'key': str(key or '').strip(), 'sections': sections}

try:
    conn = sqlite3.connect(db_path)
except Exception as e:
    print(json.dumps({'error': f'open_failed:{e}'}))
    raise SystemExit(0)

try:
    conn.row_factory = sqlite3.Row
    tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'").fetchall()]
    songs = []
    seen = set()

    title_keys = {'title','name','songtitle','song_name','caption'}
    author_keys = {'author','artist','writer','writtenby','composer'}
    lyrics_keys = {'lyrics','text','content','body','words','slide_text','verse','songtext'}

    for table in tables:
        try:
            cols = [r['name'] if isinstance(r, sqlite3.Row) else r[1] for r in conn.execute(f'PRAGMA table_info("{table}")').fetchall()]
        except Exception:
            continue
        low = {c.lower(): c for c in cols}
        title_col = next((low[k] for k in title_keys if k in low), None)
        lyrics_col = next((low[k] for k in lyrics_keys if k in low), None)
        author_col = next((low[k] for k in author_keys if k in low), None)
        key_col = low.get('key') or low.get('songkey')
        if not title_col or not lyrics_col:
            continue
        try:
            query_cols = [title_col, lyrics_col] + ([author_col] if author_col else []) + ([key_col] if key_col else [])
            query = 'SELECT ' + ','.join([f'"{c}"' for c in query_cols]) + f' FROM "{table}"'
            for row in conn.execute(query).fetchall():
                title = row[title_col]
                lyrics = row[lyrics_col]
                if not title or not lyrics:
                    continue
                author = row[author_col] if author_col else ''
                key = row[key_col] if key_col else ''
                item = norm_song(title, author, lyrics, key)
                dedupe = (item['title'].strip().lower(), item['author'].strip().lower())
                if dedupe in seen:
                    continue
                seen.add(dedupe)
                songs.append(item)
        except Exception:
            continue
    print(json.dumps({'songs': songs}))
except Exception as e:
    print(json.dumps({'error': f'query_failed:{e}'}))
`.trim();

  const { spawnSync } = require('child_process');
  for (const candidate of _songLibraryPythonCandidates()) {
    try {
      const args = candidate === 'py' ? ['-3', '-c', pyScript, filePath] : ['-c', pyScript, filePath];
      const res = spawnSync(candidate, args, {
        encoding: 'utf8',
        windowsHide: true,
        timeout: 30000,
        maxBuffer: 20 * 1024 * 1024,
      });
      const out = String(res.stdout || '').trim();
      if (!out) continue;
      const parsed = JSON.parse(out);
      if (Array.isArray(parsed?.songs)) return parsed.songs;
      if (parsed?.error?.startsWith('open_failed:')) {
        throw new Error('This file is not a readable SQLite song database.');
      }
    } catch (e) {
      if (e?.message) throw e;
    }
  }
  throw new Error('Could not read songs from database.');
}



function _extractEasyWorshipTitlesFromDb(dbPath) {
  const raw = fs.readFileSync(dbPath);
  const asciiMatches = raw.toString('latin1').match(/[\x20-\x7E]{4,}/g) || [];
  const blacklist = new Set([
    'kresttemp.DB','Title','Author','RecID','Text Percentage Bottom','Copyright','Administrator','Words','Default Background',
    'BK Type','BK Color','BK Gradient Color1','BK Gradient Color2','BK Gradient Shading','BK Gradient Variant','BK Texture',
    'BK Bitmap Name','BK Bitmap','BK AspectRatio','Favorite','Last Modified','Demo Data','Song Number','BK Thumbnail',
    'Override Enabled','Font Size Limit Default','Font Size Limit','Font Name Default','Font Name','Text Color Default','Text Color',
    'Shadow Color Default','Shadow Color','Outline Color Default','Outline Color','Shadow Text','Outline Text','Bold Text',
    'Italic Text','Text Alignment','Vert Alignment','Text Percent Rect Default','Text Percentage Left','Text Percentage Top',
    'Text Percentage Right','Vendor ID','ascii'
  ]);
  const titles = [];
  for (let i = 1; i < asciiMatches.length; i += 1) {
    const token = String(asciiMatches[i] || '').trim();
    if (!token.startsWith('{\\rtf1')) continue;
    const prev = String(asciiMatches[i - 1] || '').trim();
    if (!prev || prev.startsWith('{\\rtf1')) continue;
    if (blacklist.has(prev)) continue;
    if (/^[A-Za-z]:\\/.test(prev) || /Documents and Settings|Program Files|\\Users\\|\.jpg$|\.png$|\.bmp$|\.wmv$|\.mp4$/i.test(prev)) continue;
    if (/^d?public domain$/i.test(prev)) continue;
    if (/^[\W_]+$/.test(prev)) continue;
    if (prev.length > 220) continue;
    titles.push(prev);
  }
  return titles;
}

function _extractFullRtfBlocksFromFile(filePath) {
  const raw = fs.readFileSync(filePath);
  const marker = Buffer.from('{\\rtf1', 'latin1');
  const blocks = [];
  let pos = 0;
  while (true) {
    const start = raw.indexOf(marker, pos);
    if (start === -1) break;
    let i = start;
    let depth = 0;
    let escaped = false;
    for (; i < raw.length; i += 1) {
      const b = raw[i];
      if (escaped) {
        escaped = false;
        continue;
      }
      if (b === 0x5C) { // backslash
        escaped = true;
        continue;
      }
      if (b === 0x7B) depth += 1; // {
      else if (b === 0x7D) { // }
        depth -= 1;
        if (depth === 0) {
          blocks.push(raw.slice(start, i + 1).toString('latin1'));
          pos = i + 1;
          break;
        }
      }
    }
    if (i >= raw.length) break;
  }
  return blocks;
}

function _rtfToPlainText(rtf = '') {
  const stack = [{ ignorable: false, justOpened: false }];
  let out = '';
  let i = 0;
  const ignoreDestinations = new Set(['fonttbl','colortbl','stylesheet','info','pict','object','header','footer']);
  const current = () => stack[stack.length - 1];

  while (i < rtf.length) {
    const ch = rtf[i];
    if (ch === '{') {
      stack.push({ ignorable: current().ignorable, justOpened: true });
      i += 1;
      continue;
    }
    if (ch === '}') {
      if (stack.length > 1) stack.pop();
      i += 1;
      continue;
    }
    if (ch === '\\') {
      const next = rtf[i + 1] || '';
      if (next === '\\' || next === '{' || next === '}') {
        if (!current().ignorable) out += next;
        i += 2;
        current().justOpened = false;
        continue;
      }
      if (next === '*') {
        current().ignorable = true;
        i += 2;
        continue;
      }
      let j = i + 1;
      let word = '';
      while (j < rtf.length && /[A-Za-z]/.test(rtf[j])) {
        word += rtf[j];
        j += 1;
      }
      let num = '';
      if (rtf[j] === '-' || /\d/.test(rtf[j] || '')) {
        num += rtf[j];
        j += 1;
        }
      while (j < rtf.length && /\d/.test(rtf[j])) {
        num += rtf[j];
        j += 1;
      }
      if (rtf[j] === ' ') j += 1;
      if (current().justOpened && ignoreDestinations.has(word)) current().ignorable = true;
      current().justOpened = false;
      if (!current().ignorable) {
        if (word === 'par' || word === 'line') out += '\n';
        else if (word === 'tab') out += '\t';
        else if (word === 'u') {
          const code = parseInt(num || '0', 10);
          if (!Number.isNaN(code) && code > 0) out += String.fromCharCode(code);
        }
      }
      i = j;
      continue;
    }
    current().justOpened = false;
    if (!current().ignorable) out += ch;
    i += 1;
  }

  return out
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\u0000/g, '')
    .split('\n')
    .map(v => v.trim())
    .filter(v => v && !/^(Arial|Verdana|Tahoma|Times New Roman|Cambria|Calibri|Comic Sans MS|LucidaConsole|Monotype Corsiva|Agency FB|Aharoni|Broadway|Century Schoolbook|Baskerville Old Face|Calisto MT|Haettenschweiler)([;,\s].*)?$/i.test(v))
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function _extractEasyWorshipLyricsFromMb(mbPath) {
  const blocks = _extractFullRtfBlocksFromFile(mbPath);
  return blocks
    .map(block => _cleanupImportedLyricText(_rtfToPlainText(block)))
    .filter(Boolean);
}

function _extractEasyWorshipSongsFromDataFolder(dataDir) {
  const names = fs.readdirSync(dataDir);
  const songsDb = names.find(name => /^songs\.db$/i.test(name)) || names.find(name => /^songs\./i.test(name) && name.toLowerCase().endsWith('.db'));
  if (!songsDb) throw new Error('Songs.DB was not found in the selected folder.');
  const songsDbPath = path.join(dataDir, songsDb);
  const songsMb = names.find(name => /^songs\.mb$/i.test(name));
  const songsMbPath = songsMb ? path.join(dataDir, songsMb) : null;

  const titles = _extractEasyWorshipTitlesFromDb(songsDbPath);
  const lyrics = songsMbPath ? _extractEasyWorshipLyricsFromMb(songsMbPath) : [];
  if (!titles.length) throw new Error('No song titles were found in Songs.DB.');
  if (!lyrics.length) throw new Error('Songs.MB was not found or no readable lyrics were found. Please select the full DATA folder, not Songs.DB alone.');

  const count = Math.min(titles.length, lyrics.length);
  const songs = [];
  for (let i = 0; i < count; i += 1) {
    const title = String(titles[i] || '').trim();
    const plain = String(lyrics[i] || '').trim();
    if (!title || !plain) continue;
    const sections = _parseSongTextToSections(plain);
    if (!sections?.length || sections.every(sec => !(sec.lines || []).length)) continue;
    songs.push(_normalizeSongForIo({
      title,
      author: '',
      sections,
    }));
  }

  const deduped = [];
  const seen = new Set();
  for (const song of songs) {
    const key = `${String(song.title || '').trim().toLowerCase()}|${song.sections.map(sec => (sec.lines || []).join(' ')).join(' ').slice(0, 200).toLowerCase()}`;
    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push(song);
  }

  return {
    songs: deduped,
    dbPath: songsDbPath,
    mbPath: songsMbPath,
    titleCount: titles.length,
    lyricCount: lyrics.length,
  };
}


function _collectImportableSongFilesFromFolder(folderPath) {
  const files = [];
  const walk = (dir, depth = 0) => {
    if (depth > 4) return;
    let names = [];
    try { names = fs.readdirSync(dir); } catch (_) { return; }
    for (const name of names) {
      const abs = path.join(dir, name);
      let st;
      try { st = fs.statSync(abs); } catch (_) { continue; }
      if (st.isDirectory()) {
        if (/^(cache|logs?|tmp|temp|thumbnails?)$/i.test(name)) continue;
        walk(abs, depth + 1);
      } else {
        const ext = path.extname(name).toLowerCase();
        if (['.json', '.txt', '.rtf', '.db', '.sqlite', '.sqlite3', '.db3', '.pro', '.pro4', '.pro5', '.pro6', '.propresenter'].includes(ext)) {
          files.push(abs);
        }
      }
    }
  };
  walk(folderPath, 0);
  return Array.from(new Set(files));
}

function _detectSongImportSource(targetPath) {
  const abs = path.resolve(String(targetPath || ''));
  if (!fs.existsSync(abs)) return { kind: 'unknown', path: abs };
  const stat = fs.statSync(abs);
  if (stat.isDirectory()) {
    const names = fs.readdirSync(abs);
    const lower = names.map(v => v.toLowerCase());
    if (lower.includes('songs.db')) return { kind: 'easyworship-data', path: abs }; // legacy song library
    return { kind: 'folder-scan', path: abs, files: _collectImportableSongFilesFromFolder(abs) };
  }
  const ext = path.extname(abs).toLowerCase();
  if (['.pro', '.pro4', '.pro5', '.pro6', '.propresenter'].includes(ext)) return { kind: 'propresenter-file', path: abs, files: [abs] };
  if (/^songs\.db$/i.test(path.basename(abs))) return { kind: 'easyworship-db-file', path: abs, files: [abs] };
  if (['.json', '.txt', '.rtf', '.db', '.sqlite', '.sqlite3', '.db3'].includes(ext)) return { kind: 'direct-file', path: abs, files: [abs] };
  return { kind: 'unknown', path: abs };
}

function _parseProPresenterSongFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const baseTitle = path.basename(filePath, ext).replace(/_/g, ' ');

  // .pro files are binary Protocol Buffers, NOT XML.
  // Detect: if first bytes are not '<?xml' or 'bplist', treat as binary protobuf.
  const rawBuf = fs.readFileSync(filePath);
  const headerStr = rawBuf.slice(0, 6).toString('ascii');
  const isXml    = headerStr.startsWith('<?xml');
  const isBplist = headerStr.startsWith('bplist');
  const isBinary = !isXml && !isBplist;

  if (isBinary) {
    // ── Binary protobuf lyrics extractor ──────────────────────────────────
    // Lyrics are stored as RTF strings embedded in the binary.
    // Each slide: {\\rtf0...\\cb3 LYRIC TEXT}  preceded by a label string.
    const latin1 = rawBuf.toString('latin1');

    // Extract title from filename — most reliable source for PP files
    // (The binary title field is hard to locate reliably across PP versions)
    let title = baseTitle;
    // Try to find the song title in the binary: it appears as a short readable
    // string near the top of the file, right after the file header bytes.
    // Look in first 512 bytes only to avoid matching lyric text.
    const headerRegion = latin1.slice(0, 512);
    const titleMatch = headerRegion.match(/([A-Z][A-Za-z0-9 '!&,.-]{1,60})(?=\x22|\n)/);
    if (titleMatch && titleMatch[1].length > 2 && !/^[0-9a-f-]{8,}$/i.test(titleMatch[1])) {
      title = titleMatch[1].trim();
    }

    // ── Extract lyrics — supports PP6 (\cb3) and PP7 (\cf2) formats ───────────
    // PP7 uses \cf2 + spaces before text; lines split by \\\n within block; block ends at }0
    // PP6 uses \cb3 before text; lines split by \par; each line is a separate block

    function _cleanRtfLine(s) {
      return String(s || '')
        .replace(/\\u(\d+)\s*\?/g, (_, c) => { try { return String.fromCodePoint(parseInt(c)); } catch { return "'"; } })
        .replace(/\\[a-z*-]+\d*\s*/g, '')
        .replace(/[{}]/g, '')
        .replace(/\r/g, '')
        .replace(/^[\s\-\d]+/, '')   // strip leading junk like "-50 "
        .trim();
    }

    const allLines = [];
    const isPP7 = latin1.includes('partightenfactor') && latin1.includes('\\cf2 ');

    if (isPP7) {
      // PP7: text follows \cf2 + spaces, lines split by backslash+newline, block ends at }0
      // Use indexOf (not regex) because the backslash in the binary is a literal single byte
      const cf2marker = '\x5ccf2 ';   // literal \cf2 space
      const endmarker = '\x7d0';      // literal }0
      const lineBreak = '\x5c\x0a';  // literal backslash + newline (within-slide separator)
      let searchPos = 0;
      while (true) {
        const start = latin1.indexOf(cf2marker, searchPos);
        if (start === -1) break;
        let textStart = start + cf2marker.length;
        while (textStart < latin1.length && latin1[textStart] === ' ') textStart++;
        const end = latin1.indexOf(endmarker, textStart);
        if (end === -1) break;
        const block = latin1.slice(textStart, end);
        const lines = block.split(lineBreak)
          .map(l => _cleanRtfLine(l))
          .filter(l => l.length > 1 && /[A-Za-z]/.test(l));
        if (lines.length > 0) allLines.push({ text: lines.join('\n'), pos: start });
        searchPos = end + 2;
      }
    } else {
      // PP6: text follows \cb3, each line is a separate match
      const cb3Re = /\\cb\d+[ \t]?([^\x00-\x08\x0e-\x1f\x7f-\xff][^\x00-\x08]{2,300}?)(?=\\par|\\pard|\}0)/g;
      const rawLines = [];
      let rm6;
      while ((rm6 = cb3Re.exec(latin1)) !== null) {
        const cleaned = _cleanRtfLine(
          rm6[1].replace(/\\par\b/gi, '\n').replace(/\\[a-z]+\d*\s*/g, '')
        );
        if (cleaned.length > 2 && /[A-Za-z]/.test(cleaned)) {
          rawLines.push({ text: cleaned, pos: rm6.index });
        }
      }
      // Pair consecutive lines that are close together into one slide
      for (let i = 0; i < rawLines.length; i++) {
        const cur = rawLines[i], next = rawLines[i + 1];
        if (next && (next.pos - cur.pos) < 2000) {
          allLines.push({ text: cur.text + '\n' + next.text, pos: cur.pos });
          i++;
        } else {
          allLines.push({ text: cur.text, pos: cur.pos });
        }
      }
    }

    // Each entry in allLines is already one complete slide
    const slideGroups = allLines.map(b => ({
      lines: b.text.split('\n').filter(l => l.trim()),
      pos: b.pos
    })).filter(g => g.lines.length > 0);

    if (slideGroups.length > 0) {
      // Extract slide labels (Verse, Chorus, Bridge etc.) with positions
      const labelRe = /\b(Verse|Chorus|Bridge|Pre-?Chorus|Intro|Outro|Tag|Interlude|Ending|VERSE|CHORUS|BRIDGE)\b/g;
      const allLabels = [];
      let lm;
      while ((lm = labelRe.exec(latin1)) !== null) {
        allLabels.push({
          label: lm[1].charAt(0).toUpperCase() + lm[1].slice(1).toLowerCase(),
          pos: lm.index
        });
      }

      // Match each slide group to nearest preceding label
      const labelCounts = {};
      const sections = slideGroups.map((grp, i) => {
        let best = '', bestDist = Infinity;
        for (const l of allLabels) {
          if (l.pos < grp.pos && (grp.pos - l.pos) < 4000 && (grp.pos - l.pos) < bestDist) {
            bestDist = grp.pos - l.pos;
            best = l.label;
          }
        }
        const base = best || 'Verse';
        labelCounts[base] = (labelCounts[base] || 0) + 1;
        const label = labelCounts[base] > 1 ? `${base} ${labelCounts[base]}` : base;
        return { label, lines: grp.lines.filter(l => l.trim()) };
      }).filter(s => s.lines.length > 0);

      if (sections.length > 0) {
        return _normalizeSongForIo({ title, author: '', sections });
      }
    }
    // Fall through to XML/text parser if binary extraction found nothing
  }

  // ── XML / plist / JSON song format ─────────────────────────────────────────
  const raw = isBinary ? rawBuf.toString('latin1') : rawBuf.toString('utf8');
  let title = '', author = '', body = '';

  const tm = raw.match(/<key>\s*title\s*<\/key>\s*<string>([\s\S]*?)<\/string>/i)
    || raw.match(/<title>([\s\S]*?)<\/title>/i)
    || raw.match(/"title"\s*:\s*"([^"]+)"/i);
  if (tm) title = String(tm[1] || '').replace(/&amp;/g,'&').trim();

  const am = raw.match(/<key>\s*author\s*<\/key>\s*<string>([\s\S]*?)<\/string>/i)
    || raw.match(/<author>([\s\S]*?)<\/author>/i)
    || raw.match(/"author"\s*:\s*"([^"]+)"/i);
  if (am) author = String(am[1] || '').replace(/&amp;/g,'&').trim();

  const bm = raw.match(/<key>\s*(?:lyrics|text|plainText)\s*<\/key>\s*<string>([\s\S]*?)<\/string>/i)
    || raw.match(/<(?:lyrics|text)>([\s\S]*?)<\/(?:lyrics|text)>/i)
    || raw.match(/"plainText"\s*:\s*"([\s\S]*?)"/i)
    || raw.match(/"text"\s*:\s*"([\s\S]*?)"/i);
  if (bm) body = String(bm[1] || '');

  if (!body) {
    body = raw.replace(/<[^>]+>/g,'\n').replace(/\\n/g,'\n')
      .replace(/\r\n/g,'\n').replace(/\r/g,'\n')
      .replace(/[ \t]+\n/g,'\n').replace(/\n{3,}/g,'\n\n').trim();
  }
  body = body.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
    .replace(/&#10;/g,'\n').replace(/\\n/g,'\n').trim();

  const sections = _parseSongTextToSections(body);
  return _normalizeSongForIo({
    title: title || baseTitle || 'Imported Song',
    author,
    sections,
  });
}


function _importEasyWorshipDataFolder(dataDir, mode = 'merge') {
  const extracted = _extractEasyWorshipSongsFromDataFolder(dataDir);
  const currentSongs = _readSongLibrary();
  let output = [];
  let imported = 0;

  if (mode === 'replace') {
    const seen = new Set();
    for (const song of extracted.songs) {
      const key = `${String(song.title || '').trim().toLowerCase()}|${String(song.author || '').trim().toLowerCase()}`;
      if (seen.has(key)) continue;
      seen.add(key);
      output.push(song);
    }
    imported = output.length;
  } else {
    const merged = _dedupeSongs(currentSongs, extracted.songs);
    output = merged.songs;
    imported = merged.added;
  }

  _writeSongLibrary(output, true);
  return {
    success: true,
    imported,
    total: output.length,
    scanned: extracted.songs.length,
    source: extracted.dbPath,
    hasMb: extracted.hasMb,
    detected: 'easyworship-data',
    mode,
  };
}

function _importSongsFromSmartSource(targetPath, mode = 'merge') {
  const detected = _detectSongImportSource(targetPath);
  if (detected.kind === 'easyworship-data') {
    return _importEasyWorshipDataFolder(detected.path, mode);
  }
  if (detected.kind === 'propresenter-file') {
    const incoming = [_parseProPresenterSongFile(detected.path)];
    const currentSongs = _readSongLibrary();
    if (mode === 'replace') {
      const output = [];
      const seen = new Set();
      for (const song of incoming) {
        const key = `${String(song.title || '').trim().toLowerCase()}|${String(song.author || '').trim().toLowerCase()}`;
        if (seen.has(key)) continue;
        seen.add(key);
        output.push(song);
      }
      _writeSongLibrary(output, true);
      return { success: true, imported: output.length, total: output.length, scanned: output.length, source: detected.path, detected: 'propresenter-file', mode };
    }
    const merged = _dedupeSongs(currentSongs, incoming);
    _writeSongLibrary(merged.songs, true);
    return { success: true, imported: merged.added, total: merged.songs.length, scanned: incoming.length, source: detected.path, detected: 'propresenter-file', mode };
  }
  if (detected.kind === 'folder-scan') {
    const files = detected.files || [];
    if (!files.length) throw new Error('No importable song files were found in the selected folder.');
    return _importSongsFromPaths(files, mode);
  }
  if (detected.kind === 'direct-file') {
    return _importSongsFromPaths(detected.files || [detected.path], mode);
  }
  throw new Error('The selected file or folder is not a supported song source. Choose a song file or folder from a compatible source.');
}


function _tryImportEasyWorshipFromSelectedFiles(filePaths = [], mode = 'merge') {
  const files = (filePaths || []).map(f => path.resolve(String(f || ''))).filter(Boolean);
  if (!files.length) return null;

  const lower = files.map(f => path.basename(f).toLowerCase());
  const dbIdx = lower.findIndex(n => n === 'songs.db');
  const mbIdx = lower.findIndex(n => n === 'songs.mb');

  if (dbIdx === -1) return null;

  const dbPath = files[dbIdx];
  const dataDir = path.dirname(dbPath);
  const mbPath = mbIdx !== -1 ? files[mbIdx] : path.join(dataDir, 'Songs.MB');

  if (!fs.existsSync(dbPath)) return null;
  if (!fs.existsSync(mbPath)) {
    throw new Error('Songs.DB was selected but Songs.MB was not found. Select both files together, or select the full DATA folder.');
  }

  return _importEasyWorshipDataFolder(dataDir, mode);
}

function _importSongsFromPaths(filePaths = [], mode = 'merge') {
  // Route .pro files to their binary parser — never treat as text/db
  const proExts = ['.pro', '.pro4', '.pro5', '.pro6', '.propresenter'];
  const proPaths   = (filePaths || []).filter(f => proExts.includes(path.extname(f).toLowerCase()));
  const otherPaths = (filePaths || []).filter(f => !proExts.includes(path.extname(f).toLowerCase()));

  if (proPaths.length > 0 && otherPaths.length === 0) {
    const incoming = proPaths.map(f => _parseProPresenterSongFile(f));
    const currentSongs = _readSongLibrary();
    if (mode === 'replace') {
      const seen = new Set(); const output = [];
      for (const song of incoming) {
        const key = String(song.title||'').trim().toLowerCase();
        if (!seen.has(key)) { seen.add(key); output.push(song); }
      }
      _writeSongLibrary(output, true);
      return { success: true, imported: output.length, total: output.length, scanned: proPaths.length, detected: 'propresenter-file', mode };
    }
    const merged = _dedupeSongs(currentSongs, incoming);
    _writeSongLibrary(merged.songs, true);
    return { success: true, imported: merged.added, total: merged.songs.length, scanned: proPaths.length, detected: 'propresenter-file', mode };
  }

  const workPaths = otherPaths.length ? otherPaths : (filePaths || []);
  const easyWorshipImport = _tryImportEasyWorshipFromSelectedFiles(workPaths, mode);
  if (easyWorshipImport) return easyWorshipImport;

  let currentSongs = _readSongLibrary();
  const incoming = [];
  for (const filePath of workPaths) {
    const ext = path.extname(filePath).toLowerCase();
    if (ext === '.json') {
      const parsed = JSON.parse(fs.readFileSync(filePath, 'utf8'));
      if (Array.isArray(parsed)) incoming.push(...parsed);
      else if (Array.isArray(parsed?.songs)) incoming.push(...parsed.songs);
      else if (parsed && parsed.title && parsed.sections) incoming.push(parsed);
    } else if (['.db', '.sqlite', '.sqlite3', '.db3'].includes(ext)) {
      incoming.push(..._extractSongsFromSqliteDb(filePath));
    } else {
      const raw = fs.readFileSync(filePath, 'utf8');
      incoming.push(_parseEasyWorshipSongBuffer(filePath, raw));
    }
  }
  const normalizedIncoming = incoming.map((song, i) => _normalizeSongForIo(song, i)).filter(song => song.title && song.sections?.length);
  let output;
  let imported = 0;
  if (mode === 'replace') {
    output = [];
    for (const song of normalizedIncoming) {
      const key = `${String(song.title || '').trim().toLowerCase()}|${String(song.author || '').trim().toLowerCase()}`;
      if (!output.some(x => `${String(x.title||'').trim().toLowerCase()}|${String(x.author||'').trim().toLowerCase()}` === key)) {
        output.push(song);
      }
    }
    imported = output.length;
  } else {
    const merged = _dedupeSongs(currentSongs, normalizedIncoming);
    output = merged.songs;
    imported = merged.added;
  }
  _writeSongLibrary(output, true);
  return { success: true, imported, total: output.length, mode };
}


ipcMain.handle('import-song-library-folder', async (_, { mode = 'merge' } = {}) => {
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Select Song Library Folder',
      properties: ['openDirectory']
    });
    if (result.canceled || !result.filePaths?.length) return { success:false, canceled:true };
    return _importEasyWorshipDataFolder(result.filePaths[0], mode);
  } catch (e) {
    return { success:false, error:e.message };
  }
});


function _collectAppDataBackupEntries() {
  const entries = [];
  const pushIfExists = (relPath) => {
    const abs = path.join(DATA_DIR, relPath);
    try { if (fs.existsSync(abs)) entries.push({ abs, rel: path.join('AnchorCastData', relPath) }); } catch (_) {}
  };

  [
    'settings.json',
    'themes.json',
    'songs.json',
    'media.json',
    'presentations.json',
    'created_presentations.json',
    'transcripts.json',
    'recent_schedules.json',
    'song-backup-meta.json',
    'whisper_python.txt'
  ].forEach(pushIfExists);

  const addTree = (dirRel) => {
    const absDir = path.join(DATA_DIR, dirRel);
    if (!fs.existsSync(absDir)) return;
    const walk = (folder) => {
      for (const name of fs.readdirSync(folder)) {
        const abs = path.join(folder, name);
        const stat = fs.statSync(abs);
        if (stat.isDirectory()) walk(abs);
        else {
          const rel = path.relative(DATA_DIR, abs);
          entries.push({ abs, rel: path.join('AnchorCastData', rel) });
        }
      }
    };
    walk(absDir);
  };

  addTree('bibles');
  addTree('Schedules');
  addTree('song-library-backups');
  addTree('full-backups');
  return entries;
}

function _zipDirectorySync(fileMap, zipPath) {
  const zlib = require('zlib');
  const localHeaders = [];
  const central = [];
  let offset = 0;

  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
    return table;
  })();

  const crc32 = (buf) => {
    let crc = 0 ^ (-1);
    for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ crcTable[(crc ^ buf[i]) & 0xFF];
    return (crc ^ (-1)) >>> 0;
  };

  const dosDateTime = (date) => {
    const d = date || new Date();
    const year = Math.max(1980, d.getFullYear());
    const dosTime = ((d.getHours() & 0x1f) << 11) | ((d.getMinutes() & 0x3f) << 5) | ((Math.floor(d.getSeconds()/2)) & 0x1f);
    const dosDate = (((year - 1980) & 0x7f) << 9) | (((d.getMonth()+1) & 0xf) << 5) | (d.getDate() & 0x1f);
    return { dosTime, dosDate };
  };

  const chunks = [];
  for (const entry of fileMap) {
    const nameBuf = Buffer.from(String(entry.rel).replace(/\\/g, '/'));
    const data = fs.readFileSync(entry.abs);
    const compressed = zlib.deflateRawSync(data);
    const crc = crc32(data);
    const { dosTime, dosDate } = dosDateTime(fs.statSync(entry.abs).mtime);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(dosTime, 10);
    local.writeUInt16LE(dosDate, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);

    const centralHdr = Buffer.alloc(46);
    centralHdr.writeUInt32LE(0x02014b50, 0);
    centralHdr.writeUInt16LE(20, 4);
    centralHdr.writeUInt16LE(20, 6);
    centralHdr.writeUInt16LE(0, 8);
    centralHdr.writeUInt16LE(8, 10);
    centralHdr.writeUInt16LE(dosTime, 12);
    centralHdr.writeUInt16LE(dosDate, 14);
    centralHdr.writeUInt32LE(crc, 16);
    centralHdr.writeUInt32LE(compressed.length, 20);
    centralHdr.writeUInt32LE(data.length, 24);
    centralHdr.writeUInt16LE(nameBuf.length, 28);
    centralHdr.writeUInt16LE(0, 30);
    centralHdr.writeUInt16LE(0, 32);
    centralHdr.writeUInt16LE(0, 34);
    centralHdr.writeUInt16LE(0, 36);
    centralHdr.writeUInt32LE(0, 38);
    centralHdr.writeUInt32LE(offset, 42);

    chunks.push(local, nameBuf, compressed);
    localHeaders.push(local.length + nameBuf.length + compressed.length);
    central.push(centralHdr, nameBuf);
    offset += local.length + nameBuf.length + compressed.length;
  }

  const centralSize = central.reduce((n, b) => n + b.length, 0);
  const centralOffset = offset;
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(fileMap.length, 8);
  end.writeUInt16LE(fileMap.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(centralOffset, 16);
  end.writeUInt16LE(0, 20);

  fs.mkdirSync(path.dirname(zipPath), { recursive: true });
  fs.writeFileSync(zipPath, Buffer.concat([...chunks, ...central, end]));
  return zipPath;
}

function _extractZipWithPython(zipPath, destinationDir) {
  const pyScript = `
import os, sys, zipfile
zip_path = sys.argv[1]
dest = sys.argv[2]
os.makedirs(dest, exist_ok=True)
with zipfile.ZipFile(zip_path, 'r') as z:
    z.extractall(dest)
print(dest)
`.trim();
  for (const candidate of _songLibraryPythonCandidates()) {
    try {
      const args = candidate === 'py' ? ['-3', '-c', pyScript, zipPath, destinationDir] : ['-c', pyScript, zipPath, destinationDir];
      const res = spawnSync(candidate, args, { encoding: 'utf8', windowsHide: true, timeout: 30000, maxBuffer: 10 * 1024 * 1024 });
      if (!res.error && (res.status || 0) === 0) return true;
    } catch (_) {}
  }
  return false;
}

function _restoreAppDataFromExtractedFolder(sourceRoot) {
  const rootDir = fs.existsSync(path.join(sourceRoot, 'AnchorCastData')) ? path.join(sourceRoot, 'AnchorCastData') : sourceRoot;
  const copyTree = (srcDir, dstDir) => {
    if (!fs.existsSync(srcDir)) return;
    fs.mkdirSync(dstDir, { recursive: true });
    for (const name of fs.readdirSync(srcDir)) {
      const absSrc = path.join(srcDir, name);
      const absDst = path.join(dstDir, name);
      const st = fs.statSync(absSrc);
      if (st.isDirectory()) copyTree(absSrc, absDst);
      else fs.copyFileSync(absSrc, absDst);
    }
  };
  copyTree(rootDir, DATA_DIR);
}

// ── Songs ────────────────────────────────────────────────────────────────────
ipcMain.handle('get-songs', () => {
  try{
    if(fs.existsSync(SONGS_FILE)) return JSON.parse(fs.readFileSync(SONGS_FILE,'utf-8'));
  } catch(e){}
  return [];
});
ipcMain.handle('save-songs', (_, songs) => {
  try{
    fs.mkdirSync(path.dirname(SONGS_FILE),{recursive:true});
    fs.writeFileSync(SONGS_FILE, JSON.stringify(songs,null,2));
    // Notify main window to reload song library panel
    mainWindow?.webContents.send('songs-saved');
    return {success:true};
  } catch(e){ return {success:false, error:e.message}; }
});
ipcMain.handle('open-song-manager', (_, opts = {}) => {
  createSongManagerWindow();
  if (opts?.songId != null) {
    setTimeout(() => songManagerWindow?.webContents.send('song-manager-open-song', { songId: opts.songId }), 300);
  }
  if (opts?.newSong) {
    setTimeout(() => songManagerWindow?.webContents.send('song-manager-new-song'), 320);
  }
  return { success:true };
});

ipcMain.handle('import-songs-file', async (_, { mode = 'merge' } = {}) => {
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Import Songs',
      properties: ['openFile', 'multiSelections'],
      filters: [
        { name: 'Song Library Files', extensions: ['json', 'txt', 'rtf', 'db', 'sqlite', 'sqlite3', 'db3', 'pro', 'pro4', 'pro5', 'pro6'] },
        { name: 'All Files', extensions: ['*'] },
      ],
    });
    if (result.canceled || !result.filePaths?.length) return { success: false, canceled: true, imported: 0 };
    return _importSongsFromPaths(result.filePaths, mode);
  } catch (e) {
    return { success: false, error: e.message, imported: 0 };
  }
});

ipcMain.handle('import-smart-song-source', async (_, { mode = 'merge' } = {}) => {
  try {
    const choice = await dialog.showMessageBox(mainWindow, {
      type: 'question',
      title: 'Import From other app',
      message: 'Choose what you want to import from',
      detail: 'Pick File for a song file (JSON, TXT, or Songs.DB). Pick Folder for a song library folder.',
      buttons: ['Choose File', 'Choose Folder', 'Cancel'],
      cancelId: 2,
      defaultId: 0,
      noLink: true,
    });
    if (choice.response === 2) return { success: false, canceled: true, imported: 0 };

    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Import From other app',
      properties: choice.response === 1 ? ['openDirectory'] : ['openFile', 'multiSelections'],
      filters: [
        { name: 'Song Sources', extensions: ['json', 'txt', 'rtf', 'db', 'sqlite', 'sqlite3', 'db3', 'mb', 'pro', 'pro4', 'pro5', 'pro6', 'propresenter'] },
        { name: 'All Files', extensions: ['*'] },
      ],
    });
    if (result.canceled || !result.filePaths?.length) return { success: false, canceled: true, imported: 0 };
    return choice.response === 1
      ? _importSongsFromSmartSource(result.filePaths[0], mode)
      : _importSongsFromPaths(result.filePaths, mode);
  } catch (e) {
    return { success: false, error: e.message, imported: 0 };
  }
});

ipcMain.handle('backup-full-appdata', async () => {
  try {
    const result = await dialog.showSaveDialog(mainWindow, {
      title: 'Backup AnchorCast Data',
      defaultPath: path.join(app.getPath('documents'), `AnchorCastData-Backup-${new Date().toISOString().slice(0,10)}.zip`),
      filters: [{ name: 'ZIP Backup', extensions: ['zip'] }],
      buttonLabel: 'Create Backup',
    });
    if (result.canceled || !result.filePath) return { success: false, canceled: true };
    const entries = _collectAppDataBackupEntries();
    if (!entries.length) return { success: false, error: 'No AnchorCast data found to back up.' };
    _zipDirectorySync(entries, result.filePath);
    return { success: true, path: result.filePath, count: entries.length };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('restore-full-appdata', async () => {
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Restore AnchorCast Data Backup',
      properties: ['openFile'],
      filters: [{ name: 'ZIP Backup', extensions: ['zip'] }],
      buttonLabel: 'Restore Backup',
    });
    if (result.canceled || !result.filePaths?.length) return { success: false, canceled: true };
    const zipPath = result.filePaths[0];
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'anchorcast-restore-'));
    const ok = _extractZipWithPython(zipPath, tempDir);
    if (!ok) return { success: false, error: 'Could not extract ZIP backup. Python zip support was not available.' };
    _restoreAppDataFromExtractedFolder(tempDir);
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch (_) {}
    mainWindow?.webContents.send('songs-saved');
    mainWindow?.webContents.send('themes-updated');
    mainWindow?.webContents.send('bible-versions-updated');
    return { success: true, restoredFrom: zipPath, requiresRestart: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('get-song-library-info', () => {
  try {
    const songs = _readSongLibrary();
    const backups = fs.existsSync(SONG_BACKUP_DIR)
      ? fs.readdirSync(SONG_BACKUP_DIR).filter(f => f.endsWith('.json')).length
      : 0;
    return {
      count: songs.length,
      backups,
      path: SONGS_FILE,
      backupDir: SONG_BACKUP_DIR,
    };
  } catch (e) {
    return { count: 0, backups: 0, path: SONGS_FILE, backupDir: SONG_BACKUP_DIR, error: e.message };
  }
});

// ── Genius Lyrics Search ───────────────────────────────────────────────────────
// Runs in main process: no CORS, API key never sent to renderer.
function _getGeniusApiKey() {
  if (process.env.GENIUS_API_KEY) return process.env.GENIUS_API_KEY;
  try {
    if (fs.existsSync(SETTINGS_FILE)) {
      const s = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8'));
      if (s.geniusApiKey) return s.geniusApiKey;
    }
  } catch (_) {}
  return '';
}

ipcMain.handle('genius-search', async (_, { query }) => {
  if (!query || !query.trim()) return { results: [] };
  const GENIUS_API_KEY = _getGeniusApiKey();
  if (!GENIUS_API_KEY) return { error: 'Genius API key not configured. Add it in Settings → AI & Transcription.', results: [] };
  try {
    const url = `https://api.genius.com/search?q=${encodeURIComponent(query.trim())}&per_page=10`;
    const data = await new Promise((resolve, reject) => {
      const req = https.get(url, {
        headers: { 'Authorization': `Bearer ${GENIUS_API_KEY}`, 'User-Agent': 'AnchorCast/1.0' }
      }, (res) => {
        let body = '';
        res.on('data', chunk => body += chunk);
        res.on('end', () => {
          try { resolve(JSON.parse(body)); }
          catch (e) { reject(new Error('Invalid JSON from Genius')); }
        });
      });
      req.on('error', reject);
      req.setTimeout(8000, () => { req.destroy(); reject(new Error('Genius search timed out')); });
    });
    const hits = (data?.response?.hits || [])
      .filter(h => h.type === 'song')
      .map(h => ({
        id:          h.result.id,
        title:       h.result.title,
        artist:      h.result.primary_artist?.name || '',
        thumbnail:   h.result.song_art_image_thumbnail_url || '',
        url:         h.result.url,
        path:        h.result.path,
        fullTitle:   h.result.full_title,
      }));
    return { results: hits };
  } catch (e) {
    console.warn('[Genius] Search error:', e.message);
    return { error: e.message, results: [] };
  }
});

ipcMain.handle('genius-fetch-lyrics', async (_, { url }) => {
  if (!url) return { error: 'No URL provided' };
  try {
    const html = await new Promise((resolve, reject) => {
      const req = https.get(url, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'text/html,application/xhtml+xml',
          'Accept-Language': 'en-US,en;q=0.9',
        }
      }, (res) => {
        // Follow redirects
        if (res.statusCode === 301 || res.statusCode === 302) {
          const redirUrl = res.headers.location;
          if (redirUrl) { resolve(_geniusFetch(redirUrl)); return; }
        }
        let body = '';
        res.on('data', chunk => body += chunk);
        res.on('end', () => resolve(body));
      });
      req.on('error', reject);
      req.setTimeout(12000, () => { req.destroy(); reject(new Error('Fetch timed out')); });
    });

    // Extract lyrics from Genius page HTML
    // Genius wraps lyrics in data-lyrics-container="true" divs
    const lyrics = _parseGeniusLyrics(html);
    if (!lyrics) return { error: 'Could not extract lyrics — Genius page format may have changed' };
    return { lyrics };
  } catch (e) {
    console.warn('[Genius] Fetch error:', e.message);
    return { error: e.message };
  }
});

// ── LRCLIB fallback (free, no API key, synced + plain lyrics) ─────────────────
ipcMain.handle('lrclib-search', async (_, { query, artist }) => {
  if (!query || !query.trim()) return { results: [] };
  try {
    // Build search URL — try track_name+artist_name first, fallback to q param
    const baseUrl = 'https://lrclib.net/api';
    const searches = [];
    if (artist && artist.trim()) {
      searches.push(`${baseUrl}/search?track_name=${encodeURIComponent(query.trim())}&artist_name=${encodeURIComponent(artist.trim())}`);
    }
    searches.push(`${baseUrl}/search?q=${encodeURIComponent(query.trim())}`);

    let hits = [];
    for (const url of searches) {
      const data = await new Promise((resolve, reject) => {
        const req = https.get(url, {
          headers: { 'User-Agent': 'AnchorCast/1.0 (https://github.com/anchorcastapp-team/anchorcastapp)', 'Accept': 'application/json' }
        }, (res) => {
          let body = '';
          res.on('data', chunk => body += chunk);
          res.on('end', () => {
            try { resolve(JSON.parse(body)); }
            catch(e) { resolve([]); }
          });
        });
        req.on('error', reject);
        req.setTimeout(8000, () => { req.destroy(); reject(new Error('LRCLIB search timed out')); });
      });
      const arr = Array.isArray(data) ? data : [];
      // Prefer tracks with plain lyrics
      hits = arr
        .filter(t => t.plainLyrics || t.syncedLyrics)
        .map(t => ({
          id:          t.id,
          title:       t.trackName || '',
          artist:      t.artistName || '',
          album:       t.albumName  || '',
          duration:    t.duration   || 0,
          hasPlain:    !!t.plainLyrics,
          hasSynced:   !!t.syncedLyrics,
          plainLyrics: t.plainLyrics  || '',
          source:      'lrclib',
        }));
      if (hits.length) break; // use first successful search
    }
    return { results: hits };
  } catch(e) {
    console.warn('[LRCLIB] Search error:', e.message);
    return { error: e.message, results: [] };
  }
});

ipcMain.handle('lrclib-fetch-lyrics', async (_, { id }) => {
  if (!id) return { error: 'No ID provided' };
  try {
    const url = `https://lrclib.net/api/get/${id}`;
    const data = await new Promise((resolve, reject) => {
      const req = https.get(url, {
        headers: { 'User-Agent': 'AnchorCast/1.0 (https://github.com/anchorcastapp-team/anchorcastapp)', 'Accept': 'application/json' }
      }, (res) => {
        let body = '';
        res.on('data', chunk => body += chunk);
        res.on('end', () => {
          try { resolve(JSON.parse(body)); }
          catch(e) { resolve(null); }
        });
      });
      req.on('error', reject);
      req.setTimeout(8000, () => { req.destroy(); reject(new Error('LRCLIB fetch timed out')); });
    });
    if (!data || (!data.plainLyrics && !data.syncedLyrics)) {
      return { error: 'No lyrics found for this track' };
    }
    // Return plain lyrics, stripping LRC timestamps if needed from synced
    const plain = data.plainLyrics
      || (data.syncedLyrics || '').replace(/\[\d+:\d+\.\d+\]\s*/g, '').trim();
    return { lyrics: plain, hasSynced: !!data.syncedLyrics };
  } catch(e) {
    console.warn('[LRCLIB] Fetch error:', e.message);
    return { error: e.message };
  }
});

function _geniusFetch(url) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36' }
    }, (res) => {
      let body = '';
      res.on('data', c => body += c);
      res.on('end', () => resolve(body));
    });
    req.on('error', reject);
    req.setTimeout(12000, () => { req.destroy(); reject(new Error('Fetch timed out')); });
  });
}

function _parseGeniusLyrics(html) {
  try {
    // Strategy 1: data-lyrics-container divs (current Genius format)
    const containerMatches = [...html.matchAll(/data-lyrics-container="true"[^>]*>([\s\S]*?)<\/div>/g)];
    if (containerMatches.length) {
      const raw = containerMatches.map(m => m[1]).join('\n');
      return _cleanGeniusHtml(raw);
    }
    // Strategy 2: JSON embedded in page (window.__PRELOADED_STATE__ or similar)
    const jsonMatch = html.match(/window\.__PRELOADED_STATE__\s*=\s*JSON\.parse\('([\s\S]*?)'\)/);
    if (jsonMatch) {
      try {
        const state = JSON.parse(jsonMatch[1].replace(/\\'/g,"'").replace(/\\"/g,'"'));
        const lyricsText = state?.entities?.songs?.byId
          ? Object.values(state.entities.songs.byId)[0]?.lyricsBody
          : null;
        if (lyricsText) return lyricsText;
      } catch(_) {}
    }
    // Strategy 3: og:description meta (plain text snippet, partial)
    const descMatch = html.match(/<meta[^>]+property="og:description"[^>]+content="([^"]+)"/);
    if (descMatch) return decodeHtmlEntities(descMatch[1]);
    return null;
  } catch(e) {
    console.warn('[Genius] Parse error:', e.message);
    return null;
  }
}

function _cleanGeniusHtml(html) {
  return html
    .replace(/<br\s*\/?>/gi, '\n')        // <br> → newline
    .replace(/<\/p>/gi, '\n')             // </p> → newline
    .replace(/<a[^>]*>/gi, '')            // strip <a>
    .replace(/<\/a>/gi, '')
    .replace(/<span[^>]*>/gi, '')         // strip <span>
    .replace(/<\/span>/gi, '')
    .replace(/<i[^>]*>/gi, '')
    .replace(/<\/i>/gi, '')
    .replace(/<b[^>]*>/gi, '')
    .replace(/<\/b>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '')      // strip HTML comments
    .replace(/<[^>]+>/g, '')              // strip any remaining tags
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#x27;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/\[([^\]]+)\]/g, '\n[$1]\n')  // [Verse 1] etc → own lines
    .replace(/\n{3,}/g, '\n\n')            // collapse 3+ newlines → 2
    .trim();
}

function decodeHtmlEntities(str) {
  return str
    .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
    .replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&nbsp;/g,' ');
}

ipcMain.handle('backup-song-library', (_, { reason = 'manual' } = {}) => {
  try {
    const backup = _makeSongLibraryBackup(reason);
    return { success: true, ...backup };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('export-songs', async (_, { mode = 'easyworship-txt' } = {}) => {
  try {
    const songs = _safeJsonRead(SONGS_FILE, []).map((song, i) => _normalizeSongForIo(song, i));
    if (!songs.length) return { success: false, error: 'No songs to export' };

    if (mode === 'json') {
      const result = await dialog.showSaveDialog(mainWindow, {
        title: 'Export Song Library',
        defaultPath: path.join(app.getPath('documents'), 'AnchorCast Songs.json'),
        filters: [{ name: 'JSON', extensions: ['json'] }],
      });
      if (result.canceled || !result.filePath) return { success: false, canceled: true };
      fs.writeFileSync(result.filePath, JSON.stringify({ format: 'anchorcast-song-library', exportedAt: Date.now(), songs }, null, 2));
      return { success: true, path: result.filePath, count: songs.length };
    }

    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Choose Export Folder',
      defaultPath: path.join(app.getPath('documents'), 'AnchorCast Song Export'),
      properties: ['openDirectory', 'createDirectory'],
      buttonLabel: 'Export Here',
    });
    if (result.canceled || !result.filePaths?.length) return { success: false, canceled: true };
    const outDir = result.filePaths[0];
    fs.mkdirSync(outDir, { recursive: true });
    const index = [];
    for (const song of songs) {
      const safe = sanitizeFileName(song.title || 'song');
      const filePath = path.join(outDir, `${safe}.txt`);
      fs.writeFileSync(filePath, _serializeSongAsEasyWorshipText(song), 'utf8');
      index.push({ title: song.title, author: song.author || '', file: path.basename(filePath) });
    }
    fs.writeFileSync(path.join(outDir, '_index.json'), JSON.stringify({ format: 'easyworship-text-bundle', exportedAt: Date.now(), songs: index }, null, 2));
    return { success: true, folder: outDir, count: songs.length };
  } catch (e) {
    return { success: false, error: e.message };
  }
});


// ── Media ─────────────────────────────────────────────────────────────────────
ipcMain.handle('get-media', () => {
  try{
    const items = JSON.parse(fs.readFileSync(MEDIA_FILE,'utf8'));
    return (items || []).map(item => ({
      ...item,
      title: item.title || item.name || '',
      mute: item.mute === undefined ? (item.type === 'video' ? false : false) : item.mute,
      loop: item.loop === undefined ? true : item.loop,
      // BUG-F FIX: volume stored as 0-100 but <video>.volume expects 0.0-1.0
      volume: item.volume === undefined ? 1.0 : Math.min(1, Math.max(0, (item.volume > 1 ? item.volume / 100 : item.volume))),
      aspectRatio: item.aspectRatio || 'contain',
    }));
  }
  catch(e){ return []; }
});

ipcMain.handle('save-media', (_, items) => {
  try{
    fs.writeFileSync(MEDIA_FILE, JSON.stringify(items, null, 2));
    return { success: true };
  } catch(e){ return { success: false, error: e.message }; }
});


async function maybeTranscodeWmvToMp4(srcPath) {
  try {
    const ext = path.extname(String(srcPath || '')).toLowerCase();
    if (ext !== '.wmv') return srcPath;
    const candidates = process.platform === 'win32'
      ? ['ffmpeg.exe', 'ffmpeg', 'C:/ffmpeg/bin/ffmpeg.exe']
      : ['ffmpeg', '/usr/bin/ffmpeg', '/usr/local/bin/ffmpeg'];
    const outDir = VIDEO_ASSETS;
    fs.mkdirSync(outDir, { recursive: true });
    const outPath = uniqueDestinationPath(outDir, path.basename(srcPath, ext) + '.mp4');
    for (const ffmpegBin of candidates) {
      try {
        const probe = spawnSync(ffmpegBin, ['-version'], { encoding:'utf8', windowsHide:true, timeout:15000 });
        if (probe.error) continue;
        const run = spawnSync(ffmpegBin, ['-y', '-i', srcPath, '-c:v', 'libx264', '-preset', 'veryfast', '-c:a', 'aac', '-movflags', '+faststart', outPath], { encoding:'utf8', windowsHide:true, timeout: 10 * 60 * 1000, maxBuffer: 16 * 1024 * 1024 });
        if (!run.error && fs.existsSync(outPath) && fs.statSync(outPath).size > 0) {
          return outPath;
        }
      } catch (_) {}
    }
  } catch (_) {}
  return srcPath;
}

ipcMain.handle('import-media-files', async (_, { files = [] } = {}) => {
  try {
    const imported = [];
    const skipped = [];
    for (const src of files) {
      if (!src || !fs.existsSync(src)) { skipped.push({ path: src, reason: 'File not found' }); continue; }
      const type = normalizeMediaType(src);
      if (!type) { skipped.push({ path: src, reason: 'Unsupported format' }); continue; }
      let sourcePath = src;
      if (String(path.extname(src) || '').toLowerCase() === '.wmv') {
        sourcePath = await maybeTranscodeWmvToMp4(src);
      }
      const finalType = normalizeMediaType(sourcePath) || type;
      const destDir = mediaAssetDirForType(finalType);
      fs.mkdirSync(destDir, { recursive: true });
      const destPath = path.resolve(sourcePath).startsWith(path.resolve(destDir) + path.sep) || path.resolve(sourcePath) === path.resolve(destDir)
        ? sourcePath
        : uniqueDestinationPath(destDir, path.basename(sourcePath));
      if (path.resolve(destPath) !== path.resolve(sourcePath)) fs.copyFileSync(sourcePath, destPath);
      imported.push(toStoredMediaItem(destPath));
    }
    return { success: true, items: imported, skipped };
  } catch(e) {
    return { success: false, error: e.message, items: [] };
  }
});


ipcMain.handle('get-media-integrity', () => {
  try {
    const items = fs.existsSync(MEDIA_FILE) ? JSON.parse(fs.readFileSync(MEDIA_FILE, 'utf8')) : [];
    const list = (items || []).map(item => {
      const filePath = item?.path || '';
      const exists = !!(filePath && fs.existsSync(filePath));
      return {
        id: item?.id,
        title: item?.title || item?.name || path.basename(filePath || '') || 'Untitled',
        type: item?.type || normalizeMediaType(filePath) || 'unknown',
        storageLocation: filePath || '',
        exists,
        broken: !exists,
        missingFile: !exists,
      };
    });
    return {
      success: true,
      items: list,
      summary: {
        total: list.length,
        healthy: list.filter(x => x.exists).length,
        broken: list.filter(x => x.broken).length,
      }
    };
  } catch (e) {
    return { success: false, error: e.message, items: [], summary: { total: 0, healthy: 0, broken: 0 } };
  }
});

ipcMain.handle('repair-media-links', () => {
  try {
    const items = fs.existsSync(MEDIA_FILE) ? JSON.parse(fs.readFileSync(MEDIA_FILE, 'utf8')) : [];
    const searchDirs = [VIDEO_ASSETS, AUDIO_ASSETS, IMAGE_ASSETS, LEGACY_MEDIA_ASSETS, path.join(APPDATA_ROOT, 'Media')]
      .filter(Boolean)
      .filter(d => fs.existsSync(d));
    let repaired = 0;
    let unresolved = 0;
    const updated = (items || []).map(item => {
      if (item?.path && fs.existsSync(item.path)) return item;
      const targetBase = path.basename(item?.path || item?.title || item?.name || '');
      let found = '';
      for (const dir of searchDirs) {
        const candidate = path.join(dir, targetBase);
        if (targetBase && fs.existsSync(candidate)) { found = candidate; break; }
      }
      if (!found && targetBase) {
        for (const dir of searchDirs) {
          try {
            const match = fs.readdirSync(dir).find(name => name.toLowerCase() == targetBase.toLowerCase());
            if (match) { found = path.join(dir, match); break; }
          } catch (e) {}
        }
      }
      if (found) {
        repaired += 1;
        return { ...item, path: found, type: item?.type || normalizeMediaType(found) || item?.type };
      }
      unresolved += 1;
      return item;
    });
    fs.writeFileSync(MEDIA_FILE, JSON.stringify(updated, null, 2));
    return { success: true, repaired, unresolved };
  } catch (e) {
    return { success: false, error: e.message, repaired: 0, unresolved: 0 };
  }
});

ipcMain.handle('clear-media-cache', () => {
  try {
    const targets = [THUMBS_DIR, TEMP_MEDIA_DIR, SESSION_CACHE_DIR, CACHE_DIR, path.join(APPDATA_ROOT, 'tmp'), path.join(APPDATA_ROOT, 'temp')];
    let removedFiles = 0;
    let removedDirs = 0;
    for (const target of targets) {
      if (!target || !fs.existsSync(target)) continue;
      try {
        const stat = fs.statSync(target);
        if (stat.isDirectory()) {
          const walk = (dir) => {
            for (const name of fs.readdirSync(dir)) {
              const full = path.join(dir, name);
              const st = fs.statSync(full);
              if (st.isDirectory()) { walk(full); }
              else removedFiles += 1;
            }
          };
          walk(target);
          fs.rmSync(target, { recursive: true, force: true });
          removedDirs += 1;
        } else {
          fs.rmSync(target, { force: true });
          removedFiles += 1;
        }
      } catch (e) {}
    }
    [CACHE_DIR, THUMBS_DIR, TEMP_MEDIA_DIR, SESSION_CACHE_DIR].forEach(d => { try { fs.mkdirSync(d, { recursive: true }); } catch(e) {} });
    const staleSessionFiles = fs.existsSync(DATA_DIR) ? fs.readdirSync(DATA_DIR).filter(name => /^session.*\.(json|tmp)$/i.test(name)) : [];
    for (const name of staleSessionFiles) {
      try { fs.rmSync(path.join(DATA_DIR, name), { force: true }); removedFiles += 1; } catch (e) {}
    }
    return { success: true, removedFiles, removedDirs };
  } catch (e) {
    return { success: false, error: e.message, removedFiles: 0, removedDirs: 0 };
  }
});

ipcMain.handle('open-file-location', async (_, { filePath }) => {
  try {
    if (!filePath) return { success: false, error: 'No file path provided' };
    const target = fs.existsSync(filePath) ? filePath : path.dirname(filePath);
    if (!target || !fs.existsSync(target)) return { success: false, error: 'File not found' };
    if (fs.existsSync(filePath)) {
      shell.showItemInFolder(filePath);
    } else {
      await shell.openPath(target);
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('project-media', (_, data) => {
  currentBackgroundMedia = data || null;
  pushRenderStateToProjection(buildRenderState('media', data, { backgroundMedia: data || null }));
  return { success: true };
});

ipcMain.handle('send-logo-overlay', (_, data) => {
  currentLogoOverlayState = data || null;
  if (projectionWindow && !projectionWindow.isDestroyed()) {
    projectionWindow.webContents.send('logo-overlay', data);
  }
  return { success: true };
});

ipcMain.handle('logo-overlay-drag-update', (_, data) => {
  if (currentLogoOverlayState && data) {
    Object.assign(currentLogoOverlayState, data);
  }
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('logo-overlay-drag-update', data);
  }
  return { success: true };
});

// ── Presets ──────────────────────────────────────────────────────────────────
ipcMain.handle('get-presets', () => {
  try {
    if (fs.existsSync(PRESETS_FILE)) return JSON.parse(fs.readFileSync(PRESETS_FILE, 'utf8'));
  } catch(_) {}
  return [];
});

ipcMain.handle('save-preset', async (_, preset) => {
  try {
    let presets = [];
    try { if (fs.existsSync(PRESETS_FILE)) presets = JSON.parse(fs.readFileSync(PRESETS_FILE, 'utf8')); } catch(_) {}
    if (!Array.isArray(presets)) presets = [];
    const safeId = String(preset.id || '').replace(/[^a-zA-Z0-9_-]/g, '');
    if (!safeId) return { success: false, error: 'Invalid preset id' };
    preset.id = safeId;
    if (preset.logoOverlay?.fileData) {
      const rawExt = path.extname(preset.logoOverlay.fileName || '.png') || '.png';
      const ext = rawExt.replace(/[^a-zA-Z0-9.]/g, '');
      const safeName = `preset_logo_${safeId}${ext}`;
      const dest = path.join(PRESETS_ASSETS, safeName);
      if (!dest.startsWith(PRESETS_ASSETS)) return { success: false, error: 'Invalid path' };
      const buf = Buffer.from(preset.logoOverlay.fileData, 'base64');
      fs.writeFileSync(dest, buf);
      preset.logoOverlay.savedPath = dest;
      delete preset.logoOverlay.fileData;
    } else if (preset.logoOverlay) {
      delete preset.logoOverlay.savedPath;
      const existing = presets.find(p => p.id === safeId);
      if (existing?.logoOverlay?.savedPath) preset.logoOverlay.savedPath = existing.logoOverlay.savedPath;
    }
    const idx = presets.findIndex(p => p.id === safeId);
    if (idx >= 0) presets[idx] = preset;
    else presets.push(preset);
    fs.writeFileSync(PRESETS_FILE, JSON.stringify(presets, null, 2));
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
});

ipcMain.handle('delete-preset', (_, id) => {
  try {
    let presets = [];
    try { if (fs.existsSync(PRESETS_FILE)) presets = JSON.parse(fs.readFileSync(PRESETS_FILE, 'utf8')); } catch(_) {}
    const preset = presets.find(p => p.id === id);
    if (preset?.logoOverlay?.savedPath) {
      const resolved = path.resolve(preset.logoOverlay.savedPath);
      if (resolved.startsWith(PRESETS_ASSETS)) {
        try { fs.unlinkSync(resolved); } catch(_) {}
      }
    }
    presets = presets.filter(p => p.id !== id);
    fs.writeFileSync(PRESETS_FILE, JSON.stringify(presets, null, 2));
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
});

// Show native error dialog for unsupported media formats
ipcMain.handle('show-unsupported-dialog', async (_, { title, message, detail }) => {
  await dialog.showMessageBox(mainWindow, {
    type:    'warning',
    title:   title   || 'Unsupported Format',
    message: message || 'File not imported',
    detail:  detail  || '',
    buttons: ['OK'],
    defaultId: 0,
  });
  return { ok: true };
});


// ── Created presentations (custom slides) — stored in assets/Data/ ─────────────
const CREATED_PRES_FILE = path.join(DATA_DIR, 'created_presentations.json');

ipcMain.handle('get-created-presentations', () => {
  try { return JSON.parse(fs.readFileSync(CREATED_PRES_FILE, 'utf-8')); }
  catch(e) { return []; }
});

ipcMain.handle('save-created-presentations', (_, list) => {
  try {
    fs.writeFileSync(CREATED_PRES_FILE, JSON.stringify(list, null, 2));
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
});

// ── Presentations (PPTX / PDF → slide images) ─────────────────────────────────
const PRES_DIR  = PRES_ASSETS;
const PRES_FILE = path.join(DATA_DIR, 'presentations.json');

function readPresentations() {
  try { return JSON.parse(fs.readFileSync(PRES_FILE, 'utf-8')); }
  catch(e) { return []; }
}
function savePresentations(list) {
  fs.writeFileSync(PRES_FILE, JSON.stringify(list, null, 2));
}

// ── Migrate any existing userData data to assets/Data/ on first run ────────────
// Called after all constants are declared to avoid temporal dead zone errors.
function migrateUserData() {
  const migrate = [
    [path.join(UD, 'settings.json'),      SETTINGS_FILE],
    [path.join(UD, 'transcripts.json'),   TRANSCRIPTS_FILE],
    [path.join(UD, 'themes.json'),        THEMES_FILE],
    [path.join(UD, 'songs.json'),         SONGS_FILE],
    [path.join(UD, 'media.json'),         MEDIA_FILE],
    [path.join(UD, 'presentations.json'), PRES_FILE],
  ];
  for (const [src, dst] of migrate) {
    if (fs.existsSync(src) && !fs.existsSync(dst)) {
      try { fs.copyFileSync(src, dst); } catch(e) { /* best effort */ }
    }
  }
  // Migrate presentation slide images from legacy locations → assets/Presentation/
  for (const oldPresDir of [path.join(UD, 'presentations'), LEGACY_PRES_ASSETS]) {
    if (!fs.existsSync(oldPresDir)) continue;
    const entries = fs.readdirSync(oldPresDir, { withFileTypes: true });
    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      const srcDir = path.join(oldPresDir, entry.name);
      const dstDir = path.join(PRES_ASSETS, entry.name);
      if (!fs.existsSync(dstDir)) {
        try { fs.cpSync(srcDir, dstDir, { recursive: true }); } catch(e) {}
      }
    }
  }

  // Migrate legacy schedules folder into assets/Schedules
  if (fs.existsSync(LEGACY_SCHEDULES_DIR)) {
    try {
      const entries = fs.readdirSync(LEGACY_SCHEDULES_DIR, { withFileTypes: true });
      for (const entry of entries) {
        if (!entry.isFile()) continue;
        const src = path.join(LEGACY_SCHEDULES_DIR, entry.name);
        const dst = path.join(SCHEDULES_DIR, entry.name);
        if (!fs.existsSync(dst)) fs.copyFileSync(src, dst);
      }
    } catch(e) {}
  }

  // Normalize stored presentation metadata to the new assets/Presentation folder
  try {
    const list = readPresentations();
    let changed = false;
    for (const p of list) {
      if (!p?.outDir) continue;
      const base = path.basename(p.outDir);
      const desired = path.join(PRES_ASSETS, base);
      if (p.outDir !== desired && fs.existsSync(desired)) {
        p.outDir = desired;
        changed = true;
      }
    }
    if (changed) savePresentations(list);
  } catch(e) {}
}
migrateUserData();

// Convert a file (PPTX or PDF) into slide PNG images
// Returns { success, slides: [{ index, imagePath }], slideCount, error }
ipcMain.handle('import-presentation', async (_, { filePath }) => {
  const { execFile, spawn } = require('child_process');
  const util = require('util');
  const execFileAsync = util.promisify(execFile);

  const ext    = path.extname(filePath).toLowerCase();
  const id     = Date.now().toString();
  const outDir = path.join(PRES_DIR, id);
  fs.mkdirSync(outDir, { recursive: true });

  try {
    const pyScript = path.join(__dirname, 'pptx_to_png.py');

    // On Windows, first try native PowerPoint automation when Office is installed.
    if (process.platform === 'win32' && ['.pptx','.ppt'].includes(ext)) {
      try {
        const escapedSrc = filePath.replace(/'/g, "''");
        const escapedOut = outDir.replace(/'/g, "''");
        const ps = [
          "$ErrorActionPreference = 'Stop'",
          `$src='${escapedSrc}'`,
          `$outDir='${escapedOut}'`,
          "$ppt = New-Object -ComObject PowerPoint.Application",
          "$ppt.Visible = 1",
          "$pres = $ppt.Presentations.Open($src, $false, $false, $false)",
          "$pres.Export($outDir, 'PNG', 1920, 1080)",
          "$pres.Close()",
          "$ppt.Quit()",
          "Write-Output 'ok'"
        ].join('; ');
        await execFileAsync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps], { timeout: 240000, maxBuffer: 4 * 1024 * 1024 });
        const files = fs.readdirSync(outDir)
          .filter(f => /(^Slide\d+\.PNG$)|(^slide\d+\.png$)|(^\d+\.png$)/i.test(f))
          .sort((a,b) => {
            const na = parseInt(a.match(/(\d+)/)?.[1]||'0');
            const nb = parseInt(b.match(/(\d+)/)?.[1]||'0');
            return na - nb;
          });
        if (files.length) {
          const slides = [];
          files.forEach((f,i) => {
            const srcP = path.join(outDir, f);
            const normalized = path.join(outDir, `slide-${i+1}.png`);
            if (srcP !== normalized) {
              try { fs.copyFileSync(srcP, normalized); } catch(_) {}
            }
            slides.push({ index:i, imagePath: normalized });
          });
          const name   = path.basename(filePath, ext);
          const list   = readPresentations();
          list.push({ id, name, filePath, slideCount: slides.length, outDir, importedAt: Date.now() });
          savePresentations(list);
          return { success: true, id, name, slides, slideCount: slides.length, outDir };
        }
      } catch(e) {
        console.warn('[Presentation] PowerPoint automation failed, falling back:', e.message);
      }
    }

    // Try LibreOffice / soffice direct PDF export before Python fallback
    if (['.pptx','.ppt','.odp'].includes(ext)) {
      const sofficeCandidates = process.platform === 'win32'
        ? [
            'C:/Program Files/LibreOffice/program/soffice.exe',
            'C:/Program Files (x86)/LibreOffice/program/soffice.exe',
            'soffice.exe',
            'soffice'
          ]
        : ['/usr/bin/libreoffice', '/usr/bin/soffice', 'libreoffice', 'soffice'];
      for (const soffice of sofficeCandidates) {
        try {
          const tmpPdfDir = path.join(outDir, 'pdf-export');
          fs.mkdirSync(tmpPdfDir, { recursive: true });
          await execFileAsync(soffice, ['--headless', '--convert-to', 'pdf', '--outdir', tmpPdfDir, filePath], { timeout: 180000, maxBuffer: 2 * 1024 * 1024 });
          const pdfName = path.basename(filePath, ext) + '.pdf';
          const pdfPath = path.join(tmpPdfDir, pdfName);
          if (fs.existsSync(pdfPath)) {
            const slidePrefix = path.join(outDir, 'slide');
            await execFileAsync('pdftoppm', ['-png', '-r', '150', pdfPath, slidePrefix], { timeout: 120000 });
            const files = fs.readdirSync(outDir)
              .filter(f => f.startsWith('slide') && f.endsWith('.png'))
              .sort((a,b) => {
                const na = parseInt(a.match(/(\d+)/)?.[1]||'0');
                const nb = parseInt(b.match(/(\d+)/)?.[1]||'0');
                return na - nb;
              });
            if (files.length) {
              const slides = files.map((f,i) => ({ index:i, imagePath: path.join(outDir,f) }));
              const name   = path.basename(filePath, ext);
              const list   = readPresentations();
              list.push({ id, name, filePath, slideCount: slides.length, outDir, importedAt: Date.now() });
              savePresentations(list);
              return { success: true, id, name, slides, slideCount: slides.length, outDir };
            }
          }
        } catch (e) {}
      }
    }

    // Try multiple Python executables
    const pythonCandidates = process.platform === 'win32'
      ? ['python', 'python3', 'py']
      : ['python3', 'python'];

    let result = null;
    let lastErr = '';

    for (const py of pythonCandidates) {
      try {
        const { stdout } = await execFileAsync(py, [pyScript, filePath, outDir],
          { timeout: 180000, maxBuffer: 2 * 1024 * 1024 });
        const parsed = JSON.parse(stdout.trim());
        if (parsed.success || parsed.error) { result = parsed; break; }
      } catch(e) {
        lastErr = e.message;
        // Continue to next candidate
      }
    }

    // If Python completely unavailable, try direct pdftoppm for PDFs
    if (!result) {
      if (ext === '.pdf') {
        const slidePrefix = path.join(outDir, 'slide');
        await execFileAsync('pdftoppm', ['-png', '-r', '150', filePath, slidePrefix],
          { timeout: 120000 });
        const files = fs.readdirSync(outDir)
          .filter(f => f.startsWith('slide') && f.endsWith('.png'))
          .sort((a,b) => {
            const na = parseInt(a.match(/(\d+)/)?.[1]||'0');
            const nb = parseInt(b.match(/(\d+)/)?.[1]||'0');
            return na - nb;
          });
        if (!files.length) {
          fs.rmSync(outDir, { recursive: true, force: true });
          return { success: false, error: 'No slides generated from PDF' };
        }
        result = {
          success: true,
          slides: files.map((f,i) => ({ index:i, imagePath: path.join(outDir,f) })),
          count: files.length
        };
      } else {
        fs.rmSync(outDir, { recursive: true, force: true });
        return { success: false, error: `Python not found. Install Python 3 and run: pip install python-pptx Pillow\n\nDetails: ${lastErr}` };
      }
    }

    if (!result.success) {
      fs.rmSync(outDir, { recursive: true, force: true });
      return { success: false, error: result.error || 'Conversion failed' };
    }

    const slides = result.slides;
    const name   = path.basename(filePath, ext);
    const list   = readPresentations();
    list.push({ id, name, filePath, slideCount: slides.length, outDir, importedAt: Date.now() });
    savePresentations(list);

    return { success: true, id, name, slides, slideCount: slides.length, outDir };
  } catch(e) {
    try { fs.rmSync(outDir, { recursive: true, force: true }); } catch(_) {}
    return { success: false, error: e.message };
  }
});

ipcMain.handle('get-presentations', () => {
  const list = readPresentations().map(p => {
    if (!p?.outDir) return p;
    if (fs.existsSync(p.outDir)) return p;
    const repaired = path.join(PRES_ASSETS, path.basename(p.outDir));
    return fs.existsSync(repaired) ? { ...p, outDir: repaired } : p;
  });
  return list.filter(p => fs.existsSync(p.outDir));
});

ipcMain.handle('get-presentation-slides', (_, { id }) => {
  const list = readPresentations();
  const pres = list.find(p => p.id === id);
  if (!pres || !fs.existsSync(pres.outDir)) return { success: false, slides: [] };
  const slides = fs.readdirSync(pres.outDir)
    .filter(f => f.startsWith('slide') && f.endsWith('.png'))
    .sort((a, b) => {
      const na = parseInt(a.match(/(\d+)/)?.[1] || '0');
      const nb = parseInt(b.match(/(\d+)/)?.[1] || '0');
      return na - nb;
    })
    .map((f, i) => ({ index: i, imagePath: path.join(pres.outDir, f) }));
  return { success: true, slides, name: pres.name, slideCount: slides.length };
});

ipcMain.handle('delete-presentation', (_, { id }) => {
  let list = readPresentations();
  const pres = list.find(p => p.id === id);
  if (pres) {
    try { fs.rmSync(pres.outDir, { recursive: true, force: true }); } catch(_) {}
    list = list.filter(p => p.id !== id);
    savePresentations(list);
  }
  return { success: true };
});

ipcMain.handle('project-presentation-slide', (_, { imagePath }) => {
  const payload = { imagePath };
  pushRenderStateToProjection(buildRenderState('presentation-slide', payload));
  if (projectionWindow) {
    projectionWindow.webContents.send('show-presentation-slide', payload);
  }
  return { success: true };
});

ipcMain.handle('project-created-slide', (_, { slide }) => {
  const payload = { slide };
  pushRenderStateToProjection(buildRenderState('created-slide', payload));
  if (projectionWindow) {
    projectionWindow.webContents.send('show-created-slide', payload);
  }
  return { success: true };
});

// ── Timer state (tracked so countdown window can sync on open) ────────────────
let _activeTimer = null; // { data, startedAt, seconds, mode }

ipcMain.handle('show-timer', (_, data) => {
  // Block start if projection is not open — timer needs a screen to display on
  const projOpen = !!(projectionWindow && !projectionWindow.isDestroyed());
  if (!projOpen) {
    return { success: false, projected: false, error: 'Projection screen is not open. Click Projection to open it first.' };
  }

  const startedAt = Date.now();
  _activeTimer = {
    data,
    startedAt,
    seconds:   data.seconds || 0,
    mode:      data.mode    || 'countdown',
    label:     data.label   || '',
    scale:     data.scale   || 1,
    position:  data.position || 'edge',
  };

  const livePayload = {
    ...data,
    seconds: _activeTimer.seconds,
    startedAt,
    mode: _activeTimer.mode,
    label: _activeTimer.label || data.label || '',
    scale: _activeTimer.scale,
    position: _activeTimer.position,
    serverNow: Date.now(),
  };

  const projected = !!(projectionWindow && !projectionWindow.isDestroyed());
  if (projected) projectionWindow.webContents.send('show-timer', livePayload);

  const syncPayload = {
    ..._activeTimer,
    running: true,
    remaining: _activeTimer.seconds,
    startedAt,
    serverNow: Date.now(),
  };
  if (countdownWindow && !countdownWindow.isDestroyed()) {
    countdownWindow.webContents.send('timer-state-sync', syncPayload);
  }
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('timer-state-sync', syncPayload);
  }

  return { success: true, projected };
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
    running:   true,
    data:      _activeTimer.data,
    mode:      _activeTimer.mode,
    label:     _activeTimer.label,
    scale:     _activeTimer.scale,
    position:  _activeTimer.position,
    seconds:   _activeTimer.seconds,
    remaining: remaining,
    startedAt: _activeTimer.startedAt,
    serverNow: Date.now(),
  };
});

ipcMain.handle('timer-scale', (_, data) => {
  if (_activeTimer) { _activeTimer.scale = data.scale; _activeTimer.position = data.position; }
  if (projectionWindow && !projectionWindow.isDestroyed()) projectionWindow.webContents.send('timer-scale', data);
  if (countdownWindow && !countdownWindow.isDestroyed()) countdownWindow.webContents.send('timer-scale-sync', data);
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('timer-scale-sync', data);
  return { success: true };
});

// Flash speed control — sent by timer control window's Flash Faster button
// Forwards to projection so the blink animates faster while button is held
ipcMain.handle('timer-flash-speed', (_, data) => {
  if (projectionWindow && !projectionWindow.isDestroyed()) {
    projectionWindow.webContents.send('timer-flash-speed', data);
  }
  return { success: true };
});

// Tells countdown-window.html whether it's running in standalone timer mode
ipcMain.handle('is-timer-standalone', () => false);

// Clock / Date display on projection
let _activeClockData = null; // store current clock state for projection reopen sync

ipcMain.handle('show-clock', (_, data) => {
  _activeClockData = data;
  if (projectionWindow && !projectionWindow.isDestroyed())
    projectionWindow.webContents.send('show-clock', data);
  return { success: true };
});
ipcMain.handle('hide-clock', () => {
  _activeClockData = null;
  if (projectionWindow && !projectionWindow.isDestroyed())
    projectionWindow.webContents.send('hide-clock');
  return { success: true };
});
ipcMain.handle('get-clock-state', () => _activeClockData || null);

// Set projection background — solid color, image, or video
// Used by the standalone timer and greyed-out in main app
ipcMain.handle('set-projection-bg', (_, data) => {
  const win = projectionWindow;
  if (win && !win.isDestroyed()) {
    win.webContents.send('set-projection-bg', data);
  }
  return { success: true };
});

// File picker for bg — only used by standalone timer (greyed out in main app)
ipcMain.handle('pick-bg-file', async (_, type) => {
  const filters = type === 'image'
    ? [{ name: 'Images', extensions: ['jpg','jpeg','png','webp','gif'] }]
    : [{ name: 'Videos', extensions: ['mp4','webm','mov','mkv'] }];
  const res = await dialog.showOpenDialog(mainWindow, { filters, properties: ['openFile'] });
  if (res.canceled || !res.filePaths.length) return null;
  return res.filePaths[0];
});

ipcMain.handle('stop-timer', () => {
  _activeTimer = null;
  if (projectionWindow) projectionWindow.webContents.send('stop-timer');
  // Notify countdown window if open
  if (countdownWindow && !countdownWindow.isDestroyed()) {
    countdownWindow.webContents.send('timer-stopped');
  }
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('timer-stopped');
  }
  return { success: true };
});

ipcMain.handle('show-caption', (_, data) => {
  if (projectionWindow) projectionWindow.webContents.send('show-caption', data);
  return { success: true };
});

ipcMain.handle('show-alert', (_, data) => {
  if (projectionWindow) projectionWindow.webContents.send('show-alert', data);
  return { success: true };
});

// ── Schedule (save/load service plans) ────────────────────────────────────────
// Save schedule — shows input dialog to name it, saves to schedules dir
ipcMain.handle('save-schedule', async (_, { name, items }) => {
  try{
    fs.mkdirSync(SCHEDULES_DIR, { recursive: true });
    const safeName = (name||'schedule').replace(/[<>:"/\\|?*]/g,'_').trim() || 'schedule';
    const filePath = path.join(SCHEDULES_DIR, `${safeName}.acsch`);
    fs.writeFileSync(filePath, JSON.stringify({ name: name||safeName, items, savedAt: Date.now() }, null, 2));
    _touchRecentSchedule(filePath, name||safeName);
    return { success: true, path: filePath, name: name||safeName };
  } catch(e){ return { success: false, error: e.message }; }
});

// Save-as: native Save dialog so user can pick filename and location
ipcMain.handle('save-schedule-as', async (_, { items, currentName }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title:       'Save Schedule As',
    defaultPath: path.join(SCHEDULES_DIR, (currentName || 'My Service') + '.acsch'),
    filters:     [{ name: 'AnchorCast Schedule', extensions: ['acsch','json'] }],
    buttonLabel: 'Save Schedule',
  });
  if (result.canceled || !result.filePath) return { canceled: true };
  try{
    const name = path.basename(result.filePath, path.extname(result.filePath));
    fs.mkdirSync(path.dirname(result.filePath), { recursive: true });
    fs.writeFileSync(result.filePath, JSON.stringify({ name, items, savedAt: Date.now() }, null, 2));
    _touchRecentSchedule(result.filePath, name);
    return { success: true, path: result.filePath, name };
  } catch(e){ return { success: false, error: e.message }; }
});

// Open schedule: native Open dialog
ipcMain.handle('open-schedule-dialog', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title:       'Open Schedule',
    defaultPath: SCHEDULES_DIR,
    filters:     [{ name: 'AnchorCast Schedule', extensions: ['acsch','json'] }],
    properties:  ['openFile'],
    buttonLabel: 'Open Schedule',
  });
  if (result.canceled || !result.filePaths.length) return { canceled: true };
  try{
    const data = JSON.parse(fs.readFileSync(result.filePaths[0], 'utf-8'));
    _touchRecentSchedule(result.filePaths[0], data.name);
    return { success: true, schedule: { name: data.name, items: data.items||[], savedAt: data.savedAt } };
  } catch(e){ return { success: false, error: e.message }; }
});

// Load a specific schedule file by filename (used by Recent Schedules menu)
ipcMain.handle('consume-pending-schedule-open', () => {
  const filePath = pendingScheduleOpenPath || null;
  pendingScheduleOpenPath = null;
  return filePath;
});

ipcMain.handle('load-schedule-file', async (_, filename) => {
  let filePath;
  if (path.isAbsolute(filename)) {
    filePath = path.resolve(filename);
    if (!_looksLikeSchedulePath(filePath)) return { success: false, error: 'Invalid schedule file' };
  } else {
    const safeName = path.basename(String(filename || ''));
    if (!safeName) return { success: false, error: 'Invalid filename' };
    filePath = path.resolve(path.join(SCHEDULES_DIR, safeName));
    if (!filePath.startsWith(path.resolve(SCHEDULES_DIR) + path.sep)) return { success: false, error: 'Invalid path' };
  }
  try{
    const data = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    _touchRecentSchedule(filePath, data.name);
    return { success: true, schedule: { name: data.name, items: data.items||[], savedAt: data.savedAt } };
  } catch(e){ return { success: false, error: e.message }; }
});

ipcMain.handle('load-schedules', () => {
  try{
    if(!fs.existsSync(SCHEDULES_DIR)) return [];
    return fs.readdirSync(SCHEDULES_DIR)
      .filter(f => /\.(acsch|json)$/i.test(f))
      .map(f => {
        try{
          const data = JSON.parse(fs.readFileSync(path.join(SCHEDULES_DIR,f),'utf-8'));
          return { name: data.name || f.replace('.json',''), items: data.items || [], savedAt: data.savedAt, file: f };
        } catch(e){ return null; }
      })
      .filter(Boolean)
      .sort((a,b) => (b.savedAt||0) - (a.savedAt||0));
  } catch(e){ return []; }
});

ipcMain.handle('delete-schedule', (_, filename) => {
  try{
    const safeName = path.basename(String(filename || ''));
    if (!safeName) return { success: false, error: 'Invalid filename' };
    const target = path.resolve(path.join(SCHEDULES_DIR, safeName));
    if (!target.startsWith(path.resolve(SCHEDULES_DIR) + path.sep)) return { success: false, error: 'Invalid path' };
    fs.unlinkSync(target);
    _refreshRecentSchedulesMenu();
    return { success: true };
  } catch(e){ return { success: false }; }
});
ipcMain.handle('open-settings',(_, opts = {})=>{
  _settingsOpenData = opts || null;
  createSettingsWindow(opts?.section || '');
  return{success:true};
});
ipcMain.handle('get-settings-open-params', () => {
  const s = _settingsStartSection || '';
  _settingsStartSection = '';
  return s ? { section: s } : null;
});
// Theme designer open params — stored so renderer can query after init
let _themeDesignerOpenData = null;

ipcMain.handle('open-theme-designer',(_, data)=>{
  _themeDesignerOpenData = data || null;
  createThemeWindow(data);
  return{success:true};
});

ipcMain.handle('get-theme-designer-params', () => {
  const d = _themeDesignerOpenData;
  _themeDesignerOpenData = null; // consume once
  return d;
});
ipcMain.handle('show-remote-url',()=>{ showRemoteUrl(); return{success:true}; });
ipcMain.handle('open-external',(_, url)=>{ shell.openExternal(url); return{success:true}; });
ipcMain.handle('copy-to-clipboard', (_, text) => { try { clipboard.writeText(String(text || '')); return { success:true }; } catch (e) { return { success:false, error: e?.message || 'Copy failed' }; } });

// Remote server info and toggle
ipcMain.handle('get-remote-status',()=>({
  enabled: (currentSettings.remoteEnabled !== false) || (httpServer !== null),
  running: httpServer !== null,
  ip: localIp(),
  port: currentSettings.httpPort||8080,
  url: `http://${localIp()}:${currentSettings.httpPort||8080}/remote`,
  authRequired: isRemoteAuthRequired(currentSettings.remoteRequireAuth),
  selectedAdapter: currentSettings.networkAdapter || null,
  adapters: getAllNetworkAdapters(),
  roleLinks: makeRoleRemoteLinks(),
}));
ipcMain.handle('get-remote-info',()=>({
  ip:      localIp(),
  port:    currentSettings.httpPort||8080,
  enabled: (currentSettings.remoteEnabled !== false) || (httpServer !== null),
  running: httpServer !== null,
  authRequired: isRemoteAuthRequired(currentSettings.remoteRequireAuth),
  selectedAdapter: currentSettings.networkAdapter || null,
  adapters: getAllNetworkAdapters(),
  roleLinks: makeRoleRemoteLinks(),
}));
ipcMain.handle('get-network-adapters', () => getAllNetworkAdapters());
ipcMain.handle('set-network-adapter', async (_, adapterName) => {
  currentSettings.networkAdapter = adapterName || null;
  try{
    fs.mkdirSync(path.dirname(SETTINGS_FILE),{recursive:true});
    fs.writeFileSync(SETTINGS_FILE,JSON.stringify(currentSettings,null,2));
  }catch(e){}
  // Restart server so it reports the correct IP
  if(currentSettings.remoteEnabled !== false){
    startHttpServer(currentSettings.httpPort||8080);
  }
  return { success:true, ip: localIp() };
});

ipcMain.handle('start-remote', async ()=>{
  const gate = requiresRegistration('Remote Control');
  if (gate) return { error: gate.reason, blocked: true };
  currentSettings.remoteEnabled = true;
  try{ fs.mkdirSync(path.dirname(SETTINGS_FILE),{recursive:true}); fs.writeFileSync(SETTINGS_FILE,JSON.stringify(currentSettings,null,2)); }catch(e){}
  startHttpServer(currentSettings.httpPort||8080);
  return { success:true, enabled:true, ip:localIp(), port:currentSettings.httpPort||8080 };
});
ipcMain.handle('stop-remote', async ()=>{
  currentSettings.remoteEnabled = false;
  try{ fs.mkdirSync(path.dirname(SETTINGS_FILE),{recursive:true}); fs.writeFileSync(SETTINGS_FILE,JSON.stringify(currentSettings,null,2)); }catch(e){}
  stopHttpServer();
  mainWindow?.webContents.send('http-server-started',{ip:null,disabled:true});
  return { success:true, enabled:false, ip:null, port:currentSettings.httpPort||8080 };
});
ipcMain.handle('toggle-remote', async (_, enable)=>{
  if (enable) {
    const gate = requiresRegistration('Remote Control');
    if (gate) return { error: gate.reason, blocked: true };
  }
  currentSettings.remoteEnabled = enable;
  try{
    fs.mkdirSync(path.dirname(SETTINGS_FILE),{recursive:true});
    fs.writeFileSync(SETTINGS_FILE,JSON.stringify(currentSettings,null,2));
  }catch(e){}
  if(enable){
    startHttpServer(currentSettings.httpPort||8080);
  } else {
    stopHttpServer();
    mainWindow?.webContents.send('http-server-started',{ip:null,disabled:true});
  }
  return{ success:true, enabled:enable, ip:enable?localIp():null, port:currentSettings.httpPort||8080 };
});

ipcMain.handle('open-adaptive-management',()=>{createAdaptiveManagementWindow();return{success:true};});



ipcMain.handle('get-remote-runtime-status', () => {
  const lastSeenAt = Number(remoteRuntimeStatus.lastSeenAt || 0);
  const connected = lastSeenAt > 0 && (Date.now() - lastSeenAt) < 90000;
  return {
    connected,
    lastSeenAt: lastSeenAt || null,
    lastRole: remoteRuntimeStatus.lastRole || null,
    lastIp: remoteRuntimeStatus.lastIp || null,
    serverEnabled: currentSettings.remoteEnabled !== false
  };
});
ipcMain.handle('load-detection-review-data', () => ({
  phrases: readDetectionJson(DETECTION_PHRASES_FILE, []),
  events: readDetectionJson(DETECTION_EVENTS_FILE, []),
}));
ipcMain.handle('save-detection-feedback', (_, payload = {}) => {
  try{
    const events = readDetectionJson(DETECTION_EVENTS_FILE, []);
    events.push({
      id: payload.id || `det_evt_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      detectionId: payload.detectionId || null,
      ref: payload.ref || '',
      sourceText: payload.sourceText || '',
      action: payload.action || 'review',
      method: payload.method || '',
      confidence: Number(payload.confidence || 0),
      accepted: payload.accepted == null ? null : (payload.accepted ? 1 : 0),
      createdAt: new Date().toISOString(),
    });
    writeDetectionJson(DETECTION_EVENTS_FILE, events.slice(-3000));
    return { success:true };
  }catch(e){ return { success:false, error:e.message }; }
});

// License IPC handlers
// ── Registration IPC handlers ────────────────────────────────────────────────
ipcMain.handle('get-registration-status', () => getRegistrationStatus());
ipcMain.handle('get-license-status', () => getRegistrationStatus()); // legacy compat
ipcMain.handle('get-hardware-id', () => getHardwareId());
ipcMain.handle('send-registration-email', async (_, fullName, email, churchName) => {
  try {
    if (!fullName || !String(fullName).trim()) return { success:false, error:'Full name is required.' };
    if (!email || !String(email).includes('@')) return { success:false, error:'Valid email required.' };
    const result = await sendRegistrationEmail(String(fullName).trim(), String(email).trim(), String(churchName||'').trim());
    if (result.success) {
      // Persist pending email state so app shows success screen on relaunch
      saveEmailSentState(String(fullName).trim(), String(email).trim(), String(churchName||'').trim());
    }
    return result;
  } catch(e) {
    console.error('[Registration] Email error:', e.message);
    return { success:false, error: e.message || 'Failed to send email. Check your internet connection.' };
  }
});
ipcMain.handle('get-email-sent-state', () => getEmailSentState());
ipcMain.handle('activate-registration', (_, fullName, email, token, churchName) => {
  return activateRegistration(fullName, email, token, churchName);
});

ipcMain.handle('open-bible-manager', () => {
  createBibleManagerWindow();
  return { success:true };
});
ipcMain.handle('open-presentation-editor', (_, payload = {}) => {
  createPresentationEditorWindow(payload || {});
  return { success:true };
});
ipcMain.handle('open-countdown-window', () => {
  createCountdownWindow();
  return { success:true };
});

// BUG-G FIX: parseLicenseImportFile and getAndClearPendingLicenseImport were
// called by IPC handlers but never defined — caused crash on any license import.
let _pendingLicenseImportPath = null;

function getAndClearPendingLicenseImport() {
  const p = _pendingLicenseImportPath;
  _pendingLicenseImportPath = null;
  return p;
}

function parseLicenseImportFile(filePath) {
  try {
    if (!filePath || !fs.existsSync(filePath)) {
      return { success: false, error: 'License file not found.' };
    }
    const raw = fs.readFileSync(filePath, 'utf-8');
    let data;
    try { data = JSON.parse(raw); } catch (_) {
      return { success: false, error: 'License file is not valid JSON.' };
    }
    const fullName   = String(data.fullName   || data.name    || '').trim();
    const email      = String(data.email      || '').trim();
    const token      = String(data.token      || data.key     || '').trim();
    const churchName = String(data.churchName || data.church  || '').trim();
    const hardwareId = String(data.hwId       || data.hardwareId || '').trim();
    if (!email || !token) {
      return { success: false, error: 'License file is missing email or token.' };
    }
    return { success: true, fullName, email, token, churchName, hardwareId, key: token };
  } catch (e) {
    return { success: false, error: e.message || 'Could not parse license file.' };
  }
}
ipcMain.handle('open-license-window', () => {
  createRegistrationStatusWindow();
  return { success:true };
});

ipcMain.handle('import-license-file', async () => {
  try {
    const { canceled, filePaths } = await dialog.showOpenDialog({
      title: 'Select AnchorCast license file',
      properties: ['openFile'],
      filters: [
        { name: 'AnchorCast License Files', extensions: ['json', 'license', 'anchorcast-license'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });
    if (canceled || !filePaths || !filePaths[0]) return { success:false, canceled:true };
    return parseLicenseImportFile(filePaths[0]);
  } catch (e) {
    return { success:false, error:e.message || 'Could not read license file.' };
  }
});
ipcMain.handle('consume-pending-license-import', () => {
  const p = getAndClearPendingLicenseImport();
  if (!p) return { success:false, empty:true };
  return parseLicenseImportFile(p);
});


ipcMain.handle('save-detection-phrase', (_, payload = {}) => {
  try{
    const phrases = readDetectionJson(DETECTION_PHRASES_FILE, []);
    const phrase = String(payload.phrase || '').trim().toLowerCase();
    const ref = String(payload.ref || '').trim();
    if(!phrase || !ref) return { success:false, error:'Missing phrase or ref' };
    const idx = phrases.findIndex(p => String(p.phrase || '').toLowerCase() === phrase && String(p.ref || '').toLowerCase() === ref.toLowerCase());
    const row = {
      id: idx >= 0 ? phrases[idx].id : `det_phrase_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      phrase,
      ref,
      type: payload.type || 'learned',
      createdBy: payload.createdBy || 'user',
      usageCount: idx >= 0 ? Number(phrases[idx].usageCount || 0) + 1 : 1,
      createdAt: idx >= 0 ? phrases[idx].createdAt : new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    if(idx >= 0) phrases[idx] = { ...phrases[idx], ...row };
    else phrases.push(row);
    writeDetectionJson(DETECTION_PHRASES_FILE, phrases);
    return { success:true, count: phrases.length };
  }catch(e){ return { success:false, error:e.message }; }
});

} // end timer mode else block

#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const root = path.resolve(__dirname, '..');
const isWin = process.platform === 'win32' || process.argv.includes('--win');

function run(cmd, args, opts = {}) {
  const pretty = `${cmd} ${args.join(' ')}`;
  console.log(`[prepare-build] ${pretty}`);
  const res = spawnSync(cmd, args, {
    cwd: root,
    stdio: 'inherit',
    shell: false,
    windowsHide: false,
    env: { ...process.env, ANCHORCAST_NONINTERACTIVE: '1', ...(opts.env || {}) },
    ...opts,
  });
  if (res.error) return { ok: false, code: 1, error: res.error };
  return { ok: (res.status || 0) === 0, code: res.status || 0 };
}
function exists(p){ try { return fs.existsSync(p); } catch(_) { return false; } }

function ensureWhisperPrepared() {
  if (!isWin) return true;

  // ── Python check — OPTIONAL, warn only ───────────────────────────────────
  // If python/ folder is missing, the installer will run setup_whisper.bat
  // which downloads portable Python 3.12 and faster-whisper on the user's machine.
  const pyExe = path.join(root, 'python', 'python.exe');
  if (!exists(pyExe)) {
    console.warn('[prepare-build] ⚠ python\\python.exe not found — Python will NOT be bundled.');
    console.warn('[prepare-build]   The installer will run setup_whisper.bat on the user machine.');
    console.warn('[prepare-build]   To bundle Python: run setup_whisper.bat first, then rebuild.');
  } else {
    const check = spawnSync(pyExe, ['-c', 'from faster_whisper import WhisperModel'], {
      encoding: 'utf8', timeout: 8000, windowsHide: true
    });
    if (check.status !== 0) {
      console.warn('[prepare-build] ⚠ faster-whisper not found in bundled Python.');
      console.warn('[prepare-build]   Run setup_whisper.bat to install it before building.');
    } else {
      console.log('[prepare-build] ✓ Bundled Python + faster-whisper: OK');
    }
  }

  // ── VC++ Redistributable check — OPTIONAL, warn only ────────────────────
  const vcRedist = path.join(root, 'vc_redist.x64.exe');
  if (!exists(vcRedist)) {
    console.warn('[prepare-build] ⚠ vc_redist.x64.exe not found in project root.');
    console.warn('[prepare-build]   Download from: https://aka.ms/vs/17/release/vc_redist.x64.exe');
    console.warn('[prepare-build]   Without it, installer will try to download VC++ at install time.');
  } else {
    const size = Math.round(fs.statSync(vcRedist).size / 1024 / 1024);
    console.log(`[prepare-build] ✓ vc_redist.x64.exe found (${size} MB) — will be bundled.`);
  }

  // ── Models check — OPTIONAL, warn only ───────────────────────────────────
  // If models/ folder is missing or empty, the installer will offer to
  // download the model on first run. No need to block the build.
  const modelsDir = path.join(root, 'models');
  if (!exists(modelsDir)) {
    console.warn('[prepare-build] ⚠ models/ folder not found — no models will be bundled.');
    console.warn('[prepare-build]   Installer will offer to download (~244 MB) on first run.');
    console.warn('[prepare-build]   To bundle: copy model folders into models/ and rebuild.');
  } else {
    const modelFolders = fs.readdirSync(modelsDir).filter(f => {
      try { return fs.statSync(path.join(modelsDir, f)).isDirectory(); } catch { return false; }
    });
    if (modelFolders.length === 0) {
      console.warn('[prepare-build] ⚠ models/ folder is empty — installer will offer download on first run.');
    } else {
      console.log(`[prepare-build] ✓ Bundled models (${modelFolders.length}): ${modelFolders.join(', ')}`);
    }
  }

  return true;
}

function ensureNdiBuilt() {
  if (process.env.SKIP_NDI_BUILD === '1') return true;
  const isMac = process.platform === 'darwin' || process.argv.includes('--mac');
  const ndiAddonDir = path.join(root, 'ndi-addon');
  const nodeFile = path.join(ndiAddonDir, 'build', 'Release', 'ndi_sender.node');

  if (isWin) {
    const dllFile = path.join(ndiAddonDir, 'build', 'Release', 'Processing.NDI.Lib.x64.dll');
    if (exists(nodeFile) && exists(dllFile)) {
      console.log('[prepare-build] NDI addon already built.');
      return true;
    }
    const sdkHeader = 'C:\\Program Files\\NDI\\NDI 6 SDK\\Include\\Processing.NDI.Lib.h';
    if (!exists(sdkHeader)) {
      console.warn('[prepare-build] NDI 6 SDK not found. Skipping NDI build (MJPEG fallback only).');
      return true;
    }
  } else if (isMac) {
    const dylibFile = path.join(ndiAddonDir, 'build', 'Release', 'libndi.dylib');
    if (exists(nodeFile)) {
      console.log('[prepare-build] NDI addon already built.');
      return true;
    }
    const sdkHeader = '/Library/NDI SDK for Apple/include/Processing.NDI.Lib.h';
    if (!exists(sdkHeader)) {
      console.warn('[prepare-build] NDI SDK for Apple not found. Skipping NDI build (MJPEG fallback only).');
      return true;
    }
  } else {
    console.log('[prepare-build] NDI addon build not configured for this platform, skipping.');
    return true;
  }

  let electronVer = '';
  try {
    const epkg = path.join(root, 'node_modules', 'electron', 'package.json');
    if (exists(epkg)) electronVer = JSON.parse(fs.readFileSync(epkg, 'utf8')).version || '';
  } catch (_) {}

  console.log('[prepare-build] Installing ndi-addon dependencies...');
  run('npm', ['install', '--ignore-scripts'], { cwd: ndiAddonDir });

  const arch = isMac ? (process.arch || 'x64') : 'x64';
  if (electronVer) {
    console.log(`[prepare-build] Building NDI addon for Electron ${electronVer} (${arch})...`);
    const res = run('node-gyp', ['rebuild', `--target=${electronVer}`, `--arch=${arch}`, '--dist-url=https://electronjs.org/headers'], { cwd: ndiAddonDir });
    if (!res.ok) {
      console.warn('[prepare-build] NDI addon build (Electron target) failed. Trying system Node fallback...');
      const res2 = run('node-gyp', ['rebuild'], { cwd: ndiAddonDir });
      if (!res2.ok) {
        console.warn('[prepare-build] NDI addon build failed entirely. Build will continue with MJPEG fallback only.');
        return false;
      }
    }
  } else {
    console.log('[prepare-build] Building NDI addon for system Node.js...');
    if (isWin) {
      const buildBat = path.join(ndiAddonDir, 'build-ndi.bat');
      if (!exists(buildBat)) { console.warn('[prepare-build] build-ndi.bat not found, skipping.'); return true; }
      const res = run('cmd.exe', ['/c', buildBat], { cwd: ndiAddonDir });
      if (!res.ok) { console.warn('[prepare-build] NDI build failed. MJPEG fallback only.'); return false; }
    } else {
      const buildSh = path.join(ndiAddonDir, 'build-ndi.sh');
      if (!exists(buildSh)) { console.warn('[prepare-build] build-ndi.sh not found, skipping.'); return true; }
      const res = run('bash', [buildSh], { cwd: ndiAddonDir });
      if (!res.ok) { console.warn('[prepare-build] NDI build failed. MJPEG fallback only.'); return false; }
    }
  }
  return true;
}

ensureWhisperPrepared();
ensureNdiBuilt();
console.log('[prepare-build] Done.');

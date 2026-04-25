const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const http = require('http');
const { EventEmitter } = require('events');

const app = express();
const PORT = 5000;

const emitter = new EventEmitter();

const DATA_DIR = path.join(os.homedir(), '.anchorcast');
const PRESETS_ASSETS = path.join(DATA_DIR, 'PresetAssets');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(PRESETS_ASSETS)) fs.mkdirSync(PRESETS_ASSETS, { recursive: true });

const SETTINGS_FILE = path.join(DATA_DIR, 'settings.json');
const TRANSCRIPTS_FILE = path.join(DATA_DIR, 'transcripts.json');
const THEMES_FILE = path.join(DATA_DIR, 'themes.json');
const SONGS_FILE = path.join(DATA_DIR, 'songs.json');
const MEDIA_FILE = path.join(DATA_DIR, 'media.json');
const BIBLE_VERSIONS_DIR = path.join(DATA_DIR, 'bible');
if (!fs.existsSync(BIBLE_VERSIONS_DIR)) fs.mkdirSync(BIBLE_VERSIONS_DIR, { recursive: true });

function readJSON(file, def = {}) {
  try {
    if (fs.existsSync(file)) return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (e) {}
  return def;
}
function writeJSON(file, data) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
}

let currentRenderState = { module: 'clear', payload: null, updatedAt: 0, backgroundMedia: null };

app.use(express.json({ limit: '50mb' }));

app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/data', express.static(path.join(__dirname, 'data')));
app.use('/preset-assets', express.static(PRESETS_ASSETS));

app.use(express.static(path.join(__dirname, 'src', 'renderer'), {
  setHeaders: (res) => {
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
  }
}));

app.get('/api/settings', (req, res) => {
  const defaults = {
    whisperSource: 'local',
    onlineMode: false,
    translation: 'KJV',
    theme: 'sanctuary',
    remoteEnabled: false,
    overlayTextOnMedia: false,
    songTheme: null,
    presTheme: null,
    claudeApiKey: '',
    deepgramApiKey: '',
    geniusApiKey: '',
    hideGetStarted: false,
  };
  const settings = Object.assign({}, defaults, readJSON(SETTINGS_FILE));
  res.json(settings);
});

app.post('/api/settings', (req, res) => {
  const current = readJSON(SETTINGS_FILE, {});
  const hasWrapped = req.body && req.body._settings;
  const settings = hasWrapped ? req.body._settings : req.body;
  const opts = hasWrapped ? (req.body._opts || {}) : {};
  const updated = Object.assign({}, current, settings);
  writeJSON(SETTINGS_FILE, updated);
  emitter.emit('settings-saved', { themeOnly: !!opts.themeOnly, changedKeys: opts.changedKeys || [] });
  res.json({ ok: true });
});

app.get('/api/transcripts', (req, res) => {
  res.json(readJSON(TRANSCRIPTS_FILE, []));
});

app.post('/api/transcripts', (req, res) => {
  const transcripts = readJSON(TRANSCRIPTS_FILE, []);
  const transcript = { ...req.body, id: req.body.id || Date.now().toString() };
  const idx = transcripts.findIndex(t => t.id === transcript.id);
  if (idx >= 0) transcripts[idx] = transcript;
  else transcripts.unshift(transcript);
  writeJSON(TRANSCRIPTS_FILE, transcripts);
  res.json({ ok: true });
});

app.delete('/api/transcripts/:id', (req, res) => {
  let transcripts = readJSON(TRANSCRIPTS_FILE, []);
  transcripts = transcripts.filter(t => t.id !== req.params.id);
  writeJSON(TRANSCRIPTS_FILE, transcripts);
  res.json({ ok: true });
});

app.get('/api/themes', (req, res) => {
  res.json(readJSON(THEMES_FILE, []));
});

app.post('/api/themes', (req, res) => {
  writeJSON(THEMES_FILE, req.body);
  emitter.emit('themes-updated', req.body);
  res.json({ ok: true });
});

app.get('/api/songs', (req, res) => {
  res.json(readJSON(SONGS_FILE, []));
});

app.post('/api/songs', (req, res) => {
  writeJSON(SONGS_FILE, req.body);
  emitter.emit('songs-saved', req.body);
  res.json({ ok: true });
});

// ── Genius Lyrics Search API ──────────────────────────────────────────────────
function getGeniusApiKey() {
  if (process.env.GENIUS_API_KEY) return process.env.GENIUS_API_KEY;
  try {
    const s = readJSON(SETTINGS_FILE, {});
    if (s.geniusApiKey) return s.geniusApiKey;
  } catch (_) {}
  return '';
}

app.get('/api/genius/search', async (req, res) => {
  const q = (req.query.q || '').trim();
  if (!q) return res.json({ hits: [] });
  const GENIUS_API_KEY = getGeniusApiKey();
  if (!GENIUS_API_KEY) return res.status(500).json({ error: 'Genius API key not configured. Add it in Settings → AI & Transcription.' });
  try {
    const url = `https://api.genius.com/search?q=${encodeURIComponent(q)}&per_page=10`;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 10000);
    const resp = await fetch(url, { headers: { 'Authorization': `Bearer ${GENIUS_API_KEY}` }, signal: controller.signal });
    clearTimeout(timeout);
    if (!resp.ok) return res.status(resp.status).json({ error: resp.status === 429 ? 'Rate limited — try again shortly' : 'Search service unavailable' });
    const data = await resp.json();
    const hits = (data.response?.hits || []).map(h => ({
      id: h.result?.id,
      title: h.result?.title || '',
      artist: h.result?.primary_artist?.name || '',
      thumbnail: h.result?.song_art_image_thumbnail_url || '',
      url: h.result?.url || '',
      path: h.result?.path || '',
    }));
    res.json({ hits });
  } catch (e) {
    res.status(500).json({ error: 'Search request failed' });
  }
});

app.get('/api/genius/lyrics', async (req, res) => {
  const artist = (req.query.artist || '').trim();
  const title = (req.query.title || '').trim();
  if (!artist || !title) return res.status(400).json({ error: 'Missing artist or title parameter' });
  try {
    const url = `https://api.lyrics.ovh/v1/${encodeURIComponent(artist)}/${encodeURIComponent(title)}`;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 15000);
    const resp = await fetch(url, { signal: controller.signal });
    clearTimeout(timeout);
    if (resp.status === 404) return res.status(404).json({ error: 'Lyrics not found for this song' });
    if (resp.status === 429) return res.status(429).json({ error: 'Rate limited — try again shortly' });
    if (!resp.ok) return res.status(502).json({ error: 'Lyrics service unavailable' });
    const data = await resp.json();
    let lyrics = (data.lyrics || '').trim();
    if (!lyrics) return res.status(404).json({ error: 'No lyrics available' });
    const lines = lyrics.split('\n').map(l => l.trim());
    const cleaned = [];
    let prevBlank = false;
    for (const line of lines) {
      if (!line) {
        if (!prevBlank && cleaned.length) { cleaned.push(''); prevBlank = true; }
      } else {
        cleaned.push(line);
        prevBlank = false;
      }
    }
    while (cleaned.length && !cleaned[cleaned.length - 1]) cleaned.pop();
    while (cleaned.length && !cleaned[0]) cleaned.shift();
    res.json({ lyrics: cleaned.join('\n') });
  } catch (e) {
    res.status(500).json({ error: 'Lyrics request failed' });
  }
});

app.get('/api/media', (req, res) => {
  res.json(readJSON(MEDIA_FILE, []));
});

app.post('/api/media', (req, res) => {
  writeJSON(MEDIA_FILE, req.body);
  res.json({ ok: true });
});

app.get('/api/bible/installed', (req, res) => {
  const versions = {};
  try {
    const files = fs.readdirSync(BIBLE_VERSIONS_DIR);
    files.forEach(f => {
      if (f.endsWith('.json')) {
        versions[f.replace('.json', '')] = true;
      }
    });
  } catch (e) {}
  versions['KJV'] = true;
  res.json(versions);
});

app.get('/api/bible/load', (req, res) => {
  const kjvPath = path.join(__dirname, 'data', 'kjv.json');
  if (fs.existsSync(kjvPath)) {
    try {
      const data = JSON.parse(fs.readFileSync(kjvPath, 'utf8'));
      res.json({ KJV: data });
    } catch (e) {
      console.error('Failed to parse kjv.json:', e.message);
      res.status(500).json({ error: 'Bible data file is corrupted' });
    }
  } else {
    res.json({});
  }
});

function sanitizeTranslation(name) {
  return String(name || '').replace(/[^a-zA-Z0-9_\- ]/g, '').slice(0, 40);
}

app.post('/api/bible/save', (req, res) => {
  const { translation, data } = req.body;
  if (!translation || !data) return res.status(400).json({ error: 'Missing translation or data' });
  const safe = sanitizeTranslation(translation);
  if (!safe) return res.status(400).json({ error: 'Invalid translation name' });
  writeJSON(path.join(BIBLE_VERSIONS_DIR, `${safe}.json`), data);
  emitter.emit('bible-versions-updated');
  res.json({ ok: true });
});

app.delete('/api/bible/:translation', (req, res) => {
  const safe = sanitizeTranslation(req.params.translation);
  if (!safe) return res.status(400).json({ error: 'Invalid translation name' });
  const file = path.join(BIBLE_VERSIONS_DIR, `${safe}.json`);
  if (fs.existsSync(file)) fs.unlinkSync(file);
  emitter.emit('bible-versions-updated');
  res.json({ ok: true });
});

app.get('/api/render-state', (req, res) => {
  res.json(currentRenderState);
});

app.post('/api/render-state', (req, res) => {
  if (req.body.module === 'alert' || req.body.module === 'caption') {
    emitter.emit('render-state', req.body);
  } else if (req.body.module === 'timer-scale') {
    if (currentRenderState.module === 'timer' && currentRenderState.payload) {
      currentRenderState.payload.scale = req.body.payload?.scale ?? 1;
    }
    emitter.emit('render-state', req.body);
  } else if (req.body.module === 'logo-overlay') {
    emitter.emit('render-state', req.body);
  } else {
    currentRenderState = req.body;
    emitter.emit('render-state', currentRenderState);
  }
  res.json({ ok: true });
});

app.post('/api/logo-overlay-drag-update', (req, res) => {
  emitter.emit('logo-overlay-drag-update', req.body);
  res.json({ ok: true });
});

app.get('/api/app-info', (req, res) => {
  res.json({
    version: '1.1.0',
    platform: process.platform,
    isElectron: false,
    isWeb: true,
  });
});

const PRESETS_FILE = path.join(DATA_DIR, 'presets.json');

app.get('/api/presets', (req, res) => {
  res.json(readJSON(PRESETS_FILE, []));
});

app.post('/api/presets', (req, res) => {
  try {
    let presets = readJSON(PRESETS_FILE, []);
    if (!Array.isArray(presets)) presets = [];
    const preset = req.body;
    const safeId = String(preset.id || '').replace(/[^a-zA-Z0-9_-]/g, '');
    if (!safeId) return res.json({ success: false, error: 'Invalid preset id' });
    preset.id = safeId;
    if (preset.logoOverlay?.fileData) {
      const rawExt = path.extname(preset.logoOverlay.fileName || '.png') || '.png';
      const ext = rawExt.replace(/[^a-zA-Z0-9.]/g, '');
      const safeName = `preset_logo_${safeId}${ext}`;
      const dest = path.join(PRESETS_ASSETS, safeName);
      if (!dest.startsWith(PRESETS_ASSETS)) return res.json({ success: false, error: 'Invalid path' });
      fs.writeFileSync(dest, Buffer.from(preset.logoOverlay.fileData, 'base64'));
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
    writeJSON(PRESETS_FILE, presets);
    res.json({ success: true });
  } catch(e) { res.json({ success: false, error: e.message }); }
});

app.post('/api/presets/delete', (req, res) => {
  try {
    let presets = readJSON(PRESETS_FILE, []);
    const preset = presets.find(p => p.id === req.body.id);
    if (preset?.logoOverlay?.savedPath) {
      const resolved = path.resolve(preset.logoOverlay.savedPath);
      if (resolved.startsWith(PRESETS_ASSETS)) {
        try { fs.unlinkSync(resolved); } catch(_) {}
      }
    }
    presets = presets.filter(p => p.id !== req.body.id);
    writeJSON(PRESETS_FILE, presets);
    res.json({ success: true });
  } catch(e) { res.json({ success: false, error: e.message }); }
});

app.get('/api/remote-info', (req, res) => {
  res.json({ port: PORT, enabled: false, url: null });
});

app.get('/api/network-adapters', (req, res) => {
  const ifaces = os.networkInterfaces();
  const adapters = Object.keys(ifaces);
  res.json(adapters);
});

app.get('/api/whisper/status', (req, res) => {
  res.json({ ready: false, status: 'Web mode — local Whisper not available' });
});

app.get('/api/ndi/status', (req, res) => {
  res.json({ active: false, status: 'NDI not available in web mode' });
});

app.get('/api/schedules', (req, res) => {
  res.json([]);
});

app.get('/api/adaptive/dashboard', (req, res) => {
  res.json({ speakerProfiles: [], rules: [], vocab: [], suggestions: [] });
});

const PRESENTATIONS_FILE = path.join(DATA_DIR, 'presentations.json');
const CREATED_PRES_FILE = path.join(DATA_DIR, 'created_presentations.json');

app.get('/api/presentations', (req, res) => {
  res.json(readJSON(PRESENTATIONS_FILE, []));
});

app.get('/api/created-presentations', (req, res) => {
  res.json(readJSON(CREATED_PRES_FILE, []));
});

app.post('/api/created-presentations', (req, res) => {
  writeJSON(CREATED_PRES_FILE, req.body);
  res.json({ ok: true });
});

app.get('/api/displays', (req, res) => {
  res.json([{ id: 1, label: 'Display 1', bounds: { x: 0, y: 0, width: 1920, height: 1080 } }]);
});

const clients = new Set();

app.get('/api/events', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.flushHeaders();

  const send = (event, data) => {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };

  clients.add(send);

  const handlers = {
    'settings-saved': (d) => send('settings-saved', d),
    'themes-updated': (d) => send('themes-updated', d),
    'songs-saved': (d) => send('songs-saved', d),
    'bible-versions-updated': () => send('bible-versions-updated', {}),
    'render-state': (d) => send('render-state', d),
    'logo-overlay-drag-update': (d) => send('logo-overlay-drag-update', d),
  };

  Object.entries(handlers).forEach(([ev, fn]) => emitter.on(ev, fn));

  req.on('close', () => {
    clients.delete(send);
    Object.entries(handlers).forEach(([ev, fn]) => emitter.off(ev, fn));
  });
});

let liveState = { live: false, liveType: '', songTitle: '', songSlide: null, verse: null, mediaId: null };

function buildPreviewPayload() {
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

app.get('/api/status', (req, res) => {
  const preview = buildPreviewPayload();
  res.json({ ...liveState, preview });
});

app.post('/api/status', (req, res) => {
  Object.assign(liveState, req.body);
  res.json({ ok: true });
});

app.get('/api/bootstrap', (req, res) => {
  res.json({
    ok: true,
    role: 'admin',
    capabilities: ['status','go-live','clear','next','prev','scripture','songs','media','queue-add','set-translation','library'],
    authRequired: false,
    version: 'web'
  });
});

app.get('/api/library', (req, res) => {
  const type = req.query.type;
  if (type === 'songs') {
    const songs = readJSON(SONGS_FILE, []);
    res.json({ items: songs });
  } else if (type === 'media') {
    const media = readJSON(MEDIA_FILE, []);
    res.json({ items: media });
  } else {
    res.json({ items: [] });
  }
});

const VALID_CONTROL_ACTIONS = ['present','present-song','present-media','present-current','clear','next','prev','presentation-next','presentation-prev','present-verse'];
let pendingControlCommands = [];
let _controlCmdSeq = 0;

function _serverProjection(module, payload) {
  currentRenderState = { module, payload, updatedAt: Date.now() };
  emitter.emit('render-state', currentRenderState);
  for (const send of clients) {
    send('render-state', currentRenderState);
  }
}

let _serverLiveState = { songId: null, slideIdx: 0, mediaId: null, mode: null };

function _handleServerSideProjection(action, data) {
  if (action === 'clear') {
    _serverLiveState = { songId: null, slideIdx: 0, mediaId: null, mode: null };
    liveState = { live: false, liveType: '', songTitle: '', songSlide: null, verse: null, mediaId: null };
    _serverProjection('clear', null);
    return;
  }
  if (action === 'present-song') {
    const songs = readJSON(SONGS_FILE, []);
    const song = songs.find(s => String(s.id) === String(data?.songId));
    if (song && song.sections) {
      const idx = Math.max(0, Math.min(parseInt(data?.slideIdx ?? 0, 10) || 0, (song.sections.length || 1) - 1));
      _serverLiveState = { songId: String(song.id), slideIdx: idx, mediaId: null, mode: 'song' };
      liveState.live = true;
      liveState.liveType = 'song';
      liveState.songTitle = song.title || '';
      liveState.songSlide = idx;
      liveState.verse = null;
      liveState.mediaId = null;
      const sec = song.sections[idx];
      const lines = (sec?.lines || []).filter(l => l && l.trim());
      _serverProjection('song', {
        title: song.title || '',
        author: song.author || '',
        sectionLabel: sec?.label || '',
        lines,
        theme: 'song_sanctuary',
        songVerticalAlign: 'center',
      });
    }
    return;
  }
  if (action === 'present-media') {
    const media = readJSON(MEDIA_FILE, []);
    const item = media.find(m => String(m.id) === String(data?.mediaId));
    if (item) {
      _serverLiveState = { songId: null, slideIdx: 0, mediaId: String(item.id), mode: 'media' };
      liveState.live = true;
      liveState.liveType = 'media';
      liveState.songTitle = '';
      liveState.songSlide = null;
      liveState.verse = null;
      liveState.mediaId = String(item.id);
      _serverProjection('media', {
        id: item.id, name: item.title || item.name, path: item.path,
        type: item.type, loop: item.loop !== false, mute: item.mute === true,
      });
    }
    return;
  }
  if (action === 'present-current') {
    if (_serverLiveState.mode === 'song' && _serverLiveState.songId) {
      const songs = readJSON(SONGS_FILE, []);
      const song = songs.find(s => String(s.id) === String(_serverLiveState.songId));
      if (song?.sections) {
        const idx = _serverLiveState.slideIdx;
        const sec = song.sections[idx];
        if (sec) {
          const lines = (sec.lines || []).filter(l => l && l.trim());
          _serverProjection('song', {
            title: song.title || '', author: song.author || '',
            sectionLabel: sec.label || '', lines,
            theme: 'song_sanctuary', songVerticalAlign: 'center',
          });
        }
      }
    }
    return;
  }
  if (action === 'next' || action === 'prev') {
    const dir = action === 'next' ? 1 : -1;
    if (_serverLiveState.mode === 'song' && _serverLiveState.songId) {
      const songs = readJSON(SONGS_FILE, []);
      const song = songs.find(s => String(s.id) === String(_serverLiveState.songId));
      if (song?.sections?.length) {
        const newIdx = Math.max(0, Math.min(_serverLiveState.slideIdx + dir, song.sections.length - 1));
        _serverLiveState.slideIdx = newIdx;
        const sec = song.sections[newIdx];
        const lines = (sec?.lines || []).filter(l => l && l.trim());
        liveState.songSlide = newIdx;
        liveState.songTitle = song.title || '';
        _serverProjection('song', {
          title: song.title || '', author: song.author || '',
          sectionLabel: sec?.label || '', lines,
          theme: 'song_sanctuary', songVerticalAlign: 'center',
        });
      }
    }
    return;
  }
}

app.post('/api/control', (req, res) => {
  const { action, data } = req.body || {};
  if (!action || typeof action !== 'string' || !VALID_CONTROL_ACTIONS.includes(action)) {
    return res.status(400).json({ error: 'Invalid action' });
  }
  const cmd = { id: ++_controlCmdSeq, action, data: data || null };
  console.log(`[Control] cmd#${cmd.id} ${action} → ${clients.size} SSE client(s)`);
  const SERVER_HANDLED = ['present-song','present-media','present-current','clear','next','prev'];
  if (SERVER_HANDLED.includes(action)) {
    _handleServerSideProjection(action, data);
  } else {
    pendingControlCommands.push(cmd);
    if (pendingControlCommands.length > 50) pendingControlCommands = pendingControlCommands.slice(-50);
    if (action === 'present' || action === 'present-verse') {
      const ref = data?.rawRef || data?.ref || '';
      if (ref) {
        liveState.live = true;
        liveState.liveType = 'scripture';
        liveState.verse = { ref };
        liveState.songTitle = '';
        liveState.songSlide = null;
        liveState.mediaId = null;
        _serverLiveState = { songId: null, slideIdx: 0, mediaId: null, mode: 'scripture' };
        _serverProjection('scripture', { ref, text: data?.text || '' });
      }
    }
    emitter.emit('remote-control', cmd);
    for (const send of clients) {
      send('remote-control', cmd);
    }
  }
  res.json({ ok: true });
});

app.get('/api/control/pending', (req, res) => {
  const afterId = parseInt(req.query.after) || 0;
  const cmds = pendingControlCommands.filter(c => c.id > afterId);
  res.json({ commands: cmds });
});

app.get('/remote', (req, res) => {
  res.send(buildRemoteHTML());
});

function buildRemoteHTML() {
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
.hdr{padding:14px 16px 12px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;background:rgba(7,7,14,.95);backdrop-filter:blur(20px)}
.hdr-brand{display:flex;align-items:center;gap:8px}
.hdr-cross{font-size:18px;color:var(--gold)}
.hdr-title{font-size:13px;font-weight:700;color:var(--gold);letter-spacing:.1em;text-transform:uppercase}
.hdr-sub{font-size:8px;color:var(--text-muted);letter-spacing:.15em;text-transform:uppercase}
.onair{display:flex;align-items:center;gap:5px;padding:4px 9px;border-radius:20px;background:var(--live-dim);border:1px solid rgba(231,76,60,.3);font-size:9px;font-weight:700;color:var(--live);letter-spacing:.08em;display:none}
.onair.show{display:flex}
.onair-dot{width:6px;height:6px;border-radius:50%;background:var(--live);animation:pulse 1.4s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.status-bar{margin:12px 14px 0;background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:12px 14px;display:flex;align-items:center;gap:10px;min-height:56px}
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
.status-icon{width:34px;height:34px;border-radius:50%;background:var(--surface);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0}
.status-what{font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.status-type{font-size:9px;color:var(--text-dim);letter-spacing:.07em;text-transform:uppercase;margin-top:2px}
.body{padding:0 14px}
.sec-lbl{font-size:9px;font-weight:600;color:var(--text-muted);letter-spacing:.16em;text-transform:uppercase;padding-top:14px;margin-bottom:6px}
.tabs{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-bottom:12px}
.tab{padding:9px 3px 7px;border-radius:var(--r-sm);border:1px solid var(--border);background:var(--card);color:var(--text-muted);font-size:8px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;cursor:pointer;display:flex;flex-direction:column;align-items:center;gap:4px;transition:all .12s;font-family:inherit}
.tab .ti{font-size:16px;line-height:1}
.tab.active{background:var(--gold-dim);border-color:var(--border-gold);color:var(--gold)}
.tab:not(:disabled):active{opacity:.7;transform:scale(.97)}
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
.media-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r-sm);padding:11px 13px;display:flex;align-items:center;gap:11px;cursor:pointer;transition:all .12s}
.media-card:active{background:var(--card2)}
.media-thumb{width:38px;height:38px;border-radius:7px;background:var(--card2);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}
.media-info{flex:1;min-width:0}
.media-name{font-size:12px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.media-type{font-size:9px;color:var(--text-dim);letter-spacing:.07em;text-transform:uppercase;margin-top:2px}
.media-play{width:30px;height:30px;border-radius:50%;background:var(--gold-dim);border:1px solid var(--border-gold);color:var(--gold);font-size:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.pres-note{font-size:11px;color:var(--text-dim);text-align:center;margin-bottom:12px;line-height:1.5}
.pres-nav{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-bottom:9px}
.pres-btn{padding:15px 8px;border-radius:var(--r);border:1px solid var(--border);background:var(--card2);color:var(--text-dim);font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:7px;transition:all .12s;font-family:inherit}
.pres-btn:active{background:var(--surface);transform:scale(.97)}
.pres-present{width:100%;padding:14px;border-radius:var(--r);border:1px solid rgba(74,158,232,.3);background:rgba(74,158,232,.1);color:var(--blue);font-size:12px;font-weight:600;cursor:pointer;font-family:inherit}
.pres-present:active{opacity:.7}
.broadcast{position:fixed;bottom:0;left:0;right:0;max-width:430px;margin:0 auto;background:rgba(7,7,14,.97);border-top:1px solid var(--border);padding:10px 14px 12px;z-index:50;display:flex;flex-direction:column;gap:7px}
.go-btn{width:100%;padding:15px;border-radius:var(--r);border:1px solid rgba(231,76,60,.35);background:var(--live-dim);color:var(--live);font-size:14px;font-weight:700;letter-spacing:.12em;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:9px;transition:all .12s;font-family:inherit}
.go-btn.on{background:var(--live);color:#fff;border-color:var(--live)}
.go-btn:active{transform:scale(.98)}
.ctrl-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px}
.ctrl-btn{padding:11px 6px;border-radius:var(--r-sm);border:1px solid var(--border);background:var(--card);color:var(--text-dim);font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:5px;transition:all .12s;font-family:inherit}
.ctrl-btn .arr{font-size:16px;color:var(--text-muted)}
.ctrl-btn:active{background:var(--card2);transform:scale(.97)}
.role-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 9px;border-radius:20px;background:var(--gold-glow);border:1px solid var(--border-gold);font-size:9px;font-weight:600;color:var(--gold);letter-spacing:.07em;text-transform:uppercase;margin:10px 14px 0}
.toast-wrap{position:fixed;bottom:118px;left:50%;transform:translateX(-50%);z-index:999;width:calc(100% - 28px);max-width:400px;pointer-events:none}
#toast{font-size:12px;font-weight:500;padding:10px 14px;border-radius:var(--r-sm);background:rgba(39,174,96,.14);border:1px solid rgba(39,174,96,.32);color:#5de88a;transition:opacity .22s,transform .22s;opacity:0;transform:translateY(6px);text-align:center}
#toast.err{background:rgba(231,76,60,.12);border-color:rgba(231,76,60,.28);color:#ff7b6b}
#toast.show{opacity:1;transform:translateY(0)}
.hidden{display:none}
.home-btn{background:var(--card);border:1px solid var(--border);border-radius:var(--r-xs);color:var(--text-dim);font-size:16px;padding:6px 10px;cursor:pointer;display:flex;align-items:center;justify-content:center}
.home-btn:active{background:var(--card2)}
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
}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-brand">
    <div class="hdr-cross">\u271D</div>
    <div><div class="hdr-title">AnchorCast</div><div class="hdr-sub">Remote</div></div>
  </div>
  <div style="display:flex;align-items:center;gap:8px">
    <div class="onair" id="onairBadge"><div class="onair-dot"></div>ON AIR</div>
    <button class="home-btn" onclick="window.location.href='/'" title="Back to main app">\uD83C\uDFE0</button>
  </div>
</div>

<div class="ls-wrap">
<div class="ls-left">

<div class="status-bar">
  <div class="status-icon" id="statusIcon">\uD83D\uDCE1</div>
  <div><div class="status-what" id="statusWhat">Connecting\u2026</div><div class="status-type" id="statusType">Please wait</div></div>
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

<div class="role-badge" id="roleBadge">\uD83D\uDC64 Admin</div>

<div class="body">
  <div class="sec-lbl">Mode</div>
  <div class="tabs">
    <button class="tab active" id="modeScripture"><span class="ti">\uD83D\uDCD6</span>Scripture</button>
    <button class="tab" id="modeSongs"><span class="ti">\uD83C\uDFB5</span>Songs</button>
    <button class="tab" id="modeMedia"><span class="ti">\uD83C\uDFAC</span>Media</button>
    <button class="tab" id="modePresentation"><span class="ti">\uD83D\uDCD1</span>Slides</button>
  </div>

  <div class="hidden" id="scriptureSec">
    <div class="scripture-live" id="scriptureNowLive">
      <div class="scripture-live-dot"></div>
      <div style="font-size:12px;font-weight:600;color:var(--text);flex:1" id="scriptureNowLiveRef">\u2014</div>
      <div style="font-size:9px;color:var(--text-dim)" id="scriptureNowLiveType">LIVE</div>
    </div>
    <div class="sec-lbl" style="padding-top:0">Search &amp; Send</div>
    <div class="ref-row">
      <input class="ref-input" id="refInput" type="text" placeholder="John 3:16 \u00B7 Ps 23 \u00B7 Rom 8:28" autocomplete="off" autocorrect="off" autocapitalize="words" spellcheck="false">
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

  <div class="hidden" id="songsSec">
    <div class="active-bar" id="activeSongBar">
      <div class="active-dot"></div>
      <div class="active-name" id="activeSongLabel">No song live</div>
      <div class="active-slide" id="activeSongSlide"></div>
    </div>
    <div style="display:flex;gap:7px;margin-bottom:10px">
      <input id="songSearch" type="text" placeholder="Search songs\u2026" autocomplete="off" spellcheck="false"
        style="flex:1;background:var(--surface);border:1px solid var(--border);border-radius:var(--r-sm);
        color:var(--text);font-size:14px;padding:10px 12px;outline:none;font-family:inherit;-webkit-appearance:none">
      <button id="songSearchClear"
        style="padding:10px 13px;border-radius:var(--r-sm);border:1px solid var(--border);
        background:var(--card);color:var(--text-muted);font-size:15px;cursor:pointer;font-family:inherit;min-width:40px">\u2715</button>
    </div>
    <div class="songs-list" id="songsList">
      <div class="song-card"><div class="song-title" style="color:var(--text-dim)">Loading songs\u2026</div></div>
    </div>
  </div>

  <div class="hidden" id="mediaSec">
    <div class="media-list" id="mediaList">
      <div class="song-card"><div class="song-title" style="color:var(--text-dim)">Loading media\u2026</div></div>
    </div>
  </div>

  <div class="hidden" id="presentationSec">
    <div class="pres-note">Control the presentation from here.<br>Use the main app to load slides first.</div>
    <div class="pres-nav">
      <button class="pres-btn" id="presPrevBtn"><span style="font-size:18px">\u2039</span> Prev Slide</button>
      <button class="pres-btn" id="presNextBtn">Next Slide <span style="font-size:18px">\u203A</span></button>
    </div>
    <button class="pres-present" id="presCurrentBtn">\u25B6 \u00A0Present Current Slide</button>
  </div>
</div>

</div>
</div>

<div class="broadcast">
  <button class="go-btn" id="liveBtn"><span id="liveBtnDot">\u25CF</span><span id="liveBtnText">GO LIVE</span></button>
  <div class="ctrl-row">
    <button class="ctrl-btn" id="prevBtn"><span class="arr">\u2039</span> Prev</button>
    <button class="ctrl-btn" id="clearBtn" style="border-color:rgba(255,255,255,.05);color:var(--text-muted);font-size:11px">\u2715 Clear</button>
    <button class="ctrl-btn" id="nextBtn">Next <span class="arr">\u203A</span></button>
  </div>
</div>

<div class="toast-wrap"><div id="toast"></div></div>

<script>
(function(){
'use strict';
var activeMode='scripture';
var liveState={live:false,liveType:'',songTitle:'',songSlide:null};
var songSel=null,mediaSel=null,allSongs=[],toastTimer=null;

function $(id){return document.getElementById(id);}
function toast(msg,isErr){
  var el=$('toast');if(!el)return;
  el.textContent=msg;el.className='show'+(isErr?' err':'');
  clearTimeout(toastTimer);toastTimer=setTimeout(function(){el.className=el.className.replace('show','').trim();},2500);
}
function setStatus(what,type,icon){
  var wi=$('statusWhat'),ti=$('statusType'),si=$('statusIcon');
  if(wi)wi.textContent=what||'\\u2014';if(ti)ti.textContent=type||'';if(si)si.textContent=icon||'\\uD83D\\uDCE1';
  var b=$('onairBadge');if(b)b.className='onair'+(liveState.live?' show':'');
}

function xhr(method,url,body,cb){
  var x=new XMLHttpRequest();x.open(method,url,true);
  x.setRequestHeader('Content-Type','application/json');
  x.onreadystatechange=function(){if(x.readyState!==4)return;var d={};try{d=JSON.parse(x.responseText||'{}');}catch(e){}cb(x.status,d);};
  x.onerror=function(){cb(0,{});};
  x.send(body?JSON.stringify(body):null);
}
function cmd(action,data,okMsg){
  xhr('POST','/api/control',{action:action,data:data||null},function(status,resp){
    if(status>=200&&status<300){if(okMsg)toast(okMsg);setTimeout(poll,300);}
    else toast('Error '+status,true);
  });
}

function goLive(){
  cmd('present-current',{mode:activeMode},'\\u2713 Sent to projection');
}

function sendRef(){
  var v=($('refInput')||{}).value||'';v=v.trim();
  if(!v){toast('Enter a reference',true);return;}
  cmd('present',{ref:v,rawRef:v},'\\u2713 '+v);
}

function setMode(m){
  activeMode=m;
  if(m!=='songs'){var ss=$('songSearch');if(ss)ss.value='';if(allSongs.length)renderSongsFiltered(allSongs);}
  refreshUI();
  if(m==='songs')loadLib('songs');
  if(m==='media')loadLib('media');
}

function refreshUI(){
  ['scripture','songs','media','presentation'].forEach(function(m){
    var t=$('mode'+m.charAt(0).toUpperCase()+m.slice(1));
    if(t)t.className='tab'+(activeMode===m?' active':'');
  });
  var secMap={scripture:'scriptureSec',songs:'songsSec',media:'mediaSec',presentation:'presentationSec'};
  ['scripture','songs','media','presentation'].forEach(function(m){
    var el=$(secMap[m]);
    if(!el)return;
    el.className=el.className.replace(/\\bhidden\\b/g,'').trim();
    el.style.display=(activeMode===m)?'block':'none';
  });
  var lb=$('liveBtn'),lt=$('liveBtnText');
  if(!lb)return;
  lb.className='go-btn'+(liveState.live?' on':'');
  var lbl=activeMode==='songs'?'PRESENT SONG':activeMode==='media'?'PRESENT MEDIA':activeMode==='presentation'?'PRESENT SLIDE':'GO LIVE';
  if(liveState.live&&activeMode==='scripture')lbl='RE-SEND LIVE';
  if(lt)lt.textContent=lbl;
}

function poll(){
  xhr('GET','/api/status',null,function(status,data){
    if(status!==200){setStatus('No connection','Check network','\\u26A0');return;}
    liveState.live=!!data.live;
    liveState.liveType=data.liveType||'';
    liveState.songTitle=data.songTitle||'';
    liveState.songSlide=data.songSlide!=null?data.songSlide:null;
    var what='',type='',icon='\\uD83D\\uDCE1';
    if(data.verse&&data.verse.ref){what=data.verse.ref;type='Scripture';icon='\\u271D';}
    else if(data.songTitle){what=data.songTitle;type='Song'+(data.songSlide!=null?' \\u00B7 Slide '+(data.songSlide+1):'');icon='\\uD83C\\uDFB5';}
    else if(data.liveType==='media'){what='Media playing';type='Media';icon='\\uD83C\\uDFAC';}
    else if(data.liveType==='presentation'){what='Presentation';type='Slides';icon='\\uD83D\\uDCD1';}
    else{what='Nothing on screen';type='Clear';icon='\\u25A1';}
    setStatus(what,type+(data.live?' \\u2022 ON AIR':''),icon);
    refreshUI();
    var snl=$('scriptureNowLive'),snlRef=$('scriptureNowLiveRef');
    if(snl&&snlRef){
      if(data.verse&&data.verse.ref){snl.className='scripture-live show';snlRef.textContent=data.verse.ref;}
      else{snl.className='scripture-live';}
    }
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
    var ic=pv.mediaType==='video'?'\\uD83C\\uDFAC':pv.mediaType==='audio'?'\\uD83C\\uDFA7':'\\uD83D\\uDDBC';
    html='<div class="proj-media-icon">'+ic+'</div><div class="proj-media-name">'+esc(pv.name||'Media')+'</div>';
  }else if(pv.type==='presentation'){
    var sn=pv.slide!=null?String(Math.trunc(+pv.slide)||''):'';
    html='<div class="proj-media-icon">\\uD83D\\uDCD1</div><div class="proj-media-name">Slide '+esc(sn)+'</div>';
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

function renderSongs(items){allSongs=items||[];renderSongsFiltered(allSongs);}
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
          document.querySelectorAll('#songsList .chip').forEach(function(x){x.classList.remove('active');});
          c.classList.add('active');
          cmd('present-song',{songId:sid,slideIdx:si},'\\u2713 '+label+' \\u2014 Sent to projection');
        };
        chips.appendChild(c);
      })(s.id,idx,sec);
    });
    card.appendChild(chips);el.appendChild(card);
  });
}

function renderMedia(items){
  var el=$('mediaList');if(!el)return;
  if(!items.length){el.innerHTML='<div class="song-card"><div class="song-title" style="color:var(--text-dim)">No media files</div></div>';return;}
  el.innerHTML='';
  var icons={video:'\\uD83C\\uDFAC',image:'\\uD83D\\uDDBC',audio:'\\uD83C\\uDFA7'};
  items.forEach(function(item){
    var card=document.createElement('div');card.className='media-card';
    var thumb=document.createElement('div');thumb.className='media-thumb';thumb.textContent=icons[item.type]||'\\uD83C\\uDFAC';
    var info=document.createElement('div');info.className='media-info';
    var name=document.createElement('div');name.className='media-name';name.textContent=item.title||'Untitled';
    var mtype=document.createElement('div');mtype.className='media-type';mtype.textContent=(item.type||'').toUpperCase();
    info.appendChild(name);info.appendChild(mtype);
    var play=document.createElement('div');play.className='media-play';play.textContent='\\u25BA';
    card.appendChild(thumb);card.appendChild(info);card.appendChild(play);
    card.onclick=function(){
      mediaSel={mediaId:item.id};
      cmd('present-media',{mediaId:item.id},'\\u2713 '+item.title+' \\u2014 Sent to projection');
    };
    el.appendChild(card);
  });
}

window.addEventListener('load',function(){
  $('modeScripture').onclick=function(){setMode('scripture');};
  $('modeSongs').onclick=function(){setMode('songs');};
  $('modeMedia').onclick=function(){setMode('media');};
  $('modePresentation').onclick=function(){setMode('presentation');};
  $('liveBtn').onclick=goLive;
  $('prevBtn').onclick=function(){cmd('prev',{mode:activeMode},'\\u2713 Prev');};
  $('nextBtn').onclick=function(){cmd('next',{mode:activeMode},'\\u2713 Next');};
  $('clearBtn').onclick=function(){cmd('clear',null,'\\u2713 Display cleared');};
  $('projClearBtn').onclick=function(){cmd('clear',null,'\\u2713 Display cleared');};
  $('sendBtn').onclick=sendRef;
  $('refInput').addEventListener('keydown',function(e){if(e.key==='Enter'){e.preventDefault();sendRef();}});
  $('smpJohn316').onclick=function(){cmd('present',{ref:'John 3:16',rawRef:'John 3:16'},'\\u2713 John 3:16');};
  $('smpPsalm23').onclick=function(){cmd('present',{ref:'Psalms 23:1',rawRef:'Psalms 23:1'},'\\u2713 Psalm 23:1');};
  $('smpPhil413').onclick=function(){cmd('present',{ref:'Philippians 4:13',rawRef:'Philippians 4:13'},'\\u2713 Phil 4:13');};
  $('smpRom828').onclick=function(){cmd('present',{ref:'Romans 8:28',rawRef:'Romans 8:28'},'\\u2713 Romans 8:28');};
  $('smpIsa4031').onclick=function(){cmd('present',{ref:'Isaiah 40:31',rawRef:'Isaiah 40:31'},'\\u2713 Isaiah 40:31');};
  $('smpJer2911').onclick=function(){cmd('present',{ref:'Jeremiah 29:11',rawRef:'Jeremiah 29:11'},'\\u2713 Jeremiah 29:11');};
  var songSearchEl=$('songSearch'),songSearchClearEl=$('songSearchClear');
  if(songSearchEl){
    songSearchEl.addEventListener('input',function(){
      var q=(this.value||'').toLowerCase().trim();
      if(!q){renderSongsFiltered(allSongs);return;}
      renderSongsFiltered(allSongs.filter(function(s){return(s.title||'').toLowerCase().indexOf(q)>=0||(s.author||'').toLowerCase().indexOf(q)>=0;}));
    });
  }
  if(songSearchClearEl)songSearchClearEl.onclick=function(){if(songSearchEl)songSearchEl.value='';renderSongsFiltered(allSongs);};
  $('presPrevBtn').onclick=function(){cmd('presentation-prev',null,'\\u2713 Prev slide');};
  $('presNextBtn').onclick=function(){cmd('presentation-next',null,'\\u2713 Next slide');};
  $('presCurrentBtn').onclick=function(){cmd('present-current',{mode:'presentation'},'\\u2713 Slide presented');};
  refreshUI();loadLib('songs');loadLib('media');poll();
  setInterval(poll,2500);
});
})();
</script>
</body></html>`;
}

app.get('/{*path}', (req, res) => {
  res.sendFile(path.join(__dirname, 'src', 'renderer', 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`AnchorCast web server running on http://0.0.0.0:${PORT}`);
});

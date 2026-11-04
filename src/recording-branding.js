'use strict';
// Reusable intro/outro assets and non-destructive MP3 assembly. No Electron dependency
// in the core service; all paths supplied by the native bridge are constrained.
const fs = require('node:fs'), fsp = fs.promises, path = require('node:path'), crypto = require('node:crypto');
const https = require('node:https');
const { spawn } = require('node:child_process');
const { EventEmitter } = require('node:events');
const C = require('./recording-branding-config');
const { bundledEncoder } = require('./sermon-audio-recorder');
const { fileIdentity, validateTrim, validateRequestId } = require('./recording-trim');
const { RecordingPreview } = require('./recording-preview');
const ASSET = /^(?:voice|mix|music)-[a-f0-9]{32}\.(?:wav|mp3)$/;
const EXPORT = /^Message-\d{8}-\d{6}-[a-f0-9]{32}\.mp3$/;
const MAX_FILE = 30 * 1024 * 1024;
const clone = x => JSON.parse(JSON.stringify(x));
function aborted(signal) { if (signal?.aborted) throw Error('Operation cancelled. The original recording is unchanged.'); }
async function regular(file, maxSize = Infinity) {
  const s = await fsp.lstat(file);
  if (!s.isFile() || s.isSymbolicLink() || s.size < 1 || s.size > maxSize) throw Error('Audio file is empty, too large, or not a regular local file.');
  return s;
}
async function atomic(file, bytes) {
  const tmp = file + '.' + crypto.randomBytes(8).toString('hex') + '.tmp';
  try { await fsp.writeFile(tmp, bytes, { flag: 'wx', mode: 0o600 }); await fsp.rename(tmp, file); }
  finally { await fsp.unlink(tmp).catch(() => {}); }
}
function wavDuration(buffer) {
  if (buffer.toString('ascii', 0, 4) !== 'RIFF' || buffer.toString('ascii', 8, 12) !== 'WAVE') throw Error('Decoded audio is not a WAV file.');
  let rate = 0, size = 0;
  for (let p = 12; p + 8 <= buffer.length;) {
    const kind = buffer.toString('ascii', p, p + 4), n = buffer.readUInt32LE(p + 4);
    if (p + 8 + n > buffer.length) throw Error('Decoded audio is incomplete.');
    if (kind === 'fmt ' && n >= 16) rate = buffer.readUInt32LE(p + 16);
    if (kind === 'data') size += n;
    p += 8 + n + (n % 2);
  }
  if (!rate || !size || size / rate < .15) throw Error('No usable speech/audio was found.');
  return size / rate;
}
function requestSpeech({ key, payload, signal, request = https.request, timeoutMs = 120000 }) {
  if (!key || typeof key !== 'string' || key.length > 1000 || /[\r\n]/.test(key)) return Promise.reject(Error('Add and save your OpenAI API key in Audio & AI → General before generating a voiceover.'));
  // Do not send a known Anthropic credential to the OpenAI speech endpoint.
  // Claude's current public model API returns text, not speech audio.
  if (/^sk-ant-/i.test(key.trim())) return Promise.reject(Error('A Claude (Anthropic) API key cannot generate speech audio here. Use an OpenAI API key for Generate voice, or import your prepared narration.'));
  const body = Buffer.from(JSON.stringify(payload));
  return new Promise((resolve, reject) => {
    let settled = false, response, req, timer;
    const done = (error, data) => { if (settled) return; settled = true; clearTimeout(timer); signal?.removeEventListener('abort', cancel); error ? reject(error) : resolve(data); };
    const cancel = () => { const error = Error('Voice generation cancelled.'); done(error); req?.destroy(); response?.destroy(); };
    try {
      aborted(signal);
      req = request({ hostname: 'api.openai.com', port: 443, path: '/v1/audio/speech', method: 'POST',
        headers: { 'Authorization': 'Bearer ' + key, 'Content-Type': 'application/json', 'Content-Length': body.length } }, res => {
        response = res;
        if (res.statusCode !== 200) {
          const code = Number(res.statusCode), reasons = { 401: 'The OpenAI API key was rejected. Check the saved key.', 403: 'This API key is not allowed to generate speech.', 429: 'OpenAI rate or account usage limit reached. Check your API billing/limits before retrying.' };
          done(Error(reasons[code] || `Speech service returned HTTP ${code}. Please try again later.`));
          res.destroy(); return;
        }
        const type = String(res.headers['content-type'] || '').toLowerCase();
        if (!type.startsWith('audio/') && !type.includes('octet-stream')) { done(Error('Speech service returned an unexpected response instead of audio.')); res.destroy(); return; }
        let length = 0; const parts = [];
        res.on('data', data => { length += data.length; if (length > MAX_FILE) { done(Error('Generated voiceover is too large. Shorten the text.')); res.destroy(); return; } parts.push(data); });
        res.on('aborted', () => done(Error('Speech download was interrupted. The previous voiceover is unchanged.')));
        res.on('error', () => done(Error('Could not download the generated speech.')));
        res.on('end', () => length > 100 ? done(null, Buffer.concat(parts)) : done(Error('Speech service returned no usable audio.')));
      });
      req.on('error', () => done(Error('Could not connect to the speech service. Check your internet connection.')));
      req.setTimeout(30000, () => { done(Error('Speech service connection timed out.')); req.destroy(); });
      timer = setTimeout(() => { done(Error('Voice generation exceeded two minutes. Try a shorter script.')); req.destroy(); response?.destroy(); }, timeoutMs);
      signal?.addEventListener('abort', cancel, { once: true });
      req.end(body);
    } catch (error) { done(error); }
  });
}
function runEncoder(executable, args, { signal, progress, timeoutMs = 120000, spawnProcess = spawn } = {}) {
  return new Promise((resolve, reject) => {
    let child, timer, error = '', pending = '', reason = null, settled = false;
    const finish = e => { if (settled) return; settled = true; clearTimeout(timer); signal?.removeEventListener('abort', cancel); e ? reject(e) : resolve(); };
    const kill = message => { reason ||= Error(message); try { child?.kill('SIGKILL'); } catch (_) {} };
    const cancel = () => kill('Operation cancelled. The original recording is unchanged.');
    try {
      aborted(signal);
      child = spawnProcess(executable, ['-hide_banner', '-nostdin', '-loglevel', 'error', '-nostats', '-progress', 'pipe:1', ...args], { shell: false, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
      child.once('error', e => finish(Error('Audio encoder could not start: ' + (e.code || 'unknown error'))));
      child.stderr.on('data', c => { error = (error + c.toString()).slice(-1600); });
      child.stdout.on('data', c => {
        pending = (pending + c.toString()).slice(-10000);
        const lines = pending.split(/\r?\n/); pending = lines.pop();
        for (const line of lines) if (line.startsWith('out_time_us=')) { const n = Number(line.slice(12)) / 1000000; if (Number.isFinite(n)) progress?.(Math.max(0, n)); }
      });
      child.once('close', code => finish(reason || (code === 0 ? null : Error('Audio processing failed. Check the source audio and available disk space.' + (error.includes('No space left') ? ' The disk is full.' : '')))));
      signal?.addEventListener('abort', cancel, { once: true });
      timer = setTimeout(() => kill('Audio processing timed out. The original recording is unchanged.'), timeoutMs);
    } catch (e) { finish(e); }
  });
}
// Bundled, read-only music lives in the existing project assets directory.
// Resolve from this module, never from the working directory. Electron fs reads
// the same path inside app.asar; musicFile() stages bytes outside it for FFmpeg.
// Older projects used assets/Audio. Retain that casing as a compatibility path,
// without relying on the former src/audio-assets directory.
function bundledMusicFile(moduleDirectory = __dirname) {
  const assets = path.resolve(moduleDirectory, '..', 'assets');
  const preferred = path.join(assets, 'audio', 'warm-welcome.mp3');
  const existingCase = path.join(assets, 'Audio', 'warm-welcome.mp3');
  return !fs.existsSync(preferred) && fs.existsSync(existingCase) ? existingCase : preferred;
}
class BrandingService extends EventEmitter {
  constructor({ directory, recordings, encoder = bundledEncoder, speak = requestSpeech, builtin = bundledMusicFile(), run = runEncoder }) {
    super(); Object.assign(this, { directory, recordings, encoder, speak, builtin, run });
    this.assets = path.join(directory, 'assets'); this.exports = path.join(recordings.directory, 'Exports');
    this.configFile = path.join(directory, 'recording-settings.json'); this.job = null; this.cache = null; this.loading = null; this.durationCache = new Map(); this.previewServer = new RecordingPreview();
  }
  async load() {
    if (this.cache) return this.cache;
    // Serialize the first load so two windows cannot race the one-time migration.
    if (!this.loading) this.loading = this.loadSettings().finally(() => { this.loading = null; });
    return this.loading;
  }
  async loadSettings() {
    await fsp.mkdir(this.assets, { recursive: true, mode: 0o700 });
    try {
      await regular(this.configFile, 150000);
      const data = JSON.parse(await fsp.readFile(this.configFile, 'utf8'));
      if (data.schema !== 1 || !Number.isSafeInteger(data.revision) || data.revision < 0) throw Error('Invalid saved recording settings.');
      const upgrade = !(Number.isInteger(data.samplesRevision) && data.samplesRevision >= C.SAMPLES_REVISION);
      const migrated = upgrade ? C.migrateLegacyDefaults(data.config, data.clips) : { config: C.validate(data.config) };
      const next = { schema: 1, revision: data.revision, samplesRevision: Math.max(C.SAMPLES_REVISION, data.samplesRevision || 0),
        config: migrated.config, clips: data.clips || {}, customMusic: data.customMusic || null };
      if (upgrade) {
        if (next.revision === Number.MAX_SAFE_INTEGER) throw Error('Saved revision is out of range.');
        // Preserve the complete previous settings and clip references before migration.
        // The audio files themselves are never trimmed, deleted or regenerated here.
        try { await fsp.copyFile(this.configFile, path.join(this.directory, 'recording-settings.before-script-only-v1.json'), fs.constants.COPYFILE_EXCL); }
        catch (e) { if (e.code !== 'EEXIST') throw e; }
        next.revision++; await this.persist(next);
      } else this.cache = next;
    } catch (e) {
      if (e.code !== 'ENOENT') throw Error('Recording settings could not be read. Restore recording-settings.json from backup; the existing file was not replaced.');
      this.cache = { schema: 1, revision: 0, samplesRevision: C.SAMPLES_REVISION, config: C.defaults(), clips: {}, customMusic: null };
    }
    return this.cache;
  }
  async asset(name) { if (!ASSET.test(name || '')) throw Error('Invalid prepared audio asset.'); const file = path.join(this.assets, name); await regular(file, MAX_FILE); return file; }
  async persist(next) { await atomic(this.configFile, JSON.stringify(next, null, 2) + '\n'); this.cache = next; }
  assertFree() { if (this.job) throw Error('An audio operation is already running. Wait for it to finish or cancel it.'); }
  async operation(kind, work, requestId = null) {
    validateRequestId(requestId); this.assertFree(); const controller = new AbortController();
    const job = { id: crypto.randomUUID(), requestId, kind, phase: kind, seconds: 0, controller }; this.job = job;
    const progress = (phase, seconds = 0) => { job.phase = phase; job.seconds = seconds; this.emit('progress', { kind, phase, seconds, requestId }); };
    job.done = (async () => { try { return await work(controller.signal, progress); } finally { if (this.job === job) this.job = null; this.emit('progress', { kind, phase: 'idle', seconds: 0, requestId }); } })();
    return job.done;
  }
  async cancel({ requestId } = {}) { validateRequestId(requestId); const job = this.job; if (requestId && job?.requestId !== requestId) return { cancelled: false }; if (job) { job.controller.abort(); await job.done.catch(() => {}); } return { cancelled: !!job }; }
  async state() {
    const d = await this.load(), clips = {};
    for (const slot of ['intro', 'outro']) {
      const clip = d.clips[slot], needsRegeneration = clip?.source === 'openai' && clip.speechFormat !== C.SPEECH_FORMAT;
      let voiceCurrent = false, ready = false, error = '';
      try {
        if (clip) { await this.asset(clip.file); voiceCurrent = !needsRegeneration && clip.signature === C.textSignature(d.config, slot, clip.source === 'imported');
          if (voiceCurrent && clip.mix && clip.mix.signature === C.mixSignature(d.config, clip, d.customMusic)) { await this.asset(clip.mix.file); ready = true; } }
      } catch (e) { error = e.message; }
      clips[slot] = { exists: !!clip, voiceCurrent, ready, needsRegeneration, seconds: clip?.mix?.seconds || clip?.seconds || 0,
        source: clip?.source || null, generatedAt: clip?.generatedAt || null,
        state: ready ? 'Ready to reuse' : !clip ? 'Not prepared' : needsRegeneration ? 'Generate voice again — remove old spoken preface' : !voiceCurrent ? 'Text or voice changed — prepare again' : 'Mix settings changed — rebuild mix', error };
    }
    return { revision: d.revision, config: clone(d.config), samples: { introText: C.INTRO, outroText: C.OUTRO }, clips, hasCustomMusic: !!d.customMusic,
      customMusicName: d.customMusic?.displayName || '', job: this.job ? { kind: this.job.kind, phase: this.job.phase, seconds: this.job.seconds } : null };
  }
  async save({ config, revision }) {
    return this.operation('save-settings', async () => {
      const old = await this.load();
      if (revision !== old.revision) throw Error('Recording settings changed in another window. Reload Recording settings before saving.');
      const validated = C.validate(config); await this.persist({ ...clone(old), config: validated, revision: old.revision + 1 }); return this.state();
    });
  }
  async transcode(input, signal, { music = false } = {}) {
    const file = (music ? 'music-' : 'voice-') + crypto.randomBytes(16).toString('hex') + '.wav';
    const output = path.join(this.assets, file);
    try {
      await regular(input, MAX_FILE); aborted(signal);
      const format = { '.mp3':'mp3', '.wav':'wav', '.m4a':'mov', '.aac':'aac', '.ogg':'ogg', '.flac':'flac', '.aiff':'aiff', '.aif':'aiff' }[path.extname(input).toLowerCase()];
      if (!format) throw Error('Unsupported local audio format.');
      await this.run(this.encoder(), ['-protocol_whitelist', 'file,pipe', '-f', format, '-i', input, '-map', '0:a:0', '-vn', '-sn', '-dn', '-t', '181',
        '-af', 'loudnorm=I=-18:TP=-2:LRA=7', '-ar', '44100', '-ac', '1', '-c:a', 'pcm_s16le', '-map_metadata', '-1', '-n', output], { signal });
      const seconds = wavDuration(await fsp.readFile(output));
      if (seconds > 180) throw Error('Use an audio clip no longer than three minutes.');
      return { file, seconds };
    } catch (e) { await fsp.unlink(output).catch(() => {}); throw e; }
  }
  async musicFile(data) {
    if (data.config.music === 'custom') { if (!data.customMusic) throw Error('Import background music first, or select Warm Welcome / None.'); return this.asset(data.customMusic.file); }
    if (data.config.music !== 'warm') return null;
    // Read the bundled asset through Electron fs (works in app.asar), then copy to a real path for FFmpeg.
    let bytes;
    try { bytes = await fsp.readFile(this.builtin); }
    catch (e) {
      if (e.code === 'ENOENT') throw Error('Bundled Warm Welcome music is missing. Restore assets/audio/warm-welcome.mp3 from the patch and rebuild the app, or select None / import music.');
      throw e;
    }
    const name = 'music-' + crypto.createHash('md5').update(bytes).digest('hex') + '.mp3';
    const dest = path.join(this.assets, name);
    try { await fsp.writeFile(dest, bytes, { flag: 'wx', mode: 0o600 }); } catch (e) { if (e.code !== 'EEXIST') throw e; }
    return this.asset(name);
  }
  async renderMix(data, slot, signal) {
    const clip = data.clips[slot], c = data.config;
    if (clip?.source === 'openai' && clip.speechFormat !== C.SPEECH_FORMAT) throw Error('Generate the ' + (slot === 'intro' ? 'introduction' : 'closing') + ' voice again to remove the old spoken preface. Rebuild mix cannot remove recorded speech.');
    if (!clip || clip.signature !== C.textSignature(c, slot, clip.source === 'imported')) throw Error('Prepare a current ' + (slot === 'intro' ? 'introduction' : 'closing') + ' voice first.');
    const voice = await this.asset(clip.file), music = c.music !== 'none' && c.musicVolume > 0 ? await this.musicFile(data) : null;
    const lead = music ? c.leadSeconds : 0, tail = music ? c.tailSeconds : 0, duration = clip.seconds + lead + tail;
    const name = 'mix-' + crypto.randomBytes(16).toString('hex') + '.mp3', output = path.join(this.assets, name);
    const args = ['-protocol_whitelist', 'file,pipe', '-i', voice];
    if (music) args.push('-stream_loop', '-1', '-protocol_whitelist', 'file,pipe', '-i', music);
    let filter = `[0:a]aresample=44100,aformat=channel_layouts=mono,adelay=${Math.round(lead * 1000)}:all=1,apad,atrim=duration=${duration}[v]`;
    if (music) filter += `;[1:a]aresample=44100,aformat=channel_layouts=mono,atrim=duration=${duration},volume=${c.musicVolume / 100},afade=t=in:d=${Math.min(1, duration / 4)},afade=t=out:st=${Math.max(0, duration - Math.max(.5, tail))}:d=${Math.max(.5, tail)}[m];[v][m]amix=inputs=2:duration=first:normalize=0,alimiter=limit=0.95:level=0[out]`;
    else filter += ';[v]anull[out]';
    try {
      await this.run(this.encoder(), [...args, '-filter_complex', filter, '-map', '[out]', '-c:a', 'libmp3lame', '-b:a', '128k', '-ar', '44100', '-ac', '1', '-map_metadata', '-1',
        '-metadata', `title=AnchorCast ${slot === 'intro' ? 'Introduction' : 'Closing'}`, '-n', output], { signal });
      await regular(output, MAX_FILE); aborted(signal);
      return { file: name, seconds: duration, signature: C.mixSignature(c, clip, data.customMusic) };
    } catch (e) { await fsp.unlink(output).catch(() => {}); throw e; }
  }
  async generate({ slot, consent, revision }, key) {
    C.slotName(slot);
    if (consent !== true) throw Error('Confirm that the script may be sent to OpenAI before generating speech.');
    return this.operation('generate-' + slot, async (signal, progress) => {
      const old = await this.load(); if (old.revision !== revision) throw Error('Save or reload your Recording settings first.');
      const next = clone(old), payload = C.speechPayload(next.config, slot);
      // Resolve/validate key before any provider request. Never send sermon audio.
      if (!key || /[\r\n]/.test(key)) throw Error('Save an OpenAI API key in Audio & AI → General first.');
      this.encoder(); // Fail before a paid network call when the local encoder is missing.
      progress('Generating voice — sending this script to OpenAI');
      const bytes = await this.speak({ key, payload, signal }); aborted(signal);
      if (!Buffer.isBuffer(bytes) || bytes.length > MAX_FILE || bytes.length < 100) throw Error('Voice generation did not return a usable audio file.');
      const temp = path.join(this.assets, 'tts-' + crypto.randomUUID() + '.mp3');
      let voice, mix, committed = false;
      try {
        await fsp.writeFile(temp, bytes, { flag: 'wx', mode: 0o600 }); progress('Preparing voice audio');
        voice = await this.transcode(temp, signal);
        next.clips[slot] = { ...voice, source: 'openai', speechFormat: C.SPEECH_FORMAT, signature: C.textSignature(next.config, slot), generatedAt: new Date().toISOString() };
        progress('Mixing voice and background music'); mix = await this.renderMix(next, slot, signal); next.clips[slot].mix = mix;
        aborted(signal); next.revision++; await this.persist(next); committed = true; return this.state();
      } finally {
        await fsp.unlink(temp).catch(() => {});
        if (!committed) for (const asset of [voice, mix]) if (asset) await fsp.unlink(path.join(this.assets, asset.file)).catch(() => {});
      }
    });
  }
  async importVoice({ slot, input, revision }) {
    C.slotName(slot);
    return this.operation('import-' + slot, async (signal, progress) => {
      const old = await this.load(); if (revision !== old.revision) throw Error('Save or reload Recording settings first.');
      // For imported audio a blank script is acceptable, but unresolved contact tokens are not.
      const signature = C.textSignature(old.config, slot, true), next = clone(old); let voice, mix, committed = false;
      try {
        progress('Importing local narration'); voice = await this.transcode(input, signal);
        next.clips[slot] = { ...voice, source: 'imported', signature, generatedAt: new Date().toISOString() };
        progress('Mixing imported narration'); mix = await this.renderMix(next, slot, signal); next.clips[slot].mix = mix;
        aborted(signal); next.revision++; await this.persist(next); committed = true; return this.state();
      } finally { if (!committed) for (const a of [voice, mix]) if (a) await fsp.unlink(path.join(this.assets, a.file)).catch(() => {}); }
    });
  }
  async importMusic({ input, revision }) {
    return this.operation('import-music', async (signal, progress) => {
      const old = await this.load(); if (revision !== old.revision) throw Error('Save or reload Recording settings first.');
      progress('Importing background music'); let music, committed = false;
      try { music = await this.transcode(input, signal, { music: true }); const next = clone(old);
        next.customMusic = { ...music, displayName: path.basename(input).slice(0, 150) }; next.config.music = 'custom'; next.revision++;
        aborted(signal); await this.persist(next); committed = true; return this.state();
      } finally { if (music && !committed) await fsp.unlink(path.join(this.assets, music.file)).catch(() => {}); }
    });
  }
  async prepare({ slot, revision }) {
    C.slotName(slot);
    return this.operation('mix-' + slot, async (signal, progress) => {
      const old = await this.load(); if (revision !== old.revision) throw Error('Save or reload Recording settings first.'); const next = clone(old);
      progress('Rebuilding background mix — no cloud request'); const mix = await this.renderMix(next, slot, signal);
      try { aborted(signal); next.clips[slot].mix = mix; next.revision++; await this.persist(next); } catch (e) { await fsp.unlink(path.join(this.assets, mix.file)).catch(() => {}); throw e; }
      return this.state();
    });
  }
  async preview({ slot }) {
    C.slotName(slot); const s = await this.state(); if (!s.clips[slot].ready) throw Error('Prepare the current ' + slot + ' audio before previewing.');
    const d = await this.load(); return { bytes: new Uint8Array(await fsp.readFile(await this.asset(d.clips[slot].mix.file))) };
  }
  async previewMusic() {
    const d = await this.load(), file = await this.musicFile(d); if (!file) throw Error('Select a music track first.');
    return { bytes: new Uint8Array(await fsp.readFile(file)), mime: file.endsWith('.wav') ? 'audio/wav' : 'audio/mpeg' };
  }
  async recordingDuration(raw, stat, signal, progress = () => {}) {
    const identity = fileIdentity(stat), cached = this.durationCache.get(raw);
    if (cached?.identity === identity) return cached.seconds;
    let seconds = 0;
    progress('Reading recording duration');
    // Decode to a null sink. This gives a reliable duration even for the original
    // pipe-written MP3, which may have no Xing seek/duration header. No PCM buffer
    // is retained. Reuse the result only while the immutable source still matches.
    await this.run(this.encoder(), ['-protocol_whitelist', 'file,pipe', '-f', 'mp3', '-i', raw,
      '-map', '0:a:0', '-vn', '-sn', '-dn', '-ar', '44100', '-ac', '1', '-f', 'null', '-'],
      { signal, timeoutMs: 3600000, progress: n => { seconds = Math.max(seconds, n); progress('Reading recording duration', n); } });
    aborted(signal);
    if (!Number.isFinite(seconds) || seconds <= 0) throw Error('Could not read a usable recording duration.');
    if (fileIdentity(await regular(raw)) !== identity) throw Error('The original recording changed while its duration was being read. Open it again.');
    if (this.durationCache.size >= 20) this.durationCache.delete(this.durationCache.keys().next().value);
    this.durationCache.set(raw, { identity, seconds }); return seconds;
  }
  async recordingInfo({ name, requestId } = {}) {
    return this.operation('inspect-recording', async (signal, progress) => {
      const raw = await this.recordings.sourceFile(name), stat = await regular(raw);
      const seconds = await this.recordingDuration(raw, stat, signal, progress); aborted(signal);
      const preview = await this.previewServer.open(raw, stat);
      try { aborted(signal); return { name, seconds, bytes: stat.size, sourceIdentity: fileIdentity(stat), ...preview }; }
      catch (e) { this.previewServer.release(preview.previewToken); throw e; }
    }, requestId);
  }
  releaseRecordingPreview({ previewToken } = {}) {
    if (typeof previewToken !== 'string' || !/^[a-f0-9]{64}$/.test(previewToken)) throw Error('Invalid recording preview.');
    return this.previewServer.release(previewToken);
  }
  async dispose() { await this.cancel(); await this.previewServer.dispose(); }
  async compose({ name, intro = false, outro = false, trim = null, sourceIdentity = null, requestId = null }) {
    if (typeof intro !== 'boolean' || typeof outro !== 'boolean' || (!intro && !outro && !trim)) throw Error('Select introduction, closing, or a trim selection. The original is already saved.');
    if (sourceIdentity != null && (typeof sourceIdentity !== 'string' || !/^[a-f0-9]{64}$/.test(sourceIdentity))) throw Error('Invalid original recording identity.');
    return this.operation('export', async (signal, progress) => {
      const s = await this.state(), data = clone(await this.load());
      for (const [slot, selected] of [['intro', intro], ['outro', outro]]) if (selected && !s.clips[slot].ready) throw Error(`The ${slot} is not ready or has changed. Prepare it in Recording settings before adding it.`);
      const raw = await this.recordings.sourceFile(name), rawStat = await regular(raw), identity = fileIdentity(rawStat);
      if (sourceIdentity && sourceIdentity !== identity) throw Error('The original recording changed since preview. Close this dialog and open the recording again.');
      let selection = null;
      if (trim) selection = validateTrim(trim, await this.recordingDuration(raw, rawStat, signal, progress));
      if (!intro && !outro && !selection?.changed) throw Error('Select introduction, closing, or change the trim selection. The original is already saved.');
      const inputs = []; if (intro) inputs.push(await this.asset(data.clips.intro.mix.file));
      const sermonIndex = inputs.length; inputs.push(raw); if (outro) inputs.push(await this.asset(data.clips.outro.mix.file));
      aborted(signal); await fsp.mkdir(this.exports, { recursive: true, mode: 0o700 });
      const stamp = new Date().toISOString().replace(/[-:]/g, '').replace('T', '-').slice(0, 15);
      const outputName = `Message-${stamp}-${crypto.randomBytes(16).toString('hex')}.mp3`, dest = path.join(this.exports, outputName);
      const temp = dest + '.tmp.mp3', args = [];
      for (const input of inputs) args.push('-protocol_whitelist', 'file,pipe', '-f', 'mp3', '-i', input);
      const chains = inputs.map((_f, i) => {
        let chain = `[${i}:a]aresample=44100,aformat=sample_fmts=fltp:channel_layouts=mono`;
        // Trim only the sermon input, at decoded-sample boundaries; voiceovers and
        // their music remain intact. Reset timestamps before concatenation.
        if (i === sermonIndex && selection?.changed) chain += `,atrim=start_sample=${Math.round(selection.startSeconds * 44100)}:end_sample=${Math.round(selection.endSeconds * 44100)}`;
        chain += ',asetpts=PTS-STARTPTS';
        if (i < inputs.length - 1 && data.config.gapSeconds) chain += ',apad=pad_dur=' + data.config.gapSeconds;
        return chain + `[a${i}]`;
      });
      chains.push(inputs.map((_f, i) => `[a${i}]`).join('') + `concat=n=${inputs.length}:v=0:a=1[out]`);
      try {
        progress('Creating MP3 — original kept');
        await this.run(this.encoder(), [...args, '-filter_complex', chains.join(';'), '-map', '[out]', '-c:a', 'libmp3lame', '-b:a', '128k', '-ar', '44100', '-ac', '1', '-map_metadata', '-1',
          '-metadata', 'title=Sermon message', '-metadata', 'comment=Prepared in AnchorCast; original recording preserved.', '-n', temp], { signal, timeoutMs: 3600000, progress: seconds => progress('Creating MP3 — original kept', seconds) });
        aborted(signal); const stat = await regular(temp); if (stat.size < 200) throw Error('The finished MP3 is empty.');
        if (fileIdentity(await regular(raw)) !== identity) throw Error('The original recording changed during export. Try again.');
        // Final commit has no intervening asynchronous step after the abort check.
        // A cancellation received before commit removes only the unfinished copy.
        aborted(signal); fs.renameSync(temp, dest);
        return { saved: true, name: outputName, path: dest, originalName: name, bytes: stat.size, intro, outro,
          trim: selection?.changed ? selection : null, sourceIdentity: identity };
      } finally { await fsp.unlink(temp).catch(() => {}); }
    }, requestId);
  }
  async exportFile(name) { if (!EXPORT.test(name || '')) throw Error('Invalid finished message.'); const file = path.join(this.exports, name); await regular(file); return file; }
}
module.exports = { bundledMusicFile, BrandingService, requestSpeech, runEncoder, wavDuration, regular, ASSET, EXPORT, MAX_FILE };

'use strict';
// Short-lived, read-only capability URLs for browser audio seeking. No arbitrary
// paths, directory browsing, uploads, or externally bound listener. Large sermons
// are streamed with byte-range support, never copied through renderer IPC.
const http = require('node:http');
const fs = require('node:fs'), fsp = fs.promises;
const crypto = require('node:crypto');
const { fileIdentity } = require('./recording-trim');
const LEASE_MS = 2 * 60 * 60 * 1000;
function byteRange(header, size) {
  if (!header) return { start: 0, end: size - 1, partial: false };
  const m = /^bytes=(\d*)-(\d*)$/.exec(header);
  if (!m || (!m[1] && !m[2])) return null;
  let start, end;
  if (!m[1]) {
    const suffix = Number(m[2]);
    if (!Number.isSafeInteger(suffix) || suffix <= 0) return null;
    start = Math.max(0, size - suffix); end = size - 1;
  } else {
    start = Number(m[1]); end = m[2] ? Number(m[2]) : size - 1;
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || start >= size) return null;
    end = Math.min(size - 1, end);
  }
  return { start, end, partial: true };
}
class RecordingPreview {
  constructor({ leaseMs = LEASE_MS } = {}) {
    this.leaseMs = leaseMs; this.leases = new Map(); this.server = null; this.starting = null; this.port = null;
    this.timer = null; this.disposed = false;
  }
  async listen() {
    if (this.disposed) throw Error('Audio preview is closed.');
    if (this.port) return;
    if (!this.starting) this.starting = new Promise((resolve, reject) => {
      const server = http.createServer((req, res) => { void this.serve(req, res).catch(() => { if (!res.headersSent) res.writeHead(500); res.end(); }); });
      this.server = server;
      server.once('error', reject);
      server.listen(0, '127.0.0.1', () => {
        server.removeListener('error', reject); server.on('error', () => {});
        this.port = server.address().port; server.unref();
        this.timer = setInterval(() => this.sweep(), Math.min(this.leaseMs, 60000)); this.timer.unref();
        resolve();
      });
    }).finally(() => { this.starting = null; });
    await this.starting;
  }
  sweep() { const now = Date.now(); for (const [token, lease] of this.leases) if (lease.until < now) this.release(token); }
  async open(file, stat) {
    await this.listen(); this.sweep();
    while (this.leases.size >= 4) this.release(this.leases.keys().next().value);
    const token = crypto.randomBytes(32).toString('hex');
    this.leases.set(token, { file, identity: fileIdentity(stat), until: Date.now() + this.leaseMs, streams: new Set() });
    return { previewToken: token, previewURL: `http://127.0.0.1:${this.port}/recording/${token}` };
  }
  release(token) {
    const lease = this.leases.get(token); if (!lease) return { released: false };
    this.leases.delete(token); for (const stream of lease.streams) stream.destroy();
    return { released: true };
  }
  async serve(req, res) {
    res.setHeader('Cache-Control', 'private, no-store'); res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Referrer-Policy', 'no-referrer');
    const deny = (code, extra = {}) => { res.writeHead(code, extra); res.end(); };
    if (req.headers.host !== `127.0.0.1:${this.port}` || req.socket.remoteAddress !== '127.0.0.1') return deny(403);
    if (req.method !== 'GET' && req.method !== 'HEAD') return deny(405, { Allow: 'GET, HEAD' });
    const m = /^\/recording\/([a-f0-9]{64})$/.exec(req.url || ''); this.sweep();
    const lease = m && this.leases.get(m[1]); if (!lease) return deny(404);
    let handle;
    try {
      const flags = fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW || 0);
      const before = await fsp.lstat(lease.file);
      if (!before.isFile() || before.isSymbolicLink() || fileIdentity(before) !== lease.identity) return deny(409);
      handle = await fsp.open(lease.file, flags); const stat = await handle.stat();
      if (!stat.isFile() || fileIdentity(stat) !== lease.identity) { await handle.close(); return deny(409); }
      if (!this.leases.has(m[1]) || res.destroyed) { await handle.close(); return; }
      const range = byteRange(req.headers.range, stat.size);
      if (!range) { await handle.close(); return deny(416, { 'Content-Range': `bytes */${stat.size}` }); }
      lease.until = Date.now() + this.leaseMs;
      const headers = { 'Content-Type': 'audio/mpeg', 'Accept-Ranges': 'bytes', 'Content-Length': range.end - range.start + 1 };
      if (range.partial) headers['Content-Range'] = `bytes ${range.start}-${range.end}/${stat.size}`;
      res.writeHead(range.partial ? 206 : 200, headers);
      if (req.method === 'HEAD') { await handle.close(); res.end(); return; }
      const stream = handle.createReadStream({ start: range.start, end: range.end, autoClose: true }); handle = null;
      lease.streams.add(stream); stream.on('close', () => lease.streams.delete(stream));
      stream.on('error', () => res.destroy()); res.on('close', () => stream.destroy());
      stream.pipe(res);
    } catch (_) { await handle?.close().catch(() => {}); if (!res.headersSent) deny(404); else res.destroy(); }
  }
  async dispose() {
    this.disposed = true; await this.starting?.catch(() => {}); clearInterval(this.timer);
    for (const token of [...this.leases.keys()]) this.release(token);
    if (this.server) { const s = this.server; this.server = null; this.port = null; await new Promise(r => { s.close(r); s.closeAllConnections?.(); }); }
  }
}
module.exports = { RecordingPreview, byteRange };

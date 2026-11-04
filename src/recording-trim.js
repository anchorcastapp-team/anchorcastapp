'use strict';
// Pure, shared validation for non-destructive start/end trimming. Times are seconds.
const MIN_KEEP_SECONDS = 0.1;
function fileIdentity(stat) {
  const crypto = require('node:crypto');
  return crypto.createHash('sha256').update([stat.dev, stat.ino, stat.size, stat.mtimeMs, stat.ctimeMs].join(':')).digest('hex');
}
function validateTrim(trim, seconds) {
  if (!Number.isFinite(seconds) || seconds <= 0) throw Error('The recording has no usable audio duration.');
  if (trim === undefined || trim === null) return { startSeconds: 0, endSeconds: seconds, keptSeconds: seconds, changed: false };
  if (typeof trim !== 'object' || Array.isArray(trim)) throw Error('Invalid trim selection.');
  const { startSeconds, endSeconds } = trim;
  if (typeof startSeconds !== 'number' || typeof endSeconds !== 'number' || !Number.isFinite(startSeconds) || !Number.isFinite(endSeconds)) throw Error('Start and end must be finite time values.');
  if (startSeconds < 0 || endSeconds <= 0) throw Error('Start and end must be within the original recording.');
  // UI displays milliseconds. Accommodate rounding only, never a materially out-of-bounds end.
  if (endSeconds > seconds + 0.0011 || startSeconds >= seconds) throw Error('The trim selection extends beyond the original recording.');
  const end = Math.min(seconds, endSeconds), kept = end - startSeconds;
  if (kept < MIN_KEEP_SECONDS - 0.000001) throw Error('End must be after start. Keep at least 0.1 second of audio.');
  return { startSeconds, endSeconds: end, keptSeconds: kept, changed: startSeconds > 0.0005 || seconds - end > 0.0011 };
}
function validateRequestId(id) {
  if (id != null && (typeof id !== 'string' || !/^[A-Za-z0-9_-]{8,100}$/.test(id))) throw Error('Invalid audio operation identifier.');
  return id || null;
}
module.exports = { MIN_KEEP_SECONDS, fileIdentity, validateTrim, validateRequestId };

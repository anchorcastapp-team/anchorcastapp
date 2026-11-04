/* AnchorCast timer timing and formatting — shared by Electron and renderer windows.
 * A countdown remains running after its deadline until the operator stops it.
 * No intervals mutate the counter: each display derives its value from a baseline.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.AnchorCastTimer = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  function number(value, fallback) {
    if (value === null || value === undefined || value === '' || typeof value === 'boolean') return fallback;
    const n = Number(value);
    return Number.isFinite(n) ? n : fallback;
  }
  function normalizeMode(mode) { return mode === 'countup' ? 'countup' : 'countdown'; }
  function warningSeconds(seconds) { return seconds <= 300 ? 60 : Math.ceil(seconds * 0.1); }
  function flashSpeed(value) { return Math.max(0.2, Math.min(2, number(value, 1))); }
  function formatTime(value) {
    const seconds = Math.trunc(number(value, 0));
    const sign = seconds < 0 ? '-' : '';
    const total = Math.abs(seconds);
    const h = Math.floor(total / 3600);
    const m = String(Math.floor((total % 3600) / 60)).padStart(2, '0');
    const s = String(total % 60).padStart(2, '0');
    return sign + (h ? `${h}:${m}:${s}` : `${m}:${s}`);
  }
  // Visibility belongs to the running timer, not to a particular window. Older
  // payloads without this field remain visible. Only an explicit false hides it.
  function projectionIsVisible(state) {
    if (typeof state?.projectionVisible === 'boolean') return state.projectionVisible;
    return state?.data?.projectionVisible !== false;
  }
  function setProjectionVisibility(active, visible) {
    if (typeof visible !== 'boolean') throw new TypeError('Timer visibility must be a boolean');
    if (!active) return null;
    return { ...active, projectionVisible: visible,
      data: { ...active.data, projectionVisible: visible } };
  }
  function createState(input = {}, now = Date.now()) {
    const data = { ...input, mode: normalizeMode(input.mode),
      seconds: Math.max(0, number(input.seconds, 0)), flashSpeed: 1, projectionVisible: true };
    return { data, startedAt: now, seconds: data.seconds, mode: data.mode, projectionVisible: true,
      label: data.label || '', scale: number(data.scale, 1), position: data.position || 'edge' };
  }
  function snapshot(active, now = Date.now()) {
    if (!active) return { running: false };
    const elapsed = Math.max(0, (now - active.startedAt) / 1000);
    const remaining = active.mode === 'countdown' ? active.seconds - elapsed : elapsed;
    return { ...active, running: true, projectionVisible: projectionIsVisible(active), remaining, elapsed,
      overtime: active.mode === 'countdown' && remaining <= 0, serverNow: now };
  }
  function payload(active, now = Date.now()) {
    if (!active) return null;
    return { ...active.data, ...snapshot(active, now) };
  }
  function baseline(state = {}, now = Date.now()) {
    const data = { ...(state.data || {}), ...state };
    const mode = normalizeMode(data.mode);
    const seconds = Math.max(0, number(data.seconds, 0));
    const startedAt = number(data.startedAt, null);
    if (startedAt !== null) {
      const skew = now - number(data.serverNow, now);
      return startedAt + skew + (mode === 'countdown' ? seconds * 1000 : 0);
    }
    // Legacy payloads may contain signed remaining time but no start timestamp.
    const remaining = number(data.remaining, mode === 'countdown' ? seconds : 0);
    return mode === 'countdown' ? now + remaining * 1000 : now - Math.max(0, remaining) * 1000;
  }
  function frame(baselineMs, mode, warnSecs = 60, now = Date.now()) {
    const countdown = normalizeMode(mode) === 'countdown';
    const millis = countdown ? baselineMs - now : Math.max(0, now - baselineMs);
    // Ceiling counts down without expiring a second early. After zero it shows
    // completed overtime seconds: 00:00, then -00:01 (never -00:00).
    const seconds = countdown ? Math.ceil(millis / 1000) : Math.floor(millis / 1000);
    const overtime = countdown && millis <= 0;
    return { seconds: seconds === 0 ? 0 : seconds, text: formatTime(seconds), overtime,
      warning: countdown && !overtime && seconds <= warnSecs };
  }
  return Object.freeze({ normalizeMode, warningSeconds, flashSpeed, formatTime,
    projectionIsVisible, setProjectionVisibility, createState, snapshot, payload, baseline, frame });
});

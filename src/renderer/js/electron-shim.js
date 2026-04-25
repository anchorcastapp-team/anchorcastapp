(function() {
  'use strict';

  if (window.electronAPI && !window.electronAPI.isWeb) {
    console.log('[AnchorCast] Real Electron API detected — shim skipped');
    return;
  }

  const BASE = '';

  async function api(method, path, body) {
    const opts = { method, headers: { 'Content-Type': 'application/json' } };
    if (body !== undefined) opts.body = JSON.stringify(body);
    const r = await fetch(BASE + path, opts);
    if (!r.ok) throw new Error(`API ${path} => ${r.status}`);
    return r.json();
  }

  const _listeners = {};

  let _sseInstance = null;
  function setupSSE() {
    if (_sseInstance) { try { _sseInstance.close(); } catch(_){} }
    const es = new EventSource('/api/events');
    _sseInstance = es;
    const events = [
      'settings-saved','themes-updated','songs-saved','bible-versions-updated',
      'render-state','projection-opened','projection-closed',
      'shortcut-go-live','shortcut-next','shortcut-prev','shortcut-clear',
      'show-verse','show-song','show-media','clear-verse',
      'show-timer','stop-timer','show-caption','show-alert','http-server-started',
      'remote-present','remote-queue-add','remote-set-translation',
      'remote-present-song','remote-present-media','remote-control',
      'open-settings-modal','ndi-status','open-ndi-panel',
      'transcript-result','transcript-no-key','whisper-status',
      'menu-schedule-new','menu-schedule-save','menu-schedule-save-as',
      'menu-schedule-open','menu-schedule-export','menu-schedule-import',
      'menu-schedule-load-file','show-created-slide',
      'pres-editor-load','pres-editor-saved','show-presentation-slide',
      'theme-designer-open','song-manager-new-song','song-manager-open-song',
      'logo-overlay','logo-overlay-drag-update',
    ];
    events.forEach(ev => {
      es.addEventListener(ev, (e) => {
        let data;
        try { data = JSON.parse(e.data); } catch(_) { data = e.data; }
        if (ev === 'remote-control' && data && data.id) {
          _dispatchRemoteCmd(data);
        } else {
          const cbs = _listeners[ev] || [];
          cbs.forEach(cb => cb(data));
        }
      });
    });
    es.onerror = () => {
      try { es.close(); } catch(_){}
      _sseInstance = null;
      setTimeout(setupSSE, 3000);
    };
  }

  setupSSE();

  let _lastCmdId = 0;
  const _processedIds = new Set();

  function _dispatchRemoteCmd(cmd) {
    if (!cmd || !cmd.id) return;
    if (_processedIds.has(cmd.id)) return;
    _processedIds.add(cmd.id);
    if (_processedIds.size > 100) {
      const arr = Array.from(_processedIds).sort((a,b) => a - b);
      arr.splice(0, arr.length - 50);
      _processedIds.clear();
      arr.forEach(k => _processedIds.add(k));
    }
    if (cmd.id > _lastCmdId) _lastCmdId = cmd.id;
    const cbs = _listeners['remote-control'] || [];
    cbs.forEach(cb => cb(cmd));
  }

  setInterval(() => {
    fetch('/api/control/pending?after=' + _lastCmdId)
      .then(r => r.json())
      .then(data => {
        if (data.commands && data.commands.length) {
          data.commands.forEach(cmd => _dispatchRemoteCmd(cmd));
        }
      })
      .catch(() => {});
  }, 800);

  const noop = () => Promise.resolve(null);
  const noopFalse = () => Promise.resolve(false);

  let _projectionWindow = null;
  let _projClosePoller = null;

  function _startClosePolling() {
    if (_projClosePoller) return;
    _projClosePoller = setInterval(() => {
      if (!_projectionWindow || _projectionWindow.closed) {
        clearInterval(_projClosePoller);
        _projClosePoller = null;
        _projectionWindow = null;
        const cbs = _listeners['projection-closed'] || [];
        cbs.forEach(cb => { try { cb({}); } catch(_){} });
      }
    }, 500);
  }

  function _openProjectionWindow() {
    if (_projectionWindow && !_projectionWindow.closed) {
      _projectionWindow.focus();
      return Promise.resolve();
    }
    _projectionWindow = window.open(
      '/projection.html',
      'anchorcast-projection',
      'width=1280,height=720,menubar=no,toolbar=no,location=no,status=no,resizable=yes'
    );
    if (!_projectionWindow) {
      console.warn('[AnchorCast] Projection popup blocked — allow popups for this site');
      return Promise.resolve();
    }
    _startClosePolling();
    setTimeout(() => {
      if (_projectionWindow && !_projectionWindow.closed) {
        const cbs = _listeners['projection-opened'] || [];
        cbs.forEach(cb => { try { cb({}); } catch(_){} });
      }
    }, 400);
    return Promise.resolve();
  }

  window.electronAPI = {
    platform: navigator.platform || 'web',
    isElectron: false,
    isWeb: true,

    getSettings: () => api('GET', '/api/settings'),
    saveSettings: (s, opts) => api('POST', '/api/settings', { _settings: s, _opts: opts || {} }),
    openSettings: () => { window.open('/settings.html', '_blank'); return Promise.resolve(); },
    getSettingsOpenParams: () => Promise.resolve({}),

    getThemes: () => api('GET', '/api/themes'),
    saveThemes: (t) => api('POST', '/api/themes', t),
    openThemeDesigner: (opts) => { window.open('/theme-designer.html' + (opts?.category ? '?category=' + opts.category : ''), '_blank'); return Promise.resolve(); },
    getThemeDesignerParams: () => Promise.resolve({}),
    pickBgMedia: noop,

    getTranscripts: () => api('GET', '/api/transcripts'),
    saveTranscript: (d) => api('POST', '/api/transcripts', d),
    deleteTranscript: (id) => api('DELETE', `/api/transcripts/${id}`),
    openHistory: () => { window.open('/history.html', '_blank'); return Promise.resolve(); },
    loadDetectionReviewData: () => Promise.resolve({ detections: [], transcripts: [] }),
    saveDetectionFeedback: noop,
    saveDetectionPhrase: noop,

    geniusSearch: async ({ query } = {}) => {
      if (!query || !query.trim()) return { results: [] };
      try {
        const r = await fetch('/api/genius/search?q=' + encodeURIComponent(query.trim()));
        const data = await r.json();
        if (data.error) return { error: data.error, results: [] };
        const results = (data.hits || []).map(h => ({
          id:        h.id,
          title:     h.title     || '',
          artist:    h.artist    || '',
          thumbnail: h.thumbnail || '',
          url:       h.url       || '',
          path:      h.path      || '',
          fullTitle: h.title && h.artist ? `${h.title} by ${h.artist}` : (h.title || ''),
        }));
        return { results };
      } catch (e) {
        return { error: e.message || 'Search failed', results: [] };
      }
    },

    geniusFetchLyrics: async ({ url, artist, title } = {}) => {
      if (!artist || !title) return { error: 'Song artist and title required' };
      try {
        const params = new URLSearchParams({ artist: artist.trim(), title: title.trim() });
        const r = await fetch('/api/genius/lyrics?' + params.toString());
        const data = await r.json();
        if (data.error) return { error: data.error };
        return { lyrics: data.lyrics || '' };
      } catch (e) {
        return { error: e.message || 'Lyrics fetch failed' };
      }
    },

    getSongs: () => api('GET', '/api/songs'),
    saveSongs: (songs) => api('POST', '/api/songs', songs),
    openSongManager: () => { window.open('/song-manager.html', '_blank'); return Promise.resolve(); },
    importSongsFile: noop,
    exportSongs: noop,
    getSongLibraryInfo: () => Promise.resolve({ count: 0 }),
    backupSongLibrary: noop,
    importEasyworshipDataFolder: noop,
    importSmartSongSource: noop,

    getMedia: () => api('GET', '/api/media'),
    saveMedia: (d) => api('POST', '/api/media', d),
    importMediaFiles: noop,
    getMediaIntegrity: () => Promise.resolve({ ok: true }),
    repairMediaLinks: noop,
    clearMediaCache: noop,
    openFileLocation: noop,
    showUnsupportedDialog: noop,

    getPresentations: () => api('GET', '/api/presentations'),
    getPresentationSlides: noop,
    deletePresentation: noop,
    importPresentation: noop,
    projectPresentationSlide: (d) => api('POST', '/api/render-state', { module: 'presentation-slide', payload: d, updatedAt: Date.now() }),
    projectCreatedSlide: (d) => api('POST', '/api/render-state', { module: 'created-slide', payload: d, updatedAt: Date.now() }),
    openPresEditor: (opts) => { window.open('/presentation-editor.html' + (opts?.id ? '?id=' + opts.id : ''), '_blank'); return Promise.resolve(); },
    openPresentationEditor: (opts) => { window.open('/presentation-editor.html' + (opts?.id ? '?id=' + opts.id : ''), '_blank'); return Promise.resolve(); },
    presEditorSaved: noop,
    pickPresFile: noop,
    presentationEditorImportToLibrary: noop,
    addPresentationToScheduleFromEditor: noop,
    presBibleSearch: (q) => Promise.resolve([]),
    getCreatedPresentations: () => api('GET', '/api/created-presentations'),
    saveCreatedPresentations: (d) => api('POST', '/api/created-presentations', d),

    loadBibleData: () => api('GET', '/api/bible/load'),
    getInstalledVersions: () => api('GET', '/api/bible/installed'),
    saveBibleVersion: (trans, data) => api('POST', '/api/bible/save', { translation: trans, data }),
    deleteBibleVersion: (trans) => api('DELETE', `/api/bible/${trans}`),
    importTranslation: ({ abbrev, data }) => api('POST', '/api/bible/save', { translation: abbrev, data }),
    deleteTranslation: (trans) => api('DELETE', `/api/bible/${trans}`),

    openProjection: _openProjectionWindow,
    closeProjection: () => {
      if (_projectionWindow && !_projectionWindow.closed) {
        _projectionWindow.close();
        _projectionWindow = null;
      }
      if (_projClosePoller) { clearInterval(_projClosePoller); _projClosePoller = null; }
      const cbs = _listeners['projection-closed'] || [];
      cbs.forEach(cb => { try { cb({}); } catch(_){} });
      return Promise.resolve();
    },
    projectVerse: (d) => api('POST', '/api/render-state', { module: 'verse', payload: d, updatedAt: Date.now() }),
    projectSong: (d) => api('POST', '/api/render-state', { module: 'song', payload: d, updatedAt: Date.now() }),
    projectMedia: (d) => api('POST', '/api/render-state', { module: 'media', payload: d, updatedAt: Date.now() }),
    sendLogoOverlay: (d) => api('POST', '/api/render-state', { module: 'logo-overlay', payload: d, updatedAt: Date.now() }),
    getPresets:      ()  => api('GET', '/api/presets'),
    savePreset:      (d) => api('POST', '/api/presets', d),
    deletePreset:    (id)=> api('POST', '/api/presets/delete', { id }),
    showTimer: (d) => api('POST', '/api/render-state', { module: 'timer', payload: d, updatedAt: Date.now() }),
    stopTimer: () => api('POST', '/api/render-state', { module: 'clear', payload: null, updatedAt: Date.now() }),
    timerScale: (d) => api('POST', '/api/render-state', { module: 'timer-scale', payload: d, updatedAt: Date.now() }),
    showCaption: (d) => api('POST', '/api/render-state', { module: 'caption', payload: d, updatedAt: Date.now() }),
    showAlert: (d) => api('POST', '/api/render-state', { module: 'alert', payload: d, updatedAt: Date.now() }),
    clearProjection: () => api('POST', '/api/render-state', { module: 'clear', payload: null, updatedAt: Date.now() }),
    getDisplays: () => api('GET', '/api/displays'),
    getCurrentRenderState: () => api('GET', '/api/render-state'),

    whisperStatus: () => api('GET', '/api/whisper/status'),
    whisperStart: noop,
    whisperStop: noop,
    startWhisper: noop,
    stopWhisper: noop,
    reinforceWhisper: noop,
    setWhisperSource: noop,
    pushAudioPcm: noop,
    onTranscript: (cb) => {
      if (!_listeners['transcript-result']) _listeners['transcript-result'] = [];
      _listeners['transcript-result'].push(cb);
      return () => {
        _listeners['transcript-result'] = (_listeners['transcript-result'] || []).filter(f => f !== cb);
      };
    },

    ndiStart: noopFalse,
    ndiStop: noop,
    ndiStatus: () => Promise.resolve({ active: false }),

    getRemoteInfo: () => api('GET', '/api/remote-info'),
    getRemoteStatus: () => Promise.resolve({ enabled: false }),
    getRemoteRuntimeStatus: () => Promise.resolve({ connected: false, serverEnabled: false, lastRole: null, lastSeenAt: null, lastIp: null }),
    getNetworkAdapters: () => api('GET', '/api/network-adapters'),
    setNetworkAdapter: noop,
    toggleRemote: noop,
    startRemote: noop,
    stopRemote: noop,
    showRemoteUrl: noop,

    getAppInfo: () => api('GET', '/api/app-info'),
    openExternal: (url) => { window.open(url, '_blank'); return Promise.resolve(); },
    copyToClipboard: (text) => navigator.clipboard.writeText(text).catch(() => {}),

    exportFile: noop,
    exportTranscript: noop,
    importNotes: () => new Promise(resolve => {
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = '.txt,.md,.json';
      inp.style.cssText = 'position:fixed;opacity:0;pointer-events:none';
      document.body.appendChild(inp);
      inp.onchange = async () => {
        const file = inp.files?.[0];
        inp.remove();
        if (!file) { resolve({ success: false }); return; }
        try {
          const content = await file.text();
          resolve({ success: true, content, filePath: file.name });
        } catch(e) { resolve({ success: false, error: e.message }); }
      };
      inp.oncancel = () => { inp.remove(); resolve({ success: false }); };
      inp.click();
    }),

    saveSchedule: (d) => api('POST', '/api/settings', d).catch(() => null),
    saveScheduleAs: (d) => api('POST', '/api/settings', d).catch(() => null),
    openScheduleDialog: noop,
    loadScheduleFile: noop,
    consumePendingScheduleOpen: () => Promise.resolve(null),
    loadSchedules: () => api('GET', '/api/settings').then(s => s?.schedules || []).catch(() => []),
    deleteSchedule: noop,

    getAdaptiveDashboard: () => api('GET', '/api/adaptive/dashboard'),
    exportAdaptiveData: noop,
    importAdaptiveData: noop,
    openAdaptiveManagement: noop,
    upsertAdaptiveSpeakerProfile: noop,
    deleteAdaptiveSpeakerProfile: noop,
    upsertAdaptiveRule: noop,
    deleteAdaptiveRule: noop,
    upsertAdaptiveVocab: noop,
    deleteAdaptiveVocab: noop,
    approveAdaptiveSuggestion: noop,
    rejectAdaptiveSuggestion: noop,
    resetAdaptiveLearningData: noop,
    loadAdaptiveMemory: () => Promise.resolve(null),
    saveAdaptiveMemory: noop,
    persistTranscriptSession: noop,
    persistTranscriptSessionEnd: noop,
    persistTranscriptChunk: noop,
    persistCorrectionEvents: noop,
    loadLearningData: () => Promise.resolve(null),

    backupFullAppData: noop,
    restoreFullAppData: noop,

    openBibleManager: () => { window.open('/bible-manager.html', '_blank'); return Promise.resolve(); },

    send: (channel, ...args) => {
      const cbs = _listeners[channel] || [];
      cbs.forEach(cb => { try { cb(...args); } catch(_){} });
      return Promise.resolve();
    },

    on: (channel, cb) => {
      if (!_listeners[channel]) _listeners[channel] = [];
      _listeners[channel].push(cb);
      return () => {
        _listeners[channel] = (_listeners[channel] || []).filter(f => f !== cb);
      };
    },
  };

  console.log('[AnchorCast] Web mode: electronAPI shim installed');
})();

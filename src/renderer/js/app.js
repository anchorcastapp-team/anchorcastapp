// AnchorCast — Main Application
// Renderer process: UI logic, state, event handling

'use strict';

// ─── STATE ───────────────────────────────────────────────────────────────────
const State = {
  isRecording:   false,
  isOnline:      false,
  isLive:        false,
  isAutoMode:    false,
  isProjectionOpen: false,
  currentTranslation: 'KJV',
  currentTheme:  'sanctuary',
  currentTab:    'book',
  currentBook:   'Genesis',
  currentChapter: 1,
  previewVerse:  null,
  liveVerse:     null,
  liveContentType: null,         // null = nothing live; set when content goes live
  verseParts:    null,   // array of text parts if verse was split
  versePartIdx:  0,      // current part index
  previewViewMode: 'list', // 'list' | 'slides'
  queue:         [],
  whisperSource: 'local',    // 'deepgram' | 'local' | 'cloud' (default local offline)
  whisperLocalReady: false,
  media:         [],
  currentMediaId: null,
  currentQueueIdx: null,  // index of currently active schedule item
  presentations:   [],    // loaded presentation list
  currentPresId:   null,  // active presentation ID
  currentPresSlides: [],  // slides for active presentation
  currentPresSlideIdx: 0, // current slide index
  detections:    [],
  transcriptLines: [],
  reviewTranscriptLineId: null,
  reviewDetectionId: null,
  inlineSuggestions: [],
  replayTimeline: [],
  replayIndex: 0,
  replayAutoPlayTimer: null,
  serviceArchiveItems: [],
  serviceArchiveSelected: null,
  autoServiceSuggestions: null,
  liveSmartSuggestions: null,
  liveSmartSuggestionsTimer: null,
  sermonIntelligence: null,
  analyticsDashboard: null,
  clipCandidates: [],
  clipSelected: null,
  settings:      {},
  themes:        [],
  currentSongTheme: null,  // active song theme ID (from settings.songTheme)
  programOverlayTextOnMedia: null,
  currentPresTheme: null,  // active presentation theme ID
  generatedNotes: null,
  recordingStartTime: null,
  remoteInfo:    null,
  // Songs
  songs:           null,  // null = not loaded yet
  currentSongId:   null,
  currentSongSlideIdx: null,
  // Schedule
  scheduleName:    null,
  // Logo overlay state
  _logoActive:     false,   // true while logo is being shown on projection
  // Media editing
  _editingMediaId: null,    // id of media item currently being edited
  _mediaImporting: false,   // true while a media import is in progress
  // Song slides
  currentSongSlides: [],    // slides for the currently selected song
};


async function loadDefaultBibleReference() {
  try {
    State.currentBook = 'Genesis';
    State.currentChapter = 1;
    State.previewVerse = { book: 'Genesis', chapter: 1, verse: 1, ref: 'Genesis 1:1' };
    State.searchQuery = 'Genesis 1';
    const si = document.getElementById('searchInput');
    const st = document.getElementById('searchTranslation');
    const gt = document.getElementById('globalTranslation');
    const nav = document.getElementById('chapterNavTitle');
    if (si) si.value = 'Genesis 1';
    if (st) st.value = State.currentTranslation || 'KJV';
    if (gt) gt.value = State.currentTranslation || 'KJV';
    if (nav) nav.textContent = 'Genesis 1';
    if (typeof renderSearchResults === 'function') {
      renderSearchResults();
      setTimeout(() => { try { highlightVerseRow(1); } catch (_) {} }, 80);
    }
  } catch (_) {}
}

// ─── INIT ─────────────────────────────────────────────────────────────────────

function syncStartupBroadcastUI() {
  try {
    State.isLive = false;
    const btn = document.getElementById('goLiveBtn');
    const txt = document.getElementById('goLiveTxt');
    const tag = document.getElementById('liveTag');
    btn?.classList?.toggle('active', false);
    if (txt) txt.textContent = 'Project Live';
    if (tag) tag.style.display = 'none';

    const liveDisplay = document.getElementById('liveDisplay');
    const liveEmpty = document.getElementById('liveEmpty');
    if (liveDisplay) {
      liveDisplay.innerHTML = '';
      liveDisplay.style.display = 'none';
    }
    if (liveEmpty) liveEmpty.style.display = 'flex';
  } catch (_) {}
}

function initializeDefaultStartupCanvas() {
  try {
    // Reset live / preview selections to the required first-launch canvas
    State.isLive = false;
    State.liveVerse = null;
    State.liveContentType = null;       // FIX: null when resetting
    State.currentSongId = null;
    State.currentSongSlideIdx = null;
    State.currentMediaId = null;
    State.currentPresId = null;
    State.currentPresSlides = [];   // FIX: empty array not null (avoids .length crash)
    State.currentQueueIdx = null;
    State.scheduleName = null;
    State.queue = [];

    syncStartupBroadcastUI();

    // Return preview column to Program Preview
    try { backToVersePreview(); } catch (_) {}
    State.currentTab = 'book';
    try { switchTab('book'); } catch (_) {}

    // Force default Bible startup state
    State.currentBook = 'Genesis';
    State.currentChapter = 1;
    State.previewVerse = { book: 'Genesis', chapter: 1, verse: 1, ref: 'Genesis 1:1' };
    State.searchQuery = 'Genesis 1';

    const searchInput = document.getElementById('searchInput');
    const navTitle = document.getElementById('chapterNavTitle');
    if (searchInput) searchInput.value = 'Genesis 1';
    if (navTitle) navTitle.textContent = 'Genesis 1';

    // Ensure empty schedule panel on normal launch
    try { renderQueue(); } catch (_) {}
    try { updateScheduleNameBar(); } catch (_) {}

    // Ensure search + preview reflect Genesis 1 on initial canvas
    try { renderSearchResults(); } catch (_) {}
    setTimeout(() => {
      try { highlightVerseRow(1); } catch (_) {}
      try { syncBibleListHighlightAndScroll(); } catch (_) {}
    }, 80);
  } catch (_) {}
}



function _applyImportedLicenseFileData(r) {
  try {
    if (!r || !r.success) return;
    const nameInput = document.getElementById('nameInput');
    const emailInput = document.getElementById('emailInput');
    const keyInput = document.getElementById('keyInput');
    const res = document.getElementById('result');

    if (nameInput && r.fullName) nameInput.value = r.fullName;
    if (emailInput && r.email) emailInput.value = r.email;
    if (keyInput && r.key) keyInput.value = r.key;

    const hwId = document.getElementById('hwId')?.textContent || '';
    if (r.hardwareId && hwId && hwId !== r.hardwareId) {
      if (res) {
        res.style.color = '#f39c12';
        res.textContent = 'License file loaded, but the hardware ID inside the file does not match this computer.';
      }
      return;
    }

    if (res) {
      res.style.color = '#2ecc71';
      res.textContent = 'License file loaded. Review the fields and click Register AnchorCast.';
    }
  } catch (_) {}
}

async function _consumePendingLicenseImportIntoWindow() {
  try {
    const r = await window.electronAPI?.consumePendingLicenseImport?.();
    if (!r || r.empty || !r.success) return;
    _applyImportedLicenseFileData(r);
  } catch (_) {}
}

async function init() {
  // Load settings
  if (window.electronAPI) {
    State.settings = await window.electronAPI.getSettings();
    // Always start in LOCAL transcription + OFFLINE detection mode on load
    State.settings.onlineMode    = false;
    State.settings.whisperSource = 'local';
    applySettings(State.settings);
    State._appInitComplete = true;  // prevent applySettings from resetting mode on subsequent calls
    setupElectronEvents();
    await refreshTranslationDropdowns();
    try { await loadBibleData(); } catch (_) {}
    if (!BibleDB?.translations?.[State.currentTranslation]) State.currentTranslation = 'KJV';
    await loadDefaultBibleReference();
  } else {
    try {
      const r = await fetch('/api/settings');
      if (r.ok) State.settings = await r.json();
    } catch (_) {}
  }

  // Load themes into State so custom theme data is available for projection
  if (window.electronAPI?.getThemes) {
    window.electronAPI.getThemes().then(themes => {
      State.themes = themes || [];
      refreshThemeSwatches();
    }).catch(() => {});
  }

  // Init AI detection
  if (window.AIDetection) {
    AIDetection.init(State.settings.apiKey || '', handleDetection);
  }

  // Init Adaptive Transcript Memory
  if (window.TranscriptMemory) {
    TranscriptMemory.init().catch(() => {});
  }

  // Load remote server state for button indicator
  if (window.electronAPI?.getRemoteInfo) {
    window.electronAPI.getRemoteInfo().then(info => {
      State.remoteInfo = info;
      updateRemoteBtn(info);
    }).catch(() => {});
  }

  // Render initial search view
  renderSearchResults();
  try { syncBibleListHighlightAndScroll(); } catch (_) {}

  // Bind all UI events
  bindEvents();

  // Force the required first-launch default canvas.
  // If a schedule/file was explicitly opened, the pending launcher can still override afterward.
  initializeDefaultStartupCanvas();

  setTimeout(() => { _checkPendingScheduleLaunch(); }, 150);

  setTimeout(() => maybeShowGetStarted(), 600);

  console.log('[AnchorCast] Ready');
}


function updateOnlineUI() {
  const onlineBtn  = document.getElementById('onlineToggle');
  const offlineBtn = document.getElementById('offlineToggle');
  const pill       = document.getElementById('modeToggle');
  const label      = document.getElementById('modeLabel');
  if (onlineBtn)  onlineBtn.classList.toggle('active',  !!State.isOnline);
  if (offlineBtn) offlineBtn.classList.toggle('active', !State.isOnline);
  if (pill)  pill.className    = 'status-pill ' + (State.isOnline ? 'online' : 'offline');
  if (label) label.textContent = State.isOnline ? 'Online' : 'Offline';
}

// ─── SETTINGS ─────────────────────────────────────────────────────────────────
function applySettings(s) {
  if (!s) return;
  if (s.translation) {
    State.currentTranslation = s.translation;
    const gt = document.getElementById('globalTranslation');
    const st = document.getElementById('searchTranslation');
    if (gt) gt.value = s.translation;
    if (st) st.value = s.translation;
  }
  if (s.theme) setTheme(s.theme);
  if (s.songTheme) State.currentSongTheme = s.songTheme;
  if (s.presTheme) State.currentPresTheme = s.presTheme;
  // Store scripture ref position (used by syncProjection)
  State.settings = { ...State.settings, ...s };
  State.programOverlayTextOnMedia = !!State.settings.overlayTextOnMedia;
  updateOverlayTextOnMediaUI();
  // Only update online/whisper if applySettings is called at init time
  // (State.isOnline stays false and whisperSource stays 'local' from init defaults)
  // For subsequent calls (settings-saved, preset-load) — don't touch these values
  // so the user's current session choices are preserved
  if (!State._appInitComplete) {
    // First call from init() — apply startup defaults already set in init()
    State.isOnline      = false;
    State.whisperSource = 'local';
  }
  if (window.electronAPI?.setWhisperSource) {
    window.electronAPI.setWhisperSource(State.whisperSource).catch(() => {});
  }
  if (window.AIDetection) AIDetection.setEnabled(State.isOnline);
  if (typeof updateOnlineUI === 'function') updateOnlineUI();
  if (s.displayMode === 'auto') setDisplayMode(true);
  // Apply adaptive transcript memory settings
  if (window.TranscriptMemory) {
    TranscriptMemory.enabled = s.adaptiveEnabled !== false;
    const btn = document.getElementById('tmToggleBtn');
    if (btn) {
      btn.textContent  = TranscriptMemory.enabled ? '🧠 ON' : '🧠 OFF';
      btn.style.color  = TranscriptMemory.enabled ? 'var(--gold)' : 'var(--text-dim)';
    }
  }
  if (s.apiKey && window.AIDetection) {
    AIDetection.init(s.apiKey, handleDetection);
  }
}


function overlayTextOnMediaEnabled(kind = null) {
  const master = !!(State.programOverlayTextOnMedia ?? State.settings?.overlayTextOnMedia);
  if (!master) return false;
  if (kind === 'song') return State.settings?.overlayTextOnMediaSongs !== false;
  if (kind === 'scripture') return State.settings?.overlayTextOnMediaScripture !== false;
  return true;
}

function updateOverlayTextOnMediaUI() {
  const btn = document.getElementById('overlayTextOnMediaToggle');
  const on = overlayTextOnMediaEnabled();
  if (btn) {
    btn.classList.toggle('active', on);
    btn.style.background = on ? 'rgba(201,168,76,.22)' : '';
    btn.style.borderColor = on ? 'rgba(201,168,76,.45)' : '';
    btn.textContent = on ? '🖼＋T On' : '🖼＋T Off';
    btn.title = on ? 'Overlay on media is ON' : 'Overlay on media is OFF';
  }
  const songBtn = document.getElementById('overlaySongToggle');
  if (songBtn) {
    const songOn = overlayTextOnMediaEnabled('song');
    songBtn.classList.toggle('active', songOn);
    songBtn.style.background = songOn ? 'rgba(201,168,76,.22)' : '';
    songBtn.title = songOn ? 'Song overlay enabled' : 'Song overlay disabled';
  }
  const scrBtn = document.getElementById('overlayScriptureToggle');
  if (scrBtn) {
    const scrOn = overlayTextOnMediaEnabled('scripture');
    scrBtn.classList.toggle('active', scrOn);
    scrBtn.style.background = scrOn ? 'rgba(201,168,76,.22)' : '';
    scrBtn.title = scrOn ? 'Scripture overlay enabled' : 'Scripture overlay disabled';
  }
  const ltBtn = document.getElementById('overlayLowerThirdToggle');
  if (ltBtn) {
    const ltOn = !!State.settings?.overlayTextOnMediaLowerThird;
    ltBtn.classList.toggle('active', ltOn);
    ltBtn.style.background = ltOn ? 'rgba(201,168,76,.22)' : '';
    ltBtn.title = ltOn ? 'Lower third mode enabled' : 'Lower third mode disabled';
  }
}

function getCurrentBackgroundMediaItem() {
  if (State.currentMediaId == null) return null;
  return (State.media || []).find(m => String(m.id) === String(State.currentMediaId)) || null;
}

function _sameMediaIdentity(a, b) {
  if (!a || !b) return false;
  return String(a.type || '') === String(b.type || '') && String(a.path || '') === String(b.path || '');
}

function renderLiveBackgroundMedia(container, item) {
  if (!container || !item) return;
  const sameItem = _sameMediaIdentity(container.__mediaItem, item);
  container.style.background = '#000';
  container.__mediaItem = { ...item };
  if (item.type === 'audio') {
    if (!sameItem) {
      container.innerHTML = `<div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;background:radial-gradient(ellipse at 50% 50%, #1a0a2e 0%, #000 100%);color:#fff"><div style="font-size:42px">🎵</div><div style="font-size:12px;color:#ddd;text-align:center;padding:0 24px">${escapeHtml(item.title || item.name || '')}</div></div>`;
    }
    return;
  }
  if (item.type === 'video') {
    let vid = sameItem ? container.querySelector('video') : null;
    if (!vid) {
      container.innerHTML = '';
      vid = document.createElement('video');
      vid.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;display:block;object-fit:' + _mediaObjectFit(item);
      vid.autoplay = true;
      vid.setAttribute('playsinline', '');
      container.appendChild(vid);
    }
    vid.style.objectFit = _mediaObjectFit(item);
    vid.loop = item.loop !== false;
    vid.muted = true;
    const nextSrc = toMediaUrl(item.path);
    if (!sameItem || !vid.src || !vid.src.endsWith(encodeURI(String(item.path || '')).replace(/#/g, '%23'))) {
      vid.src = nextSrc;
      vid.load();
    }
    vid.play().catch(() => {});
    return;
  }
  let img = sameItem ? container.querySelector('img') : null;
  if (!img) {
    container.innerHTML = '';
    img = document.createElement('img');
    img.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;display:block;object-fit:' + _mediaObjectFit(item);
    container.appendChild(img);
  }
  img.style.objectFit = _mediaObjectFit(item);
  const nextSrc = toMediaUrl(item.path);
  if (!sameItem) img.src = nextSrc;
}

function getOverlayTextPrefs(kind = null) {
  const enabled = overlayTextOnMediaEnabled(kind);
  return {
    enabled,
    lowerThird: enabled && !!State.settings?.overlayTextOnMediaLowerThird,
    dim: enabled ? Math.max(0, Math.min(90, Number(State.settings?.overlayTextOnMediaDim ?? 35))) : 0,
  };
}

function getLiveTextTarget(kind = null) {
  const liveDisplay = document.getElementById('liveDisplay');
  const liveEmpty = document.getElementById('liveEmpty');
  const prefs = getOverlayTextPrefs(kind);
  const item = prefs.enabled ? getCurrentBackgroundMediaItem() : null;
  if (!liveDisplay) return { host: null, overlay: false };
  if (!item) {
    liveDisplay.querySelectorAll('video,audio').forEach(m => { try { m.pause(); } catch(e) {} });
    liveDisplay.innerHTML = '';
    liveDisplay.style.padding = '';
    liveDisplay.style.position = '';
    liveDisplay.style.display = 'flex';
    if (liveEmpty) liveEmpty.style.display = 'none';
    return { host: liveDisplay, overlay: false };
  }
  liveDisplay.style.cssText = 'display:flex;position:relative;overflow:hidden;padding:0;background:#000';
  let bg = liveDisplay.querySelector('.live-media-bg');
  let dim = liveDisplay.querySelector('.live-media-dim');
  let fg = liveDisplay.querySelector('.live-text-overlay-host');
  const canReuseBg = !!(bg && _sameMediaIdentity(bg.__mediaItem, item));
  if (!bg || !dim || !fg) {
    liveDisplay.querySelectorAll('video,audio').forEach(m => { try { m.pause(); } catch(e) {} });
    liveDisplay.innerHTML = '';
    bg = document.createElement('div');
    bg.className = 'live-media-bg';
    bg.style.cssText = 'position:absolute;inset:0;z-index:0;overflow:hidden;background:#000';
    dim = document.createElement('div');
    dim.className = 'live-media-dim';
    fg = document.createElement('div');
    fg.className = 'live-text-overlay-host';
    liveDisplay.appendChild(bg);
    liveDisplay.appendChild(dim);
    liveDisplay.appendChild(fg);
  }
  if (!canReuseBg) renderLiveBackgroundMedia(bg, item);
  else renderLiveBackgroundMedia(bg, item);
  dim.style.cssText = `position:absolute;inset:0;z-index:1;background:rgba(0,0,0,${prefs.dim/100});pointer-events:none`;
  fg.style.cssText = prefs.lowerThird
    ? 'position:absolute;left:0;right:0;bottom:0;z-index:2;display:flex;align-items:flex-end;justify-content:center;padding:0 0 5vh 0;background:transparent;min-height:42%'
    : 'position:absolute;inset:0;z-index:2;display:flex;align-items:center;justify-content:center;padding:0;background:transparent';
  fg.innerHTML = '';
  if (liveEmpty) liveEmpty.style.display = 'none';
  return { host: fg, overlay: true };
}

function getLiveThemeTextHost(display) {
  if (!display) return null;
  let fg = display.querySelector('.live-theme-fg');
  if (!fg) {
    fg = document.createElement('div');
    fg.className = 'live-theme-fg';
    fg.style.cssText = 'position:absolute;inset:0;z-index:2;display:flex;align-items:center;justify-content:center;background:transparent';
    display.appendChild(fg);
  }
  fg.innerHTML = '';
  fg.style.cssText = 'position:absolute;inset:0;z-index:2;display:flex;align-items:center;justify-content:center;background:transparent';
  return fg;
}

async function toggleOverlayTextOnMedia() {
  State.programOverlayTextOnMedia = !overlayTextOnMediaEnabled();
  updateOverlayTextOnMediaUI();
  try {
    if (window.electronAPI) {
      const s = await window.electronAPI.getSettings();
      await window.electronAPI.saveSettings({ ...s, overlayTextOnMedia: State.programOverlayTextOnMedia });
      State.settings = { ...State.settings, overlayTextOnMedia: State.programOverlayTextOnMedia };
    }
  } catch (e) {}
  refreshCanvases();
  toast(State.programOverlayTextOnMedia ? '🖼 Text overlay on media enabled' : '🖼 Text overlay on media disabled');
}
async function toggleOverlayModule(kind) {
  try {
    const s = await window.electronAPI.getSettings();
    const key = kind === 'song' ? 'overlayTextOnMediaSongs' : 'overlayTextOnMediaScripture';
    const next = !(State.settings?.[key] !== false);
    await window.electronAPI.saveSettings({ ...s, [key]: next });
    State.settings = { ...State.settings, [key]: next };
    updateOverlayTextOnMediaUI();
    refreshCanvases();
    toast(`${kind === 'song' ? '🎵 Song' : '✝ Scripture'} overlay ${next ? 'enabled' : 'disabled'}`);
  } catch (e) {}
}

async function toggleOverlayLowerThird() {
  try {
    const s = await window.electronAPI.getSettings();
    const next = !State.settings?.overlayTextOnMediaLowerThird;
    await window.electronAPI.saveSettings({ ...s, overlayTextOnMediaLowerThird: next });
    State.settings = { ...State.settings, overlayTextOnMediaLowerThird: next };
    updateOverlayTextOnMediaUI();
    refreshCanvases();
    toast(next ? '▁ Overlay lower-third enabled' : '▁ Overlay lower-third disabled');
  } catch (e) {}
}


// Populate translation dropdowns with only installed versions
async function refreshTranslationDropdowns() {
  if (!window.electronAPI) return;
  let installed = {};
  try { installed = await window.electronAPI.getInstalledVersions() || {}; }
  catch(e) {}

  // Build available list: KJV always first (built-in seed), then all installed
  const available = [];
  const seen = new Set();

  // Always include KJV
  available.push({ id: 'KJV', label: 'KJV' });
  seen.add('KJV');

  // Add all installed translations in alphabetical order
  for (const id of Object.keys(installed).sort()) {
    if (!seen.has(id)) {
      available.push({ id, label: id });
      seen.add(id);
    }
  }

  ['globalTranslation', 'searchTranslation'].forEach(selId => {
    const sel = document.getElementById(selId);
    if (!sel) return;
    const current = sel.value;
    sel.innerHTML = '';
    available.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t.id;
      opt.textContent = t.label;
      if (t.id === current) opt.selected = true;
      sel.appendChild(opt);
    });
    if (!sel.querySelector(`option[value="${current}"]`) && available.length) {
      sel.value = available[0].id;
      State.currentTranslation = available[0].id;
    }
  });
}

function openSettings() {
  if (window.electronAPI) {
    window.electronAPI.openSettings();
  }
}

// ─── WEB MENU BAR ────────────────────────────────────────────────────────────
let _wmOpen = null;
function _wmCloseAll() { if (_wmOpen) { _wmOpen.classList.remove('open'); _wmOpen = null; } }
document.addEventListener('mousedown', e => {
  if (_wmOpen && !e.target.closest('.web-menu')) _wmCloseAll();
});
document.addEventListener('DOMContentLoaded', () => {
  if (window.electronAPI && !window.electronAPI.isWeb) {
    const wb = document.getElementById('webMenubar');
    if (wb) wb.style.display = 'none';
  }
  document.querySelectorAll('.web-menu-label').forEach(lbl => {
    lbl.addEventListener('click', e => {
      e.stopPropagation();
      const menu = lbl.parentElement;
      if (menu === _wmOpen) { _wmCloseAll(); }
      else { _wmCloseAll(); menu.classList.add('open'); _wmOpen = menu; }
    });
    lbl.addEventListener('mouseenter', () => {
      if (_wmOpen && _wmOpen !== lbl.parentElement) {
        _wmCloseAll();
        lbl.parentElement.classList.add('open');
        _wmOpen = lbl.parentElement;
      }
    });
  });
});

function wmAction(id) {
  _wmCloseAll();
  switch (id) {
    case 'schedule-new':    if (typeof newSchedule === 'function') newSchedule(); break;
    case 'schedule-save':   if (typeof saveSchedule === 'function') saveSchedule(); break;
    case 'schedule-save-as':if (typeof saveScheduleAs === 'function') saveScheduleAs(); break;
    case 'schedule-open':   if (typeof openSchedule === 'function') openSchedule(); break;
    case 'schedule-export': if (typeof exportSchedule === 'function') exportSchedule(); break;
    case 'schedule-import': if (typeof importSchedule === 'function') importSchedule(); break;
    case 'preset-save':   savePresetDialog(); break;
    case 'preset-load':   togglePresetPopup(); break;
    case 'preset-manage': openManagePresets(); break;
    case 'preferences':     openSettings(); break;
    case 'sermon-notes':    showModal('notesOverlay'); break;
    case 'countdown-timer': showModal('timerOverlay'); _highlightTimerPosBtn(); break;
    case 'alerts':          showModal('alertsOverlay'); break;
    case 'theme-designer':
      if (window.electronAPI?.openThemeDesigner) window.electronAPI.openThemeDesigner({ category: 'scripture' });
      else window.open('/theme-designer.html', '_blank');
      break;
    case 'presentation-editor':
      if (window.electronAPI?.openPresentationEditor) window.electronAPI.openPresentationEditor();
      else window.open('/presentation-editor.html', '_blank');
      break;
    case 'song-manager':
      if (window.electronAPI?.openSongManager) window.electronAPI.openSongManager();
      else window.open('/song-manager.html', '_blank');
      break;
    case 'bible-manager':
      if (window.electronAPI?.openBibleManager) window.electronAPI.openBibleManager();
      else window.open('/bible-manager.html', '_blank');
      break;
    case 'sermon-history':
      if (window.electronAPI?.openHistory) window.electronAPI.openHistory();
      else window.open('/help.html', '_blank');
      break;
    case 'remote-url':      openRemotePopover(); break;
    case 'external-output':
      if (window.electronAPI?.send) window.electronAPI.send('open-ndi-panel');
      else toast('External Output requires the desktop app');
      break;
    case 'open-projection': openProjection(); break;
    case 'close-projection':
      if (window.electronAPI?.closeProjection) window.electronAPI.closeProjection();
      break;
    case 'go-live':         toggleGoLive(); break;
    case 'next-item':       if (typeof nextActiveItem === 'function') nextActiveItem(); break;
    case 'prev-item':       if (typeof prevActiveItem === 'function') prevActiveItem(); break;
    case 'clear-display':   if (typeof ccClear === 'function') ccClear(); break;
    case 'fullscreen':
      if (document.fullscreenElement) document.exitFullscreen();
      else document.documentElement.requestFullscreen();
      break;
    case 'show-occ': {
      const cc = document.getElementById('commandCenter');
      if (cc) { cc.style.display = ''; const b = document.getElementById('ccBody'); if (b) b.style.display = 'block'; }
      break;
    }
    case 'help':
      window.open('/help.html', '_blank');
      break;
    case 'shortcuts':
      window.open('/help.html', '_blank');
      break;
    case 'about':
      showAboutModal();
      break;
  }
}

// ─── EVENTS ───────────────────────────────────────────────────────────────────
function bindEvents() {
  // Titlebar
  document.getElementById('modeToggle').addEventListener('click', toggleMode);
  document.getElementById('displayModeBadge').addEventListener('click', () => setDisplayMode(!State.isAutoMode));
  document.getElementById('goLiveBtn').addEventListener('click', toggleGoLive);
  document.getElementById('settingsBtn').addEventListener('click', openSettings);
  document.getElementById('notesBtn').addEventListener('click', () => showModal('notesOverlay'));
  window.electronAPI?.on?.('open-sermon-notes', () => showModal('notesOverlay'));
  document.getElementById('overlaySongToggle')?.addEventListener('click', () => toggleOverlayModule('song'));
  document.getElementById('overlayScriptureToggle')?.addEventListener('click', () => toggleOverlayModule('scripture'));
  document.getElementById('overlayLowerThirdToggle')?.addEventListener('click', () => toggleOverlayLowerThird());
  document.getElementById('projectionBtn').addEventListener('click', openProjection);
  document.getElementById('historyBtn')?.addEventListener('click', () => {
    if (window.electronAPI) window.electronAPI.openHistory();
    else toast('ℹ Sermon History requires the desktop app');
  });
  document.getElementById('themeBtn')?.addEventListener('click', () => {
    if (window.electronAPI) window.electronAPI.openThemeDesigner({ category: State.currentTab === 'theme' ? _themeCategory : 'scripture' });
    else toast('ℹ Theme Designer requires the desktop app');
  });
  document.getElementById('remoteBtn')?.addEventListener('click', openRemotePopover);
  document.getElementById('countdownTbBtn')?.addEventListener('click', () => {
    if (window.electronAPI?.openCountdownWindow) window.electronAPI.openCountdownWindow();
    else toast('ℹ Timer window requires the desktop app');
  });

  // Songs toolbar button — opens song manager editor window
  document.getElementById('songsBtn')?.addEventListener('click', () => {
    if (window.electronAPI) window.electronAPI.openSongManager();
    else toast('ℹ Song Manager requires the desktop app');
  });

  // Bottom panel tab listeners
  document.getElementById('songLibTab')?.addEventListener('click',  () => switchTab('song'));
  document.getElementById('mediaLibTab')?.addEventListener('click', () => switchTab('media'));
  document.getElementById('presLibTab')?.addEventListener('click',  () => switchTab('pres'));
  document.getElementById('themeLibTab')?.addEventListener('click', () => switchTab('theme'));

  // Song search + new/edit buttons
  document.getElementById('songSearchInput')?.addEventListener('input', e => {
    renderSongList(e.target.value);
  });
  document.getElementById('clearSongSearchBtn')?.addEventListener('click', () => {
    const inp = document.getElementById('songSearchInput');
    if (inp) inp.value = '';
    renderSongList('');
  });
  document.getElementById('newSongBtn')?.addEventListener('click', () => {
    if (window.electronAPI) window.electronAPI.openSongManager({ newSong: true });
    else toast('ℹ Song Manager requires the desktop app');
  });
  document.getElementById('editSongBtn')?.addEventListener('click', () => {
    if (!State.currentSongId) { toast('ℹ Select a song first'); return; }
    if (window.electronAPI) window.electronAPI.openSongManager({ songId: State.currentSongId });
    else toast('ℹ Song Manager requires the desktop app');
  });

  // Song list click — single delegated handler, wired once here
  document.getElementById('songListBody')?.addEventListener('click', _songListClickHandler);
  document.getElementById('songListBody')?.addEventListener('dragstart', _songListDragHandler);
  document.getElementById('songListBody')?.addEventListener('contextmenu', _songListContextHandler);
  _makeQueueBodyDropTarget();
  // Song slides click — single delegated handler, wired once here
  document.getElementById('songSlidesList')?.addEventListener('click',   _songSlideClickHandler);
  document.getElementById('songSlidesList')?.addEventListener('dblclick', _songSlideClickHandler);

  // Media tab — add button and search
  document.getElementById('addMediaBtn')?.addEventListener('click', () => {
    document.getElementById('mediaFileInput')?.click();
  });
  document.getElementById('mediaFileInput')?.addEventListener('change', async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length) await addMediaFiles(files);
    e.target.value = '';
  });
  document.getElementById('mediaSearchInput')?.addEventListener('input', e => {
    renderMediaList(e.target.value);
  });
  document.getElementById('mediaIntegrityBtn')?.addEventListener('click', showMediaIntegrityPanel);
  document.getElementById('clearMediaCacheBtn')?.addEventListener('click', clearMediaCacheAction);
  document.getElementById('globalTranslation').addEventListener('change', e => {
    State.currentTranslation = e.target.value;
    document.getElementById('searchTranslation').value = e.target.value;
    refreshCanvases();
  });

  // Theme swatches
  document.querySelectorAll('.theme-swatch').forEach(sw => {
    sw.addEventListener('click', () => setTheme(sw.dataset.t));
  });

  // Transcript
  document.getElementById('recordBtn').addEventListener('click', toggleRecording);
  document.getElementById('clearTranscriptBtn').addEventListener('click', clearTranscript);
  document.getElementById('exportTranscriptBtn').addEventListener('click', exportTranscript);
  document.getElementById('reviewTranscriptBtn')?.addEventListener('click', openTranscriptReviewPanel);
  if (document.getElementById('serviceReplayBtn')) document.getElementById('serviceReplayBtn').addEventListener('click', openServiceReplayTimeline);
  if (document.getElementById('serviceArchiveBtn')) document.getElementById('serviceArchiveBtn').addEventListener('click', openServiceArchiveCenter);
  if (document.getElementById('autoServiceBuilderBtn')) document.getElementById('autoServiceBuilderBtn').addEventListener('click', openAutoServiceBuilder);
  if (document.getElementById('liveSmartSuggestionsBtn')) document.getElementById('liveSmartSuggestionsBtn').addEventListener('click', openLiveSmartSuggestions);
  if (document.getElementById('sermonIntelligenceBtn')) document.getElementById('sermonIntelligenceBtn').addEventListener('click', openSermonIntelligence);
  if (document.getElementById('analyticsDashboardBtn')) document.getElementById('analyticsDashboardBtn').addEventListener('click', openAnalyticsDashboard);
  if (document.getElementById('clipGeneratorBtn')) document.getElementById('clipGeneratorBtn').addEventListener('click', openClipGenerator);
  if (document.getElementById('serviceArchiveSearchInput')) document.getElementById('serviceArchiveSearchInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') runServiceArchiveSearch(); });
  document.getElementById('transcriptReviewSearch')?.addEventListener('input', renderTranscriptReviewList);

  // Audio file upload — transcribe a recorded sermon file
  document.getElementById('audioUploadBtn')?.addEventListener('click', () => {
    document.getElementById('audioFileInput')?.click();
  });
  document.getElementById('audioFileInput')?.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = ''; // reset so same file can be re-selected
    await transcribeAudioFile(file);
  });

  // Manual text input — type sermon text to test AI detection
  const manualInput = document.getElementById('manualInput');
  if (manualInput) {
    manualInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && manualInput.value.trim()) {
        const text = manualInput.value.trim();
        manualInput.value = '';
        // Ensure recording state is active
        if (!State.isRecording) {
          State.isRecording = true;
          State.recordingStartTime = Date.now();
          const btn = document.getElementById('recordBtn');
          btn.textContent = '⏹ Stop Transcript';
          btn.classList.add('active');
          startMicAnimation();
          const body = document.getElementById('transcriptBody');
          const empty = body.querySelector('.empty-state');
          if (empty) empty.remove();
        }
        pushTranscriptLine(text, true);
      }
    });
    // Show manual input when transcript starts
    document.getElementById('recordBtn')?.addEventListener('click', () => {
      setTimeout(() => {
        if (manualInput) manualInput.parentElement.style.display =
          State.isRecording ? 'block' : 'none';
      }, 100);
    });
  }

  // Preview panel
  document.getElementById('presentToLiveBtn').addEventListener('click', sendPreviewToLive);
  document.getElementById('openPreviewFullBtn').addEventListener('click', openFullscreenPreview);


  // Live panel
  document.getElementById('clearLiveBtn').addEventListener('click', blankScreen);
  document.getElementById('openLiveFullBtn').addEventListener('click', openProjection);
  _updateLogoBtnState();

  // Queue
  document.getElementById('clearQueueBtn').addEventListener('click', clearQueue);
  document.getElementById('saveScheduleBtn')?.addEventListener('click', saveSchedule);
  document.getElementById('loadScheduleBtn')?.addEventListener('click', openLoadScheduleModal);
  document.getElementById('addSectionBtn')?.addEventListener('click', addSectionHeader);

  // Presentations tab
  document.getElementById('presFileInput')?.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (file) await importPresentation(file);
    e.target.value = '';
  });
  document.getElementById('presPptxInput')?.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (file) await importPresentation(file);
    e.target.value = '';
  });
  document.getElementById('presSearchInput')?.addEventListener('input', e => {
    renderPresLibStrip(e.target.value);
  });
  const queueCurrentPresentation = () => {
    if (!State.currentPresId || !State.currentPresSlides?.length) { toast('⚠ Select a presentation first'); return; }
    addPresentationToQueue(State.currentPresId);
  };
  document.getElementById('presImportBtn')?.addEventListener('click', () => importPresentation(false));
  document.getElementById('presPreviewImportBtn')?.addEventListener('click', () => importPresentation(false));
  document.getElementById('presAddScheduleBtn')?.addEventListener('click', queueCurrentPresentation);
  document.getElementById('presPreviewScheduleBtn')?.addEventListener('click', queueCurrentPresentation);

  // Search
  document.getElementById('bookTab').addEventListener('click', () => switchTab('book'));
  // Context search tab removed — merged into Bible Search
  document.getElementById('searchInput').addEventListener('input', e => handleSearchInput(e.target.value));
  document.getElementById('searchInput').addEventListener('keydown', e => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    e.stopPropagation();
    const q = (e.currentTarget.value || '').trim();
    if (!q) return;
    smartSearch(q, { sendLive: true });
  });
  document.getElementById('clearBibleSearchBtn')?.addEventListener('click', () => {
    const input = document.getElementById('searchInput');
    if (input) input.value = '';
    State.searchQuery = '';
    // Keep the currently displayed chapter/verse intact; only clear the typed query.
    try { renderSearchResults(); } catch (_) {}
    setTimeout(() => { try { highlightVerseRow(State.previewVerse?.verse || 1); } catch (_) {} }, 50);
    input?.focus();
  });
  document.getElementById('searchTranslation').addEventListener('change', e => {
    State.currentTranslation = e.target.value;
    document.getElementById('globalTranslation').value = e.target.value;
    refreshCanvases();
    renderSearchResults();
    setTimeout(() => { try { highlightVerseRow(State.previewVerse?.verse || 1); } catch (_) {} }, 50);
  });
  document.getElementById('prevChapterBtn').addEventListener('click', prevChapter);
  document.getElementById('nextChapterBtn').addEventListener('click', nextChapter);

  // Detections
  document.getElementById('clearDetectionsBtn').addEventListener('click', clearDetections);
  document.getElementById('reviewDetectionsBtn')?.addEventListener('click', openDetectionReviewPanel);
  document.getElementById('detectionReviewSearch')?.addEventListener('input', renderDetectionReviewList);

  // Notes modal
  document.getElementById('generateNotesBtn').addEventListener('click', generateSermonNotes);
  document.getElementById('notesCancelBtn').addEventListener('click', () => closeModal('notesOverlay'));
  document.getElementById('notesClose').addEventListener('click', () => closeModal('notesOverlay'));

  // Load saved transcripts when notes modal opens
  document.getElementById('notesBtn')?.addEventListener('click', () => _loadSavedTranscripts());
  window.electronAPI?.on?.('open-sermon-notes', () => _loadSavedTranscripts());

  // Use current live transcript button
  document.getElementById('useCurrentTranscriptBtn')?.addEventListener('click', () => {
    State._selectedSavedTranscriptId = null;
    // Highlight active
    document.querySelectorAll('.saved-transcript-item').forEach(el => el.classList.remove('active'));
    document.getElementById('useCurrentTranscriptBtn').style.background = 'rgba(201,168,76,.25)';
    const info = document.getElementById('transcriptPreviewInfo');
    info.style.display = 'block';
    info.textContent = `Using live transcript — ${State.transcriptLines.length} lines`;
  });
  document.getElementById('exportNotesBtn').addEventListener('click', exportNotes);
  document.getElementById('copyNotesBtn').addEventListener('click', copyNotesToClipboard);
  document.getElementById('generatePresBtn')?.addEventListener('click', generatePresFromNotes);
  document.getElementById('importNotesBtn')?.addEventListener('click', importNotesFromFile);

  // NDI panel
  document.getElementById('ndiClose')?.addEventListener('click', () => closeModal('ndiOverlay'));
  document.getElementById('ndiCloseFooter')?.addEventListener('click', () => closeModal('ndiOverlay'));
  document.getElementById('ndiStartBtn')?.addEventListener('click', async () => {
    if (!window.electronAPI) return;
    document.getElementById('ndiStatusLabel').textContent = '⏳ Starting NDI…';
    const result = await window.electronAPI.ndiStart();
    updateNdiPanel(result);
  });
  document.getElementById('ndiStopBtn')?.addEventListener('click', async () => {
    if (!window.electronAPI) return;
    await window.electronAPI.ndiStop();
    updateNdiPanel({ status: 'disabled', method: 'none' });
  });
  // NDI copy buttons
  document.getElementById('ndiCopyObs')?.addEventListener('click', () => {
    const u = document.getElementById('ndiObsUrl')?.textContent;
    if(u){ navigator.clipboard?.writeText(u); toast('📋 OBS URL copied!'); }
  });
  document.getElementById('ndiCopyVmix')?.addEventListener('click', () => {
    const u = document.getElementById('ndiVmixUrl')?.textContent;
    if(u){ navigator.clipboard?.writeText(u); toast('📋 vMix URL copied!'); }
  });

  // Projection overlay
  document.getElementById('projClose').addEventListener('click', closeFullscreenPreview);
  document.getElementById('projOverlay').addEventListener('click', e => {
    if (e.target === document.getElementById('projOverlay') || e.target === document.getElementById('projClose'))
      closeFullscreenPreview();
  });

  // Keyboard shortcuts
  document.addEventListener('keydown', handleKeyboard);

  // Modal backdrop click
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', e => {
      if (e.target === overlay) closeModal(overlay.id);
    });
  });

  // Wire IPC menu + editor events
  _wireMenuEvents();
  _wirePresEditorSaved();
}

// ─── Whisper Setup Banner ────────────────────────────────────────────────────
function _showWhisperSetupBanner(data) {
  // Don't show duplicate banners
  if (document.getElementById('whisperSetupNotice')) return;

  const reason = data?.reason || 'no_python';
  const hasBat = data?.setupBatExists !== false;

  // Build human-readable title + detail based on exact failure reason
  let icon  = '🎙️';
  let title = 'Local Transcription not installed';
  let detail = 'Configure a transcription key in Settings (Deepgram or OpenAI), or install Local Whisper — free, offline, no API key needed.';
  let btnLabel = 'Set Up Now';

  if (reason === 'no_python') {
    icon     = '🎙️';
    title    = 'Local Whisper not installed';
    detail   = 'Python and the Whisper AI engine are not installed. Click Set Up Now for a one-time setup — installs portable Python 3.12 + Whisper automatically (~200 MB). No system changes are made.';
    btnLabel = 'Set Up Now';
  } else if (reason === 'no_faster_whisper') {
    icon     = '⚠️';
    title    = 'Whisper engine missing';
    detail   = 'Python is installed but the faster-whisper package is missing. Click Set Up Now to complete the installation.';
    btnLabel = 'Set Up Now';
  } else if (reason.startsWith('wrong_version:')) {
    const ver = reason.split(':')[1] || 'unknown';
    icon     = '⚠️';
    title    = `Python ${ver} is not supported`;
    detail   = `Local Transcription requires Python 3.8 or newer. Your system has Python ${ver}. Click Set Up Now to install portable Python 3.12 automatically.`;
    btnLabel = 'Set Up Now';
  } else if (reason === 'model_not_found') {
    icon     = '📥';
    title    = 'Whisper model not downloaded';
    detail   = 'Python and faster-whisper are ready, but the speech model file is missing. Click Download Model to get it (~244 MB).';
    btnLabel = 'Download Model';
  } else if (reason === 'missing_vcredist') {
    icon   = '⚠️';
    title  = 'Visual C++ Runtime missing';
    detail = 'The AI transcription engine requires the Microsoft Visual C++ 2015-2022 Redistributable. ' +
             'Please reinstall AnchorCast — it will install VC++ automatically. ' +
             'Or download it manually from: aka.ms/vs/17/release/vc_redist.x64.exe';
    btnLabel = 'Download VC++';
  } else if (reason === 'no_source_available') {
    icon   = '⚠️';
    title  = 'No transcription source configured';
    detail = 'Configure a transcription key in Settings (Deepgram or OpenAI), or set up Local Whisper below — free, works offline, no API key needed.';
  }

  const n = document.createElement('div');
  n.id = 'whisperSetupNotice';
  n.style.cssText = [
    'position:fixed', 'bottom:52px', 'left:50%', 'transform:translateX(-50%)',
    'z-index:8000', 'background:#1e293b', 'border:1px solid #f59e0b',
    'border-radius:10px', 'padding:14px 18px', 'display:flex',
    'align-items:center', 'gap:14px', 'box-shadow:0 8px 32px rgba(0,0,0,.6)',
    'max-width:520px', 'width:calc(100vw - 40px)',
    'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
  ].join(';');

  n.innerHTML = `
    <span style="font-size:22px;flex-shrink:0">${icon}</span>
    <div style="flex:1;min-width:0">
      <div style="font-size:13px;font-weight:700;color:#f59e0b;margin-bottom:3px">${title}</div>
      <div style="font-size:12px;color:#94a3b8;line-height:1.4">${detail}</div>
    </div>
    ${hasBat ? `<button id="whisperSetupBtn" style="
      background:#f59e0b;color:#000;border:none;border-radius:7px;
      padding:8px 15px;font-size:12px;font-weight:700;cursor:pointer;
      white-space:nowrap;flex-shrink:0">${btnLabel}</button>` : ''}
    <button id="whisperSetupDismiss" style="
      background:transparent;border:none;color:#475569;
      cursor:pointer;font-size:20px;padding:0 2px;flex-shrink:0;line-height:1">✕</button>
  `;
  document.body.appendChild(n);

  if (hasBat) {
    document.getElementById('whisperSetupBtn').onclick = async () => {
      // Special case: VC++ missing — open download page
      if (reason === 'missing_vcredist') {
        window.electronAPI?.openExternal?.('https://aka.ms/vs/17/release/vc_redist.x64.exe');
        return;
      }
      const btn = document.getElementById('whisperSetupBtn');
      btn.textContent = 'Setting up…';
      btn.disabled = true;
      // Update banner to show progress state
      n.querySelector('div[style*="flex:1"]').innerHTML = `
        <div style="font-size:13px;font-weight:700;color:#f59e0b;margin-bottom:3px">Setting up Whisper AI…</div>
        <div style="font-size:12px;color:#94a3b8;line-height:1.4">A console window is running the setup. This may take 2-5 minutes.</div>
      `;
      const r = await window.electronAPI?.runWhisperSetup?.();
      if (!r?.ok) {
        toast('⚠️ Could not start setup: ' + (r?.error || 'unknown error'));
        btn.textContent = btnLabel;
        btn.disabled = false;
      }
      // Result comes via whisper-setup-result event (see setupElectronEvents)
    };
  }
  document.getElementById('whisperSetupDismiss').onclick = () => n.remove();
}

function setupElectronEvents() {
  if (!window.electronAPI) return;
  window.electronAPI.on('projection-opened', () => { State.isProjectionOpen = true; });
  // Sync timer state from countdown window to quick timer bar in main app
  window.electronAPI.on('timer-state-sync', (state) => {
    if (!state) return;
    // running may be explicit true, or implied by having seconds > 0
    const isRunning = state.running === true || (state.seconds > 0 && state.running !== false);
    if (isRunning) {
      _startLiveTimerDisplay(state.remaining || state.seconds, state.mode, state.label || '');
    }
  });
  window.electronAPI.on('timer-stopped', () => {
    _stopLiveTimerDisplay();
  });

  window.electronAPI.on('display-warning', ({ msg, level }) => {
    // Show a prominent banner when display is disconnected
    const existing = document.getElementById('_displayWarnBanner');
    if (existing) existing.remove();
    const banner = document.createElement('div');
    banner.id = '_displayWarnBanner';
    banner.style.cssText = [
      'position:fixed;top:0;left:0;right:0;z-index:99999',
      'background:' + (level === 'error' ? '#7f1d1d' : '#78350f'),
      'color:#fff;font-size:13px;font-weight:600',
      'padding:10px 16px;display:flex;align-items:center;gap:12px',
      'box-shadow:0 2px 12px rgba(0,0,0,.5)',
      'border-bottom:2px solid ' + (level === 'error' ? '#ef4444' : '#f59e0b'),
    ].join(';');
    banner.innerHTML = (level === 'error' ? '🔌 ' : '⚠ ') + msg +
      '<button onclick="this.parentElement.remove()" style="margin-left:auto;background:rgba(255,255,255,.15);' +
      'border:none;color:#fff;padding:3px 10px;border-radius:4px;cursor:pointer;font-size:12px">✕</button>';
    document.body.appendChild(banner);
    // Auto-dismiss after 12 seconds
    setTimeout(() => { try { banner.remove(); } catch(_) {} }, 12000);
    // Also show toast
    toast(level === 'error' ? '🔌 ' + msg : '⚠ ' + msg);
  });

  window.electronAPI.on('projection-closed', () => { State.isProjectionOpen = false; });
  window.electronAPI.on('shortcut-go-live', toggleGoLive);
  // BUG-C FIX: nextActiveItem/prevActiveItem guarded — they are defined later in the file
  // so the IPC listener closure captures them safely at call time (not definition time)
  window.electronAPI.on('shortcut-next', () => { if (typeof nextActiveItem === 'function') nextActiveItem(); });
  window.electronAPI.on('shortcut-prev', () => { if (typeof prevActiveItem === 'function') prevActiveItem(); });
  window.electronAPI.on('shortcut-clear', blankScreen);
  window.electronAPI.on('open-settings-modal', openSettings);
  // BUG-16 FIX: register transcript-no-key ONCE here instead of inside
  // startPcmCapture() — which ran on every recording start and stacked
  // duplicate listeners (after 3 restarts the toast fired 3 times).
  window.electronAPI.on('transcript-no-key', () => {
    toast('⚠ Invalid API key, missing API key, or billing/fund issue. Check Settings and your provider account.');
  });
  // Reload settings when settings window saves
  window.electronAPI.on('settings-saved', async (payload) => {
    const themeOnly = !!(payload && payload.themeOnly);
    const changedKeys = (payload && Array.isArray(payload.changedKeys)) ? payload.changedKeys : [];
    const onlyThemeChanged = themeOnly;  // explicit flag is reliable; key-diff heuristic was fragile

    const prevMicId = State.settings?.microphoneId;
    State.settings = await window.electronAPI.getSettings();
    applySettings(State.settings);
    refreshTranslationDropdowns();
    if (State.settings.apiKey && window.AIDetection) {
      AIDetection.init(State.settings.apiKey, handleDetection);
    }
    updateSrcToggleUI();

    const newMicId = State.settings?.microphoneId;
    if (State.isRecording && newMicId !== prevMicId) {
      toast('🔄 Microphone changed — restarting capture…');
      if (pcmWorkletNode)  { try { pcmWorkletNode.disconnect(); }  catch(e){} pcmWorkletNode = null; }
      if (pcmSourceNode)   { try { pcmSourceNode.disconnect(); }   catch(e){} pcmSourceNode  = null; }
      if (pcmStream)       { pcmStream.getTracks().forEach(t => t.stop()); pcmStream = null; }
      if (pcmAudioContext) { try { pcmAudioContext.close(); }       catch(e){} pcmAudioContext = null; }
      stopDeepgramSocket();
      startPcmCapture();
    }

    // If only song/presentation theme keys changed, re-render the song/pres output
    // and stop — do NOT re-project scripture even if liveContentType === 'scripture'.
    if (onlyThemeChanged) {
      if (State.isLive && State.liveContentType === 'song' && State.currentSongId != null && State.currentSongSlideIdx != null) {
        const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
        if (song?.sections?.[State.currentSongSlideIdx]) {
          const sec = song.sections[State.currentSongSlideIdx];
          const lines = (sec.lines || []).filter(l => l.trim());
          renderSongLiveDisplay(song, sec, lines, getActiveSongThemeData(), 'present');
          await sendSongToProjection(song, State.currentSongSlideIdx);
        }
      }
      // Song/pres theme change: no toast, no scripture re-projection
      return;
    }

    // General settings change — re-render the active live module
    if (State.liveContentType === 'song' && State.currentSongId != null && State.currentSongSlideIdx != null) {
      State.liveVerse = null;
      const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
      if (song?.sections?.[State.currentSongSlideIdx]) {
        const sec = song.sections[State.currentSongSlideIdx];
        const lines = (sec.lines || []).filter(l => l.trim());
        const themeData = getActiveSongThemeData();
        const songMode = State.isLive ? 'present' : 'preview';
        renderSongLiveDisplay(song, sec, lines, themeData, songMode);
        if (State.isLive) await sendSongToProjection(song, State.currentSongSlideIdx);
      }
    } else if (State.isLive && State.liveContentType === 'scripture' && State.liveVerse) {
      const { book, chapter, verse, ref } = State.liveVerse;
      updateLiveDisplay(book, chapter, verse, ref);
      syncProjection();
    } else if (State.isLive && State.liveContentType === 'media' && State.currentMediaId) {
      const item = (State.media || []).find(m => m.id === State.currentMediaId);
      if (item) {
        previewMedia(item);
        presentMedia(item);
      }
    } else {
      refreshCanvases();
    }
    toast('✓ Settings applied');
  });
  window.electronAPI.on('bible-versions-updated', async () => {
    try { await loadBibleData(); } catch(_) {}
    await refreshTranslationDropdowns();
    if (!BibleDB?.translations?.[State.currentTranslation]) State.currentTranslation = 'KJV';
    const gt = document.getElementById('globalTranslation');
    const st = document.getElementById('searchTranslation');
    if (gt) gt.value = State.currentTranslation;
    if (st) st.value = State.currentTranslation;
    try { renderBibleSearch(); } catch (_) {}
    try { refreshCanvases(); } catch (_) {}
    toast('✓ Bible versions refreshed');
  });
  window.electronAPI.on('ndi-status', (info) => {
    updateNdiPanel(info);
  });
  window.electronAPI.on('open-ndi-panel', () => {
    openNdiPanel();
  });
  // Reload song library when Song Manager saves
  window.electronAPI.on('songs-saved', () => {
    const prevSongId = State.currentSongId;
    const prevSlideIdx = State.currentSongSlideIdx;
    const wasLiveSong = !!(State.isLive && State.liveContentType === 'song' && prevSongId != null);
    setTimeout(async () => {
      try {
        await reloadSongs();
        if (prevSongId != null) {
          const refreshedSong = (State.songs || []).find(s => String(s.id) === String(prevSongId));
          if (refreshedSong) {
            State.currentSongId = refreshedSong.id;
            if (prevSlideIdx != null) {
              State.currentSongSlideIdx = Math.max(0, Math.min(prevSlideIdx, (refreshedSong.sections?.length || 1) - 1));
            }
            try { selectSong(refreshedSong.id); } catch (_) {}
          }
        }
        if (wasLiveSong) {
          const liveSong = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
          const liveIdx = Math.max(0, Math.min(State.currentSongSlideIdx ?? 0, ((liveSong?.sections || []).length || 1) - 1));
          if (liveSong?.sections?.[liveIdx]) {
            const sec = liveSong.sections[liveIdx];
            const lines = (sec.lines || []).filter(l => String(l || '').trim());
            const themeData = getActiveSongThemeData();
            try { renderSongLiveDisplay(liveSong, sec, lines, themeData, 'present'); } catch (_) {}
            try { await sendSongToProjection(liveSong, liveIdx); } catch (_) {}
          }
        }
        if (State.currentTab === 'song') toast('🎵 Song library updated');
      } catch(_) {}
    }, 300);
  });

  // Reload themes when Theme Designer saves — re-apply active themes immediately
  window.electronAPI.on('themes-updated', async () => {
    try {
      const fresh = await window.electronAPI.getThemes();
      if (fresh) State.themes = fresh;
    } catch(e) {}
    // Refresh theme swatches and grid
    refreshThemeSwatches();
    if (State.currentTab === 'theme') {
      loadThemeGrid();
    }
    // Only re-render if a scripture or presentation theme changed — not song themes.
    // Song theme updates don't affect live scripture projection, so we avoid an
    // unintended re-projection of the current verse when a song theme is saved.
    const liveType = State.liveContentType || '';
    if (liveType === 'scripture' || liveType === 'presentation') {
      refreshCanvases();
    } else if (liveType === 'song' && State.currentSongId != null) {
      // Re-render song live display with updated song theme
      const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
      if (song && State.currentSongSlideIdx != null) {
        const sec   = song.sections[State.currentSongSlideIdx];
        const lines = (sec?.lines || []).filter(l => l.trim());
        renderSongLiveDisplay(song, sec, lines, getActiveSongThemeData(), 'present');
        if (State.isLive) sendSongToProjection(song, State.currentSongSlideIdx);
      }
    }
    toast('🎨 Themes updated');
  });

  // Whisper local server status
  window.electronAPI.on('whisper-status', (status) => {
    State.whisperLocalReady = !!status?.ready;
    updateSrcToggleUI();
  });

  // Detailed whisper setup notification — fires when main process cannot find
  // a working Python + faster-whisper installation
  window.electronAPI.on('whisper-setup-needed', (data) => {
    _showWhisperSetupBanner(data);
  });

  // Result from setup_whisper.bat — fires when main process detects flag file
  // Close safeguard — fires when user closes app with an unsaved/unprocessed transcript
  window.electronAPI.on('confirm-quit-with-transcript', () => {
    const lines = State.transcriptLines?.length || 0;
    if (!lines) { window.electronAPI.quitConfirmed(); return; }

    // Build dialog
    const d = document.createElement('div');
    d.id = 'quitGuardDialog';
    d.style.cssText = [
      'position:fixed','inset:0','z-index:10000',
      'background:rgba(0,0,0,.72)','display:flex',
      'align-items:center','justify-content:center',
      'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
    ].join(';');
    d.innerHTML = `
      <div style="background:#1e293b;border:1px solid #334155;border-radius:14px;
        padding:28px 32px;max-width:420px;width:90vw;box-shadow:0 24px 64px rgba(0,0,0,.6)">
        <div style="font-size:18px;font-weight:700;color:#f1f5f9;margin-bottom:8px">
          Close AnchorCast?
        </div>
        <div style="font-size:13px;color:#94a3b8;line-height:1.6;margin-bottom:24px">
          You have a transcript with <strong style="color:#e2e8f0">${lines} lines</strong>.
          It has been auto-saved to your Transcript History.<br><br>
          Would you like to generate Sermon Notes before closing?
        </div>
        <div style="display:flex;flex-direction:column;gap:8px">
          <button id="qgOpenNotes" style="
            background:linear-gradient(135deg,rgba(201,168,76,.3),rgba(201,168,76,.15));
            border:1px solid rgba(201,168,76,.5);color:#c9a84c;border-radius:8px;
            padding:10px 16px;font-size:13px;font-weight:700;cursor:pointer;text-align:left">
            ✦ Generate Sermon Notes first
          </button>
          <button id="qgQuit" style="
            background:rgba(239,68,68,.15);border:1px solid rgba(239,68,68,.3);
            color:#f87171;border-radius:8px;padding:10px 16px;
            font-size:13px;font-weight:600;cursor:pointer;text-align:left">
            ✕ Close without generating notes
          </button>
          <button id="qgCancel" style="
            background:transparent;border:1px solid #334155;color:#64748b;
            border-radius:8px;padding:10px 16px;
            font-size:13px;cursor:pointer;text-align:left">
            ← Stay in AnchorCast
          </button>
        </div>
      </div>
    `;
    document.body.appendChild(d);

    document.getElementById('qgOpenNotes').onclick = () => {
      d.remove();
      window.electronAPI.quitCancelled();
      showModal('notesOverlay');
      _loadSavedTranscripts(); // refresh list
    };
    document.getElementById('qgQuit').onclick = () => {
      d.remove();
      window.electronAPI.quitConfirmed();
    };
    document.getElementById('qgCancel').onclick = () => {
      d.remove();
      window.electronAPI.quitCancelled();
    };
  });

  window.electronAPI.on('whisper-setup-result', (data) => {
    // Remove any existing setup banner
    const existing = document.getElementById('whisperSetupNotice');
    if (existing) existing.remove();

    if (data?.needsWindowsRestart) {
      // VC++ just installed — need Windows restart
      const n = document.createElement('div');
      n.style.cssText = [
        'position:fixed','bottom:52px','left:50%','transform:translateX(-50%)',
        'z-index:8000','background:#1e293b','border:1px solid #ef4444',
        'border-radius:10px','padding:16px 20px','display:flex',
        'align-items:center','gap:14px','box-shadow:0 8px 32px rgba(0,0,0,.6)',
        'max-width:520px','width:calc(100vw - 40px)',
        'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
      ].join(';');
      n.innerHTML = `
        <span style="font-size:22px;flex-shrink:0">🔄</span>
        <div style="flex:1">
          <div style="font-size:13px;font-weight:700;color:#ef4444;margin-bottom:3px">Windows restart required</div>
          <div style="font-size:12px;color:#94a3b8;line-height:1.4">
            Visual C++ was installed. Restart Windows, then reopen AnchorCast — setup will complete automatically.
          </div>
        </div>
        <button onclick="this.parentElement.remove()" style="
          background:transparent;border:none;color:#475569;
          cursor:pointer;font-size:20px;padding:0 2px;flex-shrink:0">✕</button>
      `;
      document.body.appendChild(n);
    } else if (data?.success) {
      // Success — app is about to restart, show brief confirmation
      toast('✅ ' + (data.message || 'Whisper ready! Restarting…'));
    }
  });
  window.electronAPI.whisperStatus?.().then(s => {
    State.whisperLocalReady = !!s?.ready;
    State.whisperSource = 'local';
    window.electronAPI?.setWhisperSource?.('local');
    updateSrcToggleUI();
    // whisper-setup-needed event from main.js handles the banner when Python is missing
  });
  // Remote control from phone/tablet
  window.electronAPI.on('remote-present', (data) => {
    if (!data) return;
    const ref = data.book
      ? { book: data.book, chapter: data.chapter, verse: data.verse }
      : BibleDB.parseReference(data.rawRef || data.ref);
    if (ref) {
      _setLiveOn();
      presentVerse(ref.book, ref.chapter, ref.verse, data.ref || `${ref.book} ${ref.chapter}:${ref.verse}`);
      updateLiveDisplay(ref.book, ref.chapter, ref.verse, data.ref || `${ref.book} ${ref.chapter}:${ref.verse}`);
      syncProjection();
      if (typeof applyScriptureCapitalizationForActiveOutput === 'function') applyScriptureCapitalizationForActiveOutput();
      toast(`📱 Remote: ${data.ref || ref.book + ' ' + ref.chapter + ':' + ref.verse}`);
    }
  });
  window.electronAPI.on('remote-queue-add', (data) => {
    if (data?.book) addToQueue(data.book, data.chapter, data.verse, data.ref);
  });
  window.electronAPI.on('remote-set-translation', (translation) => {
    if (!translation) return;
    State.currentTranslation = translation;
    document.getElementById('globalTranslation').value = translation;
    document.getElementById('searchTranslation').value = translation;
    if (typeof applyScriptureCapitalizationForActiveOutput === 'function') applyScriptureCapitalizationForActiveOutput();
    refreshCanvases();
    toast(`📱 Translation → ${translation}`);
  });
  window.electronAPI.on('remote-present-song', (data) => {
    const song = (State.songs || []).find(s => String(s.id) === String(data?.songId));
    if (!song) return;
    _setLiveOn();
    if (State.currentTab !== 'song') switchTab('song');
    selectSong(song.id);
    const slideIdx = Math.max(0, Math.min(parseInt(data?.slideIdx ?? 0, 10) || 0, (song.sections?.length || 1) - 1));
    presentSongSlide(slideIdx);
    toast(`📱 Remote Song: ${song.title} — ${(song.sections?.[slideIdx]?.label) || ('Slide ' + (slideIdx + 1))}`);
  });
  window.electronAPI.on('remote-present-media', (data) => {
    const item = (State.media || []).find(m => String(m.id) === String(data?.mediaId));
    if (!item) return;
    _setLiveOn();
    if (State.currentTab !== 'media') switchTab('media');
    previewMedia(item);
    presentMedia(item);
    toast(`📱 Remote Media: ${item.title || item.name || 'Media'}`);
  });
  window.electronAPI.on('remote-control', async (data) => {
    if (!data || !data.action) return;
    const { action, data: payload } = data;
    switch (action) {
      case 'present':
      case 'present-verse': {
        const ref = BibleDB.parseReference(payload?.rawRef || payload?.ref);
        if (ref) {
          _setLiveOn();
          if (State.currentTab !== 'book') switchTab('book');
          const refStr = payload?.ref || `${ref.book} ${ref.chapter}:${ref.verse}`;
          presentVerse(ref.book, ref.chapter, ref.verse, refStr);
          if (typeof applyScriptureCapitalizationForActiveOutput === 'function') applyScriptureCapitalizationForActiveOutput();
          toast(`📱 Remote: ${refStr}`);
        }
        break;
      }
      case 'present-song': {
        if (!State.songs?.length) { try { await reloadSongs(); } catch(_){} }
        const song = (State.songs || []).find(s => String(s.id) === String(payload?.songId));
        if (!song) return;
        _setLiveOn();
        if (State.currentTab !== 'song') switchTab('song');
        selectSong(song.id);
        const slideIdx = Math.max(0, Math.min(parseInt(payload?.slideIdx ?? 0, 10) || 0, (song.sections?.length || 1) - 1));
        await presentSongSlide(slideIdx);
        await _syncCurrentProjection();
        toast(`📱 Remote Song: ${song.title} — ${(song.sections?.[slideIdx]?.label) || ('Slide ' + (slideIdx + 1))}`);
        break;
      }
      case 'present-media': {
        if (!State.media?.length) { try { await loadMedia(); } catch(_){} }
        const item = (State.media || []).find(m => String(m.id) === String(payload?.mediaId));
        if (!item) return;
        _setLiveOn();
        if (State.currentTab !== 'media') switchTab('media');
        previewMedia(item);
        await presentMedia(item);
        await _syncCurrentProjection();
        toast(`📱 Remote Media: ${item.title || item.name || 'Media'}`);
        break;
      }
      case 'present-current': {
        if (!State.isLive) _setLiveOn();
        sendPreviewToLive();
        setTimeout(() => _syncCurrentProjection(), 100);
        break;
      }
      case 'clear': {
        clearLive();
        toast('📱 Remote: Cleared');
        break;
      }
      case 'next': {
        _remoteStep(1, payload?.mode);
        setTimeout(() => _syncCurrentProjection(), 100);
        break;
      }
      case 'prev': {
        _remoteStep(-1, payload?.mode);
        setTimeout(() => _syncCurrentProjection(), 100);
        break;
      }
      case 'presentation-next': {
        _remoteStep(1, 'presentation');
        setTimeout(() => _syncCurrentProjection(), 100);
        break;
      }
      case 'presentation-prev': {
        _remoteStep(-1, 'presentation');
        setTimeout(() => _syncCurrentProjection(), 100);
        break;
      }
    }
    setTimeout(_syncRemoteLiveState, 300);
  });
  window.electronAPI.on('http-server-started', (info) => {
    State.remoteInfo = info;
    updateRemoteBtn(info);
  });
}


function syncBibleListHighlightAndScroll() {
  try {
    const body = document.getElementById('searchResults');
    if (!body) return;
    // verse-row is the actual class used in renderBibleChapter
    const items = Array.from(body.querySelectorAll('.verse-row[data-verse]'));
    if (!items.length) return;

    // Find active verse by matching State.previewVerse
    let activeVerse = null;
    if (State.previewVerse?.book === State.currentBook &&
        State.previewVerse?.chapter === State.currentChapter) {
      activeVerse = State.previewVerse.verse;
    }

    let active = null;
    items.forEach(el => {
      const isActive = activeVerse != null && parseInt(el.dataset.verse) === activeVerse;
      el.classList.toggle('active', isActive);
      el.classList.toggle('selected', isActive);
      if (isActive) active = el;
    });

    if (!active) return;
    // Scroll within the container only — NOT scrollIntoView which scrolls the page
    scrollElementIntoContainer(active, body, 'center');
  } catch (_) {}
}


function handleKeyboard(e) {
  const tag = document.activeElement?.tagName;
  const isInput = tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA';

  if (e.key === 'Escape') {
    closeFullscreenPreview();
    closeModal('notesOverlay');
    closeModal('ndiOverlay');
  }

  // Ctrl+L → Go Live toggle
  if ((e.ctrlKey || e.metaKey) && (e.key === 'l' || e.key === 'L')) {
    e.preventDefault();
    toggleGoLive();
    return;
  }

  // Ctrl+B → Bible Manager
  if ((e.ctrlKey || e.metaKey) && (e.key === 'b' || e.key === 'B') && !e.shiftKey) {
    e.preventDefault();
    wmAction('bible-manager');
    return;
  }

  // Don't fire shortcuts when user is typing
  if (isInput) return;

  if (e.key === 'Tab') { e.preventDefault(); } // Tab no longer cycles context

  if (e.key === 'Enter' || e.key === ' ') {
    // Bible list view: only Enter should send live. Space is reserved for typing/search flow.
    if (State.currentTab === 'book' && State.previewViewMode === 'list') {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (!State.previewVerse) ensureBibleSelection();
        sendPreviewToLive();
      }
      return;
    }
    e.preventDefault();
    if (State.currentTab === 'book' && !State.previewVerse) ensureBibleSelection();
    sendPreviewToLive();
    return;
  }

  // Arrow navigation — Down/Right = Next, Up/Left = Prev
  // ALL arrow keys route through _stepPreviewSlide() which reads the Program
  // Preview slide cards — never the Schedule queue.
  if (e.key === 'ArrowDown' || e.key === 'ArrowRight' ||
      e.key === 'ArrowUp'   || e.key === 'ArrowLeft') {
    e.preventDefault();
    const dir = (e.key === 'ArrowDown' || e.key === 'ArrowRight') ? 1 : -1;
    _stepPreviewSlide(dir);
    return;
  }
}

// ── Song slide navigation — updates highlight on list rows AND slide cards ────
function navigateSongSlide(direction) {
  const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
  if (!song || !song.sections?.length) return;

  const total = song.sections.length;
  // PREV/NEXT FIX: when no slide has been selected yet (null), treat the
  // current position as 0 (first slide), not -1. The old ?? -1 default meant
  // pressing NEXT jumped to slide 1 instead of previewing slide 0 first,
  // and PREV from null clamped to 0 silently doing nothing visible.
  const current = State.currentSongSlideIdx ?? 0;
  let   next    = current + direction;
  if (next < 0)      next = 0;
  if (next >= total) next = total - 1;
  if (next === current && State.currentSongSlideIdx != null) return;

  // Trigger preview/live output
  if (State.isLive) presentSongSlide(next);
  else              previewSongSlide(next);

  // Keep state in sync so the next arrow press anchors from the right position
  State.currentSongSlideIdx = next;

  // Update card highlight + scroll in slide mode
  if (State.previewViewMode === 'slides') {
    const allCards = document.querySelectorAll('#songSlidesList .slide-card');
    allCards.forEach((c, i) => c.classList.toggle('active', i === next));
    const activeCard = document.querySelector('#songSlidesList .slide-card[data-idx="' + next + '"]');
    if (activeCard) activeCard.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }
}

// ── Bible verse navigation — handles both list rows and slide cards ────────────
// Bible slide cards live in #bibleSlideGrid, list rows in #searchResults
function navigateSearchVerses(direction) {
  const isSlideMode = State.previewViewMode === 'slides' && State.currentTab === 'book';

  if (isSlideMode) {
    const cards = Array.from(document.querySelectorAll('#bibleSlideGrid .slide-card[data-verse]'));
    if (!cards.length) return false;

    let curIdx = cards.findIndex(c => c.classList.contains('active'));
    if (curIdx === -1) curIdx = direction > 0 ? -1 : cards.length;

    const nextIdx = curIdx + direction;
    if (nextIdx < 0 || nextIdx >= cards.length) return false;

    const card    = cards[nextIdx];
    const verse   = parseInt(card.dataset.verse);
    const partIdx = parseInt(card.dataset.partIdx  || '0');
    const ref     = State.currentBook + ' ' + State.currentChapter + ':' + verse;

    // Update highlight
    cards.forEach(c => c.classList.remove('active'));
    card.classList.add('active');
    card.scrollIntoView({ block: 'nearest', behavior: 'smooth' });

    // Set part state from the card's data attributes
    const parts = _splitVerseForTheme(State.currentBook, State.currentChapter, verse, ref);
    State.previewVerse  = { book: State.currentBook, chapter: State.currentChapter, verse, ref };
    State.verseParts    = parts;
    State.versePartIdx  = partIdx;
    _updateVersePartIndicator(partIdx + 1, parts.length);

    presentVerse(State.currentBook, State.currentChapter, verse, ref, { parts, partIdx });
    if (State.isLive) sendPreviewToLive();
    toast(`👁 ${ref}${parts.length > 1 ? ` · ${partIdx+1}/${parts.length}` : ''}`);
    return true;

  } else {
    // List mode — .verse-row items
    const rows = Array.from(document.querySelectorAll('#searchResults .verse-row[data-verse]'));
    if (!rows.length) return false;

    let curIdx = rows.findIndex(r => r.classList.contains('active'));
    if (curIdx === -1) curIdx = direction > 0 ? -1 : rows.length;

    const nextIdx = curIdx + direction;
    if (nextIdx < 0 || nextIdx >= rows.length) return false;

    const row   = rows[nextIdx];
    const verse = parseInt(row.dataset.verse);
    const ref   = State.currentBook + ' ' + State.currentChapter + ':' + verse;

    rows.forEach(r => {
      r.classList.remove('active','selected','current');
    });
    row.classList.add('active','selected','current');
    // Scroll within the container — not scrollIntoView which scrolls the page
    const _sr = document.getElementById('searchResults');
    if (_sr) scrollElementIntoContainer(row, _sr, 'center');

    presentVerse(State.currentBook, State.currentChapter, verse, ref);
    try { highlightVerseRow(verse); } catch (_) {}
    if (State.isLive) sendPreviewToLive();
    return true;
  }
}

// ─── REMOTE CONTROL POPOVER ───────────────────────────────────────────────────
function updateRemoteBtn(info) {
  const btn = document.getElementById('remoteBtn');
  if (!btn) return;
  const enabled = !info?.disabled && info?.ip;
  btn.classList.toggle('remote-active', !!enabled);
  btn.title = enabled
    ? `Remote: http://${info.ip}:${info.port}/remote (Click to manage)`
    : 'Remote Control (disabled)';
}

async function openRemotePopover() {
  if (!window.electronAPI) { toast('ℹ Remote Control requires the desktop app'); return; }

  // Fetch fresh info
  const info = await window.electronAPI.getRemoteInfo().catch(() => null);
  if (info) { State.remoteInfo = info; }
  const data = State.remoteInfo || info || { ip: null, port: 8080, enabled: false };

  // Remove any existing popover
  document.getElementById('remotePopover')?.remove();

  const enabled = data.enabled !== false && data.ip;
  const url = enabled ? `http://${data.ip}:${data.port}/remote` : null;

  const pop = document.createElement('div');
  pop.id = 'remotePopover';
  pop.style.cssText = `
    position:fixed; top:54px; right:16px; z-index:400;
    background:var(--panel); border:1px solid var(--border-lit);
    border-radius:10px; padding:18px 20px; width:320px;
    box-shadow:0 8px 40px rgba(0,0,0,0.7); animation:slideIn .15s ease;
  `;

  pop.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
      <div style="font-family:'Cinzel',serif;font-size:12px;font-weight:600;color:var(--gold-bright);letter-spacing:1px">📱 Remote Control</div>
      <button id="remotePopClose" style="background:none;border:none;color:var(--text-dim);font-size:16px;cursor:pointer;line-height:1;padding:2px 4px">✕</button>
    </div>

    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;padding:10px 12px;background:var(--card);border-radius:7px;border:1px solid var(--border)">
      <div>
        <div style="font-size:11px;font-weight:600;color:var(--text);margin-bottom:2px">Remote Server</div>
        <div style="font-size:10px;color:var(--text-dim)">Allow phones on same WiFi to control the app</div>
      </div>
      <div id="remoteToggle" class="remote-toggle ${enabled ? 'on' : 'off'}" title="${enabled ? 'Click to disable' : 'Click to enable'}">
        <div class="remote-toggle-knob"></div>
      </div>
    </div>

    <div id="remoteUrlSection" style="display:${enabled ? 'block' : 'none'}">
      <div style="font-size:10px;font-weight:600;color:var(--text-sub);text-transform:uppercase;letter-spacing:.7px;margin-bottom:6px">Open on phone/tablet (same WiFi):</div>
      <div style="background:var(--void);border:1px solid var(--border-lit);border-radius:6px;padding:10px 12px;display:flex;align-items:center;gap:8px;margin-bottom:8px">
        <div style="font-family:'Courier New',monospace;font-size:12px;color:var(--gold-bright);flex:1;word-break:break-all" id="remoteUrlText">${url || '—'}</div>
        <button id="remoteCopyBtn" title="Copy URL" style="padding:4px 8px;background:var(--gold-glow);border:1px solid var(--gold-dim);border-radius:4px;color:var(--gold);font-size:10px;font-weight:700;cursor:pointer;font-family:inherit;white-space:nowrap;flex-shrink:0">Copy</button>
      </div>
      <button id="remoteShareBtn" title="Copy shareable message for WhatsApp / SMS" style="width:100%;padding:7px 10px;background:var(--card);border:1px solid var(--border-lit);border-radius:6px;color:var(--text);font-size:11px;font-weight:600;cursor:pointer;font-family:inherit;display:flex;align-items:center;justify-content:center;gap:6px;margin-bottom:10px;transition:background .15s">
        📋 Copy Link to Share via WhatsApp / SMS
      </button>
      <div style="font-size:10px;color:var(--text-dim);line-height:1.6">
        Port: <strong style="color:var(--text)">${data.port}</strong>
        &nbsp;·&nbsp; IP: <strong style="color:var(--text)">${data.ip || '—'}</strong>
      </div>
      <div style="margin-top:10px">
        <div style="font-size:10px;font-weight:600;color:var(--text-sub);text-transform:uppercase;letter-spacing:.7px;margin-bottom:5px">Network Adapter</div>
        <select id="remoteAdapterSel" style="width:100%;background:var(--card);border:1px solid var(--border-lit);border-radius:5px;color:var(--text);font-family:inherit;font-size:11px;padding:6px 8px;outline:none">
          <option value="">Auto-detect</option>
          ${(data.adapters||[]).map(a=>`<option value="${a.name}" ${a.name===data.selectedAdapter?'selected':''}>${a.label}</option>`).join('')}
        </select>
        <div style="font-size:10px;color:var(--text-dim);margin-top:4px" id="adapterIpHint">Select the adapter your phone is on (same WiFi)</div>
      </div>
    </div>

    <div id="remoteOffSection" style="display:${enabled ? 'none' : 'block'}">
      <div style="font-size:11px;color:var(--text-dim);text-align:center;padding:8px 0">Remote control is disabled.<br>Toggle ON to start the server.</div>
    </div>

    <div style="margin-top:12px;padding-top:12px;border-top:1px solid var(--border);font-size:10px;color:var(--text-dim)">
      Change port in <a href="#" id="remoteOpenSettings" style="color:var(--gold);text-decoration:none">⚙ Settings → Audio tab</a>
    </div>
  `;

  document.body.appendChild(pop);

  // Toggle switch
  const toggle = pop.querySelector('#remoteToggle');
  toggle.addEventListener('click', async () => {
    const nowEnabled = !toggle.classList.contains('on');
    toggle.classList.toggle('on', nowEnabled);
    toggle.classList.toggle('off', !nowEnabled);
    const result = await window.electronAPI.toggleRemote(nowEnabled);
    State.remoteInfo = { ...data, ...result, enabled: nowEnabled };

    const urlSection = pop.querySelector('#remoteUrlSection');
    const offSection = pop.querySelector('#remoteOffSection');
    const urlEl = pop.querySelector('div[style*="Courier"]');

    if (nowEnabled && result.ip) {
      const newUrl = `http://${result.ip}:${result.port}/remote`;
      if (urlEl) urlEl.textContent = newUrl;
      urlSection.style.display = 'block';
      offSection.style.display = 'none';
      updateRemoteBtn({ ip: result.ip, port: result.port, disabled: false });
      toast(`📱 Remote ON — http://${result.ip}:${result.port}/remote`);
    } else {
      urlSection.style.display = 'none';
      offSection.style.display = 'block';
      updateRemoteBtn({ ip: null, disabled: true });
      toast('📵 Remote control disabled');
    }
  });

  // Adapter change
  pop.querySelector('#remoteAdapterSel')?.addEventListener('change', async e => {
    const adapterName = e.target.value || null;
    if (!window.electronAPI?.setNetworkAdapter) return;
    const result = await window.electronAPI.setNetworkAdapter(adapterName);
    if (result?.ip) {
      const newUrl = `http://${result.ip}:${data.port}/remote`;
      const urlEl = pop.querySelector('div[style*="Courier"]');
      const hint = pop.querySelector('#adapterIpHint');
      if (urlEl) urlEl.textContent = newUrl;
      if (hint) hint.textContent = `IP: ${result.ip} — URL updated above`;
      State.remoteInfo = { ...State.remoteInfo, ip: result.ip };
      updateRemoteBtn({ ip: result.ip, port: data.port, disabled: false });
      toast(`📶 Remote IP updated: ${result.ip}`);
    }
  });

  // Copy URL
  pop.querySelector('#remoteCopyBtn')?.addEventListener('click', () => {
    const currentUrl = pop.querySelector('#remoteUrlText')?.textContent || url;
    if (currentUrl && currentUrl !== '—') { navigator.clipboard?.writeText(currentUrl).catch(()=>{}); toast('📋 URL copied!'); }
  });

  // Copy shareable message
  pop.querySelector('#remoteShareBtn')?.addEventListener('click', () => {
    const currentUrl = pop.querySelector('#remoteUrlText')?.textContent || url;
    if (!currentUrl || currentUrl === '—') return;
    const msg = `📱 AnchorCast Remote Control\n\nOpen this link on your phone or tablet to control the worship presentation:\n\n${currentUrl}\n\n(Make sure you're connected to the same WiFi network)`;
    navigator.clipboard?.writeText(msg).then(() => {
      const btn = pop.querySelector('#remoteShareBtn');
      if (btn) { btn.textContent = '✓ Copied! Paste in WhatsApp or SMS'; setTimeout(() => { btn.innerHTML = '📋 Copy Link to Share via WhatsApp / SMS'; }, 2000); }
      toast('📋 Shareable message copied to clipboard!');
    }).catch(()=>{ toast('⚠ Could not copy to clipboard'); });
  });

  // Open settings
  pop.querySelector('#remoteOpenSettings')?.addEventListener('click', e => {
    e.preventDefault();
    openSettings();
    pop.remove();
  });

  // Close
  pop.querySelector('#remotePopClose').addEventListener('click', () => pop.remove());

  // Click outside to close
  setTimeout(() => {
    document.addEventListener('click', function outside(e) {
      if (!pop.contains(e.target) && e.target.id !== 'remoteBtn') {
        pop.remove();
        document.removeEventListener('click', outside);
      }
    });
  }, 50);
}
let recognition = null;
let demoInterval = null;
let micAnimInterval = null;

const SERMON_DEMO = [
  "Good morning church. Let us open our hearts to receive God's word today.",
  "We are going to begin in the Gospel of John, chapter three, and verse sixteen.",
  "For God so loved the world that he gave his only begotten Son.",
  "That whoever believes in him should not perish but have everlasting life.",
  "This verse is perhaps the most profound statement in all of scripture.",
  "Turn with me now to Romans chapter eight, verse twenty eight.",
  "We know that all things work together for good to those who love God.",
  "The prophet Isaiah in chapter forty verse thirty one reminds us:",
  "But those who wait on the Lord shall renew their strength.",
  "They shall mount up with wings like eagles, and run and not be weary.",
  "Paul writes to the church at Philippi, I can do all things through Christ who strengthens me.",
  "David wrote in Psalm twenty three, The Lord is my shepherd, I shall not want.",
  "Jeremiah twenty nine eleven says, I know the plans I have for you says the Lord.",
  "Plans for welfare and not for evil, to give you a future and a hope.",
  "And in John fourteen six, Jesus said, I am the way, the truth, and the life.",
];
let demoIdx = 0;

function toggleRecording() {
  if (State.isRecording) stopRecording();
  else startRecording();
}

function startRecording() {
  // Before starting mic — check if a transcription source is actually available
  const hasSrc = {
    local:    !!State.whisperLocalReady,
    deepgram: !!State.settings?.deepgramKey,
    cloud:    !!State.settings?.openAiKey,
  };
  if (!hasSrc[State.whisperSource]) {
    if (State.whisperSource === 'local') {
      // Show the detailed setup banner with a "Set Up Now" button
      _showWhisperSetupBanner({ reason: 'no_python', setupBatExists: true });
    } else if (State.whisperSource === 'deepgram') {
      toast('⚠️ No Deepgram API key — add one in Settings → Audio, or set up Local Whisper.');
    } else {
      toast('⚠️ No OpenAI API key — add one in Settings → Audio, or set up Local Whisper.');
    }
    // Don't start recording with no working source
    return;
  }

  State.isRecording = true;
  State.recordingStartTime = Date.now();
  const btn = document.getElementById('recordBtn');
  btn.textContent = '⏹ Stop Transcript';
  btn.classList.add('active');

  // Clear empty state
  const body = document.getElementById('transcriptBody');
  const empty = body.querySelector('.empty-state');
  if (empty) empty.remove();

  // Start adaptive memory session
  if (window.TranscriptMemory) {
    const profileSel = document.getElementById('tmSpeakerSelect');
    const profileId  = profileSel?.value || null;
    TranscriptMemory.startSession(profileId);
  }

  // Always use AudioWorklet PCM pipeline — works in Electron without Google
  startPcmCapture();
}

// ── AudioWorklet PCM capture pipeline ────────────────────────────────────────
// Mic → AudioContext → AudioWorklet → Float32 → downsample → Int16 PCM → IPC → Claude
let pcmAudioContext  = null;
let pcmWorkletNode   = null;
let pcmSourceNode    = null;
let pcmStream        = null;
let pcmTranscriptUnsub = null;

const TARGET_SAMPLE_RATE = 16000; // 16kHz — standard for speech recognition

function float32ToInt16(float32) {
  const out = new Int16Array(float32.length);
  for (let i = 0; i < float32.length; i++) {
    const s = Math.max(-1, Math.min(1, float32[i]));
    out[i] = s < 0 ? s * 0x8000 : s * 0x7fff;
  }
  return out;
}

function downsampleBuffer(input, inputRate, targetRate) {
  if (inputRate === targetRate) return input;
  const ratio = inputRate / targetRate;
  const outputLength = Math.round(input.length / ratio);
  const output = new Float32Array(outputLength);
  let offsetResult = 0;
  let offsetBuffer = 0;
  while (offsetResult < output.length) {
    const nextOffsetBuffer = Math.round((offsetResult + 1) * ratio);
    let accum = 0, count = 0;
    for (let i = offsetBuffer; i < nextOffsetBuffer && i < input.length; i++) {
      accum += input[i]; count++;
    }
    output[offsetResult] = count > 0 ? accum / count : 0;
    offsetResult++;
    offsetBuffer = nextOffsetBuffer;
  }
  return output;
}

async function startPcmCapture() {
  try {
    const audioConstraints = {
      channelCount: 1,
      echoCancellation: true,
      noiseSuppression: true,
      autoGainControl: true,
    };
    const savedMicId = State.settings?.microphoneId;
    if (savedMicId && savedMicId !== 'default') {
      audioConstraints.deviceId = { exact: savedMicId };
    }
    pcmStream = await navigator.mediaDevices.getUserMedia({
      audio: audioConstraints,
      video: false,
    });

    toast('🎙 Microphone active — speak now');

    // Start Deepgram WebSocket if that's the selected source
    if (State.whisperSource === 'deepgram') {
      startDeepgramStream();
    }

    // Create AudioContext
    pcmAudioContext = new AudioContext({ latencyHint: 'interactive' });
    const nativeRate = pcmAudioContext.sampleRate;

    // Load AudioWorklet processor
    // It needs to be served as a URL — our renderer server handles this
    const workletUrl = '/pcm-worklet-processor.js';
    await pcmAudioContext.audioWorklet.addModule(workletUrl);

    // Create nodes
    pcmSourceNode  = pcmAudioContext.createMediaStreamSource(pcmStream);
    pcmWorkletNode = new AudioWorkletNode(pcmAudioContext, 'pcm-capture-processor');
    pcmSourceNode.connect(pcmWorkletNode);

    startMicAnimation();

    // Receive Float32 chunks from AudioWorklet thread
    pcmWorkletNode.port.onmessage = (event) => {
      if (!State.isRecording) return;
      const floatChunk  = event.data;
      const downsampled = downsampleBuffer(floatChunk, nativeRate, TARGET_SAMPLE_RATE);
      const pcm16       = float32ToInt16(downsampled);

      if (State.whisperSource === 'deepgram' && deepgramConnected) {
        // Deepgram: send raw PCM directly over WebSocket — ~300ms latency
        sendPcmToDeepgram(pcm16.buffer);
      } else {
        // Local Whisper or OpenAI: accumulate in main process via IPC
        if (window.electronAPI?.pushAudioPcm) {
          window.electronAPI.pushAudioPcm(pcm16.buffer);
        }
      }
    };

    // Listen for transcript results from main process
    // BUG-17 FIX: guard the callback against late-arriving results after
    // stopRecording() has already been called. The !State.isRecording check
    // handles the normal case, but if the unsub function ever fails silently
    // we also check that the worklet node still exists (it is nulled in stopRecording).
    if (window.electronAPI?.onTranscript) {
      pcmTranscriptUnsub = window.electronAPI.onTranscript((text) => {
        if (!State.isRecording || !pcmWorkletNode || !text?.trim()) return;
        pushTranscriptLine(text.trim(), true);
      });
    }

    _startSilenceDetection(pcmAudioContext, pcmSourceNode);

  } catch(err) {
    if (err.name === 'NotAllowedError') {
      toast('⚠ Microphone denied — check OS Settings → Privacy → Microphone');
    } else if (err.name === 'NotFoundError') {
      toast('⚠ No microphone found — connect a mic and try again');
    } else if (err.name === 'OverconstrainedError') {
      toast('⚠ Selected microphone not available — using system default instead');
      console.warn('[PCM] Device not found, retrying with default');
      State.settings.microphoneId = 'default';
      try { return await startPcmCapture(); } catch (_) {}
      toast('⚠ No working microphone found');
      stopRecording();
      return;
    } else {
      toast(`⚠ Mic error: ${err.message}`);
      console.warn('[PCM]', err.name, err.message);
    }
    stopRecording();
  }
}

let _silenceDetectionTimer = null;
function _startSilenceDetection(audioCtx, sourceNode) {
  if (_silenceDetectionTimer) { clearTimeout(_silenceDetectionTimer); _silenceDetectionTimer = null; }
  if (!audioCtx || !sourceNode) return;
  try {
    const analyser = audioCtx.createAnalyser();
    analyser.fftSize = 2048;
    sourceNode.connect(analyser);
    const buf = new Uint8Array(analyser.frequencyBinCount);
    let silentChecks = 0;
    const CHECK_INTERVAL = 1000;
    const MAX_CHECKS = 5;
    const SILENCE_THRESHOLD = 2;

    function check() {
      if (!State.isRecording) { try { sourceNode.disconnect(analyser); } catch(_){} return; }
      analyser.getByteTimeDomainData(buf);
      let peak = 0;
      for (let i = 0; i < buf.length; i++) {
        const v = Math.abs(buf[i] - 128);
        if (v > peak) peak = v;
      }
      if (peak <= SILENCE_THRESHOLD) {
        silentChecks++;
        if (silentChecks >= MAX_CHECKS) {
          toast('⚠ No audio signal detected — check your microphone or select a different input device in Settings → Audio & AI');
          try { sourceNode.disconnect(analyser); } catch(_){}
          return;
        }
        _silenceDetectionTimer = setTimeout(check, CHECK_INTERVAL);
      } else {
        try { sourceNode.disconnect(analyser); } catch(_){}
      }
    }
    _silenceDetectionTimer = setTimeout(check, CHECK_INTERVAL);
  } catch(_) {}
}

function _stopSilenceDetection() {
  if (_silenceDetectionTimer) { clearTimeout(_silenceDetectionTimer); _silenceDetectionTimer = null; }
}

async function startSpeechAPI() {
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    startDemoMode();
    return;
  }

  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false });
    stream.getTracks().forEach(t => t.stop());
  } catch (err) {
    toast(`⚠ Microphone error: ${err.message}`);
    startDemoMode();
    return;
  }

  recognition = new SpeechRecognition();
  recognition.continuous = true;
  recognition.interimResults = true;
  recognition.lang = 'en-US';
  recognition.maxAlternatives = 1;

  recognition.onresult = (e) => {
    let interim = '';
    for (let i = e.resultIndex; i < e.results.length; i++) {
      if (e.results[i].isFinal) {
        const text = e.results[i][0].transcript.trim();
        if (text) pushTranscriptLine(text, true);
        removeInterim();
      } else {
        interim += e.results[i][0].transcript;
      }
    }
    if (interim) updateInterim(interim);
  };

  recognition.onerror = (e) => {
    if (e.error === 'network') {
      toast('⚠ Web Speech requires internet — add Claude API key in Settings for offline transcription');
      stopRecording();
    } else if (e.error === 'not-allowed') {
      toast('⚠ Microphone blocked — check OS privacy settings');
    } else if (e.error === 'no-speech') {
      return;
    }
  };

  recognition.onend = () => { if (State.isRecording) try { recognition.start(); } catch(e) {} };
  recognition.onstart = () => toast('🎙 Microphone active — speak now');

  try { recognition.start(); } catch(e) { startDemoMode(); }
}

function startDemoMode() {
  demoIdx = 0;
  toast('▶ Demo mode — simulating live sermon');
  demoInterval = setInterval(() => {
    if (!State.isRecording) return;
    const line = SERMON_DEMO[demoIdx % SERMON_DEMO.length];
    pushTranscriptLine(line, true);
    demoIdx++;
  }, 3200);
}

// ── Audio file transcription ──────────────────────────────────────────────────
// Decodes any audio/video file via Web Audio API, resamples to 16kHz PCM,
// sends through the same IPC pipeline as the live microphone
let fileTranscribeAbort = false;

async function transcribeAudioFile(file) {
  if (!window.electronAPI?.pushAudioPcm) {
    toast('⚠ Audio file transcription requires the desktop app');
    return;
  }

  // Check a transcription source is available
  if (State.whisperSource === 'local' && !State.whisperLocalReady) {
    _showWhisperSetupBanner({ reason: 'no_python', setupBatExists: true });
    return;
  }
  if (State.whisperSource === 'online' && !State.settings?.openAiKey) {
    toast('⚠ Add an OpenAI API key in Settings for cloud transcription');
    return;
  }

  toast(`📂 Loading: ${file.name}`);
  fileTranscribeAbort = false;

  try {
    // Read file as ArrayBuffer
    const arrayBuf = await file.arrayBuffer();

    // Decode audio (supports mp3, wav, m4a, mp4, ogg, webm, flac…)
    const audioCtx = new AudioContext({ sampleRate: 16000 }); // decode directly to 16kHz
    let audioBuffer;
    try {
      audioBuffer = await audioCtx.decodeAudioData(arrayBuf);
    } catch(e) {
      toast(`⚠ Cannot decode "${file.name}" — try MP3, WAV, or M4A`);
      await audioCtx.close();
      return;
    }

    const sampleRate   = audioBuffer.sampleRate;     // should be 16000
    const numChannels  = audioBuffer.numberOfChannels;
    const totalSamples = audioBuffer.length;
    await audioCtx.close();

    // Mix down to mono
    const monoData = new Float32Array(totalSamples);
    for (let ch = 0; ch < numChannels; ch++) {
      const chData = audioBuffer.getChannelData(ch);
      for (let i = 0; i < totalSamples; i++) monoData[i] += chData[i];
    }
    if (numChannels > 1) {
      for (let i = 0; i < totalSamples; i++) monoData[i] /= numChannels;
    }

    // Start "recording" state for UI
    if (State.isRecording) stopRecording();
    State.isRecording = true;
    State.recordingStartTime = Date.now();
    const btn = document.getElementById('recordBtn');
    btn.textContent = '⏹ Stop';
    btn.classList.add('active');
    startMicAnimation();
    const body = document.getElementById('transcriptBody');
    const empty = body.querySelector('.empty-state');
    if (empty) empty.remove();

    // Wire transcript result listener
    if (window.electronAPI?.onTranscript && !pcmTranscriptUnsub) {
      pcmTranscriptUnsub = window.electronAPI.onTranscript((text) => {
        if (!text?.trim()) return;
        pushTranscriptLine(text.trim(), true);
      });
    }

    const CHUNK_SAMPLES = Math.floor(sampleRate * 2.5); // 2.5s chunks
    const totalChunks   = Math.ceil(totalSamples / CHUNK_SAMPLES);
    const durationSec   = totalSamples / sampleRate;

    toast(`🎵 Transcribing ${file.name} (${Math.round(durationSec)}s)…`);
    document.getElementById('audioUploadBtn').textContent = '⏳';
    document.getElementById('audioUploadBtn').disabled = true;

    // Send chunks with small gaps so the IPC queue doesn't flood
    for (let i = 0; i < totalChunks; i++) {
      if (fileTranscribeAbort || !State.isRecording) break;

      const start  = i * CHUNK_SAMPLES;
      const end    = Math.min(start + CHUNK_SAMPLES, totalSamples);
      const chunk  = monoData.slice(start, end);

      // Convert Float32 → Int16
      const int16 = new Int16Array(chunk.length);
      for (let j = 0; j < chunk.length; j++) {
        const s = Math.max(-1, Math.min(1, chunk[j]));
        int16[j] = s < 0 ? s * 0x8000 : s * 0x7fff;
      }

      // Update mic animation to show progress
      const progress = Math.round(((i + 1) / totalChunks) * 100);
      document.querySelectorAll('#micLevels .mic-bar').forEach((b, bi) => {
        b.style.height = (progress / 100 > bi / 5 ? 12 : 4) + 'px';
      });

      await window.electronAPI.pushAudioPcm(int16.buffer);

      // Pace the chunks — don't send faster than real-time to avoid queue flood
      await new Promise(r => setTimeout(r, 400));
    }

    // Wait for last transcription to come back
    await new Promise(r => setTimeout(r, 4000));

    toast(`✓ File transcription complete`);

  } catch(e) {
    console.warn('[FileTranscribe]', e.message);
    toast(`⚠ Transcription failed: ${e.message}`);
  } finally {
    // Restore button
    const uploadBtn = document.getElementById('audioUploadBtn');
    if (uploadBtn) { uploadBtn.textContent = '🎵 File'; uploadBtn.disabled = false; }
    // Stop recording state
    if (State.isRecording) {
      State.isRecording = false;
      const btn = document.getElementById('recordBtn');
      btn.textContent = '▶ Start Transcript';
      btn.classList.remove('active');
      stopMicAnimation();
      if (pcmTranscriptUnsub) { pcmTranscriptUnsub(); pcmTranscriptUnsub = null; }
      saveTranscriptHistory();
      if (window.TranscriptMemory) {
        TranscriptMemory.endSession();
        if (State.settings?.adaptiveLearningEnabled !== false) {
          setTimeout(() => TranscriptMemory.runLearningJob(), 1200);
        }
      }
    }
  }
}

function stopRecording() {
  fileTranscribeAbort = true; // abort any in-progress file transcription
  State.isRecording = false;
  _stopSilenceDetection();
  const btn = document.getElementById('recordBtn');
  btn.textContent = '▶ Start Transcript';
  btn.classList.remove('active');
  stopMicAnimation();
  removeInterim();

  // Stop Deepgram WebSocket
  stopDeepgramSocket();

  // Stop Web Speech API fallback
  if (recognition) { try { recognition.stop(); } catch(e){} recognition = null; }

  // Stop PCM AudioWorklet pipeline
  if (pcmWorkletNode)  { try { pcmWorkletNode.disconnect(); }  catch(e){} pcmWorkletNode = null; }
  if (pcmSourceNode)   { try { pcmSourceNode.disconnect(); }   catch(e){} pcmSourceNode  = null; }
  if (pcmStream)       { pcmStream.getTracks().forEach(t => t.stop()); pcmStream = null; }
  if (pcmAudioContext) { try { pcmAudioContext.close(); }       catch(e){} pcmAudioContext = null; }
  if (pcmTranscriptUnsub) { pcmTranscriptUnsub(); pcmTranscriptUnsub = null; }

  // Stop demo mode
  if (demoInterval)    { clearInterval(demoInterval); demoInterval = null; }

  // Auto-save transcript to history on every stop
  saveTranscriptHistory().then(() => {
    window.electronAPI?.setTranscriptUnsaved?.(false);
    toast('💾 Transcript auto-saved');
  }).catch(() => {});

  // End adaptive memory session and trigger offline learning
  if (window.TranscriptMemory) {
    TranscriptMemory.endSession();
    if (State.settings?.adaptiveLearningEnabled !== false) {
      setTimeout(() => TranscriptMemory.runLearningJob(), 2000);
    }
  }

  toast('⏹ Transcription stopped');
}

// ── Client-side Biblical word corrections ────────────────────────────────────
// Runs on every transcript line as a fast pre-pass before adaptive memory.
// Fixes common Whisper mishearings of Biblical/church vocabulary.
const _BIBLICAL_FIXES = [
  [/shall not watch/gi,              'shall not want'],
  [/shall not wash/gi,               'shall not want'],
  [/shall not won/gi,                'shall not want'],
  [/I shall not pass\b/gi,           'I shall not want'],
  [/I shall not past\b/gi,           'I shall not want'],
  [/\bgreen pass\b/gi,               'green pastures'],
  [/\bgreen pace\b/gi,               'green pastures'],
  [/\bgreen pastor\b/gi,             'green pastures'],
  [/still water(?!s)/gi,             'still waters'],
  [/restoreth my son/gi,             'restoreth my soul'],
  [/restored my son/gi,              'restoreth my soul'],
  [/restore my son/gi,               'restoreth my soul'],
  [/leadeth my son/gi,               'leadeth my soul'],
  [/rivers of righteousness/gi,      'paths of righteousness'],
  [/walk through the back/gi,        'walk through the valley'],
  [/through the back\b/gi,           'through the valley'],
  [/shadow death/gi,                 'shadow of death'],
  [/I will fear and all evil/gi,     'I will fear no evil'],
  [/\bfear and all evil/gi,          'fear no evil'],
  [/my cup runny(?:\s+over)?\b/gi,   'my cup runneth over'],
  [/my cup running(?:\s+over)?\b/gi,'my cup runneth over'],
  [/\bthigh kingdom/gi,              'thy kingdom'],
  [/\bethernet\b/gi,                 'eternal'],
  [/\bhollowly\b/gi,                 'hallowed'],
  [/hollow be thy name/gi,           'hallowed be thy name'],
  [/hallow be thy name/gi,           'hallowed be thy name'],
  [/our daily broad\b/gi,            'our daily bread'],
  [/goodness mercy/gi,               'goodness and mercy'],
  [/\brod staff\b/gi,                'rod and staff'],
  [/\bsabour\b/gi,                   'saviour'],
  [/\brevelations\b/gi,              'Revelation'],
  [/\bresolutions\s+(chapter|chap\b|\d)/gi, 'Revelation $1'],
  [/\brevolution\s+(chapter|chap\b|\d)/gi, 'Revelation $1'],
  [/\bis releasing as he said\b/gi,  'is risen as he said'],
  [/\bhe is releasing\b/gi,          'He is risen'],
  [/\brising as he said\b/gi,        'risen as he said'],
  [/\bcoming out of the green\b/gi,  'coming out of the grave'],
  [/\bout of that grace\b/gi,        'out of that grave'],
  [/\bart thou cast down\b/gi,       'art thou cast down'],
  [/\bhealth of his countenance\b/gi,'help of his countenance'],
  [/\bhope go and go\b/gi,           'hope thou in God'],
  [/\bquiet type in me\b/gi,         'disquieted in me'],
  [/\bthigh rod\b/gi,                'thy rod'],
  [/\bthigh will\b/gi,               'thy will'],
  [/\bthigh word\b/gi,               'thy word'],
  [/\bthigh hand\b/gi,               'thy hand'],
  [/\bin the beginning was the world\b/gi, 'in the beginning was the Word'],
  [/\bfor God's love the world\b/gi, 'for God so loved the world'],
  [/\blamb of got\b/gi,              'Lamb of God'],
  [/\blamb of guard\b/gi,            'Lamb of God'],
  [/\bholy goes\b/gi,                'Holy Ghost'],
  [/\bholy coast\b/gi,               'Holy Ghost'],
  [/\bholy goats\b/gi,               'Holy Ghost'],
  [/\blassarod\b/gi,                 'Lazarus'],
  [/\blazarod\b/gi,                  'Lazarus'],
  [/\blazarous\b/gi,                 'Lazarus'],
  [/\blazzarus\b/gi,                 'Lazarus'],
  [/\bfaracees\b/gi,                 'Pharisees'],
  [/\bfarisees\b/gi,                 'Pharisees'],
  [/\bfarisee\b/gi,                  'Pharisee'],
  [/\bgalilee\b/gi,                  'Galilee'],
  [/\bgethsemani\b/gi,              'Gethsemane'],
  [/\bgolgotha\b/gi,                'Golgotha'],
  [/\bnazaret\b/gi,                  'Nazareth'],
  [/\bjerusalom\b/gi,               'Jerusalem'],
  [/\bjeresalem\b/gi,               'Jerusalem'],
  [/\bbethleham\b/gi,               'Bethlehem'],
  [/\baphesians\b/gi,               'Ephesians'],
  [/\befesians\b/gi,                'Ephesians'],
  [/\bgalatians\b/gi,               'Galatians'],
  [/\bfilipians\b/gi,               'Philippians'],
  [/\bfilippians\b/gi,              'Philippians'],
  [/\bcollosians\b/gi,              'Colossians'],
  [/\bcolosians\b/gi,               'Colossians'],
  [/\bthessalonions\b/gi,           'Thessalonians'],
  [/\bthesalonians\b/gi,            'Thessalonians'],
  [/\bdeuteronomy\b/gi,             'Deuteronomy'],
  [/\bduteronomy\b/gi,              'Deuteronomy'],
  [/\blevitikus\b/gi,               'Leviticus'],
  [/\becclesiastes\b/gi,            'Ecclesiastes'],
  [/\becclesiasties\b/gi,           'Ecclesiastes'],
  [/\becclesiast\b/gi,              'Ecclesiastes'],
  [/\becclesiat\b/gi,               'Ecclesiastes'],
];
function _applyBiblicalFixes(text) {
  let out = text;
  for (const [pat, rep] of _BIBLICAL_FIXES) out = out.replace(pat, rep);
  return out;
}

function pushTranscriptLine(text, detectVerses = false) {
  if (!text?.trim()) return;

  // Apply Biblical word corrections before everything else
  const correctedInput = _applyBiblicalFixes(text);

  let displayText = correctedInput;
  let chunkId = null;
  let appliedRules = [];

  // ── Adaptive correction pass ──────────────────────────────────────────────
  if (window.TranscriptMemory?.enabled) {
    const result = TranscriptMemory.process(correctedInput);
    displayText = result.corrected || correctedInput;
    chunkId = result.chunkId;
    appliedRules = result.appliedRules || [];
  }

  const line = {
    id: `tl_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
    text: displayText,
    raw: correctedInput,
    chunkId,
    time: Date.now(),
    corrected: displayText !== correctedInput,  // true when ANY fix was applied
    appliedRules,
  };
  State.transcriptLines.push(line);
  // Mark transcript as unsaved so the close-safeguard dialog fires
  window.electronAPI?.setTranscriptUnsaved?.(true);
  _pushReplayEvent('transcript-line', {
    text: line.text || '',
    transcript: line.text || '',
    liveText: line.raw && line.raw !== line.text ? `Original: ${line.raw}` : '',
    summary: line.text || ''
  });

  const body = document.getElementById('transcriptBody');
  body.querySelectorAll('.transcript-line.current').forEach(el => el.classList.remove('current'));

  const div = document.createElement('div');
  div.className = 'transcript-line current';
  div.dataset.chunkId = chunkId || '';
  div.dataset.rawText = correctedInput;
  div.dataset.lineId = line.id;
  div.textContent = displayText;
  div.addEventListener('click', () => selectTranscriptReviewLine(line.id));
  div.addEventListener('dblclick', () => { openTranscriptReviewPanel(); selectTranscriptReviewLine(line.id); });

  // Show correction indicator if text was changed
  if (displayText !== text) {  // show indicator if any fix changed the original
    div.title = `Original: "${text}"`;
    div.style.borderLeft = '2px solid rgba(201,168,76,.4)';
    div.style.paddingLeft = '6px';

    const sug = document.createElement('div');
    sug.className = 'transcript-inline-suggestion';
    sug.style.cssText = 'margin-top:6px;padding:6px 8px;border:1px solid rgba(201,168,76,.18);border-radius:6px;background:rgba(201,168,76,.06);font-size:11px;color:var(--text-dim)';
    sug.innerHTML = `Did you mean: <strong style="color:var(--gold)">${escapeHtml(displayText)}</strong> instead of <span>${escapeHtml(text)}</span>? <button type="button" style="margin-left:8px" class="btn-sm">Review</button>`;
    sug.querySelector('button')?.addEventListener('click', () => { openTranscriptReviewPanel(); selectTranscriptReviewLine(line.id); });
    div.appendChild(sug);
  }

  body.appendChild(div);
  body.scrollTop = body.scrollHeight;

  if (detectVerses && window.AIDetection) {
    AIDetection.processText(displayText, State.isOnline);
  }
}

function updateInterim(text) {
  let el = document.getElementById('interimLine');
  if (!el) {
    el = document.createElement('div');
    el.id = 'interimLine';
    el.className = 'transcript-interim';
    document.getElementById('transcriptBody').appendChild(el);
  }
  el.innerHTML = text + '<span class="cursor-blink"></span>';
  document.getElementById('transcriptBody').scrollTop = 99999;
}

function removeInterim() {
  const el = document.getElementById('interimLine');
  if (el) el.remove();
}

function clearTranscript() {
  State.transcriptLines = [];
  document.getElementById('transcriptBody').innerHTML =
    `<div class="empty-state"><span class="empty-icon">🎤</span>Start transcription to see the live sermon text here.</div>`;
  toast('↺ Transcript cleared');
}

async function exportTranscript() {
  const text = State.transcriptLines.map(l => l.text).join('\n');
  if (!text) { toast('⚠ No transcript to export'); return; }
  if (window.electronAPI) {
    const result = await window.electronAPI.exportTranscript({
      content: text,
      defaultName: `sermon-${new Date().toISOString().slice(0,10)}.txt`
    });
    if (result.success) toast('↓ Transcript exported');
  } else {
    const blob = new Blob([text], { type: 'text/plain' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `sermon-${new Date().toISOString().slice(0,10)}.txt`;
    a.click();
    toast('↓ Transcript exported');
  }
}

async function saveTranscriptHistory() {
  const text = State.transcriptLines.map(l => l.text).join('\n');
  if (!text || !window.electronAPI) return;
  const words = text.split(/\s+/).filter(Boolean).length;
  const duration = Math.round((Date.now() - (State.recordingStartTime || Date.now())) / 1000);
  await window.electronAPI.saveTranscript({
    title: `Sermon — ${new Date().toLocaleDateString('en-US', { weekday: 'long', month: 'short', day: 'numeric' })}`,
    text,
    lineCount: State.transcriptLines.length,
    verseCount: State.detections.length,
    wordCount: words,
    duration,
    detectedVerses: State.detections.map(d => d.ref),
    date: new Date().toISOString(),
  });
}

// ─── MIC ANIMATION ────────────────────────────────────────────────────────────
let _micAnalyser = null;
let _micAnalyserBuf = null;
function startMicAnimation() {
  const bars = document.querySelectorAll('#micLevels .mic-bar');
  if (pcmAudioContext && pcmSourceNode) {
    try {
      _micAnalyser = pcmAudioContext.createAnalyser();
      _micAnalyser.fftSize = 256;
      pcmSourceNode.connect(_micAnalyser);
      _micAnalyserBuf = new Uint8Array(_micAnalyser.frequencyBinCount);
    } catch(_) { _micAnalyser = null; }
  }
  micAnimInterval = setInterval(() => {
    if (_micAnalyser && _micAnalyserBuf) {
      _micAnalyser.getByteTimeDomainData(_micAnalyserBuf);
      let peak = 0;
      for (let i = 0; i < _micAnalyserBuf.length; i++) {
        const v = Math.abs(_micAnalyserBuf[i] - 128);
        if (v > peak) peak = v;
      }
      const level = peak / 128;
      bars.forEach((b, i) => {
        const threshold = (i + 1) / bars.length;
        const h = level > threshold * 0.3 ? Math.max(4, level * 18) : 4;
        b.style.height = h + 'px';
        b.style.background = level < 0.02 ? '#555' : '';
      });
    } else {
      bars.forEach(b => { b.style.height = '4px'; b.style.background = '#555'; });
    }
  }, 80);
}
function stopMicAnimation() {
  clearInterval(micAnimInterval);
  if (_micAnalyser && pcmSourceNode) {
    try { pcmSourceNode.disconnect(_micAnalyser); } catch(_) {}
  }
  _micAnalyser = null;
  _micAnalyserBuf = null;
  document.querySelectorAll('#micLevels .mic-bar').forEach(b => { b.style.height = '4px'; b.style.background = ''; });
}

// ─── AI DETECTION HANDLER ─────────────────────────────────────────────────────
function handleDetection(detection) {
  const lastLine = (State.transcriptLines || []).slice(-1)[0] || null;

  // Deduplicate exact ref + same source phrase
  if (State.detections.find(d => d.ref === detection.ref && String(d.sourceText || '') === String(lastLine?.text || lastLine?.raw || ''))) return;

  // Apply confidence threshold by detection type
  const type = String(detection.type || 'keyword').toLowerCase();
  const thresholdMap = {
    direct: parseFloat(State.settings.detectionThresholdDirect ?? 0.80),
    verbal: parseFloat(State.settings.detectionThresholdVerbal ?? 0.85),
    keyword: parseFloat(State.settings.detectionThresholdKeyword ?? 0.80),
    content: parseFloat(State.settings.detectionThresholdContent ?? 0.60),
    learned: parseFloat(State.settings.detectionThresholdLearned ?? 0.75),
  };
  const threshold = thresholdMap[type] ?? parseFloat(State.settings.confidenceThreshold || 0.75);
  if (detection.confidence < threshold) return;

  const det = {
    ...detection,
    id: `det_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
    sourceText: detection.sourcePhrase || lastLine?.text || lastLine?.raw || '',
    sourceLineId: lastLine?.id || null,
    status: 'pending',
    createdAt: Date.now(),
  };

  State.detections.unshift(det);
  _pushReplayEvent('verse-detection', {
    ref: det.ref || '',
    transcript: det.sourceText || '',
    liveText: det.text || '',
    summary: `${det.ref || ''} (${det.type || 'detection'})`
  });
  renderDetections();

  // ── Whisper Reinforcement ──────────────────────────────────────────────────
  // Feed the detected verse text back into Whisper's context prompt.
  // This biases the next 30s of transcription toward the same passage vocabulary,
  // dramatically improving detection accuracy when the preacher quotes the same verse.
  if (det.text && window.electronAPI?.reinforceWhisper) {
    // Send verse text + ref as reinforcement (e.g. "John 3:16. For God so loved the world...")
    const reinforceText = `${det.ref}. ${det.text}`.slice(0, 350);
    window.electronAPI.reinforceWhisper({ text: reinforceText, ttl: 30 }).catch(() => {});
  }

  // Auto mode: automatically present high-confidence verses
  if (State.isAutoMode && det.confidence >= 0.88) {
    navigateBibleSearch(det.book, det.chapter, det.verse);
    presentVerse(det.book, det.chapter, det.verse, det.ref);
    if (State.isLive) sendPreviewToLive();
  }
}

function renderDetections() {
  const body = document.getElementById('detectionsBody');
  const emptyEl = document.getElementById('detectionsEmpty');
  if(emptyEl) emptyEl.style.display = State.detections.length ? 'none' : 'block';

  // Remove old detection items (keep the empty-state element)
  body.querySelectorAll('.detection-item').forEach(e => e.remove());

  // State.detections[0] = newest (via unshift). Render newest at top, oldest at bottom.
  State.detections.forEach((det, i) => {
    const preview = (det.text || '').slice(0, 75) + (det.text?.length > 75 ? '…' : '');
    const confPct = Math.round(det.confidence * 100);
    const confClass = confPct >= 85 ? 'high' : confPct >= 65 ? 'med' : 'low';

    const div = document.createElement('div');
    div.className = 'detection-item';
    div.innerHTML = `
      <div class="det-header">
        <div class="det-dot ${confClass}"></div>
        <div class="det-ref">${det.ref}</div>
        <span class="det-type ${det.type || 'keyword'}">${det.type || 'AI'}</span>
        <div class="det-score">${confPct}%</div>
      </div>
      <div class="det-preview">${preview}</div>
      <div class="det-actions">
        <button class="det-btn det-btn-present">▶ Present</button>
        <button class="det-btn det-btn-queue">+ Scheduled</button>
        <button class="det-btn det-btn-review">📝 Review</button>
      </div>
    `;

    div.querySelector('.det-btn-present').addEventListener('click', () => {
      presentVerse(det.book, det.chapter, det.verse, det.ref);
      _pushReplayEvent('present-verse', { ref: det.ref || '', transcript: det.sourceText || '', liveText: det.text || '', summary: det.ref || '' });
      navigateBibleSearch(det.book, det.chapter, det.verse);
    });
    div.querySelector('.det-btn-queue').addEventListener('click', () =>
      addToQueue(det.book, det.chapter, det.verse, det.ref));
    div.querySelector('.det-btn-review').addEventListener('click', () => { openDetectionReviewPanel(); selectDetectionReview(det.id); });

    body.appendChild(div);
  });

  if (body) body.scrollTop = 0;
}

function clearDetections() {
  State.detections = [];
  // Clear transcriptLines too so old text can't re-trigger detections
  // when transcription restarts
  State.transcriptLines = [];
  renderDetections();
  // Clear the AI detection engine cache
  if (window.AIDetection) AIDetection.clearCache();
  toast('🗑 Detections cleared');
}


// ─── VERSE DETECTION REVIEW PANEL ────────────────────────────────────────────
function getDetectionReviewItem() {
  return (State.detections || []).find(d => d.id === State.reviewDetectionId) || null;
}

function openDetectionReviewPanel() {
  const modal = document.getElementById('detectionReviewModal');
  if (modal) modal.style.display = 'flex';
  renderDetectionReviewList();
  if (!getDetectionReviewItem() && State.detections?.length) {
    selectDetectionReview(State.detections[0].id);
  }
}

function closeDetectionReviewPanel() {
  const modal = document.getElementById('detectionReviewModal');
  if (modal) modal.style.display = 'none';
}

function renderDetectionReviewList() {
  const list = document.getElementById('detectionReviewList');
  const count = document.getElementById('detectionReviewCount');
  const query = (document.getElementById('detectionReviewSearch')?.value || '').trim().toLowerCase();
  if (!list) return;
  const items = (State.detections || []).filter(d => {
    if (!query) return true;
    return String(d.ref || '').toLowerCase().includes(query) || String(d.sourceText || '').toLowerCase().includes(query) || String(d.text || '').toLowerCase().includes(query);
  });
  list.innerHTML = '';
  if (count) count.textContent = `${items.length} detections`;
  if (!items.length) {
    list.innerHTML = `<div class="empty-state" style="min-height:120px"><span class="empty-icon">🔮</span>No detections to review.</div>`;
    return;
  }
  items.forEach(det => {
    const card = document.createElement('button');
    card.type = 'button';
    card.style.cssText = 'text-align:left;background:var(--card);border:1px solid var(--border);border-radius:8px;padding:10px;cursor:pointer;color:var(--text)';
    if (det.id === State.reviewDetectionId) card.style.borderColor = 'var(--gold)';
    const statusColor = det.status === 'approved' ? 'var(--live)' : det.status === 'rejected' ? '#ff6b6b' : 'var(--gold)';
    card.innerHTML = `
      <div style="display:flex;gap:8px;align-items:center;margin-bottom:4px">
        <span style="font-size:10px;color:${statusColor}">${escapeHtml(det.status || 'pending')}</span>
        <span style="font-size:10px;color:var(--text-dim)">${escapeHtml(det.type || 'ai')}</span>
        <span style="margin-left:auto;font-size:10px;color:var(--text-dim)">${Math.round(Number(det.confidence || 0) * 100)}%</span>
      </div>
      <div style="font-size:12px;color:var(--text);font-weight:700">${escapeHtml(det.ref || '')}</div>
      <div style="font-size:10px;color:var(--text-dim);margin-top:4px">${escapeHtml(String(det.sourceText || '').slice(0,120))}</div>
    `;
    card.addEventListener('click', () => selectDetectionReview(det.id));
    list.appendChild(card);
  });
}

function selectDetectionReview(id) {
  State.reviewDetectionId = id;
  const det = getDetectionReviewItem();
  if (!det) return renderDetectionReviewList();
  document.getElementById('detectionReviewRef').value = det.ref || '';
  document.getElementById('detectionReviewMeta').value = `${det.type || 'ai'} · ${Math.round(Number(det.confidence || 0) * 100)}%`;
  document.getElementById('detectionReviewSource').value = det.sourceText || '';
  document.getElementById('detectionReviewVerseText').value = det.text || '';
  document.getElementById('detectionReplaceRef').value = det.ref || '';
  const badge = document.getElementById('detectionReviewBadge');
  if (badge) badge.textContent = `${det.status || 'pending'} detection`;
  renderDetectionReviewList();
}

async function persistDetectionFeedback(det, action, accepted = null) {
  try {
    await window.electronAPI?.saveDetectionFeedback?.({
      detectionId: det.id,
      ref: det.ref,
      sourceText: det.sourceText || '',
      action,
      method: det.type || '',
      confidence: det.confidence || 0,
      accepted,
    });
  } catch(_) {}
}

async function approveDetectionReview() {
  const det = getDetectionReviewItem();
  if (!det) return;
  det.status = 'approved';
  await persistDetectionFeedback(det, 'approve', true);
  selectDetectionReview(det.id);
  toast('✓ Detection approved');
}

async function rejectDetectionReview() {
  const det = getDetectionReviewItem();
  if (!det) return;
  det.status = 'rejected';
  await persistDetectionFeedback(det, 'reject', false);
  selectDetectionReview(det.id);
  toast('↺ Detection rejected');
}

function presentDetectionReview() {
  const det = getDetectionReviewItem();
  if (!det) return;
  presentVerse(det.book, det.chapter, det.verse, det.ref);
  navigateBibleSearch(det.book, det.chapter, det.verse);
  toast('▶ Detection presented');
}

async function replaceDetectionReview() {
  const det = getDetectionReviewItem();
  if (!det) return;
  const nextRef = (document.getElementById('detectionReplaceRef')?.value || '').trim();
  const parsed = window.BibleDB?.parseReference ? window.BibleDB.parseReference(nextRef) : null;
  if (!parsed) return toast('⚠ Enter a valid Bible reference');
  const verseText = window.BibleDB?.getVerse(parsed.book, parsed.chapter, parsed.verse || 1, State.currentTranslation) || window.BibleDB?.getVerse(parsed.book, parsed.chapter, parsed.verse || 1);
  det.book = parsed.book;
  det.chapter = parsed.chapter;
  det.verse = parsed.verse || 1;
  det.ref = `${parsed.book} ${parsed.chapter}:${parsed.verse || 1}`;
  det.text = verseText || det.text;
  det.status = 'replaced';
  await persistDetectionFeedback(det, 'replace', true);
  selectDetectionReview(det.id);
  renderDetections();
  toast('✓ Detection replaced');
}

async function teachDetectionPhrase() {
  const det = getDetectionReviewItem();
  if (!det) return;
  const phrase = (document.getElementById('detectionReviewSource')?.value || det.sourceText || '').trim();
  if (!phrase) return toast('⚠ Add a source phrase first');
  const res = await window.electronAPI?.saveDetectionPhrase?.({ phrase, ref: det.ref, type: det.type || 'learned', createdBy: 'user' });
  if (res?.success) {
    det.status = 'taught';
    det.sourceText = phrase;
    await persistDetectionFeedback(det, 'teach-phrase', true);
    await window.AIDetection?.reloadLearnedPhrases?.();
    selectDetectionReview(det.id);
    toast('🧠 Phrase taught for offline detection');
  } else {
    toast('⚠ Could not teach this phrase');
  }
}


// ─── VERSE DISPLAY ────────────────────────────────────────────────────────────
function getVerseText(book, chapter, verse) {
  if (!window.BibleDB) return null;
  return BibleDB.getVerse(book, chapter, verse, State.currentTranslation);
}

const _BOOK_ABBREVS = {
  'Genesis':'Gen','Exodus':'Ex','Leviticus':'Lev','Numbers':'Num','Deuteronomy':'Deut',
  'Joshua':'Josh','Judges':'Judg','Ruth':'Ruth','1 Samuel':'1 Sam','2 Samuel':'2 Sam',
  '1 Kings':'1 Kgs','2 Kings':'2 Kgs','1 Chronicles':'1 Chr','2 Chronicles':'2 Chr',
  'Ezra':'Ezra','Nehemiah':'Neh','Esther':'Esth','Job':'Job','Psalms':'Ps','Psalm':'Ps',
  'Proverbs':'Prov','Ecclesiastes':'Eccl','Song of Solomon':'Song',
  'Isaiah':'Isa','Jeremiah':'Jer','Lamentations':'Lam','Ezekiel':'Ezek','Daniel':'Dan',
  'Hosea':'Hos','Joel':'Joel','Amos':'Amos','Obadiah':'Obad','Jonah':'Jonah',
  'Micah':'Mic','Nahum':'Nah','Habakkuk':'Hab','Zephaniah':'Zeph','Haggai':'Hag',
  'Zechariah':'Zech','Malachi':'Mal',
  'Matthew':'Matt','Mark':'Mark','Luke':'Luke','John':'John','Acts':'Acts',
  'Romans':'Rom','1 Corinthians':'1 Cor','2 Corinthians':'2 Cor',
  'Galatians':'Gal','Ephesians':'Eph','Philippians':'Phil','Colossians':'Col',
  '1 Thessalonians':'1 Thess','2 Thessalonians':'2 Thess',
  '1 Timothy':'1 Tim','2 Timothy':'2 Tim','Titus':'Titus','Philemon':'Phlm',
  'Hebrews':'Heb','James':'Jas','1 Peter':'1 Pet','2 Peter':'2 Pet',
  '1 John':'1 John','2 John':'2 John','3 John':'3 John','Jude':'Jude','Revelation':'Rev'
};

function _abbreviateRef(ref) {
  if (!ref) return ref;
  for (const [full, abbr] of Object.entries(_BOOK_ABBREVS)) {
    if (ref.startsWith(full)) return abbr + ref.slice(full.length);
  }
  return ref;
}

function _getRefOutlineCss(sfs) {
  const st = (sfs || {}).scriptureRefOutlineStyle || 'none';
  if (st === 'none') return '';
  const w = (sfs || {}).scriptureRefOutlineWidth ?? 0;
  if (w <= 0) return '';
  const c = (sfs || {}).scriptureRefOutlineColor || '#000';
  const o = ((sfs || {}).scriptureRefOutlineOpacity ?? 100) / 100;
  const r = parseInt(c.slice(1,3),16)||0, g = parseInt(c.slice(3,5),16)||0, b = parseInt(c.slice(5,7),16)||0;
  return `-webkit-text-stroke:${w}px rgba(${r},${g},${b},${o});paint-order:stroke fill;`;
}

function _formatRef(ref) {
  const sfs = State.settings || {};
  return sfs.scriptureAbbreviateBooks ? _abbreviateRef(ref) : ref;
}

function buildVerseHTML(book, chapter, verse, ref) {
  const text = getVerseText(book, chapter, verse);
  if (!text) return null;
  const transform = _getScriptureTransform();
  const sfs = State.settings || {};
  const dRef = _formatRef(ref);
  const refOnly = !!sfs.scriptureShowRefOnly;
  const extraLSStyle = sfs.scriptureAdditionalLineSpacing ? 'line-height:1.8;' : '';
  const roc = _getRefOutlineCss(sfs);
  return `
    <div class="verse-ref-label" style="${roc}">${dRef} &nbsp;·&nbsp; ${State.currentTranslation}</div>
    ${refOnly ? '' : `<div class="verse-body-text" style="text-transform:${transform};${extraLSStyle}"><span class="verse-sup" style="color:var(--gold);${roc}">${verse}</span>${text}</div>`}
    <div class="verse-trans-label">${State.currentTranslation} &nbsp;·&nbsp; ${dRef}</div>
  `;
}

function buildVerseRangeText(book, chapter, startVerse, endVerse) {
  const parts = [];
  for (let v = startVerse; v <= endVerse; v++) {
    const t = getVerseText(book, chapter, v);
    if (t) parts.push({ verse: v, text: t });
  }
  return parts;
}

function buildVerseRangeHTML(book, chapter, startVerse, endVerse, ref) {
  const verses = buildVerseRangeText(book, chapter, startVerse, endVerse);
  if (!verses.length) return null;
  const transform = _getScriptureTransform();
  const sfs = State.settings || {};
  const dRef = _formatRef(ref);
  const refOnly = !!sfs.scriptureShowRefOnly;
  const breakOnNew = !!sfs.scriptureBreakOnNewVerse;
  const sep = breakOnNew ? '<br>' : '  ';
  const roc = _getRefOutlineCss(sfs);
  const body = verses.map(v =>
    `<span class="verse-sup" style="color:var(--gold);${roc}">${v.verse}</span>${v.text}`
  ).join(sep);
  const extraLSStyle = sfs.scriptureAdditionalLineSpacing ? 'line-height:1.8;' : '';
  return `
    <div class="verse-ref-label" style="${roc}">${dRef} &nbsp;·&nbsp; ${State.currentTranslation}</div>
    ${refOnly ? '' : `<div class="verse-body-text" style="text-transform:${transform};${extraLSStyle}">${body}</div>`}
    <div class="verse-trans-label">${State.currentTranslation} &nbsp;·&nbsp; ${dRef}</div>
  `;
}

function presentVerseRange(book, chapter, startVerse, endVerse, ref, options = {}) {
  const html = buildVerseRangeHTML(book, chapter, startVerse, endVerse, ref);
  if (!html) { toast(`⚠ Verses not found: ${ref}`); return; }

  const verses = buildVerseRangeText(book, chapter, startVerse, endVerse);
  const combinedText = verses.map(v => v.text).join('  ');
  const combinedTextWithNums = verses.map(v => `${v.verse} ${v.text}`).join('  ');

  State.previewVerse = { book, chapter, verse: startVerse, endVerse, ref };
  State.liveContentType = 'scripture';

  const parts = (options.parts || [combinedText]).filter(Boolean);
  const partIdx = 0;
  State.versePartIdx = partIdx;
  State.verseParts = parts;
  _updateVersePartIndicator(1, parts.length);

  const display = document.getElementById('previewDisplay');
  const empty   = document.getElementById('previewEmpty');
  display.innerHTML = html;
  display.style.display = 'flex';
  empty.style.display = 'none';
  _renderVersePaginationControls();
  _syncVerseCardActiveState();
  try { highlightVerseRow(startVerse); } catch (_) {}

  if (State.isAutoMode || State.isLive) {
    _updateLiveDisplayRange(book, chapter, startVerse, endVerse, ref, verses);
    if (State.isLive) _syncProjectionRange(book, chapter, startVerse, endVerse, ref, verses);
  }

  if (!options.silent) toast(`👁 Preview: ${ref}`);
  try { _updatePreviewNavArrows(); } catch (_) {}
  if (State.isLive) setTimeout(_syncRemoteLiveState, 200);
}

function _updateLiveDisplayRange(book, chapter, startVerse, endVerse, ref, verses) {
  State._logoActive = false;
  _updateLogoBtnState();
  State.liveVerse = { book, chapter, verse: startVerse, endVerse, ref, rangeVerses: verses };
  const display = document.getElementById('liveDisplay');
  const empty   = document.getElementById('liveEmpty');
  if (!display) return;

  const liveSurface = getLiveTextTarget('scripture');
  const textTarget = liveSurface.host || display;
  if (!liveSurface.overlay) {
    display.style.padding  = '';
    display.style.position = '';
  }

  const t = getActiveThemeData();
  if (!liveSurface.overlay && t && !t?.boxes?.length) _applyLiveThemeBackground(display, t);
  const transform = document.documentElement.style.getPropertyValue('--theme-text-transform') || _getScriptureTransform();

  const sfs = State.settings || {};
  const showVerseNums = (sfs.showVerseNumbers !== false) && (t?.showVerseNum !== false);
  const breakOnNew = !!sfs.scriptureBreakOnNewVerse;
  const refOnly = !!sfs.scriptureShowRefOnly;
  const sep = breakOnNew ? '<br>' : '  ';
  const _rangeRefColor = t?.refColor || t?.accentColor || '#c9a84c';
  const _rangeRoc = _getRefOutlineCss(sfs);
  const combinedText = verses.map(v => {
    const num = showVerseNums ? `<span class="verse-sup" style="color:${_rangeRefColor};${_rangeRoc}">${v.verse}</span>` : '';
    return num + v.text;
  }).join(sep);

  if (t?.boxes?.length) {
    const boxHost = liveSurface.overlay ? textTarget : (() => {
      display.style.cssText = 'display:flex;align-items:center;justify-content:center;padding:0;position:relative;overflow:hidden;background:#000';
      _applyLiveThemeBackground(display, t);
      return getLiveThemeTextHost(display);
    })();
    const plainText = verses.map(v => v.text).join('  ');
    const _scrSfs = State.settings || {};
    renderBoxThemeToElement(boxHost, t, {
      ref, verse: startVerse,
      translation: State.currentTranslation,
      text: plainText,
      showVerseNumbers: showVerseNums,
      scriptureShowReference: _scrSfs.scriptureShowReference !== false && (_scrSfs.scriptureRefPosition || 'top') !== 'hidden',
      scriptureRefPosition: _scrSfs.scriptureRefPosition || 'top',
      scriptureTextTransform: _getScriptureTransform(),
      scriptureAbbreviateBooks: !!_scrSfs.scriptureAbbreviateBooks,
      scriptureShowRefOnly: !!_scrSfs.scriptureShowRefOnly,
      scriptureAdditionalLineSpacing: !!_scrSfs.scriptureAdditionalLineSpacing,
      scriptureBreakOnNewVerse: !!_scrSfs.scriptureBreakOnNewVerse,
      scriptureAutoSize: _scrSfs.scriptureAutoSize || 'resize',
      scriptureNormalizeSize: _scrSfs.scriptureNormalizeSize !== false,
      scriptureRefOutlineStyle: _scrSfs.scriptureRefOutlineStyle || 'none',
      scriptureRefOutlineColor: _scrSfs.scriptureRefOutlineColor || '#000000',
      scriptureRefOutlineWidth: _scrSfs.scriptureRefOutlineWidth ?? 0,
      scriptureRefOutlineOpacity: _scrSfs.scriptureRefOutlineOpacity ?? 100,
      rangeVerses: verses,
      fontOverrides: {
        fontFamily: _scrSfs.scriptureFontFamily, fontSize: _scrSfs.scriptureFontSize,
        fontColor: _scrSfs.scriptureFontColor, fontStyle: _scrSfs.scriptureFontStyle,
        lineSpacing: _scrSfs.scriptureLineSpacing, textAlign: _scrSfs.scriptureTextAlign,
        underline: _scrSfs.scriptureUnderline, fontOpacity: _scrSfs.scriptureFontOpacity,
        outlineStyle: _scrSfs.scriptureOutlineStyle, outlineColor: _scrSfs.scriptureOutlineColor,
        outlineWidth: _scrSfs.scriptureOutlineWidth, outlineOpacity: _scrSfs.scriptureOutlineOpacity,
        shadowEnabled: _scrSfs.scriptureShadowEnabled, shadowColor: _scrSfs.scriptureShadowColor,
        shadowBlur: _scrSfs.scriptureShadowBlur, shadowX: _scrSfs.scriptureShadowX, shadowY: _scrSfs.scriptureShadowY,
        shadowOpacity: _scrSfs.scriptureShadowOpacity,
        marginTop: _scrSfs.scriptureMarginTop, marginBottom: _scrSfs.scriptureMarginBottom,
        marginLeft: _scrSfs.scriptureMarginLeft, marginRight: _scrSfs.scriptureMarginRight,
        verticalAlign: _scrSfs.scriptureVerticalAlign,
      },
      scriptureRefFont: {
        fontFamily:    _scrSfs.scriptureRefFontFamily || 'Cinzel',
        fontSize:      _scrSfs.scriptureRefFontSize || 38,
        fontColor:     _scrSfs.scriptureRefFontColor || '#c9a84c',
        fontStyle:     _scrSfs.scriptureRefFontStyle || 'bold',
        underline:     !!_scrSfs.scriptureRefUnderline,
        uppercase:     _scrSfs.scriptureRefUppercase !== false,
        fontOpacity:   _scrSfs.scriptureRefFontOpacity ?? 90,
        textAlign:     _scrSfs.scriptureRefTextAlign || 'center',
        letterSpacing: _scrSfs.scriptureRefLetterSpacing ?? 0.25,
        shadowEnabled: !!_scrSfs.scriptureRefShadowEnabled,
        shadowColor:   _scrSfs.scriptureRefShadowColor || '#000000',
        shadowBlur:    _scrSfs.scriptureRefShadowBlur ?? 4,
        shadowX:       _scrSfs.scriptureRefShadowX ?? 1,
        shadowY:       _scrSfs.scriptureRefShadowY ?? 1,
        shadowOpacity: _scrSfs.scriptureRefShadowOpacity ?? 60,
      },
    });
    if (empty) empty.style.display = 'none';
    return;
  }

  const sfAlign  = sfs.scriptureTextAlign || 'center';
  const sfVAlign = sfs.scriptureVerticalAlign || 'bottom';
  const ldW = display.clientWidth || 450;
  const ldScale = Math.min(1, ldW / 1920);
  const sfSize   = sfs.scriptureFontSize || 80;
  const scaledFs = Math.round(sfSize * ldScale);
  const fontStyle = sfs.scriptureFontStyle || 'bold';
  const fontWeight = (fontStyle === 'bold' || fontStyle === 'bold-italic') ? 'bold' : 'normal';
  const fontItalic = (fontStyle === 'italic' || fontStyle === 'bold-italic') ? 'italic' : 'normal';
  const ls = sfs.scriptureLineSpacing || 1.4;

  let outline = '';
  const outStyle = sfs.scriptureOutlineStyle || 'none';
  if (outStyle !== 'none') {
    const ow = Math.max(1, (sfs.scriptureOutlineWidth ?? 12) * ldScale);
    const oc = sfs.scriptureOutlineColor || '#000';
    const oo = (sfs.scriptureOutlineOpacity ?? 100) / 100;
    const r = parseInt(oc.slice(1,3),16)||0, g = parseInt(oc.slice(3,5),16)||0, b = parseInt(oc.slice(5,7),16)||0;
    outline = `-webkit-text-stroke:${ow}px rgba(${r},${g},${b},${oo});paint-order:stroke fill;`;
  }

  let shadow = 'text-shadow:none;';
  if (sfs.scriptureShadowEnabled !== false) {
    const sx = ((sfs.scriptureShadowX ?? 2) * ldScale).toFixed(1);
    const sy = ((sfs.scriptureShadowY ?? 2) * ldScale).toFixed(1);
    const sb = ((sfs.scriptureShadowBlur ?? 8) * ldScale).toFixed(1);
    const so = (sfs.scriptureShadowOpacity ?? 80) / 100;
    const sc = sfs.scriptureShadowColor || '#000';
    const sr = parseInt(sc.slice(1,3),16)||0, sg = parseInt(sc.slice(3,5),16)||0, sbb = parseInt(sc.slice(5,7),16)||0;
    shadow = `text-shadow:${sx}px ${sy}px ${sb}px rgba(${sr},${sg},${sbb},${so});`;
  }

  const mT = Math.round((sfs.scriptureMarginTop ?? 22) * ldScale);
  const mB = Math.round((sfs.scriptureMarginBottom ?? 22) * ldScale);
  const mL = Math.round((sfs.scriptureMarginLeft ?? 200) * ldScale);
  const mR = Math.round((sfs.scriptureMarginRight ?? 200) * ldScale);

  let refOutline = '';
  const refOutStyle = sfs.scriptureRefOutlineStyle || 'none';
  if (refOutStyle !== 'none') {
    const row = Math.max(1, (sfs.scriptureRefOutlineWidth ?? 0) * ldScale);
    const roc = sfs.scriptureRefOutlineColor || '#000';
    const roo = (sfs.scriptureRefOutlineOpacity ?? 100) / 100;
    const rr = parseInt(roc.slice(1,3),16)||0, rg = parseInt(roc.slice(3,5),16)||0, rb = parseInt(roc.slice(5,7),16)||0;
    refOutline = `-webkit-text-stroke:${row}px rgba(${rr},${rg},${rb},${roo});paint-order:stroke fill;`;
  }

  const refPos = sfs.scriptureRefPosition || 'top';
  const showRef = sfs.scriptureShowReference !== false && refPos !== 'hidden';
  const dRef = _formatRef(ref);
  const refHtml = showRef ? `<div class="verse-ref-label" style="order:${refPos==='bottom'?'2':'0'};${refOutline}">${dRef} · ${State.currentTranslation}</div>` : '';

  const extraLS = sfs.scriptureAdditionalLineSpacing ? 0.4 : 0;
  const effLS = extraLS ? Math.min(ls + extraLS, 2.4) : ls;
  const vAlignMap = { top: 'flex-start', center: 'center', bottom: 'flex-end' };
  const bodyStyle = `font-size:${scaledFs}px;font-weight:${fontWeight};font-style:${fontItalic};text-align:${sfAlign};line-height:${effLS};text-transform:${transform};max-width:none;${outline}${shadow}`;

  textTarget.innerHTML = `
    ${refPos !== 'bottom' ? refHtml : ''}
    ${refOnly ? '' : `<div class="verse-body-text" style="${bodyStyle};order:1">${combinedText}</div>`}
    ${refPos === 'bottom' ? refHtml : ''}
  `;
  display.style.display = 'flex';
  display.style.padding = `${mT}px ${mR}px ${mB}px ${mL}px`;
  display.style.justifyContent = vAlignMap[sfVAlign] || 'flex-end';
  if (empty) empty.style.display = 'none';
}

function _syncProjectionRange(book, chapter, startVerse, endVerse, ref, verses) {
  if (!window.electronAPI && !State.isProjectionOpen) return;

  const sfs = State.settings || {};
  const showVerseNums = sfs.showVerseNumbers !== false;
  const themeData = getActiveThemeData();

  const combinedTextPlain = verses.map(v => v.text).join('  ');

  const versePayload = {
    book, chapter, verse: startVerse, endVerse, ref,
    text: combinedTextPlain,
    rangeVerses: verses,
    translation: State.currentTranslation,
    theme: State.currentTheme,
    themeData,
    partIdx: 0, partTotal: 1,
    showVerseNumbers: showVerseNums,
    scriptureRefPosition: sfs.scriptureRefPosition || 'top',
    scriptureRefLocation: sfs.scriptureRefLocation || 'each',
    scriptureShowReference: sfs.scriptureShowReference !== false && (sfs.scriptureRefPosition || 'top') !== 'hidden',
    scriptureTextTransform: _getScriptureTransform(),
    scriptureAbbreviateBooks: !!sfs.scriptureAbbreviateBooks,
    scriptureShowRefOnly: !!sfs.scriptureShowRefOnly,
    scriptureAdditionalLineSpacing: !!sfs.scriptureAdditionalLineSpacing,
    scriptureBreakOnNewVerse: !!sfs.scriptureBreakOnNewVerse,
    scriptureAutoSize: sfs.scriptureAutoSize || 'resize',
    scriptureNormalizeSize: sfs.scriptureNormalizeSize !== false,
    scriptureRefOutlineStyle: sfs.scriptureRefOutlineStyle || 'none',
    scriptureRefOutlineColor: sfs.scriptureRefOutlineColor || '#000000',
    scriptureRefOutlineWidth: sfs.scriptureRefOutlineWidth ?? 0,
    scriptureRefOutlineOpacity: sfs.scriptureRefOutlineOpacity ?? 100,
    overlayOptions: getOverlayTextPrefs('scripture'),
    scriptureFont: {
      fontFamily:      sfs.scriptureFontFamily || 'Calibri',
      fontSize:        sfs.scriptureFontSize || 80,
      fontColor:       sfs.scriptureFontColor || '#ffffff',
      fontStyle:       sfs.scriptureFontStyle || 'bold',
      underline:       !!sfs.scriptureUnderline,
      fontOpacity:     sfs.scriptureFontOpacity ?? 100,
      lineSpacing:     sfs.scriptureLineSpacing || 1.4,
      textAlign:       sfs.scriptureTextAlign || '',
      verticalAlign:   sfs.scriptureVerticalAlign || '',
      outlineStyle:    sfs.scriptureOutlineStyle || 'none',
      outlineColor:    sfs.scriptureOutlineColor || '#000000',
      outlineJoin:     sfs.scriptureOutlineJoin || 'round',
      outlineWidth:    sfs.scriptureOutlineWidth ?? 12,
      outlineOpacity:  sfs.scriptureOutlineOpacity ?? 100,
      shadowEnabled:   sfs.scriptureShadowEnabled !== false,
      shadowColor:     sfs.scriptureShadowColor || '#000000',
      shadowBlur:      sfs.scriptureShadowBlur ?? 8,
      shadowX:         sfs.scriptureShadowX ?? 2,
      shadowY:         sfs.scriptureShadowY ?? 2,
      shadowOpacity:   sfs.scriptureShadowOpacity ?? 80,
      marginTop:       sfs.scriptureMarginTop ?? 22,
      marginBottom:    sfs.scriptureMarginBottom ?? 22,
      marginLeft:      sfs.scriptureMarginLeft ?? 200,
      marginRight:     sfs.scriptureMarginRight ?? 200,
    },
    scriptureRefFont: {
      fontFamily:      sfs.scriptureRefFontFamily || 'Cinzel',
      fontSize:        sfs.scriptureRefFontSize || 38,
      fontColor:       sfs.scriptureRefFontColor || '#c9a84c',
      fontStyle:       sfs.scriptureRefFontStyle || 'bold',
      underline:       !!sfs.scriptureRefUnderline,
      uppercase:       sfs.scriptureRefUppercase !== false,
      fontOpacity:     sfs.scriptureRefFontOpacity ?? 90,
      textAlign:       sfs.scriptureRefTextAlign || 'center',
      letterSpacing:   sfs.scriptureRefLetterSpacing ?? 0.25,
      shadowEnabled:   !!sfs.scriptureRefShadowEnabled,
      shadowColor:     sfs.scriptureRefShadowColor || '#000000',
      shadowBlur:      sfs.scriptureRefShadowBlur ?? 4,
      shadowX:         sfs.scriptureRefShadowX ?? 1,
      shadowY:         sfs.scriptureRefShadowY ?? 1,
      shadowOpacity:   sfs.scriptureRefShadowOpacity ?? 60,
    },
  };
  if (window.electronAPI) {
    window.electronAPI.projectVerse(versePayload);
  } else {
    _webPostRenderState('scripture', versePayload);
  }
}

function presentVerse(book, chapter, verse, ref, options = {}) {
  const html = buildVerseHTML(book, chapter, verse, ref);
  if (!html) { toast(`⚠ Verse not found: ${ref}`); return; }

  State.previewVerse = { book, chapter, verse, ref };

  // Check if verse needs splitting for the current theme's text box
  const parts = (options.parts || _splitVerseForTheme(book, chapter, verse, ref)).filter(Boolean);
  if (!parts.length) { toast(`⚠ Could not load verse text: ${ref}`); return; }  // FIX: guard empty parts
  const partIdx = Math.max(0, Math.min(options.partIdx || 0, Math.max(0, parts.length - 1)));
  State.versePartIdx  = partIdx;
  State.verseParts    = parts;
  _updateVersePartIndicator(partIdx + 1, parts.length);

  // Update Preview — show selected part
  const display = document.getElementById('previewDisplay');
  const empty   = document.getElementById('previewEmpty');
  display.innerHTML = html;
  display.style.display = 'flex';
  empty.style.display = 'none';
  if (parts.length > 1) {
    _renderVersePartInPreview(parts[partIdx], ref, `${partIdx + 1}/${parts.length}`);
  }
  _renderVersePaginationControls();
  _syncVerseCardActiveState();
  try { highlightVerseRow(verse); } catch (_) {}

  if (State.isAutoMode || State.isLive) {
    updateLiveDisplay(book, chapter, verse, ref);
    if (State.isLive) syncProjection();
  }

  if (!options.silent) toast(`👁 Preview: ${ref}${parts.length > 1 ? ` (${partIdx + 1}/${parts.length})` : ''}`);
  // Update Program Preview PREV/NEXT overlay arrow state
  try { _updatePreviewNavArrows(); } catch (_) {}
  if (State.isLive) setTimeout(_syncRemoteLiveState, 200);
}

// Show part indicator badge on live canvas when verse is split
function _updateVersePartIndicator(part, total) {
  // Add/update a small "Part 1/2" badge on the live canvas
  let badge = document.getElementById('versePartBadge');
  if (!badge) {
    badge = document.createElement('div');
    badge.id = 'versePartBadge';
    badge.style.cssText = 'position:absolute;top:6px;right:8px;font-size:10px;font-weight:700;' +
      'background:rgba(201,168,76,.85);color:#000;padding:2px 7px;border-radius:4px;z-index:10;pointer-events:none';
    const lc = document.getElementById('liveCanvas');
    if (lc) lc.appendChild(badge);
  }
  if (total > 1) {
    badge.textContent = `${part}/${total}`;
    badge.style.display = 'block';
  } else {
    badge.style.display = 'none';
  }
}

function _syncVerseCardActiveState() {
  const cards = document.querySelectorAll('#bibleSlideGrid .slide-card[data-verse]');
  cards.forEach(c => {
    const isActive =
      String(c.dataset.verse || '') === String(State.previewVerse?.verse || '') &&
      String(c.dataset.partIdx || '0') === String(State.versePartIdx || 0);
    c.classList.toggle('active', isActive);
  });
}

function _renderVersePaginationControls() {
  const display = document.getElementById('previewDisplay');
  if (!display) return;
  let wrap = document.getElementById('versePreviewPager');
  if (!wrap) {
    wrap = document.createElement('div');
    wrap.id = 'versePreviewPager';
    wrap.style.cssText = 'position:absolute;left:10px;right:10px;bottom:10px;display:none;align-items:center;justify-content:space-between;gap:8px;z-index:8;pointer-events:none';
    wrap.innerHTML = `
      <button type="button" data-dir="-1" style="pointer-events:auto;border:1px solid rgba(201,168,76,.35);background:rgba(0,0,0,.55);color:#f3e7c0;border-radius:8px;padding:6px 10px;font:600 11px/1 inherit;cursor:pointer">◀ Prev part</button>
      <div data-role="label" style="pointer-events:none;min-width:90px;text-align:center;border:1px solid rgba(201,168,76,.25);background:rgba(0,0,0,.45);color:#f3e7c0;border-radius:999px;padding:5px 12px;font:700 11px/1 inherit;letter-spacing:.04em"></div>
      <button type="button" data-dir="1" style="pointer-events:auto;border:1px solid rgba(201,168,76,.35);background:rgba(0,0,0,.55);color:#f3e7c0;border-radius:8px;padding:6px 10px;font:600 11px/1 inherit;cursor:pointer">Next part ▶</button>`;
    display.appendChild(wrap);
    wrap.querySelectorAll('button[data-dir]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        stepVersePart(Number(btn.dataset.dir || '0'), { syncLive: !!State.liveVerse });
      });
    });
  }
  const total = State.verseParts?.length || 1;
  const idx   = Math.max(0, Math.min(State.versePartIdx || 0, total - 1));
  const label = wrap.querySelector('[data-role="label"]');
  const prev  = wrap.querySelector('button[data-dir="-1"]');
  const next  = wrap.querySelector('button[data-dir="1"]');
  if (total > 1) {
    wrap.style.display = 'flex';
    if (label) label.textContent = `Part ${idx + 1} of ${total}`;
    if (prev) prev.disabled = idx <= 0;
    if (next) next.disabled = idx >= total - 1;
    if (prev) prev.style.opacity = idx <= 0 ? '.45' : '1';
    if (next) next.style.opacity = idx >= total - 1 ? '.45' : '1';
  } else {
    wrap.style.display = 'none';
  }
}

function _applyCurrentVersePartToPreview() {
  if (!State.previewVerse) return false;
  const { book, chapter, verse, ref } = State.previewVerse;
  const html = buildVerseHTML(book, chapter, verse, ref);
  const display = document.getElementById('previewDisplay');
  const empty = document.getElementById('previewEmpty');
  if (!display || !html) return false;
  display.innerHTML = html;
  display.style.display = 'flex';
  if (empty) empty.style.display = 'none';
  const parts = State.verseParts || _splitVerseForTheme(book, chapter, verse, ref);
  State.verseParts = parts;
  const idx = Math.max(0, Math.min(State.versePartIdx || 0, Math.max(0, parts.length - 1)));
  State.versePartIdx = idx;
  if (parts.length > 1) {
    _renderVersePartInPreview(parts[idx], ref, `${idx + 1}/${parts.length}`);
  }
  _renderVersePaginationControls();
  _syncVerseCardActiveState();
  return true;
}

function stepVersePart(delta, { syncLive = false } = {}) {
  if (!State.previewVerse) return false;
  const parts = State.verseParts || _splitVerseForTheme(
    State.previewVerse.book,
    State.previewVerse.chapter,
    State.previewVerse.verse,
    State.previewVerse.ref
  );
  if (!parts || parts.length <= 1) return false;
  const nextIdx = (State.versePartIdx || 0) + delta;
  if (nextIdx < 0 || nextIdx >= parts.length) return false;
  State.verseParts = parts;
  State.versePartIdx = nextIdx;
  _updateVersePartIndicator(nextIdx + 1, parts.length);
  _applyCurrentVersePartToPreview();
  if (syncLive && State.liveVerse) {
    updateLiveDisplay(State.liveVerse.book, State.liveVerse.chapter, State.liveVerse.verse, State.liveVerse.ref);
    if (State.isLive) syncProjection();
  }
  toast(`📖 Part ${nextIdx + 1}/${parts.length}`);
  return true;
}


function _escapeHtmlBasic(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function _getScriptureTransform() {
  const s = State.settings || {};
  if (s.scriptureCapitalizeAll) return 'uppercase';
  if (s.scriptureCapitalizeFirst) return 'capitalize';
  return s.scriptureTextTransform || 'none';
}

function _getSectionLabelStyle(label) {
  if (!label) return '';
  const labels = State.settings?.songSectionLabels;
  if (!Array.isArray(labels) || !labels.length) return '';
  const norm = label.replace(/\s*\d+$/, '').toUpperCase().trim();
  const match = labels.find(l => norm === l.name || norm.startsWith(l.name));
  if (!match) return '';
  return `background:${match.bgColor};color:${match.textColor};padding:1px 6px;border-radius:3px;font-size:9px`;
}

function _getSongTransform(themeData = null, sec = null) {
  const explicit = State.settings?.songTextTransform;
  if (explicit && explicit !== 'theme') return explicit;
  const perSlide = sec?.textTransform;
  if (perSlide && perSlide !== 'inherit') return perSlide;
  const themeTransform = themeData?.textTransform || 'none';
  if (themeData && themeData.builtIn === false && !themeData.forceSongThemeTransform) return 'none';
  return themeTransform;
}

function _sanitizeSongHtml(html = '') {
  return String(html || '')
    .replace(/<(?!\/?(?:b|strong|i|em|u|span|div|p|br)\b)[^>]*>/gi, '')
    .replace(/ on\w+="[^"]*"/gi, '')
    .replace(/ on\w+='[^']*'/gi, '');
}

function _fitBoxFontSize(box = {}, content = '') {
  const raw = String(content || '').trim();
  const base = Number(box.fontSize || 52);
  if (!raw) return base;
  const width = Math.max(120, Number(box.w || 1680) - 32);
  const height = Math.max(80, Number(box.h || 680) - 32);
  const lineSpacing = Number(box.lineSpacing || 1.4);
  const isBold = box.bold || (box.fontWeight && Number(box.fontWeight) >= 700);
  const charRatio = isBold ? 0.62 : 0.58;
  let size = base;
  while (size > 16) {
    const approxCharWidth = Math.max(5, size * charRatio);
    const charsPerLine = Math.max(6, Math.floor(width / approxCharWidth));
    const wrappedLines = raw.split(/\n+/).reduce((sum, line) => sum + Math.max(1, Math.ceil((line.length || 1) / charsPerLine)), 0);
    const maxLines = Math.max(1, Math.floor(height / (size * lineSpacing)));
    if (wrappedLines <= maxLines) return size;
    size -= 2;
  }
  return Math.max(16, size);
}

function renderBoxThemeToElement(display, themeData, payload = {}) {
  if (!display || !themeData?.boxes?.length) return false;
  const forcedTextTransform = payload.songTextTransform || payload.scriptureTextTransform || null;
  const cw = display.clientWidth || 400;
  const ch = display.clientHeight || 225;
  const sc = Math.min(cw / 1920, ch / 1080);
  const globalTransform = forcedTextTransform || themeData.textTransform || 'none';
  const uf = payload.fontOverrides || {};

  const normBox = (box = {}) => ({
    ...box,
    textTransform: box.textTransform || globalTransform || 'none',
    fontWeight: Number(box.fontWeight || (box.bold ? 700 : 400)),
    lineSpacing: Number(box.lineSpacing || 1.4),
    letterSpacing: Number(box.letterSpacing || 0),
    shadow: box.shadow !== false,
    shadowColor: box.shadowColor || '#000000',
    shadowBlur: Number(box.shadowBlur ?? 8),
    shadowOffsetX: Number(box.shadowOffsetX ?? 0),
    shadowOffsetY: Number(box.shadowOffsetY ?? 2)
  });

  const rgbaFromHex = (hex, alpha) => {
    const raw = String(hex || '#000000').replace('#','');
    const full = raw.length === 3 ? raw.split('').map(x => x + x).join('') : raw.padEnd(6, '0').slice(0,6);
    const r = parseInt(full.slice(0,2),16) || 0;
    const g = parseInt(full.slice(2,4),16) || 0;
    const b = parseInt(full.slice(4,6),16) || 0;
    return `rgba(${r},${g},${b},${alpha})`;
  };

  const _buildLiveBoxFontOverrides = (sf) => {
    if (!sf || typeof sf !== 'object') return '';
    const parts = [];
    if (sf.fontFamily) {
      const fam = sf.fontFamily === 'Cinzel' ? 'Crimson Pro' : sf.fontFamily;
      parts.push(`font-family:'${fam}',serif`);
    }
    if (sf.fontColor) parts.push(`color:${sf.fontColor}`);
    const fst = sf.fontStyle || '';
    if (fst) {
      parts.push(`font-weight:${(fst==='bold'||fst==='bold-italic')?'700':'400'}`);
      parts.push(`font-style:${(fst==='italic'||fst==='bold-italic')?'italic':'normal'}`);
    }
    if (sf.textAlign) parts.push(`text-align:${sf.textAlign}`);
    if (sf.underline) parts.push('text-decoration:underline');
    if (sf.fontOpacity != null) parts.push(`opacity:${sf.fontOpacity / 100}`);
    const outStyle = sf.outlineStyle || 'none';
    if (outStyle && outStyle !== 'none') {
      const oo = (sf.outlineOpacity ?? 100) / 100;
      const oc = sf.outlineColor || '#000';
      const r2=parseInt(oc.slice(1,3),16)||0, g2=parseInt(oc.slice(3,5),16)||0, b2=parseInt(oc.slice(5,7),16)||0;
      parts.push(`-webkit-text-stroke:${Math.max(1, (sf.outlineWidth||11) * sc)}px rgba(${r2},${g2},${b2},${oo})`);
      parts.push('paint-order:stroke fill');
    }
    if (sf.shadowEnabled === true) {
      const sx=(sf.shadowX??2)*sc, sy=(sf.shadowY??2)*sc, sb=(sf.shadowBlur??8)*sc;
      const so=(sf.shadowOpacity??80)/100;
      const shc=sf.shadowColor||'#000000';
      const r3=parseInt(shc.slice(1,3),16)||0, g3=parseInt(shc.slice(3,5),16)||0, b3=parseInt(shc.slice(5,7),16)||0;
      parts.push(`text-shadow:${sx.toFixed(1)}px ${sy.toFixed(1)}px ${sb.toFixed(1)}px rgba(${r3},${g3},${b3},${so})`);
    } else if (sf.shadowEnabled === false) {
      parts.push('text-shadow:none');
    }
    return parts.join(';');
  };

  let html = '';
  themeData.boxes.forEach(rawBox => {
    const box = normBox(rawBox);
    const bx = (box.x * sc) + 'px', by = (box.y * sc) + 'px';
    const bw = (box.w * sc) + 'px', bh = (box.h * sc) + 'px';
    const role = (box.role || '').toLowerCase();
    const isMain = (role === 'main');
    let content = box.text || '';
    let contentHTML = '';
    let useHTML = false;
    let _inlineRefHTML = '';
    if (role === 'main') {
      const isSong = Array.isArray(payload.lines);
      if (!isSong && payload.scriptureShowRefOnly) {
        content = '';
      } else {
        const showNums = !isSong && (payload.showVerseNumbers !== false);
        const breakOnNew = !!payload.scriptureBreakOnNewVerse;
        const _boxRefBox = (themeData?.boxes || []).find(b => (b.role||'').toLowerCase() === 'ref');
        const _boxNumColor = box.refColor || _boxRefBox?.color || themeData?.refColor || themeData?.accentColor || '#c9a84c';
        const esc = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        const safeNum = n => String(parseInt(n, 10) || 0);

        const _refPos = payload.scriptureRefPosition || 'top';
        let _inlineRefHTML = '';
        if (!isSong && !_boxRefBox && payload.scriptureShowReference !== false) {
          let dRef = payload.ref;
          if (payload.scriptureAbbreviateBooks && dRef) dRef = _abbreviateRef(dRef);
          if (dRef) {
            const refLabel = dRef + (payload.translation ? ' (' + payload.translation + ')' : '');
            // Read ref font from Preferences (scriptureRefFont) — not from box theme properties
            const _rf = payload.scriptureRefFont || {};
            const rFF = _rf.fontFamily || 'Cinzel';
            const rFS = _rf.fontSize || 38;
            const rFC = _rf.fontColor || _boxNumColor;
            const _rfBold = _rf.fontStyle === 'bold' || _rf.fontStyle === 'bold-italic';
            const rFW = _rfBold ? 700 : 500;
            const rLS = 1.4;
            const rfOpacity = (_rf.fontOpacity ?? 90) / 100;
            const rfAlign   = _rf.textAlign || box.align || 'center';
            const rfLS_em   = _rf.letterSpacing ?? 0.25;
            const rfUpper   = _rf.uppercase !== false ? 'uppercase' : 'none';
            const _refFsSc  = Math.max(7, rFS * sc) + 'px';
            const _refMargin = _refPos === 'bottom' ? `margin-top:${Math.max(4, 12*sc)}px;margin-bottom:0` : `margin-bottom:${Math.max(4, 12*sc)}px`;
            _inlineRefHTML = `<div style="font-family:'${rFF}',serif;font-size:${_refFsSc};color:${rFC};font-weight:${rFW};line-height:${rLS};opacity:${rfOpacity};letter-spacing:${rfLS_em}em;text-transform:${rfUpper};${_refMargin};text-align:${rfAlign};flex-shrink:0">${esc(refLabel)}</div>`;
          }
        }

        const _combineRefAndVerse = (refH, verseH) => {
          if (!refH) return verseH || '';
          const wrV = verseH ? `<div style="min-height:0">${verseH}</div>` : '';
          if (_refPos === 'bottom') return wrV + refH;
          return refH + wrV;
        };

        if (!isSong && payload.rangeVerses && payload.rangeVerses.length > 0) {
          const sep = breakOnNew ? '\n' : '  ';
          content = payload.rangeVerses.map(v => {
            const num = showNums ? v.verse + ' ' : '';
            return num + v.text;
          }).join(sep);
          const htmlSep = breakOnNew ? '<br>' : '  ';
          const _boxRoc = _getRefOutlineCss(payload);
          if (showNums) {
            contentHTML = _combineRefAndVerse(_inlineRefHTML, payload.rangeVerses.map(v =>
              `<span class="verse-sup" style="color:${_boxNumColor};${_boxRoc}">${safeNum(v.verse)}</span>${esc(v.text)}`
            ).join(htmlSep));
          } else {
            contentHTML = _combineRefAndVerse(_inlineRefHTML, payload.rangeVerses.map(v => esc(v.text)).join(htmlSep));
          }
          useHTML = !!contentHTML;
        } else {
          content = isSong ? payload.lines.join('\n') : (payload.text ?? box.text ?? '');
          if (!isSong && _inlineRefHTML) {
            const _boxSingleRoc = _getRefOutlineCss(payload);
            if (showNums && payload.verse) {
              contentHTML = _combineRefAndVerse(_inlineRefHTML, `<span class="verse-sup" style="color:${_boxNumColor};${_boxSingleRoc}">${safeNum(payload.verse)}</span>${esc(payload.text ?? box.text ?? '')}`);
            } else {
              contentHTML = _combineRefAndVerse(_inlineRefHTML, esc(content));
            }
            useHTML = true;
          } else if (showNums && payload.verse) {
            content = payload.verse + ' ' + content;
            const _boxSingleRoc = _getRefOutlineCss(payload);
            contentHTML = `<span class="verse-sup" style="color:${_boxNumColor};${_boxSingleRoc}">${safeNum(payload.verse)}</span>${esc(payload.text ?? box.text ?? '')}`;
            useHTML = true;
          }
        }
      }
    } else if (role === 'ref') {
      const isSong = Array.isArray(payload.lines);
      if (isSong) {
        if (payload.showSongMetadata) {
          content = payload.title ? `${payload.title}${payload.sectionLabel ? ' · ' + payload.sectionLabel : ''}` : (box.text || '');
        } else {
          content = '';
        }
      } else if (payload.scriptureShowReference === false) {
        content = '';
      } else {
        let dRef = payload.ref;
        if (payload.scriptureAbbreviateBooks && dRef) dRef = _abbreviateRef(dRef);
        content = dRef ? `${dRef}${payload.translation ? ' · ' + payload.translation : ''}` : (box.text || '');
      }
    } else if (role === 'title') {
      const isSong = Array.isArray(payload.lines);
      content = (isSong && !payload.showSongMetadata) ? '' : (payload.title || payload.name || payload.ref || box.text || '');
    } else if (role === 'subtitle') {
      const isSong = Array.isArray(payload.lines);
      content = (isSong && !payload.showSongMetadata) ? '' : (payload.sectionLabel || payload.subtitle || box.text || '');
    }
    const bgF = box.bgOpacity > 0 && box.bgFill ? `background:${rgbaFromHex(box.bgFill, box.bgOpacity/100)};` : '';
    const bord = box.borderW > 0 ? `border:${Math.max(1, box.borderW*sc)}px solid ${box.borderColor||'#fff'};` : '';
    const transform = forcedTextTransform || box.textTransform || globalTransform;
    const shadowStyle = box.shadow === false ? 'none' : `${(box.shadowOffsetX * sc).toFixed(1)}px ${(box.shadowOffsetY * sc).toFixed(1)}px ${(Math.max(0, box.shadowBlur) * sc).toFixed(1)}px ${box.shadowColor || '#000000'}`;

    const mT = isMain && uf.marginTop != null ? Math.max(1, uf.marginTop * sc) : Math.max(6, 8*sc);
    const mB = isMain && uf.marginBottom != null ? Math.max(1, uf.marginBottom * sc) : Math.max(6, 8*sc);
    const mL = isMain && uf.marginLeft != null ? Math.max(1, uf.marginLeft * sc) : Math.max(6, 8*sc);
    const mR = isMain && uf.marginRight != null ? Math.max(1, uf.marginRight * sc) : Math.max(6, 8*sc);

    const _isSongFit = Array.isArray(payload.lines);
    const _hasInlineRefFit = isMain && !_isSongFit
      && !(themeData?.boxes || []).some(b => (b.role||'').toLowerCase() === 'ref')
      && payload.scriptureShowReference !== false && payload.ref;
    const _inlineRefHFit = _hasInlineRefFit ? ((payload.scriptureRefFont?.fontSize || 38) * 2.0 + 30) : 0;

    let fittedFontSize = box.fontSize || 52;
    if (isMain) {
      const fitBox = Object.assign({}, box);
      fitBox.w = Math.max(100, box.w - ((uf.marginLeft ?? 0) + (uf.marginRight ?? 0)));
      fitBox.h = Math.max(60, box.h - ((uf.marginTop ?? 0) + (uf.marginBottom ?? 0)) - _inlineRefHFit);
      if (uf.fontSize) fitBox.fontSize = uf.fontSize;
      if (uf.lineSpacing) fitBox.lineSpacing = uf.lineSpacing;
      if (uf.fontStyle) {
        const ufBold = uf.fontStyle === 'bold' || uf.fontStyle === 'bold-italic';
        fitBox.bold = ufBold;
        fitBox.fontWeight = ufBold ? 700 : 400;
      }
      const autoSize = payload.scriptureAutoSize || payload.autoSize || 'resize';
      const normalizeSize = payload.scriptureNormalizeSize !== undefined ? payload.scriptureNormalizeSize : (payload.normalizeSize !== false);
      if (autoSize === 'none') {
        fittedFontSize = fitBox.fontSize || box.fontSize || 52;
      } else {
        fittedFontSize = _fitBoxFontSize(fitBox, content);
        if (normalizeSize) {
          const maxNorm = Math.round((fitBox.fontSize || box.fontSize || 52) * 0.85);
          fittedFontSize = Math.min(fittedFontSize, maxNorm);
        }
      }
    }
    const fs = Math.max(7, fittedFontSize * sc) + 'px';
    const vAlign = isMain && uf.verticalAlign
      ? (uf.verticalAlign === 'top' ? 'flex-start' : uf.verticalAlign === 'bottom' ? 'flex-end' : 'center')
      : (box.valign==='top'?'flex-start':box.valign==='bottom'?'flex-end':'center');

    const baseLS = box.lineSpacing || 1.4;
    const effLS = (isMain && payload.scriptureAdditionalLineSpacing) ? Math.min(baseLS + 0.4, 2.4) : baseLS;
    const baseDivCss = `width:100%;white-space:pre-wrap;text-align:${box.align||'center'};
          font-family:'${box.fontFamily||'DM Sans'}',serif;font-size:${fs};color:${box.color||'#fff'};
          font-weight:${box.fontWeight || (box.bold?700:400)};font-style:${box.italic?'italic':'normal'};
          line-height:${effLS};letter-spacing:${(box.letterSpacing || 0) * sc}px;
          text-transform:${transform};text-shadow:${shadowStyle}`;
    const fontOvCss = isMain ? _buildLiveBoxFontOverrides(uf) : '';
    let finalDivCss = fontOvCss ? baseDivCss + ';' + fontOvCss : baseDivCss;
    if (role === 'ref') {
      const _roStyle = payload.scriptureRefOutlineStyle || 'none';
      if (_roStyle !== 'none') {
        const _roW = Math.max(1, (payload.scriptureRefOutlineWidth ?? 0) * sc);
        const _roC = payload.scriptureRefOutlineColor || '#000';
        const _roO = (payload.scriptureRefOutlineOpacity ?? 100) / 100;
        const _rR = parseInt(_roC.slice(1,3),16)||0, _rG = parseInt(_roC.slice(3,5),16)||0, _rB = parseInt(_roC.slice(5,7),16)||0;
        finalDivCss += `;-webkit-text-stroke:${_roW}px rgba(${_rR},${_rG},${_rB},${_roO});paint-order:stroke fill`;
      }
    }

    const _innerFlexCss = (useHTML && _inlineRefHTML) ? ';display:flex;flex-direction:column' : '';
    html += `<div style="position:absolute;left:${bx};top:${by};width:${bw};height:${bh};
      display:flex;flex-direction:column;align-items:${box.align==='left'?'flex-start':box.align==='right'?'flex-end':'center'};
      justify-content:${vAlign};overflow:hidden;
      padding:${mT.toFixed(1)}px ${mR.toFixed(1)}px ${mB.toFixed(1)}px ${mL.toFixed(1)}px;box-sizing:border-box;${bgF}${bord}
      ${box.borderRadius ? `border-radius:${box.borderRadius*sc}px;` : ''}">
        <div style="${finalDivCss}${_innerFlexCss}">${useHTML ? contentHTML : String(content || '')
            .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
            .replace(/\n/g,'<br>')}</div>
      </div>`;
  });

  const isSong = Array.isArray(payload.lines);
  if (isSong && payload.showSongMetadata) {
    const hasMetaBox = themeData.boxes.some(b => {
      const r = (b.role || '').toLowerCase();
      return r === 'ref' || r === 'title' || r === 'subtitle';
    });
    if (!hasMetaBox && (payload.title || payload.sectionLabel)) {
      const mainBox = themeData.boxes.find(b => (b.role||'').toLowerCase() === 'main') || themeData.boxes[0];
      const metaX = (mainBox?.x ?? 120) * sc;
      const metaW = (mainBox?.w ?? 1680) * sc;
      const metaTop = Math.max(0, (mainBox?.y ?? 100) * sc - 28 * sc);
      const accent = themeData.accentColor || '#c9a84c';
      const metaFs = Math.max(8, 13 * sc);
      const label = payload.title + (payload.sectionLabel ? ' \u00B7 ' + payload.sectionLabel : '');
      const esc = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      html += `<div style="position:absolute;left:${metaX}px;top:${metaTop}px;width:${metaW}px;
        text-align:center;font-size:${metaFs}px;color:${accent};letter-spacing:0.18em;
        text-transform:uppercase;font-family:sans-serif;opacity:0.8;z-index:2">${esc(label)}</div>`;
    }
  }

  display.innerHTML = html;
  return true;
}


function _normalizeThemeMediaUrl(src) {
  const raw = String(src || '').trim();
  if (!raw) return '';
  if (/^(file|https?|blob|data|media):/i.test(raw)) return raw;
  if (/^[A-Za-z]:\\/.test(raw) || raw.startsWith('\\')) return 'file:///' + raw.replace(/\\/g,'/');
  return raw;
}
function _applyLiveThemeBackground(display, themeData) {
  if (!display) return;
  let bg = display.querySelector('.live-theme-bg');
  if (!bg) {
    bg = document.createElement('div');
    bg.className = 'live-theme-bg';
    bg.style.cssText = 'position:absolute;inset:0;overflow:hidden;z-index:0;pointer-events:none;background:#000';
    display.prepend(bg);
  }
  let veil = display.querySelector('.live-theme-veil');
  if (!veil) {
    veil = document.createElement('div');
    veil.className = 'live-theme-veil';
    veil.style.cssText = 'position:absolute;inset:0;z-index:1;pointer-events:none;background:transparent';
    display.appendChild(veil);
  }
  const dim = Math.max(0, Math.min(1, Number(themeData?.bgOverlay || 0) / 100));
  veil.style.background = `rgba(0,0,0,${dim})`;
  if (!themeData) { bg.style.background = '#000'; return; }
  if (themeData.bgType === 'video' && themeData.bgVideo) {
    bg.style.background = '#000';
    let vid = display.querySelector('.live-theme-video');
    if (!vid) {
      vid = document.createElement('video');
      vid.className = 'live-theme-video';
      vid.muted = true; vid.defaultMuted = true; vid.loop = true; vid.autoplay = true; vid.playsInline = true; vid.preload = 'auto';
      vid.setAttribute('muted','');
      vid.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;object-fit:cover;display:block;background:#000';
      bg.appendChild(vid);
    } else { bg.appendChild(vid); }
    const url = _normalizeThemeMediaUrl(themeData.bgVideo);
    if (vid.dataset.src !== url) { vid.pause?.(); vid.src = url; vid.dataset.src = url; vid.dataset.mediaUrl = url; vid.load(); }
    vid.style.display = 'block';
    vid.play().catch(()=>{});
    return;
  }
  const oldVid = display.querySelector('.live-theme-video'); if (oldVid) { try{ oldVid.pause(); }catch(_){} oldVid.remove(); }
  if (themeData.bgType === 'image' && themeData.bgImage) {
    bg.style.background = `url('${_normalizeThemeMediaUrl(themeData.bgImage)}') center/cover no-repeat`;
    bg.style.backgroundRepeat = 'no-repeat';
  } else if (themeData.bgType === 'solid') {
    bg.style.background = themeData.bgColor1 || '#0a0a1e';
  } else if (themeData.bgType === 'linear') {
    bg.style.background = `linear-gradient(135deg,${themeData.bgColor1||'#0a0a1e'},${themeData.bgColor2||'#000'})`;
  } else {
    bg.style.background = `radial-gradient(ellipse at 35% 45%,${themeData?.bgColor1||'#0a0a1e'} 0%,${themeData?.bgColor2||'#000'} 70%,#000 100%)`;
  }
}

function renderSongLiveDisplay(song, sec, lines, themeData, mode = 'present') {
  const liveDisplay = document.getElementById('liveDisplay');
  const liveEmpty   = document.getElementById('liveEmpty');
  if (!liveDisplay) return;
  const liveSurface = mode === 'present' ? getLiveTextTarget('song') : { host: liveDisplay, overlay: false };
  if (mode !== 'present') {
    liveDisplay.querySelectorAll('video,audio').forEach(m => { try { m.pause(); } catch(e) {} });
    liveDisplay.innerHTML = '';
    liveDisplay.style.display = 'flex';
    if (liveEmpty) liveEmpty.style.display = 'none';
  }
  const textTarget = liveSurface.host || liveDisplay;
  liveDisplay.style.textTransform = 'none';
  textTarget.style.textTransform = 'none';
  if (themeData?.boxes?.length) {
    const boxHost = liveSurface.overlay ? textTarget : (() => {
      liveDisplay.style.cssText = 'display:flex;align-items:center;justify-content:center;padding:0;position:relative;overflow:hidden;background:#000';
      _applyLiveThemeBackground(liveDisplay, themeData);
      return getLiveThemeTextHost(liveDisplay);
    })();
    const _sfs = State.settings || {};
    renderBoxThemeToElement(boxHost, themeData, {
      title: song.title,
      name: song.title,
      sectionLabel: sec.label || '',
      subtitle: sec.label || '',
      lines,
      text: lines.join('\n'),
      showSongMetadata: State.settings?.showSongMetadata === true,
      songTextTransform: _getSongTransform(themeData, sec),
      overlayOptions: getOverlayTextPrefs('song'),
      fontOverrides: {
        fontFamily: _sfs.songFontFamily, fontSize: _sfs.songFontSize,
        fontColor: _sfs.songFontColor, fontStyle: _sfs.songFontStyle,
        lineSpacing: _sfs.songLineSpacing, textAlign: _sfs.songTextAlign,
        underline: _sfs.songUnderline, fontOpacity: _sfs.songFontOpacity,
        outlineStyle: _sfs.songOutlineStyle, outlineColor: _sfs.songOutlineColor,
        outlineWidth: _sfs.songOutlineWidth, outlineOpacity: _sfs.songOutlineOpacity,
        shadowEnabled: _sfs.songShadowEnabled, shadowColor: _sfs.songShadowColor,
        shadowBlur: _sfs.songShadowBlur, shadowX: _sfs.songShadowX, shadowY: _sfs.songShadowY,
        shadowOpacity: _sfs.songShadowOpacity,
        marginTop: _sfs.songMarginTop, marginBottom: _sfs.songMarginBottom,
        marginLeft: _sfs.songMarginLeft, marginRight: _sfs.songMarginRight,
        verticalAlign: _sfs.songVerticalAlign,
      },
    });
    if (liveEmpty) liveEmpty.style.display = 'none';
    return;
  }

  if (mode === 'preview') {
    const songTransform = _getSongTransform(themeData, sec);
    const richHtml = sec?.richHtml ? _sanitizeSongHtml(sec.richHtml) : '';
    textTarget.innerHTML = richHtml
      ? `<div class="verse-body-text song-lyric-block" style="font-size:clamp(14px,1.8vw,26px);line-height:1.75;text-align:center;text-transform:${songTransform}">${richHtml}</div>`
      : `<div class="verse-body-text song-lyric-block" style="font-size:clamp(14px,1.8vw,26px);line-height:1.75;text-align:center;text-transform:${songTransform}">${lines.map(l => _escapeHtmlBasic(l)).join('<br>')}</div>`;
    textTarget.querySelectorAll('.song-lyric-block, .song-lyric-block *, div[style], span[style], p[style], br').forEach(el => { try { el.style.textTransform = songTransform; } catch(_) {} });
    liveDisplay.style.display = 'flex';
    if (liveEmpty) liveEmpty.style.display = 'none';
    return;
  }

  const sfs = State.settings || {};
  const rawFf = sfs.songFontFamily || themeData.fontFamily || 'Crimson Pro';
  const ff  = (rawFf === 'Cinzel') ? 'Crimson Pro' : rawFf;
  const fs  = sfs.songFontSize     || themeData.fontSize   || 52;
  const ls  = sfs.songLineSpacing  || themeData.lineSpacing || 1.4;
  const col = sfs.songFontColor    || themeData.textColor  || '#fff';
  const fontStyle = sfs.songFontStyle || 'bold';
  const fontWeight = (fontStyle === 'bold' || fontStyle === 'bold-italic') ? 'bold' : 'normal';
  const fontItalic = (fontStyle === 'italic' || fontStyle === 'bold-italic') ? 'italic' : 'normal';
  const textAlign = sfs.songTextAlign || 'center';

  let sh = 'none';
  if (sfs.songShadowEnabled === true) {
    const sx = sfs.songShadowX ?? 2, sy = sfs.songShadowY ?? 2;
    const sb = sfs.songShadowBlur ?? 8;
    const so = (sfs.songShadowOpacity ?? 80) / 100;
    const sc = sfs.songShadowColor || '#000000';
    const r = parseInt(sc.slice(1,3),16), g = parseInt(sc.slice(3,5),16), b = parseInt(sc.slice(5,7),16);
    sh = `${sx}px ${sy}px ${sb}px rgba(${r},${g},${b},${so})`;
  } else if (sfs.songShadowEnabled !== false && themeData.shadowOn) {
    sh = '0 2px 10px rgba(0,0,0,.9)';
  }

  let outline = '';
  const outStyle = sfs.songOutlineStyle || 'none';
  if (outStyle !== 'none') {
    const ow = sfs.songOutlineWidth ?? 11;
    const oc = sfs.songOutlineColor || '#000000';
    const oo = (sfs.songOutlineOpacity ?? 100) / 100;
    const r = parseInt(oc.slice(1,3),16), g = parseInt(oc.slice(3,5),16), b = parseInt(oc.slice(5,7),16);
    outline = `-webkit-text-stroke:${ow}px rgba(${r},${g},${b},${oo});paint-order:stroke fill;`;
  }
  const underline = sfs.songUnderline ? 'text-decoration:underline;' : '';
  const fontOpacity = (sfs.songFontOpacity ?? 100) / 100;

  const tr  = _getSongTransform(themeData, sec);
  const pad = themeData.padding || 80;
  const bg  = _getBgCss(themeData);
  const showMeta = State.settings?.showSongMetadata === true;

  // FIX: Scale font to the Live Display panel size, NOT the viewport.
  // The projection handles its own sizing; the Live Display is a ~450px preview panel.
  // Using vw-relative units here made text enormous (e.g. 3.25vw at 1512px = 49px).
  // Scale fs proportionally: Live Display width / 1920 projection width.
  const ldW = liveDisplay.clientWidth || 450;
  const scale = Math.min(1, ldW / 1920);
  const scaledFs = Math.round(fs * scale);
  const scaledPad = Math.round(Math.min(pad, pad * scale * 3));  // reduced padding in preview
  const displayFs = `clamp(11px, ${scaledFs}px, ${scaledFs}px)`;
  const displayMetaFs = `clamp(8px, ${Math.max(8, Math.round(11 * scale))}px, 13px)`;

  const songVA = sfs.songVerticalAlign || 'center';
  const songVAMap = { top: 'flex-start', center: 'center', bottom: 'flex-end' };
  const songJC = songVAMap[songVA] || 'center';

  if (!liveSurface.overlay) {
    liveDisplay.style.cssText = `display:flex;align-items:center;justify-content:${songJC};flex-direction:column;
      padding:${scaledPad}px;text-align:center;position:relative;overflow:hidden;${bg}`;
    _applyLiveThemeBackground(liveDisplay, themeData);
  } else {
    textTarget.style.cssText = `position:absolute;inset:0;display:flex;align-items:center;justify-content:${songJC};flex-direction:column;
      padding:${scaledPad}px;text-align:center;overflow:hidden;background:transparent`;
  }

  const lyricStyle = `font-size:${displayFs};color:${col};font-family:'${ff}',serif;font-weight:${fontWeight};font-style:${fontItalic};line-height:${ls};text-shadow:${sh};text-transform:${tr};text-align:${textAlign};opacity:${fontOpacity};${underline}${outline}`;
  textTarget.innerHTML = `
    <div style="position:relative;z-index:1;width:100%;text-transform:none">
      ${showMeta ? `<div style="font-size:${displayMetaFs};color:${themeData.accentColor||'#c9a84c'};
        letter-spacing:.18em;text-transform:uppercase;margin-bottom:clamp(4px,1vh,10px);
        font-family:'${ff}',serif;opacity:.75">${_escapeHtmlBasic(song.title)}${sec.label ? ' · ' + _escapeHtmlBasic(sec.label) : ''}</div>` : ''}
      ${sec?.richHtml ? `<div class="song-lyric-block" style="${lyricStyle}">${_sanitizeSongHtml(sec.richHtml)}</div>` : lines.map(l=>`<div class="song-lyric-block" style="${lyricStyle}">${_escapeHtmlBasic(l)}</div>`).join('')}
    </div>`;
  textTarget.querySelectorAll('.song-lyric-block, .song-lyric-block *, div[style], span[style], p[style], br').forEach(el => { try { el.style.textTransform = tr; } catch(_) {} });
  if (liveEmpty) liveEmpty.style.display = 'none';
}

// Split verse text into parts that fit the active theme's main text box
function _estimateVerseCapacity(theme) {
  let charsPerSlide = 220;
  if (theme?.boxes?.length) {
    const mainBox = theme.boxes.find(b => b.role === 'main') || theme.boxes[theme.boxes.length - 1];
    if (mainBox) {
      const fitted = _fitBoxFontSize(mainBox, 'W'.repeat(140));
      const approxCharWidth = Math.max(5, fitted * 0.48);
      const usableW = Math.max(120, (mainBox.w || 1680) - 32);
      const usableH = Math.max(80, (mainBox.h || 680) - 32);
      const charsPerLine = Math.max(8, Math.floor(usableW / approxCharWidth));
      const lineH = fitted * (mainBox.lineSpacing || 1.4);
      const maxLines = Math.max(2, Math.floor(usableH / lineH));
      charsPerSlide = Math.floor(charsPerLine * maxLines * 0.82);
    }
  } else if (theme?.fontSize) {
    const approxCharWidth = Math.max(5, theme.fontSize * 0.48);
    const boxW = 1920 - 2 * (theme.padding || 80);
    const boxH = 1080 - 2 * (theme.padding || 80);
    const charsPerLine = Math.max(8, Math.floor(boxW / approxCharWidth));
    const lineH = theme.fontSize * 1.55;
    const maxLines = Math.max(2, Math.floor(boxH / lineH));
    charsPerSlide = Math.floor(charsPerLine * maxLines * 0.82);
  }
  return Math.max(90, charsPerSlide);
}

function _chunkLongSentence(sentence, maxChars) {
  const words = String(sentence || '').trim().split(/\s+/).filter(Boolean);
  if (!words.length) return [];
  const out = [];
  let current = '';
  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    if (candidate.length > maxChars && current) {
      out.push(current.trim());
      current = word;
    } else {
      current = candidate;
    }
  }
  if (current.trim()) out.push(current.trim());
  return out;
}

function _splitVerseForTheme(book, chapter, verse, ref) {
  const fullText = getVerseText(book, chapter, verse);
  if (!fullText) return [];  // FIX: empty array not [null]

  const theme = getActiveThemeData();
  if (!theme) return [fullText];
  const charsPerSlide = _estimateVerseCapacity(theme);
  if (fullText.length <= charsPerSlide) return [fullText];

  const parts = [];
  const sentences = fullText.match(/[^.!?;]+[.!?;]*/g) || [fullText];
  let current = '';
  for (const sentence of sentences) {
    const sent = sentence.trim();
    if (!sent) continue;
    if (sent.length > charsPerSlide) {
      const chunks = _chunkLongSentence(sent, charsPerSlide);
      for (const chunk of chunks) {
        if ((current + ' ' + chunk).trim().length > charsPerSlide && current.trim()) {
          parts.push(current.trim());
          current = chunk;
        } else {
          current = (current ? current + ' ' : '') + chunk;
        }
        if (current.length >= charsPerSlide) {
          parts.push(current.trim());
          current = '';
        }
      }
      continue;
    }
    const candidate = (current ? current + ' ' : '') + sent;
    if (candidate.length > charsPerSlide && current.trim()) {
      parts.push(current.trim());
      current = sent;
    } else {
      current = candidate;
    }
  }
  if (current.trim()) parts.push(current.trim());
  return parts.length ? parts : [fullText];
}

function updateLiveDisplay(book, chapter, verse, ref) {
  if (!book) return;
  State.liveVerse = { book, chapter, verse, ref };
  State.liveContentType = 'scripture';

  const display = document.getElementById('liveDisplay');
  const empty   = document.getElementById('liveEmpty');
  if (!display) return;

  const liveSurface = getLiveTextTarget('scripture');
  const textTarget = liveSurface.host || display;
  if (!liveSurface.overlay) {
    display.style.padding  = '';
    display.style.position = '';
  }

  const parts   = State.verseParts;
  const partIdx = State.versePartIdx || 0;
  const fullText = getVerseText(book, chapter, verse);
  const partText = (parts && parts.length > 1) ? parts[partIdx] : fullText;

  const t = getActiveThemeData();
  if (!liveSurface.overlay && t && !t?.boxes?.length) _applyLiveThemeBackground(display, t);
  const transform = document.documentElement.style.getPropertyValue('--theme-text-transform') || _getScriptureTransform();
  const refPos = State.settings?.scriptureRefPosition || 'top';

  if (t?.boxes?.length) {
    const boxHost = liveSurface.overlay ? textTarget : (() => {
      display.style.cssText = 'display:flex;align-items:center;justify-content:center;padding:0;position:relative;overflow:hidden;background:#000';
      _applyLiveThemeBackground(display, t);
      return getLiveThemeTextHost(display);
    })();
    const _scrSfs = State.settings || {};
    const showVerseNums = (_scrSfs.showVerseNumbers !== false) && (t?.showVerseNum !== false);
    renderBoxThemeToElement(boxHost, t, {
      ref,
      verse,
      translation: State.currentTranslation,
      text: partText || '',
      showVerseNumbers: showVerseNums,
      scriptureShowReference: _scrSfs.scriptureShowReference !== false && (_scrSfs.scriptureRefPosition || 'top') !== 'hidden',
      scriptureRefPosition: _scrSfs.scriptureRefPosition || 'top',
      scriptureTextTransform: _getScriptureTransform(),
      scriptureAbbreviateBooks: !!_scrSfs.scriptureAbbreviateBooks,
      scriptureShowRefOnly: !!_scrSfs.scriptureShowRefOnly,
      scriptureAdditionalLineSpacing: !!_scrSfs.scriptureAdditionalLineSpacing,
      scriptureBreakOnNewVerse: !!_scrSfs.scriptureBreakOnNewVerse,
      scriptureAutoSize: _scrSfs.scriptureAutoSize || 'resize',
      scriptureNormalizeSize: _scrSfs.scriptureNormalizeSize !== false,
      scriptureRefOutlineStyle: _scrSfs.scriptureRefOutlineStyle || 'none',
      scriptureRefOutlineColor: _scrSfs.scriptureRefOutlineColor || '#000000',
      scriptureRefOutlineWidth: _scrSfs.scriptureRefOutlineWidth ?? 0,
      scriptureRefOutlineOpacity: _scrSfs.scriptureRefOutlineOpacity ?? 100,
      fontOverrides: {
        fontFamily: _scrSfs.scriptureFontFamily, fontSize: _scrSfs.scriptureFontSize,
        fontColor: _scrSfs.scriptureFontColor, fontStyle: _scrSfs.scriptureFontStyle,
        lineSpacing: _scrSfs.scriptureLineSpacing, textAlign: _scrSfs.scriptureTextAlign,
        underline: _scrSfs.scriptureUnderline, fontOpacity: _scrSfs.scriptureFontOpacity,
        outlineStyle: _scrSfs.scriptureOutlineStyle, outlineColor: _scrSfs.scriptureOutlineColor,
        outlineWidth: _scrSfs.scriptureOutlineWidth, outlineOpacity: _scrSfs.scriptureOutlineOpacity,
        shadowEnabled: _scrSfs.scriptureShadowEnabled, shadowColor: _scrSfs.scriptureShadowColor,
        shadowBlur: _scrSfs.scriptureShadowBlur, shadowX: _scrSfs.scriptureShadowX, shadowY: _scrSfs.scriptureShadowY,
        shadowOpacity: _scrSfs.scriptureShadowOpacity,
        marginTop: _scrSfs.scriptureMarginTop, marginBottom: _scrSfs.scriptureMarginBottom,
        marginLeft: _scrSfs.scriptureMarginLeft, marginRight: _scrSfs.scriptureMarginRight,
        verticalAlign: _scrSfs.scriptureVerticalAlign,
      },
      scriptureRefFont: {
        fontFamily:    _scrSfs.scriptureRefFontFamily || 'Cinzel',
        fontSize:      _scrSfs.scriptureRefFontSize || 38,
        fontColor:     _scrSfs.scriptureRefFontColor || '#c9a84c',
        fontStyle:     _scrSfs.scriptureRefFontStyle || 'bold',
        underline:     !!_scrSfs.scriptureRefUnderline,
        uppercase:     _scrSfs.scriptureRefUppercase !== false,
        fontOpacity:   _scrSfs.scriptureRefFontOpacity ?? 90,
        textAlign:     _scrSfs.scriptureRefTextAlign || 'center',
        letterSpacing: _scrSfs.scriptureRefLetterSpacing ?? 0.25,
        shadowEnabled: !!_scrSfs.scriptureRefShadowEnabled,
        shadowColor:   _scrSfs.scriptureRefShadowColor || '#000000',
        shadowBlur:    _scrSfs.scriptureRefShadowBlur ?? 4,
        shadowX:       _scrSfs.scriptureRefShadowX ?? 1,
        shadowY:       _scrSfs.scriptureRefShadowY ?? 1,
        shadowOpacity: _scrSfs.scriptureRefShadowOpacity ?? 60,
      },
    });
    if (empty) empty.style.display = 'none';
    return;
  }

  const sfs = State.settings || {};
  const showVerseNums = (sfs.showVerseNumbers !== false) && (t?.showVerseNum !== false);
  const _liveRefColor = t?.refColor || t?.accentColor || '#c9a84c';
  const _liveRoc = _getRefOutlineCss(sfs);
  const verseNum = showVerseNums
    ? `<span class="verse-sup" style="color:${_liveRefColor};${_liveRoc}">${verse}</span>` : '';
  const partLabel = (parts && parts.length > 1)
    ? `<span style="font-size:10px;opacity:.5;margin-left:6px">${partIdx+1}/${parts.length}</span>` : '';
  const sfFamily = sfs.scriptureFontFamily || 'Calibri';
  const sfSize   = sfs.scriptureFontSize || 80;
  const sfColor  = sfs.scriptureFontColor || '#ffffff';
  const sfStyle  = sfs.scriptureFontStyle || 'bold';
  const sfWeight = (sfStyle === 'bold' || sfStyle === 'bold-italic') ? 'bold' : 'normal';
  const sfItalic = (sfStyle === 'italic' || sfStyle === 'bold-italic') ? 'italic' : 'normal';
  const sfUnderline = sfs.scriptureUnderline ? 'underline' : 'none';
  const sfOpacity = (sfs.scriptureFontOpacity ?? 100) / 100;
  const sfLineH  = sfs.scriptureLineSpacing || 1.4;
  const sfAlign  = sfs.scriptureTextAlign || 'center';

  const ldW = display.clientWidth || 450;
  const ldScale = Math.min(1, ldW / 1920);
  const scaledFs = Math.max(11, Math.round(sfSize * ldScale));

  let outlineCss = '';
  const outStyle = sfs.scriptureOutlineStyle || 'none';
  if (outStyle !== 'none') {
    const oo = (sfs.scriptureOutlineOpacity ?? 100) / 100;
    const oc = sfs.scriptureOutlineColor || '#000';
    const _r = parseInt(oc.slice(1,3),16)||0, _g = parseInt(oc.slice(3,5),16)||0, _b = parseInt(oc.slice(5,7),16)||0;
    const scaledOW = Math.max(1, Math.round((sfs.scriptureOutlineWidth||12) * ldScale));
    outlineCss = `-webkit-text-stroke:${scaledOW}px rgba(${_r},${_g},${_b},${oo});paint-order:stroke fill;`;
  }

  let shadowCss = '';
  if (sfs.scriptureShadowEnabled !== false) {
    const so = (sfs.scriptureShadowOpacity ?? 80) / 100;
    const sc = sfs.scriptureShadowColor || '#000000';
    const _sr = parseInt(sc.slice(1,3),16)||0, _sg = parseInt(sc.slice(3,5),16)||0, _sb = parseInt(sc.slice(5,7),16)||0;
    const sx = Math.round((sfs.scriptureShadowX||0) * ldScale);
    const sy = Math.round((sfs.scriptureShadowY||0) * ldScale);
    const sb = Math.max(1, Math.round((sfs.scriptureShadowBlur||0) * ldScale));
    shadowCss = `text-shadow:${sx}px ${sy}px ${sb}px rgba(${_sr},${_sg},${_sb},${so});`;
  }

  const extraLS = sfs.scriptureAdditionalLineSpacing ? 0.4 : 0;
  const effLineH = extraLS ? Math.min(sfLineH + extraLS, 2.4) : sfLineH;
  const bodyStyle = `font-size:${scaledFs}px;text-transform:${transform};font-family:'${sfFamily}',serif;font-weight:${sfWeight};font-style:${sfItalic};text-decoration:${sfUnderline};opacity:${sfOpacity};line-height:${effLineH};text-align:${sfAlign};color:${sfColor};max-width:none;${outlineCss}${shadowCss}`;

  let refOutlineCss = '';
  const refOutSt = sfs.scriptureRefOutlineStyle || 'none';
  if (refOutSt !== 'none') {
    const row = Math.max(1, (sfs.scriptureRefOutlineWidth ?? 0) * ldScale);
    const roc = sfs.scriptureRefOutlineColor || '#000';
    const roo = (sfs.scriptureRefOutlineOpacity ?? 100) / 100;
    const _rr = parseInt(roc.slice(1,3),16)||0, _rg = parseInt(roc.slice(3,5),16)||0, _rb = parseInt(roc.slice(5,7),16)||0;
    refOutlineCss = `-webkit-text-stroke:${row}px rgba(${_rr},${_rg},${_rb},${roo});paint-order:stroke fill;`;
  }

  const showRef = sfs.scriptureShowReference !== false && refPos !== 'hidden';
  const refOnly = !!sfs.scriptureShowRefOnly;
  const effectiveRefPos = showRef ? refPos : 'hidden';
  const dRef = _formatRef(ref);
  const refHtml   = `<div class="verse-ref-label" style="${refOutlineCss}">${dRef} &nbsp;·&nbsp; ${State.currentTranslation}${partLabel}</div>`;
  const textHtml  = refOnly ? '' : `<div class="verse-body-text" style="${bodyStyle}">${verseNum}${partText || ''}</div>`;
  if (effectiveRefPos === 'hidden') {
    textTarget.innerHTML = textHtml;
  } else if (effectiveRefPos === 'bottom') {
    textTarget.innerHTML = textHtml + refHtml;
  } else {
    textTarget.innerHTML = refHtml + textHtml;
  }

  const sfVAlign = sfs.scriptureVerticalAlign || 'bottom';
  const vAlignMap = { top: 'flex-start', center: 'center', bottom: 'flex-end' };
  display.style.justifyContent = vAlignMap[sfVAlign] || 'flex-end';
  const mT = Math.round((sfs.scriptureMarginTop ?? 22) * ldScale);
  const mB = Math.round((sfs.scriptureMarginBottom ?? 22) * ldScale);
  const mL = Math.round((sfs.scriptureMarginLeft ?? 200) * ldScale);
  const mR = Math.round((sfs.scriptureMarginRight ?? 200) * ldScale);
  if (!liveSurface.overlay) display.style.padding = `${mT}px ${mR}px ${mB}px ${mL}px`;

  display.style.display = 'flex';
  if (empty) empty.style.display = 'none';
}

function sendPreviewToLive() {
  State._logoActive = false;
  _updateLogoBtnState();
  // BUG-2 FIX: Song path must not require currentSongSlideIdx != null.
  // After clearLive(), currentSongSlideIdx stays null (we preserve song identity).
  // Without this fix: sendPreviewToLive falls through to the verse check and
  // shows "Select a verse" when Go Live is pressed after Clear on a song.
  // Default to slide 0 when slideIdx is null — the first slide of the song.
  if (State.currentTab === 'song' && State.currentSongId != null) {
    if (!State.isLive) _setLiveOn();
    const slideIdx = State.currentSongSlideIdx ?? 0;
    State.currentSongSlideIdx = slideIdx; // lock it in so subsequent calls anchor correctly
    presentSongSlide(slideIdx);
    return;
  }
  if (State.currentTab === 'media' && State.currentMediaId != null) {
    if (!State.isLive) _setLiveOn();
    presentCurrentMedia();
    return;
  }
  if (State.currentTab === 'pres' && State.currentPresSlides?.length) {
    if (!State.isLive) _setLiveOn();
    presentCurrentPresSlide();
    return;
  }
  if (!State.previewVerse) { toast('⚠ Select a verse in Preview first'); return; }
  if (!State.isLive) _setLiveOn();
  const { book, chapter, verse, ref, endVerse } = State.previewVerse;
  const canRange = endVerse && (State.settings || {}).scriptureShowCompletePassage;
  if (canRange) {
    const verses = buildVerseRangeText(book, chapter, verse, endVerse);
    _updateLiveDisplayRange(book, chapter, verse, endVerse, ref, verses);
    _syncProjectionRange(book, chapter, verse, endVerse, ref, verses);
  } else {
    updateLiveDisplay(book, chapter, verse, ref);
    syncProjection();
  }
  toast(`📺 Live: ${ref}`);
}

// Helper: turn live on and update UI — idempotent, single source of truth.
// BUG-11 FIX: guard against overwriting toggleGoLive()'s "On Air" label.
// If isLive is already true the UI is already correct — do nothing.
function _setLiveOn() {
  if (State.isLive) return;
  State.isLive = true;
  const btn = document.getElementById('goLiveBtn');
  const txt = document.getElementById('goLiveTxt');
  const tag = document.getElementById('liveTag');
  if (btn) btn.classList.add('active');
  if (txt) txt.textContent = 'On Air';
  if (tag) tag.style.display = 'flex';
  setTimeout(_syncRemoteLiveState, 200);
}

// Blank the projection screen without ending the live session.
// The X button in the live display header calls this — keeps isLive=true
// so the next verse/song click pushes directly to projection without
// needing to press Go Live again.
function blankScreen() {
  State._logoActive = false;
  _updateLogoBtnState();

  // Clear live display panel DOM
  State.liveVerse       = null;
  State.liveContentType = null;
  State.currentMediaId  = null;
  State.previewVerse    = null;

  const ld = document.getElementById('liveDisplay');
  if (ld) {
    ld.querySelectorAll('video,audio').forEach(m => { m.pause(); m.src = ''; });
    ld.style.padding  = '';
    ld.style.position = '';
    ld.innerHTML = '';
    ld.style.display = 'none';
  }
  const liveEmpty = document.getElementById('liveEmpty');
  if (liveEmpty) liveEmpty.style.display = 'flex';

  _renderVersePaginationControls();

  // Blank the projection — but DO NOT set isLive=false
  // The session stays live; next content push goes straight to projection
  if (window.electronAPI) window.electronAPI.clearProjection();
  else if (State.isProjectionOpen) _webPostRenderState('clear', null);

  toast('⬜ Screen blanked');
  _syncRemoteLiveState();
}

function clearLive() {
  State._logoActive = false;
  _updateLogoBtnState();
  // Guard: nothing to clear
  const hasLiveContent = !!(State.liveContentType || State.liveVerse || State.currentMediaId);
  if (!hasLiveContent) {
    // Also check if the live display panel has DOM content even without state
    // (can happen when state gets out of sync)
    const ld = document.getElementById('liveDisplay');
    if (!ld || !ld.innerHTML.trim()) { toast('ℹ Nothing is live'); return; }
  }

  // Reset ALL live state — fully symmetrical with toggleGoLive going offline
  State.liveVerse       = null;
  State.liveContentType = null;
  State.currentMediaId  = null;
  State.previewVerse    = null;
  // BUG-2/3 FIX: reset isLive so Go Live must be pressed again to re-establish
  // the live session. This forces syncProjection/presentSongSlide to run fresh.
  State.isLive = false;
  const goBtn = document.getElementById('goLiveBtn');
  const goTxt = document.getElementById('goLiveTxt');
  const goTag = document.getElementById('liveTag');
  if (goBtn) goBtn.classList.remove('active');
  if (goTxt) goTxt.textContent = 'Project Live';
  if (goTag) goTag.style.display = 'none';

  // Preserve song identity (currentSongId, currentSongSlideIdx) so the
  // operator can re-present the same song immediately after clearing.

  _renderVersePaginationControls();

  // BUG-3 FIX: fully wipe the live display DOM — including any song/media
  // content that may still be rendered there even when liveContentType was
  // already null. Without this wipe, the song lyrics stay visible on the
  // operator's live display panel while the projection is blank, and switching
  // to scripture then causes the two to get out of sync.
  const ld2 = document.getElementById('liveDisplay');
  if (ld2) {
    ld2.querySelectorAll('video,audio').forEach(m => { m.pause(); m.src = ''; });
    ld2.style.padding  = '';
    ld2.style.position = '';
    ld2.innerHTML = '';
    ld2.style.display = 'none';
  }
  const liveEmpty = document.getElementById('liveEmpty');
  if (liveEmpty) liveEmpty.style.display = 'flex';

  if (window.electronAPI) window.electronAPI.clearProjection();
  else if (State.isProjectionOpen) _webPostRenderState('clear', null);
  toast('✕ Live display cleared');
  _syncRemoteLiveState();
}

function _webPostRenderState(module, payload) {
  fetch('/api/render-state', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ module, payload, updatedAt: Date.now() })
  }).catch(() => {});
}

function syncProjection() {
  if (State.liveContentType !== 'scripture' || !State.liveVerse) return;
  if (!window.electronAPI && !State.isProjectionOpen) return;
  const { book, chapter, verse, ref } = State.liveVerse;

  // Use the split part if available, else full text
  let text = getVerseText(book, chapter, verse);
  const parts = State.verseParts;
  if (parts && parts.length > 1) {
    const idx = Math.max(0, Math.min(State.versePartIdx || 0, parts.length - 1));
    text = parts[idx];
  }

  const themeData = getActiveThemeData();
  // Inject the per-part text into the main box if box-based theme
  let projThemeData = themeData;
  if (themeData?.boxes?.length && parts && parts.length > 1) {
    projThemeData = JSON.parse(JSON.stringify(themeData));
    const mainBox = projThemeData.boxes.find(b => b.role === 'main') ||
                    projThemeData.boxes[projThemeData.boxes.length - 1];
    if (mainBox) mainBox.text = text;
  }

  const sfs = State.settings || {};
  const showVerseNums = sfs.showVerseNumbers !== false;
  if (projThemeData) projThemeData.showVerseNum = showVerseNums;
  const versePayload = {
    book, chapter, verse, ref, text,
    translation: State.currentTranslation,
    theme: State.currentTheme,
    themeData: projThemeData,
    partIdx:  (State.versePartIdx || 0),
    partTotal: parts?.length || 1,
    showVerseNumbers: showVerseNums,
    scriptureRefPosition: sfs.scriptureRefPosition || 'top',
    scriptureRefLocation: sfs.scriptureRefLocation || 'each',
    scriptureShowReference: sfs.scriptureShowReference !== false && (sfs.scriptureRefPosition || 'top') !== 'hidden',
    scriptureTextTransform: _getScriptureTransform(),
    scriptureAbbreviateBooks: !!sfs.scriptureAbbreviateBooks,
    scriptureShowRefOnly: !!sfs.scriptureShowRefOnly,
    scriptureAdditionalLineSpacing: !!sfs.scriptureAdditionalLineSpacing,
    scriptureBreakOnNewVerse: !!sfs.scriptureBreakOnNewVerse,
    scriptureAutoSize: sfs.scriptureAutoSize || 'resize',
    scriptureNormalizeSize: sfs.scriptureNormalizeSize !== false,
    scriptureRefOutlineStyle: sfs.scriptureRefOutlineStyle || 'none',
    scriptureRefOutlineColor: sfs.scriptureRefOutlineColor || '#000000',
    scriptureRefOutlineWidth: sfs.scriptureRefOutlineWidth ?? 0,
    scriptureRefOutlineOpacity: sfs.scriptureRefOutlineOpacity ?? 100,
    overlayOptions: getOverlayTextPrefs('scripture'),
    scriptureFont: {
      fontFamily:      sfs.scriptureFontFamily || 'Calibri',
      fontSize:        sfs.scriptureFontSize || 80,
      fontColor:       sfs.scriptureFontColor || '#ffffff',
      fontStyle:       sfs.scriptureFontStyle || 'bold',
      underline:       !!sfs.scriptureUnderline,
      fontOpacity:     sfs.scriptureFontOpacity ?? 100,
      lineSpacing:     sfs.scriptureLineSpacing || 1.4,
      textAlign:       sfs.scriptureTextAlign || '',
      verticalAlign:   sfs.scriptureVerticalAlign || '',
      outlineStyle:    sfs.scriptureOutlineStyle || 'none',
      outlineColor:    sfs.scriptureOutlineColor || '#000000',
      outlineJoin:     sfs.scriptureOutlineJoin || 'round',
      outlineWidth:    sfs.scriptureOutlineWidth ?? 12,
      outlineOpacity:  sfs.scriptureOutlineOpacity ?? 100,
      shadowEnabled:   sfs.scriptureShadowEnabled !== false,
      shadowColor:     sfs.scriptureShadowColor || '#000000',
      shadowBlur:      sfs.scriptureShadowBlur ?? 8,
      shadowX:         sfs.scriptureShadowX ?? 2,
      shadowY:         sfs.scriptureShadowY ?? 2,
      shadowOpacity:   sfs.scriptureShadowOpacity ?? 80,
      marginTop:       sfs.scriptureMarginTop ?? 22,
      marginBottom:    sfs.scriptureMarginBottom ?? 22,
      marginLeft:      sfs.scriptureMarginLeft ?? 200,
      marginRight:     sfs.scriptureMarginRight ?? 200,
    },
    scriptureRefFont: {
      fontFamily:      sfs.scriptureRefFontFamily || 'Cinzel',
      fontSize:        sfs.scriptureRefFontSize || 38,
      fontColor:       sfs.scriptureRefFontColor || '#c9a84c',
      fontStyle:       sfs.scriptureRefFontStyle || 'bold',
      underline:       !!sfs.scriptureRefUnderline,
      uppercase:       sfs.scriptureRefUppercase !== false,
      fontOpacity:     sfs.scriptureRefFontOpacity ?? 90,
      textAlign:       sfs.scriptureRefTextAlign || 'center',
      letterSpacing:   sfs.scriptureRefLetterSpacing ?? 0.25,
      shadowEnabled:   !!sfs.scriptureRefShadowEnabled,
      shadowColor:     sfs.scriptureRefShadowColor || '#000000',
      shadowBlur:      sfs.scriptureRefShadowBlur ?? 4,
      shadowX:         sfs.scriptureRefShadowX ?? 1,
      shadowY:         sfs.scriptureRefShadowY ?? 1,
      shadowOpacity:   sfs.scriptureRefShadowOpacity ?? 60,
    },
  };
  if (window.electronAPI) {
    window.electronAPI.projectVerse(versePayload);
  } else {
    _webPostRenderState('scripture', versePayload);
  }
}

async function _syncCurrentProjection() {
  if (!State.isLive) return;
  if (!window.electronAPI && !State.isProjectionOpen) return;
  try {
    if (State.liveContentType === 'scripture') {
      const pv = State.previewVerse || State.liveVerse;
      const canRange = pv?.endVerse && (State.settings || {}).scriptureShowCompletePassage;
      if (canRange) {
        const rv = buildVerseRangeText(pv.book, pv.chapter, pv.verse, pv.endVerse);
        _syncProjectionRange(pv.book, pv.chapter, pv.verse, pv.endVerse, pv.ref, rv);
      } else {
        syncProjection();
      }
    } else if (State.liveContentType === 'song' && State.currentSongId != null) {
      const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
      if (song) await sendSongToProjection(song, State.currentSongSlideIdx ?? 0);
    } else if (State.liveContentType === 'media' && State.currentMediaId != null) {
      const item = (State.media || []).find(m => String(m.id) === String(State.currentMediaId));
      if (item) {
        const mediaPayload = {
          id: item.id, name: item.title || item.name, path: item.path,
          type: item.type, loop: item.loop !== false, mute: item.mute === true,
          volume: Number.isFinite(Number(item.volume)) ? (Number(item.volume) > 2 ? Number(item.volume) / 100 : Number(item.volume)) : 1,
          objectFit: _mediaObjectFit(item)
        };
        if (window.electronAPI) await window.electronAPI.projectMedia(mediaPayload);
        else _webPostRenderState('media', mediaPayload);
      }
    } else if (State.liveContentType === 'presentation') {
      const slide = (State.currentPresSlides || [])[State.currentPresSlideIdx ?? 0];
      const pres = State.presentations?.find(p => String(p.id) === String(State.currentPresId));
      if (slide) {
        if (window.electronAPI) {
          if (pres?.type === 'imported') {
            await window.electronAPI.projectPresentationSlide({ imagePath: slide.imagePath });
          } else {
            await window.electronAPI.projectCreatedSlide({ slide: slide.slide || {} });
          }
        } else {
          if (pres?.type === 'imported') {
            _webPostRenderState('presentation-slide', { imagePath: slide.imagePath });
          } else {
            _webPostRenderState('created-slide', { slide: slide.slide || {} });
          }
        }
      }
    }
  } catch (e) {
    console.warn('[Projection] sync error:', e);
  }
}

function _remoteStep(dir, remoteMode) {
  const liveType = State.liveContentType;
  if (liveType === 'song' && State.currentSongId != null) {
    const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
    if (song?.sections?.length) {
      const cur = State.currentSongSlideIdx ?? 0;
      const next = Math.max(0, Math.min(cur + dir, song.sections.length - 1));
      if (next !== cur) presentSongSlide(next);
    }
  } else if (liveType === 'presentation' || remoteMode === 'presentation') {
    _stepPreviewSlide(dir);
  } else {
    if (dir > 0) nextActiveItem(); else prevActiveItem();
  }
}

function _syncRemoteLiveState() {
  if (!window.electronAPI?.isWeb) return;
  const songObj = State.currentSongId ? (State.songs || []).find(s => String(s.id) === String(State.currentSongId)) : null;
  const st = {
    live: !!State.isLive,
    liveType: State.liveContentType || '',
    songTitle: songObj?.title || '',
    songSlide: State.currentSongSlideIdx,
    verse: State.liveVerse ? { ref: State.liveVerse.ref } : null,
    mediaId: State.currentMediaId || null,
  };
  fetch('/api/status', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(st) }).catch(() => {});
}

// ── remoteStepScripture ───────────────────────────────────────────────────────
// Steps the currently live or previewed scripture verse by ±1, without touching
// State.currentTab or forcing any desktop UI changes.
// Called by the remote's Scripture mode PREV/NEXT buttons.
// Safe to call from any desktop tab context.
function remoteStepScripture(dir) {
  // BUG-1 FIX: If something other than scripture is currently live (song, media,
  // presentation), do NOT navigate bible verses — the remote scripture mode
  // PREV/NEXT should be silent in that case. liveVerse is only set for scripture
  // content so we cannot use it as the sole check; check liveContentType first.
  if (State.isLive && State.liveContentType && State.liveContentType !== 'scripture') {
    toast('\u26A0 ' + State.liveContentType.charAt(0).toUpperCase() + State.liveContentType.slice(1) + ' is live \u2014 clear it first');
    return;
  }

  // Use liveVerse first (what's projected), fall back to previewVerse — but
  // only use previewVerse when it's a genuine bible verse (not a stale value
  // left over from a previous session with a different content type live).
  const anchor = State.liveVerse || (State.liveContentType === null && State.previewVerse ? State.previewVerse : null);
  if (!anchor) { toast('\u26A0 No scripture loaded \u2014 send a verse first'); return; }

  const { book, chapter, verse } = anchor;
  if (!window.BibleDB) return;

  // Get all verse numbers for this chapter
  const chapterVerses = BibleDB.getChapter(book, chapter, State.currentTranslation) || [];
  if (!chapterVerses.length) return;
  const verseNums = chapterVerses.map(v => v.verse || v.v || v);

  const curIdx = verseNums.indexOf(verse);
  let nextIdx  = curIdx + dir;

  // At chapter boundary — step to adjacent chapter
  if (nextIdx < 0 || nextIdx >= verseNums.length) {
    const chapters = BibleDB.getChapters(book);
    const chIdx    = chapters.indexOf(chapter);
    if (dir > 0 && chIdx < chapters.length - 1) {
      const nextCh     = chapters[chIdx + 1];
      const nextVObjs  = BibleDB.getChapter(book, nextCh, State.currentTranslation) || [];
      const nextV      = nextVObjs[0];
      if (!nextV) return;
      const v = nextV.verse || nextV.v || nextV;
      const ref = `${book} ${nextCh}:${v}`;
      _remotePresentVerse(book, nextCh, v, ref);
    } else if (dir < 0 && chIdx > 0) {
      const prevCh    = chapters[chIdx - 1];
      const prevVObjs = BibleDB.getChapter(book, prevCh, State.currentTranslation) || [];
      const prevV     = prevVObjs[prevVObjs.length - 1];
      if (!prevV) return;
      const v = prevV.verse || prevV.v || prevV;
      const ref = `${book} ${prevCh}:${v}`;
      _remotePresentVerse(book, prevCh, v, ref);
    } else {
      toast(dir > 0 ? '📖 Last verse of the Bible book' : '📖 First verse');
    }
    return;
  }

  const nextVerse = verseNums[nextIdx];
  const ref = `${book} ${chapter}:${nextVerse}`;
  _remotePresentVerse(book, chapter, nextVerse, ref);
}

// Present a verse and send to live — used by remote navigation
function _remotePresentVerse(book, chapter, verse, ref) {
  State.previewVerse  = { book, chapter, verse, ref };
  State.currentBook   = book;
  State.currentChapter = chapter;
  updateLiveDisplay(book, chapter, verse, ref);
  if (State.isLive) syncProjection();
  toast(`📖 ${ref}`);
}
// ─────────────────────────────────────────────────────────────────────────────

function refreshCanvases() {
  if (State.previewVerse && State.currentTab !== 'song' && State.currentTab !== 'media' && State.currentTab !== 'pres' && State.currentTab !== 'theme') {
    const { book, chapter, verse, ref, endVerse } = State.previewVerse;
    const canRange = endVerse && (State.settings || {}).scriptureShowCompletePassage;
    if (canRange) {
      presentVerseRange(book, chapter, verse, endVerse, ref, { silent: true });
    } else {
      presentVerse(book, chapter, verse, ref, { parts: State.verseParts || undefined, partIdx: State.versePartIdx || 0, silent: true });
    }
  }

  if (!State.isLive || !State.liveContentType) return;

  if (State.liveContentType === 'song' && State.currentSongId != null && State.currentSongSlideIdx != null) {
    const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
    if (song?.sections?.[State.currentSongSlideIdx]) {
      const sec = song.sections[State.currentSongSlideIdx];
      const lines = (sec.lines || []).filter(l => l.trim());
      const themeData = getActiveSongThemeData();
      renderSongLiveDisplay(song, sec, lines, themeData, 'present');
      sendSongToProjection(song, State.currentSongSlideIdx);
    }
    return;
  }

  if (State.liveContentType === 'scripture' && State.liveVerse) {
    const { book, chapter, verse, ref, endVerse } = State.liveVerse;
    const canRange = endVerse && (State.settings || {}).scriptureShowCompletePassage;
    if (canRange) {
      const rv = buildVerseRangeText(book, chapter, verse, endVerse);
      _updateLiveDisplayRange(book, chapter, verse, endVerse, ref, rv);
      _syncProjectionRange(book, chapter, verse, endVerse, ref, rv);
    } else {
      updateLiveDisplay(book, chapter, verse, ref);
      syncProjection();
    }
    return;
  }

  if (State.liveContentType === 'media' && State.currentMediaId) {
    const item = (State.media || []).find(m => m.id === State.currentMediaId);
    if (item) {
      previewMedia(item);   // FIX: pass item object not item.id (string)
      presentMedia(item);
    }
  }
}

// ─── GO LIVE ──────────────────────────────────────────────────────────────────
function toggleGoLive() {
  State.isLive = !State.isLive;
  const btn = document.getElementById('goLiveBtn');
  const txt = document.getElementById('goLiveTxt');
  const tag = document.getElementById('liveTag');

  btn.classList.toggle('active', State.isLive);
  txt.textContent = State.isLive ? 'On Air' : 'Project Live';
  tag.style.display = State.isLive ? 'flex' : 'none';

  if (State.isLive) {
    if (State.currentTab === 'song' && State.currentSongId != null) sendPreviewToLive();
    else if (State.currentTab === 'media' && State.currentMediaId != null) sendPreviewToLive();
    else if (State.currentTab === 'pres' && State.currentPresSlides?.length) sendPreviewToLive();
    else if (State.previewVerse) sendPreviewToLive();
    toast('🔴 Now broadcasting LIVE');
  } else {
    // Going offline: blank the projection screen and Live Display
    State.liveVerse       = null;
    State.liveContentType = null;  // FIX: reset so refreshCanvases won't re-project
    document.getElementById('liveDisplay').style.display = 'none';
    document.getElementById('liveEmpty').style.display = 'flex';
    if (window.electronAPI) window.electronAPI.clearProjection();
    else if (State.isProjectionOpen) _webPostRenderState('clear', null);
    toast('⬛ Broadcast ended — screen cleared');
  }
}

// ─── QUEUE ────────────────────────────────────────────────────────────────────
function addToQueue(book, chapter, verse, ref) {
  if (!book || !chapter || !verse) return;  // guard missing params
  const safeRef = ref || `${book} ${chapter}:${verse}`;
  // BUG-4 FIX: compare ONLY against other verse items — the old check used
  // `type !== 'song'` which matched song-slide/media/presentation items too,
  // giving false negatives when mixed content was in the queue.
  if (State.queue.find(q => q.type === 'verse' && q.ref === safeRef)) {
    toast(`Already in schedule: ${safeRef}`); return;
  }
  State.queue.push({ type: 'verse', book, chapter, verse, ref: safeRef });
  renderQueue();
  toast(`📋 Added: ${safeRef}`);
}

function addSongToQueue(songId, sectionIdx) {
  const song = (State.songs || []).find(s => String(s.id) === String(songId));
  if (!song) return;

  if (sectionIdx === undefined) {
    // Add whole song as ONE item — shows as single row in schedule
    if (State.queue.find(q => q.type === 'song' && String(q.songId) === String(songId))) {
      toast(`"${song.title}" already in schedule`); return;
    }
    State.queue.push({ type: 'song', songId: song.id, songTitle: song.title, author: song.author || '' });
    toast(`🎵 Added: ${song.title}`);
  } else {
    // Add a single slide — also one item
    const sec = song.sections?.[sectionIdx];
    if (!sec) return;
    State.queue.push({ type: 'song-slide', songId: song.id, sectionIdx, songTitle: song.title, sectionLabel: sec.label });
    toast(`🎵 Added: ${song.title} — ${sec.label}`);
  }
  renderQueue();
}

function addPresentationToQueue(presId) {
  const pres = (State.presentations || []).find(p => String(p.id) === String(presId));
  if (!pres) { toast('⚠ Presentation not found'); return; }
  if (State.queue.find(q => q.type === 'presentation' && String(q.presId) === String(presId))) {
    toast(`Already in schedule: ${pres.name}`); return;
  }
  State.queue.push({ type: 'presentation', presId: pres.id, name: pres.name, slideCount: pres.slideCount || (State.currentPresSlides?.length || 0) });
  renderQueue();
  toast(`📋 Added: ${pres.name}`);
}

function renderQueue() {
  const body      = document.getElementById('queueBody');
  const emptyEl   = document.getElementById('queueEmpty');
  const countEl   = document.getElementById('queueCount');
  const posBar    = document.getElementById('schedPosBar');
  const posText   = document.getElementById('schedPosText');

  // Count only non-section items
  const realItems = State.queue.filter(q => q.type !== 'section');
  countEl.textContent = realItems.length;
  emptyEl.style.display = realItems.length === 0 ? 'block' : 'none';

  // Update position bar with inline nav
  const activeIdx = State.currentQueueIdx ?? -1;
  if (activeIdx >= 0 && realItems.length > 0) {
    const pos = realItems.findIndex(item => State.queue.indexOf(item) === activeIdx) + 1;
    if (posText) posText.textContent = `Item ${pos} of ${realItems.length}`;
    if (posBar)  posBar.style.display = 'flex';
    // Update nav button states
    const prevBtn = document.getElementById('schedPrevBtn');
    const nextBtn = document.getElementById('schedNextBtn');
    if (prevBtn) prevBtn.classList.toggle('active-nav', pos > 1);
    if (nextBtn) nextBtn.classList.toggle('active-nav', pos < realItems.length);
  } else {
    if (posBar) posBar.style.display = 'none';
  }

  // Remove all rendered items (keep emptyEl)
  body.querySelectorAll('.queue-item, .queue-section').forEach(e => e.remove());

  State.queue.forEach((item, i) => {
    // ── Section header ─────────────────────────────────────────────────────
    if (item.type === 'section') {
      const sec = document.createElement('div');
      sec.className = 'queue-section';
      sec.dataset.idx = i;
      sec.innerHTML = `
        <input type="text" value="${escapeHtml(item.label || 'Section')}"
          class="queue-section-input"
          onclick="event.stopPropagation()"
          onchange="updateSectionLabel(${i}, this.value)">
        <button class="queue-section-remove" onclick="removeFromQueue(${i})" title="Remove section">✕</button>`;
      _makeQueueItemDraggable(sec, i);
      _makeQueueDropTarget(sec, i);
      body.appendChild(sec);
      return;
    }

    // ── Regular item ───────────────────────────────────────────────────────
    const isActive = i === State.currentQueueIdx;
    let iconHtml, refText, previewText, typeLabel, typeClass;

    if (item.type === 'song') {
      iconHtml   = '🎵'; typeLabel = 'Song'; typeClass = 'song';
      refText    = item.songTitle || 'Song';
      previewText = item.author || '';
    } else if (item.type === 'song-slide') {
      iconHtml   = '🎵'; typeLabel = 'Slide'; typeClass = 'song';
      refText    = item.songTitle || 'Song';
      previewText = item.sectionLabel || `Slide ${(item.sectionIdx||0)+1}`;
    } else if (item.type === 'media') {
      iconHtml   = item.mediaType === 'video' ? '🎬' : (item.mediaType === 'audio' ? '🎵' : '🖼');
      typeLabel  = item.mediaType === 'video' ? 'Video' : (item.mediaType === 'audio' ? 'Audio' : 'Image');
      typeClass  = 'media';
      refText    = item.name || 'Media';
      previewText = '';
    } else if (item.type === 'presentation') {
      iconHtml   = '📑'; typeLabel = 'Presentation'; typeClass = 'media';
      refText    = item.name || 'Presentation';
      previewText = `${item.slideCount || 0} slides`;
    } else {
      // Bible verse
      iconHtml   = '📖'; typeLabel = 'Verse'; typeClass = 'verse';
      refText    = item.ref || '';
      const text = getVerseText(item.book, item.chapter, item.verse);
      previewText = (text || '').slice(0, 50) + (text?.length > 50 ? '…' : '');
    }

    const div = document.createElement('div');
    div.className = 'queue-item' + (isActive ? ' active' : '');
    div.dataset.idx = i;
    div.innerHTML = `
      <div class="q-handle" title="Drag to reorder">⠿</div>
      <div class="q-thumb">${iconHtml}</div>
      <div class="queue-item-info">
        <div class="queue-ref">${escapeHtml(refText)}</div>
        <div class="queue-preview">${escapeHtml(previewText)}</div>
        <span class="q-type-badge ${typeClass}">${typeLabel}</span>
      </div>
      <div class="queue-actions">
        <button class="icon-btn play" style="width:20px;height:20px;font-size:9px" title="Present">▶</button>
        <button class="icon-btn remove" style="width:20px;height:20px;font-size:9px" title="Remove">✕</button>
      </div>`;

    div.querySelector('.icon-btn.play').addEventListener('click', (e) => {
      e.stopPropagation();
      presentQueueItem(item, i);
    });
    div.querySelector('.icon-btn.remove').addEventListener('click', (e) => {
      e.stopPropagation();
      removeFromQueue(i);
    });
    div.addEventListener('click', (e) => {
      if (e.target.closest('.queue-actions') || e.target.closest('.q-handle')) return;
      presentQueueItem(item, i);
    });

    _makeQueueItemDraggable(div, i);
    _makeQueueDropTarget(div, i);
    body.appendChild(div);
  });

  // Scroll active item into view
  const activeEl = body.querySelector('.queue-item.active');
  if (activeEl) activeEl.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

// ── Drag-and-drop reorder ─────────────────────────────────────────────────────
let _dragSrcIdx = null;

function _makeQueueItemDraggable(el, idx) {
  el.draggable = true;
  el.addEventListener('dragstart', e => {
    _dragSrcIdx = idx;
    el.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
  });
  el.addEventListener('dragend', () => {
    _dragSrcIdx = null;
    el.classList.remove('dragging');
    document.querySelectorAll('.queue-item, .queue-section').forEach(e => e.classList.remove('drag-over'));
  });
  el.addEventListener('dragover', e => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    document.querySelectorAll('.queue-item, .queue-section').forEach(e => e.classList.remove('drag-over'));
    el.classList.add('drag-over');
  });
  el.addEventListener('drop', e => {
    e.preventDefault();
    el.classList.remove('drag-over');
    if (e.dataTransfer.types.includes('application/anchorcast-item')) return;
    const targetIdx = parseInt(el.dataset.idx);
    if (_dragSrcIdx === null || _dragSrcIdx === targetIdx) return;
    // Reorder
    const item = State.queue.splice(_dragSrcIdx, 1)[0];
    const insertAt = targetIdx > _dragSrcIdx ? targetIdx - 1 : targetIdx;
    State.queue.splice(insertAt, 0, item);
    // Update currentQueueIdx to follow the moved item
    if (State.currentQueueIdx === _dragSrcIdx) {
      State.currentQueueIdx = insertAt;
    } else if (State.currentQueueIdx !== null) {
      if (_dragSrcIdx < State.currentQueueIdx && insertAt >= State.currentQueueIdx)
        State.currentQueueIdx--;
      else if (_dragSrcIdx > State.currentQueueIdx && insertAt <= State.currentQueueIdx)
        State.currentQueueIdx++;
    }
    _dragSrcIdx = null;
    renderQueue();
  });
}

function _makeQueueDropTarget(el, idx) {
  el.addEventListener('dragover', e => {
    if (!e.dataTransfer.types.includes('application/anchorcast-item')) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
    document.querySelectorAll('.queue-item, .queue-section').forEach(x => x.classList.remove('drag-over'));
    el.classList.add('drag-over');
  });
  el.addEventListener('dragleave', () => el.classList.remove('drag-over'));
  el.addEventListener('drop', e => {
    const raw = e.dataTransfer.getData('application/anchorcast-item');
    if (!raw) return;
    e.preventDefault();
    e.stopPropagation();
    el.classList.remove('drag-over');
    let payload;
    try { payload = JSON.parse(raw); } catch { return; }
    let newItem;
    if (payload.type === 'verse') {
      if (State.queue.find(q => q.type === 'verse' && q.ref === payload.ref)) {
        toast(`Already in schedule: ${payload.ref}`); return;
      }
      newItem = { type: 'verse', book: payload.book, chapter: payload.chapter, verse: payload.verse, ref: payload.ref };
    } else if (payload.type === 'song') {
      if (State.queue.find(q => q.type === 'song' && String(q.songId) === String(payload.songId))) {
        toast(`"${payload.songTitle}" already in schedule`); return;
      }
      newItem = { type: 'song', songId: payload.songId, songTitle: payload.songTitle, author: payload.author || '' };
    } else if (payload.type === 'media') {
      if (State.queue.find(q => q.type === 'media' && String(q.mediaId) === String(payload.mediaId))) {
        toast(`"${payload.name}" already in schedule`); return;
      }
      newItem = { type: 'media', mediaId: payload.mediaId, name: payload.name, mediaType: payload.mediaType };
    } else { return; }
    const insertAt = idx + 1;
    State.queue.splice(insertAt, 0, newItem);
    if (State.currentQueueIdx !== null && State.currentQueueIdx >= insertAt) State.currentQueueIdx++;
    renderQueue();
    toast(`📋 Dropped: ${newItem.ref || newItem.songTitle || newItem.name}`);
  });
}

function _makeQueueBodyDropTarget() {
  const body = document.getElementById('queueBody');
  if (!body) return;
  body.addEventListener('dragover', e => {
    if (!e.dataTransfer.types.includes('application/anchorcast-item')) return;
    if (e.target.closest('.queue-item, .queue-section')) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  });
  body.addEventListener('drop', e => {
    const raw = e.dataTransfer.getData('application/anchorcast-item');
    if (!raw) return;
    if (e.target.closest('.queue-item, .queue-section')) return;
    e.preventDefault();
    let payload;
    try { payload = JSON.parse(raw); } catch { return; }
    let newItem;
    if (payload.type === 'verse') {
      if (State.queue.find(q => q.type === 'verse' && q.ref === payload.ref)) {
        toast(`Already in schedule: ${payload.ref}`); return;
      }
      newItem = { type: 'verse', book: payload.book, chapter: payload.chapter, verse: payload.verse, ref: payload.ref };
    } else if (payload.type === 'song') {
      if (State.queue.find(q => q.type === 'song' && String(q.songId) === String(payload.songId))) {
        toast(`"${payload.songTitle}" already in schedule`); return;
      }
      newItem = { type: 'song', songId: payload.songId, songTitle: payload.songTitle, author: payload.author || '' };
    } else if (payload.type === 'media') {
      if (State.queue.find(q => q.type === 'media' && String(q.mediaId) === String(payload.mediaId))) {
        toast(`"${payload.name}" already in schedule`); return;
      }
      newItem = { type: 'media', mediaId: payload.mediaId, name: payload.name, mediaType: payload.mediaType };
    } else { return; }
    State.queue.push(newItem);
    renderQueue();
    toast(`📋 Added: ${newItem.ref || newItem.songTitle || newItem.name}`);
  });
}

// ── Section headers ───────────────────────────────────────────────────────────
function addSectionHeader() {
  State.queue.push({ type: 'section', label: 'Section' });
  renderQueue();
  // Auto-focus the label input for immediate editing
  setTimeout(() => {
    const inputs = document.querySelectorAll('#queueBody .queue-section input');
    if (inputs.length) inputs[inputs.length - 1].focus();
  }, 50);
}

function updateSectionLabel(idx, val) {
  if (State.queue[idx]) State.queue[idx].label = val;
}

// ── presentQueueItem — handles ALL types including songs and media ─────────────
function presentQueueItem(item, idx) {
  // Track position for arrow navigation (skip section headers)
  if (item.type !== 'section') State.currentQueueIdx = idx;
  renderQueue();

  if (item.type === 'song') {
    if (State.currentTab !== 'song') switchTab('song');
    selectSong(item.songId);
  } else if (item.type === 'song-slide') {
    if (State.currentTab !== 'song') switchTab('song');
    if (String(State.currentSongId) !== String(item.songId)) selectSong(item.songId);
    if (State.isLive) presentSongSlide(item.sectionIdx);
    else              previewSongSlide(item.sectionIdx);
  } else if (item.type === 'media') {
    // State.media is lazy-loaded when the Media tab is first opened.
    // If it hasn't loaded yet (schedule played before tab was visited),
    // load it now before trying to find and present the item.
    const _doPresent = () => {
      const mediaItem = (State.media || []).find(m => String(m.id) === String(item.mediaId));
      if (mediaItem) {
        if (State.currentTab !== 'media') switchTab('media');
        presentMedia(mediaItem);
      }
    };
    if (!State.media || State.media.length === 0) {
      loadMedia().then(_doPresent).catch(_doPresent);
    } else {
      _doPresent();
    }
  } else if (item.type === 'presentation') {
    const pres = (State.presentations || []).find(p => String(p.id) === String(item.presId));
    if (pres) {
      if (State.currentTab !== 'pres') switchTab('pres');
      openPresentation(pres.id);
      setTimeout(() => presentCurrentPresSlide(), 60);
    }
  } else if (item.type !== 'section') {
    // Bible verse
    presentVerse(item.book, item.chapter, item.verse, item.ref);
    navigateBibleSearch(item.book, item.chapter, item.verse);
    // FIX-3: navigateBibleSearch() calls highlightVerseRow() at setTimeout(60).
    // We call it again at 120ms to guarantee .active is applied on the rendered
    // row before the user can press NEXT/PREV — this is what lets Fix-1 anchor
    // correctly without needing to fall back to the previewVerse scan.
    setTimeout(() => {
      try { highlightVerseRow(item.verse); } catch (_) {}
    }, 120);
  }
}

function removeFromQueue(i) {
  if (State.currentQueueIdx === i) State.currentQueueIdx = null;
  else if (State.currentQueueIdx > i) State.currentQueueIdx--;
  State.queue.splice(i, 1);
  renderQueue();
}

function clearQueue() {
  State.queue = [];
  State.currentQueueIdx = null;
  renderQueue();
  State.scheduleName = null;
  updateScheduleNameBar();
  toast('🗑 Schedule cleared');
}


function moveMediaSelection(direction) {
  const items = State.media || [];
  if (!items.length) return false;
  let idx = items.findIndex(m => String(m.id) === String(State.currentMediaId));
  if (idx < 0) idx = direction > 0 ? -1 : items.length;
  const next = idx + direction;
  if (next < 0 || next >= items.length) return false;
  const item = items[next];
  if (!item) return false;
  previewMedia(item);
  if (State.isLive) presentMedia(item);
  const row = document.querySelector('#mediaList .media-row[data-mediaid="' + item.id + '"]');
  if (row) row.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  return true;
}

// ── previewPrevSlide / previewNextSlide ───────────────────────────────────────
// These are called by the ‹ › overlay arrows on the Program Preview canvas.
// They walk the SLIDE CARDS visible in the Program Preview grid (#bibleSlideGrid
// in slides mode, or the verse rows in #searchResults in list mode).
// They have NOTHING to do with the Schedule queue.
function previewPrevSlide() { _stepPreviewSlide(-1); }
function previewNextSlide() { _stepPreviewSlide(1); }

function _stepPreviewSlide(dir) {
  // --- SONG TAB: step through song sections ---
  if (State.currentTab === 'song' && State.currentSongId != null) {
    const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
    if (!song?.sections?.length) { toast('ℹ No song slides loaded'); return; }
    const current = State.currentSongSlideIdx ?? 0;
    const next = current + dir;
    if (next < 0)                    { toast('🎵 First slide'); return; }
    if (next >= song.sections.length){ toast('🎵 Last slide');  return; }
    navigateSongSlide(dir);
    return;
  }
  // --- PRES TAB: step through presentation slides ---
  if (State.currentTab === 'pres' && State.currentPresSlides?.length) {
    const current = State.currentPresSlideIdx ?? 0;
    const next = current + dir;
    if (next < 0)                            { toast('📑 First slide'); return; }
    if (next >= State.currentPresSlides.length){ toast('📑 Last slide');  return; }
    navigatePresSlide(dir);
    return;
  }
  // --- BIBLE TAB: step through verse cards / rows in Program Preview ---
  if (State.currentTab === 'book') {
    // 1. Try to step within a multi-part verse first
    if (stepVersePart(dir, { syncLive: !!State.liveVerse })) return;

    // 2. Walk the visible slide cards (slides mode) or verse rows (list mode)
    const cards = _bibleSelectableElements();
    if (!cards.length) return;

    // Find current position by .active, then fall back to State.previewVerse
    let idx = cards.findIndex(el => el.classList.contains('active'));
    if (idx < 0 && State.previewVerse) {
      idx = cards.findIndex(el =>
        String(el.dataset.verse) === String(State.previewVerse.verse) &&
        String(el.dataset.partIdx || '0') === String(State.versePartIdx || 0)
      );
    }

    const next = idx + dir;
    if (next < 0) { toast('📖 First slide'); return; }
    if (next >= cards.length) { toast('📖 Last slide'); return; }

    const target = cards[next];
    if (!target) return;

    // Highlight
    cards.forEach(c => c.classList.remove('active'));
    target.classList.add('active');
    const container = target.closest('#searchResults, #bibleSlideGrid');
    if (container) scrollElementIntoContainer(target, container, 'nearest');

    // Read verse info from data attributes
    const verse   = parseInt(target.dataset.verse);
    const partIdx = parseInt(target.dataset.partIdx || '0');
    const ref     = `${State.currentBook} ${State.currentChapter}:${verse}`;
    const parts   = _splitVerseForTheme(State.currentBook, State.currentChapter, verse, ref);

    State.previewVerse  = { book: State.currentBook, chapter: State.currentChapter, verse, ref };
    State.verseParts    = parts;
    State.versePartIdx  = partIdx;
    _updateVersePartIndicator(partIdx + 1, parts.length);
    _updatePreviewNavArrows();

    presentVerse(State.currentBook, State.currentChapter, verse, ref, { parts, partIdx, silent: true });
    if (State.isLive) sendPreviewToLive();
    toast(`📖 ${ref}${parts.length > 1 ? ` · ${partIdx + 1}/${parts.length}` : ''}`);
    return;
  }
  // --- MEDIA TAB: step to adjacent media item ---
  if (State.currentTab === 'media') {
    if (!moveMediaSelection(dir)) {
      toast(dir > 0 ? '🎬 Last item' : '🎬 First item');
    }
  }
}

// Update the ‹ › arrow visibility / disabled state based on position in grid
function _updatePreviewNavArrows() {
  const prevBtn = document.getElementById('previewPrevBtn');
  const nextBtn = document.getElementById('previewNextBtn');
  const vpa     = document.getElementById('versePreviewArea');
  if (!prevBtn || !nextBtn) return;

  if (!State.previewVerse) {
    prevBtn.style.display = 'none';
    nextBtn.style.display = 'none';
    vpa?.classList.remove('has-verse');
    return;
  }

  prevBtn.style.display = '';
  nextBtn.style.display = '';
  vpa?.classList.add('has-verse');

  const cards = _bibleSelectableElements();
  let idx = cards.findIndex(el => el.classList.contains('active'));
  if (idx < 0 && State.previewVerse) {
    idx = cards.findIndex(el => String(el.dataset.verse) === String(State.previewVerse.verse));
  }

  prevBtn.style.opacity = idx <= 0 ? '0.2' : '';
  nextBtn.style.opacity = (idx < 0 || idx >= cards.length - 1) ? '0.2' : '';
  prevBtn.disabled = idx <= 0;
  nextBtn.disabled = (idx < 0 || idx >= cards.length - 1);
}

function nextActiveItem() {
  // Keyboard shortcut-next → same as clicking ›  on Program Preview
  _stepPreviewSlide(1);
}

function prevActiveItem() {
  // Keyboard shortcut-prev → same as clicking ‹  on Program Preview
  _stepPreviewSlide(-1);
}

// ── Arrow through entire service ──────────────────────────────────────────────
// Skips section headers, handles all item types
function nextQueueItem() {
  const realIdxs = State.queue
    .map((q, i) => q.type !== 'section' ? i : -1)
    .filter(i => i >= 0);
  if (!realIdxs.length) return;

  const curPos  = realIdxs.indexOf(State.currentQueueIdx ?? -1);
  const nextPos = curPos + 1;
  if (nextPos >= realIdxs.length) { toast('⚠ End of schedule'); return; }

  const nextIdx = realIdxs[nextPos];
  presentQueueItem(State.queue[nextIdx], nextIdx);

  // Update nav button styles
}

function prevQueueItem() {
  const realIdxs = State.queue
    .map((q, i) => q.type !== 'section' ? i : -1)
    .filter(i => i >= 0);
  if (!realIdxs.length) return;

  const curPos  = realIdxs.indexOf(State.currentQueueIdx ?? -1);
  const prevPos = curPos <= 0 ? 0 : curPos - 1;
  if (curPos <= 0 && State.currentQueueIdx !== null) { toast('⚠ Already at start'); return; }

  // If nothing is active yet, start at first item
  const targetPos = curPos === -1 ? 0 : prevPos;
  const prevIdx = realIdxs[targetPos];
  presentQueueItem(State.queue[prevIdx], prevIdx);
}

// nav btn state handled in renderQueue

// ─── SCHEDULE SAVE / LOAD ─────────────────────────────────────────────────────
function updateScheduleNameBar() {
  const bar = document.getElementById('scheduleNameBar');
  if (!bar) return;
  if (State.scheduleName) {
    bar.textContent = `📋 ${State.scheduleName}`;
    bar.style.display = '';
  } else {
    bar.style.display = 'none';
  }
}



// ── Save (Ctrl+S) — save to existing name, or prompt Save As if new ───────────
async function saveSchedule() {
  if (!State.queue.filter(q => q.type !== 'section').length) {
    toast('⚠ Schedule is empty'); return;
  }
  if (State.scheduleName) {
    // Already has a name — overwrite silently
    const result = await window.electronAPI?.saveSchedule({
      name: State.scheduleName, items: State.queue
    });
    if (result?.success) {
      toast(`💾 Saved: "${State.scheduleName}"`);
      updateScheduleNameBar();
      console.log('[Schedule] Saved to', result.path || State.scheduleName);
    } else {
      toast('⚠ Save failed — try Save As');
    }
  } else {
    saveScheduleAs();
  }
}

// ── Save As — native Save dialog ──────────────────────────────────────────────
async function saveScheduleAs() {
  if (!State.queue.filter(q => q.type !== 'section').length) {
    toast('⚠ Schedule is empty'); return;
  }
  if (!window.electronAPI?.saveScheduleAs) {
    // Browser fallback
    const name = `Service ${new Date().toLocaleDateString()}`;
    const json = JSON.stringify({ name, items: State.queue }, null, 2);
    const a = Object.assign(document.createElement('a'), {
      href: URL.createObjectURL(new Blob([json], { type:'application/json' })),
      download: `${name}.json`
    });
    a.click();
    toast(`💾 Exported: "${name}.json"`);
    return;
  }
  const result = await window.electronAPI.saveScheduleAs({
    items: State.queue,
    currentName: State.scheduleName || `Service ${new Date().toLocaleDateString()}`
  });
  if (result?.canceled) return;
  if (result?.success) {
    State.scheduleName = result.name;
    updateScheduleNameBar();
    toast(`💾 Saved: "${result.name}"`);
    console.log('[Schedule] Saved to', result.path || result.name);
  } else {
    toast('⚠ Save failed: ' + (result?.error || 'unknown error'));
  }
}

// ── New Schedule ───────────────────────────────────────────────────────────────
function newSchedule() {
  // BUG-8 FIX: Electron blocks native confirm() in some builds (contextIsolation=true).
  // Use a lightweight custom confirmation modal instead.
  if (State.queue.length) {
    _confirmModal(
      'Start a new empty schedule? Unsaved changes will be lost.',
      () => _doNewSchedule()
    );
  } else {
    _doNewSchedule();
  }
}
function _doNewSchedule() {
  State.queue = [];
  State.scheduleName = null;
  State.currentQueueIdx = null;
  renderQueue();
  updateScheduleNameBar();
  toast('📋 New schedule');
}

// Reusable lightweight confirm modal (avoids native confirm() which is blocked in Electron)
function _confirmModal(message, onConfirm, onCancel) {
  document.getElementById('_acConfirmModal')?.remove();
  const overlay = document.createElement('div');
  overlay.id = '_acConfirmModal';
  overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:1000;display:flex;align-items:center;justify-content:center';
  const box = document.createElement('div');
  box.style.cssText = 'background:var(--panel);border:1px solid var(--border-lit);border-radius:8px;padding:24px 22px;width:360px;box-shadow:0 8px 32px rgba(0,0,0,.7)';
  const msg = document.createElement('div');
  msg.style.cssText = 'font-size:13px;color:var(--text);line-height:1.55;margin-bottom:20px';
  msg.textContent = message;
  const btns = document.createElement('div');
  btns.style.cssText = 'display:flex;gap:8px;justify-content:flex-end';
  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'Cancel';
  cancelBtn.style.cssText = 'padding:7px 18px;background:var(--card);border:1px solid var(--border-lit);border-radius:5px;color:var(--text);font-size:12px;cursor:pointer;font-family:inherit';
  cancelBtn.onclick = () => { overlay.remove(); if (onCancel) onCancel(); };
  const okBtn = document.createElement('button');
  okBtn.textContent = 'Continue';
  okBtn.style.cssText = 'padding:7px 18px;background:var(--live);border:none;border-radius:5px;color:#fff;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit';
  okBtn.onclick = () => { overlay.remove(); onConfirm(); };
  btns.appendChild(cancelBtn);
  btns.appendChild(okBtn);
  box.appendChild(msg);
  box.appendChild(btns);
  overlay.appendChild(box);
  document.body.appendChild(overlay);
  overlay.addEventListener('click', e => { if (e.target === overlay) { overlay.remove(); if (onCancel) onCancel(); } });
  okBtn.focus();
}

// ── Open Schedule — native Open dialog ────────────────────────────────────────
async function openSchedule() {
  if (!window.electronAPI?.openScheduleDialog) {
    openLoadScheduleModal(); return; // fallback to in-app picker
  }
  const result = await window.electronAPI.openScheduleDialog();
  if (result?.canceled) return;
  if (result?.success) {
    _applyLoadedSchedule(result.schedule);
  } else {
    toast('⚠ Could not open file: ' + (result?.error || 'unknown'));
  }
}

// ── Open Schedule (in-app modal picker) ───────────────────────────────────────
async function openLoadScheduleModal() {
  if (!window.electronAPI) { toast('ℹ Schedule loading requires the desktop app'); return; }
  const schedules = await window.electronAPI.loadSchedules();

  document.getElementById('scheduleModal')?.remove();

  const modal = document.createElement('div');
  modal.id = 'scheduleModal';
  modal.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:900;display:flex;align-items:center;justify-content:center';

  const noSaved = !schedules.length;
  const inner = document.createElement('div');
  inner.style.cssText = 'background:var(--panel);border:1px solid var(--border-lit);border-radius:8px;width:420px;max-height:520px;display:flex;flex-direction:column;box-shadow:0 8px 32px rgba(0,0,0,.6)';

  // Header
  const header = document.createElement('div');
  header.style.cssText = 'padding:14px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-shrink:0';
  header.innerHTML = '<div style="font-size:13px;font-weight:600;color:var(--text)">📂 Open Schedule</div>';
  const closeX = document.createElement('button');
  closeX.textContent = '✕';
  closeX.style.cssText = 'background:none;border:none;color:var(--text-dim);font-size:16px;cursor:pointer;padding:2px 6px';
  closeX.onclick = () => modal.remove();
  header.appendChild(closeX);
  inner.appendChild(header);

  // Body — BUG-7 FIX: use data-idx instead of inline JSON to avoid parse-break
  // on schedule names that contain quotes, apostrophes or backslashes.
  const body = document.createElement('div');
  body.style.cssText = 'flex:1;overflow-y:auto;padding:8px';

  if (noSaved) {
    body.innerHTML = '<div style="padding:32px;text-align:center;color:var(--text-dim);font-size:12px">No saved schedules yet.<br>Build a schedule and press <strong>Ctrl+S</strong> to save.</div>';
  } else {
    schedules.forEach((s, idx) => {
      const row = document.createElement('div');
      row.className = '_sched-row';
      row.dataset.idx = idx;
      row.style.cssText = 'padding:10px 12px;border-radius:6px;margin-bottom:4px;cursor:pointer;border:1px solid var(--border);transition:background .12s;display:flex;align-items:center;gap:10px';

      const info = document.createElement('div');
      info.style.cssText = 'flex:1;min-width:0';
      const itemCount = (s.items || []).filter(i => i.type !== 'section').length;
      const dateStr   = s.savedAt ? new Date(s.savedAt).toLocaleDateString('en-US', { month:'short', day:'numeric', year:'numeric' }) : '';
      info.innerHTML  = `<div style="font-size:12px;font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(s.name)}</div>
        <div style="font-size:10px;color:var(--text-dim);margin-top:2px">${itemCount} items${dateStr ? ' · ' + dateStr : ''}</div>`;

      const del = document.createElement('button');
      del.textContent = '🗑';
      del.title = 'Delete this schedule';
      del.style.cssText = 'background:none;border:none;color:var(--text-dim);cursor:pointer;font-size:13px;padding:4px 6px;border-radius:4px;flex-shrink:0';
      del.onmouseenter = () => { del.style.color = 'var(--live)'; };
      del.onmouseleave = () => { del.style.color = 'var(--text-dim)'; };
      del.onclick = (e) => { e.stopPropagation(); deleteScheduleFile(s.file); };

      row.onmouseenter = () => { row.style.background = 'var(--card)'; };
      row.onmouseleave = () => { row.style.background = ''; };
      row.onclick = () => { _scheduleModalSelectIdx(schedules, idx); };

      row.appendChild(info);
      row.appendChild(del);
      body.appendChild(row);
    });
  }
  inner.appendChild(body);

  // Footer
  const foot = document.createElement('div');
  foot.style.cssText = 'padding:10px 14px;border-top:1px solid var(--border);display:flex;gap:8px;justify-content:flex-end;flex-shrink:0';
  const browseBtn = document.createElement('button');
  browseBtn.textContent = '📂 Browse Files…';
  browseBtn.style.cssText = 'padding:6px 14px;background:rgba(201,168,76,.15);border:1px solid var(--gold-dim);border-radius:5px;color:var(--gold);font-size:11px;cursor:pointer;font-family:inherit';
  browseBtn.onclick = () => { openSchedule(); modal.remove(); };
  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'Close';
  cancelBtn.style.cssText = 'padding:6px 14px;background:var(--card);border:1px solid var(--border-lit);border-radius:5px;color:var(--text);font-size:11px;cursor:pointer;font-family:inherit';
  cancelBtn.onclick = () => modal.remove();
  foot.appendChild(browseBtn);
  foot.appendChild(cancelBtn);
  inner.appendChild(foot);

  modal.appendChild(inner);
  document.body.appendChild(modal);
  modal.addEventListener('click', e => { if (e.target === modal) modal.remove(); });
}

function _scheduleModalSelectIdx(schedules, idx) {
  document.getElementById('scheduleModal')?.remove();
  const s = schedules[idx];
  if (!s) return;
  loadScheduleObj(s);
}

function loadScheduleObj(scheduleObj) {
  document.getElementById('scheduleModal')?.remove();
  if (!scheduleObj?.items?.length) { toast('⚠ Empty schedule'); return; }
  if (State.queue.length > 0) {
    _confirmModal(
      `Replace current schedule with "${scheduleObj.name}"?`,
      () => _applyLoadedSchedule(scheduleObj)
    );
    return;
  }
  _applyLoadedSchedule(scheduleObj);
}

// Keep old name for backward compat
function loadSchedule(s) { loadScheduleObj(s); }

function _applyLoadedSchedule(scheduleObj) {
  State.queue = scheduleObj.items || [];
  State.scheduleName = scheduleObj.name || null;
  State.currentQueueIdx = null;
  renderQueue();
  updateScheduleNameBar();
  const count = State.queue.filter(q=>q.type!=='section').length;
  toast(`📂 Loaded: "${scheduleObj.name}" — ${count} items`);
}

async function deleteScheduleFile(filename) {
  // BUG-10 FIX: use _confirmModal (native confirm() blocked in Electron).
  // Also await the delete before reopening the modal so deleted files don't
  // still appear in the refreshed list.
  _confirmModal('Delete this saved schedule? This cannot be undone.', async () => {
    if (window.electronAPI) await window.electronAPI.deleteSchedule(filename);
    openLoadScheduleModal();
  });
}

// ── Export / Import (file system via dialog) ──────────────────────────────────
async function exportSchedule() { saveScheduleAs(); }  // same action

async function importSchedule() {
  if (!window.electronAPI?.openScheduleDialog) return;
  const result = await window.electronAPI.openScheduleDialog();
  if (result?.canceled) return;
  if (result?.success) _applyLoadedSchedule(result.schedule);
}

// ── Wire File menu events from main process ───────────────────────────────────
function _wireMenuEvents() {
  if (!window.electronAPI) return;
  window.electronAPI.on('menu-schedule-new',       () => newSchedule());
  window.electronAPI.on('menu-schedule-save',      () => saveSchedule());
  window.electronAPI.on('menu-schedule-save-as',   () => saveScheduleAs());
  window.electronAPI.on('menu-schedule-open',      () => openLoadScheduleModal());
  window.electronAPI.on('menu-schedule-export',    () => exportSchedule());
  window.electronAPI.on('menu-schedule-import',    () => importSchedule());
  window.electronAPI.on('menu-preset-save',        () => savePresetDialog());
  window.electronAPI.on('menu-preset-load',        () => openManagePresets());
  window.electronAPI.on('menu-show-getstarted',    () => showGetStarted());
  window.electronAPI.on('menu-preset-manage',      () => openManagePresets());
  window.electronAPI.on('menu-schedule-load-file', async (file) => {
    const result = await window.electronAPI.loadScheduleFile(file);
    if (result?.success) _applyLoadedSchedule(result.schedule);
    else toast('⚠ Could not load: ' + (result?.error || file));
  });
}


async function _checkPendingScheduleLaunch(retries = 8) {
  if (!window.electronAPI?.consumePendingScheduleOpen || !window.electronAPI?.loadScheduleFile) return;
  // BUG-6 FIX: if the operator already manually loaded a schedule while we
  // were retrying, abort — never clobber a manually opened schedule.
  if (State.scheduleName || State.queue.length > 0) return;
  try {
    const file = await window.electronAPI.consumePendingScheduleOpen();
    if (!file) {
      if (retries > 0) setTimeout(() => _checkPendingScheduleLaunch(retries - 1), 250);
      return;
    }
    const result = await window.electronAPI.loadScheduleFile(file);
    if (result?.success) {
      _applyLoadedSchedule(result.schedule);
      toast(`📂 Loaded schedule: ${result.schedule?.name || ''}`.trim());
    } else {
      toast('⚠ Could not load startup schedule: ' + (result?.error || file));
    }
  } catch (_) {
    if (retries > 0) setTimeout(() => _checkPendingScheduleLaunch(retries - 1), 250);
  }
}

function moveQueueItemUp(i) {
  if (i === 0) return;
  [State.queue[i-1], State.queue[i]] = [State.queue[i], State.queue[i-1]];
  renderQueue();
}

// nextQueueItem/prevQueueItem defined above

// ─── SEARCH ───────────────────────────────────────────────────────────────────

// Navigate Bible Search panel to a specific book/chapter/verse.
// Called when presenting from AI Detections so operator sees the full chapter context.
function navigateBibleSearch(book, chapter, verse) {
  if (!book || !chapter) return;

  // Switch to Book Search tab if not already there
  if (State.currentTab !== 'book') switchTab('book');

  // Update state
  State.currentBook    = book;
  State.currentChapter = chapter;

  // Update the chapter nav title and search input
  const navTitle = document.getElementById('chapterNavTitle');
  const searchInput = document.getElementById('searchInput');
  if (navTitle) navTitle.textContent = `${book} ${chapter}`;
  if (searchInput) searchInput.value = `${book} ${chapter}`;

  // Render the chapter
  renderSearchResults();

  // Highlight and scroll to the specific verse after render
  if (verse) {
    setTimeout(() => highlightVerseRow(verse), 60);
  }
}
function handleSearchInput(val) {
  if (State.currentTab !== 'book') return;
  val = val.trim();
  const chapterNav = document.getElementById('chapterNav');
  if (chapterNav) chapterNav.style.display = '';
  if (!val) return;

  const fullRef = BibleDB.parseReference(val);
  if (fullRef) {
    State.currentBook    = fullRef.book;
    State.currentChapter = fullRef.chapter;
    document.getElementById('chapterNavTitle').textContent =
      `${State.currentBook} ${State.currentChapter}`;
    renderSearchResults();
    const sfs = State.settings || {};
    if (fullRef.endVerse && sfs.scriptureShowCompletePassage) {
      const rangeRef = `${fullRef.book} ${fullRef.chapter}:${fullRef.verse}-${fullRef.endVerse}`;
      presentVerseRange(fullRef.book, fullRef.chapter, fullRef.verse, fullRef.endVerse, rangeRef, { silent: true });
    }
    setTimeout(() => highlightVerseRow(fullRef.verse), 80);
    return;
  }

  // ── Try "Book Chapter"  e.g. "John 3" or "Romans 8" ──────────────────────
  const bookChapter = val.match(/^((?:\d\s)?[a-zA-Z]+(?:\s[a-zA-Z]+)?)\s+(\d+)$/i);
  if (bookChapter) {
    const book = BibleDB.normalizeBook(bookChapter[1]);
    if (book) {
      State.currentBook    = book;
      State.currentChapter = parseInt(bookChapter[2]);
      document.getElementById('chapterNavTitle').textContent =
        `${State.currentBook} ${State.currentChapter}`;
      renderSearchResults();
      return;
    }
  }

  // ── Try book name only  e.g. "John" or "Romans" ───────────────────────────
  const book = BibleDB.normalizeBook(val);
  if (book && BibleDB.getChapters(book).length > 0) {
    State.currentBook    = book;
    State.currentChapter = BibleDB.getChapters(book)[0];
    document.getElementById('chapterNavTitle').textContent =
      `${State.currentBook} ${State.currentChapter}`;
    renderSearchResults();
  }
}

function handleSearchEnter(val) {
  const trimmed = val.trim();
  if (!trimmed) return;
  // Auto-detect: if it looks like a Bible reference, do direct lookup; otherwise AI context search
  const looksLikeRef = /^[1-3]?\s*[a-z]/i.test(trimmed) &&
    window.BibleDB?.parseReference?.(trimmed)?.book;
  if (looksLikeRef || State.currentTab !== 'context') {
    handleSearchInput(trimmed);
  } else {
    runContextSearch(trimmed);
  }
}

// Auto-routing search: reference → Bible lookup, free text → AI context search
function smartSearch(val, options = {}) {
  const trimmed = String(val || '').trim();
  if (!trimmed) return false;
  const sendLive = !!options.sendLive;
  const ref = window.BibleDB?.parseReference?.(trimmed);

  if (ref?.book) {
    State.currentTab = 'book';
    State.currentBook = ref.book;
    State.currentChapter = ref.chapter;
    State.previewViewMode = 'list';
    State.searchQuery = trimmed;
    try { renderSearchResults(); } catch (_) {}
    const verse = ref.verse || 1;
    const sfs = State.settings || {};
    if (ref.endVerse && sfs.scriptureShowCompletePassage) {
      const rangeRef = `${ref.book} ${ref.chapter}:${ref.verse}-${ref.endVerse}`;
      presentVerseRange(ref.book, ref.chapter, ref.verse, ref.endVerse, rangeRef);
    } else {
      const fullRef = `${ref.book} ${ref.chapter}:${verse}`;
      presentVerse(ref.book, ref.chapter, verse, fullRef);
    }
    try { highlightVerseRow(verse); } catch (_) {}
    if (sendLive) {
      State.isLive = true;
      State.liveContentType = 'scripture';
      sendPreviewToLive();
    } else if (State.isLive) {
      sendPreviewToLive();
    }
    return true;
  }

  // Topic/context search path
  const chapterNav = document.getElementById('chapterNav');
  if (chapterNav) chapterNav.style.display = 'none';
  runContextSearch(trimmed);
  return true;
}

function switchTab(tab) {
  State.currentTab = tab;
  const isSong  = tab === 'song';
  const isMedia = tab === 'media';
  const isPres  = tab === 'pres';
  const isTheme = tab === 'theme';

  document.getElementById('bookTab').classList.toggle('active', tab === 'book' || tab === 'context');
  document.getElementById('songLibTab')?.classList.toggle('active', isSong);
  document.getElementById('mediaLibTab')?.classList.toggle('active', isMedia);
  document.getElementById('presLibTab')?.classList.toggle('active', isPres);
  document.getElementById('themeLibTab')?.classList.toggle('active', isTheme);

  const bibleContent = document.getElementById('bibleSearchContent');
  const songContent  = document.getElementById('songLibContent');
  const mediaContent = document.getElementById('mediaLibContent');
  const presContent  = document.getElementById('presLibContent');
  const themeContent = document.getElementById('themeLibContent');
  const isBible = !isSong && !isMedia && !isPres && !isTheme;
  if (bibleContent) bibleContent.style.display = isBible  ? 'flex' : 'none';
  if (songContent)  songContent.style.display  = isSong   ? 'flex' : 'none';
  if (mediaContent) mediaContent.style.display = isMedia  ? 'flex' : 'none';
  if (presContent)  presContent.style.display  = isPres   ? 'flex' : 'none';
  if (themeContent) themeContent.style.display = isTheme  ? 'flex' : 'none';

  // Always hide media preview + controls when leaving media tab
  if (!isMedia) {
    const mediaPreviewArea    = document.getElementById('mediaPreviewArea');
    const mediaPreviewControls = document.getElementById('mediaPreviewControls');
    if (mediaPreviewArea)     mediaPreviewArea.style.display    = 'none';
    if (mediaPreviewControls) mediaPreviewControls.style.display = 'none';
    const previewVideo = document.querySelector('#mediaPreviewCanvas video');
    if (previewVideo) { previewVideo.pause(); previewVideo.src = ''; }
  }
  // Restore default panel proportions when leaving pres tab
  if (!isPres) {
    const previewPanel = document.getElementById('previewPanel');
    const searchPanel  = document.getElementById('searchPanel');
    if (previewPanel) previewPanel.style.flex = '';
    if (searchPanel)  searchPanel.style.flex  = '';
  }

  // ── Clear Program Preview on every tab switch ───────────────────────────────
  _clearPreviewPanel(tab);

  if (isBible) {
    document.getElementById('chapterNav').style.display = tab === 'book' ? 'flex' : 'none';
    const input = document.getElementById('searchInput');
    input.placeholder = tab === 'book'
      ? 'e.g. Ge 1 · Gen 1:1 · Genesis 1:1'
      : 'Search by topic, phrase, or theme…';
    input.value = '';
    if (tab === 'book') renderSearchResults();
    else document.getElementById('searchResults').innerHTML =
      `<div class="empty-state"><span class="empty-icon">🔍</span>Type a topic or phrase and press Enter.<br>e.g. "God's love", "strength", "salvation"</div>`;
    refreshThemeSwatches();
  } else if (isSong) {
    loadSongs();
    refreshThemeSwatches();
  } else if (isMedia) {
    loadMedia();
  } else if (isPres) {
    loadPresentations();
    const previewPanel = document.getElementById('previewPanel');
    const searchPanel  = document.getElementById('searchPanel');
    if (previewPanel) previewPanel.style.flex = '3';
    if (searchPanel)  searchPanel.style.flex  = '1';
    if (State.currentPresId && State.currentPresSlides.length) {
      const pres = State.presentations.find(p => String(p.id) === String(State.currentPresId));
      if (pres) _showPresPreviewArea(pres);
    }
  } else if (isTheme) {
    loadThemeGrid();
  }
}

// Clear all Program Preview areas and reset state
function _clearPreviewPanel(tab) {
  // Stop any playing media
  const canvas = document.getElementById('mediaPreviewCanvas');
  if (canvas) {
    canvas.querySelectorAll('video,audio').forEach(m => { m.pause(); m.src = ''; });
  }
  // Hide all areas
  document.getElementById('versePreviewArea').style.display  = 'none';
  document.getElementById('bibleSlideArea').style.display    = 'none';
  document.getElementById('songPreviewArea').style.display   = 'none';
  document.getElementById('mediaPreviewArea').style.display  = 'none';
  document.getElementById('mediaPreviewControls').style.display = 'none';
  const presPA = document.getElementById('presPreviewArea');
  if (presPA) presPA.style.display = 'none';
  // Restore verse canvas from theme preview state
  const verseBgEl = document.querySelector('#previewCanvas .verse-bg');
  if (verseBgEl) verseBgEl.style.display = '';
  const thPrevVid = document.getElementById('themePreviewVideo');
  if (thPrevVid) { thPrevVid.pause(); thPrevVid.src = ''; thPrevVid.style.display = 'none'; }
  const thPrevVeil = document.getElementById('themePreviewVeil');
  if (thPrevVeil) thPrevVeil.style.display = 'none';
  const previewCanvasReset = document.getElementById('previewCanvas');
  if (previewCanvasReset && tab !== 'theme') {
    previewCanvasReset.style.background = '';
    previewCanvasReset.style.backgroundImage = '';
    previewCanvasReset.style.backgroundSize = '';
    previewCanvasReset.style.backgroundPosition = '';
  }
  const previewDisplayReset = document.getElementById('previewDisplay');
  if (previewDisplayReset && tab !== 'theme') {
    previewDisplayReset.style.cssText = '';
    previewDisplayReset.innerHTML = '';
    previewDisplayReset.style.display = 'none';
  }
  const previewEmptyReset = document.getElementById('previewEmpty');
  if (previewEmptyReset && tab !== 'theme') {
    previewEmptyReset.style.display = '';
  }
  const vpa = document.getElementById('versePreviewArea');
  if (vpa) { vpa.style.flex = '1'; vpa.style.minHeight = '0'; }
  // For Bible tab in list mode — show empty verse canvas ready for selection
  if (tab === 'book' && State.previewViewMode === 'list') {
    if (vpa) vpa.style.display = 'flex';
  }
  // For theme tab — show a theme preview panel
  if (tab === 'theme') {
    if (vpa) vpa.style.display = 'flex';
    const previewEmpty   = document.getElementById('previewEmpty');
    const previewDisplay = document.getElementById('previewDisplay');
    const previewCanvas  = document.getElementById('previewCanvas');
    if (previewCanvas) {
      previewCanvas.removeAttribute('data-theme');
      previewCanvas.style.background = '';
      previewCanvas.style.backgroundImage = '';
      const vbg = previewCanvas.querySelector('.verse-bg');
      if (vbg) vbg.style.display = 'none';
    }
    if (previewEmpty)   previewEmpty.style.display = 'none';
    if (previewDisplay) {
      previewDisplay.innerHTML = '';  // clear old verse HTML before rendering theme
      previewDisplay.style.display = 'flex';
      _renderThemePreview(previewDisplay);
    }
  }
  // Update panel title
  const pt = document.getElementById('previewPanelTitle');
  if (pt) pt.textContent = tab === 'theme' ? '🎨 Theme Preview' : '👁 Program Preview';
}

// ─── SONGS ENGINE ─────────────────────────────────────────────────────────────
// Bottom panel = song titles list only
// Program Preview = song slides when song selected
// Click slide = preview on Live Display + projection

async function loadSongs() {
  try {
    const s = window.electronAPI ? await window.electronAPI.getSongs() : [];
    State.songs = s || [];
  } catch(e) { State.songs = []; }
  renderSongList();
}

// ── Song list (bottom panel) ──────────────────────────────────────────────────
function renderSongList(filter = '') {
  const body  = document.getElementById('songListBody');
  if (!body) return;
  const empty = document.getElementById('songListEmpty');

  const q = filter.toLowerCase().trim();
  let list;
  if (!q) {
    list = (State.songs || []).slice();
  } else {
    // Score each song: exact title=4, title startsWith=3, title contains=2, author/label=1, lyrics=0
    const scored = (State.songs || []).map(s => {
      const t = (s.title || '').toLowerCase();
      const a = (s.author || '').toLowerCase();
      let score = -1;
      if (t === q) score = 4;
      else if (t.startsWith(q)) score = 3;
      else if (t.includes(q)) score = 2;
      else if (a.includes(q)) score = 1;
      else if (Array.isArray(s.sections)) {
        for (const sec of s.sections) {
          if ((sec.label || '').toLowerCase().includes(q)) { score = 1; break; }
          if (Array.isArray(sec.lines) && sec.lines.some(l => l.toLowerCase().includes(q))) { score = 0; break; }
          if ((sec.body || sec.text || sec.lyrics || '').toLowerCase().includes(q)) { score = 0; break; }
        }
      }
      if (score < 0 && (s.lyrics || '').toLowerCase().includes(q)) score = 0;
      return { s, score };
    }).filter(x => x.score >= 0);
    // Sort: higher score first, then alphabetical by title
    scored.sort((a, b) => b.score - a.score || (a.s.title || '').localeCompare(b.s.title || ''));
    list = scored.map(x => x.s);
  }

  if (!list.length) {
    body.innerHTML = '';
    if (empty) {
      empty.style.display = '';
      empty.innerHTML = filter
        ? `<span class="empty-icon">🔍</span>No songs match "<strong>${escapeHtml(filter)}</strong>"`
        : `<span class="empty-icon">🎵</span>No songs yet.<br>Click <strong>+ New</strong> or open<br><strong>🎵 Songs</strong> in the toolbar.`;
    }
    return;
  }
  if (empty) empty.style.display = 'none';

  body.innerHTML = list.map(s => {
    const isActive = String(s.id) === String(State.currentSongId);
    const slides   = s.sections?.length || 0;
    return `<div class="song-list-item${isActive ? ' active' : ''}" data-songid="${s.id}" draggable="true">
      <div class="song-item-main">
        <div class="song-item-title">${escapeHtml(s.title || 'Untitled')}</div>
        <div class="song-item-meta">${escapeHtml(s.author || '')}${s.author ? ' · ' : ''}${slides} slide${slides !== 1 ? 's' : ''}</div>
      </div>
      <div class="song-item-actions">
        <button class="song-item-btn" data-action="schedule" title="Add all slides to Schedule">+📋</button>
        <button class="song-item-btn" data-action="edit" title="Edit song">✏</button>
        <button class="song-item-btn" data-action="delete" title="Delete song" style="color:var(--live);border-color:rgba(224,82,82,.3)">🗑</button>
      </div>
    </div>`;
  }).join('');
}

// Single delegated click handler on songListBody — defined once, handles all items
function _songListClickHandler(e) {
  const item = e.target.closest('.song-list-item');
  if (!item) return;
  const rawId = item.dataset.songid;
  const song  = (State.songs || []).find(s => String(s.id) === rawId);
  if (!song) return;
  const action = e.target.closest('[data-action]')?.dataset.action;
  if (action === 'schedule') { addSongToQueue(song.id); return; }
  if (action === 'edit')     { if (window.electronAPI) window.electronAPI.openSongManager({ songId: song.id }); return; }
  if (action === 'delete')   { deleteSong(song.id); return; }
  selectSong(song.id);
}
function _songListDragHandler(e) {
  const item = e.target.closest('.song-list-item');
  if (!item) return;
  const rawId = item.dataset.songid;
  const song = (State.songs || []).find(s => String(s.id) === rawId);
  if (!song) return;
  e.dataTransfer.setData('application/anchorcast-item', JSON.stringify({
    type: 'song', songId: song.id, songTitle: song.title, author: song.author || ''
  }));
  e.dataTransfer.effectAllowed = 'copy';
}

function _songListContextHandler(e) {
  const item = e.target.closest('.song-list-item');
  if (!item) return;
  e.preventDefault();
  e.stopPropagation();
  const rawId = item.dataset.songid;
  const song = (State.songs || []).find(s => String(s.id) === rawId);
  if (!song) return;

  // Remove any existing song context menu
  document.getElementById('_songCtxMenu')?.remove();

  const menu = document.createElement('div');
  menu.id = '_songCtxMenu';
  menu.style.cssText = `
    position:fixed;z-index:99999;min-width:180px;
    background:var(--card,#16162a);border:1px solid var(--border-lit,rgba(255,255,255,.12));
    border-radius:8px;padding:4px 0;box-shadow:0 8px 24px rgba(0,0,0,.6);
    font-size:13px;color:var(--text,#e0e0e0);
  `;

  const menuItems = [
    { icon: '✏', label: 'Edit song', action: () => {
        if (window.electronAPI) window.electronAPI.openSongManager({ songId: song.id });
        else selectSong(song.id);
      }
    },
    { icon: '+📋', label: 'Add all slides to Schedule', action: () => addSongToQueue(song.id) },
    { separator: true },
    { icon: '🗑', label: 'Delete song', danger: true, action: () => deleteSong(song.id) },
  ];

  for (const mi of menuItems) {
    if (mi.separator) {
      const sep = document.createElement('div');
      sep.style.cssText = 'height:1px;background:var(--border,rgba(255,255,255,.08));margin:4px 0';
      menu.appendChild(sep);
      continue;
    }
    const btn = document.createElement('div');
    btn.style.cssText = `
      padding:8px 16px;cursor:pointer;display:flex;align-items:center;gap:8px;
      ${mi.danger ? 'color:var(--live,#e05252)' : ''}
    `;
    btn.innerHTML = `<span>${mi.icon}</span><span>${mi.label}</span>`;
    btn.addEventListener('mouseenter', () => btn.style.background = 'var(--card-hover,rgba(255,255,255,.06))');
    btn.addEventListener('mouseleave', () => btn.style.background = '');
    btn.addEventListener('click', () => { menu.remove(); mi.action(); });
    menu.appendChild(btn);
  }

  // Position near cursor, keep in viewport
  document.body.appendChild(menu);
  const mw = menu.offsetWidth, mh = menu.offsetHeight;
  const vw = window.innerWidth, vh = window.innerHeight;
  menu.style.left = Math.min(e.clientX, vw - mw - 8) + 'px';
  menu.style.top  = Math.min(e.clientY, vh - mh - 8) + 'px';

  // Close on click outside
  const close = ev => { if (!menu.contains(ev.target)) { menu.remove(); document.removeEventListener('mousedown', close, true); } };
  document.addEventListener('mousedown', close, true);
}

function deleteSong(songId) {
  const song = (State.songs || []).find(s => String(s.id) === String(songId));
  if (!song) return;
  _confirmModal(`Delete "${song.title || 'Untitled'}"? This cannot be undone.`, async () => {
    State.songs = (State.songs || []).filter(s => String(s.id) !== String(songId));
    if (String(State.currentSongId) === String(songId)) {
      State.currentSongId = null;
      State.currentSongSlides = [];
      const slidesList = document.getElementById('songSlidesList');
      if (slidesList) slidesList.innerHTML = '';
    }
    State.queue = State.queue.filter(q => !(q.type === 'song' && String(q.songId) === String(songId)) && !(q.type === 'song-slide' && String(q.songId) === String(songId)));
    if (window.electronAPI?.saveSongs) await window.electronAPI.saveSongs(State.songs);
    renderSongList();
    renderQueue();
    toast(`🗑 Deleted: "${song.title}"`);
  });
}

// ── Song selection — loads slides into Program Preview ────────────────────────
function selectSong(id) {
  State.currentSongId = id;
  State.currentSongSlideIdx = null;

  // Find song — string compare to handle large timestamp IDs safely
  const song = (State.songs || []).find(s => String(s.id) === String(id));
  if (!song) { console.warn('[Songs] selectSong: not found', id); return; }

  // Update active highlight in list without full re-render
  document.querySelectorAll('#songListBody .song-list-item').forEach(el => {
    el.classList.toggle('active', el.dataset.songid === String(id));
  });

  // Update Program Preview song info bar
  document.getElementById('songPreviewTitle').textContent = song.title || 'Untitled';
  document.getElementById('songPreviewMeta').textContent  =
    (song.author ? song.author + ' · ' : '') + (song.sections?.length || 0) + ' slides';

  // Switch Program Preview: hide verse canvas, show song slides
  document.getElementById('versePreviewArea').style.display = 'none';
  document.getElementById('songPreviewArea').style.display  = 'flex';

  // Update panel title
  const pt = document.getElementById('previewPanelTitle');
  if (pt) pt.textContent = '🎵 Song Slides';

  renderSongSlides();
}

function backToVersePreview() {
  State.currentSongId = null;
  State.currentSongSlideIdx = null;
  State.currentMediaId = null;
  document.getElementById('songPreviewArea').style.display  = 'none';
  document.getElementById('mediaPreviewArea').style.display = 'none';
  document.getElementById('bibleSlideArea').style.display   = 'none';
  // Restore verse canvas to full height (may have been shrunk by Bible slide mode)
  const vpa = document.getElementById('versePreviewArea');
  if (vpa) { vpa.style.display = 'flex'; vpa.style.flex = '1'; vpa.style.minHeight = '0'; }
  const pt = document.getElementById('previewPanelTitle');
  if (pt) pt.textContent = '👁 Program Preview';
  renderSongList(document.getElementById('songSearchInput')?.value || '');
}

// Keep backward compat alias
function backToSongList() { backToVersePreview(); }

function openSongManager() {
  if (window.electronAPI) window.electronAPI.openSongManager();
  else toast('ℹ Song Manager requires the desktop app');
}

// ── Song slides (Program Preview area) ───────────────────────────────────────
function renderSongSlides() {
  const body = document.getElementById('songSlidesList');
  if (!body) return;

  const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
  if (!song || !song.sections?.length) {
    body.innerHTML = `<div class="empty-state" style="font-size:11px">
      No slides yet.<br>Click <strong>✏ Edit</strong> to add lyrics.
    </div>`;
    return;
  }

  // ── Slide card view ──────────────────────────────────────────────────────
  if (State.previewViewMode === 'slides') {
    renderSongSlideCards(song, body);
    return;
  }

  // ── List view (default) ───────────────────────────────────────────────────
  body.innerHTML = song.sections.map((sec, i) => {
    const isActive   = i === State.currentSongSlideIdx;
    const lines      = (sec.lines || []).filter(l => l.trim());
    const lyricsHtml = lines.map(l => escapeHtml(l)).join('<br>');
    return `<div class="song-slide-row${isActive ? ' active' : ''}" data-idx="${i}"
        title="Click → Live Display · Double-click → Projection">
      <div class="song-slide-num">${i + 1}</div>
      <div class="song-slide-content">
        <div class="song-slide-tag" style="${_getSectionLabelStyle(sec.label)}">${escapeHtml(sec.label || `Slide ${i + 1}`)}</div>
        <div class="song-slide-lyrics">${lyricsHtml || '<em style="color:var(--text-dim);font-size:10px">(empty)</em>'}</div>
      </div>
      <div class="song-slide-actions">
        <button class="slide-action-btn live" data-action="live" title="Send Live">▶</button>
        <button class="slide-action-btn sched" data-action="sched" title="Add to Schedule">+📋</button>
      </div>
    </div>`;
  }).join('');
}

// Delegated handler for song slides — wired once in init()
function _songSlideClickHandler(e) {
  const row = e.target.closest('.song-slide-row');
  if (!row) return;
  const idx    = parseInt(row.dataset.idx, 10);
  const action = e.target.closest('[data-action]')?.dataset.action;
  if (action === 'live')  { presentSongSlide(idx); return; }
  if (action === 'sched') {
    const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
    if (song) addSongToQueue(song.id, idx);
    return;
  }
  // Plain row click
  if (e.type === 'dblclick') presentSongSlide(idx);
  else previewSongSlide(idx);
}

// ── Preview slide (shows in Live Display canvas + projection if Go Live) ──────
async function previewSongSlide(idx) {
  State.currentSongSlideIdx = idx;
  State.liveContentType = 'song';
  State.liveVerse = null;
  // Update active class on list rows (list mode)
  document.querySelectorAll('#songSlidesList .song-slide-row').forEach(el => {
    el.classList.toggle('active', parseInt(el.dataset.idx, 10) === idx);
  });
  // Update active class on slide cards (card mode)
  document.querySelectorAll('#songSlidesList .slide-card').forEach((el, i) => {
    el.classList.toggle('active', i === idx);
  });

  const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
  if (!song?.sections?.[idx]) return;
  const lines = (song.sections[idx].lines || []).filter(l => l.trim());

  renderSongLiveDisplay(song, song.sections[idx], lines, getActiveSongThemeData(), 'preview');

  if (State.isLive) await sendSongToProjection(song, idx);
  else if (window.electronAPI?.clearProjection) window.electronAPI.clearProjection();
}

async function presentSongSlide(idx) {
  State._logoActive = false;
  _updateLogoBtnState();
  State.currentSongSlideIdx = idx;
  State.liveVerse = null;
  // Update list rows
  document.querySelectorAll('#songSlidesList .song-slide-row').forEach(el => {
    el.classList.toggle('active', parseInt(el.dataset.idx, 10) === idx);
  });
  // Update slide cards
  document.querySelectorAll('#songSlidesList .slide-card').forEach((el, i) => {
    el.classList.toggle('active', i === idx);
  });

  const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
  if (!song?.sections?.[idx]) return;
  const lines = (song.sections[idx].lines || []).filter(l => l.trim());
  const sec   = song.sections[idx];
  const themeData = getActiveSongThemeData();

  renderSongLiveDisplay(song, sec, lines, themeData, 'present');
  State.liveContentType = 'song';

  if (State.isLive) {
    await sendSongToProjection(song, idx);
    toast(`🎵 ${song.title} — ${sec.label || 'Slide ' + (idx + 1)}`);
  } else {
    if (window.electronAPI?.clearProjection) window.electronAPI.clearProjection();
    toast(`🎵 Preview only — ${song.title} — ${sec.label || 'Slide ' + (idx + 1)}`);
  }
  setTimeout(_syncRemoteLiveState, 200);
}

async function sendSongToProjection(song, idx) {
  const sec   = song.sections[idx];
  const lines = (sec.lines || []).filter(l => l.trim());
  const themeData = getActiveSongThemeData();
  const sfs = State.settings || {};
  const songPayload = {
    title:             song.title,
    author:            song.author || '',
    sectionLabel:      sec.label || '',
    lines,
    themeData,
    theme:             State.currentSongTheme || 'song_sanctuary',
    songVerticalAlign: sfs.songVerticalAlign || '',
    showSongMetadata:  sfs.showSongMetadata === true,
    songTextTransform: _getSongTransform(themeData, sec),
    overlayOptions:    getOverlayTextPrefs('song'),
    songFont: {
      fontFamily:      sfs.songFontFamily || '',
      fontSize:        sfs.songFontSize || 0,
      fontColor:       sfs.songFontColor || '',
      fontStyle:       sfs.songFontStyle || 'bold',
      underline:       !!sfs.songUnderline,
      fontOpacity:     sfs.songFontOpacity ?? 100,
      lineSpacing:     sfs.songLineSpacing || 0,
      textAlign:       sfs.songTextAlign || 'center',
      outlineStyle:    sfs.songOutlineStyle || 'none',
      outlineColor:    sfs.songOutlineColor || '#000000',
      outlineJoin:     sfs.songOutlineJoin || 'round',
      outlineWidth:    sfs.songOutlineWidth ?? 11,
      outlineOpacity:  sfs.songOutlineOpacity ?? 100,
      shadowEnabled:   sfs.songShadowEnabled !== false,
      shadowColor:     sfs.songShadowColor || '#000000',
      shadowBlur:      sfs.songShadowBlur ?? 8,
      shadowX:         sfs.songShadowX ?? 2,
      shadowY:         sfs.songShadowY ?? 2,
      shadowOpacity:   sfs.songShadowOpacity ?? 80,
      marginTop:       sfs.songMarginTop ?? 22,
      marginBottom:    sfs.songMarginBottom ?? 22,
      marginLeft:      sfs.songMarginLeft ?? 38,
      marginRight:     sfs.songMarginRight ?? 38,
      verticalAlign:   sfs.songVerticalAlign || '',
      wordWrap:        sfs.songWordWrap !== false,
      autoSize:        sfs.songAutoSize || 'resize',
      normalizeSize:   sfs.songNormalizeSize !== false,
      capitalizeAll:   !!sfs.songCapitalizeAll,
      capitalizeFirst: !!sfs.songCapitalizeFirst,
    },
    // Copyright / CCLI — from both song-level and global settings
    showCopyright:     !!(sfs.showCopyright),
    ccliNumber:        sfs.ccliNumber        || '',
    ccliLicenseType:   sfs.ccliLicenseType   || 'streaming',
    ccliDisplayFormat: sfs.ccliDisplayFormat || 'ccli_only',
    copyright:         song.copyright        || '',
  };
  if (window.electronAPI) {
    await window.electronAPI.projectSong(songPayload);
  } else {
    _webPostRenderState('song', songPayload);
  }
}

// Reload after Song Manager saves — small delay ensures file write completes
async function reloadSongs() {
  await new Promise(r => setTimeout(r, 150)); // wait for fs.writeFileSync to complete
  try {
    const s = window.electronAPI ? await window.electronAPI.getSongs() : [];
    State.songs = s || [];
  } catch(e) { State.songs = []; }
  renderSongList(document.getElementById('songSearchInput')?.value || '');
  if (State.currentSongId) {
    const song = State.songs.find(s => String(s.id) === String(State.currentSongId));
    if (song) {
      document.getElementById('songPreviewTitle').textContent = song.title || 'Untitled';
      document.getElementById('songPreviewMeta').textContent  =
        (song.author ? song.author + ' · ' : '') + (song.sections?.length || 0) + ' slides';
      renderSongSlides();
    } else {
      backToVersePreview();
    }
  }
}


// ─── LIVE PANEL POPUP AUTO-CLOSE ──────────────────────────────────────────────
// All four live-panel popups share this logic:
// • Only one popup open at a time (opening one closes others)
// • Auto-close when user clicks outside the popup AND its trigger button
const _LIVE_POPUPS = [
  { popId: 'timerPopup',        btnId: 'timerBtn' },
  { popId: 'alertsPopup',       btnId: 'alertsBtn' },
  { popId: 'presetPopup',       btnId: 'presetBtn' },
  { popId: 'logoOverlayPopup',  btnId: 'logoOverlayBtn' },
];

function _closeAllLivePopups(exceptId = null) {
  _LIVE_POPUPS.forEach(({ popId }) => {
    if (popId === exceptId) return;
    const el = document.getElementById(popId);
    if (el) el.style.display = 'none';
  });
}

function _toggleLivePopup(popId) {
  const el = document.getElementById(popId);
  if (!el) return;
  const isOpen = el.style.display !== 'none' && el.style.display !== '';
  _closeAllLivePopups(popId);
  el.style.display = isOpen ? 'none' : (el.style.flexDirection || popId === 'timerPopup' ? 'flex' : 'flex');
  return !isOpen; // returns true if now open
}

// Single document listener — closes popups on outside click
document.addEventListener('mousedown', e => {
  _LIVE_POPUPS.forEach(({ popId, btnId }) => {
    const pop = document.getElementById(popId);
    const btn = document.getElementById(btnId);
    if (!pop || pop.style.display === 'none' || pop.style.display === '') return;
    // If click is inside popup or on its trigger button — keep open
    if (pop.contains(e.target) || btn?.contains(e.target)) return;
    pop.style.display = 'none';
  });
}, true);

// ─── TIMER & CAPTION CONTROLS ─────────────────────────────────────────────────
function toggleTimerPopup() {
  const opened = _toggleLivePopup('timerPopup');
  if (opened) {
    _highlightTimerPosBtn();
    // Restore Start button state if timer already running
    const startBtn = document.getElementById('timerInlineStartBtn');
    if (startBtn) {
      if (_liveTimerRunning) {
        startBtn.textContent = '⏸ Running';
        startBtn.style.background  = 'linear-gradient(135deg,#1e5fa8,#3b82f6)';
        startBtn.style.color       = '#fff';
        startBtn.style.borderColor = '#3b82f6';
      } else {
        startBtn.textContent = '▶ Start';
        startBtn.style.background  = 'linear-gradient(135deg,#c9a84c,#e0bf6a)';
        startBtn.style.color       = '#000';
        startBtn.style.borderColor = '#c9a84c';
      }
    }
  }
}
function _highlightTimerPosBtn() {
  const ids = [
    ['timerPosCenter','timerInlinePosCenter','center'],
    ['timerPosTop','timerInlinePosTop','top'],
    ['timerPosEdge','timerInlinePosEdge','edge'],
  ];
  ids.forEach(([modal,inline,pos]) => {
    const active = _timerPosition === pos ? 'var(--accent)' : '';
    const m = document.getElementById(modal); if (m) m.style.background = active;
    const i = document.getElementById(inline); if (i) i.style.background = active;
  });
}

let _timerScale = 1.0;
let _timerPosition = 'top';

function setTimerPosition(pos) {
  _timerPosition = pos;
  _highlightTimerPosBtn();
  if (window.electronAPI) {
    window.electronAPI.timerScale?.({ scale: _timerScale, position: _timerPosition });
  }
  toast(`⏱ Position: ${pos}`);
}

function adjustTimerSize(dir) {
  _timerScale = Math.max(0.4, Math.min(2.0, _timerScale + dir * 0.1));
  const pct = Math.round(_timerScale * 100);
  toast(`⏱ Timer size: ${pct}%`);
  if (window.electronAPI) {
    window.electronAPI.timerScale?.({ scale: _timerScale, position: _timerPosition });
  }
}


function resetProjectionTimer() {
  _timerScale = 1.0;
  _timerPosition = 'edge';

  const inlineMinutes = document.getElementById('timerInlineMinutes');
  const modalMinutes = document.getElementById('timerMinutes');
  const inlineMode = document.getElementById('timerInlineMode');
  const modalMode = document.getElementById('timerMode');
  const inlineLabel = document.getElementById('timerInlineLabel');
  const modalLabel = document.getElementById('timerLabelInput');

  if (inlineMinutes) inlineMinutes.value = 5;
  if (modalMinutes) modalMinutes.value = 5;
  if (inlineMode) inlineMode.value = 'countdown';
  if (modalMode) modalMode.value = 'countdown';
  if (inlineLabel) inlineLabel.value = 'SERVICE STARTS';
  if (modalLabel) modalLabel.value = 'SERVICE STARTS';

  _highlightTimerPosBtn();
  if (window.electronAPI) {
    window.electronAPI.timerScale?.({ scale: _timerScale, position: _timerPosition });
  }
  toast('⏱ Timer reset: Edge, 100%, 5 min');
}

// ── Live timer tracking ───────────────────────────────────────────────────────
let _liveTimerInterval  = null;
let _liveTimerEnd       = null;
let _liveTimerMode      = 'countdown';
let _liveTimerWarnSecs  = 60;
let _liveTimerRunning   = false;

function _startLiveTimerDisplay(seconds, mode, label) {
  _stopLiveTimerDisplay();
  _liveTimerMode     = mode;
  _liveTimerRunning  = true;
  _liveTimerWarnSecs = seconds <= 300 ? 60 : Math.ceil(seconds * 0.1);
  _liveTimerEnd      = mode === 'countdown'
    ? Date.now() + seconds * 1000
    : Date.now();

  // Use the permanent liveTimerLayer overlay — mirrors projection's timerLayer
  // It floats ON TOP of live content without replacing it
  const layer = document.getElementById('liveTimerLayer');
  if (!layer) return;

  // Position mirrors the timer position setting
  const pos = _timerPosition || 'edge';
  const posStyles = {
    center: 'inset:0;align-items:center;justify-content:center;',
    top:    'top:20px;left:0;right:0;align-items:center;',
    edge:   'top:14px;right:14px;align-items:flex-end;',
  };
  const isEdge = pos === 'edge';
  layer.style.cssText =
    'display:flex;position:absolute;z-index:7;flex-direction:column;pointer-events:none;' +
    (posStyles[pos] || posStyles.edge);

  // Font size proportional to live panel — fixed ratio, no scale override
  // The live display is a PREVIEW mirror, not a control surface.
  // We use a fixed proportion of the panel so it always looks right at any zoom.
  const liveCanvas = document.getElementById('liveCanvas');
  const panelW = liveCanvas ? liveCanvas.clientWidth  : 300;
  const panelH = liveCanvas ? liveCanvas.clientHeight : 200;
  // Fixed ratios matching projection appearance in a smaller canvas
  // Edge: small corner display. Center/Top: prominent display.
  const clampDisp = isEdge
    ? Math.round(Math.min(panelW * 0.08, panelH * 0.09))   // edge: small corner
    : Math.round(Math.min(panelW * 0.12, panelH * 0.15));  // center/top: moderate
  const clampLbl = Math.round(clampDisp * 0.28);

  layer.innerHTML =
    '<div id="_ltimerDisp" style="' +
      'font-family:Cinzel,serif;font-weight:700;color:#fff;line-height:1;' +
      'font-size:' + clampDisp + 'px;' +
      'letter-spacing:.04em;transition:color .3s;' +
      'text-shadow:0 0 8px rgba(0,0,0,.9),1px 1px 0 #000,-1px -1px 0 #000">00:00</div>' +
    '<div id="_ltimerLbl" style="' +
      'font-family:Cinzel,serif;font-weight:700;' +
      'font-size:' + clampLbl + 'px;' +
      'letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,.8);' +
      'margin-top:' + Math.round(clampDisp * 0.12) + 'px;' +
      'text-align:' + (isEdge ? 'right' : 'center') + ';' +
      'text-shadow:0 0 4px rgba(0,0,0,1),1px 1px 0 #000">' +
      (label || '') + '</div>';

  // Update Start button
  const startBtn = document.getElementById('timerInlineStartBtn');
  if (startBtn) {
    startBtn.textContent       = '⏸ Running';
    startBtn.style.background  = 'linear-gradient(135deg,#1e5fa8,#3b82f6)';
    startBtn.style.color       = '#fff';
    startBtn.style.borderColor = '#3b82f6';
  }

  function tick() {
    const now  = Date.now();
    let secs;
    if (_liveTimerMode === 'countdown') {
      secs = Math.max(0, Math.ceil((_liveTimerEnd - now) / 1000));
    } else {
      secs = Math.floor((now - _liveTimerEnd) / 1000);
    }
    const m = Math.floor(secs / 60);
    const s = String(secs % 60).padStart(2, '0');
    const disp = document.getElementById('_ltimerDisp');
    if (!disp) { clearInterval(_liveTimerInterval); return; }
    disp.textContent = `${m}:${s}`;
    const isWarn = _liveTimerMode === 'countdown' && secs <= _liveTimerWarnSecs && secs > 0;
    const isOver = _liveTimerMode === 'countdown' && secs === 0;
    disp.style.color = (isWarn || isOver) ? '#e74c3c' : '#fff';
    if (isOver) { _stopLiveTimerDisplay(); return; }
  }

  tick();
  _liveTimerInterval = setInterval(tick, 500);

}

function _stopLiveTimerDisplay() {
  if (_liveTimerInterval) { clearInterval(_liveTimerInterval); _liveTimerInterval = null; }
  _liveTimerRunning = false;

  // Hide the overlay — leave liveDisplay content untouched
  const layer = document.getElementById('liveTimerLayer');
  if (layer) layer.style.display = 'none';

  // Reset Start button
  const startBtn = document.getElementById('timerInlineStartBtn');
  if (startBtn) {
    startBtn.textContent       = '▶ Start';
    startBtn.style.background  = 'linear-gradient(135deg,#c9a84c,#e0bf6a)';
    startBtn.style.color       = '#000';
    startBtn.style.borderColor = '#c9a84c';
  }
}

async function startProjectionTimer() {
  const inlineOpen = document.getElementById('timerPopup')?.style.display === 'flex';
  const mode    = document.getElementById(inlineOpen ? 'timerInlineMode' : 'timerMode')?.value || 'countdown';
  const minutes = parseInt(document.getElementById(inlineOpen ? 'timerInlineMinutes' : 'timerMinutes')?.value || '5');
  const label   = document.getElementById(inlineOpen ? 'timerInlineLabel' : 'timerLabelInput')?.value || '';
  const seconds = minutes * 60;
  if (window.electronAPI) {
    await window.electronAPI.showTimer({ mode, seconds, label, scale: _timerScale, position: _timerPosition });
  }
  // Start live countdown display in operator view
  _startLiveTimerDisplay(seconds, mode, label);
  toast(`⏱ Timer started: ${minutes}min ${mode}`);
}

async function stopProjectionTimer() {
  if (window.electronAPI) await window.electronAPI.stopTimer();
  _stopLiveTimerDisplay();
  toast('⏱ Timer stopped');
}

// ─── ALERTS ───────────────────────────────────────────────────────────────────
let _nurseryAlertTimer = null;

function toggleAlertsPopup() {
  _toggleLivePopup('alertsPopup');
}

function toggleAlertsModal() {
  showModal('alertsOverlay');
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('nurseryAlertSendBtn')?.addEventListener('click', () => sendNurseryAlert());
  document.getElementById('nurseryAlertHideBtn')?.addEventListener('click', () => clearNurseryAlert());
  document.getElementById('messageAlertSendBtn')?.addEventListener('click', () => sendMessageAlert());
  document.getElementById('messageAlertHideBtn')?.addEventListener('click', () => clearMessageAlert());
  document.getElementById('nurseryAlertInput')?.addEventListener('keydown', (e) => { if (e.key === 'Enter') sendNurseryAlert(); });
  document.getElementById('messageAlertInput')?.addEventListener('keydown', (e) => { if (e.key === 'Enter') sendMessageAlert(); });

  document.getElementById('alertsTbBtn')?.addEventListener('click', () => toggleAlertsModal());
  document.getElementById('nurseryAlertModalSendBtn')?.addEventListener('click', () => sendNurseryAlert('modal'));
  document.getElementById('nurseryAlertModalHideBtn')?.addEventListener('click', () => clearNurseryAlert());
  document.getElementById('messageAlertModalSendBtn')?.addEventListener('click', () => sendMessageAlert('modal'));
  document.getElementById('messageAlertModalHideBtn')?.addEventListener('click', () => clearMessageAlert());
  document.getElementById('nurseryAlertModalInput')?.addEventListener('keydown', (e) => { if (e.key === 'Enter') sendNurseryAlert('modal'); });
  document.getElementById('messageAlertModalInput')?.addEventListener('keydown', (e) => { if (e.key === 'Enter') sendMessageAlert('modal'); });

  const alertsOv = document.getElementById('alertsOverlay');
  if (alertsOv) alertsOv.addEventListener('click', (e) => { if (e.target === alertsOv) closeModal('alertsOverlay'); });
});

function _getAlertSettings() {
  const s = State.settings || {};
  return {
    nursery: {
      enabled:    s.nurseryAlertEnabled !== false,
      bgColor:    s.nurseryAlertBgColor || '#cc0000',
      bgOpacity:  s.nurseryAlertBgOpacity ?? 100,
      position:   s.nurseryAlertPosition || 'bottom-left',
      fontSize:   s.nurseryAlertFontSize || 48,
      fontColor:  s.nurseryAlertFontColor || '#ffffff',
      fontFamily: s.nurseryAlertFontFamily || 'Tahoma',
      autoRemove: s.nurseryAlertAutoRemove !== false,
      autoRemoveTime: s.nurseryAlertAutoRemoveTime || 60,
    },
    message: {
      enabled:     s.messageAlertEnabled !== false,
      bgColor:     s.messageAlertBgColor || '#000000',
      bgOpacity:   s.messageAlertBgOpacity ?? 100,
      position:    s.messageAlertPosition || 'top',
      scrollSpeed: s.messageAlertScrollSpeed ?? 5,
      fontSize:    s.messageAlertFontSize || 48,
      fontColor:   s.messageAlertFontColor || '#f5deb3',
      fontFamily:  s.messageAlertFontFamily || 'Calibri',
      repeatCount: s.messageAlertRepeatCount || 2,
    }
  };
}

async function sendNurseryAlert(source) {
  try {
    const inputId = source === 'modal' ? 'nurseryAlertModalInput' : 'nurseryAlertInput';
    const text = document.getElementById(inputId)?.value?.trim() || '';
    if (!text) { toast('Type a nursery alert message first'); return; }
    const cfg = _getAlertSettings();
    if (!cfg.nursery.enabled) { toast('Nursery alerts are disabled in Settings'); return; }
    if (_nurseryAlertTimer) { clearTimeout(_nurseryAlertTimer); _nurseryAlertTimer = null; }
    const payload = { type: 'nursery', text, ...cfg.nursery };
    if (window.electronAPI) await window.electronAPI.showAlert(payload);
    toast(`🔔 Nursery: "${text.slice(0,30)}"`);
    if (cfg.nursery.autoRemove) {
      _nurseryAlertTimer = setTimeout(() => { clearNurseryAlert(); }, cfg.nursery.autoRemoveTime * 1000);
    }
  } catch(err) {
    console.error('[Alert] sendNurseryAlert error:', err);
    toast('Failed to send nursery alert');
  }
}

async function clearNurseryAlert() {
  try {
    if (_nurseryAlertTimer) { clearTimeout(_nurseryAlertTimer); _nurseryAlertTimer = null; }
    if (window.electronAPI) await window.electronAPI.showAlert({ type: 'nursery', text: '' });
    toast('Nursery alert cleared');
  } catch(err) {
    console.error('[Alert] clearNurseryAlert error:', err);
  }
}

async function sendMessageAlert(source) {
  try {
    const inputId = source === 'modal' ? 'messageAlertModalInput' : 'messageAlertInput';
    const text = document.getElementById(inputId)?.value?.trim() || '';
    if (!text) { toast('Type a message alert first'); return; }
    const cfg = _getAlertSettings();
    if (!cfg.message.enabled) { toast('Message alerts are disabled in Settings'); return; }
    const payload = { type: 'message', text, ...cfg.message };
    if (window.electronAPI) await window.electronAPI.showAlert(payload);
    toast(`Message: "${text.slice(0,30)}"`);
  } catch(err) {
    console.error('[Alert] sendMessageAlert error:', err);
    toast('Failed to send message alert');
  }
}

async function clearMessageAlert() {
  try {
    if (window.electronAPI) await window.electronAPI.showAlert({ type: 'message', text: '' });
    toast('Message alert cleared');
  } catch(err) {
    console.error('[Alert] clearMessageAlert error:', err);
  }
}

// ─── MEDIA ENGINE ─────────────────────────────────────────────────────────────
// Media items: { id, name, path, type:'image'|'video', ext, size, duration, loop, mute }

// Convert absolute file path → media:// URL (works from localhost renderer)
function toMediaUrl(filePath) {
  if (!filePath) return '';
  // Normalize backslashes on Windows
  const normalized = filePath.replace(/\\/g, '/');
  return `media:///${normalized.startsWith('/') ? normalized.slice(1) : normalized}`;
}
// List layout with sidebar categories

let mediaCategory = 'all'; // 'all' | 'video' | 'audio' | 'image'

// Formats supported — WMV/AVI via Electron's Chromium
const VIDEO_EXTS  = new Set(['mp4','webm','mov','mkv','avi','wmv','m4v','mpg','mpeg','3gp','flv']);
const AUDIO_EXTS  = new Set(['mp3','wav','ogg','flac','aac','m4a','wma','opus','aiff']);
const IMAGE_EXTS  = new Set(['jpg','jpeg','png','gif','webp','bmp','svg','tiff','tif']);

function getMediaType(filename) {
  const ext = (filename.split('.').pop() || '').toLowerCase();
  if (VIDEO_EXTS.has(ext)) return 'video';
  if (AUDIO_EXTS.has(ext)) return 'audio';
  if (IMAGE_EXTS.has(ext)) return 'image';
  return null;
}

function formatSize(bytes) {
  if (!bytes) return '';
  if (bytes < 1024*1024) return (bytes/1024).toFixed(0) + ' KB';
  return (bytes/(1024*1024)).toFixed(1) + ' MB';
}

function setMediaCategory(cat) {
  mediaCategory = cat;
  document.querySelectorAll('.media-cat').forEach(el => {
    el.classList.toggle('active', el.dataset.cat === cat);
  });
  renderMediaList(document.getElementById('mediaSearchInput')?.value || '');
}

async function loadMedia() {
  try {
    const items = window.electronAPI ? await window.electronAPI.getMedia() : [];
    // Ensure mute defaults to true for any legacy items
    State.media = (items || []).map(m => ({
      ...m,
      mute: m.mute === undefined ? false : m.mute,
      volume: Number.isFinite(Number(m.volume)) ? (Number(m.volume) > 2 ? Number(m.volume) / 100 : Number(m.volume)) : 1,
    }));
  } catch(e) { State.media = []; }
  renderMediaList();
}

function renderMediaList(filter = '') {
  const list  = document.getElementById('mediaList');
  const empty = document.getElementById('mediaEmpty');
  if (!list) return;

  let items = (State.media || []).filter(m => {
    const matchCat = mediaCategory === 'all' || m.type === mediaCategory;
    const matchQ   = !filter || m.name?.toLowerCase().includes(filter.toLowerCase()) || m.title?.toLowerCase().includes(filter.toLowerCase());
    return matchCat && matchQ;
  });

  list.querySelectorAll('.media-row').forEach(e => e.remove());

  if (!items.length) {
    if (empty) {
      empty.style.display = '';
      empty.innerHTML = filter
        ? `<span class="empty-icon">🔍</span>No results for "<strong>${escapeHtml(filter)}</strong>"`
        : `<span class="empty-icon">🖼</span>No ${mediaCategory === 'all' ? '' : mediaCategory + ' '}media yet.<br>Click <strong>+ Import</strong> to add files.`;
    }
    return;
  }
  if (empty) empty.style.display = 'none';

  const chunkSize = mediaCategory === 'all' ? 40 : 120;
  let idx = 0;

  const buildRow = (item) => {
    const isVideo  = item.type === 'video';
    const isActive = String(item.id) === String(State.currentMediaId);
    const row      = document.createElement('div');
    row.className  = 'media-row' + (isActive ? ' active' : '');
    row.dataset.mediaid = String(item.id);

    const useStaticVideoThumb = isVideo && mediaCategory === 'all';
    const thumbHtml = isVideo
      ? (useStaticVideoThumb
          ? `<div style="width:100%;height:100%;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#101522,#05070d);color:#d6c08a;font-size:18px">🎬</div>
             <div class="media-row-thumb-play"></div>
             ${item.duration ? `<span class="media-dur">${item.duration}</span>` : ''}`
          : `<video src="${toMediaUrl(item.path)}" preload="metadata" muted style="pointer-events:none"
               onerror="this.style.display='none';this.nextElementSibling.style.display='flex'"></video>
             <div class="media-row-thumb-play"></div>
             ${item.duration ? `<span class="media-dur">${item.duration}</span>` : ''}`)
      : `<img src="${toMediaUrl(item.path)}" loading="lazy" draggable="false"
           onerror="this.style.display='none'">`;

    row.innerHTML = `
      <div class="media-row-thumb">${thumbHtml}</div>
      <div class="media-row-info">
        <div class="media-row-title">${escapeHtml(item.title || item.name)}</div>
        <div class="media-row-meta">${escapeHtml(item.ext?.toUpperCase() || '')}${item.size ? ' · ' + formatSize(item.size) : ''}</div>
      </div>
      <div class="media-row-type">${item.type === 'video' ? '🎬 Video' : item.type === 'audio' ? '🎵 Audio' : '🖼 Image'}</div>
      <div class="media-row-actions">
        <button class="media-row-btn live" data-action="live" title="Send to projection">▶</button>
        <button class="media-row-btn" data-action="sched" title="Add to schedule">+</button>
        <button class="media-row-btn" data-action="open" title="Open file location">📂</button>
        <button class="media-row-btn" data-action="del" title="Delete from library" style="color:var(--live);border-color:rgba(224,82,82,.3)">🗑</button>
      </div>`;

    row.draggable = true;
    row.addEventListener('dragstart', e => {
      e.dataTransfer.setData('application/anchorcast-item', JSON.stringify({
        type: 'media', mediaId: item.id, name: item.name, mediaType: item.type
      }));
      e.dataTransfer.effectAllowed = 'copy';
    });
    row.addEventListener('click', e => {
      const action = e.target.closest('[data-action]')?.dataset.action;
      if (action === 'live')  { presentMedia(item); return; }
      if (action === 'sched') { addMediaToQueue(item); return; }
      if (action === 'open')  { openMediaFileLocation(item); return; }
      if (action === 'del')   { removeMedia(item.id); return; }
      document.querySelectorAll('#mediaList .media-row').forEach(el => el.classList.remove('active'));
      row.classList.add('active');
      State.currentMediaId = item.id;
    });
    row.addEventListener('dblclick', e => {
      if (!e.target.closest('[data-action]')) previewMedia(item);
    });
    row.addEventListener('contextmenu', e => {
      e.preventDefault();
      e.stopPropagation();
      State.currentMediaId = item.id;
      showMediaContextMenu(e, item);
    });
    return row;
  };

  const appendChunk = () => {
    const frag = document.createDocumentFragment();
    for (let count = 0; idx < items.length && count < chunkSize; idx += 1, count += 1) {
      frag.appendChild(buildRow(items[idx]));
    }
    list.appendChild(frag);
    if (idx < items.length) requestAnimationFrame(appendChunk);
  };
  requestAnimationFrame(appendChunk);
}

// ── Show media in Program Preview (click) ────────────────────────────────────
function previewMedia(item) {
  State.currentMediaId = item.id;

  // Highlight row in list without full re-render
  document.querySelectorAll('#mediaList .media-row').forEach(el => {
    el.classList.toggle('active', el.dataset.mediaid === String(item.id));
  });

  // Update Program Preview info bar
  const titleEl = document.getElementById('mediaPreviewTitle');
  const metaEl  = document.getElementById('mediaPreviewMeta');
  if (titleEl) titleEl.textContent = (item.title || item.name);
  if (metaEl)  metaEl.textContent  = [
    item.ext?.toUpperCase(),
    formatSize(item.size),
    item.type === 'video' ? 'Video' : item.type === 'audio' ? 'Audio' : 'Image',
  ].filter(Boolean).join(' · ');

  // Show/hide video controls (loop/mute only relevant for video/audio)
  const controls = document.getElementById('mediaPreviewControls');
  if (controls) controls.style.display = (item.type === 'video' || item.type === 'audio') ? 'flex' : 'none';

  // Sync checkboxes
  const loopChk = document.getElementById('mediaLoopChk');
  const muteChk = document.getElementById('mediaMuteChk');
  if (loopChk) loopChk.checked = item.loop !== false;
  if (muteChk) muteChk.checked = !!item.mute;

  // Build media element using DOM API
  const canvas = document.getElementById('mediaPreviewCanvas');
  if (canvas) {
    // Stop and remove any existing media
    const oldMedia = canvas.querySelector('video,audio');
    if (oldMedia) { oldMedia.pause(); oldMedia.src = ''; }
    canvas.innerHTML = '';

    const ext = (item.ext || '').toLowerCase();

    if (item.type === 'audio') {
      // Audio file — show waveform placeholder + HTML5 audio controls
      const wrapper = document.createElement('div');
      wrapper.style.cssText = 'display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:14px;padding:20px';
      wrapper.innerHTML = `
        <div style="font-size:40px">🎵</div>
        <div style="font-size:13px;font-weight:600;color:#ddd;text-align:center">${escapeHtml(item.title || item.name)}</div>
        <div style="font-size:10px;color:#888">${ext.toUpperCase()}</div>`;
      const aud = document.createElement('audio');
      aud.controls = true;
      aud.loop  = item.loop !== false;
      aud.style.cssText = 'width:100%;max-width:280px;margin-top:4px';
      aud.src = toMediaUrl(item.path);
      aud.load();
      wrapper.appendChild(aud);
      canvas.appendChild(wrapper);

    } else if (item.type === 'video') {
      {
        const vid = document.createElement('video');
        vid.style.cssText = 'width:100%;height:100%;object-fit:contain;display:block';
        vid.loop     = item.loop !== false;
        vid.muted    = true;
        vid.autoplay = true;
        vid.controls = true; // show native controls for scrubbing
        vid.setAttribute('playsinline', '');
        vid.src = toMediaUrl(item.path);
        vid.onerror = () => {
          canvas.innerHTML = `
            <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
              height:100%;color:#888;text-align:center;padding:20px;gap:8px">
              <div style="font-size:20px">⚠</div>
              <div style="font-size:11px">Cannot preview this format</div>
              <div style="font-size:10px;color:#666">File will still play on projection</div>
            </div>`;
        };
        canvas.appendChild(vid);
        vid.load();
        vid.play().catch(() => {
          const overlay = document.createElement('div');
          overlay.style.cssText = 'position:absolute;inset:0;display:flex;align-items:center;justify-content:center;cursor:pointer;background:rgba(0,0,0,.3)';
          overlay.innerHTML = '<div style="font-size:36px;color:rgba(255,255,255,.8)">▶</div>';
          overlay.onclick = () => { vid.play(); overlay.remove(); };
          canvas.style.position = 'relative';
          canvas.appendChild(overlay);
        });
      }
    } else {
      // Image
      const img = document.createElement('img');
      img.style.cssText = 'width:100%;height:100%;object-fit:contain;display:block';
      img.src = toMediaUrl(item.path);
      img.onerror = () => {
        canvas.innerHTML = `<div style="color:#888;font-size:11px;text-align:center;padding:20px">⚠ Cannot load image</div>`;
      };
      canvas.appendChild(img);
    }
  }

  // Switch Program Preview to media area
  document.getElementById('versePreviewArea').style.display = 'none';
  document.getElementById('songPreviewArea').style.display  = 'none';
  document.getElementById('mediaPreviewArea').style.display = 'flex';

  const icons = { video: '🎬', audio: '🎵', image: '🖼' };
  const pt = document.getElementById('previewPanelTitle');
  if (pt) pt.textContent = `${icons[item.type] || '🖼'} ${item.type === 'video' ? 'Video' : item.type === 'audio' ? 'Audio' : 'Image'} Preview`;
}

function _mediaObjectFit(item) {
  if (item.stretchToOutput) return 'fill';
  const ar = (item.aspectRatio || 'auto').toLowerCase();
  if (ar === 'cover' || ar === 'fill') return 'cover';
  if (ar === 'contain' || ar === 'fit') return 'contain';
  return 'contain';
}

// ── Send to projection (double-click or Go Live button) ───────────────────────
async function presentMedia(item) {
  State._logoActive = false;
  _updateLogoBtnState();
  State.liveContentType = 'media';
  State.liveVerse = null;
  State.currentMediaId = item.id;
  if (!State.isLive) _setLiveOn();
  previewMedia(item);

  const liveDisplay = document.getElementById('liveDisplay');
  const liveEmpty   = document.getElementById('liveEmpty');
  if (liveDisplay) {
    const oldMedia = liveDisplay.querySelector('video,audio');
    if (oldMedia) { oldMedia.pause(); oldMedia.src = ''; }
    liveDisplay.innerHTML = '';

    if (item.type === 'audio') {
      liveDisplay.style.padding = '';
      liveDisplay.style.position = 'relative';
      liveDisplay.innerHTML = `
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
          height:100%;gap:10px;padding:16px;text-align:center">
          <div style="font-size:32px">🎵</div>
          <div style="font-size:11px;font-weight:600;color:#ddd">${escapeHtml(item.title || item.name)}</div>
        </div>`;
      const aud = document.createElement('audio');
      aud.style.cssText = 'position:absolute;bottom:12px;left:10px;right:10px;width:calc(100% - 20px)';
      aud.controls = true;
      aud.loop = item.loop !== false;
      aud.src  = toMediaUrl(item.path);
      aud.load(); aud.play().catch(()=>{});
      liveDisplay.appendChild(aud);
    } else if (item.type === 'video') {
      liveDisplay.style.padding  = '0';
      liveDisplay.style.position = 'relative';
      const fitMode = _mediaObjectFit(item);
      const vid = document.createElement('video');
      vid.style.cssText = `position:absolute;inset:0;width:100%;height:100%;object-fit:${fitMode};display:block`;
      vid.loop     = item.loop !== false;
      vid.muted    = true;
      vid.autoplay = true;
      vid.setAttribute('playsinline', '');
      vid.src = toMediaUrl(item.path);
      liveDisplay.appendChild(vid);
      vid.load();
      vid.play().catch(() => {
        const ov = document.createElement('div');
        ov.style.cssText = 'position:absolute;inset:0;display:flex;align-items:center;justify-content:center;cursor:pointer;background:rgba(0,0,0,.5)';
        ov.innerHTML = '<span style="font-size:32px;color:#fff">▶</span>';
        ov.onclick = () => { vid.play(); ov.remove(); };
        liveDisplay.appendChild(ov);
      });
    } else {
      liveDisplay.style.padding  = '0';
      liveDisplay.style.position = 'relative';
      const fitMode = _mediaObjectFit(item);
      const img = document.createElement('img');
      img.style.cssText = `position:absolute;inset:0;width:100%;height:100%;object-fit:${fitMode};display:block`;
      img.src = toMediaUrl(item.path);
      liveDisplay.appendChild(img);
    }

    liveDisplay.style.display = 'flex';
    if (liveEmpty) liveEmpty.style.display = 'none';
  }

  const effectiveAR = _mediaObjectFit(item);
  const mediaPayload = {
    id: item.id, name: item.title || item.name, path: item.path,
    type: item.type, loop: item.loop !== false, mute: item.mute === true,
    volume: Number.isFinite(Number(item.volume)) ? (Number(item.volume) > 2 ? Number(item.volume) / 100 : Number(item.volume)) : 1, objectFit: effectiveAR
  };
  if (window.electronAPI) await window.electronAPI.projectMedia(mediaPayload);
  else if (State.isProjectionOpen) _webPostRenderState('media', mediaPayload);
  const icons = { video:'🎬', audio:'🎵', image:'🖼' };
  toast(`${icons[item.type] || '🖼'} ${item.name}`);
}

// ── Helper: present currently previewed media ─────────────────────────────────
function presentCurrentMedia() {
  const item = (State.media || []).find(m => String(m.id) === String(State.currentMediaId));
  if (item) presentMedia(item);
}

// ── Helper: add currently previewed media to schedule ────────────────────────
function addMediaToQueueCurrent() {
  const item = (State.media || []).find(m => String(m.id) === String(State.currentMediaId));
  if (item) addMediaToQueue(item);
}

// ── Update loop/mute option live ──────────────────────────────────────────────
function updateMediaOption(key, val) {
  const item = (State.media || []).find(m => String(m.id) === String(State.currentMediaId));
  if (!item) return;
  item[key] = val;
  // Re-render the preview video with new settings
  if (item.type === 'video') {
    const v = document.querySelector('#mediaPreviewCanvas video');
    if (v) { v.loop = item.loop !== false; v.muted = true; }
  }
  // Save updated settings
  if (window.electronAPI) window.electronAPI.saveMedia(State.media);
}

// ── Import files ──────────────────────────────────────────────────────────────
// ── Media import progress bar ─────────────────────────────────────────────────
function _showMediaProgress(label, count) {
  const prog  = document.getElementById('mediaImportProgress');
  const lbl   = document.getElementById('mediaImportLabel');
  const cnt   = document.getElementById('mediaImportCount');
  const bar   = document.getElementById('mediaImportBar');
  if (!prog) return;
  if (lbl) lbl.textContent = label || 'Importing media…';
  if (cnt) cnt.textContent = count ? `0 / ${count}` : '';
  if (bar) bar.style.width = '0%';
  prog.style.display = 'block';
  // Disable import button while in progress
  const btn = document.getElementById('addMediaBtn');
  if (btn) { btn.disabled = true; btn.style.opacity = '0.5'; }
}

function _updateMediaProgress(done, total) {
  const bar = document.getElementById('mediaImportBar');
  const cnt = document.getElementById('mediaImportCount');
  if (bar) bar.style.width = total > 0 ? Math.round((done / total) * 100) + '%' : '0%';
  if (cnt) cnt.textContent = `${done} / ${total}`;
}

function _hideMediaProgress(success) {
  const bar  = document.getElementById('mediaImportBar');
  const prog = document.getElementById('mediaImportProgress');
  const btn  = document.getElementById('addMediaBtn');
  if (bar) bar.style.width = success ? '100%' : '0%';
  if (btn) { btn.disabled = false; btn.style.opacity = ''; }
  setTimeout(() => {
    if (prog) prog.style.display = 'none';
    if (bar)  bar.style.width = '0%';
  }, success ? 1000 : 200);
}

// Formats that Chromium/Electron CANNOT play — must be rejected at import
const UNSUPPORTED_FORMATS = new Set(['asf']);

// Human-readable format suggestions shown in the rejection dialog
const FORMAT_SUGGESTIONS = {
  asf: 'MP3, AAC, or MP4',
};

async function addMediaFiles(files) {
  if (State._mediaImporting) return;
  State._mediaImporting = true;
  const acceptedPaths = [];
  const rejected  = [];
  try {

  for (const file of files) {
    const ext  = (file.name.split('.').pop() || '').toLowerCase();
    const type = getMediaType(file.name);

    if (UNSUPPORTED_FORMATS.has(ext)) {
      rejected.push({ name: file.name, ext });
      continue;
    }
    if (!type) continue;
    if (!file.path) {
      rejected.push({ name: file.name, ext, noPath: true });
      continue;
    }
    acceptedPaths.push(file.path);
  }

  if (rejected.length > 0) {
    const lines = rejected.map(r => {
      if (r.noPath) return `• ${r.name}\n  ↳ AnchorCast could not read the source file path for this item.`;
      const suggest = FORMAT_SUGGESTIONS[r.ext] || 'MP4 or MP3';
      return `• ${r.name}\n  ↳ ${r.ext.toUpperCase()} is not supported. Please convert to ${suggest} first.`;
    }).join('\n\n');

    if (window.electronAPI?.showUnsupportedDialog) {
      await window.electronAPI.showUnsupportedDialog({
        title:   'Unsupported Format',
        message: `${rejected.length} file${rejected.length > 1 ? 's were' : ' was'} not imported`,
        detail:  lines + '\n\nTip: Use a free converter like HandBrake, VLC, or Cloudconvert.com to convert your files.',
      });
    } else {
      rejected.forEach(r => {
        const suggest = FORMAT_SUGGESTIONS[r.ext] || 'MP4 or MP3';
        toast(`❌ ${r.name} — ${r.ext ? r.ext.toUpperCase() : 'UNKNOWN'} not supported. Convert to ${suggest} first.`);
      });
    }
  }

  if (acceptedPaths.length === 0) return;

  // Show progress bar now — files are validated and import is starting
  const total = acceptedPaths.length;
  _showMediaProgress(
    `Importing ${total} file${total !== 1 ? 's' : ''}…`,
    total
  );
  // Animate to 50% immediately to show activity during IPC call
  setTimeout(() => _updateMediaProgress(Math.ceil(total * 0.5), total), 80);

  const result = window.electronAPI?.importMediaFiles
    ? await window.electronAPI.importMediaFiles({ files: acceptedPaths })
    : { success: false, error: 'Desktop import API unavailable', items: [] };

  if (!result?.success) {
    _hideMediaProgress(false);
    toast('⚠ Media import failed: ' + (result?.error || 'unknown error'));
    return;
  }

  const imported = (result.items || []).map(m => ({
    ...m,
    mute: m.mute === undefined ? false : m.mute,
    volume: Number.isFinite(Number(m.volume)) ? (Number(m.volume) > 2 ? Number(m.volume) / 100 : Number(m.volume)) : 1,
  }));

  if (!imported.length) {
    _hideMediaProgress(false);
    toast('⚠ No media files were imported');
    return;
  }

  // Fill bar to 100% before finalising
  _updateMediaProgress(total, total);
  await new Promise(r => setTimeout(r, 0));

  State.media = [...(State.media || []), ...imported];
  if (window.electronAPI) await window.electronAPI.saveMedia(State.media);
  await new Promise(r => setTimeout(r, 0));
  renderMediaList(document.getElementById('mediaSearchInput')?.value || '');

  _hideMediaProgress(true);
  toast(`🖼 Imported ${imported.length} item${imported.length !== 1 ? 's' : ''}`);
  return imported;
} finally {
  State._mediaImporting = false;
}
}

async function openMediaFileLocation(item) {
  if (!item?.path) {
    toast('⚠ No file path available for this media');
    return;
  }
  if (!window.electronAPI?.openFileLocation || window.electronAPI?.isWeb) {
    toast(`📂 ${item.path}`);
    return;
  }
  const result = await window.electronAPI.openFileLocation({ filePath: item.path });
  if (!result?.success) {
    toast('⚠ Could not open file location: ' + (result?.error || 'unknown error'));
  }
}

async function importReplacementMediaForType(type) {
  const input = document.getElementById('mediaQuickImportInput');
  if (!input) return;
  const acceptMap = {
    video: '.mp4,.mov,.m4v,.webm,.avi,.mkv,.wmv',
    audio: '.mp3,.wav,.ogg,.flac,.aac,.m4a,.wma,.opus,.aiff',
    image: '.jpg,.jpeg,.png,.gif,.webp,.bmp,.svg,.tiff,.tif'
  };
  input.value = '';
  input.accept = acceptMap[type] || '*/*';
  input.onchange = async () => {
    const files = Array.from(input.files || []);
    if (files.length) await addMediaFiles(files);
  };
  input.click();
}

async function editCurrentMediaProperties() {
  const item = (State.media || []).find(m => String(m.id) === String(document.getElementById('mediaContextMenu')?.dataset.mediaid || State.currentMediaId));
  hideMediaContextMenu();
  if (!item) return;
  State._editingMediaId = item.id;

  const overlay = document.getElementById('mediaPropsOverlay');
  const titleEl = document.getElementById('mediaPropsTitle');
  const subEl = document.getElementById('mediaPropsSubtitle');
  const nameEl = document.getElementById('mediaPropName');
  const aspectEl = document.getElementById('mediaPropAspect');
  const repeatEl = document.getElementById('mediaPropRepeat');
  const muteEl = document.getElementById('mediaPropMute');
  const gainEl = document.getElementById('mediaPropGain');
  const gainLabel = document.getElementById('mediaPropGainLabel');
  const pathEl = document.getElementById('mediaPropsPath');
  const typeChip = document.getElementById('mediaPropsTypeChip');
  const previewInner = document.getElementById('mediaPropsPreviewInner');
  const audioSection = document.getElementById('mediaAudioSection');
  const logoBgEl = document.getElementById('mediaPropLogoBg');
  const stretchEl = document.getElementById('mediaPropStretch');

  if (!overlay || !nameEl) {
    toast('⚠ Media properties dialog is unavailable');
    return;
  }

  titleEl.textContent = `${(item.type || 'media').charAt(0).toUpperCase() + (item.type || 'media').slice(1)} Properties`;
  subEl.textContent = item.name || item.title || 'Configure playback, layout, and audio output';
  nameEl.value = item.title || item.name || '';
  aspectEl.value = item.aspectRatio || 'auto';
  repeatEl.value = item.loop === false ? 'once' : (item.loop ? 'loop' : 'auto');
  muteEl.checked = item.mute === true;
  gainEl.value = Number.isFinite(Number(item.volume)) ? Math.max(0, Math.min(200, Math.round((Number(item.volume) > 2 ? Number(item.volume) / 100 : Number(item.volume)) * 100))) : 100;
  gainLabel.textContent = `${gainEl.value}%`;
  pathEl.textContent = item.path || 'No stored path';
  typeChip.textContent = (item.type || 'media').toUpperCase();
  logoBgEl.checked = !!item.useAsLogoBackground || String(State.settings?.logoMediaId) === String(item.id);
  stretchEl.checked = !!item.stretchToOutput;

  const safePath = String(item.path || '').split('\\').join('/');
  if (item.type === 'image') {
    previewInner.innerHTML = item.path ? `<img src="file://${safePath}" alt="preview">` : '<div><span class="media-icon">🖼</span>No preview</div>';
  } else if (item.type === 'video') {
    previewInner.innerHTML = item.path ? `<video src="file://${safePath}" muted playsinline></video>` : '<div><span class="media-icon">🎬</span>No preview</div>';
    const v = previewInner.querySelector('video');
    if (v) { v.currentTime = 0; v.play().catch(()=>{}); }
  } else if (item.type === 'audio') {
    previewInner.innerHTML = `<div style="padding:18px;width:100%;text-align:center"><span class="media-icon">🎵</span><div style="margin-bottom:10px">${escapeHtml(item.name || item.title || 'Audio file')}</div>${item.path ? `<audio controls src="file://${safePath}"></audio>` : ''}</div>`;
  } else {
    previewInner.innerHTML = `<div><span class="media-icon">📄</span>${escapeHtml(item.name || 'Media item')}</div>`;
  }

  const hasAudio = item.type === 'video' || item.type === 'audio';
  audioSection.style.display = hasAudio ? '' : 'none';

  gainEl.oninput = () => {
    gainLabel.textContent = `${gainEl.value}%`;
  };

  overlay.classList.add('show');
}

function closeMediaPropertiesDialog() {
  document.getElementById('mediaPropsOverlay')?.classList.remove('show');
}

async function saveMediaPropertiesDialog() {
  const item = (State.media || []).find(m => String(m.id) === String(State._editingMediaId));
  if (!item) return closeMediaPropertiesDialog();
  item.title = document.getElementById('mediaPropName')?.value?.trim() || item.title || item.name;
  item.name = item.title;
  item.aspectRatio = document.getElementById('mediaPropAspect')?.value || 'auto';
  const repeat = document.getElementById('mediaPropRepeat')?.value || 'auto';
  item.loop = repeat === 'loop' ? true : (repeat === 'once' ? false : item.loop ?? true);
  const wantsLogo = !!document.getElementById('mediaPropLogoBg')?.checked;
  item.useAsLogoBackground = wantsLogo;
  item.stretchToOutput = !!document.getElementById('mediaPropStretch')?.checked;
  if (item.type === 'video' || item.type === 'audio') {
    item.mute = !!document.getElementById('mediaPropMute')?.checked;
    item.volume = Math.max(0, Math.min(2, Number(document.getElementById('mediaPropGain')?.value || 100) / 100));
  }
  if (wantsLogo && (item.type === 'image' || item.type === 'video')) {
    State.settings.logoMediaId = String(item.id);
    if (window.electronAPI) await window.electronAPI.saveSettings(State.settings, { themeOnly: true });
    _updateLogoBtnState();
  } else if (!wantsLogo && String(State.settings?.logoMediaId) === String(item.id)) {
    State.settings.logoMediaId = null;
    if (window.electronAPI) await window.electronAPI.saveSettings(State.settings, { themeOnly: true });
    _updateLogoBtnState();
  }
  if (window.electronAPI) await window.electronAPI.saveMedia(State.media);
  renderMediaList(document.getElementById('mediaSearchInput')?.value || '');
  closeMediaPropertiesDialog();
  const isCurrentMedia = String(State.currentMediaId) === String(item.id);
  if (isCurrentMedia) previewMedia(item);
  if (isCurrentMedia && State.liveContentType === 'media') presentMedia(item);
  toast('✓ Media properties saved');
}

function previewCurrentMediaFromDialog() {
  const item = (State.media || []).find(m => String(m.id) === String(State._editingMediaId));
  if (!item) return;
  State.currentMediaId = item.id;
  previewMedia(item);
}

function importReplacementForCurrentMedia() {
  const item = (State.media || []).find(m => String(m.id) === String(State._editingMediaId));
  if (!item) { toast('⚠ No media selected'); return; }
  const input = document.getElementById('mediaQuickImportInput');
  if (!input) return;
  const acceptMap = {
    video: '.mp4,.mov,.m4v,.webm,.avi,.mkv,.wmv',
    audio: '.mp3,.wav,.ogg,.flac,.aac,.m4a,.wma,.opus,.aiff',
    image: '.jpg,.jpeg,.png,.gif,.webp,.bmp,.svg,.tiff,.tif'
  };
  input.value = '';
  input.accept = acceptMap[item.type] || '*/*';
  input.onchange = async () => {
    const files = Array.from(input.files || []);
    if (!files.length) return;
    const imported = await addMediaFiles(files);
    if (imported && imported.length > 0) {
      closeMediaPropertiesDialog();
      toast(`📥 Imported ${imported.length} file(s) to library`);
    }
  };
  input.click();
}

function copyCurrentMediaToTheme() {
  const item = (State.media || []).find(m => String(m.id) === String(State._editingMediaId));
  if (!item) { toast('⚠ No media selected'); return; }
  if (item.type !== 'image' && item.type !== 'video') {
    toast('⚠ Only images and videos can be used as theme backgrounds');
    return;
  }
  const category = item.type === 'image' || item.type === 'video' ? 'song' : 'presentation';
  if (window.electronAPI?.openThemeDesigner) {
    window.electronAPI.openThemeDesigner({ category, backgroundMedia: item.path || item.name || '' });
    toast('🎨 Opened Theme Designer with media as background');
  } else {
    toast('⚠ Theme Designer is not available');
  }
}

function escapeHtml(s) {
  return String(s || '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}

function setCurrentMediaLoop(value) {
  const id = document.getElementById('mediaContextMenu')?.dataset.mediaid || State.currentMediaId;
  const item = (State.media || []).find(m => String(m.id) === String(id));
  hideMediaContextMenu();
  if (!item) return;
  item.loop = !!value;
  window.electronAPI?.saveMedia(State.media);
  if (String(State.currentMediaId) === String(item.id)) previewMedia(item);
}

function setCurrentMediaMute(value) {
  const id = document.getElementById('mediaContextMenu')?.dataset.mediaid || State.currentMediaId;
  const item = (State.media || []).find(m => String(m.id) === String(id));
  hideMediaContextMenu();
  if (!item) return;
  item.mute = !!value;
  window.electronAPI?.saveMedia(State.media);
  if (String(State.currentMediaId) === String(item.id)) previewMedia(item);
}

async function removeCurrentMedia() {
  const id = document.getElementById('mediaContextMenu')?.dataset.mediaid || State.currentMediaId;
  hideMediaContextMenu();
  if (!id) return;
  await removeMedia(id);
  toast('🗑 Media removed from library');
}

function showMediaContextMenu(e, item) {
  const menu = document.getElementById('mediaContextMenu');
  if (!menu) return;
  if (item?.id != null) menu.dataset.mediaid = String(item.id);
  menu.style.display = 'block';
  const x = Math.min(e.clientX, window.innerWidth - 220);
  const y = Math.min(e.clientY, window.innerHeight - 160);
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';
  setTimeout(() => document.addEventListener('click', hideMediaContextMenu, { once: true }), 10);
}

function hideMediaContextMenu() {
  const menu = document.getElementById('mediaContextMenu');
  if (menu) menu.style.display = 'none';
}

async function setCurrentMediaAsBgLogo() {
  const id = document.getElementById('mediaContextMenu')?.dataset.mediaid || State.currentMediaId;
  hideMediaContextMenu();
  const item = (State.media || []).find(m => String(m.id) === String(id));
  if (!item) { toast('⚠ No media selected'); return; }
  if (item.type !== 'image' && item.type !== 'video') {
    toast('⚠ Only images and videos can be used as a background');
    return;
  }
  State.settings.logoMediaId = String(item.id);
  if (window.electronAPI) await window.electronAPI.saveSettings(State.settings, { themeOnly: true });
  toast(`🏷 "${item.title || item.name}" set as background`);
}
function setCurrentMediaAsLogo() { return setCurrentMediaAsBgLogo(); }

function _updateBgLogoBtnState() {
  const btn = document.getElementById('bgLogoBtn');
  if (!btn) return;
  if (State._logoActive) {
    btn.classList.add('logo-active');
    btn.title = 'Background is showing — click to turn off';
  } else {
    btn.classList.remove('logo-active');
    const logoItem = _getLogoMediaItem();
    btn.title = logoItem ? `Show background (${logoItem.title || logoItem.name})` : 'Show background (none set — right-click a media item to set one)';
  }
}
function _updateLogoBtnState() { _updateBgLogoBtnState(); }

function _getLogoMediaItem() {
  const logoId = State.settings?.logoMediaId;
  if (!logoId) return null;
  return (State.media || []).find(m => String(m.id) === String(logoId)) || null;
}

async function toggleBgLogo() {
  if (State._logoActive) {
    clearLive();
    State._logoActive = false;
    _updateBgLogoBtnState();
    return;
  }
  const item = _getLogoMediaItem();
  if (!item) {
    toast('⚠ No background set — right-click a media item in the Media tab and choose "Set as Background"');
    return;
  }
  State._logoActive = true;
  State.liveContentType = 'media';
  State.liveVerse = null;
  State.currentMediaId = item.id;
  if (!State.isLive) _setLiveOn();

  const liveDisplay = document.getElementById('liveDisplay');
  const liveEmpty = document.getElementById('liveEmpty');
  if (liveDisplay) {
    const oldMedia = liveDisplay.querySelector('video,audio');
    if (oldMedia) { oldMedia.pause(); oldMedia.src = ''; }
    liveDisplay.innerHTML = '';
    liveDisplay.style.padding = '0';
    liveDisplay.style.position = 'relative';
    const fitMode = _mediaObjectFit(item);
    if (item.type === 'video') {
      const vid = document.createElement('video');
      vid.style.cssText = `position:absolute;inset:0;width:100%;height:100%;object-fit:${fitMode};display:block`;
      vid.loop = true;
      vid.muted = true;
      vid.autoplay = true;
      vid.setAttribute('playsinline', '');
      vid.src = toMediaUrl(item.path);
      liveDisplay.appendChild(vid);
      vid.load();
      vid.play().catch(() => {});
    } else {
      const img = document.createElement('img');
      img.style.cssText = `position:absolute;inset:0;width:100%;height:100%;object-fit:${fitMode};display:block`;
      img.src = toMediaUrl(item.path);
      liveDisplay.appendChild(img);
    }
    liveDisplay.style.display = 'flex';
    if (liveEmpty) liveEmpty.style.display = 'none';
  }

  const logoFit = _mediaObjectFit(item);
  const mediaPayload = {
    id: item.id, name: item.title || item.name, path: item.path,
    type: item.type, loop: true, mute: item.mute === true,
    volume: 1, objectFit: logoFit
  };
  if (window.electronAPI) await window.electronAPI.projectMedia(mediaPayload);
  else if (State.isProjectionOpen) _webPostRenderState('media', mediaPayload);

  _updateBgLogoBtnState();
  toast(`🏷 Background: ${item.title || item.name}`);
}
function toggleLogo() { return toggleBgLogo(); }

// ── Logo Overlay (new draggable/resizable logo watermark) ──────────────────
State._logoOverlay = {
  src: null, type: null, fileName: null,
  position: 'bottom-right',
  xPct: 85, yPct: 85,
  sizePct: 20, opacity: 100,
  visible: false
};

function toggleLogoOverlayPopup() {
  const popup = document.getElementById('logoOverlayPopup');
  if (!popup) return;
  popup.style.flexDirection = 'column';
  _toggleLivePopup('logoOverlayPopup');
}

function pickLogoOverlayFile() {
  const inp = document.createElement('input');
  inp.type = 'file';
  inp.accept = 'image/*,video/*';
  inp.onchange = () => {
    const f = inp.files?.[0];
    if (!f) return;
    const url = URL.createObjectURL(f);
    const isVideo = f.type.startsWith('video');
    State._logoOverlay.src = url;
    State._logoOverlay.type = isVideo ? 'video' : 'image';
    State._logoOverlay.fileName = f.name;
    State._logoOverlay._file = f;
    const nameEl = document.getElementById('logoOverlayFileName');
    if (nameEl) nameEl.textContent = f.name;
    toast(`Logo: ${f.name} loaded`);
  };
  inp.click();
}

function _logoPositionToCoords(pos) {
  const pad = 5;
  switch(pos) {
    case 'top-left':     return { xPct: pad, yPct: pad };
    case 'top-right':    return { xPct: 100 - pad, yPct: pad };
    case 'bottom-left':  return { xPct: pad, yPct: 100 - pad };
    case 'bottom-right': return { xPct: 100 - pad, yPct: 100 - pad };
    case 'center':       return { xPct: 50, yPct: 50 };
    default:             return null;
  }
}

function setLogoOverlayPosition(pos) {
  State._logoOverlay.position = pos;
  const coords = _logoPositionToCoords(pos);
  if (coords) {
    State._logoOverlay.xPct = coords.xPct;
    State._logoOverlay.yPct = coords.yPct;
  }
  document.querySelectorAll('.logo-pos-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.pos === pos);
  });
  if (State._logoOverlay.visible) _sendLogoOverlay();
}

function updateLogoOverlay() {
  const sizeSlider = document.getElementById('logoOverlaySize');
  const opacitySlider = document.getElementById('logoOverlayOpacity');
  const sizeVal = document.getElementById('logoOverlaySizeVal');
  const opacityVal = document.getElementById('logoOverlayOpacityVal');
  if (sizeSlider) {
    State._logoOverlay.sizePct = parseInt(sizeSlider.value, 10);
    if (sizeVal) sizeVal.textContent = sizeSlider.value + '%';
  }
  if (opacitySlider) {
    State._logoOverlay.opacity = parseInt(opacitySlider.value, 10);
    if (opacityVal) opacityVal.textContent = opacitySlider.value + '%';
  }
  if (State._logoOverlay.visible) _sendLogoOverlay();
}

function toggleLogoOverlayVisibility() {
  if (!State._logoOverlay.src) {
    toast('⚠ Choose a logo file first');
    return;
  }
  State._logoOverlay.visible = !State._logoOverlay.visible;
  _sendLogoOverlay();
  _updateLogoOverlayBtn();
}

function hideLogoOverlay() {
  State._logoOverlay.visible = false;
  _sendLogoOverlay();
  _updateLogoOverlayBtn();
}

function _updateLogoOverlayBtn() {
  const btn = document.getElementById('logoOverlayBtn');
  if (!btn) return;
  if (State._logoOverlay.visible) {
    btn.classList.add('logo-active');
    btn.title = 'Logo overlay is showing — click to toggle popup';
  } else {
    btn.classList.remove('logo-active');
    btn.title = 'Logo overlay';
  }
}

function _sendLogoOverlay() {
  const ov = State._logoOverlay;
  const payload = {
    visible: ov.visible,
    src: ov.src || null,
    type: ov.type || 'image',
    xPct: ov.xPct,
    yPct: ov.yPct,
    sizePct: ov.sizePct,
    opacity: ov.opacity / 100,
    position: ov.position,
    draggable: ov.position === 'custom'
  };
  if (window.electronAPI?.sendLogoOverlay) {
    window.electronAPI.sendLogoOverlay(payload);
  } else if (State.isProjectionOpen) {
    _webPostRenderState('logo-overlay', payload);
  }
  _updateLiveLogoPreview(payload);
}

function _updateLiveLogoPreview(payload) {
  const overlay = document.getElementById('liveLogoOverlay');
  const img = document.getElementById('liveLogoImg');
  if (!overlay || !img) return;
  if (!payload || !payload.visible || !payload.src) {
    overlay.style.display = 'none';
    return;
  }
  const canvas = document.getElementById('liveCanvas');
  if (!canvas) return;
  const cw = canvas.offsetWidth;
  const ch = canvas.offsetHeight;
  if (img.src !== payload.src) img.src = payload.src;
  const size = Math.round(cw * payload.sizePct / 100);
  overlay.style.width = size + 'px';
  overlay.style.height = size + 'px';
  overlay.style.opacity = payload.opacity;
  const anchorX = cw * payload.xPct / 100;
  const anchorY = ch * payload.yPct / 100;
  let left, top;
  const pos = payload.position;
  if (pos === 'top-left') { left = anchorX; top = anchorY; }
  else if (pos === 'top-right') { left = anchorX - size; top = anchorY; }
  else if (pos === 'bottom-left') { left = anchorX; top = anchorY - size; }
  else if (pos === 'bottom-right') { left = anchorX - size; top = anchorY - size; }
  else if (pos === 'center') { left = anchorX - size / 2; top = anchorY - size / 2; }
  else { left = anchorX - size / 2; top = anchorY - size / 2; }
  overlay.style.left = Math.max(0, Math.min(cw - size, left)) + 'px';
  overlay.style.top = Math.max(0, Math.min(ch - size, top)) + 'px';
  overlay.style.display = 'block';
  _initLiveLogoDrag();
}

let _liveLogoDrag = null;
let _liveLogoResize = null;

function _initLiveLogoDrag() {
  const overlay = document.getElementById('liveLogoOverlay');
  const handle = document.getElementById('liveLogoResizeHandle');
  if (!overlay || !handle || overlay._dragInit) return;
  overlay._dragInit = true;

  overlay.addEventListener('mouseenter', () => { handle.style.opacity = '1'; });
  overlay.addEventListener('mouseleave', () => { if (!_liveLogoResize) handle.style.opacity = '0'; });

  overlay.addEventListener('mousedown', (e) => {
    if (e.target === handle) return;
    e.preventDefault();
    overlay.style.cursor = 'grabbing';
    _liveLogoDrag = {
      startX: e.clientX, startY: e.clientY,
      origLeft: overlay.offsetLeft, origTop: overlay.offsetTop
    };
  });

  handle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    e.stopPropagation();
    _liveLogoResize = {
      startX: e.clientX, startY: e.clientY,
      origW: overlay.offsetWidth
    };
  });

  document.addEventListener('mousemove', (e) => {
    const canvas = document.getElementById('liveCanvas');
    if (!canvas) return;
    const cw = canvas.offsetWidth;
    const ch = canvas.offsetHeight;

    if (_liveLogoDrag) {
      const dx = e.clientX - _liveLogoDrag.startX;
      const dy = e.clientY - _liveLogoDrag.startY;
      const w = overlay.offsetWidth;
      const h = overlay.offsetHeight;
      overlay.style.left = Math.max(0, Math.min(cw - w, _liveLogoDrag.origLeft + dx)) + 'px';
      overlay.style.top = Math.max(0, Math.min(ch - h, _liveLogoDrag.origTop + dy)) + 'px';
    }
    if (_liveLogoResize) {
      const dx = e.clientX - _liveLogoResize.startX;
      const dy = e.clientY - _liveLogoResize.startY;
      const delta = Math.max(dx, dy);
      const maxSize = Math.min(cw, ch);
      const newSize = Math.max(15, Math.min(maxSize, _liveLogoResize.origW + delta));
      overlay.style.width = newSize + 'px';
      overlay.style.height = newSize + 'px';
    }
  });

  document.addEventListener('mouseup', () => {
    const canvas = document.getElementById('liveCanvas');
    if (!canvas) return;
    const cw = canvas.offsetWidth;
    const ch = canvas.offsetHeight;
    let changed = false;

    if (_liveLogoDrag) {
      overlay.style.cursor = 'grab';
      const w = overlay.offsetWidth;
      const cx = overlay.offsetLeft + w / 2;
      const cy = overlay.offsetTop + w / 2;
      State._logoOverlay.xPct = (cx / cw) * 100;
      State._logoOverlay.yPct = (cy / ch) * 100;
      State._logoOverlay.position = 'custom';
      _liveLogoDrag = null;
      changed = true;
    }
    if (_liveLogoResize) {
      const newSize = Math.min(overlay.offsetWidth, cw, ch);
      overlay.style.width = newSize + 'px';
      overlay.style.height = newSize + 'px';
      State._logoOverlay.sizePct = (newSize / cw) * 100;
      const cx = overlay.offsetLeft + newSize / 2;
      const cy = overlay.offsetTop + newSize / 2;
      State._logoOverlay.xPct = (cx / cw) * 100;
      State._logoOverlay.yPct = (cy / ch) * 100;
      State._logoOverlay.position = 'custom';
      _liveLogoResize = null;
      changed = true;
    }
    if (changed) {
      const sizeSlider = document.getElementById('logoOverlaySize');
      const sizeVal = document.getElementById('logoOverlaySizeVal');
      if (sizeSlider) sizeSlider.value = Math.round(State._logoOverlay.sizePct);
      if (sizeVal) sizeVal.textContent = Math.round(State._logoOverlay.sizePct) + '%';
      document.querySelectorAll('.logo-pos-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.pos === State._logoOverlay.position);
      });
      _sendLogoOverlay();
    }
  });
}

function _handleLogoOverlayDragUpdate(data) {
  if (data && typeof data.xPct === 'number') {
    State._logoOverlay.xPct = data.xPct;
    State._logoOverlay.yPct = data.yPct;
    if (typeof data.sizePct === 'number') {
      State._logoOverlay.sizePct = data.sizePct;
      const sizeSlider = document.getElementById('logoOverlaySize');
      const sizeVal = document.getElementById('logoOverlaySizeVal');
      if (sizeSlider) sizeSlider.value = Math.round(data.sizePct);
      if (sizeVal) sizeVal.textContent = Math.round(data.sizePct) + '%';
    }
    State._logoOverlay.position = data.position || 'custom';
    document.querySelectorAll('.logo-pos-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.pos === State._logoOverlay.position);
    });
    _updateLiveLogoPreview({
      visible: State._logoOverlay.visible,
      src: State._logoOverlay.src,
      type: State._logoOverlay.type,
      xPct: State._logoOverlay.xPct,
      yPct: State._logoOverlay.yPct,
      sizePct: State._logoOverlay.sizePct,
      opacity: State._logoOverlay.opacity / 100,
      position: State._logoOverlay.position
    });
  }
}

if (window.electronAPI?.on) {
  window.electronAPI.on('logo-overlay-drag-update', _handleLogoOverlayDragUpdate);
}

async function openCurrentMediaFileLocation() {
  const id = State._editingMediaId || document.getElementById('mediaContextMenu')?.dataset.mediaid || State.currentMediaId;
  const item = (State.media || []).find(m => String(m.id) === String(id));
  hideMediaContextMenu();
  if (item) await openMediaFileLocation(item);
  else toast('⚠ No media selected');
}


async function removeMedia(id) {
  const item = (State.media || []).find(m => String(m.id) === String(id));
  const name = item?.title || item?.name || 'this item';
  _confirmModal(`Delete "${name}"? This cannot be undone.`, async () => {
    State.media = (State.media || []).filter(m => String(m.id) !== String(id));
    if (String(State.currentMediaId) === String(id)) State.currentMediaId = null;
    State.queue = State.queue.filter(q => !(q.type === 'media' && String(q.mediaId) === String(id)));
    if (window.electronAPI) await window.electronAPI.saveMedia(State.media);
    renderMediaList();
    renderQueue();
    toast(`🗑 Deleted: "${name}"`);
  });
}

function addMediaToQueue(item) {
  if (State.queue.find(q => q.type === 'media' && String(q.mediaId) === String(item.id))) {
    toast(`"${item.name}" already in schedule`); return;
  }
  State.queue.push({ type:'media', mediaId:item.id, name:item.name, mediaType:item.type });
  renderQueue();
  toast(`📋 Added: ${item.name}`);
}

// ─── PRESENTATIONS ENGINE ──────────────────────────────────────────────────────
// Two types of presentations:
//   1. IMPORTED  — PPTX/PDF converted to PNG images (read-only)
//   2. CREATED   — Built in the editor, stored as slide objects in presentations.json
//
// State.presentations = [{ id, name, type:'imported'|'created', slideCount, slides?, outDir? }]

// ── Load & render library strip ───────────────────────────────────────────────
async function loadPresentations() {
  // Reset cache so we always get fresh data from disk on load
  _createdPresCache = null;
  try {
    const list = window.electronAPI ? await window.electronAPI.getPresentations() : [];
    State.presentations = list || [];
  } catch(e) { State.presentations = []; }
  // Merge in locally-created presentations from assets/Data/
  const local = await _getLocalPresentations();
  local.forEach(lp => {
    if (!State.presentations.find(p => p.id === lp.id)) State.presentations.push(lp);
  });
  renderPresLibStrip();
}

function renderPresLibStrip(filter = '') {
  const panel = document.getElementById('presListPanel');
  const empty  = document.getElementById('presLibEmpty');
  if (!panel) return;

  panel.querySelectorAll('.pres-row').forEach(e => e.remove());

  const list = State.presentations.filter(p =>
    !filter || p.name?.toLowerCase().includes(filter.toLowerCase())
  );

  if (!list.length) {
    if (empty) empty.style.display = '';
    return;
  }
  if (empty) empty.style.display = 'none';

  list.forEach(pres => {
    const isActive = String(pres.id) === String(State.currentPresId);

    // Thumbnail — first slide image for imported, gradient for created
    const imgUrl = pres.type === 'imported' && pres.outDir
      ? toMediaUrl(pres.outDir.replace(/\\/g, '/') + '/slide-1.png')
      : '';
    const bgColor = pres.slides?.[0]?.bg || '#0a0a1e';

    const row = document.createElement('div');
    row.className = 'pres-row';
    row.dataset.presid = String(pres.id);
    row.style.cssText = `
      display:flex;align-items:center;gap:8px;padding:7px 8px;
      border-bottom:1px solid var(--border);cursor:pointer;
      border-left:3px solid ${isActive ? 'var(--gold)' : 'transparent'};
      background:${isActive ? 'rgba(201,168,76,.06)' : 'transparent'};
      transition:background .1s;`;

    row.innerHTML = `
      <div style="width:56px;height:32px;flex-shrink:0;border-radius:3px;overflow:hidden;
        border:1px solid ${isActive ? 'var(--gold)' : 'var(--border)'};
        background:${imgUrl ? `url('${imgUrl}') center/cover no-repeat` : bgColor}"></div>
      <div style="flex:1;min-width:0">
        <div style="font-size:11px;font-weight:600;
          color:${isActive ? 'var(--gold)' : 'var(--text)'};
          white-space:nowrap;overflow:hidden;text-overflow:ellipsis"
          title="${escapeHtml(pres.name)}">${escapeHtml(pres.name)}</div>
        <div style="font-size:9px;color:var(--text-dim);margin-top:1px">
          ${pres.slideCount || 0} slides · ${pres.type === 'imported' ? 'Imported' : 'Custom'}
        </div>
      </div>
      <div style="display:flex;gap:3px;flex-shrink:0">
        <button class="btn-sm" onclick="event.stopPropagation();newPresentation()"
          style="font-size:9px;padding:2px 6px;color:var(--gold);border-color:var(--gold-dim)" title="New Presentation">+ New</button>
        <button class="btn-sm" onclick="event.stopPropagation();_editPres('${pres.id}')"
          style="font-size:9px;padding:2px 5px" title="Edit">✏</button>
        <button class="btn-sm" onclick="event.stopPropagation();_deletePres('${pres.id}')"
          style="font-size:9px;padding:2px 5px;color:var(--live);border-color:rgba(224,82,82,.3)" title="Delete">🗑</button>
      </div>`;

    row.addEventListener('click', () => selectPresentation(pres.id));
    row.addEventListener('contextmenu', e => {
      e.preventDefault();
      e.stopPropagation();
      State.currentPresId = pres.id;
      showPresContextMenu(e);
    });
    row.addEventListener('mouseover', () => {
      if (!isActive) row.style.background = 'var(--card)';
    });
    row.addEventListener('mouseout', () => {
      if (!isActive) row.style.background = 'transparent';
    });

    panel.appendChild(row);
  });
}

function _editPres(id) {
  State.currentPresId = id;
  editCurrentPresentation();
}

function _deletePres(id) {
  State.currentPresId = id;
  deleteCurrentPresentation();
}
async function selectPresentation(id) {
  State.currentPresId = id;
  State.currentPresSlideIdx = 0;

  // Highlight in list
  document.querySelectorAll('.pres-row').forEach(el => {
    const active = el.dataset.presid === String(id);
    el.style.borderLeft  = active ? '3px solid var(--gold)' : '3px solid transparent';
    el.style.background  = active ? 'rgba(201,168,76,.06)' : 'transparent';
  });

  const pres = State.presentations.find(p => String(p.id) === String(id));
  if (!pres) return;

  let slides = [];

  if (pres.type === 'imported') {
    const result = window.electronAPI
      ? await window.electronAPI.getPresentationSlides({ id })
      : { success: false, slides: [] };
    if (result.success) slides = result.slides;
  } else {
    // Created presentation — slides are stored locally
    slides = (pres.slides || []).map((s, i) => ({ index: i, slide: s }));
  }

  State.currentPresSlides = slides;
  State.currentPresSlideIdx = 0;

  // Show the presentation slides in Program Preview
  if (slides.length) previewPresSlide(0);
  _showPresPreviewArea(pres);

  const pt = document.getElementById('previewPanelTitle');
  if (pt) pt.textContent = `📑 ${pres.name}`;
}

// ── Preview slide in Program Preview ─────────────────────────────────────────
function previewPresSlide(idx) {
  State.currentPresSlideIdx = idx;

  const pres  = State.presentations.find(p => String(p.id) === String(State.currentPresId));
  const slide = State.currentPresSlides[idx];
  if (!slide) return;

  // Update the thumbnail grid in presPreviewArea (highlight active, rebuild if needed)
  const presPA = document.getElementById('presPreviewArea');
  if (presPA && presPA.style.display !== 'none') {
    _updatePresPreviewArea(pres, idx);
  }

  // Update Live Display (right panel) — always sync
  const ld = document.getElementById('liveDisplay');
  const le = document.getElementById('liveEmpty');
  if (ld) {
    ld.style.padding = '0';
    ld.style.position = 'relative';
    if (pres?.type === 'imported') {
      ld.innerHTML = `<img src="${toMediaUrl(slide.imagePath)}"
        style="position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:#000">`;
    } else {
      const s = slide.slide || {};
      _renderCreatedSlideInto(ld, s, false);
    }
    ld.style.display = 'flex';
    if (le) le.style.display = 'none';
  }
}

// Render a created slide's objects[] into a container element
function _renderCreatedSlideInto(container, s, isPreview) {
  const CW = 1920, CH = 1080;
  container.style.background = s.bg || '#0a0a1e';
  container.style.backgroundImage = s.bgImage ? `url('${toMediaUrl(s.bgImage)}')` : 'none';
  container.style.backgroundSize = 'cover';
  container.style.backgroundPosition = 'center';
  container.style.position = 'relative';
  container.style.overflow = 'hidden';

  // Remove old rendered content
  container.querySelectorAll('.pres-obj-render').forEach(e => e.remove());
  if (!container.offsetWidth) return; // not yet in DOM

  const cw = container.offsetWidth;
  const ch = container.offsetHeight;
  const sc = Math.min(cw / CW, ch / CH);
  const ox = (cw - CW * sc) / 2;
  const oy = (ch - CH * sc) / 2;

  (s.objects || []).forEach(obj => {
    const el = document.createElement('div');
    el.className = 'pres-obj-render';
    el.style.position   = 'absolute';
    el.style.left       = (ox + obj.x * sc) + 'px';
    el.style.top        = (oy + obj.y * sc) + 'px';
    el.style.width      = (obj.w * sc) + 'px';
    el.style.height     = (obj.h * sc) + 'px';
    el.style.transform  = `rotate(${obj.rotation || 0}deg)`;
    el.style.overflow   = 'hidden';
    el.style.boxSizing  = 'border-box';

    if (obj.type === 'text') {
      const alpha = (obj.textBgAlpha || 0) / 100;
      if (alpha > 0) {
        const r = parseInt((obj.textBg||'#000').slice(1,3),16)||0;
        const g = parseInt((obj.textBg||'#000').slice(3,5),16)||0;
        const b = parseInt((obj.textBg||'#000').slice(5,7),16)||0;
        el.style.background = `rgba(${r},${g},${b},${alpha})`;
      }
      el.style.display        = 'flex';
      el.style.alignItems     = 'center';
      el.style.justifyContent = obj.textAlign === 'left' ? 'flex-start'
        : obj.textAlign === 'right' ? 'flex-end' : 'center';
      const t = document.createElement('div');
      t.style.fontFamily  = `'${obj.font || 'Cinzel'}', serif`;
      t.style.fontSize    = ((obj.fontSize || 48) * sc) + 'px';
      t.style.color       = obj.textColor || '#fff';
      t.style.fontWeight  = String(Number(obj.fontWeight || (obj.bold ? 700 : 400)));
      t.style.fontStyle   = obj.italic ? 'italic' : 'normal';
      t.style.textAlign   = obj.textAlign || 'center';
      t.style.textShadow  = obj.shadow === false ? 'none' : `${Number(obj.shadowOffsetX || 0) * sc}px ${Number(obj.shadowOffsetY || 2) * sc}px ${Math.max(0, Number(obj.shadowBlur ?? 8)) * sc}px ${obj.shadowColor || '#000000'}`;
      t.style.whiteSpace  = 'pre-wrap';
      t.style.lineHeight  = String(obj.lineSpacing || 1.25);
      t.style.letterSpacing = (Number(obj.letterSpacing || 0) * sc) + 'px';
      t.style.textTransform = obj.textTransform || 'none';
      t.style.width       = '100%';
      t.textContent = obj.text || '';
      el.appendChild(t);

    } else if (obj.type === 'shape') {
      el.style.background   = obj.shapeFill || '#3b82f6';
      el.style.opacity      = String((obj.shapeOpacity ?? 80) / 100);
      el.style.borderRadius = obj.shape === 'ellipse' ? '50%' : ((obj.shapeRadius || 0) + '%');
      const bw = (obj.shapeBorderW || 0) * sc;
      if (bw > 0) el.style.border = `${bw}px solid ${obj.shapeBorder || '#fff'}`;

    } else if (obj.type === 'line') {
      const bw = Math.max(1, (obj.shapeBorderW || 2) * sc);
      el.style.borderTop   = `${bw}px solid ${obj.shapeFill || '#fff'}`;
      el.style.opacity     = String((obj.shapeOpacity ?? 80) / 100);
      el.style.height      = `${bw}px`;
      el.style.top         = (oy + (obj.y + obj.h / 2) * sc) + 'px';

    } else if (obj.type === 'svg') {
      // SVG shapes rendered at actual element pixel size
      el.style.overflow = 'visible';
      el.style.opacity  = String((obj.shapeOpacity ?? 90) / 100);
      const fill   = obj.shapeFill  || '#c9a84c';
      const stroke = (obj.shapeBorderW||0) > 0 ? obj.shapeBorder||'#fff' : 'none';
      const sw_px  = parseFloat(el.style.width)  || obj.w * sc;
      const sh_px  = parseFloat(el.style.height) || obj.h * sc;
      el.innerHTML = _svgShapePath(obj.shape, sw_px, sh_px, fill, stroke, (obj.shapeBorderW||0)*sc);

    } else if (obj.type === 'image') {
      const img = document.createElement('img');
      img.src = toMediaUrl(obj.src || '');
      img.style.width       = '100%';
      img.style.height      = '100%';
      img.style.objectFit   = obj.imgFit || 'contain';
      img.style.opacity     = String((obj.imgOpacity ?? 100) / 100);
      img.style.display     = 'block';
      img.style.pointerEvents = 'none';
      el.appendChild(img);
    }
    container.appendChild(el);
  });
}

// Shared SVG path generator used by editor, preview and projection
function _svgShapePath(shape, w, h, fill, stroke, strokeW) {
  const cx = w/2, cy = h/2;
  let pathD = '';
  if (shape === 'star') {
    const pts=5, or=Math.min(w,h)/2, ir=or*0.42;
    let d=''; for(let p=0;p<pts*2;p++){const r=p%2===0?or:ir;const a=(p*Math.PI/pts)-Math.PI/2;d+=(p===0?'M':'L')+(cx+r*Math.cos(a)).toFixed(2)+','+(cy+r*Math.sin(a)).toFixed(2);} pathD=d+'Z';
  } else if (shape === 'triangle')  { pathD=`M${cx},0 L${w},${h} L0,${h} Z`; }
  else if (shape === 'diamond')     { pathD=`M${cx},0 L${w},${cy} L${cx},${h} L0,${cy} Z`; }
  else if (shape === 'cross')       { const t=w*.3,u=h*.3; pathD=`M${(w-t)/2},0 L${(w+t)/2},0 L${(w+t)/2},${(h-u)/2} L${w},${(h-u)/2} L${w},${(h+u)/2} L${(w+t)/2},${(h+u)/2} L${(w+t)/2},${h} L${(w-t)/2},${h} L${(w-t)/2},${(h+u)/2} L0,${(h+u)/2} L0,${(h-u)/2} L${(w-t)/2},${(h-u)/2} Z`; }
  else if (shape === 'arrow')       { const aw=w*.4,ah=h*.5; pathD=`M0,${(h-ah)/2} L${w-aw},${(h-ah)/2} L${w-aw},0 L${w},${cy} L${w-aw},${h} L${w-aw},${(h+ah)/2} L0,${(h+ah)/2} Z`; }
  else if (shape === 'hexagon')     { const r=Math.min(w,h)/2; let d=''; for(let i=0;i<6;i++){const a=i*Math.PI/3-Math.PI/6;d+=(i===0?'M':'L')+(cx+r*Math.cos(a)).toFixed(2)+','+(cy+r*Math.sin(a)).toFixed(2);} pathD=d+'Z'; }
  else if (shape === 'heart')       { pathD=`M${cx},${h*.3} C${cx},${h*.1} 0,${h*.1} 0,${h*.4} C0,${h*.7} ${cx},${h*.9} ${cx},${h} C${cx},${h*.9} ${w},${h*.7} ${w},${h*.4} C${w},${h*.1} ${cx},${h*.1} ${cx},${h*.3}`; }
  else if (shape === 'octagon')     { const r=Math.min(w,h)/2; let d=''; for(let i=0;i<8;i++){const a=i*Math.PI/4-Math.PI/8;d+=(i===0?'M':'L')+(cx+r*Math.cos(a)).toFixed(2)+','+(cy+r*Math.sin(a)).toFixed(2);} pathD=d+'Z'; }
  else if (shape === 'rhombus')     { pathD=`M${cx},0 L${w},${cy} L${cx},${h} L0,${cy} Z`; }
  else if (shape === 'speech')      { pathD=`M0,0 L${w},0 L${w},${h*.7} L${w*.4},${h*.7} L${w*.25},${h} L${w*.2},${h*.7} L0,${h*.7} Z`; }
  else if (shape === 'cloud')       { pathD=`M${w*.25},${h*.8} Q${w*.05},${h*.8} ${w*.05},${h*.55} Q${w*.05},${h*.3} ${w*.25},${h*.3} Q${w*.3},${h*.05} ${w*.55},${h*.1} Q${w*.75},${h*.05} ${w*.8},${h*.3} Q${w},${h*.3} ${w*.95},${h*.55} Q${w*.95},${h*.8} ${w*.75},${h*.8} Z`; }
  else if (shape === 'banner')      { pathD=`M0,${h*.2} Q${w*.05},0 ${w*.1},${h*.2} L${w*.9},${h*.2} Q${w*.95},0 ${w},${h*.2} L${w},${h*.8} Q${w*.95},${h} ${w*.9},${h*.8} L${w*.1},${h*.8} Q${w*.05},${h} 0,${h*.8} Z`; }
  else { pathD=`M0,0 L${w},0 L${w},${h} L0,${h} Z`; } // rect fallback

  return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;overflow:visible"><path d="${pathD}" fill="${fill}" stroke="${stroke||'none'}" stroke-width="${strokeW||0}"/></svg>`;
}

// Populate the Program Preview panel with all slides of the selected presentation
function _showPresPreviewArea(pres) {
  ['versePreviewArea','bibleSlideArea','songPreviewArea','mediaPreviewArea'].forEach(id => {
    const el = document.getElementById(id); if (el) el.style.display = 'none';
  });
  const presPA = document.getElementById('presPreviewArea');
  if (presPA) presPA.style.display = 'flex';
  const pt = document.getElementById('previewPanelTitle');
  if (pt) pt.textContent = '\u{1F4D1} ' + (pres?.name || 'Presentation');
  _updatePresPreviewArea(pres, State.currentPresSlideIdx ?? 0);
}

function _updatePresPreviewArea(pres, activeIdx) {
  const titleEl = document.getElementById('presPreviewTitle');
  const strip   = document.getElementById('presPreviewStrip');
  const navPos  = document.getElementById('presNavPos');
  if (!strip) return;

  if (titleEl) titleEl.textContent = pres?.name || '';
  const total = State.currentPresSlides.length;
  if (navPos) navPos.textContent = total ? 'Slide '+(activeIdx+1)+' of '+total : '';

  strip.innerHTML = '';
  if (!total) {
    const e = document.getElementById('presPreviewEmpty');
    if (e) { strip.appendChild(e); e.style.display = ''; }
    return;
  }

  State.currentPresSlides.forEach((slide, i) => {
    const isActive = i === activeIdx;
    const bg = slide.slide?.bg || '#0a0a1e';
    const box = document.createElement('div');
    box.style.cssText =
      'position:relative;border-radius:5px;overflow:hidden;cursor:pointer;aspect-ratio:16/9;background:'+bg+';'+
      'border:2px solid '+(isActive?'var(--gold)':'var(--border)')+';'+
      'box-shadow:'+(isActive?'0 0 0 2px rgba(201,168,76,.25)':'none')+';'+
      'transition:border-color .12s,transform .1s;';

    if (pres?.type === 'imported') {
      box.style.backgroundImage = "url('"+toMediaUrl(slide.imagePath)+"')";
      box.style.backgroundSize  = 'cover';
      box.style.backgroundPosition = 'center';
    } else {
      const s = slide.slide || {};
      if (s.bgImage) { box.style.backgroundImage="url('"+toMediaUrl(s.bgImage)+"')"; box.style.backgroundSize='cover'; }

      // Render ALL objects proportionally as mini preview
      const CW = 1920, CH = 1080;
      (s.objects || []).forEach(obj => {
        const el = document.createElement('div');
        el.style.position   = 'absolute';
        el.style.left       = ((obj.x / CW) * 100) + '%';
        el.style.top        = ((obj.y / CH) * 100) + '%';
        el.style.width      = ((obj.w / CW) * 100) + '%';
        el.style.height     = ((obj.h / CH) * 100) + '%';
        el.style.transform  = 'rotate('+(obj.rotation||0)+'deg)';
        el.style.overflow   = 'hidden';
        el.style.boxSizing  = 'border-box';

        if (obj.type === 'text') {
          el.style.display = 'flex';
          el.style.alignItems = 'center';
          el.style.justifyContent = obj.textAlign==='left'?'flex-start':obj.textAlign==='right'?'flex-end':'center';
          const alpha = (obj.textBgAlpha||0)/100;
          if (alpha > 0) {
            const r=parseInt((obj.textBg||'#000').slice(1,3),16)||0;
            const g=parseInt((obj.textBg||'#000').slice(3,5),16)||0;
            const b2=parseInt((obj.textBg||'#000').slice(5,7),16)||0;
            el.style.background = 'rgba('+r+','+g+','+b2+','+alpha+')';
          }
          const t = document.createElement('div');
          t.style.cssText = 'font-size:'+Math.max(5,Math.round((obj.fontSize||48)/22))+'px;'+
            'color:'+(obj.textColor||'#fff')+';font-family:"'+(obj.font||'Cinzel')+'",serif;'+
            'font-weight:'+(obj.bold?'700':'400')+';font-style:'+(obj.italic?'italic':'normal')+';'+
            'text-shadow:'+(obj.shadow?'0 1px 2px rgba(0,0,0,.9)':'none')+';'+
            'line-height:1.2;overflow:hidden;width:100%;text-align:'+(obj.textAlign||'center')+';';
          t.textContent = obj.text?.slice(0,60) || '';
          el.appendChild(t);

        } else if (obj.type === 'shape') {
          el.style.background   = obj.shapeFill || '#3b82f6';
          el.style.opacity      = String((obj.shapeOpacity??80)/100);
          el.style.borderRadius = obj.shape==='ellipse' ? '50%' : ((obj.shapeRadius||0)+'%');

        } else if (obj.type === 'svg') {
          // Use percentage-based positioning so SVG fills its box correctly
          const fill   = obj.shapeFill || '#c9a84c';
          const stroke = (obj.shapeBorderW||0) > 0 ? (obj.shapeBorder||'#fff') : 'none';
          const strokeW = obj.shapeBorderW || 0;
          // Compute pixel dimensions based on % of thumbnail (approx 130x73 px)
          const thumbW = (obj.w / CW) * 130;
          const thumbH = (obj.h / CH) * 73;
          el.style.opacity  = String((obj.shapeOpacity??90)/100);
          el.style.overflow = 'visible';
          el.innerHTML = _svgShapePath(obj.shape, thumbW, thumbH, fill, stroke, strokeW);

        } else if (obj.type === 'line') {
          el.style.borderTop = '2px solid '+(obj.shapeFill||'#fff');
          el.style.height = '2px';
          el.style.top = (((obj.y + obj.h/2) / CH) * 100) + '%';
          el.style.opacity = String((obj.shapeOpacity??80)/100);
        }
        box.appendChild(el);
      });
    }

    const num = document.createElement('div');
    num.style.cssText = 'position:absolute;bottom:3px;left:4px;font-size:8px;font-weight:700;background:rgba(0,0,0,.75);color:#fff;padding:1px 4px;border-radius:3px;';
    num.textContent = i + 1;
    box.appendChild(num);

    if (isActive) {
      const badge = document.createElement('div');
      badge.style.cssText = 'position:absolute;top:3px;right:3px;font-size:8px;font-weight:700;background:var(--gold);color:#000;padding:1px 5px;border-radius:3px;';
      badge.textContent = 'LIVE';
      box.appendChild(badge);
    }

    box.addEventListener('mouseenter', () => { box.style.transform='scale(1.04)'; if(!isActive)box.style.borderColor='rgba(201,168,76,.5)'; });
    box.addEventListener('mouseleave', () => { box.style.transform=''; box.style.borderColor=isActive?'var(--gold)':'var(--border)'; });
    box.addEventListener('click', () => { State.currentPresSlideIdx=i; presentPresSlide(i); });
    box.addEventListener('dblclick', () => presentPresSlide(i));
    strip.appendChild(box);
  });
}

async function presentPresSlide(idx, { silent = false } = {}) {
  State._logoActive = false;
  _updateLogoBtnState();
  previewPresSlide(idx);
  State.liveContentType = 'presentation';
  const pres  = State.presentations.find(p => String(p.id) === String(State.currentPresId));
  const slide = (State.currentPresSlides || [])[idx];
  if (!slide) return;

  if (window.electronAPI) {
    if (pres?.type === 'imported') {
      await window.electronAPI.projectPresentationSlide({ imagePath: slide.imagePath });
    } else {
      await window.electronAPI.projectCreatedSlide({ slide: slide.slide || {} });
    }
  } else if (State.isProjectionOpen) {
    if (pres?.type === 'imported') {
      _webPostRenderState('presentation-slide', { imagePath: slide.imagePath });
    } else {
      _webPostRenderState('created-slide', { slide: slide.slide || {} });
    }
  }
  if (!silent) toast(`📑 Slide ${idx + 1} of ${(State.currentPresSlides || []).length}`);
}

function presentCurrentPresSlide() { presentPresSlide(State.currentPresSlideIdx ?? 0); }

function navigatePresSlide(dir) {
  const total = State.currentPresSlides?.length || 0;
  if (!total) return;
  let next = (State.currentPresSlideIdx ?? 0) + dir;
  if (next < 0) next = 0;
  if (next >= total) next = total - 1;
  if (next === State.currentPresSlideIdx) return;
  // BUG-5 FIX: only project to live output when the operator has enabled live mode.
  // Previously always called presentPresSlide() which sent to projection unconditionally.
  if (State.isLive) {
    presentPresSlide(next);
  } else {
    previewPresSlide(next);
  }
}

function updatePresSlidePos() {
  // Nav position is now managed by _updatePresPreviewArea via #presNavPos
  const navPos = document.getElementById('presNavPos');
  const total  = State.currentPresSlides.length;
  if (navPos && total) navPos.textContent = `Slide ${(State.currentPresSlideIdx ?? 0) + 1} of ${total}`;
}

// ── Context menu ──────────────────────────────────────────────────────────────
function showPresContextMenu(e) {
  e.preventDefault();
  const menu = document.getElementById('presContextMenu');
  if (!menu) return;
  menu.style.display = 'block';
  // Position — keep inside viewport
  const x = Math.min(e.clientX, window.innerWidth  - 220);
  const y = Math.min(e.clientY, window.innerHeight - 200);
  menu.style.left = x + 'px';
  menu.style.top  = y + 'px';
  const hide = () => { menu.style.display = 'none'; document.removeEventListener('click', hide); };
  setTimeout(() => document.addEventListener('click', hide), 10);
}

// ─── Open the Presentation Editor window ─────────────────────────────────────
function newPresentation() {
  window.electronAPI?.openPresEditor({ id: null, name: 'New Presentation', slides: [] });
}

function editCurrentPresentation() {
  const pres = State.presentations.find(p => String(p.id) === String(State.currentPresId));
  if (!pres) { toast('Select a presentation first'); return; }
  if (pres.type === 'imported') {
    toast('ℹ Imported presentations are read-only — create a new one to edit.'); return;
  }
  window.electronAPI?.openPresEditor({ id: pres.id, name: pres.name, slides: pres.slides || [] });
}

function openPresEditor() { newPresentation(); } // legacy alias

// Receive saved data back from editor window
function _wirePresEditorSaved() {
  window.electronAPI?.on('pres-editor-saved', async (pres) => {
    if (!pres) return;
    // Persist to assets/Data/created_presentations.json
    await _saveLocalPresentation(pres);
    // Update in-memory state
    const idx = State.presentations.findIndex(p => p.id === pres.id);
    if (idx >= 0) State.presentations[idx] = pres;
    else          State.presentations.unshift(pres);

    State.currentPresId = pres.id;
    renderPresLibStrip();
    selectPresentation(pres.id);
    toast(`💾 Saved: "${pres.name}"`);
  });
}

// ── Created presentations — persisted to assets/Data/created_presentations.json ─
// In-memory cache so we don't hit IPC on every read
let _createdPresCache = null;

async function _getLocalPresentations() {
  if (_createdPresCache !== null) return _createdPresCache;
  try {
    if (window.electronAPI?.getCreatedPresentations) {
      _createdPresCache = await window.electronAPI.getCreatedPresentations() || [];
    } else {
      // Web fallback
      _createdPresCache = JSON.parse(localStorage.getItem('anchorcast_presentations') || '[]');
    }
  } catch(e) { _createdPresCache = []; }
  return _createdPresCache;
}

async function _saveLocalPresentation(pres) {
  const list = (await _getLocalPresentations()).filter(p => p.id !== pres.id);
  list.unshift(pres);
  _createdPresCache = list;
  if (window.electronAPI?.saveCreatedPresentations) {
    await window.electronAPI.saveCreatedPresentations(list);
  } else {
    localStorage.setItem('anchorcast_presentations', JSON.stringify(list));
  }
}

async function _deleteLocalPresentation(id) {
  const list = (await _getLocalPresentations()).filter(p => p.id !== id);
  _createdPresCache = list;
  if (window.electronAPI?.saveCreatedPresentations) {
    await window.electronAPI.saveCreatedPresentations(list);
  } else {
    localStorage.setItem('anchorcast_presentations', JSON.stringify(list));
  }
}

// ── Import presentation file (PPTX / PDF) via native dialog ──────────────────
async function importPresentation(pptxOnly = false) {
  if (!window.electronAPI) { toast('ℹ Import requires the desktop app'); return; }

  const picked = await window.electronAPI.pickPresFile({ pptxOnly });
  if (picked?.canceled || !picked?.filePath) return;

  const filePath = picked.filePath;
  const ext = (filePath.split('.').pop() || '').toLowerCase();

  const overlay = document.getElementById('presImportProgress');
  if (overlay) {
    overlay.style.display = 'flex';
    overlay.innerHTML = `
      <div style="text-align:center;padding:30px 40px;background:var(--panel);border-radius:10px;
        border:1px solid var(--border-lit);min-width:280px">
        <div style="font-size:28px;margin-bottom:12px">📑</div>
        <div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:4px">
          Importing ${escapeHtml(filePath.split(/[/\\]/).pop())}</div>
        <div style="font-size:11px;color:var(--text-dim);margin-bottom:16px">
          ${ext === 'pdf' ? 'Rendering pages…' : 'Converting via LibreOffice…'}</div>
        <div style="background:var(--border);border-radius:2px;overflow:hidden;height:4px">
          <div class="pres-progress-bar"></div></div>
        <div style="font-size:10px;color:var(--text-dim);margin-top:10px">
          Large files may take up to a minute</div>
      </div>`;
  }

  try {
    const result = await window.electronAPI.importPresentation({ filePath });
    if (overlay) overlay.style.display = 'none';
    if (!result.success) { toast(`⚠ Import failed: ${result.error}`); return; }

    const pres = { id: result.id, name: result.name, type: 'imported',
      slideCount: result.slideCount, outDir: result.outDir || '' };
    State.presentations.unshift(pres);
    renderPresLibStrip();
    await selectPresentation(result.id);
    toast(`📑 Imported "${result.name}" — ${result.slideCount} slides`);
  } catch(e) {
    if (overlay) overlay.style.display = 'none';
    toast(`⚠ Import error: ${e.message}`);
  }
}

async function deleteCurrentPresentation() {
  const pres = State.presentations.find(p => String(p.id) === String(State.currentPresId));
  if (!pres) { toast('Select a presentation first'); return; }
  // BUG-12 FIX: native confirm() is blocked in Electron contextIsolation builds.
  _confirmModal(`Delete "${pres.name}"? This cannot be undone.`, async () => {
    if (pres.type === 'imported' && window.electronAPI) {
      window.electronAPI.deletePresentation({ id: pres.id });
    } else {
      await _deleteLocalPresentation(pres.id);
    }
    State.presentations = State.presentations.filter(p => p.id !== pres.id);
    State.currentPresId = null;
    State.currentPresSlides = [];
    State.currentPresSlideIdx = 0;
    renderPresLibStrip();
    const presPA = document.getElementById('presPreviewArea');
    if (presPA) presPA.style.display = 'none';
    const vpa = document.getElementById('versePreviewArea');
    if (vpa) vpa.style.display = 'flex';
    const pt = document.getElementById('previewPanelTitle');
    if (pt) pt.textContent = '👁 Program Preview';
    toast(`🗑 Deleted: "${pres.name}"`);
  });
}

function runContextSearch(query) {
  if (!query.trim()) return;
  toast('🤖 Searching…');

  // If API key available, use Claude
  const apiKey = State.settings.apiKey;
  if (apiKey && State.isOnline) {
    runClaudeSearch(query);
  } else {
    // Fallback: local keyword search
    if (!query || query.length < 2) return;  // FIX: guard too-short queries
    const results = BibleDB.searchVerses(query, State.currentTranslation, 10);
    renderContextResults(results, query);
  }
}

async function runClaudeSearch(query) {
  const body = document.getElementById('searchResults');
  body.innerHTML = '<div class="empty-state"><span class="empty-icon">⏳</span>Searching with AI…</div>';

  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 500,
        messages: [{
          role: 'user',
          content: `Find Bible verses related to: "${query}"
          
Available refs: ${BibleDB.getAllRefs().map(r => r.ref).join(', ')}

Return JSON array of up to 6 matches: [{"ref": "Book Ch:V", "relevance": 0.0-1.0, "reason": "brief why"}]
JSON only, no markdown.`
        }]
      })
    });
    const data = await response.json();
    const text = data.content?.[0]?.text || '[]';
    const parsed = JSON.parse(text.replace(/```json|```/g, '').trim());

    const results = parsed.map(item => {
      const r = BibleDB.getAllRefs().find(x => x.ref === item.ref);
      if (!r) return null;
      return {
        ...r,
        text: BibleDB.getVerse(r.book, r.chapter, r.verse, State.currentTranslation),
        score: item.relevance,
        reason: item.reason,
      };
    }).filter(Boolean);

    renderContextResults(results, query);
  } catch (e) {
    const results = BibleDB.searchVerses(query, State.currentTranslation, 8);
    renderContextResults(results, query);
  }
}

function renderContextResults(results, query) {
  const body = document.getElementById('searchResults');
  body.innerHTML = '';

  const header = document.createElement('div');
  header.style.cssText = 'padding:8px 14px;font-size:10px;color:var(--gold);text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid var(--border)';
  header.textContent = `AI Results for: "${query}"`;
  body.appendChild(header);

  if (!results.length) {
    body.innerHTML += '<div class="empty-state"><span class="empty-icon">🔍</span>No matching verses found.</div>';
    return;
  }

  results.forEach(r => {
    const div = document.createElement('div');
    div.className = 'verse-row';
    div.innerHTML = `
      <div class="vr-num">${r.verse}</div>
      <div style="flex:1">
        <div style="font-family:'Cinzel',serif;font-size:10.5px;color:var(--gold);margin-bottom:3px">${r.ref}</div>
        <div class="vr-text">${r.text || ''}</div>
      </div>
      <div class="vr-actions">
        <button class="icon-btn play" title="Present">▶</button>
        <button class="icon-btn queue" title="Queue">+</button>
      </div>
    `;
    div.querySelector('.icon-btn.play').addEventListener('click', () => {
      presentVerse(r.book, r.chapter, r.verse, r.ref);
      navigateBibleSearch(r.book, r.chapter, r.verse);
    });
    div.querySelector('.icon-btn.queue').addEventListener('click', () =>
      addToQueue(r.book, r.chapter, r.verse, r.ref));
    body.appendChild(div);
  });
}

function renderSearchResults() {
  const body = document.getElementById('searchResults');
  body.innerHTML = '';
  const chapterNav = document.getElementById('chapterNav');
  if (chapterNav) chapterNav.style.display = '';
  document.getElementById('chapterNavTitle').textContent =
    `${State.currentBook} ${State.currentChapter}`;

  const verses = BibleDB.getChapter(
    State.currentBook, State.currentChapter, State.currentTranslation
  );

  if (!verses.length) {
    body.innerHTML =
      '<div class="empty-state"><span class="empty-icon">📖</span>' +
      'No verses found for this chapter.<br>Try a different book or chapter.</div>';
    return;
  }

  // ── Slide card view ────────────────────────────────────────────────────────
  if (State.previewViewMode === 'slides') {
    const verseObjs = verses
      .filter(({ text }) => !(text?.startsWith('[') && text?.includes('upload kjv.json')))
      .map(({ verse, text }) => ({
        book: State.currentBook, chapter: State.currentChapter,
        verse, text,
        ref: `${State.currentBook} ${State.currentChapter}:${verse}`,
      }));

    // Render into col 3 Program Preview (bibleSlideArea) — like song slides
    const grid     = document.getElementById('bibleSlideGrid');
    const titleEl  = document.getElementById('bibleSlideTitle');
    const subtitleEl = document.getElementById('bibleSlideSubtitle');
    if (titleEl)    titleEl.textContent    = `${State.currentBook} ${State.currentChapter}`;
    if (subtitleEl) subtitleEl.textContent = `${verseObjs.length} verses · ${State.currentTranslation}`;

    if (grid) {
      renderVerseCards(verseObjs, grid);
    }

    // Show ONLY the Bible slide area — hide verse canvas completely
    document.getElementById('versePreviewArea').style.display  = 'none';
    document.getElementById('songPreviewArea').style.display   = 'none';
    document.getElementById('mediaPreviewArea').style.display  = 'none';
    document.getElementById('bibleSlideArea').style.display    = 'flex';
    const pt = document.getElementById('previewPanelTitle');
    if (pt) pt.textContent = `📖 ${State.currentBook} ${State.currentChapter}`;
    return;
  }

  // Show a one-time banner if all verses are placeholders (no Bible data loaded)
  const allPlaceholders = verses.every(v => v.text && v.text.startsWith('[') && v.text.includes('upload kjv.json'));
  if (allPlaceholders) {
    const banner = document.createElement('div');
    banner.style.cssText = 'padding:10px 14px;background:rgba(201,168,76,.07);border-bottom:1px solid rgba(201,168,76,.2);display:flex;align-items:center;gap:10px;flex-shrink:0;';
    banner.innerHTML = `
      <span style="font-size:16px">📥</span>
      <div style="flex:1;font-size:11px;color:var(--gold);line-height:1.5">
        <strong>Full Bible text not loaded.</strong><br>
        <span style="color:var(--text-sub)">Go to <strong>⚙ Settings → 📖 Bible Versions</strong> and upload kjv.json to unlock all 31,102 verses.</span>
      </div>
      <button onclick="window.electronAPI?.openSettings()" style="padding:5px 10px;background:rgba(201,168,76,.15);border:1px solid var(--gold-dim);border-radius:5px;color:var(--gold);font-size:10px;font-weight:700;cursor:pointer;font-family:inherit;white-space:nowrap;text-transform:uppercase;letter-spacing:.4px">Open Settings</button>
    `;
    body.appendChild(banner);
  }

  verses.forEach(({ verse, text }) => {
    const isPlaceholder = text && text.startsWith('[') && text.includes('upload kjv.json');
    const isActive =
      !isPlaceholder &&
      State.previewVerse?.book === State.currentBook &&
      State.previewVerse?.chapter === State.currentChapter &&
      State.previewVerse?.verse === verse;

    const div = document.createElement('div');
    div.className = 'verse-row' + (isActive ? ' active' : '') + (isPlaceholder ? ' placeholder' : '');
    div.dataset.verse = verse;
    if (!isPlaceholder) div.title = 'Click to preview · Double-click to send Live';
    div.innerHTML = `
      <div class="vr-num">${verse}</div>
      <div class="vr-text">${isPlaceholder ? `<span style="color:var(--text-dim);font-style:italic;font-size:11px">${verse} — Full text not loaded. Go to Settings › Bible Versions to upload kjv.json</span>` : text}</div>
      ${!isPlaceholder ? `<div class="vr-actions">
        <button class="icon-btn play" title="Preview (click)">▶</button>
        <button class="icon-btn queue" title="Add to queue">+</button>
      </div>` : ''}
    `;

    const ref = `${State.currentBook} ${State.currentChapter}:${verse}`;

    if (!isPlaceholder) {
      div.draggable = true;
      div.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('application/anchorcast-item', JSON.stringify({
          type: 'verse', book: State.currentBook, chapter: State.currentChapter, verse, ref
        }));
        e.dataTransfer.effectAllowed = 'copy';
      });
      div.addEventListener('click', (e) => {
        if (e.target.closest('.vr-actions')) return;
        presentVerse(State.currentBook, State.currentChapter, verse, ref);
        if (State.isLive) sendPreviewToLive();
      });
      div.addEventListener('dblclick', (e) => {
        if (e.target.closest('.vr-actions')) return;
        presentVerse(State.currentBook, State.currentChapter, verse, ref);
        sendPreviewToLive();
      });
      div.querySelector('.icon-btn.play').addEventListener('click', (e) => {
        e.stopPropagation();
        presentVerse(State.currentBook, State.currentChapter, verse, ref);
        if (State.isLive) sendPreviewToLive();
      });
      div.querySelector('.icon-btn.queue').addEventListener('click', (e) => {
        e.stopPropagation();
        addToQueue(State.currentBook, State.currentChapter, verse, ref);
      });
    }

    body.appendChild(div);
  });
  const sr = document.getElementById('searchResults');
  try { syncBibleListHighlightAndScroll(); } catch (_) {}
}


function scrollElementIntoContainer(target, container, block = 'center') {
  if (!target || !container) return;
  const cRect = container.getBoundingClientRect();
  const tRect = target.getBoundingClientRect();
  const topGap = tRect.top - cRect.top;
  const bottomGap = tRect.bottom - cRect.bottom;
  if (block === 'center') {
    container.scrollTop += topGap - (cRect.height / 2 - tRect.height / 2);
  } else if (topGap < 0) {
    container.scrollTop += topGap - 8;
  } else if (bottomGap > 0) {
    container.scrollTop += bottomGap + 8;
  }
}

function highlightVerseRow(verse) {
  const container = document.getElementById('searchResults');
  const rows = document.querySelectorAll('#searchResults .verse-row[data-verse]');
  rows.forEach(r => {
    const active = parseInt(r.dataset.verse) === verse;
    r.classList.toggle('active', active);
    r.classList.toggle('selected', active);
    r.classList.toggle('current', active);
  });
  const target = document.querySelector(`#searchResults .verse-row[data-verse="${verse}"]`);
  if (target && container) scrollElementIntoContainer(target, container, 'center');
}

function _bibleSelectableElements() {
  const slideCards = Array.from(document.querySelectorAll('#bibleSlideGrid .slide-card[data-verse]'));
  if (slideCards.length) return slideCards;
  return Array.from(document.querySelectorAll('#searchResults .verse-row[data-verse]'));
}

function ensureBibleSelection() {
  if (State.previewVerse) return true;
  const items = _bibleSelectableElements();
  if (!items.length) return false;
  items[0].click();
  return true;
}

function navigateBibleSelection(delta) {
  const items = _bibleSelectableElements();
  if (!items.length) return false;

  // Primary: find the element already marked .active
  let idx = items.findIndex(el => el.classList.contains('active'));

  // FIX-1A: If no .active element, anchor to State.previewVerse so that
  // arrow-key navigation starts from the correct verse rather than index 0.
  // This covers the case where the verse was loaded via presentQueueItem()
  // (clicking a schedule row) which calls highlightVerseRow() asynchronously —
  // the .active class may not be applied yet when the user presses NEXT/PREV.
  if (idx < 0 && State.previewVerse) {
    idx = items.findIndex(el =>
      String(el.dataset.verse) === String(State.previewVerse.verse) &&
      String(el.dataset.partIdx || '0') === String(State.versePartIdx || 0)
    );
  }

  // FIX-1B: If we still have no anchor, return false so the caller can decide
  // what to do rather than blindly clamping to verse 1.
  if (idx < 0) return false;

  const next = idx + delta;
  // Reached boundary — let the caller know rather than repeating the same verse
  if (next < 0 || next >= items.length) return false;

  const target = items[next];
  if (!target) return false;
  target.click();
  const container = target.closest('#searchResults, #bibleSlideGrid');
  if (container) scrollElementIntoContainer(target, container, 'nearest');
  return true;
}

function prevChapter() {
  const chapters = BibleDB.getChapters(State.currentBook);
  const idx = chapters.indexOf(State.currentChapter);
  if (idx > 0) { State.currentChapter = chapters[idx - 1]; renderSearchResults(); return; }
  // Go to prev book
  const books = BibleDB.ALL_BOOKS.filter(b => BibleDB.getChapters(b).length > 0);
  const bi = books.indexOf(State.currentBook);
  if (bi > 0) {
    State.currentBook = books[bi - 1];
    const chs = BibleDB.getChapters(State.currentBook);
    State.currentChapter = chs[chs.length - 1];
    renderSearchResults();
  }
}

function nextChapter() {
  const chapters = BibleDB.getChapters(State.currentBook);
  const idx = chapters.indexOf(State.currentChapter);
  if (idx < chapters.length - 1) { State.currentChapter = chapters[idx + 1]; renderSearchResults(); return; }
  const books = BibleDB.ALL_BOOKS.filter(b => BibleDB.getChapters(b).length > 0);
  const bi = books.indexOf(State.currentBook);
  if (bi < books.length - 1) {
    State.currentBook = books[bi + 1];
    State.currentChapter = BibleDB.getChapters(State.currentBook)[0];
    renderSearchResults();
  }
}

// ─── THEME ────────────────────────────────────────────────────────────────────
function getActiveThemeData() {
  // Search all themes (State.themes + builtins) for current theme ID
  const all = typeof _getAllThemes === 'function' ? _getAllThemes() : State.themes;
  if (all.length) {
    const found = all.find(t => t.id === State.currentTheme);
    if (found) return found;
  }
  // Legacy built-in fallback
  const BUILTIN = {
    sanctuary: { bgType:'radial', bgColor1:'#10082a', bgColor2:'#000518', textColor:'#ede6d8', accentColor:'#c9a84c', refColor:'#c9a84c', transColor:'#6a5220', fontFamily:'Crimson Pro', fontSize:52, fontStyle:'normal', textAlign:'center', padding:80 },
    dawn:      { bgType:'radial', bgColor1:'#1a0a08', bgColor2:'#080010', textColor:'#f0e8d8', accentColor:'#e8904c', refColor:'#e8904c', transColor:'#8a5020', fontFamily:'Crimson Pro', fontSize:52, fontStyle:'normal', textAlign:'center', padding:80 },
    deep:      { bgType:'radial', bgColor1:'#001028', bgColor2:'#000408', textColor:'#e0ecf8', accentColor:'#4c9ae8', refColor:'#4c9ae8', transColor:'#204870', fontFamily:'Crimson Pro', fontSize:52, fontStyle:'normal', textAlign:'center', padding:80 },
    minimal:   { bgType:'solid',  bgColor1:'#000000', bgColor2:'#000000', textColor:'#ffffff', accentColor:'#aaaaaa', refColor:'#888888', transColor:'#555555', fontFamily:'DM Sans',     fontSize:56, fontStyle:'normal', textAlign:'center', padding:80 },
  };
  return BUILTIN[State.currentTheme] || BUILTIN.sanctuary;
}

// Generate inline CSS background string from a theme object
function _getBgCss(t) {
  if (!t) return 'background:#0a0a1e;';
  if (t.bgType === 'image' && t.bgImage) {
    return `background:url('${t.bgImage}') center/cover no-repeat;`;
  }
  if (t.bgType === 'solid') return `background:${t.bgColor1 || '#0a0a1e'};`;
  if (t.bgType === 'linear') return `background:linear-gradient(135deg,${t.bgColor1},${t.bgColor2});`;
  return `background:radial-gradient(ellipse at 35% 45%,${t.bgColor1} 0%,${t.bgColor2} 70%,#000 100%);`;
}

function getActiveSongThemeData() {
  const id = State.currentSongTheme;
  if (id && State.themes.length) {
    const found = State.themes.find(t => t.id === id && t.category === 'song');
    if (found) return found;
  }
  const songTheme = State.themes.find(t => t.category === 'song');
  if (songTheme) return songTheme;
  return { bgType:'solid', bgColor1:'#0a0a1e', bgColor2:'#000010', textColor:'#fff',
    accentColor:'#c9a84c', fontFamily:'Crimson Pro', fontSize:56, fontStyle:'normal',
    textTransform:'none', lineSpacing:1.4, padding:80, shadowOn:true };  // FIX: Crimson Pro has lowercase
}

function getActivePresThemeData() {
  const id = State.currentPresTheme;
  if (id && State.themes.length) {
    const found = State.themes.find(t => t.id === id && t.category === 'presentation');
    if (found) return found;
  }
  const presTheme = State.themes.find(t => t.category === 'presentation');
  if (presTheme) return presTheme;
  return { bgType:'solid', bgColor1:'#0a0a1e', bgColor2:'#000010', titleColor:'#fff',
    subtitleColor:'#c9a84c', noteColor:'#888', fontFamily:'Cinzel',
    titleSize:72, subtitleSize:36, textAlign:'center', padding:80, shadowOn:true };
}

function setTheme(theme) {
  State.currentTheme = theme;
  document.getElementById('previewCanvas').dataset.theme = theme;
  document.getElementById('liveCanvas').dataset.theme = theme;
  document.querySelectorAll('.theme-swatch').forEach(sw => {
    sw.classList.toggle('active', sw.dataset.t === theme);
  });
  const themeObj = State.themes?.find(t => t.id === theme);
  const transform = (themeObj?.textTransform && themeObj.textTransform !== 'none') ? themeObj.textTransform : (State.settings?.scriptureTextTransform || 'none');
  const previewDisplay = document.getElementById('previewDisplay');
  const liveDisplay = document.getElementById('liveDisplay');
  previewDisplay?.querySelectorAll('.verse-body-text').forEach(el => {
    el.style.textTransform = transform;
  });
  if (State.liveContentType === 'scripture') {
    liveDisplay?.querySelectorAll('.verse-body-text').forEach(el => {
      el.style.textTransform = transform;
    });
  }
  document.documentElement.style.setProperty('--theme-text-transform', transform);
  if (State.isLive && State.liveContentType === 'scripture' && State.liveVerse) syncProjection();

  // BUG-19 FIX: when the theme changes, the verse card grid (Program Preview
  // slides mode) must be re-rendered — cards use _splitVerseForTheme() which
  // is theme-aware, so the old cards show the wrong split/layout.
  if (State.currentTab === 'book' && State.previewViewMode === 'slides') {
    const grid = document.getElementById('bibleSlideGrid');
    if (grid && State.currentBook && State.currentChapter) {
      try { renderSearchResults(); } catch (_) {}
    }
  }
}

// Rebuild theme swatches to include custom themes
function refreshThemeSwatches() {
  const container = document.querySelector('.theme-swatches');
  if (!container || !State.themes.length) return;
  container.innerHTML = '';

  // Show scripture themes in Bible/scripture context, song themes in song context
  const isSongTab = State.currentTab === 'song';
  const category  = isSongTab ? 'song' : 'scripture';
  const activeId  = isSongTab ? State.currentSongTheme : State.currentTheme;

  const filtered = State.themes.filter(t => (t.category || 'scripture') === category);
  const list = filtered.length ? filtered : State.themes.slice(0, 6);

  list.forEach(t => {
    const sw = document.createElement('div');
    sw.className = 'theme-swatch' + (t.id === activeId ? ' active' : '');
    sw.dataset.t = t.id;
    sw.title = t.name;
    const bg = t.bgType === 'solid'
      ? t.bgColor1
      : `linear-gradient(135deg, ${t.bgColor1 || '#111'}, ${t.bgColor2 || '#000'})`;
    sw.style.background = bg;
    sw.addEventListener('click', () => {
      if (isSongTab) {
        State.currentSongTheme = t.id;
        refreshThemeSwatches();
        toast(`🎵 Song theme: ${t.name}`);
      } else {
        setTheme(t.id);
      }
    });
    container.appendChild(sw);
  });
}

// ─── WHISPER SOURCE TOGGLE ────────────────────────────────────────────────────
// ─── TRANSCRIPTION SOURCE ─────────────────────────────────────────────────────
// Sources: 'deepgram' (real-time WS) | 'local' (Whisper) | 'cloud' (OpenAI)
// Deepgram streams directly from the renderer — no IPC needed
// Local/Cloud go through the existing PCM → main process pipeline

let deepgramSocket    = null;
let deepgramConnected = false;

const SRC_ORDER = ['deepgram', 'local', 'cloud']; // cycle order
const SRC_CONFIG = {
  deepgram: { label: 'LIVE',  color: '#2ecc71', title: 'Deepgram real-time streaming (~300ms)' },
  local:    { label: 'LOCAL', color: '#3498db', title: 'Local Whisper (offline)'               },
  cloud:    { label: 'CLOUD', color: '#f39c12', title: 'OpenAI Whisper API (cloud)'            },
};

function cycleTranscriptSource() {
  const sources = SRC_ORDER; // ['deepgram', 'local', 'cloud']

  // Which sources are currently available
  const available = {
    deepgram: !!State.settings?.deepgramKey,
    local:    !!State.whisperLocalReady,
    cloud:    !!State.settings?.openAiKey,
  };

  const curIdx = sources.indexOf(State.whisperSource);

  // Find next available source starting from curIdx+1 (wraps around)
  let picked = null;
  for (let i = 1; i <= sources.length; i++) {
    const candidate = sources[(curIdx + i) % sources.length];
    if (available[candidate]) { picked = candidate; break; }
  }

  if (!picked) {
    // Nothing configured — show specific guidance based on what's missing
    _showWhisperSetupBanner({
      reason: 'no_python',
      setupBatExists: true,
      context: 'no_source_available',
    });
    return;
  }

  // Show warning if skipped an unavailable source
  const nextIdx = (curIdx + 1) % sources.length;
  const next    = sources[nextIdx];
  if (!available[next] && picked !== next) {
    const skippedMsgs = {
      deepgram: '(Deepgram key not set — add in Settings)',
      local:    '(Local Whisper not set up — click Set Up Now in the notification)',
      cloud:    '(OpenAI key not set — add in Settings)',
    };
    toast(`⏭ Skipped ${next.toUpperCase()} ${skippedMsgs[next]}`);
  }

  State.whisperSource = picked;

  // If recording, switch live
  if (State.isRecording) {
    stopDeepgramSocket();
    if (State.whisperSource === 'deepgram') startDeepgramStream();
  }

  window.electronAPI?.setWhisperSource?.(State.whisperSource);
  updateSrcToggleUI();

  const msgs = {
    deepgram: '⚡ Deepgram — real-time streaming active',
    local:    '💻 Local Whisper — offline active',
    cloud:    '☁ OpenAI Whisper — cloud active',
  };
  toast(msgs[State.whisperSource]);
}

function updateSrcToggleUI() {
  const dot   = document.getElementById('srcToggleDot');
  const label = document.getElementById('srcToggleLabel');
  const wrap  = document.getElementById('srcToggleWrap');
  if (!dot || !label) return;

  const src = State.whisperSource;
  const cfg = SRC_CONFIG[src] || SRC_CONFIG.local;

  const available = {
    deepgram: !!State.settings?.deepgramKey,
    local:    !!State.whisperLocalReady,
    cloud:    !!State.settings?.openAiKey,
  };
  const ok = available[src];

  // Dot colour: configured colour if available, grey if not
  dot.style.background   = ok ? cfg.color : '#555';
  dot.style.boxShadow    = ok ? `0 0 5px ${cfg.color}88` : 'none';
  label.textContent      = cfg.label;
  label.style.color      = ok ? cfg.color : 'var(--text-dim)';

  // Tooltip shows all available sources
  const avail = SRC_ORDER.filter(s => available[s]).map(s => SRC_CONFIG[s].label).join(' · ');
  wrap.title = `Source: ${cfg.label} — click to cycle (${avail || 'none configured'})`;

  // Always show the toggle
  wrap.style.display = 'flex';
}

// ── Deepgram WebSocket Streaming ──────────────────────────────────────────────
// Sends raw PCM directly to Deepgram over WebSocket — ~300ms latency
// No chunking, no IPC, results stream back word-by-word

function startDeepgramStream() {
  const key = State.settings?.deepgramKey;
  if (!key) return;

  stopDeepgramSocket(); // close any existing

  // Deepgram streaming URL — Nova-2 model, English, smart formatting
  const url = `wss://api.deepgram.com/v1/listen?` + new URLSearchParams({
    model:             'nova-2',
    language:          'en-US',
    smart_format:      'true',
    punctuate:         'true',
    interim_results:   'true',
    utterance_end_ms:  '1200',
    vad_events:        'true',
    encoding:          'linear16',
    sample_rate:       '16000',
    channels:          '1',
  });

  try {
    deepgramSocket = new WebSocket(url, ['token', key]);
  } catch(e) {
    console.warn('[Deepgram] WebSocket error:', e.message);
    return;
  }

  deepgramSocket.onopen = () => {
    deepgramConnected = true;
    console.log('[Deepgram] Connected — streaming live');
    toast('⚡ Deepgram connected — speak now');
    updateSrcToggleUI();
  };

  deepgramSocket.onmessage = (evt) => {
    try {
      const msg = JSON.parse(evt.data);

      // Real-time interim transcript
      if (msg.type === 'Results') {
        const alt = msg.channel?.alternatives?.[0];
        if (!alt?.transcript) return;

        if (msg.is_final) {
          const text = alt.transcript.trim();
          if (text) {
            removeInterim();
            pushTranscriptLine(text, true);
          }
        } else {
          // Show interim words as they come in
          if (alt.transcript.trim()) updateInterim(alt.transcript);
        }
      }
    } catch(e) { /* ignore parse errors */ }
  };

  deepgramSocket.onerror = (e) => {
    console.warn('[Deepgram] WebSocket error');
    deepgramConnected = false;
  };

  deepgramSocket.onclose = (evt) => {
    deepgramConnected = false;
    console.log(`[Deepgram] Closed (${evt.code})`);
    if (evt.code === 1008) toast('⚠ Deepgram: invalid API key — check Settings');
    else if (State.isRecording && State.whisperSource === 'deepgram') {
      // Auto-reconnect after 2s if still recording
      setTimeout(() => {
        if (State.isRecording && State.whisperSource === 'deepgram') startDeepgramStream();
      }, 2000);
    }
  };
}

function stopDeepgramSocket() {
  if (deepgramSocket) {
    try {
      // Send KeepAlive close signal to Deepgram
      if (deepgramSocket.readyState === WebSocket.OPEN) {
        deepgramSocket.send(JSON.stringify({ type: 'CloseStream' }));
      }
      deepgramSocket.close();
    } catch(e) {}
    deepgramSocket    = null;
    deepgramConnected = false;
  }
  removeInterim();
}

// Send PCM to Deepgram socket (called from AudioWorklet handler when source=deepgram)
function sendPcmToDeepgram(int16Buffer) {
  if (!deepgramSocket || deepgramSocket.readyState !== WebSocket.OPEN) return;
  deepgramSocket.send(int16Buffer);
}

// ── Old toggle aliases (keep for compat) ──────────────────────────────────────
function toggleWhisperSource() { cycleTranscriptSource(); }
function updateWhisperToggleUI() { updateSrcToggleUI(); }

// ─── PREVIEW VIEW MODE (List ↔ Slides) ────────────────────────────────────────
// Toggles between scrollable list and card grid
// Applies to: Bible chapter verses, Song slides

function setPreviewViewMode(mode) {
  State.previewViewMode = mode;
  const listBtn  = document.getElementById('viewListBtn');
  const slideBtn = document.getElementById('viewSlideBtn');
  if (listBtn)  { listBtn.style.background  = mode === 'list'   ? 'var(--gold)' : 'transparent';
                  listBtn.style.color        = mode === 'list'   ? '#000' : 'var(--text-dim)'; }
  if (slideBtn) { slideBtn.style.background = mode === 'slides' ? 'var(--gold)' : 'transparent';
                  slideBtn.style.color       = mode === 'slides' ? '#000' : 'var(--text-dim)'; }

  const tab = State.currentTab;

  if (mode === 'list') {
    // Always hide slide grid
    document.getElementById('bibleSlideArea').style.display = 'none';

    if (tab === 'song') {
      // Stay on song view — show song slides if one is selected, else empty
      if (State.currentSongId) {
        document.getElementById('versePreviewArea').style.display  = 'none';
        document.getElementById('mediaPreviewArea').style.display  = 'none';
        document.getElementById('songPreviewArea').style.display   = 'flex';
        renderSongSlides();
      }
      // else: no song selected — leave preview as-is (empty)
    } else if (tab === 'media') {
      // Stay on media view — show media preview if one is selected
      if (State.currentMediaId) {
        document.getElementById('versePreviewArea').style.display  = 'none';
        document.getElementById('songPreviewArea').style.display   = 'none';
        document.getElementById('mediaPreviewArea').style.display  = 'flex';
      }
      // else: leave as-is
    } else {
      // Bible / context tab — show verse canvas
      const vpa = document.getElementById('versePreviewArea');
      if (vpa) { vpa.style.display = 'flex'; vpa.style.flex = '1'; vpa.style.minHeight = '0'; }
      document.getElementById('songPreviewArea').style.display  = 'none';
      document.getElementById('mediaPreviewArea').style.display = 'none';
      const pt = document.getElementById('previewPanelTitle');
      if (pt) pt.textContent = '👁 Program Preview';
      renderSearchResults();
    }
  } else {
    // Slides mode — only acts on book or song tabs
    if (tab === 'song' && State.currentSongId) {
      renderSongSlides();
    } else if (tab === 'book') {
      renderSearchResults();
    }
    // Media tab: slides toggle has no effect
  }
}

// Render Bible verses as slide cards into a container
function renderVerseCards(verses, container) {
  container.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'slide-card-grid';

  verses.forEach((v) => {
    // Calculate parts for this verse based on active theme
    const parts = _splitVerseForTheme(v.book, v.chapter, v.verse, v.ref);
    const needsSplit = parts.length > 1;

    parts.forEach((partText, partIdx) => {
      const card = document.createElement('div');
      const isActivePart = State.previewVerse?.verse === v.verse &&
                           State.previewVerse?.chapter === v.chapter &&
                           State.previewVerse?.book === v.book &&
                           (State.versePartIdx || 0) === partIdx;

      card.className = 'slide-card verse-card-item' + (isActivePart ? ' active' : '');
      card.dataset.verse   = v.verse;
      card.dataset.partIdx = partIdx;
      card.dataset.partTotal = parts.length;

      // Show full part text in the card (no truncation within parts)
      const displayText = partText.length > 180 ? partText.slice(0, 180) + '…' : partText;

      // Part badge for multi-part verses
      const partBadgeHtml = needsSplit
        ? `<div class="slide-card-label">${partIdx + 1}/${parts.length}</div>`
        : '';

      // Footer: ref + part indicator
      const footerText = needsSplit
        ? `${escapeHtml(v.ref || '')} · Part ${partIdx + 1}`
        : escapeHtml(v.ref || '');

      card.innerHTML = `
        <span class="slide-card-num">${v.verse}</span>
        ${partBadgeHtml}
        <div class="slide-card-content verse-card">${escapeHtml(displayText)}</div>
        <div class="slide-card-footer">${footerText}</div>`;

      card.addEventListener('click', () => {
        // Highlight this card
        grid.querySelectorAll('.slide-card').forEach(c => c.classList.remove('active'));
        card.classList.add('active');

        presentVerse(v.book, v.chapter, v.verse, v.ref, { parts, partIdx });
      });

      card.addEventListener('dblclick', () => {
        State.verseParts   = parts;
        State.versePartIdx = partIdx;
        _updateVersePartIndicator(partIdx + 1, parts.length);
        presentVerse(v.book, v.chapter, v.verse, v.ref, { parts, partIdx });
        sendPreviewToLive();
        if (State.isLive) syncProjection();
      });

      grid.appendChild(card);
    });
  });

  container.appendChild(grid);
}

// Render a specific verse part text into the previewDisplay with theme styling
function _renderVersePartInPreview(partText, ref, partLabel) {
  const t = getActiveThemeData();
  const display = document.getElementById('previewDisplay');
  if (!display) return;
  if (t && !t?.boxes?.length) _applyLiveThemeBackground(display, t);

  if (t?.boxes?.length) {
    // Box-based theme — render using theme boxes
    const canvas = document.getElementById('previewCanvas');
    const cw = canvas?.clientWidth  || 400;
    const ch = canvas?.clientHeight || 225;
    const sc2 = Math.min(cw / 1920, ch / 1080);

    let html = '';
    t.boxes.forEach(box => {
      const bx = (box.x * sc2) + 'px', by = (box.y * sc2) + 'px';
      const bw = (box.w * sc2) + 'px', bh = (box.h * sc2) + 'px';
      const fs = Math.max(7, (box.fontSize || 52) * sc2) + 'px';
      const textContent = (box.role === 'ref' || box.role === 'title') ? ref : partText;
      const bgF = box.bgOpacity > 0 && box.bgFill ? `background:rgba(0,0,0,${box.bgOpacity/100});` : '';
      const bord = box.borderW > 0 ? `border:${Math.max(1, box.borderW*sc2)}px solid ${box.borderColor||'#fff'};` : '';
      html += `<div style="position:absolute;left:${bx};top:${by};width:${bw};height:${bh};
        display:flex;align-items:${box.valign==='top'?'flex-start':box.valign==='bottom'?'flex-end':'center'};
        justify-content:${box.align==='left'?'flex-start':box.align==='right'?'flex-end':'center'};
        text-align:${box.align||'center'};overflow:hidden;padding:6px;box-sizing:border-box;${bgF}${bord}">
        <div style="font-size:${fs};color:${box.color||'#fff'};
          font-family:'${box.fontFamily||'Crimson Pro'}',serif;font-weight:${Number(box.fontWeight || (box.bold?700:400))};
          font-style:${box.italic?'italic':'normal'};line-height:${box.lineSpacing||1.4};
          letter-spacing:${(Number(box.letterSpacing||0) * sc2)}px;
          text-shadow:${box.shadow === false ? 'none' : `${Number(box.shadowOffsetX||0)*sc2}px ${Number(box.shadowOffsetY||2)*sc2}px ${Math.max(0, Number(box.shadowBlur ?? 8))*sc2}px ${box.shadowColor || '#000000'}`};width:100%;text-align:${box.align||'center'};
          text-transform:${box.textTransform || t.textTransform || State.settings?.scriptureTextTransform || 'none'};
          white-space:pre-wrap">${_escapeHtmlBasic(textContent)}</div>
      </div>`;
    });
    display.style.cssText = 'display:flex;position:absolute;inset:0;overflow:hidden;background:transparent';
    display.innerHTML = `<div style="position:absolute;inset:0">${html}</div>`;
  }
  // Legacy themes: the existing buildVerseHTML already handles display
}

// Render song sections as slide cards
function renderSongSlideCards(song, container) {
  if (!song?.sections?.length) return;
  container.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'slide-card-grid';

  song.sections.forEach((sec, i) => {
    const card = document.createElement('div');
    card.className = 'slide-card song-card';
    if (i === State.currentSongSlideIdx) card.classList.add('active');

    const lines = (sec.lines || []).filter(l => l.trim());
    const text  = lines.join('\n');

    card.innerHTML = `
      <span class="slide-card-num">${i + 1}</span>
      ${sec.label ? `<span class="slide-card-label" style="${_getSectionLabelStyle(sec.label)}">${escapeHtml(sec.label)}</span>` : ''}
      <div class="slide-card-content">${lines.map(l => escapeHtml(l)).join('<br>')}</div>
      <div class="slide-card-footer">${escapeHtml(song.title || '')}</div>`;

    card.dataset.idx = i; // needed for arrow key navigation
    card.addEventListener('click',  () => previewSongSlide(i));
    card.addEventListener('dblclick', () => presentSongSlide(i));
    grid.appendChild(card);
  });

  container.appendChild(grid);
}

// ─── MODE TOGGLES ─────────────────────────────────────────────────────────────
function toggleMode() {
  State.isOnline = !State.isOnline;
  State.settings.onlineMode = State.isOnline;  // FIX: persist across restart
  const pill  = document.getElementById('modeToggle');
  const label = document.getElementById('modeLabel');
  pill.className  = 'status-pill ' + (State.isOnline ? 'online' : 'offline');
  label.textContent = State.isOnline ? 'Online' : 'Offline';
  // Save immediately so the mode survives app restart
  if (window.electronAPI) {
    window.electronAPI.getSettings().then(s => {
      window.electronAPI.saveSettings({ ...s, onlineMode: State.isOnline });
    }).catch(() => {});
  }
  // NOTE: offline mode only disables Claude AI — keyword/verbal/direct detection always runs
  if (window.AIDetection) AIDetection.setEnabled(State.isOnline);
  toast(State.isOnline
    ? '🌐 Online — Claude AI + keyword detection active'
    : '📴 Offline — keyword & verbal detection active');
}

function setDisplayMode(auto) {
  State.isAutoMode = auto;
  const badge = document.getElementById('displayModeBadge');
  badge.textContent = auto ? 'Auto' : 'Manual';
  badge.className = 'mode-badge ' + (auto ? 'auto' : 'manual');
  toast(auto ? '🤖 Auto mode active' : '🖐 Manual mode active');
}

// ─── PROJECTION ───────────────────────────────────────────────────────────────
async function openProjection() {
  if (window.electronAPI) {
    const displays = await window.electronAPI.getDisplays();
    const secondDisplay = displays.find(d => !d.isPrimary);
    await window.electronAPI.openProjection(secondDisplay?.id || displays[0].id);
    toast('⛶ Projection window opened');
    setTimeout(() => _syncCurrentProjection(), 500);
  } else {
    window.open('/projection.html', 'anchorcast-projection', 'width=1280,height=720,menubar=no,toolbar=no,location=no,status=no');
    State.isProjectionOpen = true;
    toast('⛶ Projection window opened');
    setTimeout(() => _syncCurrentProjection(), 500);
  }
}

function openFullscreenPreview() {
  const overlay = document.getElementById('projOverlay');
  const inner = document.getElementById('projInner');

  if (State.previewVerse || State.liveVerse) {
    const v = State.previewVerse || State.liveVerse;
    const text = getVerseText(v.book, v.chapter, v.verse);
    if (text) {
      inner.innerHTML = `
        <div class="proj-ref">${v.ref} &nbsp;·&nbsp; ${State.currentTranslation}</div>
        <div class="proj-text"><sup style="font-size:26px;color:var(--gold)">${v.verse}</sup>${text}</div>
        <div class="proj-trans">${State.currentTranslation}</div>
      `;
    }
  } else {
    inner.innerHTML = `<div style="font-size:80px;opacity:0.08;font-family:'Cinzel',serif">✝</div><div style="color:#555;font-size:18px;margin-top:20px">No verse selected</div>`;
  }

  overlay.classList.add('show');
  overlay.style.display = 'flex';
}

function closeFullscreenPreview() {
  const overlay = document.getElementById('projOverlay');
  overlay.classList.remove('show');
  overlay.style.display = 'none';
}

// ─── SERMON NOTES ─────────────────────────────────────────────────────────────
// Load saved transcripts into the Sermon Notes picker
async function _loadSavedTranscripts() {
  const list = document.getElementById('savedTranscriptsList');
  if (!list) return;
  list.innerHTML = '<div style="font-size:11px;color:var(--text-dim);padding:4px">Loading...</div>';
  const transcripts = await window.electronAPI?.getTranscripts?.() || [];
  if (!transcripts.length) {
    list.innerHTML = '<div style="font-size:11px;color:var(--text-dim);padding:4px">No saved transcripts yet.<br>Transcripts are saved automatically when you stop recording.</div>';
    return;
  }
  list.innerHTML = '';
  transcripts.forEach(t => {
    const date = new Date(t.savedAt || t.date).toLocaleDateString('en-US', { month:'short', day:'numeric', year:'numeric' });
    const words = t.wordCount ? `${t.wordCount} words` : '';
    const item = document.createElement('div');
    item.className = 'saved-transcript-item';
    item.dataset.id = t.id;
    item.style.cssText = [
      'padding:7px 8px','border-radius:6px','border:1px solid var(--border)',
      'background:var(--card)','cursor:pointer','font-size:11px',
      'display:flex','align-items:flex-start','gap:6px',
      'transition:background .15s',
    ].join(';');
    item.innerHTML = `
      <div style="flex:1;min-width:0">
        <div style="font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${t.title || 'Sermon'}</div>
        <div style="color:var(--text-dim);margin-top:1px">${date}${words ? ' · ' + words : ''}</div>
      </div>
      <button data-del="${t.id}" style="
        background:transparent;border:none;color:var(--text-dim);
        cursor:pointer;font-size:14px;padding:0 2px;flex-shrink:0;
        line-height:1;opacity:.6" title="Delete">🗑</button>
    `;
    // Select transcript
    item.addEventListener('click', (e) => {
      if (e.target.dataset.del) return; // don't select on delete click
      document.querySelectorAll('.saved-transcript-item').forEach(el => {
        el.style.background = 'var(--card)';
        el.style.borderColor = 'var(--border)';
      });
      if (document.getElementById('useCurrentTranscriptBtn')) {
        document.getElementById('useCurrentTranscriptBtn').style.background = 'rgba(201,168,76,.15)';
      }
      item.style.background = 'rgba(56,189,248,.1)';
      item.style.borderColor = 'rgba(56,189,248,.3)';
      State._selectedSavedTranscriptId = t.id;
      State._selectedSavedTranscriptText = t.text;
      const info = document.getElementById('transcriptPreviewInfo');
      info.style.display = 'block';
      info.textContent = `Selected: ${t.title || 'Sermon'} · ${words}`;
    });
    // Delete transcript
    item.querySelector('[data-del]').addEventListener('click', async (e) => {
      e.stopPropagation();
      if (!confirm(`Delete transcript "${t.title || 'Sermon'}"?`)) return;
      await window.electronAPI?.deleteTranscript?.(t.id);
      if (State._selectedSavedTranscriptId === t.id) {
        State._selectedSavedTranscriptId = null;
        State._selectedSavedTranscriptText = null;
      }
      _loadSavedTranscripts(); // refresh list
      toast('🗑 Transcript deleted');
    });
    list.appendChild(item);
  });
}

async function generateSermonNotes() {
  const output = document.getElementById('notesOutput');
  // Use selected saved transcript OR current live transcript
  let transcriptText = null;
  if (State._selectedSavedTranscriptId && State._selectedSavedTranscriptText) {
    transcriptText = State._selectedSavedTranscriptText;
  } else if (State.transcriptLines?.length) {
    transcriptText = State.transcriptLines.map(l => l.text).join('\n');
  }
  if (!transcriptText) {
    output.textContent = '⚠ No transcript selected. Pick a saved transcript or start a live transcription.';
    return;
  }
  if (!State.settings.apiKey) {
    output.textContent = '⚠ Claude API key required. Add it in Settings.';
    return;
  }

  output.textContent = '⏳ Generating notes with AI…';

  try {
    const notes = await AIDetection.generateSermonNotes(transcriptText);
    State.generatedNotes = notes;

    let formatted = '';
    if (notes.title) formatted += `TITLE: ${notes.title}\n\n`;
    if (notes.topic) formatted += `TOPIC: ${notes.topic}\n\n`;
    if (notes.summary) formatted += `SUMMARY:\n${notes.summary}\n\n`;
    if (notes.mainPoints?.length) {
      formatted += 'MAIN POINTS:\n';
      notes.mainPoints.forEach((p, i) => {
        formatted += `\n${i+1}. ${p.heading}\n${p.content}`;
        if (p.scriptures?.length) formatted += `\n   Scripture: ${p.scriptures.join(', ')}`;
        formatted += '\n';
      });
    }
    if (notes.keyVerses?.length) {
      formatted += '\nKEY VERSES:\n';
      notes.keyVerses.forEach(v => { formatted += `• ${v.ref}: ${v.text}\n`; });
    }
    if (notes.practicalApplications?.length) {
      formatted += '\nPRACTICAL APPLICATIONS:\n';
      notes.practicalApplications.forEach(a => { formatted += `• ${a}\n`; });
    }
    if (notes.closingThought) formatted += `\nCLOSING THOUGHT:\n${notes.closingThought}\n`;

    output.textContent = formatted || JSON.stringify(notes, null, 2);
    toast('✦ Sermon notes generated');
  } catch (e) {
    output.textContent = `⚠ Error: ${e.message}`;
  }
}

async function copyNotesToClipboard() {
  const content = document.getElementById('notesOutput').textContent;
  if (!content || content.startsWith('Notes will')) { toast('⚠ Generate notes first'); return; }
  try {
    await navigator.clipboard.writeText(content);
    toast('📋 Notes copied to clipboard');
  } catch (e) {
    const ta = document.createElement('textarea');
    ta.value = content;
    ta.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    ta.remove();
    toast('📋 Notes copied to clipboard');
  }
}

async function exportNotes() {
  const content = document.getElementById('notesOutput').textContent;
  if (!content || content.startsWith('Notes will')) { toast('⚠ Generate notes first'); return; }
  if (window.electronAPI?.exportTranscript) {
    const title = State.generatedNotes?.title
      ? State.generatedNotes.title.replace(/[^a-z0-9]/gi, '-').toLowerCase()
      : 'sermon-notes';
    await window.electronAPI.exportTranscript({
      content: State.generatedNotes
        ? JSON.stringify(State.generatedNotes, null, 2)
        : content,
      defaultName: `${title}.json`,
    });
    toast('↓ Notes exported (JSON — reimportable)');
  } else {
    const blob = new Blob([content], { type: 'text/plain' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'sermon-notes.txt';
    a.click();
    toast('↓ Notes exported');
  }
}

async function generatePresFromNotes() {
  const notes = State.generatedNotes;
  if (!notes || typeof notes !== 'object') {
    toast('⚠ Generate or import notes first'); return;
  }

  const DARK_BG = '#0a0a1e';
  const GOLD    = '#c9a84c';
  const WHITE   = '#ffffff';
  const SUBTEXT = '#94a3b8';

  function txt(text, x, y, w, h, opts = {}) {
    return {
      type: 'text', text,
      x, y, w, h,
      fontSize:   opts.fontSize   ?? 28,
      bold:       opts.bold       ?? false,
      italic:     opts.italic     ?? false,
      color:      opts.color      ?? WHITE,
      align:      opts.align      ?? 'center',
      shadow:     opts.shadow     ?? false,
      textBg:     '#000', textBgAlpha: 0,
      rotation:   0,
    };
  }

  function slide(bgColor = DARK_BG) {
    return { bg: bgColor, bgImage: null, objects: [] };
  }

  const slides = [];

  const titleSlide = slide();
  if (notes.title) {
    titleSlide.objects.push(txt(notes.title, 120, 280, 1680, 120,
      { fontSize: 52, bold: true, color: GOLD }));
  }
  if (notes.topic) {
    titleSlide.objects.push(txt(notes.topic, 120, 430, 1680, 60,
      { fontSize: 24, color: SUBTEXT }));
  }
  slides.push(titleSlide);

  if (notes.summary) {
    const s = slide();
    s.objects.push(txt('Summary', 120, 80, 1680, 80,
      { fontSize: 32, bold: true, color: GOLD }));
    s.objects.push(txt(notes.summary, 120, 200, 1680, 520,
      { fontSize: 22, align: 'left', color: WHITE }));
    slides.push(s);
  }

  (notes.mainPoints || []).forEach((point, i) => {
    const s = slide();
    s.objects.push(txt(`${i + 1}. ${point.heading}`, 120, 80, 1680, 90,
      { fontSize: 34, bold: true, color: GOLD }));
    if (point.content) {
      s.objects.push(txt(point.content, 120, 210, 1680, 400,
        { fontSize: 22, align: 'left', color: WHITE }));
    }
    if (point.scriptures?.length) {
      s.objects.push(txt('📖 ' + point.scriptures.join('  •  '), 120, 640, 1680, 50,
        { fontSize: 18, color: SUBTEXT, align: 'left' }));
    }
    slides.push(s);
  });

  if (notes.keyVerses?.length) {
    const s = slide();
    s.objects.push(txt('Key Verses', 120, 80, 1680, 80,
      { fontSize: 32, bold: true, color: GOLD }));
    const verseText = notes.keyVerses
      .map(v => `${v.ref}\n${v.text}`)
      .join('\n\n');
    s.objects.push(txt(verseText, 120, 200, 1680, 500,
      { fontSize: 20, align: 'left', color: WHITE }));
    slides.push(s);
  }

  if (notes.practicalApplications?.length) {
    const s = slide();
    s.objects.push(txt('Practical Applications', 120, 80, 1680, 80,
      { fontSize: 32, bold: true, color: GOLD }));
    const appText = notes.practicalApplications.map(a => `• ${a}`).join('\n\n');
    s.objects.push(txt(appText, 120, 200, 1680, 500,
      { fontSize: 22, align: 'left', color: WHITE }));
    slides.push(s);
  }

  if (notes.closingThought) {
    const s = slide();
    s.objects.push(txt('Closing Thought', 120, 200, 1680, 80,
      { fontSize: 30, bold: true, color: GOLD }));
    s.objects.push(txt(notes.closingThought, 120, 330, 1680, 300,
      { fontSize: 24, color: WHITE }));
    slides.push(s);
  }

  if (!slides.length) { toast('⚠ No content to generate slides from'); return; }

  const presName = notes.title ? `${notes.title} — Sermon` : 'Sermon Presentation';
  const pres = {
    id:         Date.now(),
    name:       presName,
    type:       'created',
    slideCount: slides.length,
    slides,
  };

  await _saveLocalPresentation(pres);
  _createdPresCache = null;
  const idx = State.presentations ? State.presentations.findIndex(p => p.id === pres.id) : -1;
  if (idx >= 0) State.presentations[idx] = pres;
  else if (State.presentations) State.presentations.unshift(pres);
  else State.presentations = [pres];

  closeModal('notesOverlay');
  switchTab('pres');
  renderPresLibStrip();
  await selectPresentation(pres.id);
  toast(`📑 "${presName}" — ${slides.length} slides added to library`);
}

async function importNotesFromFile() {
  try {
    let result;
    if (window.electronAPI?.importNotes) {
      result = await window.electronAPI.importNotes();
    } else {
      result = await new Promise(resolve => {
        const inp = document.createElement('input');
        inp.type = 'file'; inp.accept = '.txt,.md,.json';
        inp.onchange = async () => {
          const file = inp.files[0];
          if (!file) { resolve({ success: false }); return; }
          const content = await file.text();
          resolve({ success: true, content, filePath: file.name });
        };
        inp.click();
      });
    }

    if (!result?.success) return;

    const content = result.content || '';
    const output  = document.getElementById('notesOutput');

    if ((result.filePath || '').toLowerCase().endsWith('.json')) {
      try {
        const parsed = JSON.parse(content);
        if (parsed && (parsed.title || parsed.mainPoints || parsed.summary)) {
          State.generatedNotes = parsed;
          let formatted = '';
          if (parsed.title)   formatted += `TITLE: ${parsed.title}\n\n`;
          if (parsed.topic)   formatted += `TOPIC: ${parsed.topic}\n\n`;
          if (parsed.summary) formatted += `SUMMARY:\n${parsed.summary}\n\n`;
          if (parsed.mainPoints?.length) {
            formatted += 'MAIN POINTS:\n';
            parsed.mainPoints.forEach((p, i) => {
              formatted += `\n${i+1}. ${p.heading}\n${p.content}`;
              if (p.scriptures?.length) formatted += `\n   Scripture: ${p.scriptures.join(', ')}`;
              formatted += '\n';
            });
          }
          if (parsed.keyVerses?.length) {
            formatted += '\nKEY VERSES:\n';
            parsed.keyVerses.forEach(v => { formatted += `• ${v.ref}: ${v.text}\n`; });
          }
          if (parsed.practicalApplications?.length) {
            formatted += '\nPRACTICAL APPLICATIONS:\n';
            parsed.practicalApplications.forEach(a => { formatted += `• ${a}\n`; });
          }
          if (parsed.closingThought) formatted += `\nCLOSING THOUGHT:\n${parsed.closingThought}\n`;
          output.textContent = formatted || content;
          toast('↑ Notes imported — ready to generate presentation');
          return;
        }
      } catch (_) { }
    }

    output.textContent = content;
    State.generatedNotes = _parseTextNotes(content);
    toast('↑ Notes imported');

  } catch (e) {
    toast('⚠ Import failed: ' + (e.message || 'unknown error'));
  }
}

function _parseTextNotes(text) {
  const lines  = text.split('\n');
  const notes  = { title:'', topic:'', summary:'', mainPoints:[], keyVerses:[], practicalApplications:[], closingThought:'' };
  let section  = null;
  let buf      = [];
  let curPoint = null;

  const flush = () => {
    if (!section) return;
    const content = buf.join('\n').trim();
    if (section === 'SUMMARY')   notes.summary = content;
    if (section === 'CLOSING')   notes.closingThought = content;
    if (section === 'POINT' && curPoint) { curPoint.content = content; notes.mainPoints.push(curPoint); curPoint = null; }
    buf = [];
  };

  for (const line of lines) {
    const t = line.trim();
    if (!t) continue;
    if (t.startsWith('TITLE:'))   { notes.title   = t.slice(6).trim(); continue; }
    if (t.startsWith('TOPIC:'))   { notes.topic   = t.slice(6).trim(); continue; }
    if (t === 'SUMMARY:')         { flush(); section = 'SUMMARY'; continue; }
    if (t === 'MAIN POINTS:')     { flush(); section = 'POINTS'; continue; }
    if (t === 'KEY VERSES:')      { flush(); section = 'VERSES'; continue; }
    if (t === 'PRACTICAL APPLICATIONS:') { flush(); section = 'APPS'; continue; }
    if (t === 'CLOSING THOUGHT:') { flush(); section = 'CLOSING'; continue; }

    if (section === 'POINTS' && /^\d+\./.test(t)) {
      flush(); section = 'POINT';
      curPoint = { heading: t.replace(/^\d+\.\s*/, ''), content: '', scriptures: [] };
      continue;
    }
    if (section === 'POINT' && t.startsWith('Scripture:')) {
      if (curPoint) curPoint.scriptures = t.slice(10).split(',').map(s => s.trim());
      continue;
    }
    if (section === 'VERSES' && t.startsWith('•')) {
      const [ref, ...rest] = t.slice(1).split(':');
      if (ref) notes.keyVerses.push({ ref: ref.trim(), text: rest.join(':').trim() });
      continue;
    }
    if (section === 'APPS' && t.startsWith('•')) {
      notes.practicalApplications.push(t.slice(1).trim());
      continue;
    }
    buf.push(line);
  }
  flush();

  if (!notes.title && !notes.summary && !notes.mainPoints.length) {
    notes.summary = text.trim();
  }
  return notes;
}

// ─── NDI PANEL ────────────────────────────────────────────────────────────────
async function openNdiPanel() {
  showModal('ndiOverlay');
  if (!window.electronAPI) return;
  const info = await window.electronAPI.ndiStatus().catch(() => ({}));
  updateNdiPanel(info);
}

function updateNdiPanel(info) {
  if (!info) return;
  const statusLabel = document.getElementById('ndiStatusLabel');
  const methodLabel = document.getElementById('ndiMethodLabel');
  const startBtn    = document.getElementById('ndiStartBtn');
  const stopBtn     = document.getElementById('ndiStopBtn');
  const activeInfo  = document.getElementById('ndiGrandioseOk');
  const streamUrls  = document.getElementById('ndiStreamUrls');
  const offInfo     = document.getElementById('ndiOffInfo');
  const obsUrlEl    = document.getElementById('ndiObsUrl');
  const vmixUrlEl   = document.getElementById('ndiVmixUrl');

  if (!statusLabel) return;

  const { status, ndiSdkActive, obsUrl, vmixUrl, clientCount, addonState, addonLabel } = info;
  const isActive = status === 'running' || status === 'fallback';

  if (isActive) {
    if (ndiSdkActive) {
      statusLabel.textContent = '🟢 NDI Output Running';
      methodLabel.textContent = 'In OBS/vMix: look for "AnchorCast" in NDI Sources';
      if (activeInfo) {
        activeInfo.querySelector('div').innerHTML =
          '✓ <strong>NDI active</strong> — AnchorCast appears as an NDI source in OBS, vMix, Wirecast etc.';
      }
    } else {
      statusLabel.textContent = `🟡 Live Stream Running — MJPEG (${clientCount || 0} receiver(s))`;
      methodLabel.textContent = 'Add as Browser Source in OBS, or Web Browser input in vMix';
      if (activeInfo) {
        activeInfo.querySelector('div').innerHTML =
          '✓ <strong>MJPEG stream active</strong> — connect OBS or vMix using the URLs below. ' +
          '<span style="color:var(--text-dim)">Build ndi-addon for NDI output instead.</span>';
      }
    }
  } else {
    const addonReady = addonState === 'ready';
    statusLabel.textContent = 'External Output Off';
    methodLabel.textContent = addonReady
      ? 'Click Start — NDI detected, will appear as NDI source'
      : 'Click Start — MJPEG stream (build ndi-addon for NDI)';
  }

  if (startBtn) startBtn.style.display = isActive ? 'none' : '';
  if (stopBtn)  stopBtn.style.display  = isActive ? ''     : 'none';
  if (activeInfo) activeInfo.style.display = isActive ? '' : 'none';
  if (streamUrls) streamUrls.style.display = (isActive && !ndiSdkActive) ? '' : 'none';
  if (offInfo)    offInfo.style.display    = isActive ? 'none' : '';

  if (obsUrlEl  && obsUrl)  obsUrlEl.textContent  = obsUrl;
  if (vmixUrlEl && vmixUrl) vmixUrlEl.textContent = vmixUrl;
}

function showAboutModal() {
  let existing = document.getElementById('aboutOverlay');
  if (existing) { existing.classList.add('show'); return; }
  const overlay = document.createElement('div');
  overlay.id = 'aboutOverlay';
  overlay.className = 'modal-overlay show';
  overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.7);display:flex;align-items:center;justify-content:center;z-index:9999;';
  overlay.innerHTML = `
    <div style="background:#1a1a2e;border:1px solid #333;border-radius:14px;padding:36px 40px;max-width:440px;width:90%;color:#e0e0e0;font-family:inherit;text-align:center;box-shadow:0 8px 40px rgba(0,0,0,.6);position:relative;">
      <div style="font-size:14px;letter-spacing:2px;text-transform:uppercase;color:#888;margin-bottom:6px;">&#10022;</div>
      <h2 style="margin:0 0 4px;font-size:24px;color:#fff;font-weight:700;">AnchorCast</h2>
      <div style="font-size:13px;color:#888;margin-bottom:18px;">v1.2.0</div>
      <p style="font-size:14px;line-height:1.7;color:#bbb;margin:0 0 18px;">
        AI-powered worship presentation, live sermon transcription &amp; Bible verse display for churches.
      </p>
      <div style="font-size:13px;color:#999;line-height:1.8;margin-bottom:18px;">
        Live Transcription &middot; AI Verse Detection &middot; Song Manager<br>
        Presentation Editor &middot; NDI Output &middot; Remote Control<br>
        Theme Designer &middot; Sermon Intelligence &middot; Analytics
      </div>
      <div style="border-top:1px solid #333;padding-top:16px;margin-bottom:14px;">
        <div style="font-size:13px;color:#888;margin-bottom:6px;">Developed by</div>
        <div style="font-size:15px;color:#e0e0e0;font-weight:600;">Godbless Keku &amp; NK Ofodile</div>
      </div>
      <div style="margin-bottom:18px;">
        <span style="display:inline-block;background:#22c55e22;color:#4ade80;font-size:12px;padding:3px 12px;border-radius:20px;font-weight:600;letter-spacing:.5px;">FREE &amp; OPEN SOURCE</span>
      </div>
      <div style="border-top:1px solid #333;padding-top:16px;margin-bottom:14px;margin-top:6px;">
        <div style="font-size:12px;color:#888;margin-bottom:10px;">Links</div>
        <div style="display:flex;flex-direction:column;gap:8px;align-items:center;">
          <a href="https://www.anchorcastapp.com" target="_blank" rel="noopener"
             style="display:inline-flex;align-items:center;gap:6px;color:#c9a84c;font-size:13px;text-decoration:none;">
            🌐 anchorcastapp.com
          </a>
          <a href="https://github.com/anchorcastapp-team/anchorcastapp" target="_blank" rel="noopener"
             style="display:inline-flex;align-items:center;gap:6px;color:#7aa2f7;font-size:13px;text-decoration:none;">
            <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"/></svg>
            GitHub Repository
          </a>
          <a href="https://github.com/anchorcastapp-team/anchorcastapp/wiki" target="_blank" rel="noopener"
             style="display:inline-flex;align-items:center;gap:6px;color:#7aa2f7;font-size:13px;text-decoration:none;">
            📚 Documentation Wiki
          </a>
          <a href="mailto:info@kaitamtech.com"
             style="display:inline-flex;align-items:center;gap:6px;color:#888;font-size:13px;text-decoration:none;">
            ✉ info@kaitamtech.com
          </a>
          <a href="https://github.com/anchorcastapp-team/anchorcastapp/issues" target="_blank" rel="noopener"
             style="display:inline-flex;align-items:center;gap:6px;color:#888;font-size:13px;text-decoration:none;">
            🐛 Report an Issue
          </a>
        </div>
      </div>
      <div style="margin-top:16px;">
        <button onclick="document.getElementById('aboutOverlay').classList.remove('show')"
                style="background:#333;color:#fff;border:none;padding:8px 28px;border-radius:6px;cursor:pointer;font-size:13px;">Close</button>
      </div>
    </div>`;
  overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.classList.remove('show'); });
  document.body.appendChild(overlay);
}

// ─── GET STARTED WIZARD ───────────────────────────────────────────────────────
let _gsStep = 0;
const _GS_STEPS = [
  {
    title: 'Welcome',
    icon: '✝', iconColor: '#c9a84c',
    heading: 'Welcome to AnchorCast',
    body: `AnchorCast listens to your preacher in real-time, automatically detects Bible references, and displays them on your projection screen.<br><br>
      Use this guide to get set up in a few minutes. You can revisit it anytime from <strong>Help → Get Started / Welcome</strong>.`,
    tip: 'AnchorCast works offline with local Whisper, or online with Deepgram for best accuracy.',
    buttons: [
      { id: '_gsBtnGuide', label: '📖 Get Started — Step by Step Guide', style: 'primary', action: 'next' },
      { id: '_gsBtnDocs',  label: '📄 Open Full Documentation',          style: 'secondary', action: 'docs' },
    ],
  },
  {
    title: 'Software Layout',
    icon: '🖥', iconColor: '#4a9ee8',
    heading: 'Understanding the Interface',
    body: `<strong style="color:#e0e0e0">Schedule (left)</strong> — Drag songs, verses, media, and presentations for the service into order.<br><br>
      <strong style="color:#e0e0e0">Program Preview (centre)</strong> — See what the operator is previewing before sending it live.<br><br>
      <strong style="color:#e0e0e0">Live Display (right)</strong> — What is currently showing on the projection screen.<br><br>
      <strong style="color:#e0e0e0">AI Detections (bottom-right)</strong> — Bible verses the AI automatically detected from the live sermon.`,
    tip: 'Drag the column dividers left or right to resize panels to suit your workflow.',
  },
  {
    title: 'Bible Setup',
    icon: '📖', iconColor: '#c9a84c',
    heading: 'Install Bible Translations',
    body: `KJV is bundled by default. To add more translations:<br><br>
      1. Open <strong>Settings → Bible Versions</strong><br>
      2. Paste or import a JSON Bible file<br>
      3. The translation appears instantly in the search bar<br><br>
      Switch translations at any time using the <strong>KJV</strong> dropdown in the top-right toolbar.`,
    tip: 'Supported formats: OSIS JSON, AnchorCast Bible JSON.',
  },
  {
    title: 'Audio & Transcription',
    icon: '🎤', iconColor: '#e74c3c',
    heading: 'Set Up Live Transcription',
    body: `AnchorCast transcribes the preacher's voice in real-time to auto-detect Bible verses.<br><br>
      1. Go to <strong>Settings → Audio & Transcription</strong><br>
      2. Select your microphone or audio interface<br>
      3. Choose <strong>Online</strong> (Deepgram — best) or <strong>Offline</strong> (local Whisper)<br>
      4. Press <strong>Start Transcript</strong> (top-left) to begin<br><br>
      Detected verses appear in the <strong>AI Detections</strong> panel automatically.`,
    tip: 'Online mode requires a free Deepgram API key — add it in Settings.',
  },
  {
    title: 'Multi-Monitor Setup',
    icon: '📺', iconColor: '#27ae60',
    heading: 'Open the Projection Window',
    body: `Connect your second screen (projector or TV), then:<br><br>
      1. Press <kbd style="background:#333;padding:2px 7px;border-radius:4px;font-size:12px">Ctrl+P</kbd> or go to <strong>Display → Open Projection</strong><br>
      2. Drag the projection window to your second screen<br>
      3. Press <kbd style="background:#333;padding:2px 7px;border-radius:4px;font-size:12px">F11</kbd> to make it fullscreen<br><br>
      The screen stays black until you press <strong>Project Live</strong>.`,
    tip: 'Use Ctrl+Shift+O to open the Operator Command Center — a compact floating live-control panel.',
  },
  {
    title: 'Build a Schedule',
    icon: '📋', iconColor: '#9b59b6',
    heading: 'Plan Your Service',
    body: `Add items to the schedule before the service starts:<br><br>
      • <strong>Songs</strong> — Search the Songs tab, click <strong>+Schedule</strong><br>
      • <strong>Bible verses</strong> — Search a reference, click <strong>+Schedule</strong><br>
      • <strong>Media</strong> — Import images or videos from the Media tab<br>
      • <strong>Section dividers</strong> — Click the § button to add labelled sections<br><br>
      Drag items to reorder. Click <strong>▶ Auto</strong> to auto-advance through the schedule.`,
    tip: 'Save your whole setup — schedule, theme, logo — as a Preset for next Sunday.',
  },
  {
    title: 'Go Live',
    icon: '🔴', iconColor: '#e74c3c',
    heading: 'Present to Your Congregation',
    body: `When you're ready to display content on screen:<br><br>
      1. Click any verse, song slide, or media item to <strong>preview</strong> it<br>
      2. Press <strong style="color:#e74c3c">Project Live</strong> (top-right) to send it to the screen<br>
      3. Use <strong>← →</strong> arrows to navigate between slides<br>
      4. Press <strong>Clear Screen</strong> to blank the display between items<br><br>
      The <strong>Remote Control</strong> feature lets your phone control the display over Wi-Fi.`,
    tip: 'Shortcuts: Ctrl+L = go live, Ctrl+← / Ctrl+→ = navigate, Ctrl+Backspace = clear.',
  },
];

async function maybeShowGetStarted() {
  // First-run: show the simple welcome card, not the full wizard
  const s = State.settings || {};
  if (s.hideGetStarted) return;
  if (document.getElementById('getStartedOverlay')) return;
  _gsShowSimple();
}

function showGetStarted() {
  // Help menu: always show the full step-by-step wizard
  const existing = document.getElementById('getStartedOverlay');
  if (existing) existing.remove();
  _gsStep = 0;
  _gsCreateOverlay();
  _gsRender();
}

// ── Simple first-run card ─────────────────────────────────────────────────────
function _gsShowSimple() {
  const ov = document.createElement('div');
  ov.id = 'getStartedOverlay';
  ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.78);display:flex;align-items:center;justify-content:center;z-index:9999;';
  ov.innerHTML = `
    <div style="background:#1a1a2e;border:1px solid #2a2a42;border-radius:14px;padding:36px 40px;
      max-width:500px;width:92%;color:#e0e0e0;font-family:inherit;text-align:center;
      box-shadow:0 8px 48px rgba(0,0,0,.65);">
      <div style="font-size:36px;margin-bottom:8px;color:#c9a84c">✝</div>
      <h2 style="margin:0 0 4px;font-size:24px;color:#fff;font-weight:700">Welcome to AnchorCast</h2>
      <div style="font-size:12px;color:#666;margin-bottom:22px;letter-spacing:.04em">AI-powered worship presentation for churches</div>
      <div style="text-align:left;margin-bottom:22px;display:flex;flex-direction:column;gap:12px">
        <div style="display:flex;align-items:flex-start;gap:10px">
          <span style="color:#c9a84c;font-weight:700;flex-shrink:0">1.</span>
          <span style="font-size:13px;color:#bbb;line-height:1.6"><strong style="color:#ddd">Set up your Bible</strong> — Go to <strong>Settings → Bible Versions</strong>. KJV is bundled by default.</span>
        </div>
        <div style="display:flex;align-items:flex-start;gap:10px">
          <span style="color:#c9a84c;font-weight:700;flex-shrink:0">2.</span>
          <span style="font-size:13px;color:#bbb;line-height:1.6"><strong style="color:#ddd">Choose your microphone</strong> — Select your audio input in <strong>Settings → Audio</strong> for live transcription.</span>
        </div>
        <div style="display:flex;align-items:flex-start;gap:10px">
          <span style="color:#c9a84c;font-weight:700;flex-shrink:0">3.</span>
          <span style="font-size:13px;color:#bbb;line-height:1.6"><strong style="color:#ddd">Open projection window</strong> — Press <kbd style="background:#222;padding:1px 6px;border-radius:3px;font-size:11px;border:1px solid #444">Ctrl+P</kbd> and drag it to your second screen.</span>
        </div>
      </div>
      <button id="_gsSimpleWizard" style="background:linear-gradient(135deg,#c9a84c,#a88a30);color:#000;
        border:none;padding:9px 26px;border-radius:7px;cursor:pointer;font-size:13px;font-weight:600;
        margin-bottom:16px;width:100%">Get Started — Step by Step Guide ›</button>
      <label style="display:flex;align-items:center;justify-content:center;gap:7px;cursor:pointer;
        font-size:12px;color:#666;margin-bottom:14px">
        <input type="checkbox" id="_gsSimpleChk" style="accent-color:#c9a84c;width:13px;height:13px;cursor:pointer">
        Show on startup
      </label>
      <button id="_gsSimpleClose" style="background:#222;color:#666;border:1px solid #2a2a3a;
        padding:7px 28px;border-radius:6px;cursor:pointer;font-size:12px">Skip for now</button>
    </div>`;
  document.body.appendChild(ov);

  // Default: checked = show on startup
  const chk = ov.querySelector('#_gsSimpleChk');
  chk.checked = !State.settings?.hideGetStarted;

  chk.addEventListener('change', async e => {
    const hide = !e.target.checked;
    State.settings = { ...(State.settings||{}), hideGetStarted: hide };
    try {
      if (window.electronAPI?.saveSettings) {
        const s = await window.electronAPI.getSettings();
        await window.electronAPI.saveSettings({ ...s, hideGetStarted: hide });
      }
    } catch(_) {}
  });

  ov.querySelector('#_gsSimpleClose').addEventListener('click', () => {
    ov.style.transition = 'opacity .15s'; ov.style.opacity = '0';
    setTimeout(() => ov.remove(), 160);
  });

  ov.querySelector('#_gsSimpleWizard').addEventListener('click', () => {
    ov.remove();
    _gsStep = 0;
    _gsCreateOverlay();
    _gsRender();
  });

  ov.addEventListener('click', e => {
    if (e.target === ov) { ov.style.transition='opacity .15s'; ov.style.opacity='0'; setTimeout(()=>ov.remove(),160); }
  });
}

// ── Wizard overlay (persistent — never removed during navigation) ──────────────
let _gsOverlay = null;

function _gsCreateOverlay() {
  _gsOverlay = document.createElement('div');
  _gsOverlay.id = 'getStartedOverlay';
  _gsOverlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.82);display:flex;align-items:center;justify-content:center;z-index:9999;';
  // Inner wrapper — only this is replaced on navigation, not the backdrop
  _gsOverlay.innerHTML = `<div id="_gsWrap"></div>`;
  document.body.appendChild(_gsOverlay);
  _gsOverlay.addEventListener('click', e => { if (e.target === _gsOverlay) _gsClose(); });
}

function _gsRender() {
  const wrap = document.getElementById('_gsWrap');
  if (!wrap) return; // overlay not created yet

  const step   = _GS_STEPS[_gsStep];
  const isLast = _gsStep === _GS_STEPS.length - 1;
  const showOnStartup = !State.settings?.hideGetStarted;

  // Replace ONLY the inner wrap — the backdrop stays in the DOM so the page never repaints
  wrap.innerHTML = `
    <style>
      #_gsW{display:flex;width:860px;max-width:96vw;height:540px;max-height:92vh;
        background:#12121f;border-radius:12px;overflow:hidden;
        box-shadow:0 24px 80px rgba(0,0,0,.75);border:1px solid #252538;font-family:inherit}
      #_gsS{width:200px;flex-shrink:0;background:#0d0d1a;border-right:1px solid #1a1a2e;
        display:flex;flex-direction:column;padding:0}
      #_gsSL{padding:18px 16px 14px;border-bottom:1px solid #1a1a2e;
        font-size:12px;font-weight:700;color:#c9a84c;letter-spacing:.15em;text-transform:uppercase}
      ._gi{display:flex;align-items:center;gap:8px;padding:8px 16px;cursor:pointer;
        font-size:12px;color:#444;transition:color .1s,background .1s;border-left:3px solid transparent;line-height:1.3}
      ._gi:hover{color:#777;background:rgba(255,255,255,.03)}
      ._gi.active{color:#e8e8f0;background:rgba(201,168,76,.09);border-left-color:#c9a84c;font-weight:600}
      ._gi.done{color:#555}._gi.done ._gm{color:#c9a84c}
      ._gm{width:14px;font-size:11px;flex-shrink:0;text-align:center;opacity:.7}
      #_gR{flex:1;display:flex;flex-direction:column;min-width:0;overflow:hidden}
      #_gT{padding:16px 24px 0;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
      #_gTL{font-size:10px;font-weight:700;color:#2a2a40;letter-spacing:.18em;text-transform:uppercase}
      #_gSU{display:flex;align-items:center;gap:6px;font-size:11px;color:#555;cursor:pointer;user-select:none}
      #_gSU input{accent-color:#c9a84c;cursor:pointer;width:13px;height:13px}
      #_gM{flex:1;padding:18px 24px 12px;overflow-y:auto}
      #_gI{font-size:42px;margin-bottom:10px;line-height:1}
      #_gH{font-size:18px;font-weight:700;color:#fff;margin:0 0 14px}
      #_gC{font-size:13px;color:#999;line-height:1.85}
      #_gTp{margin-top:14px;padding:9px 13px;background:rgba(201,168,76,.07);
        border-left:3px solid rgba(201,168,76,.35);border-radius:0 5px 5px 0;
        font-size:11.5px;color:#777;line-height:1.6}
      #_gD{display:flex;gap:5px;justify-content:center;padding:8px 0;flex-shrink:0}
      ._gd{width:7px;height:7px;border-radius:50%;background:#1e1e32;cursor:pointer;
        transition:background .1s,transform .1s;border:none;padding:0}
      ._gd.on{background:#c9a84c;transform:scale(1.25)}._gd:hover{background:#555}
      #_gF{padding:12px 24px 18px;display:flex;align-items:center;justify-content:space-between;
        border-top:1px solid #1a1a2e;flex-shrink:0;gap:8px}
      ._gb{padding:8px 20px;border-radius:6px;font-size:12px;font-weight:600;
        cursor:pointer;border:none;font-family:inherit;transition:opacity .1s}
      ._gb:hover{opacity:.82}
      #_gBX{background:#1a1a2e;color:#666}
      #_gBP{background:#1a1a2e;color:#666;min-width:70px}
      #_gBP:disabled{opacity:.22;cursor:default}
      #_gBN{background:linear-gradient(135deg,#c9a84c,#a88a30);color:#000;min-width:130px}
    </style>
    <div id="_gsW">
      <div id="_gsS">
        <div id="_gsSL">✝ AnchorCast</div>
        ${_GS_STEPS.map((s,i) => `
          <div class="_gi ${i===_gsStep?'active':i<_gsStep?'done':''}" data-i="${i}">
            <span class="_gm">${i<_gsStep?'✓':i===_gsStep?'›':''}</span>
            ${s.title}
          </div>`).join('')}
      </div>
      <div id="_gR">
        <div id="_gT">
          <div id="_gTL">Getting Started</div>
          <label id="_gSU">
            <input type="checkbox" id="_gSUc" ${showOnStartup?'checked':''}>
            Show on startup
          </label>
        </div>
        <div id="_gM">
          <div id="_gI" style="color:${step.iconColor}">${step.icon}</div>
          <div id="_gH">${step.heading}</div>
          <div id="_gC">${step.body}</div>
          ${step.buttons ? `<div id="_gBtns" style="display:flex;flex-direction:column;gap:9px;margin-top:20px">
            ${step.buttons.map(b => `
              <button id="${b.id}" style="
                padding:11px 20px;border-radius:8px;font-size:13px;font-weight:600;
                cursor:pointer;border:none;font-family:inherit;text-align:left;
                transition:opacity .15s;
                ${b.style==='primary'
                  ? 'background:linear-gradient(135deg,#c9a84c,#a88a30);color:#000;'
                  : 'background:#1e1e32;color:#aaa;border:1px solid #2a2a3e;'}
              ">${b.label}</button>
            `).join('')}
          </div>` : ''}
          ${step.tip?`<div id="_gTp">💡 ${step.tip}</div>`:''}
        </div>
        <div id="_gD">
          ${_GS_STEPS.map((_,i)=>`<button class="_gd${i===_gsStep?' on':''}" data-d="${i}" aria-label="Step ${i+1}"></button>`).join('')}
        </div>
        <div id="_gF">
          <button class="_gb" id="_gBX">Close</button>
          <div style="display:flex;gap:8px">
            <button class="_gb" id="_gBP" ${_gsStep===0?'disabled':''}>‹ Prev</button>
            <button class="_gb" id="_gBN">
              ${isLast ? 'Finish ✓' : 'Next: '+_GS_STEPS[_gsStep+1].title+' ›'}
            </button>
          </div>
        </div>
      </div>
    </div>`;

  // Sidebar
  wrap.querySelectorAll('._gi').forEach(el =>
    el.addEventListener('click', () => { _gsStep = +el.dataset.i; _gsRender(); })
  );
  // Dots
  wrap.querySelectorAll('._gd').forEach(el =>
    el.addEventListener('click', () => { _gsStep = +el.dataset.d; _gsRender(); })
  );
  // Prev
  wrap.querySelector('#_gBP').addEventListener('click', () => {
    if (_gsStep > 0) { _gsStep--; _gsRender(); }
  });
  // Next / Finish
  wrap.querySelector('#_gBN').addEventListener('click', () => {
    if (isLast) { _gsClose(); } else { _gsStep++; _gsRender(); }
  });
  // Close
  wrap.querySelector('#_gBX').addEventListener('click', _gsClose);
  // Step action buttons (Welcome screen)
  wrap.querySelector('#_gsBtnGuide')?.addEventListener('click', () => {
    _gsStep = 1; _gsRender();
  });
  wrap.querySelector('#_gsBtnDocs')?.addEventListener('click', () => {
    if (window.electronAPI?.openHelpWindow) {
      window.electronAPI.openHelpWindow();
    } else {
      window.open('/help.html', '_blank');
    }
  });
  // Show on startup
  wrap.querySelector('#_gSUc').addEventListener('change', async e => {
    const hide = !e.target.checked;
    State.settings = { ...(State.settings||{}), hideGetStarted: hide };
    try {
      if (window.electronAPI?.saveSettings) {
        const saved = await window.electronAPI.getSettings();
        await window.electronAPI.saveSettings({ ...saved, hideGetStarted: hide });
      }
    } catch(_) {}
  });
}


async function _gsClose() {
  const ov = document.getElementById('getStartedOverlay');
  if (!ov) return;
  ov.style.transition = 'opacity .18s ease';
  ov.style.opacity = '0';
  setTimeout(() => { if (ov.parentNode) ov.remove(); }, 200);
}

// ─── MODAL HELPERS ────────────────────────────────────────────────────────────
function showModal(id) {
  const el = document.getElementById(id);
  el.classList.add('show');
}
function closeModal(id) {
  const el = document.getElementById(id);
  el.classList.remove('show');
}

// ─── TOAST ────────────────────────────────────────────────────────────────────
let toastTimer = null;
function toast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 2600);
}

// ─── THEME LIBRARY TAB ────────────────────────────────────────────────────────

let _themeCategory = 'song'; // active category in Themes tab
let _themeFilter   = '';
let _themeCtxId    = null;   // theme id under right-click

// Built-in themes mirrored here so the grid works without opening the full designer
const _defaultBox = (overrides) => ({
  text:'', bold:false, italic:false, lineSpacing:1.4, textTransform:'none', fontWeight:400, letterSpacing:0,
  shadow:true, shadowColor:'#000000', shadowBlur:8, shadowOffsetX:0, shadowOffsetY:2,
  bgFill:'', bgOpacity:0, borderW:0, borderColor:'#ffffff', borderRadius:0, ...overrides
});

const THEME_BUILTINS = [
  { id:'song_sanctuary', category:'song', name:'Sanctuary', builtIn:true,
    bgType:'solid', bgColor1:'#0a0a1e', textColor:'#fff', accentColor:'#c9a84c', fontFamily:'Crimson Pro', fontSize:56,
    boxes:[
      _defaultBox({ role:'main', x:80, y:200, w:1760, h:680, text:'song text line one\nsong text line two\nsong text line three',
        fontFamily:'DM Sans', fontSize:72, color:'#ffffff', align:'center', valign:'center', bold:true, fontWeight:700 })
    ]},
  { id:'song_dawn', category:'song', name:'Dawn Worship', builtIn:true,
    bgType:'radial', bgColor1:'#1a0c08', bgColor2:'#080010', textColor:'#f0e8d8', accentColor:'#e8904c', fontFamily:'Crimson Pro', fontSize:56,
    boxes:[
      _defaultBox({ role:'main', x:80, y:200, w:1760, h:680, text:'song text line one\nsong text line two\nsong text line three',
        fontFamily:'Crimson Pro', fontSize:72, color:'#f0e8d8', align:'center', valign:'center', bold:true, fontWeight:700 })
    ]},
  { id:'song_minimal', category:'song', name:'Minimal', builtIn:true,
    bgType:'solid', bgColor1:'#000', textColor:'#fff', accentColor:'#aaa', fontFamily:'DM Sans', fontSize:56,
    boxes:[
      _defaultBox({ role:'main', x:80, y:200, w:1760, h:680, text:'song text line one\nsong text line two\nsong text line three',
        fontFamily:'DM Sans', fontSize:72, color:'#ffffff', align:'center', valign:'center', bold:true, fontWeight:700 })
    ]},
  { id:'song_midnight', category:'song', name:'Midnight Blue', builtIn:true,
    bgType:'linear', bgColor1:'#001040', bgColor2:'#000820', textColor:'#e0ecff', accentColor:'#4c9ae8', fontFamily:'DM Sans', fontSize:54,
    boxes:[
      _defaultBox({ role:'main', x:80, y:200, w:1760, h:680, text:'song text line one\nsong text line two\nsong text line three',
        fontFamily:'DM Sans', fontSize:72, color:'#e0ecff', align:'center', valign:'center', bold:true, fontWeight:700 })
    ]},
  { id:'sanctuary', category:'scripture', name:'Sanctuary', builtIn:true,
    bgType:'radial', bgColor1:'#10082a', bgColor2:'#000518', textColor:'#ede6d8', accentColor:'#c9a84c', refColor:'#c9a84c', fontFamily:'Crimson Pro', fontSize:48,
    boxes:[
      _defaultBox({ role:'main', x:60, y:60, w:1800, h:960,
        text:'Genesis 1:1-2 (KJV)\nIn the beginning God created the heaven and the earth.',
        fontFamily:'Crimson Pro', fontSize:56, color:'#ede6d8', align:'center', valign:'center', lineSpacing:1.7,
        refFontFamily:'Cinzel', refFontSize:52, refColor:'#c9a84c', refBold:true, refLineSpacing:1.4 })
    ]},
  { id:'dawn', category:'scripture', name:'Dawn', builtIn:true,
    bgType:'radial', bgColor1:'#1a0c08', bgColor2:'#080010', textColor:'#f0e8d8', accentColor:'#e8904c', refColor:'#e8904c', fontFamily:'Crimson Pro', fontSize:48,
    boxes:[
      _defaultBox({ role:'main', x:60, y:60, w:1800, h:960,
        text:'Genesis 1:1-2 (KJV)\nIn the beginning God created the heaven and the earth.',
        fontFamily:'Crimson Pro', fontSize:56, color:'#f0e8d8', align:'center', valign:'center', lineSpacing:1.7,
        refFontFamily:'Cinzel', refFontSize:52, refColor:'#e8904c', refBold:true, refLineSpacing:1.4 })
    ]},
  { id:'deep', category:'scripture', name:'Deep Water', builtIn:true,
    bgType:'radial', bgColor1:'#001028', bgColor2:'#000408', textColor:'#e0ecf8', accentColor:'#4c9ae8', refColor:'#4c9ae8', fontFamily:'Crimson Pro', fontSize:48,
    boxes:[
      _defaultBox({ role:'main', x:60, y:60, w:1800, h:960,
        text:'Genesis 1:1-2 (KJV)\nIn the beginning God created the heaven and the earth.',
        fontFamily:'Crimson Pro', fontSize:56, color:'#e0ecf8', align:'center', valign:'center', lineSpacing:1.7,
        refFontFamily:'Cinzel', refFontSize:52, refColor:'#4c9ae8', refBold:true, refLineSpacing:1.4 })
    ]},
  { id:'minimal', category:'scripture', name:'Minimal', builtIn:true,
    bgType:'solid', bgColor1:'#000', textColor:'#fff', accentColor:'#aaa', refColor:'#aaa', fontFamily:'DM Sans', fontSize:52,
    boxes:[
      _defaultBox({ role:'main', x:60, y:60, w:1800, h:960,
        text:'Genesis 1:1-2 (KJV)\nIn the beginning God created the heaven and the earth.',
        fontFamily:'DM Sans', fontSize:52, color:'#ffffff', align:'center', valign:'center', lineSpacing:1.7,
        refFontFamily:'DM Sans', refFontSize:48, refColor:'#aaaaaa', refBold:true, refLineSpacing:1.4 })
    ]},
  { id:'pres_classic', category:'presentation', name:'Classic Dark', builtIn:true,
    bgType:'solid', bgColor1:'#0a0a1e', titleColor:'#fff', subtitleColor:'#c9a84c', fontFamily:'Cinzel', titleSize:72,
    boxes:[
      _defaultBox({ role:'title', x:120, y:280, w:1680, h:320, text:'PRESENTATION TITLE',
        fontFamily:'Cinzel', fontSize:80, color:'#ffffff', align:'center', valign:'center', bold:true, fontWeight:700, lineSpacing:1.2 }),
      _defaultBox({ role:'subtitle', x:120, y:640, w:1680, h:200, text:'Subtitle or verse reference',
        fontFamily:'Cinzel', fontSize:40, color:'#c9a84c', align:'center', valign:'center', italic:true, lineSpacing:1.3 })
    ]},
  { id:'pres_modern', category:'presentation', name:'Modern', builtIn:true,
    bgType:'linear', bgColor1:'#1a1a2e', bgColor2:'#16213e', titleColor:'#e8e8f0', subtitleColor:'#4c9ae8', fontFamily:'DM Sans', titleSize:64,
    boxes:[
      _defaultBox({ role:'title', x:120, y:280, w:1680, h:320, text:'PRESENTATION TITLE',
        fontFamily:'DM Sans', fontSize:72, color:'#e8e8f0', align:'center', valign:'center', bold:true, fontWeight:700, lineSpacing:1.2 }),
      _defaultBox({ role:'subtitle', x:120, y:640, w:1680, h:200, text:'Subtitle or verse reference',
        fontFamily:'DM Sans', fontSize:36, color:'#4c9ae8', align:'center', valign:'center', lineSpacing:1.3 })
    ]},
  { id:'pres_worship', category:'presentation', name:'Worship', builtIn:true,
    bgType:'radial', bgColor1:'#1a0a08', bgColor2:'#080010', titleColor:'#f0e8d8', subtitleColor:'#e8904c', fontFamily:'Cinzel', titleSize:68,
    boxes:[
      _defaultBox({ role:'title', x:120, y:280, w:1680, h:320, text:'PRESENTATION TITLE',
        fontFamily:'Cinzel', fontSize:72, color:'#f0e8d8', align:'center', valign:'center', bold:true, fontWeight:700, lineSpacing:1.2 }),
      _defaultBox({ role:'subtitle', x:120, y:640, w:1680, h:200, text:'Subtitle or verse reference',
        fontFamily:'Cinzel', fontSize:38, color:'#e8904c', align:'center', valign:'center', italic:true, lineSpacing:1.3 })
    ]},
];

function _getAllThemes() {
  // Prefer State.themes (full data from disk, may have bgImage/bgVideo)
  // Fall back to THEME_BUILTINS for any that aren't in State.themes
  const stateMap = new Map((State.themes || []).map(t => [t.id, t]));
  const result = [];
  // Add all State.themes first (they have full data)
  (State.themes || []).forEach(t => result.push(t));
  // Add any THEME_BUILTINS not already in State.themes
  THEME_BUILTINS.forEach(bt => {
    if (!stateMap.has(bt.id)) result.unshift(bt);
  });
  return result;
}

function setThemeCategory(cat) {
  _themeCategory = cat;
  ['song','scripture','presentation'].forEach(c => {
    document.getElementById('tcat-'+c)?.classList.toggle('active', c === cat);
  });
  const ctxLabel = document.getElementById('themeCtxLabel');
  if (ctxLabel) ctxLabel.textContent = cat.charAt(0).toUpperCase() + cat.slice(1);
  renderThemeGrid();
  // Refresh Program Preview to show this category's default theme
  if (State.currentTab === 'theme') {
    const previewDisplay = document.getElementById('previewDisplay');
    if (previewDisplay) {
      previewDisplay.innerHTML = '';  // FIX: clear stale verse content
      previewDisplay.style.display = 'flex';
      _renderThemePreview(previewDisplay);
    }
  }
}

function filterThemeGrid(val) {
  _themeFilter = val;
  renderThemeGrid();
}

async function loadThemeGrid() {
  if (window.electronAPI) {
    try {
      const loaded = await window.electronAPI.getThemes() || [];
      // Merge: loaded themes (may have full bgImage data) + any THEME_BUILTINS not in loaded
      const loadedIds = new Set(loaded.map(t => t.id));
      const missingBuiltins = THEME_BUILTINS.filter(b => !loadedIds.has(b.id));
      State.themes = [...missingBuiltins, ...loaded];
    } catch(e) {
      State.themes = [...THEME_BUILTINS];
    }
  } else {
    State.themes = [...THEME_BUILTINS];
  }
  renderThemeGrid();
  // Show preview of current category default theme (only if still on theme tab)
  if (State.currentTab === 'theme') {
    const previewDisplay = document.getElementById('previewDisplay');
    if (previewDisplay) {
      previewDisplay.innerHTML = '';  // FIX: clear any stale verse content
      previewDisplay.style.display = 'flex';
      _renderThemePreview(previewDisplay);
    }
  }
}

function _themeBgStyle(t) {
  if (!t) return '#111';
  if (t.bgType === 'image' && t.bgImage) {
    // bgImage may be base64 or a file path - both work as CSS url()
    return `url('${t.bgImage}') center/cover no-repeat`;
  }
  if (t.bgType === 'video') {
    // Can't show video frame as CSS, use a teal-tinted dark gradient as hint
    return 'linear-gradient(135deg,#001a28 0%,#002030 50%,#000a0f 100%)';
  }
  if (t.bgType === 'solid')  return t.bgColor1 || '#111';
  if (t.bgType === 'linear') return `linear-gradient(135deg,${t.bgColor1||'#111'},${t.bgColor2||'#000'})`;
  return `radial-gradient(ellipse at 35% 45%,${t.bgColor1||'#111'} 0%,${t.bgColor2||'#000'} 70%,#000 100%)`;
}

function renderThemeGrid() {
  const grid  = document.getElementById('themeGrid');
  const empty = document.getElementById('themeGridEmpty');
  if (!grid) return;

  grid.querySelectorAll('.theme-grid-thumb').forEach(e => e.remove());

  const defaultId = _themeCategory === 'song' ? State.currentSongTheme
    : _themeCategory === 'presentation' ? State.currentPresTheme
    : State.currentTheme;

  const all = _getAllThemes().filter(t =>
    t.category === _themeCategory &&
    (!_themeFilter || t.name.toLowerCase().includes(_themeFilter.toLowerCase()))
  );

  if (!all.length) {
    if (empty) empty.style.display = '';
    return;
  }
  if (empty) empty.style.display = 'none';

  all.forEach(t => {
    const isDefault = t.id === defaultId;
    const card = document.createElement('div');
    card.className = 'theme-grid-thumb' + (isDefault ? ' active' : '');
    card.dataset.tid = t.id;

    // Canvas preview — 16:9 mini render
    const canvas = document.createElement('div');
    canvas.className = 'theme-grid-canvas';
    canvas.style.background = _themeBgStyle(t);
    canvas.style.position = 'relative';
    canvas.style.overflow = 'hidden';
    if (t.bgType === 'image' && t.bgImage) {
      const imgUrl = _normalizeThemeMediaUrl(t.bgImage);
      canvas.style.backgroundImage = `url('${imgUrl}')`;
      canvas.style.backgroundSize = 'cover';
      canvas.style.backgroundPosition = 'center';
      canvas.style.backgroundRepeat = 'no-repeat';
    }
    if (t.bgType === 'video' && t.bgVideo) {
      const v = document.createElement('video');
      v.muted = true; v.defaultMuted = true; v.loop = true; v.autoplay = true; v.playsInline = true;
      v.setAttribute('muted','');
      v.src = _normalizeThemeMediaUrl(t.bgVideo);
      v.preload = 'metadata';
      v.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;object-fit:cover;z-index:0;background:#000';
      canvas.appendChild(v);
      v.play().catch(()=>{});
    }

    // Video/image badge
    if (t.bgType === 'video') {
      const vb = document.createElement('div');
      vb.style.cssText = 'position:absolute;top:3px;right:3px;font-size:8px;background:rgba(0,40,60,.8);color:#4cf;padding:1px 5px;border-radius:3px;font-weight:700;z-index:2';
      vb.textContent = '▶ Video';
      canvas.appendChild(vb);
    } else if (t.bgType === 'image' && t.bgImage) {
      const ib = document.createElement('div');
      ib.style.cssText = 'position:absolute;top:3px;right:3px;font-size:8px;background:rgba(0,0,0,.7);color:#afc;padding:1px 5px;border-radius:3px;z-index:2';
      ib.textContent = '🖼 Image';
      canvas.appendChild(ib);
    }

    // Sample text inside preview
    const preview = document.createElement('div');
    preview.style.cssText = 'position:relative;z-index:1;text-align:center;padding:6px;width:100%;overflow:hidden;background:linear-gradient(180deg,rgba(0,0,0,.08),rgba(0,0,0,.28))';

    if (t.category === 'presentation') {
      preview.innerHTML = `
        <div style="font-size:9px;font-weight:700;color:${t.titleColor||'#fff'};font-family:'${t.fontFamily||'Cinzel'}',serif;line-height:1.2;text-shadow:0 1px 3px rgba(0,0,0,.8)">LOVE OF GOD</div>
        <div style="font-size:6px;color:${t.subtitleColor||'#c9a84c'};font-family:'${t.fontFamily||'Cinzel'}',serif;margin-top:3px;font-style:italic">Sunday Service</div>`;
    } else if (t.category === 'song') {
      preview.innerHTML = `
        <div style="font-size:6px;color:${t.accentColor||'#c9a84c'};letter-spacing:.1em;text-transform:uppercase;margin-bottom:3px;font-family:'${t.fontFamily||'Cinzel'}',serif">Amazing Grace · Verse 1</div>
        <div style="font-size:8px;color:${t.textColor||'#fff'};font-family:'${t.fontFamily||'Cinzel'}',serif;line-height:1.3;text-shadow:0 1px 3px rgba(0,0,0,.8)">Amazing grace<br>how sweet the sound</div>`;
    } else {
      preview.innerHTML = `
        <div style="font-size:6px;color:${t.accentColor||'#c9a84c'};letter-spacing:.1em;text-transform:uppercase;margin-bottom:3px;font-family:'${t.fontFamily||'Crimson Pro'}',serif">John 3:16</div>
        <div style="font-size:7px;color:${t.textColor||'#ede6d8'};font-family:'${t.fontFamily||'Crimson Pro'}',serif;line-height:1.4;text-shadow:0 1px 3px rgba(0,0,0,.8)">For God so loved<br>the world…</div>`;
    }

    canvas.appendChild(preview);
    card.appendChild(canvas);

    // Label
    const label = document.createElement('div');
    label.className = 'theme-grid-label';
    label.textContent = t.name;
    card.appendChild(label);

    const tag = document.createElement('div');
    tag.className = 'theme-grid-tag';
    tag.textContent = t.builtIn ? 'Built-in' : 'Custom';
    card.appendChild(tag);

    // Click → preview only (does NOT set as default)
    card.addEventListener('click', () => selectThemeInGrid(t.id));
    // Right-click → context menu
    card.addEventListener('contextmenu', e => {
      e.preventDefault();
      _themeCtxId = t.id;
      State.currentThemeCtxId = t.id;
      showThemeContextMenu(e);
    });

    grid.appendChild(card);
  });
}

// Render a full-quality theme preview into the Program Preview panel
function _renderThemePreview(container, id) {
  const themeId = id || (
    _themeCategory === 'song'         ? State.currentSongTheme :
    _themeCategory === 'presentation' ? State.currentPresTheme :
    State.currentTheme
  );
  const t = _getAllThemes().find(x => x.id === themeId) ||
            _getAllThemes().find(x => x.category === _themeCategory);
  if (!t) return;

  // Apply background to the PARENT canvas (previewCanvas), not just the display div
  const canvas = document.getElementById('previewCanvas');
  if (canvas) {
    canvas.style.background = '';
    canvas.removeAttribute('data-theme');
    const verseBg = canvas.querySelector('.verse-bg');
    if (verseBg) verseBg.style.display = 'none';
    const emptyEl = document.getElementById('previewEmpty');
    if (emptyEl) emptyEl.style.display = 'none';

    if (t.bgType === 'image' && t.bgImage) {
      const imgUrl = _normalizeThemeMediaUrl(t.bgImage);
      canvas.style.backgroundImage = `url('${imgUrl}')`;
      canvas.style.backgroundSize = 'cover';
      canvas.style.backgroundPosition = 'center';
      canvas.style.backgroundRepeat = 'no-repeat';
    } else if (t.bgType === 'video' && t.bgVideo) {
      canvas.style.backgroundImage = 'none';
      canvas.style.background = '#000';
      // Create / reuse a video element in the canvas
      let vid = document.getElementById('themePreviewVideo');
      if (!vid) {
        vid = document.createElement('video');
        vid.id = 'themePreviewVideo';
        vid.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;object-fit:cover;z-index:0;pointer-events:none';
        vid.loop = true; vid.muted = true; vid.playsInline = true;
        canvas.insertBefore(vid, canvas.firstChild);
      }
      const videoUrl = _normalizeThemeMediaUrl(t.bgVideo);
      if (vid.dataset.mediaUrl !== videoUrl) { vid.src = videoUrl; vid.dataset.mediaUrl = videoUrl; }
      vid.style.display = 'block';
      vid.play().catch(() => {});
    } else {
      // Stop any preview video
      const vid = document.getElementById('themePreviewVideo');
      if (vid) { vid.pause(); vid.src = ''; vid.style.display = 'none'; }
      canvas.style.backgroundImage = 'none';
      if (t.bgType === 'solid') canvas.style.background = t.bgColor1 || '#0a0a1e';
      else if (t.bgType === 'linear') canvas.style.background = `linear-gradient(135deg,${t.bgColor1||'#0a0a1e'},${t.bgColor2||'#000'})`;
      else canvas.style.background = `radial-gradient(ellipse at 35% 45%,${t.bgColor1||'#0a0a1e'} 0%,${t.bgColor2||'#000'} 70%,#000 100%)`;
    }

    // Overlay
    const ovAlpha = (t.bgOverlay || 0) / 100;
    let bgVeil = document.getElementById('themePreviewVeil');
    if (ovAlpha > 0) {
      if (!bgVeil) {
        bgVeil = document.createElement('div');
        bgVeil.id = 'themePreviewVeil';
        bgVeil.style.cssText = 'position:absolute;inset:0;pointer-events:none;z-index:1';
        canvas.insertBefore(bgVeil, canvas.firstChild);
      }
      bgVeil.style.background = `rgba(0,0,0,${ovAlpha})`;
      bgVeil.style.display = 'block';
    } else if (bgVeil) {
      bgVeil.style.display = 'none';
    }
  }

  // Now set content in the container (previewDisplay)
  container.style.cssText = 'display:flex;align-items:center;justify-content:center;' +
    'height:100%;position:relative;z-index:2;overflow:hidden;background:transparent;padding:0;';
  container.style.pointerEvents = 'none';

  // ── Box-based theme (new format) ──
  if (t.boxes && t.boxes.length > 0) {
    const CW = 1920, CH = 1080;
    const cw = canvas?.clientWidth  || container.clientWidth  || 400;
    const ch = canvas?.clientHeight || container.clientHeight || 225;
    const sc = Math.min(cw / CW, ch / CH);

    // Use category-appropriate PLACEHOLDER text — never box.text which may contain
    // live verse content stored when the theme was last applied.
    const _placeholderMain = t.category === 'song'
      ? 'Amazing grace, how sweet the sound\nThat saved a wretch like me'
      : t.category === 'presentation'
      ? 'LOVE OF GOD\nSunday Morning Service'
      : 'For God so loved the world, that He gave His only begotten Son.';
    const _placeholderRef = t.category === 'song'
      ? 'Amazing Grace · Verse 1'
      : t.category === 'presentation'
      ? 'Introduction'
      : 'John 3:16 · KJV';

    const _hasRefBox = t.boxes.some(b => (b.role||'').toLowerCase() === 'ref');
    let boxHtml = '';
    t.boxes.forEach(box => {
      const role = (box.role || '').toLowerCase();
      const isMain = role === 'main';
      const bx = (box.x * sc) + 'px', by = (box.y * sc) + 'px';
      const bw = (box.w * sc) + 'px', bh = (box.h * sc) + 'px';
      const fs = Math.max(7, (box.fontSize || 52) * sc) + 'px';
      const sh = '0 1px 4px rgba(0,0,0,.9)';
      const bgF = box.bgOpacity > 0 && box.bgFill
        ? `background:rgba(${parseInt(box.bgFill.slice(1,3)||'0',16)},${parseInt(box.bgFill.slice(3,5)||'0',16)},${parseInt(box.bgFill.slice(5,7)||'0',16)},${box.bgOpacity/100});` : '';
      const bord = box.borderW > 0 ? `border:${Math.max(1,box.borderW*sc)}px solid ${box.borderColor||'#fff'};` : '';
      const previewText = role === 'ref' ? _placeholderRef : _placeholderMain;
      const _prevRefPos = State.settings?.scriptureRefPosition || 'top';
      const _prevShowRef = State.settings?.scriptureShowReference !== false && _prevRefPos !== 'hidden';
      let inlineRefHtml = '';
      if (isMain && !_hasRefBox && t.category === 'scripture' && _prevShowRef) {
        // Read ref font from Preferences (State.settings) — not from box theme properties
        const _rf = State.settings || {};
        const rFF = _rf.scriptureRefFontFamily || 'Cinzel';
        const rFS = Math.max(7, (_rf.scriptureRefFontSize || 38) * sc) + 'px';
        const rFC = _rf.scriptureRefFontColor || t.refColor || t.accentColor || '#c9a84c';
        const _rfBold = _rf.scriptureRefFontStyle === 'bold' || _rf.scriptureRefFontStyle === 'bold-italic';
        const rFW = _rfBold ? 700 : 500;
        const refMargin = _prevRefPos === 'bottom'
          ? `margin-top:${Math.max(2, 6*sc)}px;margin-bottom:0`
          : `margin-bottom:${Math.max(2, 6*sc)}px`;
        inlineRefHtml = `<div style="font-family:'${rFF}',serif;font-size:${rFS};color:${rFC};font-weight:${rFW};line-height:1.4;${refMargin};text-align:${_rf.scriptureRefTextAlign||box.align||'center'};flex-shrink:0">${escapeHtml(_placeholderRef)}</div>`;
      }
      const mainFlexCss = inlineRefHtml ? 'display:flex;flex-direction:column;' : '';
      const contentHtml = inlineRefHtml
        ? (_prevRefPos === 'bottom'
          ? `<div style="min-height:0">${escapeHtml(previewText)}</div>${inlineRefHtml}`
          : `${inlineRefHtml}<div style="min-height:0">${escapeHtml(previewText)}</div>`)
        : escapeHtml(previewText);
      boxHtml += `<div style="position:absolute;left:${bx};top:${by};width:${bw};height:${bh};
        display:flex;flex-direction:column;align-items:${box.align==='left'?'flex-start':box.align==='right'?'flex-end':'center'};
        justify-content:${box.valign==='top'?'flex-start':box.valign==='bottom'?'flex-end':'center'};
        text-align:${box.align||'center'};overflow:hidden;padding:4px;box-sizing:border-box;${bgF}${bord}">
        <div style="font-size:${fs};color:${box.color||'#fff'};
          font-family:'${box.fontFamily||'DM Sans'}',sans-serif;font-weight:${box.bold?'700':'400'};
          font-style:${box.italic?'italic':'normal'};line-height:${box.lineSpacing||1.4};
          text-shadow:${sh};white-space:pre-wrap;width:100%;text-align:${box.align||'center'};${mainFlexCss}"
        >${contentHtml}</div>
      </div>`;
    });
    container.innerHTML = `<div style="position:absolute;inset:0">${boxHtml}</div>`;
    container.style.position = 'absolute';
    container.style.inset = '0';
    return;
  }

  // ── Legacy property-based theme ──
  const ff  = t.fontFamily || (t.category === 'scripture' ? 'Crimson Pro' : 'Cinzel');
  const col = t.textColor  || (t.category === 'scripture' ? '#ede6d8' : '#fff');
  const acc = t.accentColor || t.refColor || '#c9a84c';
  const fs  = t.fontSize || (t.category === 'presentation' ? (t.titleSize||72) : 48);
  const ls  = t.lineSpacing || 1.5;
  const sh  = '0 2px 10px rgba(0,0,0,.9)';

  let inner = '';
  if (t.category === 'presentation') {
    inner = `<div style="width:100%;text-align:center">
      <div style="font-size:clamp(16px,${(t.titleSize||72)/16}vw,${t.titleSize||72}px);color:${t.titleColor||'#fff'};
        font-family:'${ff}',serif;font-weight:700;text-shadow:${sh};line-height:1.2;margin-bottom:14px">LOVE OF GOD</div>
      <div style="font-size:clamp(10px,${(t.subtitleSize||36)/16}vw,${t.subtitleSize||36}px);color:${t.subtitleColor||acc};
        font-family:'${ff}',serif;font-style:italic;text-shadow:${sh}">Sunday Morning Service</div>
    </div>`;
  } else if (t.category === 'song') {
    inner = `<div style="width:100%;text-align:${t.textAlign||'center'}">
      <div style="font-size:clamp(8px,1.1vw,14px);color:${acc};letter-spacing:.18em;
        text-transform:uppercase;margin-bottom:12px;font-family:'${ff}',serif;opacity:.8">Amazing Grace · Verse 1</div>
      <div style="font-size:clamp(13px,${fs/16}vw,${fs}px);color:${col};font-family:'${ff}',serif;
        font-style:${t.fontStyle||'normal'};line-height:${ls};text-shadow:${sh}">
        Amazing grace, how sweet the sound<br>That saved a wretch like me</div>
    </div>`;
  } else {
    const numHtml = t.showVerseNum !== false
      ? `<span style="color:${acc};vertical-align:super;font-size:.55em;margin-right:4px">16</span>` : '';
    inner = `<div style="width:100%;text-align:${t.textAlign||'center'}">
      <div style="font-size:clamp(9px,1.2vw,15px);color:${t.refColor||acc};letter-spacing:.22em;
        text-transform:uppercase;margin-bottom:12px;font-family:'${ff}',serif;font-weight:500">John 3:16</div>
      <div style="font-size:clamp(12px,${fs/16}vw,${fs}px);color:${col};font-family:'${ff}',serif;
        font-style:${t.fontStyle||'normal'};line-height:${ls};text-shadow:${sh}">
        ${numHtml}For God so loved the world that He gave His only begotten Son.</div>
      <div style="font-size:clamp(7px,.8vw,11px);color:${t.transColor||'#555'};letter-spacing:.25em;
        text-transform:uppercase;margin-top:10px;opacity:.7;font-family:'${ff}',serif">NKJV</div>
    </div>`;
  }
  container.innerHTML = inner;
}
function showThemeContextMenu(e) {
  const menu = document.getElementById('themeContextMenu');
  if (!menu) return;
  menu.style.display = 'block';
  const x = Math.min(e.clientX, window.innerWidth  - 230);
  const y = Math.min(e.clientY, window.innerHeight - 240);
  menu.style.left = x + 'px';
  menu.style.top  = y + 'px';
  // Update dynamic label
  const cat = _themeCategory.charAt(0).toUpperCase() + _themeCategory.slice(1);
  const ctxLabel = document.getElementById('themeCtxLabel');
  if (ctxLabel) ctxLabel.textContent = cat;
  // Show/hide Edit based on whether a theme is selected
  const editBtn = document.getElementById('themeCtxEdit');
  if (editBtn) editBtn.style.display = _themeCtxId ? '' : 'none';
  setTimeout(() => document.addEventListener('click', _hideThemeCtx, { once: true }), 10);
}

function _hideThemeCtx() {
  const menu = document.getElementById('themeContextMenu');
  if (menu) menu.style.display = 'none';
  _themeCtxId = null;
}

function openThemeEditorFor(id) {
  _hideThemeCtx();
  if (window.electronAPI) {
    window.electronAPI.openThemeDesigner({
      themeId:  id || null,
      category: _themeCategory
    });
  }
}

function selectThemeInGrid(id) {
  // Highlight + preview AND apply the theme live immediately
  _themeCtxId = id;
  State.currentThemeCtxId = id;
  _applyThemeLive(id);
  // For song and presentation themes, save immediately on click (single-click = apply + persist)
  // Scripture themes require explicit "Set As Default" to avoid accidental live changes
  if (_themeCategory === 'song' || _themeCategory === 'presentation') {
    applyThemeFromGrid(id);
    return; // applyThemeFromGrid already calls renderThemeGrid + preview
  }
  renderThemeGrid();
  const previewDisplay = document.getElementById('previewDisplay');
  if (previewDisplay) {
    previewDisplay.style.display = 'flex';
    _renderThemePreview(previewDisplay, id);
  }
}

// Apply theme live without saving as default
function _applyThemeLive(id) {
  const t = _getAllThemes().find(x => x.id === id);
  if (!t) return;
  if (_themeCategory === 'scripture') {
    State.currentTheme = id;
    // Update theme swatches UI
    document.querySelectorAll('.theme-swatch').forEach(sw => {
      sw.classList.toggle('active', sw.dataset.t === id);
    });
    // Update canvas preview
    const pc = document.getElementById('previewCanvas');
    const lc = document.getElementById('liveCanvas');
    if (pc) pc.dataset.theme = id;
    if (lc) lc.dataset.theme = id;
    if (State.isLive && State.liveVerse) syncProjection();
  } else if (_themeCategory === 'song') {
    State.currentSongTheme = id;
    // If a song is currently live, re-send with new theme
    if (State.isLive && State.currentSongId !== null && State.currentSongSlideIdx !== null) {
      const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
      if (song) sendSongToProjection(song, State.currentSongSlideIdx);
    }
  } else if (_themeCategory === 'presentation') {
    State.currentPresTheme = id;
  }
}

function applyThemeFromGrid(id) {
  // Set as default (saves to settings disk) AND apply live
  const t = _getAllThemes().find(x => x.id === id);
  _applyThemeLive(id);

  if (_themeCategory === 'scripture') {
    if (window.electronAPI) {
      window.electronAPI.getSettings().then(s =>
        window.electronAPI.saveSettings({...s, theme: id}, { themeOnly: true, changedKeys: ['theme'] })
      );
    }
  } else if (_themeCategory === 'song') {
    if (window.electronAPI) {
      window.electronAPI.getSettings().then(s =>
        window.electronAPI.saveSettings({...s, songTheme: id}, { themeOnly: true, changedKeys: ['songTheme'] })
      );
    }
  } else if (_themeCategory === 'presentation') {
    if (window.electronAPI) {
      window.electronAPI.getSettings().then(s =>
        window.electronAPI.saveSettings({...s, presTheme: id}, { themeOnly: true, changedKeys: ['presTheme'] })
      );
    }
  }

  State.currentThemeCtxId = id;
  renderThemeGrid();
  const previewDisplay = document.getElementById('previewDisplay');
  if (previewDisplay) {
    previewDisplay.style.display = 'flex';
    _renderThemePreview(previewDisplay, id);
  }
  toast(`⭐ "${t?.name || id}" set as default`);
}

function setThemeAsDefault() {
  const id = _themeCtxId || State.currentThemeCtxId;
  if (id) applyThemeFromGrid(id);
  _hideThemeCtx();
}

function setThemeAsAltDefault() {
  const id = _themeCtxId || State.currentThemeCtxId;
  _hideThemeCtx();
  if (!id) return;
  const t = _getAllThemes().find(x => x.id === id);
  if (!t) return;
  if (window.electronAPI) {
    window.electronAPI.getSettings().then(s => {
      const update = {};
      if (_themeCategory === 'song')         update.songThemeAlt = id;
      else if (_themeCategory === 'presentation') update.presThemeAlt = id;
      else                                   update.themeAlt = id;
      window.electronAPI.saveSettings({ ...s, ...update });
    });
  }
  toast(`☆ "${t.name}" set as alternate default`);
}

async function deleteThemeFromGrid() {
  const id = _themeCtxId || State.currentThemeCtxId;
  const t = _getAllThemes().find(x => x.id === id);
  _hideThemeCtx();
  if (!t) return;
  if (t.builtIn) { toast('Cannot delete built-in themes'); return; }
  // BUG-13 FIX: native confirm() is blocked in Electron contextIsolation builds.
  _confirmModal(`Delete "${t.name}"? This cannot be undone.`, async () => {
    State.themes = (State.themes || []).filter(x => x.id !== id);
    if (window.electronAPI) await window.electronAPI.saveThemes(State.themes);
    renderThemeGrid();
    toast(`🗑 Deleted: "${t.name}"`);
  });
}

// ─── START ────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', init);


async function showMediaIntegrityPanel() {
  const overlay = document.getElementById('mediaIntegrityOverlay');
  if (!overlay) return;
  overlay.style.display = 'flex';
  await refreshMediaIntegrityPanel();
}
function closeMediaIntegrityPanel() {
  const overlay = document.getElementById('mediaIntegrityOverlay');
  if (overlay) overlay.style.display = 'none';
}
async function refreshMediaIntegrityPanel() {
  const body = document.getElementById('mediaIntegrityBody');
  const summary = document.getElementById('mediaIntegritySummary');
  if (!body || !summary) return;
  body.innerHTML = '<div class="empty-state"><span class="empty-icon">🧪</span>Checking media library…</div>';
  if (!window.electronAPI?.getMediaIntegrity) {
    summary.textContent = 'Media integrity tools are unavailable in this build.';
    return;
  }
  const res = await window.electronAPI.getMediaIntegrity();
  if (!res?.success) {
    summary.textContent = 'Could not inspect media library.';
    body.innerHTML = `<div class="empty-state"><span class="empty-icon">⚠</span>${escapeHtml(res?.error || 'Unknown error')}</div>`;
    return;
  }
  const items = res.items || [];
  summary.textContent = `${res.summary?.healthy || 0} healthy, ${res.summary?.broken || 0} broken, ${res.summary?.total || items.length} total`;
  if (!items.length) {
    body.innerHTML = '<div class="empty-state"><span class="empty-icon">🗂</span>No media items found.</div>';
    return;
  }
  body.innerHTML = `
    <table style="width:100%;border-collapse:collapse;font-size:12px">
      <thead>
        <tr>
          <th style="text-align:left;padding:8px;border-bottom:1px solid var(--border)">Title</th>
          <th style="text-align:left;padding:8px;border-bottom:1px solid var(--border)">Type</th>
          <th style="text-align:left;padding:8px;border-bottom:1px solid var(--border)">Status</th>
          <th style="text-align:left;padding:8px;border-bottom:1px solid var(--border)">Storage Location</th>
        </tr>
      </thead>
      <tbody>
        ${items.map(item => `
          <tr>
            <td style="padding:8px;border-bottom:1px solid var(--border)">${escapeHtml(item.title || '')}</td>
            <td style="padding:8px;border-bottom:1px solid var(--border)">${escapeHtml(item.type || '')}</td>
            <td style="padding:8px;border-bottom:1px solid var(--border);color:${item.broken ? '#ff9d9d' : 'var(--text)'}">${item.broken ? 'Broken / Missing' : 'OK'}</td>
            <td style="padding:8px;border-bottom:1px solid var(--border);font-family:monospace;font-size:11px;word-break:break-all">${escapeHtml(item.storageLocation || '')}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>`;
}
async function repairMediaIntegrity() {
  if (!window.electronAPI?.repairMediaLinks) return;
  const res = await window.electronAPI.repairMediaLinks();
  if (res?.success) {
    toast(`✓ Media repair complete: ${res.repaired || 0} repaired, ${res.unresolved || 0} unresolved`);
    await loadMedia();
    await refreshMediaIntegrityPanel();
  } else {
    toast('⚠ Media repair failed: ' + (res?.error || 'unknown error'));
  }
}
async function clearMediaCacheAction() {
  if (!window.electronAPI?.clearMediaCache) return;
  const res = await window.electronAPI.clearMediaCache();
  if (res?.success) {
    toast(`✓ Cache cleared: ${res.removedFiles || 0} files removed`);
  } else {
    toast('⚠ Could not clear cache: ' + (res?.error || 'unknown error'));
  }
}
window.showMediaIntegrityPanel = showMediaIntegrityPanel;
window.closeMediaIntegrityPanel = closeMediaIntegrityPanel;
window.refreshMediaIntegrityPanel = refreshMediaIntegrityPanel;
window.repairMediaIntegrity = repairMediaIntegrity;
window.clearMediaCacheAction = clearMediaCacheAction;

// ─── ADAPTIVE TRANSCRIPT MEMORY UI ───────────────────────────────────────────

/** Populate the speaker selector dropdown from TranscriptMemory profiles. */
function refreshSpeakerSelect() {
  const sel = document.getElementById('tmSpeakerSelect');
  if (!sel || !window.TranscriptMemory) return;
  const profiles = TranscriptMemory.getProfiles();
  const current  = sel.value;
  sel.innerHTML  = '<option value="">👤 Speaker…</option>';
  profiles.forEach(p => {
    const opt = document.createElement('option');
    opt.value       = p.id;
    opt.textContent = p.name;
    if (p.isDefault && !current) opt.selected = true;
    if (p.id === current)        opt.selected = true;
    sel.appendChild(opt);
  });
}

/** Toggle adaptive corrections on/off. */
async function toggleAdaptiveCorrections() {
  if (!window.TranscriptMemory) return;
  TranscriptMemory.enabled = !TranscriptMemory.enabled;
  const btn = document.getElementById('tmToggleBtn');
  if (btn) {
    btn.textContent = TranscriptMemory.enabled ? '🧠 ON' : '🧠 OFF';
    btn.style.color = TranscriptMemory.enabled ? 'var(--gold)' : 'var(--text-dim)';
    btn.style.borderColor = TranscriptMemory.enabled ? 'var(--gold-d)' : 'var(--border)';
  }
  try {
    const existing = await window.electronAPI?.getSettings?.();
    if (existing && window.electronAPI?.saveSettings) {
      await window.electronAPI.saveSettings({ ...existing, adaptiveEnabled: !!TranscriptMemory.enabled });
    }
  } catch(_) {}
  toast(TranscriptMemory.enabled ? '🧠 Adaptive corrections ON' : '🧠 Adaptive corrections OFF');
}

/** Open the adaptive settings panel in the settings window. */
function openAdaptiveSettings() {
  if (window.electronAPI?.openSettings) {
    window.electronAPI.openSettings({ section: 'audio', tab: 'learning' });
  }
}


// ─── TRANSCRIPT REVIEW PANEL ────────────────────────────────────────────────
function getTranscriptReviewLine() {
  return (State.transcriptLines || []).find(l => l.id === State.reviewTranscriptLineId) || null;
}

function openTranscriptReviewPanel() {
  const modal = document.getElementById('transcriptReviewModal');
  if (modal) modal.style.display = 'flex';
  renderTranscriptReviewList();
  if (!getTranscriptReviewLine() && State.transcriptLines?.length) {
    selectTranscriptReviewLine(State.transcriptLines[State.transcriptLines.length - 1].id);
  }
}

function closeTranscriptReviewPanel() {
  const modal = document.getElementById('transcriptReviewModal');
  if (modal) modal.style.display = 'none';
}

function renderTranscriptReviewList() {
  const list = document.getElementById('transcriptReviewList');
  const count = document.getElementById('transcriptReviewCount');
  const query = (document.getElementById('transcriptReviewSearch')?.value || '').trim().toLowerCase();
  if (!list) return;
  const items = (State.transcriptLines || []).filter(l => {
    if (!query) return true;
    return String(l.text || '').toLowerCase().includes(query) || String(l.raw || '').toLowerCase().includes(query);
  }).slice().reverse();
  list.innerHTML = '';
  if (count) count.textContent = `${items.length} lines`;
  if (!items.length) {
    list.innerHTML = `<div class="empty-state" style="min-height:120px"><span class="empty-icon">📝</span>No transcript lines to review.</div>`;
    return;
  }
  items.forEach(line => {
    const card = document.createElement('button');
    card.type = 'button';
    card.style.cssText = 'text-align:left;background:var(--card);border:1px solid var(--border);border-radius:8px;padding:10px;cursor:pointer;color:var(--text)';
    if (line.id === State.reviewTranscriptLineId) card.style.borderColor = 'var(--gold)';
    const changed = line.corrected && line.raw && line.raw !== line.text;
    card.innerHTML = `
      <div style="display:flex;gap:8px;align-items:center;margin-bottom:4px">
        <span style="font-size:10px;color:${changed ? 'var(--gold)' : 'var(--text-dim)'}">${changed ? 'Corrected' : 'Transcript'}</span>
        <span style="margin-left:auto;font-size:10px;color:var(--text-dim)">${new Date(line.time || Date.now()).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'})}</span>
      </div>
      <div style="font-size:12px;line-height:1.45;color:var(--text)">${escapeHtml(String(line.text || '').slice(0,160))}</div>
      ${changed ? `<div style="font-size:10px;color:var(--text-dim);margin-top:5px">Original: ${escapeHtml(String(line.raw || '').slice(0,120))}</div>` : ''}
    `;
    card.addEventListener('click', () => selectTranscriptReviewLine(line.id));
    list.appendChild(card);
  });
}

function selectTranscriptReviewLine(lineId) {
  State.reviewTranscriptLineId = lineId;
  const line = getTranscriptReviewLine();
  const badge = document.getElementById('transcriptReviewBadge');
  if (!line) return renderTranscriptReviewList();
  document.getElementById('transcriptReviewOriginal').value = line.raw || line.text || '';
  document.getElementById('transcriptReviewEdited').value = line.text || '';
  document.getElementById('transcriptRuleSource').value = line.raw || line.text || '';
  document.getElementById('transcriptRuleTarget').value = line.text || '';
  if (badge) badge.textContent = line.corrected ? 'Corrected line' : 'Raw transcript line';
  renderTranscriptReviewList();
}

async function approveTranscriptReview() {
  const line = getTranscriptReviewLine();
  if (!line) return;
  if (window.TranscriptMemory && line.appliedRules?.length) {
    line.appliedRules.forEach(ruleId => {
      try { TranscriptMemory.recordCorrectionFeedback(line.chunkId, ruleId, true); } catch(_) {}
    });
  }
  toast('✓ Transcript correction approved');
}

async function rejectTranscriptReview() {
  const line = getTranscriptReviewLine();
  if (!line) return;
  const original = line.raw || '';
  if (!original) return;
  if (window.TranscriptMemory && line.appliedRules?.length) {
    line.appliedRules.forEach(ruleId => {
      try { TranscriptMemory.recordCorrectionFeedback(line.chunkId, ruleId, false); } catch(_) {}
    });
  }
  line.text = original;
  line.corrected = false;
  const bodyLine = document.querySelector(`.transcript-line[data-line-id="${line.id}"]`);
  if (bodyLine) {
    bodyLine.textContent = original;
    bodyLine.style.borderLeft = '';
    bodyLine.style.paddingLeft = '';
    bodyLine.title = '';
  }
  selectTranscriptReviewLine(line.id);
  toast('↺ Correction rejected for this line');
}

async function applyTranscriptEditedText() {
  const line = getTranscriptReviewLine();
  if (!line) return;
  const edited = (document.getElementById('transcriptReviewEdited')?.value || '').trim();
  if (!edited) return;
  const before = line.text || line.raw || '';
  line.text = edited;
  line.corrected = edited !== (line.raw || '');
  const bodyLine = document.querySelector(`.transcript-line[data-line-id="${line.id}"]`);
  if (bodyLine) {
    bodyLine.textContent = edited;
    if (line.corrected) {
      bodyLine.style.borderLeft = '2px solid rgba(201,168,76,.4)';
      bodyLine.style.paddingLeft = '6px';
      bodyLine.title = `Original: "${line.raw || ''}"`;
    }
  }
  if (window.TranscriptMemory && line.chunkId && line.raw && edited !== line.raw) {
    try { TranscriptMemory.recordUserCorrection(line.chunkId, line.raw, edited); } catch(_) {}
  } else if (window.TranscriptMemory && line.chunkId && before !== edited) {
    try { TranscriptMemory.recordUserCorrection(line.chunkId, before, edited); } catch(_) {}
  }
  selectTranscriptReviewLine(line.id);
  toast('✓ Edited transcript line applied');
}

async function saveTranscriptReviewAsRule() {
  const source = (document.getElementById('transcriptRuleSource')?.value || '').trim();
  const target = (document.getElementById('transcriptRuleTarget')?.value || '').trim();
  const scope = document.getElementById('transcriptRuleScope')?.value || 'global';
  const ruleType = document.getElementById('transcriptRuleType')?.value || 'phrase';
  if (!source || !target) return toast('⚠ Enter both wrong text and correct text');
  if (!window.TranscriptMemory) return;
  try {
    const speakerSel = document.getElementById('tmSpeakerSelect');
    TranscriptMemory.addRule({
      sourceText: source,
      targetText: target,
      scope,
      ruleType,
      speakerProfileId: scope === 'speaker' ? (speakerSel?.value || TranscriptMemory.activeProfile || null) : null,
    });
    const line = getTranscriptReviewLine();
    if (line?.chunkId) {
      try { TranscriptMemory.recordUserCorrection(line.chunkId, source, target); } catch(_) {}
    }
    toast('✓ Correction rule saved');
  } catch (e) {
    toast('⚠ Could not save correction rule');
  }
}

// Wire up speaker select + refresh on transcript memory init
document.addEventListener('DOMContentLoaded', () => {
  // Delay so TranscriptMemory.init() completes first
  setTimeout(async () => {
    refreshSpeakerSelect();
    const sel = document.getElementById('tmSpeakerSelect');
    try {
      const settings = await window.electronAPI?.getSettings?.();
      if (sel && settings?.adaptiveSpeakerProfile) {
        sel.value = settings.adaptiveSpeakerProfile;
        if (window.TranscriptMemory) TranscriptMemory.setProfile(sel.value);
      }
    } catch(_) {}
    if (sel) {
      sel.addEventListener('change', async () => {
        if (window.TranscriptMemory) TranscriptMemory.setProfile(sel.value);
        try {
          const existing = await window.electronAPI?.getSettings?.();
          if (existing && window.electronAPI?.saveSettings) {
            await window.electronAPI.saveSettings({ ...existing, adaptiveSpeakerProfile: sel.value || 'default' });
          }
        } catch(_) {}
      });
    }
  }, 1000);
});

document.addEventListener('click', (e) => {
  const modal = document.getElementById('transcriptReviewModal');
  if (modal && e.target === modal) closeTranscriptReviewPanel();
});

document.addEventListener('click', (e) => {
  const modal = document.getElementById('detectionReviewModal');
  if (modal && e.target === modal) closeDetectionReviewPanel();
});


// ─── COMMAND CENTER LOGIC ───────────────────────────────────────────────
function _ccFormatQueueItem(item){
  if (!item) return 'None';
  if (item.type === 'song') return item.songTitle || item.title || 'Song';
  if (item.type === 'song-slide') return item.songTitle || item.title || `Song Slide ${Number(item.sectionIdx || 0) + 1}`;
  if (item.type === 'media') return item.title || item.name || 'Media';
  if (item.type === 'presentation') return item.title || item.name || 'Presentation';
  return item.ref || item.title || 'Item';
}

function _ccGetLiveLabel(){
  try{
    if (State.liveContentType === 'scripture' && State.liveVerse?.ref) return State.liveVerse.ref;
    if (State.liveContentType === 'song') {
      const song = (State.songs || []).find(s => String(s.id) === String(State.currentSongId));
      const sec = song?.sections?.[State.currentSongSlideIdx ?? 0];
      if (song?.title) return sec?.label ? `${song.title} — ${sec.label}` : song.title;
    }
    if (State.liveContentType === 'media' && State.currentMediaId != null) {
      const media = (State.media || []).find(m => String(m.id) === String(State.currentMediaId));
      if (media) return media.title || media.name || 'Media';
    }
    if (State.liveContentType === 'presentation' && State.currentPresId != null) {
      const pres = (State.presentations || []).find(p => String(p.id) === String(State.currentPresId));
      if (pres) return pres.name || pres.title || 'Presentation';
    }
    if (State.previewVerse?.ref && State.isLive) return State.previewVerse.ref;
  }catch(_){}
  return 'None';
}

function _ccGetNextLabel(){
  try{
    const realItems = (State.queue || []).filter(q => q.type !== 'section');
    if (!realItems.length) return 'None';
    if (State.currentQueueIdx == null || State.currentQueueIdx < 0) return _ccFormatQueueItem(realItems[0]);
    const currentRealPos = realItems.findIndex(item => (State.queue || []).indexOf(item) === State.currentQueueIdx);
    if (currentRealPos >= 0 && currentRealPos + 1 < realItems.length) return _ccFormatQueueItem(realItems[currentRealPos + 1]);
    return 'End of schedule';
  } catch(_) { return 'None'; }
}

async function updateCommandCenter(){
  try{
    const liveEl = document.getElementById('ccLive');
    const nextEl = document.getElementById('ccNext');
    const transcriptEl = document.getElementById('ccTranscript');
    const detectionEl = document.getElementById('ccDetection');
    const remoteEl = document.getElementById('ccRemote');

    if (liveEl) liveEl.textContent = _ccGetLiveLabel();
    if (nextEl) nextEl.textContent = _ccGetNextLabel();
    if (transcriptEl) transcriptEl.textContent = State.isRecording ? 'Recording' : 'Idle';
    if (detectionEl) detectionEl.textContent = State.detections?.length ? 'Active' : 'Idle';

    if (remoteEl && window.electronAPI?.getRemoteRuntimeStatus) {
      try {
        const rs = await window.electronAPI.getRemoteRuntimeStatus();
        if (rs?.connected) {
          const role = rs.lastRole ? ` (${rs.lastRole})` : '';
          remoteEl.textContent = `Connected${role}`;
        } else {
          remoteEl.textContent = (rs && rs.serverEnabled === false) ? 'Disabled' : 'Disconnected';
        }
      } catch (_) {
        remoteEl.textContent = 'Disconnected';
      }
    } else if (remoteEl && window.electronAPI?.getRemoteStatus) {
      try {
        const rs = await window.electronAPI.getRemoteStatus();
        if (rs?.enabled) {
          remoteEl.textContent = rs.connected ? 'Connected' : 'Listening';
        } else {
          remoteEl.textContent = 'Disabled';
        }
      } catch (_) {
        remoteEl.textContent = 'Disconnected';
      }
    } else if (remoteEl) {
      remoteEl.textContent = 'N/A';
    }
  }catch(_){}
}

async function ccGoLive(){
  try{
    if (!State.isLive) {
      State.isLive = true;
      const btn = document.getElementById('goLiveBtn');
      const txt = document.getElementById('goLiveTxt');
      const tag = document.getElementById('liveTag');
      btn?.classList?.toggle('active', true);
      if (txt) txt.textContent = 'On Air';
      if (tag) tag.style.display = 'flex';

      if (State.currentTab === 'song' && State.currentSongId != null && State.currentSongSlideIdx != null) presentSongSlide(State.currentSongSlideIdx);
      else if (State.currentTab === 'media' && State.currentMediaId != null) presentCurrentMedia();
      else if (State.currentTab === 'pres' && State.currentPresSlides?.length) presentCurrentPresSlide();
      else if (State.previewVerse) sendPreviewToLive();
      else if (State.liveVerse) syncProjection();

      toast('🔴 Projection live');
    } else {
      if (State.liveContentType === 'song' && State.currentSongSlideIdx != null) presentSongSlide(State.currentSongSlideIdx);
      else if (State.liveContentType === 'media' && State.currentMediaId != null) presentCurrentMedia();
      else if (State.liveContentType === 'scripture' && State.liveVerse) syncProjection();
      else if (State.currentTab === 'pres' && State.currentPresSlides?.length) presentCurrentPresSlide();
      toast('🔴 Projection refreshed');
    }
    updateCommandCenter();
  } catch(_) {}
}

function ccStopLive(){
  try{
    State.isLive = false;
    const btn = document.getElementById('goLiveBtn');
    const txt = document.getElementById('goLiveTxt');
    const tag = document.getElementById('liveTag');
    btn?.classList?.toggle('active', false);
    if (txt) txt.textContent = 'Go Live';
    if (tag) tag.style.display = 'none';
    clearLive();
    toast('⬛ Projection stopped');
    updateCommandCenter();
  }catch(_){}
}

function ccClear(){
  try{
    clearLive();
    toast('✕ Display cleared');
    updateCommandCenter();
  }catch(_){}
}

setInterval(() => { updateCommandCenter(); }, 1500);

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('ccToggle')?.addEventListener('click',()=>{
    const body=document.getElementById('ccBody');
    const btn=document.getElementById('ccToggle');
    const collapsed = body.style.display==='none';
    body.style.display = collapsed ? 'block':'none';
    btn.textContent = collapsed ? '—' : '□';
  });

  document.getElementById('ccClose')?.addEventListener('click',()=>{
    const cc=document.getElementById('commandCenter');
    if(cc) cc.style.display='none';
  });
});

// Restore OCC from View menu
if(window.electronAPI?.on) {
  window.electronAPI.on('menu-show-occ', () => {
    const cc=document.getElementById('commandCenter');
    if(cc){ cc.style.display=''; const body=document.getElementById('ccBody'); if(body) body.style.display='block'; }
  });
}

// ─── POST SERVICE REPORT ─────────────────────────────────────────────
function generatePostServiceReport(){
  const lines = State.transcriptLines || [];
  const detections = State.detections || [];

  const report = {
    totalLines: lines.length,
    correctedLines: lines.filter(l => l.corrected).length,
    detections: detections.length,
    approved: detections.filter(d => d.status === 'approved').length,
    rejected: detections.filter(d => d.status === 'rejected').length,
    verses: detections.map(d => d.ref)
  };

  let html = `
    <div><strong>Total transcript lines:</strong> ${report.totalLines}</div>
    <div><strong>Corrected lines:</strong> ${report.correctedLines}</div>
    <div><strong>Detections:</strong> ${report.detections}</div>
    <div><strong>Approved:</strong> ${report.approved}</div>
    <div><strong>Rejected:</strong> ${report.rejected}</div>
    <hr>
    <div><strong>Detected verses:</strong><br>${report.verses.join('<br>') || 'None'}</div>
  `;

  document.getElementById('psrContent').innerHTML = html;
}

function openPostServiceReport(){
  generatePostServiceReport();
  document.getElementById('postServiceReportModal').style.display = 'flex';
}

function closePostServiceReport(){
  document.getElementById('postServiceReportModal').style.display = 'none';
}

// Auto trigger when transcript stops
const originalStop = window.stopRecording;
window.stopRecording = function(){
  if(originalStop) originalStop();
  setTimeout(openPostServiceReport, 800);
};


// ─── SERVICE REPLAY TIMELINE ──────────────────────────────────────────
function _pushReplayEvent(type, payload = {}) {
  try {
    State.replayTimeline = Array.isArray(State.replayTimeline) ? State.replayTimeline : [];
    State.replayTimeline.push({
      id: `replay_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      at: Date.now(),
      type,
      payload
    });
    if (State.replayTimeline.length > 5000) State.replayTimeline = State.replayTimeline.slice(-5000);
  } catch (_) {}
}

function buildReplaySummary(event) {
  if (!event) return 'No event';
  const t = new Date(event.at || Date.now()).toLocaleTimeString([], {hour:'2-digit', minute:'2-digit', second:'2-digit'});
  return `${t} · ${event.type}`;
}

function renderReplayTimelineList() {
  const list = document.getElementById('replayTimelineList');
  if (!list) return;
  const items = State.replayTimeline || [];
  list.innerHTML = '';
  if (!items.length) {
    list.innerHTML = '<div style="padding:14px;border:1px dashed #2f3650;border-radius:10px;color:#9aa4c7;text-align:center">No replay events yet.</div>';
    return;
  }
  items.forEach((ev, idx) => {
    const row = document.createElement('button');
    row.type = 'button';
    row.style.cssText = `text-align:left;background:${idx===State.replayIndex ? 'rgba(212,175,55,.12)' : '#11131a'};border:1px solid ${idx===State.replayIndex ? '#d4af37' : '#232838'};color:#edf1ff;border-radius:8px;padding:8px;cursor:pointer`;
    row.innerHTML = `<div style="font-size:11px;color:#9aa4c7">${buildReplaySummary(ev)}</div><div style="font-size:12px">${(ev.payload?.ref || ev.payload?.title || ev.payload?.text || ev.payload?.summary || '').toString().slice(0,90)}</div>`;
    row.addEventListener('click', () => setReplayTimelineIndex(idx));
    list.appendChild(row);
  });
}

function renderReplayEvent() {
  const items = State.replayTimeline || [];
  const idx = Math.max(0, Math.min(State.replayIndex || 0, Math.max(0, items.length - 1)));
  State.replayIndex = idx;
  const ev = items[idx];
  const idxLabel = document.getElementById('replayIndexLabel');
  const range = document.getElementById('replayRange');
  if (idxLabel) idxLabel.textContent = items.length ? `${idx + 1} / ${items.length}` : '0 / 0';
  if (range) {
    range.max = String(Math.max(0, items.length - 1));
    range.value = String(idx);
  }
  document.getElementById('replayEventMeta').textContent = ev ? buildReplaySummary(ev) : 'No event';
  document.getElementById('replayTranscriptText').textContent = ev ? (ev.payload?.transcript || ev.payload?.text || '') : '';
  document.getElementById('replayLiveText').textContent = ev ? (ev.payload?.liveText || ev.payload?.ref || ev.payload?.summary || '') : '';
  renderReplayTimelineList();
}

function openServiceReplayTimeline() {
  const modal = document.getElementById('serviceReplayModal');
  if (modal) modal.style.display = 'flex';
  if ((State.replayTimeline || []).length) State.replayIndex = Math.max(0, (State.replayTimeline || []).length - 1);
  renderReplayEvent();
}

function closeServiceReplayTimeline() {
  const modal = document.getElementById('serviceReplayModal');
  if (modal) modal.style.display = 'none';
  if (State.replayAutoPlayTimer) {
    clearInterval(State.replayAutoPlayTimer);
    State.replayAutoPlayTimer = null;
  }
  const btn = document.getElementById('replayPlayBtn');
  if (btn) btn.textContent = '▶ Play';
}

function setReplayTimelineIndex(idx) {
  State.replayIndex = idx;
  renderReplayEvent();
}

function stepReplayTimeline(delta) {
  const len = (State.replayTimeline || []).length;
  if (!len) return;
  State.replayIndex = Math.max(0, Math.min(len - 1, (State.replayIndex || 0) + delta));
  renderReplayEvent();
}

function toggleReplayAutoPlay() {
  const btn = document.getElementById('replayPlayBtn');
  if (State.replayAutoPlayTimer) {
    clearInterval(State.replayAutoPlayTimer);
    State.replayAutoPlayTimer = null;
    if (btn) btn.textContent = '▶ Play';
    return;
  }
  if (btn) btn.textContent = '⏸ Pause';
  State.replayAutoPlayTimer = setInterval(() => {
    const len = (State.replayTimeline || []).length;
    if (!len) return;
    if ((State.replayIndex || 0) >= len - 1) {
      clearInterval(State.replayAutoPlayTimer);
      State.replayAutoPlayTimer = null;
      if (btn) btn.textContent = '▶ Play';
      return;
    }
    State.replayIndex += 1;
    renderReplayEvent();
  }, 1200);
}

document.addEventListener('click', (e) => {
  const modal = document.getElementById('serviceReplayModal');
  if (modal && e.target === modal) closeServiceReplayTimeline();
});


// ─── SERVICE ARCHIVE + SEARCH CENTER ───────────────────────────────────
function buildCurrentPostServiceReportData() {
  const lines = State.transcriptLines || [];
  const detections = State.detections || [];
  return {
    totalLines: lines.length,
    correctedLines: lines.filter(l => l.corrected).length,
    detections: detections.length,
    approved: detections.filter(d => d.status === 'approved').length,
    rejected: detections.filter(d => d.status === 'rejected').length,
    verses: detections.map(d => d.ref).filter(Boolean),
  };
}

async function saveCurrentServiceToArchive() {
  // BUG-14 FIX: window.prompt() is fully blocked in Electron (contextIsolation=true)
  // — it always returns null, so the operator could never set a custom service title.
  // Replace with a lightweight text-input modal.
  const defaultTitle = `Service ${new Date().toLocaleDateString()}`;
  _inputModal('Save Service to Archive', 'Service title', defaultTitle, async (title) => {
    try {
      const speaker = (document.getElementById('tmSpeakerSelect')?.selectedOptions?.[0]?.textContent || State.settings?.adaptiveSpeakerProfile || '').trim();
      const mediaUsed = (State.media || []).filter(m => m.presented || m.used || m.wasLive).map(m => ({ id: m.id, title: m.title || m.name || '', type: m.type || '' }));
      const payload = {
        title: title || defaultTitle,
        serviceDate: new Date().toISOString(),
        speaker,
        transcriptLines: State.transcriptLines || [],
        detections: State.detections || [],
        replayTimeline: State.replayTimeline || [],
        report: buildCurrentPostServiceReportData(),
        schedule: State.schedule || null,
        mediaUsed,
        notes: ''
      };
      const res = await window.electronAPI?.saveServiceArchive?.(payload);
      if (res?.success) {
        toast('🗂 Service archived');
        await loadServiceArchiveIndex();
      } else {
        toast('⚠ Could not save service archive');
      }
    } catch (_) {
      toast('⚠ Could not save service archive');
    }
  });
}

// Reusable text-input modal (replaces blocked window.prompt() in Electron)
function _inputModal(heading, label, defaultValue, onConfirm) {
  document.getElementById('_acInputModal')?.remove();
  const overlay = document.createElement('div');
  overlay.id = '_acInputModal';
  overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:1000;display:flex;align-items:center;justify-content:center';
  const box = document.createElement('div');
  box.style.cssText = 'background:var(--panel);border:1px solid var(--border-lit);border-radius:8px;padding:22px;width:380px;box-shadow:0 8px 32px rgba(0,0,0,.7)';
  const h = document.createElement('div');
  h.style.cssText = 'font-size:13px;font-weight:700;color:var(--text);margin-bottom:14px';
  h.textContent = heading;
  const lbl = document.createElement('label');
  lbl.style.cssText = 'display:block;font-size:11px;color:var(--text-dim);margin-bottom:6px';
  lbl.textContent = label;
  const inp = document.createElement('input');
  inp.type = 'text';
  inp.value = defaultValue;
  inp.style.cssText = 'width:100%;box-sizing:border-box;background:var(--card);border:1px solid var(--border-lit);border-radius:5px;color:var(--text);font-size:12px;padding:8px 10px;font-family:inherit;outline:none';
  inp.addEventListener('focus', () => inp.select());
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') { overlay.remove(); onConfirm(inp.value.trim()); } if (e.key === 'Escape') overlay.remove(); });
  const btns = document.createElement('div');
  btns.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;margin-top:16px';
  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'Cancel';
  cancelBtn.style.cssText = 'padding:7px 16px;background:var(--card);border:1px solid var(--border-lit);border-radius:5px;color:var(--text);font-size:12px;cursor:pointer;font-family:inherit';
  cancelBtn.onclick = () => overlay.remove();
  const okBtn = document.createElement('button');
  okBtn.textContent = 'Save';
  okBtn.style.cssText = 'padding:7px 16px;background:var(--gold);border:none;border-radius:5px;color:#000;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit';
  okBtn.onclick = () => { overlay.remove(); onConfirm(inp.value.trim()); };
  btns.appendChild(cancelBtn);
  btns.appendChild(okBtn);
  box.appendChild(h);
  box.appendChild(lbl);
  box.appendChild(inp);
  box.appendChild(btns);
  overlay.appendChild(box);
  document.body.appendChild(overlay);
  overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
  setTimeout(() => inp.focus(), 50);
}

async function loadServiceArchiveIndex(query = '') {
  const res = query
    ? await window.electronAPI?.searchServiceArchive?.(query)
    : await window.electronAPI?.getServiceArchiveIndex?.();
  State.serviceArchiveItems = Array.isArray(res?.items) ? res.items : [];
  renderServiceArchiveList();
}

function renderServiceArchiveList() {
  const list = document.getElementById('serviceArchiveList');
  if (!list) return;
  list.innerHTML = '';
  const items = State.serviceArchiveItems || [];
  if (!items.length) {
    list.innerHTML = '<div style="padding:16px;border:1px dashed #2f3650;border-radius:10px;color:#9aa4c7;text-align:center">No archived services yet.</div>';
    return;
  }
  items.forEach(item => {
    const row = document.createElement('button');
    row.type = 'button';
    row.style.cssText = `text-align:left;background:${State.serviceArchiveSelected?.id===item.id ? 'rgba(212,175,55,.12)' : '#11131a'};border:1px solid ${State.serviceArchiveSelected?.id===item.id ? '#d4af37' : '#232838'};color:#edf1ff;border-radius:8px;padding:10px;cursor:pointer`;
    row.innerHTML = `<div style="font-size:12px;font-weight:700">${escapeHtml(item.title || 'Untitled Service')}</div><div style="font-size:11px;color:#9aa4c7">${new Date(item.serviceDate || item.savedAt || Date.now()).toLocaleString()} · ${escapeHtml(item.speaker || 'Unknown speaker')}</div><div style="font-size:11px;color:#9aa4c7">Transcript ${Number(item.transcriptCount||0)} · Verses ${Number(item.detectionCount||0)} · Media ${Number(item.mediaCount||0)}</div>`;
    row.addEventListener('click', () => loadServiceArchivePreview(item.id));
    list.appendChild(row);
  });
}

async function loadServiceArchivePreview(archiveId) {
  const res = await window.electronAPI?.loadServiceArchiveItem?.(archiveId);
  if (!res?.success || !res.item) return toast('⚠ Could not load archived service');
  State.serviceArchiveSelected = res.item;
  renderServiceArchiveList();
  const meta = document.getElementById('serviceArchivePreviewMeta');
  const transcript = document.getElementById('serviceArchivePreviewTranscript');
  const detections = document.getElementById('serviceArchivePreviewDetections');
  if (meta) meta.innerHTML = `<strong>${escapeHtml(res.item.title || 'Untitled Service')}</strong><br>${new Date(res.item.serviceDate || res.item.savedAt || Date.now()).toLocaleString()} · ${escapeHtml(res.item.speaker || 'Unknown speaker')}`;
  if (transcript) transcript.textContent = (res.item.transcriptLines || []).slice(0,18).map(x => x.text || x.raw || '').join('\n') || 'No transcript';
  if (detections) detections.textContent = (res.item.detections || []).map(x => x.ref || '').filter(Boolean).join('\n') || 'No detected verses';
}

function openServiceArchiveCenter() {
  const modal = document.getElementById('serviceArchiveModal');
  if (modal) modal.style.display = 'flex';
  loadServiceArchiveIndex();
}

function closeServiceArchiveCenter() {
  const modal = document.getElementById('serviceArchiveModal');
  if (modal) modal.style.display = 'none';
}

function runServiceArchiveSearch() {
  const q = document.getElementById('serviceArchiveSearchInput')?.value || '';
  loadServiceArchiveIndex(q);
}

function loadArchivedServiceIntoReplay() {
  const item = State.serviceArchiveSelected;
  if (!item) return toast('⚠ Select an archived service first');
  State.replayTimeline = Array.isArray(item.replayTimeline) ? item.replayTimeline : [];
  State.replayIndex = 0;
  closeServiceArchiveCenter();
  openServiceReplayTimeline();
}

function openArchivedReplay() {
  loadArchivedServiceIntoReplay();
}

document.addEventListener('click', (e) => {
  const modal = document.getElementById('serviceArchiveModal');
  if (modal && e.target === modal) closeServiceArchiveCenter();
});


// ─── AUTO SERVICE BUILDER ─────────────────────────────────────────────
function openAutoServiceBuilder() {
  const modal = document.getElementById('autoServiceBuilderModal');
  if (modal) modal.style.display = 'flex';
  const speakerInput = document.getElementById('autoServiceSpeakerInput');
  if (speakerInput && !speakerInput.value) {
    speakerInput.value = document.getElementById('tmSpeakerSelect')?.selectedOptions?.[0]?.textContent || '';
  }
}

function closeAutoServiceBuilder() {
  const modal = document.getElementById('autoServiceBuilderModal');
  if (modal) modal.style.display = 'none';
}

async function runAutoServiceBuilder() {
  const speaker = (document.getElementById('autoServiceSpeakerInput')?.value || '').trim();
  const res = await window.electronAPI?.getAutoServiceSuggestions?.({ speaker, limit: 12 });
  if (!res?.success) return toast('⚠ Could not generate auto service suggestions');
  State.autoServiceSuggestions = res;
  renderAutoServiceBuilder();
  toast('🪄 Auto service suggestions ready');
}

function renderAutoServiceBuilder() {
  const res = State.autoServiceSuggestions;
  if (!res) return;

  const templateEl = document.getElementById('autoServiceTemplate');
  const songsEl = document.getElementById('autoServiceSongs');
  const versesEl = document.getElementById('autoServiceVerses');
  const momentsEl = document.getElementById('autoServiceMoments');

  if (templateEl) {
    templateEl.innerHTML = (res.template || []).map(slot => {
      const suggestions = (slot.suggestions || []).map(x => `<div style="font-size:12px;color:#edf1ff">${escapeHtml(x.value)} <span style="color:#9aa4c7">(${Number(x.count || 0)})</span></div>`).join('');
      return `<div style="padding:10px;border:1px solid #232838;border-radius:8px;background:#11131a"><div style="font-size:12px;color:#d4af37;font-weight:700;margin-bottom:6px">${escapeHtml(slot.slot)}</div>${suggestions || '<div style="font-size:12px;color:#9aa4c7">No suggestion</div>'}</div>`;
    }).join('');
  }

  const renderList = (items) => items && items.length
    ? items.map(x => `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px">${escapeHtml(x.value)} <span style="color:#9aa4c7">(${Number(x.count || 0)})</span></div>`).join('')
    : '<div style="padding:10px;border:1px dashed #2f3650;border-radius:8px;color:#9aa4c7">No suggestions yet.</div>';

  if (songsEl) songsEl.innerHTML = renderList(res.openingSongs || []);
  if (versesEl) versesEl.innerHTML = renderList(res.scriptures || []);
  if (momentsEl) momentsEl.innerHTML = renderList(res.likelyMoments || []);
}

async function buildAutoServiceSchedule() {
  const speaker = (document.getElementById('autoServiceSpeakerInput')?.value || '').trim();
  const title = `Auto Service ${new Date().toLocaleDateString()}`;
  const res = await window.electronAPI?.buildAutoServiceSchedule?.({ speaker, title, limit: 12 });
  if (!res?.success || !res?.schedule) return toast('⚠ Could not build auto service schedule');

  const sched = res.schedule;
  if (window.State) {
    State.schedule = sched;
    if (Array.isArray(sched.items)) {
      State.queue = sched.items.map(item => ({
        id: item.id,
        type: item.type,
        title: item.title,
        ref: item.ref || item.title
      }));
    }
  }
  closeAutoServiceBuilder();
  try { if (typeof renderQueue === 'function') renderQueue(); } catch (_) {}
  try { if (typeof renderSchedule === 'function') renderSchedule(); } catch (_) {}
  toast('🪄 Auto-built service schedule loaded');
}

document.addEventListener('click', (e) => {
  const modal = document.getElementById('autoServiceBuilderModal');
  if (modal && e.target === modal) closeAutoServiceBuilder();
});


// ─── LIVE SMART SUGGESTIONS ───────────────────────────────────────────
function _getLiveTranscriptSeed() {
  const lines = (State.transcriptLines || []).slice(-10);
  return lines.map(x => x.text || x.raw || '').join(' ').trim();
}

async function refreshLiveSmartSuggestions() {
  try {
    const transcript = _getLiveTranscriptSeed();
    const speaker = (document.getElementById('tmSpeakerSelect')?.selectedOptions?.[0]?.textContent || State.settings?.adaptiveSpeakerProfile || '').trim();
    const currentBook = String(window.currentLiveVerse?.book || State.previewVerse?.book || '').trim();
    const res = await window.electronAPI?.getLiveSmartSuggestions?.({
      transcript,
      speaker,
      currentBook,
      limit: 15
    });
    if (!res?.success) return;
    State.liveSmartSuggestions = res;
    renderLiveSmartSuggestions();
  } catch (_) {}
}

function _renderSuggestionList(items, type) {
  if (!items || !items.length) {
    return '<div style="padding:10px;border:1px dashed #2f3650;border-radius:8px;color:#9aa4c7">No suggestions yet.</div>';
  }
  return items.map(x => {
    const value = escapeHtml(x.value || '');
    const score = Number(x.score || 0);
    if (type === 'scripture') {
      return `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px"><div style="font-size:12px;color:#edf1ff">${value}</div><div style="display:flex;gap:8px;align-items:center;margin-top:4px"><span style="font-size:11px;color:#9aa4c7">score ${score}</span><button onclick="applyLiveSmartScripture('${String(x.value || '').replace(/'/g, '&#39;')}')" style="margin-left:auto">Use</button></div></div>`;
    }
    return `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px"><div style="font-size:12px;color:#edf1ff">${value}</div><div style="font-size:11px;color:#9aa4c7;margin-top:4px">score ${score}</div></div>`;
  }).join('');
}

function renderLiveSmartSuggestions() {
  const res = State.liveSmartSuggestions;
  if (!res) return;
  const scr = document.getElementById('liveSuggestionScriptures');
  const songs = document.getElementById('liveSuggestionSongs');
  const moments = document.getElementById('liveSuggestionMoments');
  if (scr) scr.innerHTML = _renderSuggestionList(res.scriptureSuggestions || [], 'scripture');
  if (songs) songs.innerHTML = _renderSuggestionList(res.songSuggestions || [], 'song');
  if (moments) moments.innerHTML = _renderSuggestionList(res.momentSuggestions || [], 'moment');
}

function openLiveSmartSuggestions() {
  const modal = document.getElementById('liveSmartSuggestionsModal');
  if (modal) modal.style.display = 'flex';
  refreshLiveSmartSuggestions();
  if (State.liveSmartSuggestionsTimer) clearInterval(State.liveSmartSuggestionsTimer);
  State.liveSmartSuggestionsTimer = setInterval(refreshLiveSmartSuggestions, 8000);
}

function closeLiveSmartSuggestions() {
  const modal = document.getElementById('liveSmartSuggestionsModal');
  if (modal) modal.style.display = 'none';
  if (State.liveSmartSuggestionsTimer) {
    clearInterval(State.liveSmartSuggestionsTimer);
    State.liveSmartSuggestionsTimer = null;
  }
}

function applyLiveSmartScripture(ref) {
  try {
    const parsed = window.BibleDB?.parseReference ? window.BibleDB.parseReference(ref) : null;
    if (!parsed) return toast('⚠ Could not parse suggested scripture');
    navigateBibleSearch(parsed.book, parsed.chapter, parsed.verse || 1);
    presentVerse(parsed.book, parsed.chapter, parsed.verse || 1, ref);
    toast('💡 Suggested scripture applied');
  } catch (_) {
    toast('⚠ Could not apply suggested scripture');
  }
}

document.addEventListener('click', (e) => {
  const modal = document.getElementById('liveSmartSuggestionsModal');
  if (modal && e.target === modal) closeLiveSmartSuggestions();
});


// ─── SERMON INTELLIGENCE ENGINE ───────────────────────────────────────
async function refreshSermonIntelligence() {
  try {
    const res = await window.electronAPI?.generateSermonIntelligence?.({
      transcriptLines: State.transcriptLines || [],
      detections: State.detections || [],
      useArchive: true
    });
    if (!res?.success) return toast('⚠ Could not generate sermon intelligence');
    State.sermonIntelligence = res;
    renderSermonIntelligence();
    toast('🧠 Sermon intelligence generated');
  } catch (_) {
    toast('⚠ Could not generate sermon intelligence');
  }
}

function _renderIntelList(items, formatter) {
  if (!items || !items.length) {
    return '<div style="padding:10px;border:1px dashed #2f3650;border-radius:8px;color:#9aa4c7">No data yet.</div>';
  }
  return items.map(formatter).join('');
}

function renderSermonIntelligence() {
  const res = State.sermonIntelligence;
  if (!res) return;
  const titles = document.getElementById('sermonIntelTitles');
  const keywords = document.getElementById('sermonIntelKeywords');
  const verses = document.getElementById('sermonIntelVerses');
  const points = document.getElementById('sermonIntelPoints');
  const structure = document.getElementById('sermonIntelStructure');
  const archiveThemes = document.getElementById('sermonIntelArchiveThemes');
  const stats = document.getElementById('sermonIntelStats');

  if (titles) titles.innerHTML = _renderIntelList(res.titleSuggestions || [], x => `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px">${escapeHtml(x)}</div>`);
  if (keywords) keywords.innerHTML = _renderIntelList(res.topKeywords || [], x => `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px">${escapeHtml(x.word)} <span style="color:#9aa4c7">(${Number(x.count || 0)})</span></div>`);
  if (verses) verses.innerHTML = _renderIntelList(res.recurringVerses || [], x => `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px">${escapeHtml(x.ref)} <span style="color:#9aa4c7">(${Number(x.count || 0)})</span></div>`);
  if (points) points.innerHTML = _renderIntelList(res.keyPoints || [], x => `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px">${escapeHtml(x)}</div>`);
  if (structure) {
    const s = res.structure || {};
    structure.innerHTML = `
      <div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px"><strong style="color:#d4af37">Intro</strong><br>${(s.intro || []).map(x => escapeHtml(x)).join('<br>') || 'None'}</div>
      <div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px"><strong style="color:#d4af37">Body</strong><br>${(s.body || []).map(x => escapeHtml(x)).join('<br>') || 'None'}</div>
      <div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px"><strong style="color:#d4af37">Closing</strong><br>${(s.closing || []).map(x => escapeHtml(x)).join('<br>') || 'None'}</div>
    `;
  }
  if (archiveThemes) archiveThemes.innerHTML = _renderIntelList(res.archiveThemes || [], x => `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px">${escapeHtml(x.ref)} <span style="color:#9aa4c7">(${Number(x.count || 0)})</span></div>`);
  if (stats) {
    const s = res.stats || {};
    stats.innerHTML = `
      <div style="display:flex;gap:14px;flex-wrap:wrap">
        <div>Transcript lines: <strong>${Number(s.transcriptLines || 0)}</strong></div>
        <div>Transcript words: <strong>${Number(s.transcriptWords || 0)}</strong></div>
        <div>Detections: <strong>${Number(s.detections || 0)}</strong></div>
        <div>Unique verses: <strong>${Number(s.uniqueVerses || 0)}</strong></div>
      </div>
    `;
  }
}

function openSermonIntelligence() {
  const modal = document.getElementById('sermonIntelligenceModal');
  if (modal) modal.style.display = 'flex';
  refreshSermonIntelligence();
}

function closeSermonIntelligence() {
  const modal = document.getElementById('sermonIntelligenceModal');
  if (modal) modal.style.display = 'none';
}

document.addEventListener('click', (e) => {
  const modal = document.getElementById('sermonIntelligenceModal');
  if (modal && e.target === modal) closeSermonIntelligence();
});


// ─── ANALYTICS DASHBOARD ──────────────────────────────────────────────
async function refreshAnalyticsDashboard() {
  try {
    const res = await window.electronAPI?.getAnalyticsDashboard?.({ limit: 120 });
    if (!res?.success) return toast('⚠ Could not load analytics dashboard');
    State.analyticsDashboard = res;
    renderAnalyticsDashboard();
    toast('📊 Analytics refreshed');
  } catch (_) {
    toast('⚠ Could not load analytics dashboard');
  }
}

function _renderAnalyticsList(items, labelKey) {
  if (!items || !items.length) {
    return '<div style="padding:10px;border:1px dashed #2f3650;border-radius:8px;color:#9aa4c7">No data yet.</div>';
  }
  return items.map(x => `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px">${escapeHtml(x[labelKey] || '')} <span style="color:#9aa4c7">(${Number(x.count || 0)})</span></div>`).join('');
}

function renderAnalyticsDashboard() {
  const res = State.analyticsDashboard;
  if (!res) return;

  const totals = document.getElementById('analyticsTotals');
  const books = document.getElementById('analyticsBooks');
  const verses = document.getElementById('analyticsVerses');
  const songs = document.getElementById('analyticsSongs');
  const speakers = document.getElementById('analyticsSpeakers');
  const keywords = document.getElementById('analyticsKeywords');
  const lengths = document.getElementById('analyticsServiceLengths');

  const t = res.totals || {};
  if (totals) {
    const cells = [
      ['Archived Services', t.archivedServices || 0],
      ['Detections', t.totalDetections || 0],
      ['Approved', t.totalApproved || 0],
      ['Rejected', t.totalRejected || 0],
      ['Avg Transcript Lines', t.avgTranscriptLines || 0],
    ];
    totals.innerHTML = cells.map(([label, value]) => `<div style="background:#171922;border:1px solid #232838;border-radius:10px;padding:12px"><div style="font-size:11px;color:#9aa4c7">${escapeHtml(label)}</div><div style="font-size:22px;font-weight:800;color:#edf1ff;margin-top:6px">${Number(value)}</div></div>`).join('');
  }

  if (books) books.innerHTML = _renderAnalyticsList(res.mostUsedBooks || [], 'book');
  if (verses) verses.innerHTML = _renderAnalyticsList(res.mostQuotedVerses || [], 'ref');
  if (songs) songs.innerHTML = _renderAnalyticsList(res.mostUsedSongs || [], 'title');
  if (speakers) speakers.innerHTML = _renderAnalyticsList(res.speakerPatterns || [], 'speaker');
  if (keywords) keywords.innerHTML = _renderAnalyticsList(res.keywordFrequency || [], 'keyword');
  if (lengths) {
    const items = res.serviceLengths || [];
    lengths.innerHTML = items.length
      ? items.map(x => `<div style="padding:8px;border:1px solid #232838;border-radius:8px;background:#11131a;margin-bottom:6px"><div style="font-size:12px;color:#edf1ff">${escapeHtml(x.title || 'Untitled Service')}</div><div style="font-size:11px;color:#9aa4c7">${new Date(x.serviceDate || Date.now()).toLocaleDateString()} · transcript ${Number(x.transcriptLines || 0)} · detections ${Number(x.detections || 0)} · replay ${Number(x.replayEvents || 0)}</div></div>`).join('')
      : '<div style="padding:10px;border:1px dashed #2f3650;border-radius:8px;color:#9aa4c7">No service length data yet.</div>';
  }
}

function openAnalyticsDashboard() {
  const modal = document.getElementById('analyticsDashboardModal');
  if (modal) modal.style.display = 'flex';
  refreshAnalyticsDashboard();
}

function closeAnalyticsDashboard() {
  const modal = document.getElementById('analyticsDashboardModal');
  if (modal) modal.style.display = 'none';
}

document.addEventListener('click', (e) => {
  const modal = document.getElementById('analyticsDashboardModal');
  if (modal && e.target === modal) closeAnalyticsDashboard();
});


// ─── CLIP GENERATOR ───────────────────────────────────────────────────
function _buildClipCandidates() {
  const lines = State.transcriptLines || [];
  const detections = State.detections || [];
  const replay = State.replayTimeline || [];
  const candidates = [];

  detections.slice(0, 20).forEach((det, idx) => {
    const ref = det.ref || '';
    const sourceText = String(det.sourceText || '').trim();
    let snippet = sourceText;
    if (!snippet && lines.length) {
      const near = lines.slice(Math.max(0, idx - 2), Math.min(lines.length, idx + 3)).map(x => x.text || x.raw || '').join(' ');
      snippet = near.trim();
    }
    const relatedReplay = replay.filter(ev => {
      const summary = String(ev?.payload?.summary || ev?.payload?.ref || ev?.payload?.text || '').toLowerCase();
      return (ref && summary.includes(ref.toLowerCase())) || (sourceText && summary.includes(sourceText.toLowerCase().slice(0, 20)));
    }).slice(0, 12);

    candidates.push({
      id: `clip_${Date.now()}_${idx}_${Math.random().toString(36).slice(2,5)}`,
      title: ref ? `Clip — ${ref}` : `Clip ${idx + 1}`,
      transcriptSnippet: snippet || 'No transcript snippet',
      verseRefs: ref ? [ref] : [],
      replayEvents: relatedReplay,
      score: (ref ? 3 : 0) + (snippet ? 2 : 0) + relatedReplay.length
    });
  });

  if (candidates.length < 4) {
    replay.slice(-20).forEach((ev, idx) => {
      const summary = String(ev?.payload?.summary || ev?.payload?.ref || ev?.payload?.text || '').trim();
      if (!summary) return;
      candidates.push({
        id: `clip_replay_${Date.now()}_${idx}_${Math.random().toString(36).slice(2,5)}`,
        title: `Moment — ${summary.slice(0,40)}`,
        transcriptSnippet: String(ev?.payload?.transcript || ev?.payload?.text || summary).trim(),
        verseRefs: ev?.payload?.ref ? [ev.payload.ref] : [],
        replayEvents: [ev],
        score: 1 + (ev?.payload?.ref ? 2 : 0)
      });
    });
  }

  const seen = new Set();
  return candidates
    .sort((a,b) => b.score - a.score)
    .filter(c => {
      const key = `${c.title}|${c.transcriptSnippet}`.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, 18);
}

function renderClipGenerator() {
  const list = document.getElementById('clipGeneratorList');
  const meta = document.getElementById('clipPreviewMeta');
  const transcript = document.getElementById('clipPreviewTranscript');
  const verses = document.getElementById('clipPreviewVerses');
  const titleInput = document.getElementById('clipTitleInput');
  if (!list) return;

  list.innerHTML = '';
  const candidates = State.clipCandidates || [];
  if (!candidates.length) {
    list.innerHTML = '<div style="padding:10px;border:1px dashed #2f3650;border-radius:8px;color:#9aa4c7">No clip candidates yet.</div>';
  } else {
    candidates.forEach(c => {
      const row = document.createElement('button');
      row.type = 'button';
      row.style.cssText = `text-align:left;background:${State.clipSelected?.id===c.id ? 'rgba(212,175,55,.12)' : '#11131a'};border:1px solid ${State.clipSelected?.id===c.id ? '#d4af37' : '#232838'};color:#edf1ff;border-radius:8px;padding:10px;cursor:pointer;margin-bottom:6px;width:100%`;
      row.innerHTML = `<div style="font-size:12px;font-weight:700">${escapeHtml(c.title || 'Clip')}</div><div style="font-size:11px;color:#9aa4c7">${escapeHtml(String(c.transcriptSnippet || '').slice(0,120))}</div><div style="font-size:11px;color:#9aa4c7;margin-top:4px">Verses ${Number((c.verseRefs || []).length)} · Replay events ${Number((c.replayEvents || []).length)} · Score ${Number(c.score || 0)}</div>`;
      row.addEventListener('click', () => {
        State.clipSelected = c;
        const input = document.getElementById('clipTitleInput');
        if (input) input.value = c.title || '';
        renderClipGenerator();
      });
      list.appendChild(row);
    });
  }

  const selected = State.clipSelected;
  if (meta) meta.textContent = selected ? `${selected.title} · replay ${Number((selected.replayEvents || []).length)} events` : 'No clip selected';
  if (transcript) transcript.textContent = selected ? (selected.transcriptSnippet || '') : '';
  if (verses) verses.textContent = selected ? ((selected.verseRefs || []).join('\n') || 'No verses linked') : '';
  if (titleInput && selected && !titleInput.value) titleInput.value = selected.title || '';
}

function refreshClipGenerator() {
  State.clipCandidates = _buildClipCandidates();
  State.clipSelected = (State.clipCandidates || [])[0] || null;
  const titleInput = document.getElementById('clipTitleInput');
  if (titleInput) titleInput.value = State.clipSelected?.title || '';
  renderClipGenerator();
  toast('🎬 Clip candidates generated');
}

function openClipGenerator() {
  const modal = document.getElementById('clipGeneratorModal');
  if (modal) modal.style.display = 'flex';
  refreshClipGenerator();
}

function closeClipGenerator() {
  const modal = document.getElementById('clipGeneratorModal');
  if (modal) modal.style.display = 'none';
}

async function saveSelectedClipPackage() {
  const selected = State.clipSelected;
  if (!selected) return toast('⚠ Select a clip first');
  const title = (document.getElementById('clipTitleInput')?.value || selected.title || '').trim() || 'Clip Package';
  const res = await window.electronAPI?.saveClipPackage?.({
    title,
    transcriptSnippet: selected.transcriptSnippet || '',
    verseRefs: selected.verseRefs || [],
    replayEvents: selected.replayEvents || [],
    source: 'clip-generator'
  });
  if (res?.success) {
    toast('🎬 Clip package saved');
  } else {
    toast('⚠ Could not save clip package');
  }
}

document.addEventListener('click', (e) => {
  const modal = document.getElementById('clipGeneratorModal');
  if (modal && e.target === modal) closeClipGenerator();
});

function _updateLogoOverlayControls() {
  const lo = State._logoOverlay;
  if (!lo) return;
  const sizeSlider = document.getElementById('logoOverlaySize');
  const sizeVal = document.getElementById('logoOverlaySizeVal');
  const opSlider = document.getElementById('logoOverlayOpacity');
  const opVal = document.getElementById('logoOverlayOpacityVal');
  const showBtn = document.getElementById('logoOverlayShowBtn');
  if (sizeSlider) sizeSlider.value = lo.sizePct || 15;
  if (sizeVal) sizeVal.textContent = (lo.sizePct || 15) + '%';
  if (opSlider) opSlider.value = lo.opacity ?? 100;
  if (opVal) opVal.textContent = (lo.opacity ?? 100) + '%';
  if (showBtn) showBtn.textContent = lo.visible !== false ? 'Showing' : 'Show';
  const posButtons = document.querySelectorAll('#logoOverlayPopup [data-pos]');
  posButtons.forEach(b => {
    b.style.background = b.dataset.pos === lo.position ? 'var(--gold)' : '';
    b.style.color = b.dataset.pos === lo.position ? '#000' : '';
  });
}

// ── Auto-Play Schedule ─────────────────────────────────────────────────────

let _autoPlayActive = false;
let _autoPlayTimer = null;
const AUTO_PLAY_IMAGE_HOLD = 5000;
const AUTO_PLAY_VERSE_HOLD = 8000;
const AUTO_PLAY_SONG_HOLD = 10000;

function toggleAutoPlay() {
  if (_autoPlayActive) stopAutoPlay();
  else startAutoPlay();
}

function startAutoPlay() {
  const realItems = (State.queue || []).filter(q => q.type !== 'section');
  if (realItems.length === 0) { toast('Schedule is empty'); return; }
  _autoPlayActive = true;
  _updateAutoPlayBtn();
  toast('▶ Auto-play started');
  const firstRealIdx = State.queue.findIndex(q => q.type !== 'section');
  if (firstRealIdx >= 0) {
    presentQueueItem(State.queue[firstRealIdx], firstRealIdx);
    _scheduleAutoAdvance(State.queue[firstRealIdx]);
  }
}

function stopAutoPlay(silent = false) {
  _autoPlayActive = false;
  if (_autoPlayTimer) { clearTimeout(_autoPlayTimer); _autoPlayTimer = null; }
  _updateAutoPlayBtn();
  _cleanupAutoPlayMediaListeners();
  if (!silent) toast('⏹ Auto-play stopped');
}

function _updateAutoPlayBtn() {
  const btn = document.getElementById('autoPlayBtn');
  if (!btn) return;
  if (_autoPlayActive) {
    btn.textContent = '⏹ Stop';
    btn.style.color = '#ff6b6b';
    btn.style.background = 'rgba(255,107,107,.12)';
    btn.style.borderColor = 'rgba(255,107,107,.3)';
  } else {
    btn.textContent = '▶ Auto';
    btn.style.color = '';
    btn.style.background = '';
    btn.style.borderColor = '';
  }
}

function _scheduleAutoAdvance(item) {
  if (!_autoPlayActive) return;
  if (_autoPlayTimer) { clearTimeout(_autoPlayTimer); _autoPlayTimer = null; }
  _cleanupAutoPlayMediaListeners();

  if (item.type === 'media') {
    const mediaItem = (State.media || []).find(m => String(m.id) === String(item.mediaId));
    if (mediaItem && (mediaItem.type === 'video' || mediaItem.type === 'audio')) {
      _waitForMediaEnd();
      return;
    }
    _autoPlayTimer = setTimeout(_autoAdvance, AUTO_PLAY_IMAGE_HOLD);
  } else if (item.type === 'song' || item.type === 'song-slide') {
    _autoPlayTimer = setTimeout(_autoAdvance, AUTO_PLAY_SONG_HOLD);
  } else if (item.type === 'presentation') {
    _autoPlayTimer = setTimeout(_autoAdvance, AUTO_PLAY_IMAGE_HOLD);
  } else {
    _autoPlayTimer = setTimeout(_autoAdvance, AUTO_PLAY_VERSE_HOLD);
  }
}

let _autoPlayMediaEndHandler = null;
function _waitForMediaEnd() {
  const liveDisplay = document.getElementById('liveDisplay');
  if (!liveDisplay) { _autoPlayTimer = setTimeout(_autoAdvance, AUTO_PLAY_IMAGE_HOLD); return; }
  const mediaEl = liveDisplay.querySelector('video, audio');
  if (!mediaEl) { _autoPlayTimer = setTimeout(_autoAdvance, AUTO_PLAY_IMAGE_HOLD); return; }
  mediaEl.loop = false;
  _autoPlayMediaEndHandler = () => {
    _autoPlayMediaEndHandler = null;
    if (_autoPlayActive) _autoAdvance();
  };
  mediaEl.addEventListener('ended', _autoPlayMediaEndHandler, { once: true });
}

function _cleanupAutoPlayMediaListeners() {
  if (_autoPlayMediaEndHandler) {
    const liveDisplay = document.getElementById('liveDisplay');
    const mediaEl = liveDisplay?.querySelector('video, audio');
    if (mediaEl) mediaEl.removeEventListener('ended', _autoPlayMediaEndHandler);
    _autoPlayMediaEndHandler = null;
  }
}

function _autoAdvance() {
  if (!_autoPlayActive) return;
  const currentIdx = State.currentQueueIdx ?? -1;
  let nextIdx = currentIdx + 1;
  while (nextIdx < State.queue.length && State.queue[nextIdx].type === 'section') nextIdx++;
  if (nextIdx >= State.queue.length) {
    stopAutoPlay(true); // silent — we show our own "finished" toast below
    // Clear the projection screen when autoplay reaches the end of the schedule.
    // Force-clear state and projection directly so the guard in clearLive()
    // ("Nothing is live") doesn't short-circuit when the last item was a section.
    State.liveVerse       = null;
    State.liveContentType = null;
    State.currentMediaId  = null;
    State.isLive          = false;
    const goBtn = document.getElementById('goLiveBtn');
    const goTxt = document.getElementById('goLiveTxt');
    const goTag = document.getElementById('liveTag');
    if (goBtn) goBtn.classList.remove('active');
    if (goTxt) goTxt.textContent = 'Project Live';
    if (goTag) goTag.style.display = 'none';
    const ld = document.getElementById('liveDisplay');
    if (ld) {
      ld.querySelectorAll('video,audio').forEach(m => { m.pause(); m.src = ''; });
      ld.innerHTML = '';
      ld.style.display = 'none';
    }
    const liveEmpty = document.getElementById('liveEmpty');
    if (liveEmpty) liveEmpty.style.display = 'flex';
    if (window.electronAPI) window.electronAPI.clearProjection();
    else _webPostRenderState('clear', null);
    State.currentQueueIdx = null;
    renderQueue();
    if (typeof _syncRemoteLiveState === 'function') _syncRemoteLiveState();
    toast('⏹ Auto-play finished — screen cleared');
    return;
  }
  presentQueueItem(State.queue[nextIdx], nextIdx);
  _scheduleAutoAdvance(State.queue[nextIdx]);
}

// ── Preset System ──────────────────────────────────────────────────────────

function togglePresetPopup() {
  const pop = document.getElementById('presetPopup');
  if (!pop) return;
  const opened = _toggleLivePopup('presetPopup');
  if (opened) _refreshPresetList();
}

async function _refreshPresetList() {
  const list = document.getElementById('presetList');
  if (!list) return;
  let presets = [];
  try { presets = await window.electronAPI.getPresets(); } catch(_) {}
  if (!Array.isArray(presets) || presets.length === 0) {
    list.innerHTML = '<div style="font-size:10px;color:var(--text-dim);padding:4px 0">No presets saved</div>';
    return;
  }
  list.innerHTML = '';
  presets.forEach(p => {
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;align-items:center;gap:6px;padding:3px 6px;border-radius:4px;cursor:pointer;font-size:11px;color:var(--text)';
    row.onmouseenter = () => row.style.background = 'var(--hover)';
    row.onmouseleave = () => row.style.background = '';
    const name = document.createElement('span');
    name.style.cssText = 'flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap';
    name.textContent = p.name || 'Untitled';
    const info = document.createElement('span');
    info.style.cssText = 'font-size:9px;color:var(--text-dim);white-space:nowrap';
    const parts = [];
    if (p.settings) parts.push('Set');
    if (p.theme) parts.push('Thm');
    if (p.schedule?.items?.length) parts.push('Sch');
    if (p.bgMediaId) parts.push('BG');
    if (p.logoOverlay?.src || p.logoOverlay?.savedPath) parts.push('Logo');
    info.textContent = parts.join('+') || 'empty';
    const loadBtn = document.createElement('button');
    loadBtn.className = 'btn-sm';
    loadBtn.style.cssText = 'font-size:9px;padding:2px 6px';
    loadBtn.textContent = 'Load';
    loadBtn.onclick = (e) => { e.stopPropagation(); loadPreset(p); };
    row.appendChild(name);
    row.appendChild(info);
    row.appendChild(loadBtn);
    list.appendChild(row);
  });
}

function savePresetDialog() {
  const overlay = document.getElementById('savePresetOverlay');
  if (!overlay) return;
  overlay.classList.add('show');
  const inp = document.getElementById('presetNameInput');
  if (inp) { inp.value = ''; inp.focus(); }
  const settingsInfo = document.getElementById('presetSaveSettingsInfo');
  if (settingsInfo) settingsInfo.innerHTML = 'Settings: <strong>all current preferences</strong>';
  const themeInfo = document.getElementById('presetSaveThemeInfo');
  if (themeInfo) {
    const themeId = State.currentTheme || 'sanctuary';
    themeInfo.innerHTML = 'Theme: <strong>' + themeId + '</strong>';
  }
  const schedInfo = document.getElementById('presetSaveScheduleInfo');
  if (schedInfo) {
    const items = (State.queue || []).filter(q => q.type !== 'section');
    if (items.length > 0) {
      schedInfo.innerHTML = 'Schedule: <strong>' + (State.scheduleName || 'Untitled') + '</strong> (' + items.length + ' items)';
    } else {
      schedInfo.innerHTML = 'Schedule: <em>empty (will clear schedule on load)</em>';
    }
  }
  const bgInfo = document.getElementById('presetSaveBgInfo');
  if (bgInfo) {
    const bgId = State.settings?.logoMediaId;
    const bgItem = bgId ? State.media.find(m => m.id === bgId) : null;
    bgInfo.innerHTML = 'Background: ' + (bgItem ? `<strong>${bgItem.name || bgItem.fileName}</strong>` : '<em>none</em>');
  }
  const logoInfo = document.getElementById('presetSaveLogoInfo');
  if (logoInfo) {
    const lo = State._logoOverlay;
    logoInfo.innerHTML = 'Logo Overlay: ' + (lo?.src ? `<strong>${lo.fileName || 'image'}</strong> (${lo.sizePct || 15}%, ${lo.opacity || 100}% opacity)` : '<em>none</em>');
  }
}

function closeSavePresetDialog() {
  const overlay = document.getElementById('savePresetOverlay');
  if (overlay) overlay.classList.remove('show');
}

async function confirmSavePreset() {
  const nameInput = document.getElementById('presetNameInput');
  const name = (nameInput?.value || '').trim();
  if (!name) { toast('Please enter a preset name'); if (nameInput) nameInput.focus(); return; }
  const settingsSnapshot = State.settings ? JSON.parse(JSON.stringify(State.settings)) : {};
  delete settingsSnapshot.apiKey;
  const preset = {
    id: 'preset_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6),
    name,
    createdAt: Date.now(),
    bgMediaId: State.settings?.logoMediaId || null,
    schedule: {
      name: State.scheduleName || null,
      items: JSON.parse(JSON.stringify(State.queue || [])),
    },
    settings: settingsSnapshot,
    theme: State.currentTheme || null,
    songTheme: State.currentSongTheme || null,
    presTheme: State.currentPresTheme || null,
  };
  const lo = State._logoOverlay;
  if (lo?.src) {
    const posDefaults = { 'top-left': [5,5], 'top-right': [95,5], 'bottom-left': [5,95], 'bottom-right': [95,95], 'center': [50,50] };
    const posDef = posDefaults[lo.position || 'top-right'] || [95,5];
    const overlayData = {
      fileName: lo.fileName || 'logo.png',
      type: lo.type || 'image',
      position: lo.position || 'top-right',
      xPct: lo.xPct ?? posDef[0],
      yPct: lo.yPct ?? posDef[1],
      sizePct: lo.sizePct || 15,
      opacity: lo.opacity ?? 100,
      visible: lo.visible !== false,
    };
    try {
      const resp = await fetch(lo.src);
      const blob = await resp.blob();
      const reader = new FileReader();
      const b64 = await new Promise((resolve, reject) => {
        reader.onload = () => {
          const dataUrl = reader.result;
          resolve(dataUrl.split(',')[1]);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
      overlayData.fileData = b64;
    } catch(e) {
      console.warn('Could not read logo overlay file for preset', e);
    }
    preset.logoOverlay = overlayData;
  }
  try {
    const res = await window.electronAPI.savePreset(preset);
    if (res?.success) {
      toast('Preset saved: ' + name);
      closeSavePresetDialog();
      _refreshPresetList();
    } else {
      toast('Failed to save preset: ' + (res?.error || 'unknown error'));
    }
  } catch(e) {
    toast('Error saving preset');
    console.error(e);
  }
}

async function loadPreset(preset) {
  if (!preset) return;
  console.log('[Preset] Loading preset:', preset.name, 'keys:', Object.keys(preset));
  console.log('[Preset] Has schedule:', !!preset.schedule, 'items:', preset.schedule?.items?.length);
  console.log('[Preset] Has logoOverlay:', !!preset.logoOverlay, 'savedPath:', preset.logoOverlay?.savedPath);
  const loaded = [];

  // Helper: persist current State.settings to disk
  const _persist = async () => {
    try {
      if (window.electronAPI?.saveSettings) {
        await window.electronAPI.saveSettings(State.settings);
      }
    } catch(e) { console.warn('[Preset] Could not persist settings:', e.message); }
  };

  if (preset.settings && typeof preset.settings === 'object') {
    const saved = JSON.parse(JSON.stringify(preset.settings));
    delete saved.apiKey;
    const currentKey = State.settings?.apiKey;
    Object.assign(State.settings, saved);
    if (currentKey) State.settings.apiKey = currentKey;
    applySettings(State.settings);
    await _persist();
    loaded.push('Settings');
  }

  if (preset.theme) {
    if (typeof setTheme === 'function') setTheme(preset.theme);
    loaded.push('Theme');
  }
  if (preset.songTheme) State.currentSongTheme = preset.songTheme;
  if (preset.presTheme) State.currentPresTheme = preset.presTheme;

  if (preset.schedule) {
    try {
      State.queue = Array.isArray(preset.schedule.items) ? JSON.parse(JSON.stringify(preset.schedule.items)) : [];
      State.scheduleName = preset.schedule.name || null;
      console.log('[Preset] Restored queue with', State.queue.length, 'items, scheduleName:', State.scheduleName);
      if (typeof renderQueue === 'function') renderQueue();
      if (typeof updateScheduleNameBar === 'function') updateScheduleNameBar();
      loaded.push('Schedule' + (preset.schedule.name ? ` "${preset.schedule.name}"` : ''));
    } catch(e) { console.error('[Preset] Schedule load error:', e); }
  }

  if (preset.bgMediaId) {
    const item = State.media.find(m => m.id === preset.bgMediaId);
    if (item) {
      State.settings.logoMediaId = item.id;
      await _persist();
      _updateBgLogoBtnState();
      loaded.push('Background');
    } else {
      toast('Background media not found in library');
    }
  } else {
    if (State.settings.logoMediaId) {
      State.settings.logoMediaId = null;
      await _persist();
      _updateBgLogoBtnState();
    }
  }

  if (preset.logoOverlay) {
    try {
      const lo = preset.logoOverlay;
      let src = lo.src || '';
      if (lo.savedPath) {
        if (!window.electronAPI?.isWeb) {
          // Ensure file:/// with three slashes for absolute paths (required on Windows)
          const normalized = lo.savedPath.replace(/\\/g, '/');
          src = normalized.startsWith('/') ? 'file://' + normalized : 'file:///' + normalized;
        } else {
          const fileName = lo.savedPath.split(/[\\\/]/).pop();
          src = '/preset-assets/' + encodeURIComponent(fileName);
        }
      }
      if (!src && lo.fileData) {
        const mimeType = (lo.type === 'image') ? 'image/png' : 'video/mp4';
        src = 'data:' + mimeType + ';base64,' + lo.fileData;
      }
      console.log('[Preset] Logo src:', src ? src.substring(0, 80) : '(empty)');
      const posDefaults = { 'top-left': [5,5], 'top-right': [95,5], 'bottom-left': [5,95], 'bottom-right': [95,95], 'center': [50,50] };
      const pos = lo.position || 'top-right';
      const defaults = posDefaults[pos] || [95,5];
      State._logoOverlay = {
        src,
        type: lo.type || 'image',
        fileName: lo.fileName || 'logo.png',
        position: pos,
        xPct: lo.xPct ?? defaults[0],
        yPct: lo.yPct ?? defaults[1],
        sizePct: lo.sizePct || 15,
        opacity: lo.opacity ?? 100,
        visible: lo.visible !== false,
      };
      if (src) {
        _sendLogoOverlay();
        _updateLogoOverlayControls();
        _updateLogoOverlayBtn();
        loaded.push('Logo Overlay');
      } else {
        console.warn('[Preset] Logo overlay has no valid source');
      }
    } catch(e) { console.error('[Preset] Logo load error:', e); }
  } else {
    if (State._logoOverlay?.src) {
      State._logoOverlay = { src: '', visible: false, sizePct: 15, opacity: 100, position: 'top-right' };
      _sendLogoOverlay();
      _updateLogoOverlayControls();
    }
  }

  toast('✓ Preset loaded: ' + (loaded.length ? loaded.join(', ') : preset.name));
}

async function openManagePresets() {
  const overlay = document.getElementById('managePresetsOverlay');
  if (!overlay) return;
  overlay.classList.add('show');
  await _refreshManagePresetsList();
}

function closeManagePresets() {
  const overlay = document.getElementById('managePresetsOverlay');
  if (overlay) overlay.classList.remove('show');
}

async function _refreshManagePresetsList() {
  const list = document.getElementById('managePresetsList');
  if (!list) return;
  let presets = [];
  try { presets = await window.electronAPI.getPresets(); } catch(_) {}
  if (!Array.isArray(presets) || presets.length === 0) {
    list.innerHTML = '<div style="font-size:12px;color:var(--text-dim);padding:12px 0;text-align:center">No presets saved</div>';
    return;
  }
  list.innerHTML = '';
  presets.forEach(p => {
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;align-items:center;gap:8px;padding:8px 10px;border-radius:6px;background:var(--deep);border:1px solid var(--border)';
    const info = document.createElement('div');
    info.style.cssText = 'flex:1;min-width:0';
    const nameEl = document.createElement('div');
    nameEl.style.cssText = 'font-size:12px;font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap';
    nameEl.textContent = p.name || 'Untitled';
    const meta = document.createElement('div');
    meta.style.cssText = 'font-size:10px;color:var(--text-dim);margin-top:2px';
    const parts = [];
    if (p.settings) parts.push('Settings');
    if (p.theme) parts.push('Theme');
    if (p.schedule?.items?.length) parts.push('Schedule (' + p.schedule.items.filter(i => i.type !== 'section').length + ')');
    if (p.bgMediaId) parts.push('BG');
    if (p.logoOverlay?.savedPath || p.logoOverlay?.src) parts.push('Logo');
    meta.textContent = parts.join(' + ') || 'Empty preset';
    if (p.createdAt) meta.textContent += ' \u2022 ' + new Date(p.createdAt).toLocaleDateString();
    info.appendChild(nameEl);
    info.appendChild(meta);
    const loadBtn = document.createElement('button');
    loadBtn.className = 'btn-sm';
    loadBtn.style.cssText = 'font-size:10px;padding:4px 10px';
    loadBtn.textContent = 'Load';
    loadBtn.onclick = () => { loadPreset(p); closeManagePresets(); };
    const delBtn = document.createElement('button');
    delBtn.className = 'btn-sm';
    delBtn.style.cssText = 'font-size:10px;padding:4px 8px;color:#ff6b6b';
    delBtn.textContent = 'Delete';
    delBtn.onclick = async () => {
      if (delBtn.dataset.confirmPending) {
        try {
          await window.electronAPI.deletePreset(p.id);
          toast('Preset deleted');
          _refreshManagePresetsList();
          _refreshPresetList();
        } catch(e) { toast('Error deleting preset'); }
      } else {
        delBtn.dataset.confirmPending = '1';
        delBtn.textContent = 'Confirm?';
        delBtn.style.color = '#fff';
        delBtn.style.background = '#c0392b';
        setTimeout(() => {
          delete delBtn.dataset.confirmPending;
          delBtn.textContent = 'Delete';
          delBtn.style.color = '#ff6b6b';
          delBtn.style.background = '';
        }, 3000);
      }
    };
    row.appendChild(info);
    row.appendChild(loadBtn);
    row.appendChild(delBtn);
    list.appendChild(row);
  });
}

window.electronAPI?.on?.('license-file-opened', () => { _consumePendingLicenseImportIntoWindow(); });

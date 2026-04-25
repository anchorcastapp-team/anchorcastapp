
  // ── All DOM refs — declared first so all functions can access them ─────────
  const stage       = document.getElementById('stage');
  const empty       = document.getElementById('emptyState');
  const refLabel    = document.getElementById('refLabel');
  const verseEl     = document.getElementById('verseText');
  const transEl     = document.getElementById('transLabel');
  const mediaLayer  = document.getElementById('mediaLayer');
  const mediaImg    = document.getElementById('mediaImg');
  const mediaVideo  = document.getElementById('mediaVideo');
  const captionLayer= document.getElementById('captionLayer');
  const captionText = document.getElementById('captionText');
  const timerLayer  = document.getElementById('timerLayer');
  const timerDisp   = document.getElementById('timerDisplay');
  const timerLabel  = document.getElementById('timerLabel');
  const fadeOverlay = document.getElementById('fadeOverlay');
  const partBadge   = document.getElementById('partBadge');


  function clearProjectionMediaState() {
    if (partBadge) { partBadge.style.display = 'none'; partBadge.textContent = ''; }
    mediaVideo.pause?.();
    mediaVideo.src = '';
    mediaVideo.style.display = 'none';
    mediaImg.src = '';
    mediaImg.style.display = 'none';
    const oldAudio = mediaLayer.querySelector('audio');
    if (oldAudio) { oldAudio.pause?.(); oldAudio.src = ''; oldAudio.remove(); }
    const oldOverlay = mediaLayer.querySelector('.audio-overlay');
    if (oldOverlay) oldOverlay.remove();
    mediaLayer.querySelectorAll('.created-slide-content').forEach(e => e.remove());
    document.querySelectorAll('.theme-box-el').forEach(e => e.remove());
    const bgVid = document.getElementById('themeBgVideo');
    if (bgVid) { bgVid.pause?.(); bgVid.src = ''; bgVid.style.display = 'none'; }
    const bgOv = document.getElementById('themeBgOverlay');
    if (bgOv) bgOv.style.background = 'transparent';
    mediaLayer.style.background = '#000';
    mediaLayer.style.backgroundImage = 'none';
    mediaLayer.classList.remove('active');
    stage.style.display = 'none';
    stage.classList.remove('out');
  }


  function updatePartBadge() {
    if (!partBadge) return;
    partBadge.style.display = 'none';
    partBadge.textContent = '';
  }

  function getBoxContent(box, data) {
    const role = (box.role || '').toLowerCase();
    if (role === 'main') {
      if (Array.isArray(data.lines)) return data.lines.join('\n');
      return data.text || box.text || '';
    }
    if (role === 'ref') {
      return data.ref ? `${data.ref}${data.translation ? ' · ' + data.translation : ''}` : (box.text || '');
    }
    if (role === 'title') {
      return data.title || data.name || box.text || '';
    }
    if (role === 'subtitle') {
      return data.sectionLabel || data.subtitle || box.text || '';
    }
    return box.text || '';
  }

  function getBoxTextTransform(box, data, theme) {
    return data.songTextTransform
      || data.scriptureTextTransform
      || box.textTransform
      || theme?.textTransform
      || 'none';
  }

  function fitProjectionBoxFontSize(box, content) {
    const raw = String(content || '').trim();
    const base = Number(box?.fontSize || 52);
    if (!raw) return base;
    const width = Math.max(120, Number(box?.w || 1680) - 40);
    const height = Math.max(80, Number(box?.h || 680) - 40);
    const lineSpacing = Number(box?.lineSpacing || 1.4);
    let size = base;
    while (size > 16) {
      const approxCharWidth = Math.max(5, size * 0.5);
      const charsPerLine = Math.max(6, Math.floor(width / approxCharWidth));
      const wrappedLines = raw.split(/\n+/).reduce((sum, line) => sum + Math.max(1, Math.ceil((line.length || 1) / charsPerLine)), 0);
      const maxLines = Math.max(1, Math.floor(height / (size * lineSpacing)));
      if (wrappedLines <= maxLines) return size;
      size -= 2;
    }
    return Math.max(16, size);
  }

  function playProjectionVideo(videoEl) {
    if (!videoEl) return;
    const forceAudioState = () => {
      const wantMuted = !!videoEl.dataset.wantMuted && videoEl.dataset.wantMuted !== 'false';
      videoEl.muted = wantMuted;
      videoEl.defaultMuted = wantMuted;
      if (wantMuted) videoEl.setAttribute('muted', '');
      else videoEl.removeAttribute('muted');
      const vol = Math.max(0, Math.min(1, Number(videoEl.dataset.wantVolume || '1')));
      videoEl.volume = vol;
    };
    forceAudioState();
    videoEl.onloadedmetadata = () => {
      forceAudioState();
      videoEl.play().catch(() => {});
    };
    setTimeout(() => { forceAudioState(); videoEl.play().catch(() => {}); }, 40);
    setTimeout(() => { forceAudioState(); videoEl.play().catch(() => {}); }, 250);
  }

  function applyTheme(data) {
    const t = data.themeData;
    const bg = document.querySelector('.bg');
    if (!t) {
      document.body.setAttribute('data-theme', data.theme || 'sanctuary');
      bg.style.background = '';
      verseEl.style.cssText = '';
      refLabel.style.color  = '';
      transEl.style.color   = '';
      stage.style.textAlign = 'center';
      stage.style.padding   = '80px 120px';
      return;
    }

    // ── Background ──
    if (t.bgType === 'image' && t.bgImage) {
      bg.style.background = `url('${t.bgImage}') center/cover no-repeat`;
      bg.style.boxShadow  = `inset 0 0 0 9999px rgba(0,0,0,${(t.bgOverlay||0)/100})`;
      // Stop any video
      mediaVideo.pause?.(); mediaVideo.src = '';
    } else if (t.bgType === 'video' && t.bgVideo) {
      // Video plays in .bg layer behind stage text (z-index:2)
      bg.style.background = '#000';
      bg.style.overflow   = 'hidden';
      bg.style.boxShadow  = 'none';
      // Overlay darkness on top of video
      const ovAlpha = (t.bgOverlay || 0) / 100;
      // Create or reuse dedicated bg video element
      let bgVid = document.getElementById('themeBgVideo');
      if (!bgVid) {
        bgVid = document.createElement('video');
        bgVid.id       = 'themeBgVideo';
        bgVid.loop     = true;
        bgVid.muted    = true;
        bgVid.playsInline = true;
        bgVid.autoplay = true;
        bgVid.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;' +
          'object-fit:cover;z-index:0;pointer-events:none;display:block';
        bg.appendChild(bgVid);
      }
      // Always update src (file:// paths are stable, comparison safe)
      bgVid.src = t.bgVideo;
      bgVid.style.display = 'block';
      bgVid.play().catch(e => console.warn('bgVid play:', e));
      // Overlay div for darkness
      let bgOv = document.getElementById('themeBgOverlay');
      if (!bgOv) {
        bgOv = document.createElement('div');
        bgOv.id = 'themeBgOverlay';
        bgOv.style.cssText = 'position:absolute;inset:0;z-index:1;pointer-events:none';
        bg.appendChild(bgOv);
      }
      bgOv.style.background = ovAlpha > 0 ? `rgba(0,0,0,${ovAlpha})` : 'transparent';
    } else {
      bg.style.boxShadow = 'none';
      mediaVideo.pause?.(); mediaVideo.src = '';
      // Stop any theme bg video
      const bgVid = document.getElementById('themeBgVideo');
      if (bgVid) { bgVid.pause(); bgVid.src = ''; bgVid.style.display = 'none'; }
      if (t.bgType === 'solid')  bg.style.background = t.bgColor1 || '#000';
      else if (t.bgType === 'linear') bg.style.background = `linear-gradient(135deg, ${t.bgColor1}, ${t.bgColor2||'#000'})`;
      else bg.style.background = `radial-gradient(ellipse at 35% 45%, ${t.bgColor1||'#10082a'} 0%, ${t.bgColor2||'#000'} 65%, #000 100%)`;
    }

    // ── New box-based theme format ──
    if (t.boxes && t.boxes.length > 0) {
      const W = 1920, H = 1080;
      const toP = (v, total) => (v / total * 100).toFixed(3) + '%';

      // Hide default stage elements — we render all boxes manually
      refLabel.style.cssText  = 'display:none';
      transEl.style.cssText   = 'display:none';

      // Remove any previously rendered theme boxes
      document.querySelectorAll('.theme-box-el').forEach(e => e.remove());

      // Render every box absolutely positioned on the projection
      t.boxes.forEach(box => {
        const el = document.createElement('div');
        el.className = 'theme-box-el';
        el.style.cssText = `
          position:absolute; overflow:hidden; box-sizing:border-box;
          left:${toP(box.x, W)}; top:${toP(box.y, H)};
          width:${toP(box.w, W)}; height:${toP(box.h, H)};
          display:flex; flex-direction:column;
          align-items:${box.align==='left'?'flex-start':box.align==='right'?'flex-end':'center'};
          justify-content:${box.valign==='top'?'flex-start':box.valign==='bottom'?'flex-end':'center'};
          text-align:${box.align||'center'};
          padding:20px;
          ${box.bgOpacity>0 ? `background:${_hexToRgba(box.bgFill||'#000',box.bgOpacity/100)};` : ''}
          ${box.borderW>0 ? `border:${box.borderW}px solid ${box.borderColor||'#fff'};` : ''}
          ${box.borderRadius>0 ? `border-radius:${box.borderRadius}px;` : ''}
          z-index:3;
        `;
        const inner = document.createElement('div');
        inner.style.cssText = `
          font-family:'${box.fontFamily||'Crimson Pro'}',Georgia,serif;
          font-size:clamp(16px,${fitProjectionBoxFontSize(box, getBoxContent(box, data))/16}vw,${fitProjectionBoxFontSize(box, getBoxContent(box, data))}px);
          line-height:${box.lineSpacing||1.6};
          color:${box.color||'#fff'};
          font-weight:${box.bold?'700':'400'};
          font-style:${box.italic?'italic':'normal'};
          text-shadow:0 2px 12px rgba(0,0,0,0.9);
          width:100%;
          text-align:${box.align||'center'};
          white-space:pre-wrap;
          text-transform:${getBoxTextTransform(box, data, t)};
        `;
        inner.textContent = getBoxContent(box, data);
        el.appendChild(inner);
        document.body.appendChild(el);
      });

      // Hide the default stage (we use custom boxes)
      stage.style.display = 'none';
      return;
    }

    // ── Legacy property-based format (backward compat) ──
    const fontSize = `clamp(24px, ${(t.fontSize||52)/16}vw, ${t.fontSize||52}px)`;
    verseEl.style.cssText = `
      font-family: '${t.fontFamily || 'Crimson Pro'}', Georgia, serif;
      font-size: ${fontSize}; line-height: 1.6; color: ${t.textColor || '#ede6d8'};
      font-style: ${t.fontStyle || 'normal'}; text-transform: ${t.textTransform || 'none'};
      text-shadow: 0 2px 40px rgba(0,0,0,0.95); max-width: 1200px;
    `;
    refLabel.style.color = t.refColor || t.accentColor || '#c9a84c';
    transEl.style.color  = t.transColor || '#6a5220';
    stage.style.textAlign = t.textAlign || 'center';
    stage.style.padding   = `${t.padding || 80}px ${Math.round((t.padding||80)*1.5)}px`;
  }

  function _hexToRgba(hex, alpha) {
    const r=parseInt(hex.slice(1,3)||'0',16),g=parseInt(hex.slice(3,5)||'0',16),b=parseInt(hex.slice(5,7)||'0',16);
    return `rgba(${r},${g},${b},${alpha})`;
  }

  function toProjectionMediaUrl(input) {
    const raw = String(input || '').trim();
    if (!raw) return '';
    if (/^(media|file|https?|blob|data):/i.test(raw)) return raw;
    const normalized = raw.replace(/\\/g, '/');
    return `media:///${normalized.startsWith('/') ? normalized.slice(1) : normalized}`;
  }

  function getMediaPath(data) {
    if (!data || typeof data !== 'object') return '';
    return data.path || data.filePath || data.src || data.url || '';
  }

  let lastProjectedBackgroundMedia = null;

function showMediaBackground(data) {
  const mediaPath = getMediaPath(data);
  if (!data || !mediaPath) return false;
  lastProjectedBackgroundMedia = { ...data, path: mediaPath };
  mediaLayer.querySelectorAll('.audio-overlay').forEach(e => e.remove());
  mediaImg.style.display = 'none';
  mediaVideo.style.display = 'none';
  mediaVideo.pause?.();
  mediaVideo.src = '';
  const filePath = mediaPath.replace(/\\/g, '/');
  const fileUrl  = `file:///${filePath.startsWith('/') ? filePath.slice(1) : filePath}`;
  if (data.type === 'video') {
    mediaVideo.src = fileUrl;
    mediaVideo.loop = data.loop !== false;
    mediaVideo.dataset.wantMuted = String(!!data.mute);
    mediaVideo.dataset.wantVolume = String(Math.max(0, Math.min(1, Number.isFinite(Number(data.volume)) ? Number(data.volume) : 1)));
    mediaVideo.style.objectFit = data.aspectRatio === 'cover' ? 'cover' : 'contain';
    mediaVideo.style.display = 'block';
    mediaVideo.load();
    playProjectionVideo(mediaVideo);
  } else if (data.type === 'image') {
    mediaImg.src = fileUrl;
    mediaImg.style.objectFit = data.aspectRatio === 'cover' ? 'cover' : 'contain';
    mediaImg.style.display = 'block';
  } else if (data.type === 'audio') {
    const overlay = document.createElement('div');
    overlay.className = 'audio-overlay';
    overlay.style.cssText = `
      position:absolute;inset:0;display:flex;flex-direction:column;
      align-items:center;justify-content:center;gap:20px;
      background:radial-gradient(ellipse at 50% 50%,#1a0a2e 0%,#000 100%);
    `;
    overlay.innerHTML = `
      <div style="font-size:80px;opacity:.8">🎵</div>
      <div style="font-family:'Cinzel',serif;font-size:clamp(18px,3vw,42px);
        color:#ede6d8;text-align:center;padding:0 60px;line-height:1.4">
        ${escapeHtml(data.name || '')}
      </div>`;
    const aud = document.createElement('audio');
    aud.loop = data.loop !== false;
    aud.muted = false;
    aud.volume = Math.max(0, Math.min(1, (data.volume ?? 100) / 100));
    aud.src = fileUrl;
    aud.load();
    aud.play().catch(()=>{});
    overlay.appendChild(aud);
    mediaLayer.appendChild(overlay);
  }
  mediaLayer.classList.add('active');
  empty.style.display = 'none';
  return true;
}

  function applyOverlayPresentation(options = {}) {
  const bgOv = document.getElementById('themeBgOverlay');
  const enabled = !!(options.preserveMedia && options.backgroundMedia);
  const dim = Math.max(0, Math.min(90, Number(options.overlayOptions?.dim ?? 35))) / 100;
  if (bgOv) bgOv.style.background = enabled ? `rgba(0,0,0,${dim})` : 'rgba(0,0,0,0)';
  if (options.overlayOptions?.lowerThird) {
    stage.style.justifyContent = 'flex-end';
    stage.style.paddingTop = '40px';
    stage.style.paddingBottom = '6vh';
  } else {
    stage.style.justifyContent = 'center';
    stage.style.paddingTop = '';
    stage.style.paddingBottom = '';
  }
}

function showVerse(data, options = {}) {
    if (options.preserveMedia && options.backgroundMedia) { document.querySelectorAll('.theme-box-el').forEach(e => e.remove()); showMediaBackground(options.backgroundMedia); }
    else { clearProjectionMediaState(); }
    applyOverlayPresentation(options);
    // Remove any leftover custom box elements first
    document.querySelectorAll('.theme-box-el').forEach(e => e.remove());
    updatePartBadge(data);

    try {
      applyTheme(data);
    } catch (err) {
      console.warn('Projection scripture theme render failed, using fallback', err);
      document.querySelectorAll('.theme-box-el').forEach(e => e.remove());
      refLabel.style.cssText = '';
      transEl.style.cssText = '';
      verseEl.style.cssText = '';
      stage.style.textAlign = 'center';
      stage.style.padding = '80px 120px';
      document.body.setAttribute('data-theme', data.theme || 'sanctuary');
    }

    // Box-based theme: applyTheme already rendered all boxes — stage is hidden
    if (data.themeData?.boxes?.length && document.querySelectorAll('.theme-box-el').length) {
      empty.style.display = 'none';
      return;
    }

    // Legacy theme: use stage with position-aware ref label
    const refPos = data.scriptureRefPosition || 'top';
    const scriptureTransform = (data.themeData?.textTransform && data.themeData.textTransform !== 'none')
      ? (data.scriptureTextTransform || data.themeData.textTransform)
      : (data.scriptureTextTransform || 'none');
    const refText = `${data.ref}${data.translation ? '  ·  ' + data.translation : ''}`;
    const verseNum = data.themeData?.showVerseNum !== false
      ? `<span class="verse-num">${data.verse}</span>` : '';
    const verseHtml = verseNum + (data.text || '');

    // Reset stage
    stage.style.justifyContent = 'center';
    stage.style.paddingTop     = '';
    stage.style.paddingBottom  = '';
    verseEl.style.textTransform = scriptureTransform;

    if (refPos === 'hidden') {
      refLabel.style.display = 'none';
      transEl.style.display  = 'none';
    } else if (refPos === 'bottom') {
      // Move refLabel below verse text by reordering flex children
      refLabel.style.display = '';
      transEl.style.display  = 'none';
      refLabel.style.order = '2';
      refLabel.style.marginBottom = '0';
      refLabel.style.marginTop    = 'clamp(20px,3vh,50px)';
      verseEl.style.order  = '1';
      transEl.style.order  = '3';
      refLabel.textContent = refText;
    } else {
      // top (default)
      refLabel.style.display = '';
      transEl.style.display  = 'none';
      refLabel.style.order = '';
      refLabel.style.marginBottom = '';
      refLabel.style.marginTop    = '';
      verseEl.style.order  = '';
      transEl.style.order  = '';
      refLabel.textContent = refText;
    }

    verseEl.innerHTML = verseHtml;
    transEl.textContent = '';
    transEl.style.display = 'none';
    empty.style.display = 'none';
    stage.style.display = 'flex';
    stage.classList.remove('out');
  }


  function sanitizeProjectionSongHtml(html) {
    return String(html || '')
      .replace(/<(?!\/?(?:b|strong|i|em|u|span|div|p|br)\b)[^>]*>/gi, '')
      .replace(/ on\w+="[^"]*"/gi, '')
      .replace(/ on\w+='[^']*'/gi, '');
  }

  function showSong(data, options = {}) {
    if (options.preserveMedia && options.backgroundMedia) showMediaBackground(options.backgroundMedia);
    else clearProjectionMediaState();
    applyOverlayPresentation(options);
    updatePartBadge({ partTotal: 1, partIdx: 0 });
    try {
      applyTheme(data);
    } catch (err) {
      console.warn('Projection song theme render failed, using fallback', err);
      document.querySelectorAll('.theme-box-el').forEach(e => e.remove());
      refLabel.style.cssText = '';
      transEl.style.cssText = '';
      verseEl.style.cssText = '';
      stage.style.textAlign = 'center';
      stage.style.padding = '80px 120px';
      document.body.setAttribute('data-theme', data.theme || 'sanctuary');
    }
    refLabel.textContent = '';
    refLabel.style.display = 'none';
    transEl.textContent = '';
    transEl.style.display = 'none';

    const t = data.themeData;
    if (t?.boxes?.length && document.querySelectorAll('.theme-box-el').length) {
      empty.style.display = 'none';
      return;
    }
    const ls = t?.lineSpacing || 1.4;
    const sh = t?.shadowOn !== false ? '0 2px 16px rgba(0,0,0,.95)' : 'none';
    verseEl.style.lineHeight = String(ls);
    verseEl.style.textShadow = sh;
    const richHtml = data.html ? sanitizeProjectionSongHtml(data.html) : '';
    verseEl.innerHTML = richHtml || (data.lines || []).map(l =>
      l.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    ).join('<br>');

    if (data.showSongMetadata && data.sectionLabel) {
      refLabel.style.cssText = `display:block;font-size:clamp(11px,1.2vw,18px);
        letter-spacing:.18em;text-transform:uppercase;opacity:.7;margin-bottom:1.2em;
        color:${t?.accentColor||'#c9a84c'}`;
      refLabel.textContent = `${data.title || ''} · ${data.sectionLabel}`;
    }

    const align = data.songVerticalAlign || 'center';
    const justifyMap = { top:'flex-start', center:'center', bottom:'flex-end' };
    stage.style.justifyContent = justifyMap[align] || 'center';
    stage.style.paddingTop    = align === 'top'    ? 'clamp(40px,8vh,120px)' : '';
    stage.style.paddingBottom = align === 'bottom' ? 'clamp(40px,8vh,120px)' : '';
    verseEl.style.textTransform = data.songTextTransform || (data.sectionTextTransform && data.sectionTextTransform !== 'inherit' ? data.sectionTextTransform : (t?.textTransform || 'none'));

    empty.style.display = 'none';
    stage.style.display = 'flex';
    stage.classList.remove('out');
  }

    // ── Fade helper ─────────────────────────────────────────────────────────────
  function fadeTransition(fn) {
    fadeOverlay.classList.add('fading');
    setTimeout(() => {
      fn();
      fadeOverlay.classList.remove('fading');
    }, 300);
  }

  // ── showCreatedSlide (editor-built slides) ────────────────────────────────────
  function showCreatedSlide(data) {
    const s = data.slide || {};
    fadeTransition(() => {
      stage.classList.add('out');
      stage.style.display = 'none';
      mediaVideo.pause(); mediaVideo.src = '';
      mediaVideo.style.display = 'none';
      mediaImg.style.display   = 'none';
      const oldAudio = mediaLayer.querySelector('audio');
      if (oldAudio) { oldAudio.pause(); oldAudio.src = ''; oldAudio.remove(); }

      // Slide background
      mediaLayer.style.background = s.bg || '#0a0a1e';
      mediaLayer.style.backgroundImage = s.bgImage
        ? `url('${toFileUrl(s.bgImage)}')` : 'none';
      mediaLayer.style.backgroundSize     = 'cover';
      mediaLayer.style.backgroundPosition = 'center';
      mediaLayer.style.display     = 'block';
      mediaLayer.style.position    = 'relative';
      mediaLayer.classList.add('active');

      // Remove old rendered content
      mediaLayer.querySelectorAll('.created-slide-content').forEach(e => e.remove());

      const PW = 1920, PH = 1080;
      const W  = mediaLayer.offsetWidth  || window.innerWidth;
      const H  = mediaLayer.offsetHeight || window.innerHeight;
      const sc = Math.min(W / PW, H / PH);

      // Container centred on screen
      const container = document.createElement('div');
      container.className = 'created-slide-content';
      container.style.cssText = `
        position:absolute;
        left:${(W - PW*sc)/2}px; top:${(H - PH*sc)/2}px;
        width:${PW*sc}px; height:${PH*sc}px;
        overflow:hidden; pointer-events:none;`;

      // Render each object
      (s.objects || []).forEach(obj => {
        const el = document.createElement('div');
        el.style.cssText = `
          position:absolute;
          left:${obj.x*sc}px; top:${obj.y*sc}px;
          width:${obj.w*sc}px; height:${obj.h*sc}px;
          transform:rotate(${obj.rotation||0}deg);
          overflow:hidden;`;

        if (obj.type === 'text') {
          const alpha = (obj.textBgAlpha||0) / 100;
          if (alpha > 0) el.style.background = hexToRgba(obj.textBg||'#000000', alpha);
          el.style.display        = 'flex';
          el.style.alignItems     = 'center';
          el.style.justifyContent = obj.textAlign === 'left' ? 'flex-start'
            : obj.textAlign === 'right' ? 'flex-end' : 'center';

          const t = document.createElement('div');
          t.style.cssText = `
            font-family:'${obj.font||'Cinzel'}',serif;
            font-size:${(obj.fontSize||48)*sc}px;
            color:${obj.textColor||'#ffffff'};
            font-weight:${obj.bold?'700':'400'};
            font-style:${obj.italic?'italic':'normal'};
            text-align:${obj.textAlign||'center'};
            text-shadow:${obj.shadow?'0 2px 8px rgba(0,0,0,.8)':'none'};
            white-space:pre-wrap; line-height:1.25; width:100%;`;
          t.textContent = obj.text || '';
          el.appendChild(t);

        } else if (obj.type === 'shape') {
          const opacity  = (obj.shapeOpacity ?? 80) / 100;
          const radius   = obj.shape === 'ellipse' ? '50%' : ((obj.shapeRadius||0) + '%');
          const bw       = (obj.shapeBorderW||0) * sc;
          el.style.background   = obj.shapeFill || '#3b82f6';
          el.style.opacity      = opacity;
          el.style.borderRadius = radius;
          if (bw > 0) el.style.border = `${bw}px solid ${obj.shapeBorder||'#fff'}`;

        } else if (obj.type === 'line') {
          const bw = Math.max(1, (obj.shapeBorderW||2) * sc);
          el.style.borderTop = `${bw}px solid ${obj.shapeFill||'#fff'}`;
          el.style.opacity   = (obj.shapeOpacity??80)/100;
          el.style.height    = `${bw}px`;
          el.style.top       = `${(obj.y + obj.h/2) * sc}px`;

        } else if (obj.type === 'svg') {
          // SVG path shapes — star, triangle, diamond, cross, arrow, hexagon, heart, etc.
          el.style.overflow = 'visible';
          el.style.opacity  = String((obj.shapeOpacity ?? 90) / 100);
          const fill    = obj.shapeFill  || '#c9a84c';
          const stroke  = (obj.shapeBorderW||0) > 0 ? (obj.shapeBorder||'#fff') : 'none';
          const strokeW = (obj.shapeBorderW||0) * sc;
          const sw = obj.w * sc, sh = obj.h * sc;
          const cx = sw/2, cy = sh/2;
          let pathD = '';
          if (obj.shape === 'star') {
            const or=Math.min(sw,sh)/2, ir=or*0.42; let d='';
            for(let p=0;p<10;p++){const r=p%2===0?or:ir;const a=(p*Math.PI/5)-Math.PI/2;d+=(p===0?'M':'L')+(cx+r*Math.cos(a)).toFixed(2)+','+(cy+r*Math.sin(a)).toFixed(2);} pathD=d+'Z';
          } else if (obj.shape==='triangle') { pathD=`M${cx},0 L${sw},${sh} L0,${sh} Z`; }
          else if (obj.shape==='diamond')    { pathD=`M${cx},0 L${sw},${cy} L${cx},${sh} L0,${cy} Z`; }
          else if (obj.shape==='cross')      { const t=sw*.3,u=sh*.3; pathD=`M${(sw-t)/2},0 L${(sw+t)/2},0 L${(sw+t)/2},${(sh-u)/2} L${sw},${(sh-u)/2} L${sw},${(sh+u)/2} L${(sw+t)/2},${(sh+u)/2} L${(sw+t)/2},${sh} L${(sw-t)/2},${sh} L${(sw-t)/2},${(sh+u)/2} L0,${(sh+u)/2} L0,${(sh-u)/2} L${(sw-t)/2},${(sh-u)/2} Z`; }
          else if (obj.shape==='arrow')      { const aw=sw*.4,ah=sh*.5; pathD=`M0,${(sh-ah)/2} L${sw-aw},${(sh-ah)/2} L${sw-aw},0 L${sw},${cy} L${sw-aw},${sh} L${sw-aw},${(sh+ah)/2} L0,${(sh+ah)/2} Z`; }
          else if (obj.shape==='hexagon')    { const r=Math.min(sw,sh)/2; let d=''; for(let i=0;i<6;i++){const a=i*Math.PI/3-Math.PI/6;d+=(i===0?'M':'L')+(cx+r*Math.cos(a)).toFixed(2)+','+(cy+r*Math.sin(a)).toFixed(2);} pathD=d+'Z'; }
          else if (obj.shape==='heart')      { pathD=`M${cx},${sh*.3} C${cx},${sh*.1} 0,${sh*.1} 0,${sh*.4} C0,${sh*.7} ${cx},${sh*.9} ${cx},${sh} C${cx},${sh*.9} ${sw},${sh*.7} ${sw},${sh*.4} C${sw},${sh*.1} ${cx},${sh*.1} ${cx},${sh*.3}`; }
          else if (obj.shape==='octagon')    { const r=Math.min(sw,sh)/2; let d=''; for(let i=0;i<8;i++){const a=i*Math.PI/4-Math.PI/8;d+=(i===0?'M':'L')+(cx+r*Math.cos(a)).toFixed(2)+','+(cy+r*Math.sin(a)).toFixed(2);} pathD=d+'Z'; }
          else if (obj.shape==='rhombus')    { pathD=`M${cx},0 L${sw},${cy} L${cx},${sh} L0,${cy} Z`; }
          else if (obj.shape==='speech')     { pathD=`M0,0 L${sw},0 L${sw},${sh*.7} L${sw*.4},${sh*.7} L${sw*.25},${sh} L${sw*.2},${sh*.7} L0,${sh*.7} Z`; }
          else if (obj.shape==='cloud')      { pathD=`M${sw*.25},${sh*.8} Q${sw*.05},${sh*.8} ${sw*.05},${sh*.55} Q${sw*.05},${sh*.3} ${sw*.25},${sh*.3} Q${sw*.3},${sh*.05} ${sw*.55},${sh*.1} Q${sw*.75},${sh*.05} ${sw*.8},${sh*.3} Q${sw},${sh*.3} ${sw*.95},${sh*.55} Q${sw*.95},${sh*.8} ${sw*.75},${sh*.8} Z`; }
          else if (obj.shape==='banner')     { pathD=`M0,${sh*.2} Q${sw*.05},0 ${sw*.1},${sh*.2} L${sw*.9},${sh*.2} Q${sw*.95},0 ${sw},${sh*.2} L${sw},${sh*.8} Q${sw*.95},${sh} ${sw*.9},${sh*.8} L${sw*.1},${sh*.8} Q${sw*.05},${sh} 0,${sh*.8} Z`; }
          else { pathD=`M0,0 L${sw},0 L${sw},${sh} L0,${sh} Z`; }

          el.innerHTML = `<svg width="${sw}" height="${sh}" viewBox="0 0 ${sw} ${sh}" xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;overflow:visible"><path d="${pathD}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeW}"/></svg>`;

        } else if (obj.type === 'image') {
          const img = document.createElement('img');
          img.src = toFileUrl(obj.src || '');
          img.style.cssText = `width:100%;height:100%;
            object-fit:${obj.imgFit||'contain'};
            opacity:${(obj.imgOpacity??100)/100};display:block;`;
          el.appendChild(img);
        }
        container.appendChild(el);
      });

      mediaLayer.appendChild(container);
      empty.style.display = 'none';
    });
  }

  function toFileUrl(p) {
    if (!p) return '';
    if (p.startsWith('http') || p.startsWith('file://') || p.startsWith('media://')) return p;
    return 'file:///' + p.replace(/\\/g, '/');
  }

  function hexToRgba(hex, alpha) {
    const r = parseInt(hex.slice(1,3),16)||0;
    const g = parseInt(hex.slice(3,5),16)||0;
    const b = parseInt(hex.slice(5,7),16)||0;
    return `rgba(${r},${g},${b},${alpha})`;
  }

  // ── showPresentationSlide ─────────────────────────────────────────────────────
  function showPresentationSlide(data) {
    fadeTransition(() => {
      // Hide verse/song stage
      stage.classList.add('out');
      stage.style.display = 'none';

      clearProjectionMediaState();

      // Normalize file path to file:// URL
      const filePath = (data.imagePath || '').replace(/\\/g, '/');
      const fileUrl  = `file:///${filePath.startsWith('/') ? filePath.slice(1) : filePath}`;

      // Show the slide image using the media layer
      mediaImg.src           = fileUrl;
      mediaImg.style.display = 'block';
      mediaImg.style.objectFit = 'contain'; // keep slide aspect ratio
      mediaLayer.style.background = '#000'; // black bars instead of cut
      mediaLayer.classList.add('active');

      empty.style.display = 'none';

      // Keep caption visible if present
      // (timer stays as-is)
    });
  }

  // ── showMedia ────────────────────────────────────────────────────────────────
  function showMedia(data) {
    fadeTransition(() => {
      // Hide verse stage
      stage.classList.add('out');
      stage.style.display = 'none';

      clearProjectionMediaState();

      // Normalize path — works on both Windows and Mac
      const filePath = mediaPath.replace(/\\/g, '/');
      const fileUrl  = `file:///${filePath.startsWith('/') ? filePath.slice(1) : filePath}`;

      if (data.type === 'audio') {
        // Audio-only: show full-screen music note + song name, play audio
        const overlay = document.createElement('div');
        overlay.className = 'audio-overlay';
        overlay.style.cssText = `
          position:absolute;inset:0;display:flex;flex-direction:column;
          align-items:center;justify-content:center;gap:20px;
          background:radial-gradient(ellipse at 50% 50%,#1a0a2e 0%,#000 100%);
        `;
        overlay.innerHTML = `
          <div style="font-size:80px;opacity:.8">🎵</div>
          <div style="font-family:'Cinzel',serif;font-size:clamp(18px,3vw,42px);
            color:#ede6d8;text-align:center;padding:0 60px;line-height:1.4">
            ${escapeHtml(data.name || '')}
          </div>`;
        const aud = document.createElement('audio');
        aud.loop  = data.loop !== false;
        aud.muted = false; // audio is the point — never mute
        aud.volume = Math.max(0, Math.min(1, (data.volume ?? 100) / 100));
        aud.src   = fileUrl;
        aud.load();
        aud.play().catch(()=>{});
        overlay.appendChild(aud);
        mediaLayer.appendChild(overlay);

      } else if (data.type === 'video') {
        mediaVideo.src   = fileUrl;
        mediaVideo.loop  = data.loop !== false;
        mediaVideo.dataset.wantMuted = String(!!data.mute);
        mediaVideo.dataset.wantVolume = String(Math.max(0, Math.min(1, Number.isFinite(Number(data.volume)) ? Number(data.volume) : 1)));
        mediaVideo.style.objectFit = data.aspectRatio === 'cover' ? 'cover' : 'contain';
        mediaVideo.style.display = 'block';
        mediaVideo.load();
        playProjectionVideo(mediaVideo);

      } else {
        // Image
        mediaImg.src           = fileUrl;
        mediaImg.style.objectFit = data.aspectRatio === 'cover' ? 'cover' : 'contain';
        mediaImg.style.display = 'block';
      }

      mediaLayer.classList.add('active');
      empty.style.display = 'none';

      // Caption
      if (data.caption) {
        captionText.textContent = data.caption;
        captionLayer.classList.add('active');
      } else {
        captionLayer.classList.remove('active');
      }
    });
  }

  // Simple HTML-escape for projection use
  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  // ── clearVerse (also clears media + audio) ───────────────────────────────────
  function clearVerse() {
    fadeTransition(() => {
      stage.classList.add('out');
      setTimeout(() => { stage.style.display = 'none'; stage.classList.remove('out'); }, 200);

      clearProjectionMediaState();
      captionLayer.classList.remove('active');
      empty.style.display = 'flex';
    });
  }

  // ── Timer engine ─────────────────────────────────────────────────────────────
  let timerInterval = null;
  let timerEndTime  = null;
  let timerMode     = 'countdown'; // 'countdown' | 'countup'

  function startTimer(data) {
    stopTimer();
    timerMode    = data.mode || 'countdown';
    timerEndTime = data.mode === 'countdown'
      ? Date.now() + (data.seconds || 0) * 1000
      : Date.now();
    timerLabel.textContent = data.label || (timerMode === 'countdown' ? 'COUNTDOWN' : 'ELAPSED');
    timerLayer.classList.add('active');
    updateTimer();
    timerInterval = setInterval(updateTimer, 500);
  }

  function stopTimer() {
    clearInterval(timerInterval);
    timerInterval = null;
    timerLayer.classList.remove('active');
    timerDisp.classList.remove('warning', 'overtime');
  }

  function updateTimer() {
    const now     = Date.now();
    let   elapsed = 0;

    if (timerMode === 'countdown') {
      elapsed = Math.max(0, timerEndTime - now);
      const secs = Math.ceil(elapsed / 1000);
      timerDisp.textContent = formatTime(secs);
      timerDisp.classList.toggle('warning',  secs <= 60 && secs > 0);
      timerDisp.classList.toggle('overtime', secs === 0);
    } else {
      elapsed = now - timerEndTime;
      timerDisp.textContent = formatTime(Math.floor(elapsed / 1000));
    }
  }

  function formatTime(totalSecs) {
    const h = Math.floor(totalSecs / 3600);
    const m = Math.floor((totalSecs % 3600) / 60);
    const s = totalSecs % 60;
    if (h > 0) return `${h}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
    return `${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
  }


  // ── Central projection render engine ────────────────────────────────────────
  let lastRenderState = { module: 'clear', payload: null, backgroundMedia: null, overlayOptions: {} };
  function normalizeRenderState(state) {
    if (!state || typeof state !== 'object') return { module: 'clear', payload: null, backgroundMedia: null, overlayOptions: {} };
    const overlayOptions = state.overlayOptions || {};
    const explicitBg = state.backgroundMedia || null;
    const effectiveBg = explicitBg || ((overlayOptions.enabled || overlayOptions.preserveMedia) ? lastProjectedBackgroundMedia : null);
    return {
      module: state.module || 'clear',
      payload: state.payload ?? null,
      updatedAt: state.updatedAt || Date.now(),
      backgroundMedia: effectiveBg,
      overlayOptions,
    };
  }
  function dispatchRenderState(rawState) {
    const state = normalizeRenderState(rawState);
    lastRenderState = state;
    switch (state.module) {
      case 'scripture':
        showVerse(state.payload || {}, { preserveMedia: !!state.backgroundMedia, backgroundMedia: state.backgroundMedia || null, overlayOptions: state.overlayOptions || {} });
        break;
      case 'song':
        showSong(state.payload || {}, { preserveMedia: !!state.backgroundMedia, backgroundMedia: state.backgroundMedia || null, overlayOptions: state.overlayOptions || {} });
        break;
      case 'media':
        lastProjectedBackgroundMedia = state.payload || null;
        showMedia(state.payload || {});
        break;
      case 'presentation-slide':
        showPresentationSlide(state.payload || {});
        break;
      case 'created-slide':
        showCreatedSlide(state.payload || {});
        break;
      case 'clear':
      default:
        clearVerse();
        break;
    }
  }

  // ── Receive from main process ────────────────────────────────────────────────
  if (window.electronAPI) {
    window.electronAPI.on('show-verse',  showVerse);
    window.electronAPI.on('show-song',   showSong);
    window.electronAPI.on('show-media',  showMedia);
    window.electronAPI.on('show-presentation-slide', showPresentationSlide);
    window.electronAPI.on('show-created-slide',      showCreatedSlide);
    window.electronAPI.on('render-state', dispatchRenderState);
    window.electronAPI.on('clear-verse', clearVerse);
    window.electronAPI.on('show-timer',  startTimer);
    window.electronAPI.on('stop-timer',  stopTimer);
    window.electronAPI.on('show-caption', (data) => {
      captionText.textContent = data.text || '';
      captionLayer.classList.toggle('active', !!data.text);
    });
    window.electronAPI.getCurrentRenderState?.().then(dispatchRenderState).catch(() => {});
  }

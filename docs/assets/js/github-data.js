(function(){
  'use strict';

  var doc = document;
  var html = doc.documentElement;
  var API_ENDPOINT = '/api/github';
  var CACHE_KEY = 'anchorcast.github-data.v3';
  var RELEASES_URL = 'https://github.com/anchorcastapp-team/anchorcastapp/releases/latest';
  var CACHE_MAX_AGE = 7 * 24 * 60 * 60 * 1000;

  html.classList.add('github-loading');

  function selectAll(selector){ return Array.prototype.slice.call(doc.querySelectorAll(selector)); }
  function setState(state){
    html.classList.remove('github-loading','github-ready','github-stale','github-error');
    html.classList.add('github-' + state);
  }
  function setText(selector, value){
    selectAll(selector).forEach(function(el){ el.textContent = value; });
  }
  function formatNumber(value){
    var n = Number(value);
    return Number.isFinite(n) ? n.toLocaleString() : 'Unavailable';
  }
  function formatBytes(value){
    var n = Number(value);
    if(!Number.isFinite(n) || n <= 0) return '';
    var units = ['B','KB','MB','GB'];
    var i = Math.min(units.length - 1, Math.floor(Math.log(n) / Math.log(1024)));
    return (n / Math.pow(1024, i)).toFixed(i > 1 ? 1 : 0) + ' ' + units[i];
  }
  function formatDate(value){
    if(!value) return 'Date unavailable';
    var date = new Date(value);
    if(Number.isNaN(date.getTime())) return 'Date unavailable';
    return new Intl.DateTimeFormat(undefined,{year:'numeric',month:'short',day:'numeric'}).format(date);
  }
  function versionWithoutV(tag){ return String(tag || '').replace(/^v/i,''); }
  function releaseSummary(body){
    var text = String(body || '')
      .replace(/```[\s\S]*?```/g,' ')
      .replace(/`([^`]+)`/g,'$1')
      .replace(/!\[[^\]]*\]\([^)]*\)/g,' ')
      .replace(/\[([^\]]+)\]\([^)]*\)/g,'$1')
      .replace(/^#{1,6}\s+/gm,'')
      .replace(/^[-*+]\s+/gm,'')
      .replace(/^>\s?/gm,'')
      .replace(/[\*_~]/g,'')
      .replace(/\s+/g,' ')
      .trim();
    if(!text) return 'See the GitHub release notes for everything included in this release.';
    return text.length > 280 ? text.slice(0,277).trim() + '…' : text;
  }
  function animateNumber(selector, value){
    var target = Number(value);
    if(!Number.isFinite(target)){
      setText(selector,'Unavailable');
      return;
    }
    var els = selectAll(selector);
    if(!els.length) return;
    var started = performance.now();
    var duration = 650;
    function frame(now){
      var p = Math.min(1,(now-started)/duration);
      var eased = 1-Math.pow(1-p,3);
      var text = Math.round(target*eased).toLocaleString();
      els.forEach(function(el){ el.textContent = text; });
      if(p < 1) requestAnimationFrame(frame);
    }
    requestAnimationFrame(frame);
  }

  function slotSpec(slot){
    var specs = {
      'windows-full': {platform:'windows',variant:'full',arch:null},
      'windows-light': {platform:'windows',variant:'light',arch:null},
      'windows-portable': {platform:'windows',variant:'portable',arch:null},
      'macos-full-arm64': {platform:'macos',variant:'full',arch:'arm64'},
      'macos-full-x64': {platform:'macos',variant:'full',arch:'x64'},
      'macos-light-arm64': {platform:'macos',variant:'light',arch:'arm64'},
      'macos-light-x64': {platform:'macos',variant:'light',arch:'x64'},
      'macos-full-universal': {platform:'macos',variant:'full',arch:'universal'},
      'macos-light-universal': {platform:'macos',variant:'light',arch:'universal'}
    };
    return specs[slot] || null;
  }
  function scoreAsset(asset, spec){
    if(!asset || !spec) return -10000;
    var name = String(asset.name || '').toLowerCase();
    var ext = String(asset.extension || name.split('.').pop() || '').toLowerCase();
    if(['yml','yaml','blockmap','sha256','sha512','txt'].indexOf(ext) >= 0) return -10000;
    var score = 0;
    if(asset.platform === spec.platform) score += 100; else score -= 1000;
    if(spec.variant){
      if(asset.variant === spec.variant) score += 55;
      else if(asset.variant === 'standard' && spec.variant === 'full') score += 18;
      else score -= 24;
    }
    if(spec.arch){
      if(asset.arch === spec.arch) score += 35;
      else if(asset.arch === 'universal') score += 12;
      else score -= 15;
    }
    if(spec.platform === 'windows' && ext === 'exe') score += 16;
    if(spec.platform === 'macos' && ext === 'dmg') score += 16;
    if(name.indexOf('setup') >= 0 || name.indexOf('anchorcast') >= 0) score += 5;
    return score;
  }
  function findAsset(assets, slot){
    var spec = slotSpec(slot);
    if(!spec || !Array.isArray(assets)) return null;
    var scored = assets.map(function(asset){ return {asset:asset,score:scoreAsset(asset,spec)}; })
      .sort(function(a,b){ return b.score-a.score; });
    return scored.length && scored[0].score >= 70 ? scored[0].asset : null;
  }
  function applyDownloadLinks(latest){
    var assets = Array.isArray(latest.assets) ? latest.assets : [];
    selectAll('[data-download]').forEach(function(link){
      var slot = link.getAttribute('data-download');
      var asset = findAsset(assets, slot);
      link.setAttribute('aria-busy','false');
      link.classList.remove('is-unavailable');
      if(asset){
        link.href = asset.url;
        link.removeAttribute('aria-disabled');
        link.title = asset.name + (asset.downloadCount != null ? ' · ' + formatNumber(asset.downloadCount) + ' downloads' : '');
      }else{
        link.href = latest.url || RELEASES_URL;
        link.classList.add('is-unavailable');
        link.setAttribute('aria-disabled','true');
        link.title = 'This build was not found in the latest release. Open the release page to see available files.';
      }
    });
    selectAll('[data-asset-name]').forEach(function(el){
      var asset = findAsset(assets, el.getAttribute('data-asset-name'));
      el.textContent = asset ? asset.name : 'Not included in this release';
    });
  }
  function renderAssets(latest){
    selectAll('[data-release-assets]').forEach(function(container){
      container.textContent = '';
      var assets = (latest.assets || []).filter(function(asset){
        return ['yml','yaml','blockmap'].indexOf(String(asset.extension || '').toLowerCase()) < 0;
      });
      if(!assets.length){
        var empty = doc.createElement('div');
        empty.className = 'release-assets-empty';
        empty.textContent = 'No downloadable assets are currently attached to this release.';
        container.appendChild(empty);
        return;
      }
      assets.forEach(function(asset){
        var card = doc.createElement('div');
        card.className = 'release-asset-live';
        var main = doc.createElement('div');
        main.className = 'release-asset-live-main';
        var name = doc.createElement('div');
        name.className = 'release-asset-live-name';
        name.textContent = asset.name;
        var meta = doc.createElement('div');
        meta.className = 'release-asset-live-meta';
        var details = [];
        var size = formatBytes(asset.size);
        if(size) details.push(size);
        details.push(formatNumber(asset.downloadCount) + ' downloads');
        meta.textContent = details.join(' · ');
        main.appendChild(name); main.appendChild(meta);
        var button = doc.createElement('a');
        button.className = 'btn-sm-dl';
        button.href = asset.url;
        button.textContent = 'Download';
        button.setAttribute('rel','noopener');
        card.appendChild(main); card.appendChild(button);
        container.appendChild(card);
      });
    });
  }
  function updateStructuredData(data){
    var latest = data.latest || {};
    var script = doc.getElementById('anchorcast-structured-data');
    if(!script){
      script = doc.createElement('script');
      script.type = 'application/ld+json';
      script.id = 'anchorcast-structured-data';
      doc.head.appendChild(script);
    }
    script.textContent = JSON.stringify({
      '@context':'https://schema.org',
      '@type':'SoftwareApplication',
      name:'AnchorCast',
      applicationCategory:'MultimediaApplication',
      operatingSystem:'Windows 10, Windows 11, macOS 12 or later',
      softwareVersion:versionWithoutV(latest.tag),
      downloadUrl:latest.url || RELEASES_URL,
      codeRepository:(data.repo && data.repo.url) || 'https://github.com/anchorcastapp-team/anchorcastapp',
      offers:{'@type':'Offer',price:'0',priceCurrency:'USD'},
      license:'https://opensource.org/license/mit'
    });
  }
  function applyData(data, stale){
    var latest = data.latest || {};
    var repo = data.repo || {};
    var totals = data.totals || {};
    var tag = latest.tag || 'Latest';

    setText('[data-github="version"]',tag);
    setText('[data-github="version-number"]',versionWithoutV(tag));
    animateNumber('[data-github="downloads"]',totals.downloads);
    animateNumber('[data-github="stars"]',repo.stars);
    animateNumber('[data-github="forks"]',repo.forks);
    setText('[data-github="release-date"]',formatDate(latest.publishedAt));
    setText('[data-github="release-name"]',latest.name || ('AnchorCast ' + tag));
    setText('[data-github="release-summary"]',releaseSummary(latest.body));
    setText('[data-github="asset-count"]',formatNumber((latest.assets || []).length));
    setText('[data-github="latest-downloads"]',formatNumber(latest.downloads));

    selectAll('[data-release-link]').forEach(function(link){ link.href = latest.url || RELEASES_URL; });
    applyDownloadLinks(latest);
    renderAssets(latest);
    updateStructuredData(data);

    var status = stale
      ? 'Showing the last successful GitHub update; a fresh request is temporarily unavailable.'
      : 'Live release data, download totals, and stars are synced from GitHub.';
    setText('[data-github-status]',status);
    setState(stale ? 'stale' : 'ready');
    doc.dispatchEvent(new CustomEvent('anchorcast:github-ready',{detail:{data:data,stale:stale}}));
  }
  function applyError(){
    setText('[data-github="version"]','Latest');
    setText('[data-github="version-number"]','Latest');
    setText('[data-github="downloads"]','Unavailable');
    setText('[data-github="stars"]','Unavailable');
    setText('[data-github="forks"]','Unavailable');
    setText('[data-github="release-date"]','Open GitHub for date');
    setText('[data-github="release-name"]','Latest AnchorCast release');
    setText('[data-github="release-summary"]','GitHub data is temporarily unavailable. The release-page links still work.');
    setText('[data-github-status]','GitHub data is temporarily unavailable. No number is being guessed or hardcoded.');
    selectAll('[data-download]').forEach(function(link){
      link.href = RELEASES_URL;
      link.setAttribute('aria-busy','false');
      link.classList.add('is-unavailable');
      link.title = 'Open the latest GitHub release to choose a download.';
    });
    setState('error');
  }
  function readCache(){
    try{
      var raw = localStorage.getItem(CACHE_KEY);
      if(!raw) return null;
      var parsed = JSON.parse(raw);
      if(!parsed || !parsed.savedAt || !parsed.data) return null;
      if(Date.now() - parsed.savedAt > CACHE_MAX_AGE) return null;
      return parsed.data;
    }catch(_error){ return null; }
  }
  function writeCache(data){
    try{ localStorage.setItem(CACHE_KEY,JSON.stringify({savedAt:Date.now(),data:data})); }catch(_error){}
  }
  function fetchData(){
    var controller = 'AbortController' in window ? new AbortController() : null;
    var timer = window.setTimeout(function(){ if(controller) controller.abort(); },10000);
    return fetch(API_ENDPOINT,{
      method:'GET',
      headers:{Accept:'application/json'},
      credentials:'same-origin',
      signal:controller ? controller.signal : undefined
    }).then(function(response){
      if(!response.ok) throw new Error('GitHub proxy returned HTTP ' + response.status);
      return response.json();
    }).then(function(data){
      if(!data || data.ok !== true || !data.latest) throw new Error('GitHub proxy returned invalid data');
      return data;
    }).finally(function(){ window.clearTimeout(timer); });
  }

  selectAll('[data-download]').forEach(function(link){
    link.setAttribute('aria-busy','true');
    if(!link.getAttribute('href')) link.href = RELEASES_URL;
  });

  fetchData().then(function(data){
    writeCache(data);
    applyData(data,false);
  }).catch(function(error){
    console.warn('[AnchorCast] Live GitHub data unavailable:',error);
    var cached = readCache();
    if(cached) applyData(cached,true); else applyError();
  });
})();

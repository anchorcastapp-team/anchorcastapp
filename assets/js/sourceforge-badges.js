/* SourceForge badges for AnchorCast (project 4094952).
 * Keep both official badge roots in the HTML, but request the provider once.
 * No release/download counters are read or changed by this module.
 */
(function () {
  'use strict';
  var doc = document;
  var roots = Array.from(doc.querySelectorAll('.ac-sf-badge .sf-root[data-id="4094952"]'));
  if (!roots.length || window.__anchorcastSourceForgeBadges) return;
  window.__anchorcastSourceForgeBadges = true;

  var providerURL = 'https://b.sf-syn.com/badge_js?sf_id=4094952';
  var fallback = roots.map(function (root) { return root.innerHTML; });
  var observer;
  var deadline;
  var settled = false;

  function labelFor(root) {
    return root.parentElement.getAttribute('data-sf-label') || 'SourceForge badge';
  }
  function hasArtwork(root) {
    return Array.from(root.querySelectorAll('img')).some(function (img) {
      return img.complete && img.naturalWidth > 0;
    }) || !!root.querySelector('svg');
  }
  function noteVisible(show) {
    doc.querySelectorAll('.ac-sf-service-note').forEach(function (note) {
      note.hidden = !show;
    });
  }
  function accessible(root) {
    var label = labelFor(root);
    root.querySelectorAll('a[target="_blank"]').forEach(function (link) {
      var rel = new Set((link.getAttribute('rel') || '').split(/\s+/).filter(Boolean));
      rel.add('noopener'); rel.add('noreferrer');
      var next = Array.from(rel).join(' ');
      if (link.getAttribute('rel') !== next) link.setAttribute('rel', next);
      if (!link.hasAttribute('aria-label')) {
        link.setAttribute('aria-label', 'AnchorCast on SourceForge — ' + label + ' (opens in a new tab)');
      }
    });
    root.querySelectorAll('img').forEach(function (img) {
      if (!img.alt) img.alt = 'SourceForge ' + label + ' badge for AnchorCast';
      if (!img.complete && !img.dataset.acSfObserved) {
        img.dataset.acSfObserved = 'true';
        img.addEventListener('load', refresh, {once:true});
        img.addEventListener('error', refresh, {once:true});
      }
    });
    root.querySelectorAll('svg').forEach(function (svg) {
      if (!svg.hasAttribute('role')) svg.setAttribute('role', 'img');
      if (!svg.hasAttribute('aria-label')) svg.setAttribute('aria-label', 'SourceForge ' + label + ' badge for AnchorCast');
    });
  }
  function refresh() {
    roots.forEach(function (root) {
      accessible(root);
      if (hasArtwork(root)) root.dataset.sfState = 'ready';
    });
    if (roots.every(hasArtwork)) {
      settled = true;
      window.clearTimeout(deadline);
      if (observer) observer.disconnect();
      noteVisible(false);
    }
  }
  function showFallback() {
    if (settled) return;
    settled = true;
    window.clearTimeout(deadline);
    if (observer) observer.disconnect();
    roots.forEach(function (root, index) {
      if (!hasArtwork(root)) {
        root.innerHTML = fallback[index];
        root.dataset.sfState = 'fallback';
      }
      accessible(root);
    });
    noteVisible(roots.some(function (root) { return !hasArtwork(root); }));
  }

  roots.forEach(function (root) { root.dataset.sfState = 'loading'; });
  if ('MutationObserver' in window) {
    observer = new MutationObserver(refresh);
    roots.forEach(function (root) {
      observer.observe(root, {childList:true, subtree:true});
    });
  }
  // A blocked or unavailable badge must not leave blank boxes indefinitely.
  deadline = window.setTimeout(showFallback, 12000);
  var existing = Array.from(doc.scripts).find(function (script) { return script.src === providerURL; });
  if (existing) {
    existing.addEventListener('load', refresh, {once:true});
    existing.addEventListener('error', showFallback, {once:true});
    refresh();
    return;
  }
  var script = doc.createElement('script');
  script.id = 'anchorcast-sourceforge-badge-provider';
  script.src = providerURL;
  script.async = true;
  script.addEventListener('load', refresh, {once:true});
  script.addEventListener('error', showFallback, {once:true});
  doc.head.appendChild(script);
}());

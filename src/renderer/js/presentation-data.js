/* AnchorCast 1.8.1 — presentation records shared by library and editor. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PresentationData = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  function slide(value) {
    const s = value && typeof value === 'object' ? value : {};
    let objects = Array.isArray(s.objects) ? s.objects.map(o => ({...o,
      textColor: o.textColor || o.color || '#ffffff', textAlign: o.textAlign || o.align || 'center'
    })) : [];
    if (!Array.isArray(s.objects)) {
      if (s.title) objects.push({type:'text', text:String(s.title), x:120,y:260,w:1680,h:260,
        font:'Cinzel',fontSize:s.titleSize||60,textColor:s.textColor||'#fff',textAlign:'center',bold:true});
      if (s.body) objects.push({type:'text',text:String(s.body),x:120,y:540,w:1680,h:420,
        font:'Crimson Pro',fontSize:s.bodySize||34,textColor:s.textColor||'#ddd',textAlign:'center'});
    }
    return {...s,bg:s.bg||'#0a0a1e',bgImage:s.bgImage||null,objects};
  }
  function normalize(value, kind) {
    if (!value || typeof value !== 'object' || value.id == null) throw Error('Invalid presentation record.');
    const inferred = value.type === 'imported' || (value.type !== 'created' && (value.outDir || value.resolvedPath))
      ? 'imported' : (value.type === 'created' || Array.isArray(value.slides) ? 'created' : kind || 'created');
    const p = {...value,id:String(value.id),name:String(value.name||'Untitled'),type:inferred};
    if (inferred === 'created') {p.slides=(value.slides||[]).map(slide);p.slideCount=p.slides.length;}
    return p;
  }
  function mediaUrl(value) {
    if (!value) return '';
    let p = String(value);
    if (/^(?:data:image\/|blob:|https?:\/\/|media:\/\/)/i.test(p)) return p;
    if (/^file:\/\//i.test(p)) {try {p=decodeURIComponent(new URL(p).pathname);if(/^\/[A-Za-z]:/.test(p))p=p.slice(1);}catch(_) {return '';}}
    return 'media:///' + p.replace(/\\/g,'/').replace(/^\//,'').split('/').map(encodeURIComponent).join('/');
  }
  function hasTemporaryImages(p) {
    return (p.slides||[]).some(s => /^blob:/i.test(s.bgImage||'') || (s.objects||[]).some(o=>/^blob:/i.test(o.src||'')));
  }
  return {normalize,slide,mediaUrl,hasTemporaryImages,revision:'presentations-persistence-v1.8.1-r1'};
});

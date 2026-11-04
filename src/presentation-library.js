'use strict';
/** Durable library. Kept separate from main.js so local SMTP settings are untouched. */
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const Data=require('./renderer/js/presentation-data');
function inside(root, file) {const r=path.relative(root,file);return !!r && r!=='..' && !r.startsWith('..'+path.sep) && !path.isAbsolute(r);}
function createStore({dataDir,presentationDir}) {
  fs.mkdirSync(dataDir,{recursive:true});fs.mkdirSync(presentationDir,{recursive:true});
  const roots={imported:path.join(dataDir,'presentations.json'),created:path.join(dataDir,'created_presentations.json')};
  const root=fs.realpathSync(presentationDir);
  function read(kind) {
    const file=roots[kind];let bytes;
    try {if(fs.lstatSync(file).isSymbolicLink())throw Error('Presentation library is a symbolic link.');bytes=fs.readFileSync(file,'utf8');}
    catch(e) {if(e.code==='ENOENT')return [];throw e;}
    try {const list=JSON.parse(bytes);if(!Array.isArray(list))throw Error('Expected a list.');return list.map(p=>Data.normalize(p,kind));}
    catch(e) {throw Error('Cannot read '+path.basename(file)+'. The file was not replaced. Restore its backup or contact support. '+e.message);}
  }
  function write(kind,list) {
    if (!Array.isArray(list) || list.length>10000) throw Error('Invalid presentation list.');
    const file=roots[kind],tmp=file+'.'+crypto.randomUUID()+'.tmp';
    // Validate existing data before replacing it: never turn a read error into an empty library.
    if(fs.existsSync(file))read(kind);
    const normalized=list.map(p=>Data.normalize(p,kind));
    const bytes=JSON.stringify(normalized,null,2)+'\n';
    if(Buffer.byteLength(bytes)>128*1024*1024)throw Error('Presentation library is too large to save safely.');
    try {
      const fd=fs.openSync(tmp,'wx',0o600);try{fs.writeFileSync(fd,bytes);fs.fsyncSync(fd);}finally{fs.closeSync(fd);}
      if(fs.existsSync(file))fs.copyFileSync(file,file+'.bak');
      fs.renameSync(tmp,file);
      // Directory fsync is best effort on platforms which support it.
      try{const fd=fs.openSync(dataDir,'r');try{fs.fsyncSync(fd);}finally{fs.closeSync(fd);}}catch(_){}
    } finally {try{fs.unlinkSync(tmp);}catch(_) {}}
    return normalized;
  }
  function directory(p) {
    // Prefer the durable managed folder. Never follow metadata to arbitrary directories.
    const basename=String(p.outDir||'').replace(/\\/g,'/').split('/').pop();
    for(const leaf of [String(p.id),basename].filter(Boolean)) {
      if(leaf==='.'||leaf==='..'||/[\\/\0]/.test(leaf))continue;
      const dir=path.join(path.resolve(presentationDir),leaf);
      try {if(fs.lstatSync(dir).isSymbolicLink())continue;const real=fs.realpathSync(dir);
        if(inside(root,real)&&fs.statSync(real).isDirectory())return dir;}catch(_){}
    }
    return null;
  }
  function files(dir) {
    if(!dir)return [];
    return fs.readdirSync(dir).filter(f=>/^slide-?\d+\.png$/i.test(f)).filter(f=>{
      const t=path.join(dir,f);return !fs.lstatSync(t).isSymbolicLink() && fs.statSync(t).isFile();
    }).sort((a,b)=>Number(a.match(/\d+/)[0])-Number(b.match(/\d+/)[0])).map((f,i)=>({index:i,imagePath:path.join(dir,f)}));
  }
  function imported() {
    const list=read('imported');let repair=false;
    const out=list.map(p=>{const dir=directory(p),slides=files(dir);
      if(p.type!=='imported'||(dir&&p.outDir!==dir))repair=true;
      return {...p,type:'imported',...(dir?{outDir:dir}:{}),slideCount:slides.length||p.slideCount||0,
        assetsMissing:!slides.length,thumbnailPath:slides[0]?.imagePath||null};});
    // Persist repaired paths and explicitly persist the discriminator on legacy rows.
    if(fs.existsSync(roots.imported)) {
      const raw=JSON.parse(fs.readFileSync(roots.imported,'utf8'));
      if(repair||raw.some(p=>p.type!=='imported'))write('imported',out.map(({assetsMissing,thumbnailPath,...p})=>p));
    }
    return out;
  }
  function saveCreated(payload) {
    let list;
    if(payload?.operation==='upsert') {
      const p=Data.normalize(payload.presentation,'created');
      if(p.type!=='created')throw Error('Only editable presentations can be saved by the editor.');
      if(Data.hasTemporaryImages(p))throw Error('A slide still contains a temporary image. Re-select its image/background before saving.');
      list=read('created').filter(x=>String(x.id)!==p.id);list.unshift(p);
      write('created',list);return {success:true,presentation:p};
    }
    if(!Array.isArray(payload))throw Error('Invalid presentation save request.');
    if(payload.some(Data.hasTemporaryImages))throw Error('Temporary images cannot be saved; re-select the affected images.');
    list=write('created',payload);return {success:true,count:list.length};
  }
  function remove({id,type}={}) {
    if(id==null)throw Error('Select a presentation first.');id=String(id);
    const kinds=type==='created'?['created']:type==='imported'?['imported']:['imported','created'];
    let found=false,warning='';
    for(const kind of kinds) {
      const list=read(kind),p=list.find(x=>String(x.id)===id);if(!p)continue;
      // Commit removal FIRST. A cleanup error cannot resurrect the catalog record.
      write(kind,list.filter(x=>String(x.id)!==id));found=true;
      if(kind==='imported') {const dir=directory(p);if(dir)try{fs.rmSync(dir,{recursive:true,force:true});}catch(e){warning='Removed from library; slide files could not be cleaned up: '+e.message;}}
    }
    return {success:true,removed:found,warning};
  }
  function getSlides(id) {
    const p=read('imported').find(x=>String(x.id)===String(id));
    if(!p)return {success:false,slides:[],error:'Presentation is no longer in the library. Refresh the list.'};
    const dir=directory(p),slides=files(dir);
    if(!slides.length)return {success:false,slides:[],error:'The saved slide images are missing. Re-import the original PowerPoint/PDF; the library entry has been kept.'};
    return {success:true,slides,name:p.name,type:'imported',outDir:dir,slideCount:slides.length};
  }
  return {read,write,imported,saveCreated,remove,getSlides,files};
}
function trusted(event,options) {
  const windows=[options.getMainWindow?.(),options.getEditorWindow?.()].filter(w=>w&&!w.isDestroyed?.());
  const win=windows.find(w=>w.webContents===event?.sender);
  if(!win||event.senderFrame!==win.webContents.mainFrame)return false;
  try {
    const u=new URL(event.senderFrame.url);
    const expected=win===options.getEditorWindow?.()?'presentation-editor.html':'index.html';
    if(u.protocol==='file:')return path.normalize(require('node:url').fileURLToPath(u))===path.join(__dirname,'renderer',expected);
    const port=options.getRendererPort?.();
    return u.protocol==='http:'&&u.hostname==='127.0.0.1'&&!!port&&Number(u.port)===Number(port)&&
      (u.pathname==='/'+expected||(expected==='index.html'&&u.pathname==='/'));
  }catch(_){return false;}
}
function install(options) {
  const {ipcMain}=options,store=createStore(options);
  const {createImporter}=require('./presentation-import');
  const importer=createImporter({...options,store});
  const handlers={
    'get-presentations':()=>store.imported(),
    'get-created-presentations':()=>store.read('created'),
    'save-created-presentations':(_,payload)=>store.saveCreated(payload),
    'get-presentation-slides':(_,payload)=>store.getSlides(payload?.id),
    'delete-presentation':(_,payload)=>store.remove(payload),
    'import-presentation':(_,payload)=>importer.importFile(payload?.filePath)
  };
  for(const [channel,fn] of Object.entries(handlers)) {
    ipcMain.removeHandler(channel);
    ipcMain.handle(channel,async(event,payload)=>{
      if(!trusted(event,options))throw Error('Presentation operation denied for this window.');
      try{return await fn(event,payload);}catch(e){
        if(channel.startsWith('get-')&&channel!=='get-presentation-slides')throw e;
        return {success:false,error:e.message};
      }
    });
  }
  return store;
}
module.exports={createStore,install,trusted,inside};

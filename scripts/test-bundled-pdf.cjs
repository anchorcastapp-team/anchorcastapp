'use strict';
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),os=require('node:os'),crypto=require('node:crypto');
const P=require('../src/presentation-pdf');
const {createImporter}=require('../src/presentation-import');
const {createStore}=require('../src/presentation-library');
const {prepare}=require('./prepare-pdf-renderer.cjs');
const buildHook=require('./check-pdf-renderer-build.cjs');
const ROOT=path.resolve(__dirname,'..'),FIX=path.join(__dirname,'fixtures/presentations');
function env(t){const root=fs.mkdtempSync(path.join(os.tmpdir(),'anchorcast-bundled-pdf-'));t.after(()=>fs.rmSync(root,{recursive:true,force:true}));return root;}
function fakePackage(root,version=P.PDFJS_VERSION){
 const dest=path.join(root,'node_modules/pdfjs-dist');fs.mkdirSync(path.join(dest,'legacy/build'),{recursive:true});
 fs.writeFileSync(path.join(root,'package.json'),'{}');fs.writeFileSync(path.join(dest,'package.json'),JSON.stringify({name:'pdfjs-dist',version,license:'Apache-2.0'}));
 for(const f of ['pdf.mjs','pdf.worker.mjs','pdf.renderer.mjs'])fs.writeFileSync(path.join(dest,'legacy/build',f),'// TEST STUB - no PDF rendering');
 for(const f of ['cmaps','standard_fonts','wasm','iccs']){fs.mkdirSync(path.join(dest,f),{recursive:true});fs.writeFileSync(path.join(dest,f,'test.bin'),'test only');}
 fs.writeFileSync(path.join(dest,'LICENSE'),'Test license placeholder');return dest;
}
async function job(t,options={}){const root=env(t),outDir=path.join(root,'slides'),packageRoot=fakePackage(root);const j=await P.startJob({input:path.join(FIX,'restart-test.pdf'),outDir,packageRoot,...options});t.after(()=>j.close(false));return {...j,root,outDir,packageRoot};}
function endpoint(j,route){return j.origin+'/'+j.token+'/'+route;}
function post(j,route,data,json=true){return fetch(endpoint(j,route),{method:'POST',headers:{Origin:j.origin,'X-AnchorCast-Pdf':j.token,'Content-Type':json?'application/json':'image/png'},body:json?JSON.stringify(data):data});}
const png=()=>fs.readFileSync(path.join(FIX,'picture.png'));
for(const platform of ['win32','darwin','linux'])test('Bundled-first import uses no external converter on '+platform,async t=>{
 const root=env(t),opts={dataDir:path.join(root,'data'),presentationDir:path.join(root,'slides')},store=createStore(opts);let external=0,called=0;
 const r=await createImporter({...opts,store,platform,exec:async()=>{external++;throw Error('No external tools installed');},pdfRenderer:async(input,dir)=>{called++;fs.writeFileSync(path.join(dir,'slide-0001.png'),png());return {method:'pdfjs-bundled',pageCount:1};}}).importFile(path.join(FIX,'restart-test.pdf'));
 assert.equal(r.success,true,r.error);assert.equal(r.converter,'pdfjs-bundled');assert.equal(called,1);assert.equal(external,0);assert.equal(createStore(opts).getSlides(r.id).slides.length,1);
 store.remove({id:r.id,type:'imported'});assert.equal(createStore(opts).imported().length,0);
});
for(const code of ['PDF_CANCELLED','PDF_PASSWORD','PDF_INVALID','PDF_LIMIT','PDF_BUSY','PDF_UNSUPPORTED'])test(code+' never falls through to Python',async t=>{
 const root=env(t),opts={dataDir:path.join(root,'data'),presentationDir:path.join(root,'slides')},store=createStore(opts);let count=0;
 const r=await createImporter({...opts,store,exec:async()=>{count++;throw Error('must not run');},pdfRenderer:async()=>{throw Object.assign(Error('test '+code),{code});}}).importFile(path.join(FIX,'restart-test.pdf'));
 assert.equal(r.success,false);assert.equal(r.code,code);assert.equal(r.cancelled,code==='PDF_CANCELLED');assert.equal(count,0);assert.equal(store.read('imported').length,0);assert.equal(fs.readdirSync(opts.presentationDir).length,0);
});
test('Working Poppler remains fallback when bundle is unavailable',async t=>{
 const root=env(t),opts={dataDir:path.join(root,'data'),presentationDir:path.join(root,'slides')},store=createStore(opts);
 const r=await createImporter({...opts,store,pdfRenderer:async()=>{throw Object.assign(Error('missing'),{code:'PDF_RENDERER_MISSING'});},tools:{pdftoppm:['pdftoppm']},exec:async(cmd,args)=>{assert.equal(cmd,'pdftoppm');fs.writeFileSync(args.at(-1)+'-1.png',png());return {};}}).importFile(path.join(FIX,'restart-test.pdf'));
 assert.equal(r.success,true);assert.equal(r.converter,'pdf-poppler');
});
test('Failed bundled conversion removes partial PNG before successful fallback',async t=>{
 const root=env(t),opts={dataDir:path.join(root,'data'),presentationDir:path.join(root,'slides')},store=createStore(opts);
 const r=await createImporter({...opts,store,pdfRenderer:async(_,dir)=>{fs.writeFileSync(path.join(dir,'slide-0999.png'),png());throw Error('renderer failed');},tools:{pdftoppm:['poppler']},exec:async(cmd,args)=>{fs.writeFileSync(args.at(-1)+'-1.png',png());return {};}}).importFile(path.join(FIX,'restart-test.pdf'));
 assert.equal(r.success,true);assert.equal(r.slideCount,1);assert.equal(path.basename(r.slides[0].imagePath),'slide-1.png');
});
test('Missing built-in and external renderers produce PDF-specific setup guidance',async t=>{
 const root=env(t),opts={dataDir:path.join(root,'data'),presentationDir:path.join(root,'slides')},store=createStore(opts);
 const r=await createImporter({...opts,store,pdfRenderer:async()=>{throw Error('Install rebuilt AnchorCast with PDF support');},pythonCandidates:[],tools:{pdftoppm:[]}}).importFile(path.join(FIX,'restart-test.pdf'));
 assert.match(r.error,/PDF import could not finish/);assert.match(r.error,/Install rebuilt/);assert.doesNotMatch(r.error,/For faithful slide layout/);assert.equal(r.success,false);
});
test('Installed package without staged assets is not mistaken for packaged support',t=>{const root=env(t);fakePackage(root);const s=P.getRendererStatus({appPath:root});assert.equal(s.ready,false);assert.match(s.error,/prepare-pdf-renderer/);});
test('Staging copies matching worker, resources and notices; source metadata untouched',t=>{const root=env(t);fakePackage(root);const before=fs.readFileSync(path.join(root,'package.json'));const r=prepare(root);assert(r.count>5);assert(P.getRendererStatus({appPath:root}).ready);assert.equal(buildHook.verifyBundle(root).version,P.PDFJS_VERSION);assert.deepEqual(fs.readFileSync(path.join(root,'package.json')),before);prepare(root);assert(P.getRendererStatus({appPath:root}).ready);});
test('Wrong dependency version cannot replace a working bundle',t=>{const root=env(t),p=fakePackage(root);prepare(root);const before=fs.readFileSync(path.join(root,'assets/pdfjs/package.json'));fs.writeFileSync(path.join(p,'package.json'),JSON.stringify({version:'0.0.1'}));assert.throws(()=>prepare(root),/required/);assert.deepEqual(fs.readFileSync(path.join(root,'assets/pdfjs/package.json')),before);});
test('Missing worker fails staging without destroying previous bundle',t=>{const root=env(t),p=fakePackage(root);prepare(root);fs.unlinkSync(path.join(p,'legacy/build/pdf.worker.mjs'));assert.throws(()=>prepare(root),/Missing renderer/);assert(P.getRendererStatus({appPath:root}).ready);});
test('Build guard refuses missing staged renderer',async t=>{const root=env(t);await assert.rejects(buildHook({appDir:root}),/missing or incomplete/);});
test('Build guard checks full resource hashes and rejects omission rules',async t=>{const root=env(t);fakePackage(root);prepare(root);await buildHook({appDir:root,packager:{config:{files:['src/**/*','assets/**/*']}}});await assert.rejects(buildHook({appDir:root,packager:{config:{files:['src/**/*']}}}),/include/);await assert.rejects(buildHook({appDir:root,packager:{config:{files:['assets/**/*','!assets/pdfjs/**']}}}),/excludes/);fs.appendFileSync(path.join(root,'assets/pdfjs/legacy/build/pdf.worker.mjs'),'changed');assert.throws(()=>buildHook.verifyBundle(root),/differs/);});
test('Asset mapping excludes traversal, unrelated JS and package files',t=>{const root=env(t),p=fakePackage(root);assert(P.safeAsset(p,'legacy/build/pdf.mjs'));for(const s of ['../../secret','../package.json','package.json','legacy/build/../../../package.json','cmaps/../../file','cmaps/x%2fy','/etc/passwd'])assert.equal(P.safeAsset(p,s),null,s);});
test('Png validation accepts fixture and rejects oversize dimensions and invalid bytes',()=>{assert(P.validatePng(png()).width>0);assert.throws(()=>P.validatePng(Buffer.from('fake')),/invalid/);const bad=Buffer.from(png());bad.writeUInt32BE(99999,16);assert.throws(()=>P.validatePng(bad),/limits/);});
test('PDF errors distinguish encrypted, invalid and unsupported documents',()=>{assert.equal(P.sanitizeError('PasswordException','secret').code,'PDF_PASSWORD');assert.doesNotMatch(P.sanitizeError('PasswordException','secret').message,/secret/);assert.equal(P.sanitizeError('InvalidPDFException','bad').code,'PDF_INVALID');assert.equal(P.sanitizeError('PDF_UNSUPPORTED','not supported').code,'PDF_UNSUPPORTED');});
test('Local server serves only selected PDF and allowlisted renderer resources',async t=>{const j=await job(t);assert.equal((await fetch(endpoint(j,'index.html'))).status,200);const pdf=await fetch(endpoint(j,'document.pdf'));assert.equal(pdf.headers.get('content-type'),'application/pdf');assert.deepEqual(Buffer.from(await pdf.arrayBuffer()),fs.readFileSync(path.join(FIX,'restart-test.pdf')));assert.equal((await fetch(j.origin+'/wrong/index.html')).status,404);assert.equal((await fetch(endpoint(j,'pdfjs/package.json'))).status,404);assert.equal((await fetch(endpoint(j,'pdfjs/legacy/build/pdf.mjs'))).status,200);});
test('Local service denies unauthenticated writes and still accepts valid requests',async t=>{const j=await job(t);assert.equal((await fetch(endpoint(j,'meta'),{method:'POST',body:'{"pages":1}'})).status,403);assert.equal((await post(j,'meta',{pages:1})).status,200);});
test('Local service uses restrictive CSP and no cache',async t=>{const j=await job(t),r=await fetch(endpoint(j,'index.html'));assert.match(r.headers.get('content-security-policy'),/object-src 'none'/);assert.match(r.headers.get('content-security-policy'),/connect-src 'self'/);assert.doesNotMatch(r.headers.get('content-security-policy'),/unsafe-inline/);assert.equal(r.headers.get('cache-control'),'no-store');});
test('Two page completion writes durable ordered PNG files',async t=>{const j=await job(t);assert.equal((await post(j,'meta',{pages:2})).status,200);for(let i=1;i<=2;i++)assert.equal((await post(j,'page/'+i,png(),false)).status,200);await post(j,'done',{});const r=await j.result;assert.equal(r.pageCount,2);await j.close(true);assert.deepEqual(fs.readdirSync(j.outDir),['slide-0001.png','slide-0002.png']);});
test('No completion accepted until every page has been saved',async t=>{const j=await job(t);await post(j,'meta',{pages:2});await post(j,'page/1',png(),false);await post(j,'done',{});await assert.rejects(j.result,/every page/);await j.close(false);assert.deepEqual(fs.readdirSync(j.outDir),[]);});
test('Page order and total page count are bounded',async t=>{const j=await job(t);await post(j,'meta',{pages:1});await post(j,'page/2',png(),false);await assert.rejects(j.result,/out of order/);});
test('Reject too many pages before any page write',async t=>{const j=await job(t);await post(j,'meta',{pages:501});await assert.rejects(j.result,/500 pages/);assert.deepEqual(fs.readdirSync(j.outDir),[]);});
test('Cancel removes partial derivative and preserves original hash',async t=>{const input=path.join(FIX,'restart-test.pdf'),hash=()=>crypto.createHash('sha256').update(fs.readFileSync(input)).digest('hex'),before=hash(),j=await job(t);await post(j,'meta',{pages:2});await post(j,'page/1',png(),false);await post(j,'cancel',{});await assert.rejects(j.result,e=>e.code==='PDF_CANCELLED');await j.close(false);assert.equal(hash(),before);assert.deepEqual(fs.readdirSync(j.outDir),[]);});
test('PDF job refuses an existing page folder instead of overwriting it',async t=>{const root=env(t),dir=path.join(root,'slides');fs.mkdirSync(dir);fs.writeFileSync(path.join(dir,'slide-0001.png'),png());await assert.rejects(P.startJob({input:path.join(FIX,'restart-test.pdf'),outDir:dir,packageRoot:fakePackage(root)}),/not empty/);assert.equal(fs.readdirSync(dir).length,1);});
test('Invalid PDF header rejected before creating output',async t=>{const root=env(t),input=path.join(root,'wrong.pdf');fs.writeFileSync(input,'not a PDF');await assert.rejects(P.startJob({input,outDir:path.join(root,'slides'),packageRoot:root}),/PDF header/);assert.equal(fs.existsSync(path.join(root,'slides')),false);});
test('Stalled page times out and releases local service',async t=>{const j=await job(t,{limits:{...P.LIMITS,pageTimeout:60,totalTimeout:500}});await assert.rejects(j.result,e=>e.code==='PDF_TIMEOUT');await j.close(false);await assert.rejects(fetch(endpoint(j,'index.html')));});
test('Renderer failure becomes an actionable non-Python PDF error',async t=>{const j=await job(t);await post(j,'error',{name:'PasswordException',message:'password needed'});await assert.rejects(j.result,e=>e.code==='PDF_PASSWORD');});
test('Library metadata cannot commit an incomplete built-in page set',async t=>{const root=env(t),opts={dataDir:path.join(root,'data'),presentationDir:path.join(root,'slides')},store=createStore(opts);const r=await createImporter({...opts,store,pdfRenderer:async(_,dir)=>{fs.writeFileSync(path.join(dir,'slide-0001.png'),png());return {pageCount:2};},tools:{pdftoppm:[]},pythonCandidates:[]}).importFile(path.join(FIX,'restart-test.pdf'));assert.equal(r.success,false);assert.equal(store.read('imported').length,0);});
test('Main renderer preferences retain sandbox and disable document privileges',()=>{const s=fs.readFileSync(path.join(ROOT,'src/presentation-pdf.js'),'utf8'),r=fs.readFileSync(path.join(ROOT,'src/pdf-renderer/render.mjs'),'utf8');for(const check of ['nodeIntegration:false','contextIsolation:true','sandbox:true','webSecurity:true','plugins:false'])assert(s.includes(check),check);assert.match(r,/isEvalSupported:false/);assert.match(r,/enableScripting:false/);assert.doesNotMatch(s,/shell\.openExternal|execFile\(/);});
test('Renderer resource policy allows own assets and generated images, not network or file URLs',()=>{
 const origin='http://127.0.0.1:4567',token='abc';
 for(const u of [origin+'/abc/index.html',origin+'/abc/pdfjs/wasm/a.wasm','blob:'+origin+'/id','data:image/png;base64,AA'])assert(P.allowedResourceURL(u,origin,token),u);
 for(const u of ['https://example.com/a',origin+'/other/file','file:///etc/passwd','data:text/javascript,alert(1)','blob:https://example.com/id','http://127.0.0.1:45678/abc/'])assert.equal(P.allowedResourceURL(u,origin,token),false,u);
});
function fakeElectron(mode='success') {
 const {EventEmitter}=require('node:events'),instances=[];
 class Window extends EventEmitter{
  constructor(options){super();this.options=options;this.destroyed=false;instances.push(this);const wc=new EventEmitter(),ses=new EventEmitter();
   ses.setPermissionRequestHandler=f=>ses.request=f;ses.setPermissionCheckHandler=f=>ses.check=f;ses.webRequest={onBeforeRequest:f=>ses.filter=f};
   wc.session=ses;wc.setWindowOpenHandler=f=>wc.open=f;this.webContents=wc;}
  isDestroyed(){return this.destroyed;}
  async loadURL(url){this.url=url;await fetch(url);}
  show(){setImmediate(async()=>{try{
   if(mode==='cancel'){this.destroy();return;}
   if(mode==='crash'){this.webContents.emit('render-process-gone',{}, {reason:'crashed'});return;}
   const u=new URL(this.url),j={origin:u.origin,token:u.pathname.split('/')[1]};
   await post(j,'meta',{pages:2});for(let i=1;i<=2;i++)await post(j,'page/'+i,png(),false);await post(j,'done',{});
  }catch(e){this.webContents.emit('render-process-gone',{}, {reason:'test-error'});}});}
  destroy(){if(this.destroyed)return;this.destroyed=true;this.emit('closed');}
 }
 return {app:{isReady:()=>true},BrowserWindow:Window,instances};
}
test('Electron bridge lifecycle is bounded and uses isolated no-privilege window (mock Electron)',async t=>{const root=env(t);fakePackage(root);prepare(root);const el=fakeElectron();const r=await P.renderPdf(path.join(FIX,'restart-test.pdf'),path.join(root,'output'),{appPath:root,electron:el});assert.equal(r.pageCount,2);const w=el.instances[0];assert(w.isDestroyed());assert.equal(w.options.webPreferences.nodeIntegration,false);assert.equal(w.options.webPreferences.sandbox,true);assert.equal(w.webContents.session.check(),false);assert.equal(w.webContents.open().action,'deny');});
for(const mode of ['cancel','crash'])test('Electron '+mode+' releases service/window and next import works (mock Electron)',async t=>{const root=env(t);fakePackage(root);prepare(root);const el=fakeElectron(mode);await assert.rejects(P.renderPdf(path.join(FIX,'restart-test.pdf'),path.join(root,'output'),{appPath:root,electron:el}),e=>e.code===(mode==='cancel'?'PDF_CANCELLED':'PDF_RENDER_FAILED'));assert(el.instances[0].isDestroyed());const success=await P.renderPdf(path.join(FIX,'restart-test.pdf'),path.join(root,'next'),{appPath:root,electron:fakeElectron()});assert.equal(success.pageCount,2);});
function importUi(result,filePath){
 const vm=require('node:vm'),text=fs.readFileSync(path.join(ROOT,'src/renderer/js/app.js'),'utf8'),start=text.indexOf('async function importPresentation('),end=text.indexOf('\nfunction _showPresentationImportError(',start),errors=[],messages=[];
 const ctx={window:{electronAPI:{pickPresFile:async()=>({filePath}),importPresentation:async()=>result}},document:{getElementById:()=>null},toast:m=>messages.push(m),_showPresentationImportError:(...args)=>errors.push(args)};
 vm.runInNewContext(text.slice(start,end)+';globalThis.run=importPresentation;',ctx);return {run:ctx.run,errors,messages};
}
test('Cancelling a PDF is not shown as an import failure in main UI',async()=>{const ui=importUi({success:false,cancelled:true},'C:\\Sample.pdf');await ui.run();assert.equal(ui.errors.length,0);assert.match(ui.messages[0],/cancelled/);});
test('PDF errors use PDF-specific title; PPTX errors retain presentation title',async()=>{for(const [ext,title] of [['pdf','PDF import failed'],['pptx','Presentation import failed']]){const ui=importUi({success:false,error:'test'},'/test.'+ext);await ui.run();assert.equal(ui.errors[0][1],title);}});

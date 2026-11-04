'use strict';
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),os=require('node:os');
const {spawnSync}=require('node:child_process');
const vm=require('node:vm');
const ROOT=path.resolve(__dirname,'..'),{createStore,trusted}=require('../src/presentation-library'),Data=require('../src/renderer/js/presentation-data');
const {createImporter,pythonCandidates,toolCandidates,parseOutput}=require('../src/presentation-import');
const fixture=path.join(__dirname,'fixtures','presentations');
function setup(t){const root=fs.mkdtempSync(path.join(os.tmpdir(),'ac-presentations-'));t.after(()=>fs.rmSync(root,{recursive:true,force:true}));const options={dataDir:path.join(root,'Data'),presentationDir:path.join(root,'Presentation')};return {root,options,store:createStore(options)};}
function imported(f,id='123',filenames=['slide-01.png','slide-02.png']) {
 const dir=path.join(f.options.presentationDir,String(id));fs.mkdirSync(dir,{recursive:true});
 for(const file of filenames)fs.copyFileSync(path.join(fixture,'picture.png'),path.join(dir,file));
 const row={id,name:'Imported Test',outDir:dir,slideCount:filenames.length};f.store.write('imported',[row]);return row;
}
const deck={id:77,name:'Created Test',type:'created',slides:[{title:'Saved title',body:'Saved body',bg:'#123456'}]};
test('Legacy imported records without type reopen as imported and persist discriminator',t=>{const f=setup(t);const p=imported(f);const file=path.join(f.options.dataDir,'presentations.json');fs.writeFileSync(file,JSON.stringify([p]));const fresh=createStore(f.options);assert.equal(fresh.imported()[0].type,'imported');assert.equal(JSON.parse(fs.readFileSync(file))[0].type,'imported');assert.equal(fresh.getSlides(123).slides.length,2);});
test('Slide filenames retain real padded numbering and numeric sort',t=>{const f=setup(t);imported(f,'id',['slide-10.png','Slide2.PNG','slide-1.png']);assert.deepEqual(f.store.getSlides('id').slides.map(s=>path.basename(s.imagePath)),['slide-1.png','Slide2.PNG','slide-10.png']);assert.equal(path.basename(f.store.imported()[0].thumbnailPath),'slide-1.png');});
test('Imported delete awaits durable catalog removal and survives independent reload',t=>{const f=setup(t);imported(f);assert.equal(f.store.remove({id:123,type:'imported'}).success,true);assert.equal(createStore(f.options).imported().length,0);assert.equal(fs.existsSync(path.join(f.options.presentationDir,'123')),false);});
test('Created deck saves, migrates legacy text and deletes across independent store instances',t=>{const f=setup(t);f.store.saveCreated({operation:'upsert',presentation:deck});const fresh=createStore(f.options);const p=fresh.read('created')[0];assert.equal(p.id,'77');assert.equal(p.slides[0].objects[0].text,'Saved title');fresh.remove({id:77,type:'created'});assert.deepEqual(createStore(f.options).read('created'),[]);});
test('Numeric and string IDs update one record instead of creating duplicate decks',t=>{const f=setup(t);f.store.saveCreated({operation:'upsert',presentation:deck});f.store.saveCreated({operation:'upsert',presentation:{...deck,id:'77',name:'Updated'}});assert.equal(f.store.read('created').length,1);assert.equal(f.store.read('created')[0].name,'Updated');});
test('Independent upserts preserve other decks instead of overwriting a cached list',t=>{const f=setup(t),second=createStore(f.options);f.store.saveCreated({operation:'upsert',presentation:deck});second.saveCreated({operation:'upsert',presentation:{...deck,id:88}});assert.equal(f.store.read('created').length,2);f.store.remove({id:77,type:'created'});second.saveCreated({operation:'upsert',presentation:{...deck,id:88,name:'later'}});assert.deepEqual(f.store.read('created').map(p=>p.id),['88']);});
test('Corrupted library is preserved and never silently replaced with empty data',t=>{const f=setup(t),file=path.join(f.options.dataDir,'created_presentations.json');fs.writeFileSync(file,'{broken');assert.throws(()=>f.store.saveCreated({operation:'upsert',presentation:deck}),/not replaced/);assert.equal(fs.readFileSync(file,'utf8'),'{broken');});
test('Atomic replacement failure leaves the original file and removes temporary output',t=>{const f=setup(t);f.store.saveCreated({operation:'upsert',presentation:deck});const file=path.join(f.options.dataDir,'created_presentations.json'),before=fs.readFileSync(file);const rename=fs.renameSync;try{fs.renameSync=()=>{throw Error('simulated write failure');};assert.throws(()=>f.store.saveCreated({operation:'upsert',presentation:{...deck,name:'not saved'}}),/simulated/);}finally{fs.renameSync=rename;}assert.deepEqual(fs.readFileSync(file),before);assert(!fs.readdirSync(f.options.dataDir).some(s=>s.endsWith('.tmp')));});
test('Missing images remain visible as a repairable record rather than disappearing',t=>{const f=setup(t);f.store.write('imported',[{id:'missing',name:'Missing',type:'imported',outDir:'/old/missing'}]);assert.equal(f.store.imported().length,1);assert.equal(f.store.imported()[0].assetsMissing,true);assert.match(f.store.getSlides('missing').error,/Re-import/);});
test('Relocated outDir is repaired to the existing managed ID directory',t=>{const f=setup(t);const p=imported(f);f.store.write('imported',[{...p,outDir:'C:\\old\\Presentation\\123'}]);assert.equal(fs.realpathSync(f.store.imported()[0].outDir),fs.realpathSync(path.join(f.options.presentationDir,'123')));assert.equal(f.store.getSlides('123').slides.length,2);});
test('Deletion never removes an arbitrary directory referenced by metadata',t=>{const f=setup(t),other=path.join(f.root,'Keep');fs.mkdirSync(other);fs.writeFileSync(path.join(other,'safe'),'keep');f.store.write('imported',[{id:'escape',outDir:other}]);f.store.remove({id:'escape',type:'imported'});assert.equal(fs.readFileSync(path.join(other,'safe'),'utf8'),'keep');});
test('Managed-directory symlink is not followed for preview or deletion',t=>{const f=setup(t),other=path.join(f.root,'Other');fs.mkdirSync(other);fs.copyFileSync(path.join(fixture,'picture.png'),path.join(other,'slide-1.png'));try{fs.symlinkSync(other,path.join(f.options.presentationDir,'link'),'dir');}catch(e){t.skip('Symlink creation unavailable');return;}f.store.write('imported',[{id:'link',outDir:other}]);assert.equal(f.store.getSlides('link').success,false);f.store.remove({id:'link',type:'imported'});assert(fs.existsSync(path.join(other,'slide-1.png')));});
test('Temporary blob URLs are refused on save without deleting previous content',t=>{const f=setup(t);f.store.saveCreated({operation:'upsert',presentation:deck});assert.throws(()=>f.store.saveCreated({operation:'upsert',presentation:{...deck,slides:[{bgImage:'blob:expired',objects:[]}]}}),/temporary/);assert.equal(f.store.read('created')[0].slides[0].objects[0].text,'Saved title');});
test('Durable local image references and data images survive JSON roundtrip',t=>{const f=setup(t),s={bgImage:'/saved/image.png',objects:[{type:'image',src:'data:image/png;base64,YQ=='}]};f.store.saveCreated({operation:'upsert',presentation:{...deck,slides:[s]}});assert.equal(createStore(f.options).read('created')[0].slides[0].bgImage,s.bgImage);});
test('Generated notes colour/alignment aliases and legacy titles normalize without losing text',()=>{const s=Data.slide({objects:[{type:'text',text:'Title',color:'#c9a84c',align:'left'}]});assert.equal(s.objects[0].textColor,'#c9a84c');assert.equal(s.objects[0].textAlign,'left');assert.equal(Data.slide({title:'Legacy'}).objects[0].text,'Legacy');});
test('Media URLs preserve web/data URLs and encode spaces/hash characters in managed paths',()=>{assert.equal(Data.mediaUrl('data:image/png;base64,AA=='),'data:image/png;base64,AA==');assert.equal(Data.mediaUrl('/a path/pic#1.png'),'media:///a%20path/pic%231.png');assert.equal(Data.mediaUrl('C:\\photos\\x.png'),'media:///C%3A/photos/x.png');});
test('Python discovery includes bundled Mac/Windows runtimes and py uses -3',()=>{
  const mac=pythonCandidates({platform:'darwin',appPath:'/My App',resourcesPath:'/Resources'});
  assert(mac.some(p=>p.command==='/Resources/python/bin/python3'));
  assert(mac.some(p=>p.command==='/opt/homebrew/bin/python3'));
  const win=pythonCandidates({platform:'win32',appPath:'C:/App',resourcesPath:'C:/Resources'});
  assert(win.some(p=>p.command==='C:\\Resources\\python\\python.exe'));
  assert(win.some(p=>p.command==='C:\\App\\python\\python.exe'));
  assert.deepEqual(win.find(p=>p.command==='py').args,['-3']);
});

// Exercise the production module with Windows/POSIX path APIs on every host.
// This simulates string/path semantics, not running a foreign executable.
function discoveryOnHost(host) {
  const filename=path.join(ROOT,'src/presentation-import.js');
  const mod={exports:{}};
  const api=host==='win32'?path.win32:path.posix;
  const home=host==='win32'?'C:\\Users\\Operator':'/Users/Operator';
  vm.runInNewContext(fs.readFileSync(filename,'utf8'),{
    module:mod,exports:mod.exports,
    require:id=>id==='node:path'?api:id==='node:os'?{homedir:()=>home}:require(id),
    process:{platform:host,resourcesPath:undefined},__dirname:path.dirname(filename),
  },{filename});
  return mod.exports;
}
for(const host of ['win32','darwin','linux']) {
  for(const target of ['win32','darwin','linux']) {
    test('Discovery target '+target+' is independent of '+host+' host path rules',()=>{
      const discovery=discoveryOnHost(host),windows=target==='win32';
      const root=windows?'C:\\Anchor Cast\\resources':'/Anchor Cast/resources';
      const app=windows?'C:\\Anchor Cast\\resources\\app.asar':'/Anchor Cast/resources/app.asar';
      const expected=windows?'C:\\Anchor Cast\\resources\\python\\python.exe':'/Anchor Cast/resources/python/bin/python3';
      const rows=discovery.pythonCandidates({platform:target,resourcesPath:root,appPath:app});
      assert.equal(rows[0].command,expected);
      assert(rows.some(p=>p.command===expected));
      assert.equal(new Set(rows.map(p=>p.command)).size,rows.length);
      if(windows)assert.deepEqual(Array.from(rows.find(p=>p.command==='py').args),['-3']);
      else assert(rows.every(p=>!p.command.includes('\\')));
    });
  }
  test('Discovery defaults to '+host+' host and preserves custom interpreter precedence',()=>{
    const discovery=discoveryOnHost(host),windows=host==='win32';
    const root=windows?'D:\\App Resources':'/App Resources';
    const override=windows?'D:\\Custom Python\\python.exe':'/Custom Python/bin/python3';
    const expected=windows?'D:\\App Resources\\python\\python.exe':'/App Resources/python/bin/python3';
    const extra=[override,override,'',null,42];
    const rows=discovery.pythonCandidates({resourcesPath:root,appPath:root,extra});
    assert.equal(rows[0].command,override);
    assert.equal(rows.filter(p=>p.command===override).length,1);
    assert(rows.some(p=>p.command===expected));
    assert.deepEqual(extra,[override,override,'',null,42]);
  });
  test('Mac per-user LibreOffice path stays POSIX on '+host+' host',()=>{
    const discovery=discoveryOnHost(host);
    assert(discovery.toolCandidates('soffice','darwin','/Users/Operator Name').includes('/Users/Operator Name/Applications/LibreOffice.app/Contents/MacOS/soffice'));
    assert(discovery.toolCandidates('pdftoppm','darwin').includes('/opt/homebrew/bin/pdftoppm'));
  });
}
test('Windows discovery preserves UNC roots and recognises executable Python launchers',()=>{
  const root='\\\\server\\share name\\resources';
  const extra=['py','py.exe','C:\\Windows\\py.exe','C:\\Tools\\PY.EXE','python.exe'];
  const rows=pythonCandidates({platform:'win32',resourcesPath:root,appPath:root,extra});
  assert(rows.some(p=>p.command==='\\\\server\\share name\\resources\\python\\python.exe'));
  for(const command of extra.slice(0,4))assert.deepEqual(rows.find(p=>p.command===command).args,['-3']);
  assert.deepEqual(rows.find(p=>p.command==='python.exe').args,[]);
});
test('Mac converter discovery includes application-bundle LibreOffice and Homebrew Poppler',()=>{assert(toolCandidates('soffice','darwin').includes('/Applications/LibreOffice.app/Contents/MacOS/soffice'));assert(toolCandidates('pdftoppm','darwin').includes('/usr/local/bin/pdftoppm'));});
test('Converter diagnostic JSON preserves package errors instead of calling Python absent',()=>{assert.equal(parseOutput('{"success":false,"error":"Missing Pillow"}').error,'Missing Pillow');assert.equal(parseOutput('status\n{"success":true}').success,true);assert.throws(()=>parseOutput('traceback'),/invalid JSON/);});
function importerOptions(f,run){return {...f.options,store:f.store,platform:'darwin',exec:run,tools:{soffice:[],pdftoppm:[]},pythonCandidates:[{command:'bundled-python',args:[]},{command:'system-python',args:[]}]};}
test('Installed Python without packages tries next interpreter and saves typed catalog',async t=>{const f=setup(t),calls=[];const run=async(cmd,args)=>{calls.push([cmd,args]);if(args.includes('-c'))return {stdout:'3\n'};if(cmd==='bundled-python')return {stdout:JSON.stringify({success:false,error:'Python is installed; Pillow missing'})};fs.copyFileSync(path.join(fixture,'picture.png'),path.join(args.at(-1),'slide-1.png'));assert(!args[0].includes('app.asar'));return {stdout:JSON.stringify({success:true,method:'python-basic',warning:'Basic rendering'})};};const r=await createImporter(importerOptions(f,run)).importFile(path.join(fixture,'restart-test.pptx'));assert.equal(r.success,true);assert.equal(r.warning,'Basic rendering');assert.equal(createStore(f.options).imported()[0].type,'imported');assert(calls.some(([c])=>c==='system-python'));});
test('Conversion failure keeps missing-module details and removes only its staging directory',async t=>{const f=setup(t);const run=async(_,args)=>args.includes('-c')?{stdout:'3'}:{stdout:'{"success":false,"error":"Python is installed but python-pptx is missing"}'};const r=await createImporter(importerOptions(f,run)).importFile(path.join(fixture,'restart-test.pptx'));assert.equal(r.success,false);assert.match(r.error,/python-pptx is missing/);assert.doesNotMatch(r.error,/No usable Python/);assert.equal(fs.readdirSync(f.options.presentationDir).length,0);assert.equal(f.store.read('imported').length,0);});
test('Unavailable Python produces a distinct interpreter diagnostic',async t=>{const f=setup(t);const run=async()=>{throw Object.assign(Error('absent'),{code:'ENOENT'});};const r=await createImporter(importerOptions(f,run)).importFile(path.join(fixture,'restart-test.pptx'));assert.match(r.error,/No usable Python 3/);});
test('Import rereads latest catalog before commit and cannot resurrect a deleted deck',async t=>{const f=setup(t);imported(f,'old');const run=async(_,args)=>{if(args.includes('-c'))return {stdout:'3'};f.store.remove({id:'old',type:'imported'});fs.copyFileSync(path.join(fixture,'picture.png'),path.join(args.at(-1),'slide-1.png'));return {stdout:'{"success":true}'};};const r=await createImporter(importerOptions(f,run)).importFile(path.join(fixture,'restart-test.pptx'));assert(r.success);assert.deepEqual(f.store.read('imported').map(p=>p.id),[r.id]);});
test('Partial failed converter images cannot leak into a later successful fallback',async t=>{const f=setup(t);const run=async(cmd,args)=>{if(args.includes('-c'))return {stdout:'3'};fs.copyFileSync(path.join(fixture,'picture.png'),path.join(args.at(-1),cmd==='bundled-python'?'slide-99.png':'slide-1.png'));return {stdout:JSON.stringify({success:cmd!=='bundled-python',error:'invalid slide'})};};const r=await createImporter(importerOptions(f,run)).importFile(path.join(fixture,'restart-test.pptx'));assert(r.success);assert.equal(r.slideCount,1);});
test('Unsupported formats are rejected before creating asset folders',async t=>{const f=setup(t),r=await createImporter({...f.options,store:f.store}).importFile('/test.key');assert.equal(r.success,false);assert.match(r.error,/Export/);assert.deepEqual(fs.readdirSync(f.options.presentationDir),[]);});
test('Native IPC trusts only the actual main/editor main frame on the renderer origin',()=>{const frame={url:'http://127.0.0.1:3000/index.html'},wc={mainFrame:frame},win={webContents:wc,isDestroyed:()=>false};const opts={getMainWindow:()=>win,getEditorWindow:()=>null,getRendererPort:()=>3000};assert(trusted({sender:wc,senderFrame:frame},opts));assert(!trusted({sender:{},senderFrame:frame},opts));assert(!trusted({sender:wc,senderFrame:{...frame}},opts));frame.url='http://127.0.0.1:9999/index.html';assert(!trusted({sender:wc,senderFrame:frame},opts));frame.url='https://evil.test/index.html';assert(!trusted({sender:wc,senderFrame:frame},opts));});
const available=cmd=>spawnSync(cmd,[cmd==='pdftoppm'?'-v':'--version'],{timeout:8000,encoding:'utf8'}).status===0;
test('Real PDF import -> new process/store reload -> deletion uses durable PNGs',{skip:!available('pdftoppm')},async t=>{const f=setup(t);const r=await createImporter({...f.options,store:f.store}).importFile(path.join(fixture,'restart-test.pdf'));assert.equal(r.success,true,r.error);assert.equal(r.slideCount,2);assert.equal(createStore(f.options).getSlides(r.id).slides.length,2);const cli=spawnSync(process.execPath,['-e',`const {createStore}=require(${JSON.stringify(path.join(ROOT,'src/presentation-library.js'))});process.stdout.write(String(createStore(${JSON.stringify(f.options)}).getSlides(${JSON.stringify(r.id)}).slides.length))`],{encoding:'utf8',env:{...process.env,FORCE_COLOR:'1'}});assert.equal(cli.status,0,cli.stderr);assert.equal(cli.stdout,'2');f.store.remove({id:r.id,type:'imported'});assert.equal(createStore(f.options).imported().length,0);});
test('Real PPTX conversion via installed LibreOffice produces two durable slides',{skip:!available('soffice')||!available('pdftoppm'),timeout:120000},async t=>{const f=setup(t);const r=await createImporter({...f.options,store:f.store}).importFile(path.join(fixture,'restart-test.pptx'));assert.equal(r.success,true,r.error);assert.equal(r.slideCount,2);assert.match(r.converter,/libreoffice/);assert.equal(createStore(f.options).getSlides(r.id).slides.length,2);});
test('Basic Python fallback returns a warning and preserves two slide images',{skip:spawnSync('python',['-c','import pptx,PIL']).status!==0},async t=>{const f=setup(t);const r=spawnSync('python',[path.join(ROOT,'src/pptx_to_png.py'),path.join(fixture,'restart-test.pptx'),f.options.presentationDir],{encoding:'utf8',timeout:30000});const out=JSON.parse(r.stdout);assert.equal(out.success,true,out.error);assert.match(out.warning,/Basic/);assert.equal(out.count,2);});
test('New action exists once in library toolbar, not in per-deck rows',()=>{const html=fs.readFileSync(path.join(ROOT,'src/renderer/index.html'),'utf8'),src=fs.readFileSync(path.join(ROOT,'src/renderer/js/app.js'),'utf8');assert.equal((html.match(/id="presNewBtn"/g)||[]).length,1);const rows=src.slice(src.indexOf('function renderPresLibStrip('),src.indexOf('function _editPres('));assert.doesNotMatch(rows,/newPresentation\(\)/);assert.match(rows,/data-pres-action="delete"/);});
test('Editor save awaits native persistence and refuses blob URLs before closing',()=>{const text=fs.readFileSync(path.join(ROOT,'src/renderer/presentation-editor.html'),'utf8');assert.match(text,/await api\.saveCreatedPresentations\(payload\)/);assert.match(text,/hasTemporaryImages/);assert.doesNotMatch(text,/URL\.createObjectURL\(file\)/);assert.match(text,/api\.getPathForFile/);});

test('Canonical aliases keep the configured media-root prefix while checking real containment',t=>{const f=setup(t);imported(f);const alias=path.join(f.root,'Alias');try{fs.symlinkSync(f.options.presentationDir,alias,'dir');}catch(e){t.skip('Symlink creation unavailable');return;}const store=createStore({...f.options,presentationDir:alias});const file=store.getSlides('123').slides[0].imagePath;assert(file.startsWith(alias+path.sep));assert(fs.existsSync(file));});
test('Uppercase PowerPoint extension uses the correct output PDF basename',{skip:!available('soffice')||!available('pdftoppm'),timeout:120000},async t=>{const f=setup(t),src=path.join(f.root,'UPPER Deck.PPTX');fs.copyFileSync(path.join(fixture,'restart-test.pptx'),src);const r=await createImporter({...f.options,store:f.store}).importFile(src);assert.equal(r.success,true,r.error);assert.equal(r.name,'UPPER Deck');assert.equal(r.slideCount,2);assert.match(r.converter,/libreoffice/);});

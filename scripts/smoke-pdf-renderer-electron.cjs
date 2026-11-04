#!/usr/bin/env node
'use strict';
/** Native smoke test: REAL installed Electron + staged PDF.js. No fallback converters and no network. */
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const electron=require('electron');
if(!electron.app){console.error('Run: npx --no-install electron scripts/smoke-pdf-renderer-electron.cjs');process.exitCode=1;}
else{
 const ROOT=path.resolve(__dirname,'..'),temp=fs.mkdtempSync(path.join(os.tmpdir(),'anchorcast-native-pdf-'));
 electron.app.setPath('userData',path.join(temp,'electron-profile'));
 // Keep smoke tests alive between short-lived import windows.
 electron.app.on('window-all-closed',()=>{});
 electron.app.whenReady().then(async()=>{
  try{
   require('./check-pdf-renderer-build.cjs').verifyBundle(ROOT);
   const {createStore}=require('../src/presentation-library'),{createImporter}=require('../src/presentation-import');
   const options={dataDir:path.join(temp,'data'),presentationDir:path.join(temp,'slides'),app:{getAppPath:()=>ROOT},tools:{pdftoppm:[],soffice:[]},pythonCandidates:[],exec:async()=>{throw Error('No external converter is permitted in this smoke test.');}};
   const store=createStore(options),input=path.join(__dirname,'fixtures/presentations/restart-test.pdf');
   const hash=p=>crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex'),before=hash(input);
   const result=await createImporter({...options,store}).importFile(input);
   assert.equal(result.success,true,result.error);assert.equal(result.converter,'pdfjs-bundled');assert.equal(result.slideCount,2);
   for(const slide of result.slides){const image=electron.nativeImage.createFromPath(slide.imagePath);assert.equal(image.isEmpty(),false);const size=image.getSize();assert(size.width>100&&size.height>100);assert(Math.max(size.width,size.height)<=1920);}
   assert.equal(createStore(options).getSlides(result.id).slides.length,2);assert.equal(hash(input),before);
   store.remove({id:result.id,type:'imported'});assert.equal(createStore(options).imported().length,0);
   console.log('PASS Real bundled PDF import: 2 decodable PNGs, saved-library reload, deletion, original unchanged.');
   console.log('PASS No Python, Poppler, PowerPoint or LibreOffice fallback permitted.');
   console.log('Engine: '+require('../src/presentation-pdf').PDFJS_VERSION+'; Electron '+process.versions.electron+'; Chromium '+process.versions.chrome+'; '+process.platform+'/'+process.arch);
   // Chromium may keep its temporary profile files locked until process exit on Windows.
   // Do not turn a passed import into a false test failure by deleting that active profile.
   console.log('Temporary smoke-test profile (safe to remove after exit): '+temp);
   electron.app.exit(0);
  }catch(e){console.error('FAIL Native bundled PDF import:',e.stack||e);console.error('External converters are intentionally disabled in this smoke test. A fallback Python warning here is not an instruction to install Python.');console.error('Temporary diagnostics:',temp);electron.app.exit(1);}
 });
}

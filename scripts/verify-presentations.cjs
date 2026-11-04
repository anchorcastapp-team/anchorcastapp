#!/usr/bin/env node
'use strict';
/** Read-only verification of the manually applied presentations update. */
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto'),{spawnSync}=require('node:child_process');
const root=path.resolve(__dirname,'..');let failures=0;
function check(name,fn){try{fn();console.log('PASS '+name);}catch(e){failures++;console.error('FAIL '+name+' — '+e.message);}}
console.log('AnchorCast 1.8.2 — presentation persistence and bundled PDF renderer');
check('Application version',()=>{const p=JSON.parse(fs.readFileSync(path.join(root,'package.json')));if(p.name!=='anchorcast'||p.version!=='1.8.2')throw Error('This release is AnchorCast v1.8.2. Preserve your dependency graph and update only application-version metadata before verifying.');});
check('Root lockfile application version',()=>{
 const file=path.join(root,'package-lock.json');if(!fs.existsSync(file))return;
 const lock=JSON.parse(fs.readFileSync(file,'utf8'));
 if(lock.version!=='1.8.2'||(lock.packages&&lock.packages['']&&lock.packages[''].version!=='1.8.2'))throw Error('Set only the top-level and packages[""].version application fields to 1.8.2; leave dependency versions unchanged.');
});
let manifest;
check('Patch manifest',()=>{
 const candidate=JSON.parse(fs.readFileSync(path.join(root,'PRESENTATIONS_PATCH_MANIFEST.json')));
 if(candidate.revision!=='presentations-persistence-v1.8.2'||candidate.pdfRendererRevision!=='bundled-pdf-v1.8.2-r2')
  throw Error('Copy PRESENTATIONS_PATCH_MANIFEST.json from the PDF lifecycle fix together with its files. The earlier manifest describes different source versions.');
 if(!candidate.files||typeof candidate.files!=='object'||Array.isArray(candidate.files))throw Error('Invalid source manifest.');
 manifest=candidate;
});
for(const [name,expected] of Object.entries(manifest?.files||{}))check(name,()=>{
 const bytes=fs.readFileSync(path.join(root,name));const value=expected.text?Buffer.from(bytes.toString('utf8').replace(/\r\n/g,'\n')):bytes;
 const hash=crypto.createHash('sha256').update(value).digest('hex');
 const accepted=[expected.sha256,...(expected.alternateSha256||[])];
 if(!accepted.includes(hash))throw Error('File differs from this patch. Preserve intentional local edits and review the diff.');
});
check('Native manual hook installed once',()=>{
 const main=fs.readFileSync(path.join(root,'src/main.js'),'utf8');
 const pattern=/require\(['"]\.\/presentation-library['"]\)\.install\(/g;
 if((main.match(pattern)||[]).length!==1)throw Error('Add PRESENTATION_MAIN_HOOK.txt once INSIDE the main-app block, just before the final "} // end timer mode else block".');
 for(const key of ['getRendererPort','getMainWindow','getEditorWindow','getPythonCandidates'])if(!main.includes(key+':'))throw Error('Incomplete manual hook: '+key);
});
for(const name of ['src/main.js','src/presentation-library.js','src/presentation-import.js','src/renderer/js/presentation-data.js','src/renderer/js/app.js','src/presentation-pdf.js','src/pdf-renderer/render.mjs'])check(name+' syntax',()=>{const r=spawnSync(process.execPath,['--check',path.join(root,name)],{encoding:'utf8',timeout:15000});if(r.error||r.status)throw Error(r.error?.message||r.stderr);});
for(const name of ['preload.js','src/preload.js'])check(name+' existing presentation/image APIs',()=>{const t=fs.readFileSync(path.join(root,name),'utf8');for(const method of ['getPresentations','getCreatedPresentations','saveCreatedPresentations','deletePresentation','getPathForFile','importThemeBgFile'])if(!new RegExp('\\b'+method+'\\s*:').test(t))throw Error('Missing existing API: '+method+'. Do not weaken context isolation; use the complete supported v1.8.1 preload.');});
check('Bundled PDF renderer resources',()=>require('./check-pdf-renderer-build.cjs').verifyBundle(root));
console.log(failures?'Resolve failed checks before running/building.':'Presentation source checks passed. Native Windows/macOS behavior still needs testing on the intended build.');
process.exitCode=failures?1:0;

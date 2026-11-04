#!/usr/bin/env node
'use strict';
const fs=require('node:fs'),path=require('node:path'),{spawnSync}=require('node:child_process');
const root=path.resolve(__dirname,'..');let fail=0;
// Read-only target selection. It does not waive the checks for the selected installer family.
const args=process.argv.slice(2);
const targetArg=args.find(a=>a.startsWith('--platform='));
const target=targetArg?targetArg.slice('--platform='.length):'all';
if(!['all','mac','win'].includes(target)||args.some(a=>a!==targetArg)){
 console.error('Usage: node scripts/verify-pdf-renderer.cjs [--platform=mac|win|all]');
 process.exit(1);
}
const check=(name,fn)=>{try{fn();console.log('PASS '+name);}catch(e){fail++;console.error('FAIL '+name+' — '+e.message);}};
console.log('AnchorCast v1.8.2 — bundled offline PDF renderer');
for(const file of ['src/presentation-import.js','src/presentation-pdf.js','src/pdf-renderer/render.mjs','scripts/prepare-pdf-renderer.cjs','scripts/check-pdf-renderer-build.cjs'])check(file+' syntax',()=>{const r=spawnSync(process.execPath,['--check',path.join(root,file)],{encoding:'utf8'});if(r.status)throw Error(r.stderr);});
check('Bundled PDF.js and all staged resources',()=>require('./check-pdf-renderer-build.cjs').verifyBundle(root));
check('PDF renderer is connected to the presentation importer',()=>{if(!fs.readFileSync(path.join(root,'src/presentation-import.js'),'utf8').includes("require('./presentation-pdf').renderPdf"))throw Error('Copy the updated presentation-import.js.');});
check('Existing presentation hook',()=>{const s=fs.readFileSync(path.join(root,'src/main.js'),'utf8');if((s.match(/require\(['"]\.\/presentation-library['"]\)\.install\(/g)||[]).length!==1)throw Error('The prior presentation-library hook must remain present once.');});
const configs=fs.readdirSync(root).filter(f=>/^electron-builder-(win|mac)-.*\.json$/.test(f)&&
 (target==='all'||f.startsWith('electron-builder-'+target+'-')));
check('Build configuration files ('+target+')',()=>{if(!configs.length)throw Error('No matching electron-builder configuration found.');});
for(const file of configs)check(file+' asset inclusion and build guard',()=>{
 const config=JSON.parse(fs.readFileSync(path.join(root,file),'utf8'));
 if(!config.files?.some(f=>typeof f==='string'&&['assets/**/*','assets/**','assets/pdfjs/**/*'].includes(f)))throw Error('Add assets/pdfjs/**/* to files.');
 if(config.beforePack!=='scripts/check-pdf-renderer-build.cjs')throw Error('Add beforePack: scripts/check-pdf-renderer-build.cjs, or call that hook from your existing beforePack hook.');
});
if(target!=='all')console.log('Other installer platforms were not checked. Use --platform=all before publishing both platforms.');
console.log('Real conversion check: npx electron scripts/smoke-pdf-renderer-electron.cjs');
console.log(fail?'Resolve failures before building.':'Dependency/source checks passed. Run the real Electron conversion check before distributing installers.');process.exitCode=fail?1:0;

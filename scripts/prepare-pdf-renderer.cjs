#!/usr/bin/env node
'use strict';
/** Build-time staging only. Downloads nothing and never modifies user data or native app settings. */
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const {createRequire}=require('node:module');
const {PDFJS_VERSION}=require('../src/presentation-pdf');
function prepare(root=path.resolve(__dirname,'..')){
  const req=createRequire(path.join(root,'package.json'));
  let source;try{source=path.dirname(req.resolve('pdfjs-dist/package.json'));}catch(_){throw Error(`Run npm install --save-dev --save-exact pdfjs-dist@${PDFJS_VERSION} first.`);}
  const pkg=JSON.parse(fs.readFileSync(path.join(source,'package.json'),'utf8'));
  if(pkg.version!==PDFJS_VERSION)throw Error(`PDF.js ${PDFJS_VERSION} is required; installed ${pkg.version}.`);
  const dest=path.join(root,'assets','pdfjs'),stage=dest+'.stage-'+crypto.randomUUID();
  const backup=dest+'.previous-'+crypto.randomUUID();
  let backedUp=false,committed=false;
  function copyDir(name,required=true){
    const from=path.join(source,name);
    if(!fs.existsSync(from)){if(required)throw Error('Missing dependency resources: '+name);return;}
    const target=path.join(stage,name);fs.mkdirSync(target,{recursive:true});
    for(const entry of fs.readdirSync(from,{withFileTypes:true})){
      if(entry.isSymbolicLink())throw Error('Unexpected dependency symlink: '+name+'/'+entry.name);
      if(entry.isFile())fs.copyFileSync(path.join(from,entry.name),path.join(target,entry.name));
    }
  }
  try{
    fs.mkdirSync(path.join(stage,'legacy','build'),{recursive:true});
    for(const entry of fs.readdirSync(path.join(source,'legacy','build'))){
      if(/\.mjs$/.test(entry) && entry!=='pdf.sandbox.mjs')fs.copyFileSync(path.join(source,'legacy','build',entry),path.join(stage,'legacy','build',entry));
    }
    for(const name of ['cmaps','standard_fonts','wasm'])copyDir(name);
    copyDir('iccs',false);
    fs.copyFileSync(path.join(source,'LICENSE'),path.join(stage,'LICENSE'));
    fs.writeFileSync(path.join(stage,'package.json'),JSON.stringify({name:'pdfjs-dist',version:pkg.version,license:pkg.license},null,2)+'\n');
    const files={};
    function hash(dir,rel=''){for(const e of fs.readdirSync(dir,{withFileTypes:true})){const r=rel?rel+'/'+e.name:e.name,p=path.join(dir,e.name);if(e.isDirectory())hash(p,r);else files[r]=crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex');}}
    hash(stage);fs.writeFileSync(path.join(stage,'ANCHORCAST_PDF_BUNDLE.json'),JSON.stringify({version:pkg.version,files},null,2)+'\n');
    for(const name of ['pdf.mjs','pdf.worker.mjs'])if(!fs.existsSync(path.join(stage,'legacy','build',name)))throw Error('Missing renderer '+name);
    if(fs.existsSync(dest)){if(fs.lstatSync(dest).isSymbolicLink())throw Error('assets/pdfjs must not be a symlink.');fs.renameSync(dest,backup);backedUp=true;}
    fs.renameSync(stage,dest);committed=true;
    if(backedUp)fs.rmSync(backup,{recursive:true,force:true});
    console.log(`Staged PDF.js ${pkg.version}: ${Object.keys(files).length} files in assets/pdfjs.`);
    return {version:pkg.version,count:Object.keys(files).length,dest};
  }finally{
    fs.rmSync(stage,{recursive:true,force:true});
    if(!committed&&backedUp&&!fs.existsSync(dest))fs.renameSync(backup,dest);
  }
}
if(require.main===module){try{prepare();}catch(e){console.error('PDF renderer setup failed: '+e.message);process.exitCode=1;}}
module.exports={prepare};

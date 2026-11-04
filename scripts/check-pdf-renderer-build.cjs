'use strict';
/** electron-builder beforePack hook. Fails a release build rather than silently omitting PDF support. */
const path=require('node:path'),fs=require('node:fs'),crypto=require('node:crypto');
const {getRendererStatus,PDFJS_VERSION}=require('../src/presentation-pdf');
function verifyBundle(appPath){
 const status=getRendererStatus({appPath});if(!status.ready)throw Error(status.error+' '+status.detail);
 const manifest=JSON.parse(fs.readFileSync(path.join(status.root,'ANCHORCAST_PDF_BUNDLE.json'),'utf8'));
 if(manifest.version!==PDFJS_VERSION||!manifest.files||Object.keys(manifest.files).length<4)throw Error('Invalid staged PDF renderer manifest. Run node scripts/prepare-pdf-renderer.cjs.');
 for(const [relative,expected] of Object.entries(manifest.files)){
  if(relative.includes('..')||path.isAbsolute(relative)||relative.includes('\\'))throw Error('Invalid bundled PDF path.');
  const file=path.join(status.root,...relative.split('/'));
  if(crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex')!==expected)throw Error('Staged PDF resource differs: '+relative);
 }
 return status;
}
module.exports=async function beforePack(context){
 const root=context.appDir||context.packager?.appDir||path.resolve(__dirname,'..');
 verifyBundle(root);
 const files=context.packager?.config?.files;
 if(files){const patterns=files.flatMap(f=>typeof f==='string'?[f]:(f.filter||[]));
  if(!patterns.some(p=>['**/*','assets/**/*','assets/**','assets/pdfjs/**/*','assets/pdfjs/**'].includes(p)))
    throw Error('Build files must include assets/pdfjs/**/* (assets/**/* also includes it).');
  if(patterns.some(p=>/^!assets(?:\/pdfjs)?\/(?:\*\*|\*\*\/\*)$/.test(p)))throw Error('Build configuration excludes the PDF renderer assets.');
 }
};
module.exports.verifyBundle=verifyBundle;

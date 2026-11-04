// No Node APIs, remote resources, PDF scripting manager, links or interactive document viewer.
// v1.8.2 PDF lifecycle fix: the loading task owns worker destruction, not PDFDocumentProxy.
const base=new URL('./',import.meta.url),token=base.pathname.split('/').filter(Boolean)[0];
const status=document.getElementById('status'),progress=document.getElementById('progress'),cancel=document.getElementById('cancel');
let loadingTask=null,renderTask=null,pdf=null,cancelled=false;
let destroyPromise=null,errorReport=null,cancelReport=null;
async function post(route,body,json=true){
  const response=await fetch(new URL(route,base),{method:'POST',credentials:'omit',headers:{'X-AnchorCast-Pdf':token,'Content-Type':json?'application/json':'image/png'},body:json?JSON.stringify(body):body});
  if(!response.ok){let msg='PDF renderer communication failed.';try{msg=(await response.json()).error||msg;}catch(_){}throw Error(msg);}
}
// Share one promise across success/error/password/cancel paths. Do not destroy a task twice,
// and do not assume the resolved PDFDocumentProxy has a destroy() method (PDF.js 6 does not).
function destroyLoadingTask(){
  if(destroyPromise)return destroyPromise;
  if(!loadingTask)return Promise.resolve();
  const task=loadingTask;
  destroyPromise=Promise.resolve().then(()=>task.destroy()).finally(()=>{loadingTask=null;pdf=null;});
  return destroyPromise;
}
function assertActive(){
  if(cancelled||errorReport){const error=new Error('PDF import interrupted.');error.name='AbortError';throw error;}
}
function reportError(error){
  if(cancelled)return Promise.resolve();
  if(errorReport)return errorReport;
  status.textContent='Could not finish PDF import.';
  cancel.disabled=true;
  errorReport=post('error',{name:String(error.name||'Error').slice(0,80),message:String(error.message||error).slice(0,600)}).catch(()=>{});
  return errorReport;
}
cancel.addEventListener('click',async()=>{
  if(cancelled||errorReport)return;
  cancelled=true;cancel.disabled=true;status.textContent='Cancelling import…';
  try{renderTask?.cancel();}catch(_){}
  // Notify the owner first so it can release the window/service even if worker shutdown stalls.
  cancelReport=post('cancel',{}).catch(()=>{});
  await Promise.all([cancelReport,destroyLoadingTask().catch(()=>{})]);
});
try{
  const configResponse=await fetch(new URL('job.json',base),{credentials:'omit'});
  if(!configResponse.ok)throw Error('The PDF import job could not be loaded.');
  const config=await configResponse.json();
  assertActive();
  const pdfjs=await import('./pdfjs/legacy/build/pdf.mjs');
  assertActive();
  if(pdfjs.version!==config.version)throw Error('PDF library and expected version differ. Reinstall the complete PDF dependency.');
  pdfjs.GlobalWorkerOptions.workerSrc=new URL('./pdfjs/legacy/build/pdf.worker.mjs',base).href;
  const resource=dir=>new URL('./pdfjs/'+dir+'/',base).href;
  loadingTask=pdfjs.getDocument({url:new URL('document.pdf',base).href,withCredentials:false,
    cMapUrl:resource('cmaps'),cMapPacked:true,standardFontDataUrl:resource('standard_fonts'),
    wasmUrl:resource('wasm'),iccUrl:resource('iccs'),
    isEvalSupported:false,enableScripting:false,enableXfa:false,stopAtErrors:true,
    disableRange:true,disableStream:true,disableAutoFetch:true,
    isImageDecoderSupported:false,canvasMaxAreaInBytes:32*1024*1024,
    verbosity:pdfjs.VerbosityLevel.ERRORS});
  loadingTask.onPassword=()=>{
    const error=new Error('Password protected PDF');error.name='PasswordException';
    // Record the real reason once, before aborting the task can reject its promise.
    const report=reportError(error);
    void Promise.all([report,destroyLoadingTask().catch(()=>{})]);
  };
  pdf=await loadingTask.promise;
  assertActive();
  if(pdf.isPureXfa){const error=new Error('This PDF uses XFA forms. Print or export it to an ordinary flattened PDF before importing.');error.name='PDF_UNSUPPORTED';throw error;}
  if(pdf.numPages>config.maxPages){const error=new Error('Import is limited to 500 pages. Split this PDF into smaller documents.');error.name='PDF_LIMIT';throw error;}
  const totalPages=pdf.numPages;
  await post('meta',{pages:totalPages});progress.max=totalPages;progress.value=0;
  for(let i=1;i<=totalPages;i++){
    assertActive();
    status.textContent=`Rendering page ${i} of ${totalPages}…`;
    const page=await pdf.getPage(i),unit=page.getViewport({scale:1});
    assertActive();
    if(!Number.isFinite(unit.width)||!Number.isFinite(unit.height)||unit.width<=0||unit.height<=0)throw Error('Invalid PDF page size.');
    const scale=Math.min(config.maxEdge/Math.max(unit.width,unit.height),Math.sqrt(config.maxPixels/(unit.width*unit.height)));
    const viewport=page.getViewport({scale}),canvas=document.createElement('canvas');
    canvas.width=Math.min(config.maxEdge,Math.max(1,Math.ceil(viewport.width)));canvas.height=Math.min(config.maxEdge,Math.max(1,Math.ceil(viewport.height)));
    try{
      renderTask=page.render({canvasContext:canvas.getContext('2d',{alpha:false}),viewport,background:'rgb(255,255,255)',annotationMode:pdfjs.AnnotationMode.ENABLE});
      await renderTask.promise;renderTask=null;
      assertActive();
      const blob=await new Promise((resolve,reject)=>canvas.toBlob(v=>v?resolve(v):reject(Error('Could not encode the rendered page.')),'image/png'));
      assertActive();
      await post('page/'+i,blob,false);progress.value=i;
    }finally{renderTask=null;canvas.width=0;canvas.height=0;page.cleanup();}
  }
  assertActive();
  status.textContent='Saving slides…';
  // Finish all page writes and release the worker before telling the main process to commit.
  await destroyLoadingTask();
  assertActive();
  cancel.disabled=true;
  await post('done',{});
}catch(error){
  if(!cancelled)await reportError(error);
}finally{
  if(cancelReport)await cancelReport;
  if(errorReport)await errorReport;
  // Never replace a useful rendering/password error with a secondary cleanup exception.
  await destroyLoadingTask().catch(()=>{});
}

'use strict';
/** Local-only PDF.js bridge. No Python, shell commands, remote CDN or privileged preload.
 * PDF parsing runs in an isolated Electron renderer, not the main process.
 */
const fs = require('node:fs');
const fsp = fs.promises;
const path = require('node:path');
const http = require('node:http');
const crypto = require('node:crypto');
const PDFJS_VERSION = '6.5.438';
const LIMITS = Object.freeze({ maxBytes: 128 * 1024 * 1024, maxPages: 500, maxEdge: 1920,
  maxPixels: 1920 * 1920, maxPngBytes: 20 * 1024 * 1024, maxOutputBytes: 1024 * 1024 * 1024,
  pageTimeout: 90000, totalTimeout: 15 * 60 * 1000 });
const PNG = Buffer.from('89504e470d0a1a0a', 'hex');
function failure(code, message) { const error = new Error(message); error.code = code; return error; }
function getRendererStatus({ appPath = path.resolve(__dirname, '..') } = {}) {
  try {
    const root = path.join(appPath, 'assets', 'pdfjs');
    const pkg = JSON.parse(fs.readFileSync(path.join(root,'package.json'), 'utf8'));
    if (pkg.version !== PDFJS_VERSION) throw Error(`Expected pdfjs-dist ${PDFJS_VERSION}; found ${pkg.version}.`);
    for (const relative of ['legacy/build/pdf.mjs', 'legacy/build/pdf.worker.mjs', 'LICENSE']) {
      if (!fs.statSync(path.join(root, relative)).isFile()) throw Error(`Missing ${relative}.`);
    }
    for (const directory of ['cmaps', 'standard_fonts', 'wasm']) {
      if (!fs.statSync(path.join(root, directory)).isDirectory() || !fs.readdirSync(path.join(root, directory)).length)
        throw Error(`Missing PDF resource directory: ${directory}.`);
    }
    return { ready: true, version: pkg.version, root };
  } catch (e) {
    return { ready: false, code: 'PDF_RENDERER_MISSING', error:
      'The built-in PDF renderer is missing or incomplete in this copy of AnchorCast. ' +
      'Install a rebuilt AnchorCast installer that includes PDF support. ' +
      `For a source/build checkout, run npm install --save-dev --save-exact pdfjs-dist@${PDFJS_VERSION}, ` +
      'then node scripts/prepare-pdf-renderer.cjs, verify with node scripts/verify-pdf-renderer.cjs, and rebuild the installer.', detail: e.message };
  }
}
function allowedResourceURL(url, origin, token) {
  return typeof url === 'string' && (url.startsWith(origin + '/' + token + '/') ||
    url.startsWith('blob:' + origin + '/') || /^data:(?:image|font)\//.test(url));
}
function sanitizeError(name, message) {
  if (name === 'PasswordException') return failure('PDF_PASSWORD',
    'This PDF is password protected. Open it in your PDF application and save an authorized unlocked copy, then import that copy.');
  if (name === 'InvalidPDFException' || name === 'FormatError') return failure('PDF_INVALID',
    'The selected PDF could not be read. Open it in a PDF application and export a fresh PDF, then try again.');
  if (name === 'PDF_LIMIT') return failure('PDF_LIMIT', String(message).slice(0,500));
  if (name === 'PDF_UNSUPPORTED') return failure('PDF_UNSUPPORTED', String(message).slice(0,500));
  return failure('PDF_RENDER_FAILED', 'Built-in PDF conversion failed: ' + String(message || name || 'Unknown rendering error').slice(0,600));
}
function safeAsset(root, relative) {
  // No arbitrary path, scripts outside the pinned package, directory listings or URL decoding escapes.
  if (!/^(?:legacy\/build\/[A-Za-z0-9_.-]+\.mjs|(?:cmaps|standard_fonts|wasm|iccs)\/[A-Za-z0-9_.-]+)$/.test(relative)) return null;
  const full = path.join(root, ...relative.split('/'));
  try {
    const realRoot = fs.realpathSync(root), real = fs.realpathSync(full), rel = path.relative(realRoot, real);
    if (!rel || rel.startsWith('..' + path.sep) || rel === '..' || path.isAbsolute(rel)) return null;
    if (!fs.statSync(real).isFile()) return null;
    return full;
  } catch (_) { return null; }
}
function validatePng(bytes, limits = LIMITS) {
  if (!Buffer.isBuffer(bytes) || bytes.length < 33 || bytes.length > limits.maxPngBytes || !bytes.subarray(0,8).equals(PNG) || bytes.toString('ascii',12,16) !== 'IHDR')
    throw failure('PDF_PROTOCOL', 'The PDF renderer returned an invalid page image.');
  const width = bytes.readUInt32BE(16), height = bytes.readUInt32BE(20);
  if (!width || !height || width > limits.maxEdge || height > limits.maxEdge || width * height > limits.maxPixels)
    throw failure('PDF_LIMIT', 'A rendered page exceeds the configured image limits.');
  return { width, height };
}
function boundedBody(request, max) {
  return new Promise((resolve, reject) => {
    let total = 0, parts = [];
    request.on('data', chunk => {
      total += chunk.length;
      if (total > max) { reject(failure('PDF_LIMIT', 'PDF renderer message exceeds the permitted size.')); request.destroy(); }
      else parts.push(chunk);
    });
    request.on('end', () => resolve(Buffer.concat(parts)));
    request.on('error', reject);
    request.on('aborted', () => reject(failure('PDF_CANCELLED', 'PDF import was cancelled.')));
  });
}
/** Exported for integration tests. The only file exposed is the selected PDF; only page PNGs can be written. */
async function startJob({ input, outDir, packageRoot, limits = LIMITS, onProgress = () => {} }) {
  const handle = await fsp.open(input, 'r');
  const stat = await handle.stat();
  if (!stat.isFile() || !stat.size || stat.size > limits.maxBytes) {
    await handle.close(); throw failure('PDF_LIMIT', 'Choose a nonempty PDF smaller than 128 MB. Split larger documents before importing.');
  }
  const signature = Buffer.alloc(Math.min(1024, stat.size));
  await handle.read(signature, 0, signature.length, 0);
  if (!signature.includes(Buffer.from('%PDF-'))) { await handle.close(); throw failure('PDF_INVALID', 'The selected file does not contain a PDF header.'); }
  let root;
  try {
    await fsp.mkdir(outDir, { recursive:true });
    if ((await fsp.readdir(outDir)).some(f => /^slide-?\d+\.png$/i.test(f)))
      throw failure('PDF_OUTPUT_NOT_EMPTY', 'The PDF output folder is not empty. No existing slides were changed.');
    root = await fsp.realpath(outDir);
  } catch(e) { await handle.close(); throw e; }
  const token = crypto.randomBytes(32).toString('hex');
  let origin = '', total = 0, next = 1, outputBytes = 0, done = false, writing = false;
  let resolveResult, rejectResult, idleTimer, totalTimer;
  const written = [];
  const result = new Promise((resolve, reject) => { resolveResult = resolve; rejectResult = reject; });
  // Prevent an unhandled rejection if Electron fails before the caller awaits result.
  result.catch(() => {});
  function settle(error) {
    if (done) return; done = true; clearTimeout(idleTimer); clearTimeout(totalTimer);
    if (error) rejectResult(error); else resolveResult({method:'pdfjs-bundled', pageCount:total, version:PDFJS_VERSION});
  }
  function tick() { clearTimeout(idleTimer); idleTimer = setTimeout(() => settle(failure('PDF_TIMEOUT', 'PDF import made no progress for 90 seconds. Try a smaller or re-exported PDF.')), limits.pageTimeout); }
  function json(res, code, value) { res.writeHead(code, {'Content-Type':'application/json; charset=utf-8'});res.end(JSON.stringify(value)); }
  const server = http.createServer(async (req, res) => {
    try {
      res.setHeader('Cache-Control','no-store'); res.setHeader('X-Content-Type-Options','nosniff');
      res.setHeader('Referrer-Policy','no-referrer');
      res.setHeader('Content-Security-Policy', "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self'; connect-src 'self'; img-src 'self' data: blob:; font-src 'self' blob:; style-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'");
      if (req.headers.host !== new URL(origin).host) return json(res,403,{error:'Forbidden host'});
      const u = new URL(req.url, origin), prefix = '/' + token + '/';
      if (!u.pathname.startsWith(prefix) || u.search || u.hash || done) return json(res,404,{error:'No such job'});
      const route = u.pathname.slice(prefix.length);
      if (req.method === 'POST') {
        if (req.headers.origin !== origin || req.headers['x-anchorcast-pdf'] !== token) return json(res,403,{error:'Forbidden operation'});
        if (route === 'cancel') { json(res,200,{success:true});settle(failure('PDF_CANCELLED','PDF import cancelled. The original file and library are unchanged.'));return; }
        if (route === 'meta') {
          const meta = JSON.parse((await boundedBody(req,4096)).toString('utf8'));
          if (total || !Number.isInteger(meta.pages) || meta.pages < 1 || meta.pages > limits.maxPages) throw failure('PDF_LIMIT','Import is limited to 500 pages. Split this PDF into smaller documents.');
          total = meta.pages;tick();onProgress({page:0,total});return json(res,200,{success:true});
        }
        if (/^page\/\d+$/.test(route)) {
          const page = Number(route.slice(5));
          if (writing || !total || page !== next || page > total) throw failure('PDF_PROTOCOL','PDF pages arrived out of order.');
          writing = true;
          try {
            const bytes = await boundedBody(req,limits.maxPngBytes);validatePng(bytes,limits);
            if (done) throw failure('PDF_CANCELLED','PDF import was cancelled.');
            outputBytes += bytes.length;
            if (outputBytes > limits.maxOutputBytes) throw failure('PDF_LIMIT','Converted slides exceed 1 GB. Split this PDF into smaller documents.');
            if (await fsp.realpath(outDir) !== root) throw failure('PDF_PROTOCOL','The PDF output directory changed during import.');
            const dest = path.join(outDir,'slide-'+String(page).padStart(4,'0')+'.png');
            await fsp.writeFile(dest, bytes, {flag:'wx',mode:0o600});written.push(dest);
            next++;tick();onProgress({page,total});return json(res,200,{success:true});
          } finally { writing = false; }
        }
        if (route === 'done') {
          if (!total || next !== total + 1 || writing) throw failure('PDF_PROTOCOL','PDF rendering finished before every page was saved.');
          json(res,200,{success:true});settle();return;
        }
        if (route === 'error') {
          const error = JSON.parse((await boundedBody(req,4096)).toString('utf8'));
          json(res,200,{success:true});settle(sanitizeError(error.name,error.message));return;
        }
        return json(res,404,{error:'Unknown operation'});
      }
      if (req.method !== 'GET' && req.method !== 'HEAD') return json(res,405,{error:'Method not allowed'});
      if (route === 'job.json') return json(res,200,{version:PDFJS_VERSION,maxPages:limits.maxPages,maxEdge:limits.maxEdge,maxPixels:limits.maxPixels});
      if (route === 'document.pdf') {
        res.setHeader('Content-Type','application/pdf');res.setHeader('Content-Length',stat.size);
        if (req.method === 'HEAD') { res.end();return; }
        const stream = handle.createReadStream({start:0,end:stat.size-1,autoClose:false});
        stream.on('error',e=>res.destroy(e));res.on('close',()=>stream.destroy());stream.pipe(res);return;
      }
      let file;
      if (['index.html','render.mjs','style.css'].includes(route)) file=path.join(__dirname,'pdf-renderer',route);
      else if (route.startsWith('pdfjs/')) file=safeAsset(packageRoot,route.slice(6));
      if (!file) return json(res,404,{error:'Unknown resource'});
      const data = await fsp.readFile(file);
      res.setHeader('Content-Type',file.endsWith('.html')?'text/html; charset=utf-8':/\.m?js$/.test(file)?'text/javascript; charset=utf-8':file.endsWith('.css')?'text/css; charset=utf-8':file.endsWith('.wasm')?'application/wasm':'application/octet-stream');
      res.setHeader('Content-Length',data.length);res.end(req.method==='HEAD'?undefined:data);
    } catch (e) {
      if (!res.headersSent && !res.destroyed) json(res,400,{error:String(e.message).slice(0,600)});
      else if (!res.destroyed) res.destroy();
      if (req.method==='POST') settle(e.code ? e : failure('PDF_PROTOCOL','PDF renderer sent an invalid message.'));
    }
  });
  server.requestTimeout = 60000;server.headersTimeout=15000;
  server.on('error',e=>settle(failure('PDF_SERVER',e.message)));
  try { await new Promise((resolve,reject) => {server.once('error',reject);server.listen(0,'127.0.0.1',resolve);}); }
  catch(e){await handle.close();throw e;}
  origin='http://127.0.0.1:'+server.address().port;
  tick();totalTimer=setTimeout(()=>settle(failure('PDF_TIMEOUT','PDF import exceeded 15 minutes. Split the document and try again.')),limits.totalTimeout);
  let closed=false;
  return {origin,token,url:origin+'/'+token+'/index.html',result,
    cancel: () => settle(failure('PDF_CANCELLED','PDF import cancelled. The original file and library are unchanged.')),
    fail: e => settle(e),
    async close(success=false) {
      if(closed)return;closed=true;
      if(!done)settle(failure('PDF_CANCELLED','PDF import was cancelled.'));
      server.closeAllConnections?.();await new Promise(resolve=>server.close(resolve));
      await handle.close().catch(()=>{});
      // Wait for an already acknowledged upload to finish before removing partial pages.
      while(writing)await new Promise(resolve=>setTimeout(resolve,5));
      if(!success)for(const file of written)await fsp.unlink(file).catch(()=>{});
    }
  };
}
let active = false;
async function renderPdf(input, outDir, options={}) {
  if(active)throw failure('PDF_BUSY','Another PDF is being imported. Finish or cancel it before importing another.');
  active=true;
  let job,win,success=false;
  try {
    const status=getRendererStatus({appPath:options.appPath});
    if(!status.ready)throw failure(status.code,status.error);
    let electron;
    try{electron=options.electron||require('electron');}catch(_){throw failure('PDF_RENDERER_UNAVAILABLE','Built-in PDF rendering runs inside the AnchorCast desktop app.');}
    if(!electron.BrowserWindow || !electron.app?.isReady())throw failure('PDF_RENDERER_UNAVAILABLE','Open AnchorCast before using the built-in PDF renderer.');
    job=await startJob({input,outDir,packageRoot:status.root,onProgress:options.onProgress});
    const parent=options.parent && !options.parent.isDestroyed()?options.parent:undefined;
    win=new electron.BrowserWindow({width:520,height:270,show:false,title:'Import PDF — AnchorCast',
      ...(parent?{parent}:{}),resizable:false,minimizable:false,maximizable:false,
      autoHideMenuBar:true,webPreferences:{nodeIntegration:false,contextIsolation:true,sandbox:true,
        webSecurity:true,allowRunningInsecureContent:false,webviewTag:false,plugins:false,
        backgroundThrottling:false,spellcheck:false,partition:'anchorcast-pdf-'+job.token}});
    const wc=win.webContents, ses=wc.session;
    ses.setPermissionRequestHandler((_,__,callback)=>callback(false));
    ses.setPermissionCheckHandler(()=>false);
    ses.on('will-download',event=>event.preventDefault());
    ses.webRequest.onBeforeRequest((details,callback)=>{
      const allowed=allowedResourceURL(details.url,job.origin,job.token);
      callback({cancel:!allowed});
    });
    wc.setWindowOpenHandler(()=>({action:'deny'}));
    wc.on('will-navigate',event=>event.preventDefault());
    wc.on('will-attach-webview',event=>event.preventDefault());
    wc.on('render-process-gone',(_,details)=>job.fail(failure('PDF_RENDER_FAILED','The PDF renderer stopped ('+details.reason+'). Try a smaller or re-exported PDF.')));
    win.on('closed',()=>job.cancel());
    await win.loadURL(job.url);
    if(!win.isDestroyed())win.show();
    const result=await job.result;success=true;return result;
  } finally {
    if(win&&!win.isDestroyed())win.destroy();
    if(job)await job.close(success);
    active=false;
  }
}
module.exports={PDFJS_VERSION,LIMITS,getRendererStatus,renderPdf,startJob,safeAsset,validatePng,sanitizeError,allowedResourceURL};

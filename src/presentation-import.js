'use strict';
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),crypto=require('node:crypto');
const {execFile}=require('node:child_process');
const {promisify}=require('node:util');
const {pathToFileURL}=require('node:url');
const exec=promisify(execFile);
function pythonCandidates({platform=process.platform,resourcesPath=process.resourcesPath,appPath='',userDataDir='',extra=[]}={}) {
  // Candidate paths belong to the requested platform, not necessarily the
  // Node host running a cross-platform test. Keep native filesystem operations
  // in createImporter on the host path API.
  const candidatePath=platform==='win32'?path.win32:path.posix;
  const roots=[resourcesPath,appPath,candidatePath.dirname(appPath||'.'),userDataDir].filter(Boolean);
  const list=[...extra];
  for(const root of roots) {
    if(platform==='win32')list.push(candidatePath.join(root,'python','python.exe'));
    else list.push(candidatePath.join(root,'python','bin','python3'),candidatePath.join(root,'python','python3'));
  }
  if(platform==='darwin')list.push('/opt/homebrew/bin/python3','/usr/local/bin/python3','/Library/Frameworks/Python.framework/Versions/Current/bin/python3','/usr/bin/python3');
  list.push(...(platform==='win32'?['python','python3','py']:['python3','python']));
  return [...new Set(list.filter(x=>typeof x==='string'&&x))].map(command=>({command,args:/^py(?:\.exe)?$/i.test(candidatePath.basename(command))?['-3']:[]}));
}
function toolCandidates(tool,platform=process.platform,home=os.homedir()) {
  const candidatePath=platform==='win32'?path.win32:path.posix;
  if(tool==='soffice')return [...new Set([
    ...(platform==='darwin'?['/Applications/LibreOffice.app/Contents/MacOS/soffice',candidatePath.join(home,'Applications/LibreOffice.app/Contents/MacOS/soffice')]:[]),
    ...(platform==='win32'?['C:/Program Files/LibreOffice/program/soffice.exe','C:/Program Files (x86)/LibreOffice/program/soffice.exe']:['/usr/bin/libreoffice','/usr/bin/soffice']),
    'soffice','libreoffice'])];
  return [...new Set([...(platform==='darwin'?['/opt/homebrew/bin/pdftoppm','/usr/local/bin/pdftoppm']:[]),'pdftoppm'])];
}
function parseOutput(stdout) {
  const text=String(stdout||'').trim();
  try{return JSON.parse(text);}catch(_){}
  const last=text.split(/\r?\n/).filter(Boolean).pop();
  try{return JSON.parse(last);}catch(_){throw Error('The presentation converter returned invalid JSON.');}
}
function createImporter(options) {
  const {store,presentationDir,dataDir,userDataDir=''}=options;
  const platform=options.platform||process.platform, run=options.exec||exec;
  const appPath=options.app?.getAppPath?.()||path.resolve(__dirname,'..');
  const resourcesPath=options.resourcesPath||process.resourcesPath;
  function clearSlides(dir) {for(const s of store.files(dir))try{fs.unlinkSync(s.imagePath);}catch(_) {}}
  const settings={timeout:180000,maxBuffer:4*1024*1024,windowsHide:true};
  function nativeTools(kind){return options.tools?.[kind]??toolCandidates(kind,platform);}
  async function python(input,outDir,diagnostics) {
    const helper=path.join(__dirname,'pptx_to_png.py');
    if(!fs.existsSync(helper)){diagnostics.push('Conversion helper pptx_to_png.py is missing from this installation.');return false;}
    const bytes=fs.readFileSync(helper),tag=crypto.createHash('sha256').update(bytes).digest('hex').slice(0,16);
    const toolsDir=path.join(dataDir,'presentation-tools');fs.mkdirSync(toolsDir,{recursive:true});
    const script=path.join(toolsDir,'pptx_to_png-'+tag+'.py');
    // External Python cannot read Electron's virtual app.asar paths.
    if(!fs.existsSync(script))fs.writeFileSync(script,bytes,{flag:'wx',mode:0o600});
    const extra=options.getPythonCandidates?.()||[];
    let usable=0;
    for(const p of (options.pythonCandidates??pythonCandidates({platform,appPath,resourcesPath,userDataDir,extra}))) {
      if(path.isAbsolute(p.command)&&!fs.existsSync(p.command)&&!options.exec)continue;
      try {
        const probe=await run(p.command,[...p.args,'-c','import sys; print(sys.version_info[0]); sys.exit(0 if sys.version_info[0] == 3 else 1)'],{...settings,timeout:8000});
        if(String(probe.stdout).trim()!=='3')continue;
        usable++;
      } catch(e){if(e.code!=='ENOENT')diagnostics.push(path.basename(p.command)+': Python could not start ('+(e.code||'runtime error')+').');continue;}
      try {
        const result=await run(p.command,[...p.args,script,input,outDir],settings);
        const parsed=parseOutput(result.stdout);
        if(parsed.success&&store.files(outDir).length)return {method:parsed.method||'python',warning:parsed.warning||'',python:p.command};
        clearSlides(outDir);
        diagnostics.push(path.basename(p.command)+': '+String(parsed.error||'No slide images produced.').slice(0,700));
      } catch(e){clearSlides(outDir);diagnostics.push(path.basename(p.command)+': '+String(e.stderr||e.message||e).slice(0,700));}
    }
    if(!usable)diagnostics.push('No usable Python 3 interpreter was found in bundled or system locations.');
    return false;
  }
  async function rasterize(pdf,outDir,diagnostics) {
    // First choice: offline PDF.js shipped in assets/pdfjs. No external Python needed.
    // Injection is only for trusted main-process tests; renderer IPC cannot set it.
    if (options.pdfRenderer !== false) {
      try {
        const render = options.pdfRenderer || require('./presentation-pdf').renderPdf;
        const result = await render(pdf, outDir, { appPath, parent: options.getMainWindow?.() });
        const pages = store.files(outDir);
        if (!result || !pages.length || (result.pageCount && pages.length !== result.pageCount))
          throw Error('The built-in PDF renderer did not save every page.');
        return { method: result.method || 'pdfjs-bundled' };
      } catch (error) {
        clearSlides(outDir);
        // Cancellation, input restrictions and resource limits must never trigger another converter.
        if (['PDF_CANCELLED','PDF_PASSWORD','PDF_INVALID','PDF_LIMIT','PDF_BUSY','PDF_UNSUPPORTED'].includes(error.code)) throw error;
        diagnostics.push('Built-in PDF renderer: ' + String(error.message || error).slice(0,900));
      }
    }
    for(const cmd of nativeTools('pdftoppm')) {
      try {await run(cmd,['-png','-scale-to','1920',pdf,path.join(outDir,'slide')],settings);
        if(store.files(outDir).length)return {method:'pdf-poppler'};
      }catch(e){clearSlides(outDir);if(e.code!=='ENOENT')diagnostics.push('PDF renderer: '+String(e.stderr||e.message).slice(0,500));}
    }
    return python(pdf,outDir,diagnostics);
  }
  async function importFile(input) {
    if(typeof input!=='string'||!input||input.includes('\0'))return {success:false,error:'Select a PowerPoint, OpenDocument presentation or PDF file.'};
    const src=path.resolve(input),originalExt=path.extname(src),ext=originalExt.toLowerCase();
    if(!['.pptx','.ppt','.odp','.pdf'].includes(ext))return {success:false,error:'Unsupported presentation format. Export the file as .pptx or .pdf and import that file.'};
    let stat;try{stat=fs.statSync(src);}catch(_){return {success:false,error:'The selected file no longer exists or cannot be read.'};}
    if(!stat.isFile())return {success:false,error:'Select a file, not a folder.'};
    const id=crypto.randomUUID(),outDir=path.join(presentationDir,id),diagnostics=[];
    fs.mkdirSync(outDir,{recursive:true});
    let method=null;
    try {
      if(platform==='win32'&&['.pptx','.ppt'].includes(ext)) {
        const script=path.join(outDir,'convert.ps1');
        fs.writeFileSync(script,[
          'param([string]$src, [string]$outDir)',"$ErrorActionPreference = 'Stop'",'$ppt = $null; $pres = $null','try {',
          '$ppt = New-Object -ComObject PowerPoint.Application', '$ppt.AutomationSecurity = 3',
          '$pres = $ppt.Presentations.Open($src, $true, $false, $false)',"$pres.Export($outDir, 'PNG', 1920, 1080)",
          '} finally { if ($pres) { $pres.Close() }; if ($ppt) { $ppt.Quit() } }'
        ].join('\r\n'));
        try{await run('powershell.exe',['-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',script,src,outDir],settings);
          if(store.files(outDir).length)method={method:'powerpoint'};
        }catch(e){clearSlides(outDir);diagnostics.push('PowerPoint automation was unavailable or failed.');}
        finally{try{fs.unlinkSync(script);}catch(_) {}}
      }
      if(!method&&ext!=='.pdf') {
        for(const cmd of nativeTools('soffice')) {
          const work=path.join(outDir,'pdf-export'),profile=path.join(outDir,'office-profile');fs.mkdirSync(work,{recursive:true});
          try {
            await run(cmd,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--convert-to','pdf','--outdir',work,src],settings);
            const pdf=path.join(work,path.basename(src,originalExt)+'.pdf');
            if(fs.existsSync(pdf)) {method=await rasterize(pdf,outDir,diagnostics);if(method){method.method='libreoffice+'+method.method;break;}}
          }catch(e){if(e.code!=='ENOENT')diagnostics.push('LibreOffice conversion failed: '+String(e.stderr||e.message).slice(0,400));}
        }
      }
      if(!method)method=ext==='.pdf'?await rasterize(src,outDir,diagnostics):await python(src,outDir,diagnostics);
      if(!method) {
        const unique = [...new Set(diagnostics)];
        if (ext === '.pdf') {
          const builtIn = unique.find(message => message.startsWith('Built-in PDF renderer:'));
          throw Error('PDF import could not finish.\n' +
            (builtIn || 'The built-in PDF renderer is not available in this installation.') +
            '\n\nPDF files do not require PowerPoint or LibreOffice. If an older external renderer is used, install Poppler or PyMuPDF in that Python interpreter.\n\n' +
            unique.filter(message => message !== builtIn).slice(-3).join('\n'));
        }
        throw Error('Presentation import could not finish with the available converters.\n'+
          (ext==='.pptx'?'Basic PPTX import needs python-pptx and Pillow in the same Python interpreter used by AnchorCast. ':'')+
          'For faithful PowerPoint layout, use PowerPoint on Windows or LibreOffice. You can also export to PDF and use AnchorCast’s built-in PDF importer.\n\n'+unique.slice(-5).join('\n'));
      }
      const slides=store.files(outDir);
      if(!slides.length)throw Error('The converter produced no slide images.');
      for(const s of slides) {
        const fd=fs.openSync(s.imagePath,'r'),header=Buffer.alloc(8);try{fs.readSync(fd,header,0,8,0);}finally{fs.closeSync(fd);}
        if(header.toString('hex')!=='89504e470d0a1a0a')throw Error('A converted slide is not a valid PNG image.');
      }
      // Reload immediately before commit: concurrent imports/deletes cannot resurrect old lists.
      const record={id,name:path.basename(src,originalExt),type:'imported',resolvedPath:src,outDir,slideCount:slides.length,importedAt:Date.now(),converter:method.method};
      const list=store.read('imported');list.unshift(record);store.write('imported',list);
      for(const d of ['pdf-export','office-profile'])try{fs.rmSync(path.join(outDir,d),{recursive:true,force:true});}catch(_){}
      return {success:true,...record,slides,warning:method.warning||'',diagnostics};
    }catch(e){try{fs.rmSync(outDir,{recursive:true,force:true});}catch(_){}return {success:false,cancelled:e.code==='PDF_CANCELLED',code:e.code||'IMPORT_FAILED',error:e.message,diagnostics:[...new Set(diagnostics)]};}
  }
  return {importFile};
}
module.exports={createImporter,pythonCandidates,toolCandidates,parseOutput};

'use strict';
// Independent local PCM -> MP3 recording. This module makes no network calls.
const fs=require('node:fs'),fsp=fs.promises,path=require('node:path'),crypto=require('node:crypto');
const {spawn}=require('node:child_process'),{EventEmitter}=require('node:events'),{fileURLToPath}=require('node:url');
const NAME=/^Sermon-\d{8}-\d{6}-[a-f0-9]{32}(?:\.interrupted)?\.mp3$/;
const PART=/^Sermon-\d{8}-\d{6}-[a-f0-9]{32}\.partial\.mp3$/;
const MAX_CHUNK=65536,MAX_PENDING=1048576;
function deadline(p,ms,label){let t;return Promise.race([p,new Promise((_,reject)=>{t=setTimeout(()=>reject(Error(label)),ms);})]).finally(()=>clearTimeout(t));}
function bundledEncoder(){
  let p;try{p=path.join(path.dirname(require.resolve('ffmpeg-static/package.json')),process.platform==='win32'?'ffmpeg.exe':'ffmpeg');}catch(_){throw Error('MP3 encoder is not installed. Install ffmpeg-static, then rebuild/restart AnchorCast.');}
  if(typeof p!=='string'||!path.isAbsolute(p))throw Error('MP3 encoder is unavailable on this platform.');
  p=p.replace(/app\.asar([\\/])/,'app.asar.unpacked$1');
  if(!fs.existsSync(p))throw Error('Bundled MP3 encoder is missing. Rebuild or reinstall this installer.');
  return p;
}
class RecorderService extends EventEmitter{
  constructor({directory,encoder=bundledEncoder,spawnProcess=spawn}){super();Object.assign(this,{directory,encoder,spawnProcess});this.active=null;this.starting=false;}
  get busy(){return this.starting||!!this.active;}
  async begin({sampleRate=48000}={}){
    if(this.busy)throw Error('An audio recording is already running.');
    if(!Number.isInteger(sampleRate)||sampleRate<8000||sampleRate>192000)throw Error('Invalid microphone sample rate.');
    this.starting=true;let s,output;
    try{
      const executable=this.encoder();await fsp.mkdir(this.directory,{recursive:true,mode:0o700});
      const stamp=new Date().toISOString().replace(/[-:]/g,'').replace('T','-').slice(0,15),base=`Sermon-${stamp}-${crypto.randomBytes(16).toString('hex')}`;
      const partial=path.join(this.directory,base+'.partial.mp3');
      output=fs.createWriteStream(partial,{flags:'wx',mode:0o600,flush:true});
      // Attach error handling before spawning; disk failures must not crash Electron.
      let earlyOutputError=null;output.on('error',e=>{earlyOutputError=e;});
      const child=this.spawnProcess(executable,['-hide_banner','-loglevel','error','-f','s16le','-ar',String(sampleRate),'-ac','1','-i','pipe:0','-vn','-c:a','libmp3lame','-b:a','128k','-ar','44100','-f','mp3','pipe:1'],{shell:false,windowsHide:true,stdio:['pipe','pipe','pipe']});
      s={id:crypto.randomUUID(),base,partial,output,child,sampleRate,samples:0,sequence:0,pending:0,tail:Promise.resolve(),stderr:'',error:null,ending:false};this.active=s;
      const fail=e=>{s.error ||= e instanceof Error?e:Error(String(e));if(!s.ending)this.emit('fault',s.error.message);try{child.kill();}catch(_){}};
      s.exited=new Promise(resolve=>{
        child.once('error',e=>{s.error=e;resolve(-1);if(!s.ending)this.emit('fault',e.message);output.end();});
        child.once('close',code=>{if(code!==0)s.error ||= Error(s.stderr||'MP3 encoder stopped unexpectedly.');resolve(code);if(!s.ending&&s.error)this.emit('fault',s.error.message);});
      });
      s.closed=new Promise(resolve=>{output.once('close',resolve);output.on('error',fail);});
      child.stdin.on('error',e=>{s.error ||= e;if(!s.ending)this.emit('fault',e.message);});
      child.stderr.on('data',c=>{s.stderr=(s.stderr+c.toString()).slice(-3000);});
      child.stdout.on('error',fail);child.stdout.pipe(output);
      await deadline(Promise.all([
        new Promise((yes,no)=>{child.once('spawn',yes);child.once('error',no);}),
        new Promise((yes,no)=>{output.once('open',yes);output.once('error',no);})
      ]),5000,'MP3 encoder did not start.');
      if(s.error||earlyOutputError)throw s.error||earlyOutputError;
      return {id:s.id,sampleRate,bitrate:128,channels:1};
    }catch(e){
      if(s){s.ending=true;try{s.child.kill();}catch(_){}s.output.destroy();await deadline(s.closed,2000,'Close timeout').catch(()=>{});await fsp.unlink(s.partial).catch(()=>{});}else output?.destroy();
      this.active=null;throw e;
    }finally{this.starting=false;}
  }
  append({id,sequence,pcm}={}){
    const s=this.active;
    if(!s||id!==s.id||s.ending)return Promise.reject(Error('No matching active recording.'));
    if(s.error)return Promise.reject(s.error);
    if(!Number.isInteger(sequence)||sequence!==s.sequence)return Promise.reject(Error('Audio chunk out of order.'));
    if(!(pcm instanceof Uint8Array)&&!(pcm instanceof ArrayBuffer))return Promise.reject(Error('Audio chunk must be binary PCM.'));
    const view=pcm instanceof ArrayBuffer?new Uint8Array(pcm):pcm;
    if(!view.byteLength||view.byteLength>MAX_CHUNK||view.byteLength%2)return Promise.reject(Error('Invalid audio chunk length.'));
    if(s.pending+view.byteLength>MAX_PENDING)return Promise.reject(Error('Audio encoder cannot keep up.'));
    const data=Buffer.from(view);s.pending+=data.length;s.sequence++;
    const op=s.tail.then(async()=>{if(s.error)throw s.error;await deadline(new Promise((yes,no)=>s.child.stdin.write(data,e=>e?no(e):yes())),10000,'MP3 write timed out.');s.samples+=data.length/2;return {seconds:s.samples/s.sampleRate};})
      .catch(e=>{s.error ||= e;throw e;}).finally(()=>{s.pending-=data.length;});
    s.tail=op.catch(()=>{});return op;
  }
  async finish({id,interrupted=false}={}){
    const s=this.active;if(!s||id!==s.id)throw Error('No matching active recording.');if(s.finishPromise)return s.finishPromise;s.ending=true;
    s.finishPromise=(async()=>{let completed=false;try{
      await deadline(s.tail,12000,'Pending audio could not be saved.');s.child.stdin.end();
      await deadline(Promise.all([s.exited,s.closed]),15000,'MP3 encoder did not finish.');
      if(!s.samples){await fsp.unlink(s.partial).catch(()=>{});if(s.error)throw s.error;completed=true;return {saved:false,empty:true};}
      const stat=await fsp.stat(s.partial);if(stat.size<200)throw Error('No usable MP3 audio was produced.');
      const isPartial=interrupted||!!s.error,name=s.base+(isPartial?'.interrupted':'')+'.mp3',file=path.join(this.directory,name);
      await fsp.rename(s.partial,file);completed=true;
      return {saved:true,name,path:file,seconds:s.samples/s.sampleRate,bytes:stat.size,interrupted:isPartial,error:s.error?.message||null};
    }catch(e){try{s.child.kill();}catch(_){}s.output.destroy();throw Error(e.message+' Any written audio remains in Recordings as a partial MP3.');}
    finally{if(this.active===s)this.active=null;this.emit('finished',{completed});}})();return s.finishPromise;
  }
  async list(){
    await fsp.mkdir(this.directory,{recursive:true,mode:0o700});const rows=[];
    for(const name of await fsp.readdir(this.directory)){
      if(!NAME.test(name)&&!PART.test(name))continue;const file=path.join(this.directory,name);if(this.active?.partial===file)continue;
      const stat=await fsp.lstat(file).catch(()=>null);if(!stat?.isFile()||stat.isSymbolicLink()||stat.size<200)continue;
      rows.push({name,bytes:stat.size,modifiedAt:stat.mtimeMs,interrupted:PART.test(name)||name.includes('.interrupted.')});
    }return rows.sort((a,b)=>b.modifiedAt-a.modifiedAt).slice(0,30);
  }
  async sourceFile(name){
    if(typeof name!=='string'||(!NAME.test(name)&&!PART.test(name)))throw Error('Invalid recording name.');
    const file=path.join(this.directory,name),stat=await fsp.lstat(file);
    if(!stat.isFile()||stat.isSymbolicLink()||file===this.active?.partial)throw Error('Recording is not ready to export.');return file;
  }
}
function install({app,ipcMain,dialog,shell,powerSaveBlocker,getMainWindow,getRendererPort,systemPreferences}){
  const service=new RecorderService({directory:path.join(app.getPath('userData'),'AnchorCastData','Recordings')});
  let blocker=null,pendingClose=null,closeTimer=null;
  function trusted(e){
    const w=getMainWindow();if(!w||w.isDestroyed()||e.sender!==w.webContents||e.senderFrame!==e.sender.mainFrame)return false;
    try{const u=new URL(e.senderFrame.url);return u.protocol==='http:'?u.hostname==='127.0.0.1'&&u.port===String(getRendererPort())&&u.pathname==='/index.html':u.protocol==='file:'&&path.resolve(fileURLToPath(u))===path.join(__dirname,'renderer','index.html');}catch(_){return false;}
  }
  function send(c,d){const w=getMainWindow();if(w&&!w.isDestroyed())w.webContents.send(c,d);}
  function releaseBlocker(){if(blocker!==null){try{powerSaveBlocker.stop(blocker);}catch(_){}blocker=null;}}
  function handler(name,fn){ipcMain.handle('sermon-audio:'+name,async(e,d)=>{if(!trusted(e))return {success:false,error:'Audio recorder is available only in the main application window.'};try{return {success:true,...await fn(d||{})};}catch(err){return {success:false,error:err.message};}});}
  handler('info',async()=>{let error=null;try{bundledEncoder();}catch(e){error=e.message;}return {revision:'audio-recorder-v1',available:!error,error,recordings:await service.list()};});
  handler('permission',async()=>{if(process.platform==='darwin'&&systemPreferences&&!await systemPreferences.askForMediaAccess('microphone'))throw Error('Allow AnchorCast microphone access in macOS System Settings → Privacy & Security → Microphone.');return {allowed:true};});
  handler('begin',async d=>{const r=await service.begin(d);try{blocker=powerSaveBlocker.start('prevent-app-suspension');}catch(_){}return r;});
  handler('append',d=>service.append(d));handler('finish',d=>service.finish(d));
  handler('open-folder',async()=>{await fsp.mkdir(service.directory,{recursive:true});const error=await shell.openPath(service.directory);if(error)throw Error(error);return {};});
  handler('save-copy',async({name})=>{
    const src=await service.sourceFile(name),choice=await dialog.showSaveDialog(getMainWindow(),{title:'Save sermon recording as MP3',defaultPath:path.join(app.getPath('music'),name.replace('.partial.','.interrupted.')),filters:[{name:'MP3 audio',extensions:['mp3']}],properties:['showOverwriteConfirmation','createDirectory']});
    if(choice.canceled||!choice.filePath)return {cancelled:true};let dest=choice.filePath;if(path.extname(dest).toLowerCase()!=='.mp3')dest+='.mp3';
    if(path.resolve(src)===path.resolve(dest))return {path:src};
    // When extension normalization changes the selected path, confirm separately.
    if(dest!==choice.filePath&&fs.existsSync(dest)){
      const confirm=await dialog.showMessageBox(getMainWindow(),{type:'question',message:'Replace the existing MP3?',detail:dest,buttons:['Cancel','Replace'],cancelId:0,defaultId:0});if(confirm.response!==1)return {cancelled:true};
    }
    const temp=dest+'.'+crypto.randomBytes(8).toString('hex')+'.tmp';
    try{await fsp.copyFile(src,temp,fs.constants.COPYFILE_EXCL);await fsp.rename(temp,dest);}finally{await fsp.unlink(temp).catch(()=>{});}return {path:dest};
  });
  service.on('fault',message=>send('sermon-audio-fault',message));
  service.on('finished',({completed})=>{releaseBlocker();if(!pendingClose)return;clearTimeout(closeTimer);const next=pendingClose;pendingClose=null;if(!completed){send('sermon-audio-fault','MP3 could not be finalized. AnchorCast stayed open. Check Recordings for partial audio.');return;}setTimeout(()=>{if(next==='quit')app.quit();else{const w=getMainWindow();if(w&&!w.isDestroyed())w.close();}},150);});
  function guard(e,kind){
    const w=getMainWindow();if(!service.busy||!w||w.isDestroyed())return;e.preventDefault();if(pendingClose)return;
    const index=dialog.showMessageBoxSync(w,{type:'question',title:'Sermon audio is recording',message:'Stop and save the recording before closing?',detail:'Keep recording leaves AnchorCast open. Stop and save finishes the MP3 in Recordings.',buttons:['Keep recording','Stop and save'],defaultId:0,cancelId:0,noLink:true});
    if(index===1){pendingClose=kind;send('sermon-audio-stop-request',{});closeTimer=setTimeout(()=>{pendingClose=null;send('sermon-audio-fault','Audio did not finish in time. AnchorCast stayed open; use Stop & Save or open Recordings.');},20000);}
  }
  app.prependListener('before-quit',e=>guard(e,'quit'));
  app.on('browser-window-created',(_event,win)=>{
    win.on('close',e=>{if(win===getMainWindow())guard(e,'window');});
    win.webContents.on('render-process-gone',()=>{if(win===getMainWindow()&&service.active)void service.finish({id:service.active.id,interrupted:true}).catch(()=>{});});
    win.webContents.on('did-start-navigation',(_e,_url,_inPlace,isMainFrame)=>{if(isMainFrame&&win===getMainWindow()&&service.active)void service.finish({id:service.active.id,interrupted:true}).catch(()=>{});});
  });return service;
}
module.exports={RecorderService,install,bundledEncoder,NAME,PART,MAX_CHUNK,MAX_PENDING};

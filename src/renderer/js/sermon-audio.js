/* Opt-in recorder independent of transcription and its registration controls. */
(function(root){'use strict';
  function timed(p,ms,label){let id;return Promise.race([p,new Promise((_,no)=>{id=setTimeout(()=>no(Error(label)),ms);})]).finally(()=>clearTimeout(id));}
  class SermonAudioRecorder {
    constructor({bridge,mediaDevices=root.navigator?.mediaDevices,AudioContextClass=root.AudioContext,WorkletClass=root.AudioWorkletNode,getMicrophone=()=>'',onChange=()=>{},workletUrl='js/sermon-audio-worklet.js'}){
      Object.assign(this,{bridge,mediaDevices,AudioContextClass,WorkletClass,getMicrophone,onChange,workletUrl});
      this.state='idle';this.seconds=0;this.peak=0;this.message='';this.saved=null;this.session=null;this.startPromise=null;this.stopPromise=null;this.cancelStart=false;
    }
    emit(){this.onChange({state:this.state,seconds:this.seconds,peak:this.peak,message:this.message,saved:this.saved});}
    async call(method,data){const r=await this.bridge[method](data);if(!r?.success)throw Error(r?.error||`Audio recorder ${method} failed.`);return r;}
    start(){if(!['idle','saved','error'].includes(this.state))return this.startPromise||Promise.resolve();
      this.cancelStart=false;this.state='starting';this.seconds=0;this.message='Requesting microphone access…';this.saved=null;this.emit();
      this.startPromise=this._start().finally(()=>{this.startPromise=null;});return this.startPromise;
    }
    async _start(){try{
      if(!this.bridge?.begin)throw Error('Audio recorder bridge is missing. Apply the setup script and fully restart AnchorCast.');
      const info=await this.call('info');if(!info.available)throw Error(info.error||'MP3 encoder is unavailable.');
      await this.call('permission');if(this.cancelStart)throw Error('Recording cancelled.');
      if(!this.mediaDevices?.getUserMedia)throw Error('Microphone recording is unavailable in this window.');
      const microphone=this.getMicrophone();
      this.stream=await this.mediaDevices.getUserMedia({audio:{channelCount:1,echoCancellation:true,noiseSuppression:true,autoGainControl:true,...(microphone&&microphone!=='default'?{deviceId:{exact:microphone}}:{})},video:false});
      if(this.cancelStart)throw Error('Recording cancelled.');
      this.context=new this.AudioContextClass({sampleRate:48000,latencyHint:'interactive'});await this.context.resume();await this.context.audioWorklet.addModule(this.workletUrl);
      if(this.cancelStart)throw Error('Recording cancelled.');
      this.session=await this.call('begin',{sampleRate:this.context.sampleRate});if(this.cancelStart)throw Error('Recording cancelled.');
      this.tail=Promise.resolve();this.writeError=null;this.inflight=0;
      this.node=new this.WorkletClass(this.context,'sermon-pcm-capture',{numberOfInputs:1,numberOfOutputs:1,outputChannelCount:[1]});
      this.node.onprocessorerror=()=>void this.stop('Audio processor failed.');
      this.node.port.onmessage=({data})=>{
        if(data.type==='flushed'){this.resolveFlush?.();return;}
        if(data.type==='fault'){void this.stop(data.message);return;}
        if(data.type!=='pcm'||!this.session)return;
        this.seconds=data.samples/this.context.sampleRate;this.peak=data.peak;this.emit();const sessionId=this.session.id;this.inflight++;
        this.tail=this.tail.then(async()=>{if(this.writeError)return;
          try{await timed(this.call('append',{id:sessionId,sequence:data.sequence,pcm:new Uint8Array(data.pcm)}),12000,'Audio save timed out.');}
          catch(e){this.writeError=e;void this.stop(e.message);}
        }).finally(()=>{this.inflight--;this.node?.port.postMessage({type:'ack'});});
      };
      this.source=this.context.createMediaStreamSource(this.stream);this.source.connect(this.node);this.node.connect(this.context.destination);
      for(const t of this.stream.getAudioTracks())t.onended=()=>void this.stop('Microphone disconnected.');
      this.context.onstatechange=()=>{if(this.state==='recording'&&this.context?.state==='suspended')void this.stop('Audio input was suspended. Recording stopped to avoid a silent gap.');};
      this.state='recording';this.message='Recording locally to MP3';this.emit();
    }catch(e){
      if(this.session){await this.call('finish',{id:this.session.id,interrupted:true}).catch(()=>{});this.session=null;}
      await this.cleanup();this.state=this.cancelStart?'idle':'error';this.message=this.cancelStart?'Recording cancelled.':this.friendlyError(e);this.emit();
    }}
    friendlyError(e){return ({NotAllowedError:'Microphone access denied. Allow AnchorCast microphone access in your operating system settings.',NotFoundError:'No microphone was found. Connect an input device and try again.',OverconstrainedError:'The selected microphone is unavailable. Choose another input in Settings → Audio & AI.',NotReadableError:'The microphone is in use or unavailable. Check the audio device and try again.'})[e.name]||e.message||String(e);}
    pause(){if(!['recording','paused'].includes(this.state))return;const paused=this.state==='recording';this.node.port.postMessage({type:'pause',value:paused});this.state=paused?'paused':'recording';this.peak=0;this.message=paused?'Paused — audio is not added to the file':'Recording locally to MP3';this.emit();}
    stop(reason=''){
      if(this.state==='starting'){this.cancelStart=true;this.message='Cancelling microphone startup…';this.emit();return this.startPromise||Promise.resolve();}
      if(this.stopPromise)return this.stopPromise;if(!this.session)return Promise.resolve();
      this.state='stopping';this.message='Finishing and saving MP3…';this.emit();this.stopPromise=this._stop(reason).finally(()=>{this.stopPromise=null;});return this.stopPromise;
    }
    async _stop(reason){let error=reason;try{
      if(this.node){const flushed=new Promise(yes=>{this.resolveFlush=yes;});this.node.port.postMessage({type:'stop'});await timed(flushed,2000,'Capture flush timed out.').catch(e=>{error ||= e.message;});}
      try{this.source?.disconnect();}catch(_){}for(const t of this.stream?.getTracks()||[]){t.onended=null;t.stop();}
      await timed(this.tail||Promise.resolve(),15000,'Audio save timed out.').catch(e=>{error ||= e.message;});error ||= this.writeError?.message||'';
      const result=await timed(this.call('finish',{id:this.session.id,interrupted:!!error}),30000,'MP3 finalization timed out. Open Recordings for partial audio.');
      this.saved=result.saved?result:null;this.state=result.saved?'saved':'idle';
      this.message=error?error+(result.saved?' Available audio saved as an interrupted MP3.':''):result.error|| (result.saved?`Saved: ${result.name}`:'No audio was captured.');
    }catch(e){this.state='error';this.message=this.friendlyError(e);}finally{this.session=null;this.resolveFlush=null;await this.cleanup();this.peak=0;this.emit();}}
    async cleanup(){try{this.source?.disconnect();}catch(_){}try{this.node?.disconnect();}catch(_){}
      for(const t of this.stream?.getTracks()||[]){t.onended=null;t.stop();}if(this.context){this.context.onstatechange=null;await this.context.close().catch(()=>{});}
      this.source=null;this.node=null;this.context=null;this.stream=null;
    }
  }
  root.SermonAudioRecorder=SermonAudioRecorder;if(typeof module!=='undefined'&&module.exports)module.exports={SermonAudioRecorder,timed};
})(typeof window!=='undefined'?window:globalThis);

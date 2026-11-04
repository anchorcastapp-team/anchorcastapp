/* Bounded PCM capture. Encoder acknowledgements prevent unbounded buffering.
   Outputs are always silent: this recorder never monitors the mic to speakers. */
class SermonPcmCapture extends AudioWorkletProcessor {
  constructor(){super();this.buffer=new Int16Array(8192);this.used=0;this.pending=0;this.active=true;this.paused=false;this.sequence=0;this.samples=0;this.peak=0;
    this.port.onmessage=({data})=>{
      if(data.type==='ack')this.pending=Math.max(0,this.pending-1);
      if(data.type==='pause'){this.paused=!!data.value;this.flush();this.port.postMessage({type:'paused',value:this.paused});}
      if(data.type==='stop'){this.active=false;this.flush();this.port.postMessage({type:'flushed',samples:this.samples});}
    };
  }
  flush(){if(!this.used)return;const pcm=this.buffer.slice(0,this.used);this.samples+=this.used;this.pending++;
    this.port.postMessage({type:'pcm',sequence:this.sequence++,pcm:pcm.buffer,samples:this.samples,peak:this.peak},[pcm.buffer]);this.used=0;this.peak=0;
  }
  process(inputs,outputs){
    for(const output of outputs)for(const channel of output)channel.fill(0);
    if(!this.active||this.paused)return true;
    if(this.pending>=32){this.active=false;this.port.postMessage({type:'fault',message:'Audio could not be saved fast enough. Recording stopped; review the saved partial audio.'});return true;}
    const channels=inputs[0];if(!channels?.length)return true;
    for(let i=0;i<channels[0].length;i++){let sample=0;for(const channel of channels)sample+=channel[i]||0;
      sample=Math.max(-1,Math.min(1,sample/channels.length));this.peak=Math.max(this.peak,Math.abs(sample));
      this.buffer[this.used++]=sample<0?Math.round(sample*32768):Math.round(sample*32767);if(this.used===this.buffer.length)this.flush();
    }return true;
  }
}
registerProcessor('sermon-pcm-capture',SermonPcmCapture);

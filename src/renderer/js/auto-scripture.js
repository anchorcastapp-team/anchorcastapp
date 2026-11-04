/* A bounded, cancellable presentation scheduler. Detection remains independent
 * of Manual/Auto. No old backlog is replayed when changing mode or going live. */
(function(root){
  'use strict';
  function create(options){
    const now=options.now||Date.now,set=options.setTimeout||setTimeout,clear=options.clearTimeout||clearTimeout;
    let timer=null,epoch=0,lastPresented=0,manualCutoff=0,lastSourceStart=-Infinity,inFlight=false;
    const pending=new Map(),shown=new Set();
    const log=(code,d={},detail='')=>options.onDecision?.(code,{...d,detail});
    const verdict=d=>root.DetectionPolicy.eligible(d,options.getState().settings,now());
    const presentationKey=d=>`${d.discussionId||d.passageId||d.eventKey}:${d.ref}`;
    function cancel(reason='cancelled',manual=false){epoch++;clear(timer);timer=null;for(const d of pending.values())log('auto-cancelled',d,reason);pending.clear();if(manual)manualCutoff=now();}
    function allowed(d){const s=options.getState();if(!s.auto)return 'manual-mode';if(!s.live)return 'not-live';if((d.sourceTime||0)<=manualCutoff)return 'operator-override';if(d.announcedPassage && Number(d.verse)!==d.passageVerses?.[0])return 'range-member-awaiting-reading';if(shown.has(presentationKey(d)))return 'already-presented';if(Number.isFinite(d.sourceStart)&&d.sourceStart<lastSourceStart)return 'older-than-current-output';const v=verdict(d);return v.ok?'':v.reason;}
    function schedule(){
      clear(timer);timer=null;if(!pending.size||inFlight)return;
      const wait=Math.max(0,lastPresented+1800-now(),Math.min(...[...pending.values()].map(d=>d._readyAt))-now());
      timer=set(()=>{timer=null;void drain();},wait);
    }
    async function drain(){
      if(inFlight)return;
      for(const [key,d]of pending){const reason=allowed(d);if(reason){pending.delete(key);log('auto-held',d,reason);}}
      if(!pending.size)return;
      // Newer spoken evidence outranks old network results. For a spoken list /
      // range sharing one occurrence, start at its first verse, never its last.
      const candidates=[...pending.values()].filter(d=>d._readyAt<=now()).sort((a,b)=>(b.sourceStart||0)-(a.sourceStart||0)||(a.verse-b.verse));
      if(!candidates.length){schedule();return;}
      const d=candidates[0];for(const [k,other]of pending){if(other!==d)log('auto-held',other,'superseded-by-current-passage');pending.delete(k);}
      const token=epoch,reason=allowed(d);if(reason){log('auto-held',d,reason);return;}
      inFlight=true;
      try{
        log('auto-present-requested',d);
        // present() must call this guard at the final commit boundary if it awaits.
        const result=await options.present(d,()=>token===epoch&&!allowed(d));
        if(token!==epoch){log('auto-cancelled',d,'mode-or-operator-changed');return;}
        if(result===false||result?.success===false){log('auto-present-failed',d,result?.error||'Presentation was not committed');return;}
        shown.add(presentationKey(d));if(shown.size>2000)shown.delete(shown.values().next().value);lastPresented=now();lastSourceStart=Math.max(lastSourceStart,d.sourceStart??-Infinity);
        log('auto-present-dispatched',d,'Sent to live rendering; not a physical-screen acknowledgement');
      }catch(e){log('auto-present-failed',d,e.message);}finally{inFlight=false;schedule();}
    }
    function offer(d){
      const reason=allowed(d);if(reason){log('auto-held',d,reason);return false;}
      const queueKey=presentationKey(d);
      const old=pending.get(queueKey);pending.set(queueKey,{...d,_readyAt:old?old._readyAt:now()+650});
      if(pending.size>8){const oldest=[...pending.values()].sort((a,b)=>(a.sourceStart||0)-(b.sourceStart||0))[0];pending.delete(presentationKey(oldest));log('auto-held',oldest,'queue-capacity');}
      log('auto-queued',d);schedule();return true;
    }
    return {offer,cancel,reset(){cancel('new-transcript');shown.clear();lastPresented=0;manualCutoff=0;lastSourceStart=-Infinity;},pendingCount:()=>pending.size,inspect:()=>({epoch,lastPresented,manualCutoff,lastSourceStart,pending:pending.size})};
  }
  root.AutoScripture={create};if(typeof module!=='undefined'&&module.exports)module.exports={create};
})(typeof window!=='undefined'?window:globalThis);

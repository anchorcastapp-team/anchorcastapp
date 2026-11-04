/* Shared confidence, evidence and response rules. Scores are heuristic match
 * scores, not calibrated probabilities. Completeness is a separate requirement. */
(function(root){
  'use strict';
  const DEFAULT_THRESHOLD=.75;
  const STOP=new Set('the and of to in a is that for it with was on are be this have from or had by but not you all can his one our out they we were which will an do how if my no he me us so up him has may who did get let say says said been come into made than them then what when your also just more some only i shall unto thy thee thou thine ye upon hath doth art am as at nor right know bible scripture amen hallelujah both god lord jesus christ holy spirit power therefore because even very'.split(' '));
  function threshold(settings={}){const n=settings.confidenceThreshold==null||settings.confidenceThreshold===''?NaN:Number(settings.confidenceThreshold);return Number.isFinite(n)&&n>=0&&n<=1?n:DEFAULT_THRESHOLD;}
  function words(text,db){return (db?.normalizeForSearch?.(String(text||''))||String(text||'').toLowerCase().replace(/[^a-z0-9\s]/g,' ').replace(/\s+/g,' ').trim()).split(/\s+/).filter(Boolean);}
  function quoteSupport(source,verse,db){
    const a=words(source,db),b=words(verse,db);let longest=0,bestDistinct=0;
    for(let i=0;i<a.length;i++)for(let j=0;j<b.length;j++){
      let n=0;while(i+n<a.length&&j+n<b.length&&a[i+n]===b[j+n])n++;
      if(n>=4){const d=new Set(a.slice(i,i+n).filter(w=>w.length>2&&!STOP.has(w))).size;if(n>longest){longest=n;bestDistinct=d;}else if(n===longest)bestDistinct=Math.max(bestDistinct,d);}
    }
    const qa=a.filter(w=>w.length>2&&!STOP.has(w)),vb=new Set(b.filter(w=>w.length>2&&!STOP.has(w)));
    const shared=[...new Set(qa)].filter(w=>vb.has(w)).length;
    let ordered=0,j=0;for(const w of qa){const at=b.indexOf(w,j);if(at>=0){ordered++;j=at+1;}}
    // Modest variation is allowed only with several distinctive ordered words.
    const supported=(longest>=5&&bestDistinct>=2)||(longest>=4&&shared>=4&&ordered>=4)||(longest>=5&&bestDistinct>=1&&longest/Math.max(1,b.length)>=.5);
    return {supported,longest,distinctive:shared,ordered};
  }
  function eligible(det,settings={},now=Date.now()){
    if(!det)return {ok:false,reason:'no-candidate'};
    if(['rejected','dismissed','superseded'].includes(det.status))return {ok:false,reason:det.status};
    if(!Number.isFinite(det.confidence)||det.confidence<threshold(settings))return {ok:false,reason:'below-threshold'};
    if(!['reference','quotation','learned'].includes(det.confirmation)||det.autoEligible===false)return {ok:false,reason:det.holdReason||'suggestion-only'};
    if(now-Number(det.sourceTime||det.createdAt||now)>20000)return {ok:false,reason:'stale-evidence'};
    return {ok:true,reason:'confirmed'};
  }
  // Stable evidence indicator: presentation freshness is separate from confidence.
  function indicator(det,settings={}){
    if(!det||['rejected','dismissed','superseded'].includes(det.status))return {className:'low',label:'Rejected or superseded reference'};
    if(!Number.isFinite(det.confidence)||det.confidence<threshold(settings))return {className:'low',label:'Below selected confidence threshold'};
    if(!['reference','quotation','learned'].includes(det.confirmation)||det.autoEligible===false)return {className:'med',label:'Unconfirmed suggestion — review first'};
    return {className:'high',label:'Confirmed match above selected threshold; see status for presentation readiness'};
  }
  function failure(stage,message){const e=new Error(message);e.stage=stage;return e;}
  function parseModelArray(input){
    if(typeof input!=='string'||input.length>65536)throw failure('model-json','Model response missing or too large');
    const text=input.replace(/^\uFEFF/,'').trim();
    let value;
    try{value=JSON.parse(text);}catch(_){
      // Extract ONE complete JSON array, respecting escaped quotes and nested
      // brackets. Never repair a truncated object or invent missing values.
      const start=text.indexOf('[');if(start<0)throw failure('model-json','No complete JSON array in model response');
      let quoted=false,escaped=false,depth=0,end=-1;
      for(let i=start;i<text.length;i++){const c=text[i];if(escaped){escaped=false;continue;}if(quoted&&c==='\\'){escaped=true;continue;}if(c==='"'){quoted=!quoted;continue;}if(quoted)continue;if(c==='[')depth++;if(c===']'&&--depth===0){end=i+1;break;}}
      if(end<0)throw failure('model-json','Truncated JSON array; no results accepted');
      if(/[\[\]{}]/.test(text.slice(0,start))||/[\[\]{}]/.test(text.slice(end)))throw failure('model-json','Multiple or wrapped JSON values; no results accepted');
      try{value=JSON.parse(text.slice(start,end));}catch(_){throw failure('model-json','Invalid model JSON; no results accepted');}
    }
    if(!Array.isArray(value)||value.length>30)throw failure('model-schema','Expected at most 30 detection objects');
    const items=[],rejected=[];
    for(const item of value){
      if(!item||typeof item!=='object'||Array.isArray(item)||typeof item.ref!=='string'||item.ref.length>100||typeof item.evidence!=='string'||item.evidence.length>3000||typeof item.confidence!=='number'||!Number.isFinite(item.confidence)||item.confidence<0||item.confidence>1){rejected.push('invalid-item-schema');continue;}
      items.push({ref:item.ref,confidence:item.confidence,evidence:item.evidence,type:typeof item.type==='string'?item.type:'quote'});
    }
    return {items,rejected};
  }
  function modelText(response){
    if(response?.stop_reason==='max_tokens')throw failure('model-truncated','Model response reached its output limit; no partial detections accepted');
    const parts=response?.content?.filter(c=>c.type==='text'||(!c.type&&typeof c.text==='string')).map(c=>c.text);
    if(!parts?.length)throw failure('response-schema','No text content in provider response');return parts.join('\n');
  }
  const api={threshold,eligible,indicator,quoteSupport,parseModelArray,modelText,failure,DEFAULT_THRESHOLD};
  root.DetectionPolicy=api;if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(typeof window!=='undefined'?window:globalThis);

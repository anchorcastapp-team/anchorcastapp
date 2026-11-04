/* AnchorCast 1.7.2 — reference grammar. No model calls, transcript rewriting,
 * verse-1 defaults, or corrections based on a preferred Bible passage.
 * Marked numbers can be dictated digit-by-digit; informal pairs stay pairs.
 */
(function (root) {
  'use strict';
  const UNITS='one|two|three|four|five|six|seven|eight|nine';
  const TEENS='ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen';
  const TENS='twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety';
  const ORDINALS='first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth';
  const SMALL=`(?:(?:${TENS})(?:[ -]+(?:${UNITS}))?|${TEENS}|${UNITS})`;
  const NUMBER=`(?:\\d{1,3}|(?:one[ -]+)?hundred(?:[ -]+(?:and[ -]+)?${SMALL})?|${SMALL}|${ORDINALS})`;
  const VALUE={zero:0,oh:0,one:1,two:2,three:3,four:4,five:5,six:6,seven:7,eight:8,nine:9,ten:10,eleven:11,twelve:12,thirteen:13,fourteen:14,fifteen:15,sixteen:16,seventeen:17,eighteen:18,nineteen:19,twenty:20,thirty:30,forty:40,fifty:50,sixty:60,seventy:70,eighty:80,ninety:90,first:1,second:2,third:3,fourth:4,fifth:5,sixth:6,seventh:7,eighth:8,ninth:9,tenth:10};
  const DIGIT=`(?:[0-9]|zero|oh|${UNITS})`;
  const digitsRE=new RegExp(`^(${DIGIT}(?:[ -]+${DIGIT}){1,})(?![a-z0-9])`,'i');
  const numberRE=new RegExp(`^(${NUMBER})(?![a-z0-9])`,'i');
  const escapeRE=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  function wordsToNumber(s){
    const v=String(s||'').toLowerCase().trim();
    if(/^\d{1,3}$/.test(v))return +v;
    if(!(new RegExp(`^(?:${NUMBER})$`,'i')).test(v))return null;
    let n=0;for(const w of v.split(/[ -]+/)){if(w==='and')continue;if(w==='hundred')n=(n||1)*100;else if(w in VALUE)n+=VALUE[w];else return null;}return n||null;
  }
  function readNumber(text,offset=0){
    const start=offset+/^\s*/.exec(text.slice(offset))[0].length,m=numberRE.exec(text.slice(start));
    return m?{value:wordsToNumber(m[1]),start,end:start+m[1].length,raw:m[1]}:null;
  }
  function readMarkedNumber(text,offset=0){
    const start=offset+/^\s*/.exec(text.slice(offset))[0].length,m=digitsRE.exec(text.slice(start));
    if(m && !/\d-\d/.test(m[1])){const ds=m[1].toLowerCase().split(/[ -]+/).map(w=>/^\d$/.test(w)?w:String(VALUE[w]));
      return {value:ds.length<=3?Number(ds.join('')):null,start,end:start+m[1].length,raw:m[1],digits:true,invalid:ds.length>3};}
    return readNumber(text,offset);
  }
  const measurement=/^\s+(?:people|persons|children|members|years?|months?|days?|hours?|minutes?|seconds?|dollars?|percent|times|cars?|items?)\b/i;
  function listAfter(text,first,offset,marked=false){
    const values=[first];let end=offset,incomplete=false;
    for(let count=0;count<30;count++){
      const m=/^[\s.]*([,]|and\b|to\b|through\b|[-–—])\s*(?:(?:verses?|v)\.?\s+)?/i.exec(text.slice(end));
      if(!m)break;
      const num=(marked?readMarkedNumber:readNumber)(text,end+m[0].length);
      if(!num){if(!text.slice(end+m[0].length).trim() && m[1]!==',')incomplete=true;break;}
      if(num.invalid||num.value<1||measurement.test(text.slice(num.end)))break;
      const from=values.at(-1),to=num.value;
      if(/^(?:to|through|[-–—])$/i.test(m[1])){if(to<from||to-from>50){incomplete=true;break;}for(let v=from+1;v<=to;v++)values.push(v);}
      else if(!values.includes(to))values.push(to);
      end=num.end;
    }
    return {verses:values,end,incomplete};
  }
  function correctionAt(text,end){return /^[\s,.;]*(?:i mean|sorry(?:\s+i mean)?|rather|or(?:\s+or)?)\s*[,;:]?\s*(?:(from\s+)?(chapter|chap|ch|verse|verses)\.?\s*)?/i.exec(text.slice(end));}
  function correctedList(text,verse,end,marked=true){
    let list=listAfter(text,verse,end,marked),corrected=false;
    for(let i=0;i<3;i++){
      const c=correctionAt(text,list.end);if(!c||c[2]?.startsWith('ch'))break;
      const n=readMarkedNumber(text,list.end+c[0].length);if(!n||!n.value||n.invalid)break;
      list=listAfter(text,n.value,n.end,true);corrected=true;
    }
    return {...list,corrected};
  }
  function create(db){
    const aliases=new Map();
    for(const book of db.ALL_BOOKS||[]){aliases.set(book.toLowerCase(),book);const m=/^([123]) (.+)$/.exec(book);if(m){aliases.set(['','first','second','third'][+m[1]]+' '+m[2].toLowerCase(),book);aliases.set(m[1]+['','st','nd','rd'][+m[1]]+' '+m[2].toLowerCase(),book);aliases.set(m[1]+m[2].toLowerCase(),book);}}
    aliases.set('psalm','Psalms');aliases.set('hebrew','Hebrews');aliases.set('song of songs','Song of Solomon');
    for(const [alias,idx]of Object.entries(db.referenceAliases||{})){const b=db.indexToBookName(idx);if(b)aliases.set(alias.toLowerCase(),b);}
    const pattern=[...aliases.keys()].sort((a,b)=>b.length-a.length).map(x=>escapeRE(x).replace(/ /g,'\\s+')).join('|');
    const bookRE=new RegExp(`\\b(${pattern})(?![a-z])`,'gi');
    function scan(text){
      text=String(text||'');const books=[...text.matchAll(new RegExp(bookRE.source,'gi'))],out=[];
      for(let i=0;i<books.length;i++){
        const m=books[i],book=aliases.get(m[1].toLowerCase().replace(/\s+/g,' ')),limit=books[i+1]?.index??text.length;
        const base=m.index+m[0].length,segment=text.slice(base,limit);
        const reverse=new RegExp(`\\bchapter\\s+(${NUMBER})[\\s.,]*(?:of|in)\\s+(?:the\\s+)?(?:book\\s+of\\s+)?$`,'i').exec(text.slice(0,m.index));
        const prefix=/^\.?\s*(?:(chapter|chap|ch|chapters)\.?\s*(?:number\s+)?)?/i.exec(segment);
        let first=reverse?{value:wordsToNumber(reverse[1]),end:0,raw:reverse[1]}:readNumber(segment,prefix[0].length);
        const start=reverse?reverse.index:m.index;
        if(!first){if(prefix[1])out.push({book,chapter:null,verses:[],start,end:base+prefix[0].length,incomplete:true,pendingReason:'chapter-number'});continue;}
        // Join digit spelling ONLY when chapter/verse grammar disambiguates it.
        const marked=reverse?first:readMarkedNumber(segment,prefix[0].length);
        if(marked?.digits && (prefix[1]||/^[\s.,;]*(?:verses?\b|:)/i.test(segment.slice(marked.end))))first=marked;
        let ch=first.value,verse=null,end=first.end,restoreContext=false,corrected=false;
        let incomplete=!!first.invalid,pendingReason=incomplete?'ambiguous-number':null;
        // An explicitly stuttered chapter: "chapter five five verse seven".
        if(first.digits&&/^([a-z0-9]+)\s+\1$/i.test(first.raw)&&/^\s+verse\b/i.test(segment.slice(end))){ch=readNumber(first.raw).value;}
        const plural=prefix[1]==='chapters';
        const marker=/^[\s.,;]*(?:(?:and\s+)?(?:verses?|v)\.?\s*(?:number\s+)?)|^\s*[:.]\s*(?=\d)/i.exec(segment.slice(end));
        if(marker&&marker[0]){
          const vn=readMarkedNumber(segment,end+marker[0].length);
          if(vn){verse=vn.value;end=vn.end;incomplete ||=!!vn.invalid;
            if(vn.digits&&/\b(?:zero|oh|0)$/i.test(vn.raw)&&!segment.slice(vn.end).trim()){incomplete=true;pendingReason='unfinished-digit-sequence';}}
          else{end+=marker[0].length;incomplete=true;pendingReason='verse-number';}
        }else if(!plural&&!reverse){
          const vn=/^\s+/.test(segment.slice(end))?readNumber(segment,end):null;
          if(vn){verse=vn.value;end=vn.end;
            const extra=readNumber(segment,end);if(extra){incomplete=true;pendingReason='ambiguous-number';end=extra.end;}}
        }
        let list=verse?listAfter(segment,verse,end,!!marker):{verses:[],end};end=list.end;incomplete ||=!!list.incomplete;
        for(let n=0;n<3;n++){
          const c=correctionAt(segment,end);if(!c)break;
          const one=readMarkedNumber(segment,end+c[0].length);if(!one||one.invalid)break;
          const pair=/^\s*(?::|verse\s+)\s*/i.exec(segment.slice(one.end));
          const two=pair?readMarkedNumber(segment,one.end+pair[0].length):null;
          if(two){ch=one.value;verse=two.value;end=two.end;restoreContext=false;}
          else if(c[2]?.startsWith('verse')){if(!verse)restoreContext=true;verse=one.value;end=one.end;}
          else if(c[2]?.startsWith('ch')){ch=one.value;verse=null;end=one.end;restoreContext=false;}
          else{if(verse)verse=one.value;else ch=one.value;end=one.end;}
          corrected=true;incomplete=false;pendingReason=null;
          list=verse?listAfter(segment,verse,end,true):{verses:[],end};end=list.end;incomplete ||=!!list.incomplete;
        }
        out.push({book,chapter:restoreContext?null:ch,verses:incomplete?[]:list.verses,start,end:base+end,
          chapterOnly:!verse&&!incomplete,explicitChapter:!!prefix[1]||!!reverse,incomplete,pendingReason,
          restoreContext,corrected,source:text.slice(start,base+end)});
      }
      return out;
    }
    return {scan};
  }
  function orphanVerses(text){
    const out=[],re=/\b(?:verses?|v)\b\.?\s*(?:number\s+)?/gi;let m;
    while((m=re.exec(text))){
      // Avoid matching the v in an ordinary word.
      if(m[0].toLowerCase()==='v' && /[a-z]/i.test(text[re.lastIndex]||''))continue;
      const num=readMarkedNumber(text,re.lastIndex);
      if(!num){out.push({verses:[],start:m.index,end:re.lastIndex,incomplete:true,source:text.slice(m.index,re.lastIndex)});continue;}
      const list=correctedList(text,num.value,num.end);
      out.push({...list,verses:num.invalid?[]:list.verses,start:m.index,incomplete:num.invalid||list.incomplete,source:text.slice(m.index,list.end)});re.lastIndex=list.end;
    }return out;
  }
  function continuation(text,lastVerse){
    const m=/^\s*(and|to|through|[-–—])\s+(?:(?:verses?|v)\.?\s+)?/i.exec(text);if(!m)return null;
    const num=readMarkedNumber(text,m[0].length);if(!num||num.invalid)return null;
    const list=listAfter(text,num.value,num.end,true);
    if(!/^[\s.,!;?]*(?:amen[\s.,!;?]*)?$/i.test(text.slice(list.end)))return null;
    if(/^(?:to|through|[-–—])$/i.test(m[1])){if(num.value<=lastVerse||num.value-lastVerse>50)return null;list.verses=[...Array(num.value-lastVerse)].map((_,i)=>lastVerse+i+1).concat(list.verses.slice(1));}
    return {...list,start:m.index,source:text.slice(m.index,list.end)};
  }
  function normalizeSpokenReferences(text){
    // Leave split-digit references intact for the citation parser. Partial global
    // replacements used to turn "Psalm one one nine" into "Psalm 1 one nine".
    return String(text||'')
      .replace(new RegExp(`\\bchapter\\s+(${NUMBER})(?![a-z0-9])\\s*[,.;]?\\s+verse\\s+(${NUMBER})(?![a-z0-9])`,'gi'),(_,ch,v)=>`${wordsToNumber(ch)}:${wordsToNumber(v)}`)
      .replace(new RegExp(`\\bpsalms?\\s+(${NUMBER})(?![a-z0-9])`,'gi'),(whole,ch,offset,all)=>{
        const after=all.slice(offset+whole.length);return readNumber(after)||/^\s+(?:zero|oh)\b/i.test(after)?whole:`Psalm ${wordsToNumber(ch)}`;
      });
  }
  const api={create,wordsToNumber,readNumber,readMarkedNumber,orphanVerses,continuation,normalizeSpokenReferences,NUMBER};
  root.ScriptureReferences=api;if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(typeof window!=='undefined'?window:globalThis);

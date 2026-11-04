"""Production presentation components in Chromium with simulated Electron IPC.
Disk storage uses the real Node store in a separate process for every request.
The media:// protocol is simulated using data URLs; no physical projection.
"""
import argparse,base64,json,re,subprocess,tempfile
from pathlib import Path
from playwright.sync_api import sync_playwright
parser=argparse.ArgumentParser();parser.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1]);args=parser.parse_args()
ROOT=args.root.resolve();OUT=ROOT/'test-results/presentation-persistence';OUT.mkdir(parents=True,exist_ok=True)
checks=[];errors=[]
with tempfile.TemporaryDirectory(prefix='ac-presentation-ui-') as temp:
 def rpc(req):
  p=subprocess.run(['node',str(ROOT/'scripts/presentation-test-rpc.cjs'),temp],input=json.dumps(req),text=True,capture_output=True)
  data=json.loads(p.stdout)
  if 'error' in data:raise Exception(data['error'])
  return data['result']
 rpc({'method':'seed'})
 picture='data:image/png;base64,'+base64.b64encode((ROOT/'scripts/fixtures/presentations/picture.png').read_bytes()).decode()
 def browser_rpc(req):
  data=rpc(req)
  # The real Electron protocol serves these paths. Simulate it for plain Chromium.
  if req['method']=='getPresentationSlides' and data.get('slides'):
   for sl in data['slides']:sl['imagePath']=picture
  if req['method']=='getPresentations':
   for pres in data:pres['thumbnailPath']=picture
  return data
 with sync_playwright() as pw:
  browser=pw.chromium.launch(executable_path='/usr/bin/chromium',headless=True,args=['--no-sandbox'])
  app=(ROOT/'src/renderer/js/app.js').read_text();engine=app[app.index('// ─── PRESENTATIONS ENGINE'):app.index('function runContextSearch(')]
  def check(page,name,expr):
   if not page.evaluate(expr):raise AssertionError(name)
   checks.append(name);print("PASS",name,flush=True)
  def library_page():
   page=browser.new_page(viewport={'width':1400,'height':900});page.on('pageerror',lambda e:errors.append(str(e)))
   page.route('**/*',lambda r:r.abort());page.expose_function('__native',browser_rpc)
   html=(ROOT/'src/renderer/index.html').read_text();html=re.sub(r'<script\b[^>]*>[\s\S]*?</script>','',html,flags=re.I)
   page.set_content(html);page.add_style_tag(content=re.sub(r'^@import[^\n]*','',(ROOT/'src/renderer/css/app.css').read_text(),flags=re.M))
   page.add_script_tag(path=str(ROOT/'src/renderer/js/presentation-data.js'))
   page.add_script_tag(content='''const State={presentations:[],currentPresId:null,currentPresSlides:[],currentPresSlideIdx:0,queue:[],isLive:false};
     window.messages=[];window.editorRequests=[];window.events={};window.failDelete=false;
     const toast=s=>messages.push(s);const renderQueue=()=>{};const cancelAutomaticScripture=()=>{};
     const escapeHtml=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
     const _confirmModal=(m,fn)=>{window.deletion=Promise.resolve().then(fn)};
     window.electronAPI={on:(name,fn)=>events[name]=fn,openPresEditor:d=>editorRequests.push(d),
       getPresentations:()=>__native({method:'getPresentations'}),getCreatedPresentations:()=>__native({method:'getCreatedPresentations'}),
       getPresentationSlides:d=>__native({method:'getPresentationSlides',payload:d}),
       deletePresentation:d=>failDelete?Promise.resolve({success:false,error:'Disk is read-only'}):__native({method:'deletePresentation',payload:d}),
       saveCreatedPresentations:d=>__native({method:'saveCreatedPresentations',payload:d})};
   ''')
   page.add_script_tag(content=engine)
   page.evaluate('''document.getElementById('appBootOverlay')?.remove();document.getElementById('bibleSearchContent')?.style.setProperty('display','none');
     for(const id of ['songLibContent','mediaLibContent','themeLibContent'])document.getElementById(id)?.style.setProperty('display','none');
     document.getElementById('presLibContent').style.display='flex';
     document.getElementById('bookTab').classList.remove('active');document.getElementById('presLibTab').classList.add('active');
     document.getElementById('previewPanel').style.flex='3';document.getElementById('searchPanel').style.flex='1';
     document.querySelector('.app-logo img')?.remove();
     document.getElementById('presNewBtn').addEventListener('click',newPresentation);
     document.getElementById('presSearchInput').addEventListener('input',e=>renderPresLibStrip(e.target.value));
   ''')
   page.evaluate('loadPresentations()');return page
  page=library_page()
  check(page,'Legacy and created decks load from real JSON storage',"State.presentations.length===2 && State.presentations.find(p=>p.id==='101').type==='imported'")
  page.evaluate("selectPresentation('101')")
  check(page,'Imported selection loads both slides after startup',"State.currentPresSlides.length===2 && document.getElementById('presNavPos').textContent==='Slide 1 of 2'")
  page.wait_for_function("document.querySelector('#liveDisplay img')?.naturalWidth>0")
  check(page,'Imported slide visible in preview',"document.querySelector('#liveDisplay img').naturalWidth>0")
  check(page,'Library toolbar has one New; rows have none',"document.querySelectorAll('#presNewBtn').length===1 && ![...document.querySelectorAll('.pres-row')].some(r=>r.textContent.includes('+ New'))")
  page.locator('#presNewBtn').click();check(page,'New toolbar button opens blank editor',"editorRequests.length===1 && editorRequests[0].id===null")
  check(page,'New, Import and Schedule share their toolbar',"document.getElementById('presNewBtn').parentElement===document.getElementById('presImportBtn').parentElement && document.getElementById('presNewBtn').parentElement===document.getElementById('presAddScheduleBtn').parentElement")
  page.evaluate("selectPresentation('202')")
  check(page,'Created legacy title and body render on reopen',"document.getElementById('liveDisplay').textContent.includes('Welcome back') && document.getElementById('liveDisplay').textContent.includes('survives restarting')")
  check(page,'Switching imported to created does not retain stale imported image',"!document.querySelector('#liveDisplay > img')")
  check(page,'Image objects appear in created slide thumbnails',"document.querySelectorAll('#presPreviewStrip img').length===1")
  page.screenshot(path=str(OUT/'restored-created-presentation.png'),full_page=True)
  page.evaluate("previewPresSlide(1)")
  page.wait_for_function("document.querySelector('#liveDisplay img')?.naturalWidth>0")
  check(page,'Stored image displays after reopening a created slide',"document.querySelector('#liveDisplay img').naturalWidth>0")
  # Selection race: the slower first fetch must never overwrite the later deck.
  page.evaluate("window.originalSlides=electronAPI.getPresentationSlides;electronAPI.getPresentationSlides=d=>new Promise(resolve=>window.releaseOld=()=>originalSlides(d).then(resolve));void selectPresentation('101')")
  page.evaluate("selectPresentation('202')");page.evaluate('releaseOld()');page.wait_for_timeout(120)
  check(page,'Slow earlier fetch does not overwrite the newly selected deck',"State.currentPresId==='202' && State.currentPresSlides[0].slide.objects[0].text==='Welcome back'")
  page.evaluate('() => { electronAPI.getPresentationSlides=originalSlides; }')
  page.evaluate("selectPresentation('101')");page.evaluate('failDelete=true;deleteCurrentPresentation();deletion')
  check(page,'Failed deletion keeps the row and explains the save failure',"State.presentations.some(p=>p.id==='101') && messages.at(-1).includes('read-only')")
  page.evaluate('failDelete=false;deleteCurrentPresentation();deletion')
  check(page,'Successful imported deletion removes the row after storage acknowledgment',"!State.presentations.some(p=>p.id==='101')")
  page.close();page=library_page()
  check(page,'Deleted imported presentation stays deleted across a fresh browser session',"State.presentations.length===1 && !State.presentations.some(p=>p.id==='101')")
  page.evaluate("selectPresentation('202');deleteCurrentPresentation();deletion")
  page.close();page=library_page()
  check(page,'Deleted custom presentation stays deleted across a fresh browser session',"State.presentations.length===0")
  # Editor uses production JS and markup, with native invocations routed to disk.
  editor=browser.new_page(viewport={'width':1280,'height':840});editor.on('pageerror',lambda e:errors.append(str(e)))
  editor.route('**/*',lambda r:r.abort());editor.expose_function('__native',browser_rpc)
  html=(ROOT/'src/renderer/presentation-editor.html').read_text();scripts=re.findall(r'<script\b[^>]*>([\s\S]*?)</script>',html,re.I)
  editor.set_content(re.sub(r'<script\b[^>]*>[\s\S]*?</script>','',html,flags=re.I));editor.add_script_tag(path=str(ROOT/'src/renderer/js/presentation-data.js'))
  editor.add_script_tag(content='''window.closedBySave=false;window.notifications=[];window.saveFailure=false;window.deferSave=false;
     window.close=()=>closedBySave=true;window.electronAPI={isWeb:false,on:(n,fn)=>window.editorLoad=fn,
      getPathForFile:f=>window.fixturePath,importThemeBgFile:p=>__native({method:'importThemeBgFile',payload:p}),
      saveCreatedPresentations:d=>saveFailure?Promise.resolve({success:false,error:'Disk full'}):(deferSave?new Promise(resolve=>window.completeSave=()=>__native({method:'saveCreatedPresentations',payload:d}).then(resolve)):__native({method:'saveCreatedPresentations',payload:d})),
      presEditorSaved:p=>notifications.push(p),addPresentationToScheduleFromEditor:p=>window.scheduled=p};''')
  editor.evaluate('(p)=>window.fixturePath=p',str(ROOT/'scripts/fixtures/presentations/picture.png'))
  for script in scripts:
   if script.strip():editor.add_script_tag(content=script)
  editor.evaluate("window.dispatchEvent(new Event('DOMContentLoaded'));editorLoad({id:'303',name:'Editor restart test',slides:[{title:'Saved from editor',body:'Durable text'}]})")
  editor.evaluate('saveFailure=true;savePres()')
  check(editor,'Save failure leaves the editor open with the text intact',"!closedBySave && slides[0].objects[0].text==='Saved from editor'")
  editor.evaluate('saveFailure=false;deferSave=true;void savePres()')
  check(editor,'Editor remains open while save acknowledgment is pending',"!closedBySave && _savingPresentation")
  editor.evaluate('completeSave()');editor.wait_for_function('closedBySave')
  check(editor,'Editor closes only after persistence and sends a saved notification',"notifications.length===1 && notifications[0]._persisted===true")
  check(editor,'Save controls re-enable after completion',"!_savingPresentation")
  editor.evaluate("closedBySave=false;deferSave=false;editorLoad({id:'304',name:'Temporary image deck',slides:[{objects:[],bgImage:'blob:expired'}]});savePres()")
  check(editor,'Old broken blob images cannot silently be saved as durable media',"!closedBySave && !_savingPresentation")
  editor.evaluate("editorLoad({id:'305',name:'Persistent selected image',slides:[{objects:[]}]})")
  editor.locator('#objImgIn').set_input_files(str(ROOT/'scripts/fixtures/presentations/picture.png'))
  editor.wait_for_function("slides[0].objects[0]?.src?.includes('persisted-test-image.png')")
  check(editor,'New image selection is copied into persistent storage, not a blob URL',"!slides[0].objects[0].src.startsWith('blob:')")
  editor.evaluate('savePres()');editor.wait_for_function('closedBySave')
  persisted=rpc({'method':'getCreatedPresentations'})
  assert len(persisted)==2 and any(p['id']=='303' for p in persisted);checks.append('Actual editor save is readable by a separate Node process')
  imagepath=next(p for p in persisted if p['id']=='305')['slides'][0]['objects'][0]['src'];assert Path(imagepath).is_file();checks.append('Saved image asset exists on disk after editor closure')
  page.close();page=library_page();page.evaluate("selectPresentation('303')")
  check(page,'Created deck saved by editor reopens in the main preview',"document.getElementById('liveDisplay').textContent.includes('Saved from editor')")
  if errors:raise AssertionError(errors)
  checks.append('No uncaught browser errors in tested flows')
  browser.close()
OUT.joinpath('results.json').write_text(json.dumps({'passed':len(checks),'checks':checks,'limitations':'Chromium component tests with simulated Electron IPC/media; storage uses fresh Node subprocesses.'},indent=2))
print(json.dumps({'passed':len(checks),'output':str(OUT)},indent=2))

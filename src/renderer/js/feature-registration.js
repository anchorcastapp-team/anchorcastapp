/* Registration-gated network controls. No permission is stored in localStorage.
 * Grey controls stay focusable so mouse/keyboard users can ask how to unlock
 * them. All service operations are also checked in the Electron main process.
 */
(function () {
  'use strict';
  const labels = { remote:'Remote Control', ndi:'External Output (NDI)' };
  let registered = null;
  let checking = null;
  let promptPending = false;

  const style = document.createElement('style');
  style.textContent = `
    [data-registration-feature][data-registration-locked="true"] {
      opacity:.55 !important; filter:grayscale(1); cursor:pointer !important;
    }
    .toggle[data-registration-locked="true"],
    .remote-toggle[data-registration-locked="true"] {
      background:#484b53 !important; border-color:#676b75 !important;
    }
    [data-registration-locked="true"] .toggle-knob { left:3px !important; }
    [data-registration-locked="true"] .remote-toggle-knob { transform:none !important; }
    [data-registration-feature]:focus-visible {
      outline:2px solid #8bcaff !important; outline-offset:4px;
    }
    .registration-feature-note { color:var(--text-sub,#bcc2d1); font-size:12px;
      line-height:1.6; margin:8px 0 14px; }
    .registration-feature-note button { font:inherit; color:var(--gold,#dfbe69);
      background:transparent; border:0; padding:4px 2px; text-decoration:underline; cursor:pointer; }
    #featureRegistrationDialog { max-width:480px; width:calc(100% - 40px);
      max-height:calc(100vh - 40px); overflow:auto; box-sizing:border-box;
      margin:auto; padding:26px; border:1px solid #a88a41; border-radius:14px;
      background:#151821; color:#f3f4f6; box-shadow:0 18px 80px #0009; font-family:inherit; }
    #featureRegistrationDialog::backdrop { background:#0009; }
    #featureRegistrationDialog h2 { font-size:20px; margin:0 0 12px; }
    #featureRegistrationDialog p { font-size:14px; line-height:1.7; margin:0 0 16px; }
    #featureRegistrationDialog .actions { display:flex; justify-content:flex-end; flex-wrap:wrap; gap:10px; }
    #featureRegistrationDialog button { min-height:42px; padding:10px 16px;
      border:1px solid #626b7e; border-radius:7px; background:#252a38;
      color:#fff; cursor:pointer; font:inherit; }
    #featureRegistrationDialog #featureRegistrationOpen { border-color:#c9a84c;
      background:#c9a84c; color:#16130a; font-weight:700; }
    #featureRegistrationDialog button:focus-visible { outline:2px solid #8bcaff; outline-offset:3px; }
  `;
  document.head.appendChild(style);

  function apply(root = document) {
    const locked = registered !== true;
    root.querySelectorAll('[data-registration-feature]').forEach(el => {
      const name = labels[el.dataset.registrationFeature];
      if (!name) return;
      if (el.dataset.registrationTitle === undefined) el.dataset.registrationTitle = el.title || '';
      el.dataset.registrationLocked = String(locked);
      el.setAttribute('aria-disabled', String(locked));
      el.title = locked ? `${name} requires registration. Click to register.` : el.dataset.registrationTitle;
      // Do not set native disabled: the control opens an explanation while locked.
      if (el.tagName === 'BUTTON') el.disabled = false;
      if (el.tagName !== 'BUTTON' && el.tagName !== 'A') {
        if (!el.hasAttribute('role')) el.setAttribute('role', 'button');
        el.tabIndex = 0;
      }
      if (locked && el.getAttribute('role') === 'switch') {
        el.setAttribute('aria-checked','false');
        el.classList.remove('on');
        el.classList.add('off');
      }
    });
    root.querySelectorAll('[data-registration-note]').forEach(el => { el.hidden = !locked; });
  }

  async function readStatus() {
    if (checking) return checking;
    const task = (async () => {
      let timer;
      try {
        if (typeof window.electronAPI?.getRegistrationStatus !== 'function') throw new Error('Missing bridge');
        const status = await Promise.race([
          window.electronAPI.getRegistrationStatus(),
          new Promise((_,reject) => { timer=setTimeout(()=>reject(new Error('Timed out')),5000); }),
        ]);
        registered = status?.registered === true ? true : status?.registered === false ? false : null;
      } catch (_) { registered=null; }
      finally { clearTimeout(timer); }
      apply();
      return registered;
    })();
    checking=task;
    try { return await task; } finally { if (checking===task) checking=null; }
  }

  function prompt(feature, unavailable=false) {
    const name=labels[feature];
    if (!name) return;
    let dialog=document.getElementById('featureRegistrationDialog');
    if (!dialog) {
      dialog=document.createElement('dialog');
      dialog.id='featureRegistrationDialog';
      dialog.setAttribute('aria-labelledby','featureRegistrationTitle');
      dialog.setAttribute('aria-describedby','featureRegistrationMessage');
      dialog.innerHTML=`<h2 id="featureRegistrationTitle">Registration required</h2>
        <p id="featureRegistrationMessage"></p>
        <p>You can continue using the basic features and register later from <strong>Help → Registration</strong>.</p>
        <p id="featureRegistrationStatus" role="status" aria-live="polite"></p>
        <div class="actions"><button type="button" id="featureRegistrationDismiss">Not now</button>
        <button type="button" id="featureRegistrationOpen">Register now</button></div>`;
      document.body.appendChild(dialog);
      dialog.querySelector('#featureRegistrationDismiss').onclick=()=>dialog.close();
      dialog.addEventListener('close',()=>{ if(dialog._returnFocus?.isConnected) dialog._returnFocus.focus(); });
      dialog.addEventListener('keydown',event=>{
        if(event.key!=='Tab') return;
        const buttons=Array.from(dialog.querySelectorAll('button:not(:disabled)'));
        if(!buttons.length) return;
        const index=buttons.indexOf(document.activeElement);
        if((event.shiftKey && index<=0) || (!event.shiftKey && index===buttons.length-1)) {
          event.preventDefault();
          buttons[event.shiftKey ? buttons.length-1 : 0].focus();
        }
      });
      dialog.querySelector('#featureRegistrationOpen').onclick=async()=>{
        const button=dialog.querySelector('#featureRegistrationOpen');
        const message=dialog.querySelector('#featureRegistrationStatus');
        button.disabled=true;
        let timer;
        try {
          if(typeof window.electronAPI?.openRegistrationWindow !== 'function') {
            throw new Error('Open Help → Registration in the desktop app. Restart after applying the complete patch.');
          }
          const result=await Promise.race([
            window.electronAPI.openRegistrationWindow(),
            new Promise((_,reject)=>{ timer=setTimeout(()=>reject(new Error('Registration did not open. Use Help → Registration.')),8000); }),
          ]);
          if(result?.success !== true) throw new Error(result?.error || 'Use Help → Registration to register.');
          dialog.close();
        } catch(error) { message.textContent=error.message; }
        finally { clearTimeout(timer); button.disabled=false; }
      };
    }
    dialog.querySelector('#featureRegistrationTitle').textContent=unavailable ? 'Unable to check registration' : 'Registration required';
    dialog.querySelector('#featureRegistrationMessage').textContent=unavailable
      ? `AnchorCast could not verify this device’s registration. ${name} remains off. Restart the desktop app and try again, or register from Help → Registration.`
      : `${name} requires registration. Please register AnchorCast and complete activation before using this feature.`;
    dialog.querySelector('#featureRegistrationStatus').textContent='';
    if(!dialog.open) {
      dialog._returnFocus=document.activeElement;
      dialog.showModal();
      dialog.querySelector('#featureRegistrationOpen').focus();
    }
  }

  async function ensure(feature) {
    const status=await readStatus();
    if(status===true) return true;
    prompt(feature,status===null);
    return false;
  }
  function handleBlocked(feature,result) {
    if(!result?.blocked && !result?.registrationRequired) return false;
    registered=result.code==='REGISTRATION_UNAVAILABLE' ? null : false;
    apply();
    prompt(feature,registered===null);
    return true;
  }

  // Capture runs before inline handlers, preventing even a transient ON state.
  document.addEventListener('click',event=>{
    const register=event.target.closest?.('[data-registration-prompt]');
    if(register) { event.preventDefault(); prompt(register.dataset.registrationPrompt,registered===null); return; }
    const control=event.target.closest?.('[data-registration-feature]');
    if(!control || registered===true) return;
    event.preventDefault(); event.stopImmediatePropagation();
    if(promptPending) return;
    promptPending=true;
    ensure(control.dataset.registrationFeature).finally(()=>{ promptPending=false; });
    // A newly activated user clicks again deliberately; never replay an enable.
  },true);
  document.addEventListener('keydown',event=>{
    const control=event.target.closest?.('[data-registration-feature]');
    if(!control || !['Enter',' '].includes(event.key)) return;
    event.preventDefault(); event.stopPropagation();
    control.click();
  },true);

  window.RegistrationAccess=Object.freeze({ refresh:readStatus, ensure, prompt, handleBlocked, apply,
    isGranted:()=>registered===true });
  function init() {
    apply();
    readStatus();
    window.addEventListener('focus',()=>readStatus());
    document.addEventListener('visibilitychange',()=>{ if(!document.hidden) readStatus(); });
    window.electronAPI?.on?.('registration-complete',async()=>{
      await readStatus();
      const dialog=document.getElementById('featureRegistrationDialog');
      if(registered===true && dialog?.open) dialog.close();
    });
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init,{once:true});
  else init();
})();

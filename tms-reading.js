(function () {
  'use strict';
  const sizes={normal:{label:'Normal',scale:1},large:{label:'Grande',scale:1.2},extra:{label:'Muy grande',scale:1.4}};
  const valid=value=>Object.prototype.hasOwnProperty.call(sizes,value);
  let currentKey='',size='large',frame=0;
  function preferenceKey(){const profile=typeof APP!=='undefined'?APP.currentUserProfile:null;return 'tms_reading_size_v1:'+String(profile?.authUserId||profile?.userId||profile?.email||'guest').toLowerCase();}
  function load(){const key=preferenceKey();if(key===currentKey)return;currentKey=key;try{const saved=localStorage.getItem(key);size=valid(saved)?saved:'large';}catch(error){size='large';}}
  function syncControls(){
    const modal=document.getElementById('readingSizeDialog');
    if(modal)modal.querySelectorAll('[data-reading-size]').forEach(button=>button.setAttribute('aria-pressed',String(button.dataset.readingSize===size)));
    document.querySelectorAll('[data-open-reading]').forEach(button=>button.title='Tamaño de letra: '+sizes[size].label);
  }
  function resizeText(){
    frame=0;load();
    const body=document.body;body.dataset.readingActive='false';body.dataset.readingSize=size;
    const elements=[...document.querySelectorAll('#mainLayout *,dialog:not(#readingSizeDialog) *,.inv-modal *')].filter(el=>{
      if(['SCRIPT','STYLE','SVG','PATH','CANVAS'].includes(el.tagName)||el.closest('svg,#readingSizeDialog'))return false;
      return (el.matches('input,select,textarea')||[...el.childNodes].some(node=>node.nodeType===3&&node.textContent.trim()))&&el.getClientRects().length;
    });
    // Read all base sizes before applying the scale, so nested text is never scaled twice.
    const bases=elements.map(el=>{const style=getComputedStyle(el);return {el,font:parseFloat(style.fontSize),line:parseFloat(style.lineHeight)};});
    for(const {el,font,line} of bases){if(!Number.isFinite(font)||font<=0){el.removeAttribute('data-tms-reading');el.style.removeProperty('--reading-base-font');el.style.removeProperty('--reading-base-line');continue;}el.dataset.tmsReading='';el.style.setProperty('--reading-base-font',font+'px');el.style.setProperty('--reading-base-line',(Number.isFinite(line)?line:font*1.4)+'px');}
    body.style.setProperty('--reading-scale',sizes[size].scale);body.dataset.readingActive=String(size!=='normal');syncControls();
  }
  function schedule(){if(!frame)frame=requestAnimationFrame(resizeText);}
  function choose(next){if(!valid(next))return;size=next;currentKey=preferenceKey();try{localStorage.setItem(currentKey,size);}catch(error){}resizeText();}
  function open(){
    load();let modal=document.getElementById('readingSizeDialog');
    if(!modal){modal=document.createElement('dialog');modal.id='readingSizeDialog';modal.setAttribute('aria-labelledby','readingSizeTitle');
      modal.innerHTML='<header><h2 id="readingSizeTitle">Tamaño de letra</h2><button type="button" data-reading-close aria-label="Cerrar tamaño de letra">×</button></header><div class="reading-size-options">'+Object.entries(sizes).map(([id,entry])=>'<button type="button" data-reading-size="'+id+'"><strong class="reading-sample reading-sample-'+id+'">Aa</strong><span>'+entry.label+'</span></button>').join('')+'</div><p>Se recuerda para tu usuario en este navegador.</p>';
      document.body.appendChild(modal);modal.addEventListener('click',event=>{const button=event.target.closest('button[data-reading-size]');if(button)choose(button.dataset.readingSize);else if(event.target.closest('[data-reading-close]'))modal.close();else if(event.target===modal){const r=modal.getBoundingClientRect();if(event.clientX<r.left||event.clientX>r.right||event.clientY<r.top||event.clientY>r.bottom)modal.close();}});
    }
    syncControls();if(!modal.open)modal.showModal();
  }
  function ensureButtons(){
    const header=document.getElementById('appHeader');
    if(header&&!document.getElementById('readingSizeButton')){const button=document.createElement('button');button.id='readingSizeButton';button.dataset.openReading='';button.type='button';button.textContent='Aa';button.setAttribute('aria-label','Cambiar tamaño de letra');button.setAttribute('aria-haspopup','dialog');button.setAttribute('aria-controls','readingSizeDialog');header.insertBefore(button,document.getElementById('userBadge'));}
    const bar=document.querySelector('.mobile-module-bar');
    if(bar&&!document.getElementById('mobileReadingSizeButton')){const button=document.createElement('button');button.id='mobileReadingSizeButton';button.dataset.openReading='';button.type='button';button.textContent='Aa';button.setAttribute('aria-label','Cambiar tamaño de letra');button.setAttribute('aria-haspopup','dialog');button.setAttribute('aria-controls','readingSizeDialog');bar.insertBefore(button,document.getElementById('mobileModulesButton'));}
  }
  function init(){
    ensureButtons();load();resizeText();
    document.addEventListener('click',event=>{if(event.target.closest('[data-open-reading]'))open();});
    window.addEventListener('tms-permissions-changed',()=>{ensureButtons();schedule();});window.addEventListener('resize',schedule);
    window.addEventListener('toggle',schedule,true);window.addEventListener('storage',event=>{if(event.key===preferenceKey()){currentKey='';schedule();}});
    new MutationObserver(records=>{if(records.some(record=>{const el=record.target.nodeType===1?record.target:record.target.parentElement;return el?.closest('#mainLayout,dialog:not(#readingSizeDialog),.inv-modal');}))schedule();}).observe(document.body,{subtree:true,childList:true,characterData:true});
    const oldLogout=window.cerrarSesion;window.cerrarSesion=function(...args){document.getElementById('readingSizeDialog')?.close();return oldLogout.apply(this,args);};
  }
  window.addEventListener('DOMContentLoaded',init);
})();

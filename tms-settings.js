(function(){
 'use strict';
 const panels=[['clientes','Clientes'],['rutas','Rutas'],['costos','Costos'],['accesos','Usuarios y roles'],['apariencia','Apariencia']];
 function select(panel){
  const view=document.getElementById('configClientes');
  if(!view||!panels.some(item=>item[0]===panel))return;
  view.dataset.settingsPanel=panel;
  document.querySelectorAll('[data-settings-tab]').forEach(button=>{const active=button.dataset.settingsTab===panel;button.setAttribute('aria-selected',String(active));button.tabIndex=active?0:-1;});
 }
 window.addEventListener('DOMContentLoaded',()=>{
  const view=document.getElementById('configClientes'),title=view?.querySelector('.view-title'),search=document.getElementById('configSearch');
  if(!view||!title||!search)return;
  search.closest('.card').classList.add('settings-client-toolbar');
  const toolbar=search.closest('.view-toolbar');
  const more=document.createElement('details');more.className='settings-more';more.innerHTML='<summary>Más opciones</summary><div></div>';
  const moreContent=more.querySelector('div');
  ['configFile'].forEach(id=>{const input=document.getElementById(id);if(input)moreContent.appendChild(input);});
  [...toolbar.querySelectorAll('button[onclick]')].forEach(button=>moreContent.appendChild(button));
  document.getElementById('settingsNewClient').after(more);

  const nav=document.createElement('nav');nav.id='settingsNavigation';nav.setAttribute('aria-label','Secciones de configuración');nav.setAttribute('role','tablist');
  nav.innerHTML=panels.map(([id,label])=>`<button type="button" role="tab" data-settings-tab="${id}" aria-selected="false">${label}</button>`).join('');title.after(nav);
  nav.addEventListener('click',event=>{const button=event.target.closest('[data-settings-tab]');if(button)select(button.dataset.settingsTab);});
  nav.addEventListener('keydown',event=>{if(!['ArrowLeft','ArrowRight','Home','End'].includes(event.key))return;event.preventDefault();const current=panels.findIndex(([id])=>id===view.dataset.settingsPanel),index=event.key==='Home'?0:event.key==='End'?panels.length-1:(current+(event.key==='ArrowRight'?1:-1)+panels.length)%panels.length;select(panels[index][0]);nav.querySelector(`[data-settings-tab="${panels[index][0]}"]`).focus();});
  document.getElementById('settingsNewClient').addEventListener('click',()=>{const form=document.getElementById('auditCreateClient');if(!form)return;const open=form.classList.toggle('settings-create-open');document.getElementById('settingsNewClient').setAttribute('aria-expanded',String(open));if(open)document.getElementById('auditClientCode')?.focus();});
  document.getElementById('settingsNewClient').setAttribute('aria-controls','auditCreateClient');document.getElementById('settingsNewClient').setAttribute('aria-expanded','false');
  const updateAccess=()=>{document.getElementById('settingsNewClient').disabled=!window.TMSAuditCore?.hasPermission('configuracion','editar');};
  window.addEventListener('tms-permissions-changed',updateAccess);updateAccess();
  select('clientes');
 });
})();

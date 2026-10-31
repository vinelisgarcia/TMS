(function () {
  'use strict';
  const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const paths = {
    importar:'M12 3v12m-4-4 4 4 4-4M4 16v5h16v-5',
    dashboard:'M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h7v7h-7z',
    calendario:'M4 5h16v16H4zM4 10h16M8 3v4m8-4v4M8 14h2m4 0h2m-8 3h2',
    inventario:'m3 7 9-4 9 4v10l-9 4-9-4zM3 7l9 4 9-4m-9 4v10M7 5l10 5',
    comercial:'M4 4h16v16H4zM8 8h8M8 12h8M8 16h5',
    materialesTransito:'M3 6h11v12H3zM14 10h4l3 4v4h-7M7 18v2m10-2v2',
    prioridades:'m5 7 2 2 4-4m2 2h7M5 15l2 2 4-4m2 2h7',
    incidenciasDespacho:'m12 3 10 18H2zM12 9v5m0 3v1',
    configClientes:'M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8M12 2v3m0 14v3M2 12h3m14 0h3M5 5l2 2m10 10 2 2M5 19l2-2M17 7l2-2',
    solicitudesAlmacen:'M5 4h14v17H5zM9 4V2h6v2M9 9h6M9 13h6M9 17h4',
    menu:'M4 5h16M4 12h16M4 19h16',
    undo:'M8 5 3 10l5 5M3 10h10a7 7 0 0 1 7 7',
    logout:'M10 3H4v18h6m4-14 5 5-5 5m-6-5h13'
  };
  const icon = name => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="${paths[name] || paths.comercial}"/></svg>`;
  const labels = {importar:'Importar SAP',dashboard:'Dashboard',calendario:'Calendario',inventario:'Inventario',comercial:'Comercial',materialesTransito:'En tránsito',prioridades:'Prioridades',incidenciasDespacho:'Incidencias',configClientes:'Configuración',solicitudesAlmacen:'Solicitudes',auditoriaJamaica:'Auditoría Jamaica'};
  let onlyDeliveries = true, frame = 0, calendarToolsMobile = null;
  const activeId = () => document.querySelector('.view.active')?.id || '';
  function modules() {
    return [...(document.getElementById('mobileModuleSelect')?.options || [])].filter(option => {
      const nav = document.getElementById('nav-'+option.value);
      return !option.hidden && !option.disabled && (!nav || nav.style.display !== 'none');
    }).map(option => ({id:option.value,label:labels[option.value] || option.textContent.trim()}));
  }
  function updateNavigation() {
    const main = document.getElementById('mainLayout'), mount = document.getElementById('mobileModuleNav');
    if (!main || !mount) return;
    let trigger = document.getElementById('mobileModulesButton');
    if (!trigger) {
      const row = document.createElement('div'); row.className = 'mobile-module-bar';
      row.innerHTML = '<strong id="mobileCurrentModule" role="heading" aria-level="1"></strong><button type="button" id="mobileModulesButton" aria-haspopup="dialog" aria-controls="mobileModulesDialog" aria-expanded="false">'+icon('menu')+'<span>Módulos</span></button>';
      mount.appendChild(row); trigger = row.querySelector('button'); trigger.addEventListener('click', openModules);
      const bottom = document.createElement('nav'); bottom.id = 'mobileQuickNav'; bottom.setAttribute('aria-label','Accesos rápidos'); main.appendChild(bottom);
      bottom.addEventListener('click',event=>{const button=event.target.closest('[data-mobile-module]');if(button)navigate(button.dataset.mobileModule);});
    }
    const allowed = modules(), current = activeId();
    document.getElementById('mobileCurrentModule').textContent = allowed.find(row=>row.id===current)?.label || 'TMS';
    const preferred = ['calendario','inventario','importar'].map(id=>allowed.find(row=>row.id===id)).filter(Boolean);
    for (const row of allowed) if (preferred.length<3 && !preferred.some(item=>item.id===row.id)) preferred.push(row);
    document.getElementById('mobileQuickNav').innerHTML = preferred.map(row=>`<button type="button" data-mobile-module="${esc(row.id)}" ${current===row.id?'aria-current="page"':''}>${icon(row.id)}<span>${esc(row.id==='importar'?'Importar':row.label)}</span></button>`).join('')+`<button type="button" data-mobile-module="menu" ${preferred.some(row=>row.id===current)?'':'aria-current="page"'}>${icon('dashboard')}<span>Más</span></button>`;
    const undo = document.getElementById('undoStateBtn');
    if (undo && !undo.querySelector('svg')) { undo.setAttribute('aria-label','Revertir último cambio');undo.innerHTML=icon('undo')+'<span>Revertir último cambio</span>'; }
    const logout = document.getElementById('logoutBtn');
    if (logout && !logout.querySelector('svg')) { logout.setAttribute('aria-label','Salir'); logout.innerHTML=icon('logout')+'<span>Salir</span>'; }
    const modal = document.getElementById('mobileModulesDialog');
    if (modal?.open) renderModules(modal);
  }
  function renderModules(modal) {
    const current=activeId();
    modal.innerHTML=`<header><div><p>Álvarez · TMS</p><h2 id="mobileModulesTitle">Ir a un módulo</h2></div><button type="button" data-mobile-close aria-label="Cerrar módulos">×</button></header><div class="mobile-module-grid">${modules().map(row=>`<button type="button" data-mobile-module="${esc(row.id)}" ${row.id===current?'aria-current="page"':''}>${icon(row.id)}<span>${esc(row.label)}</span>${row.id==='auditoriaJamaica'?'<small>Nueva pestaña</small>':''}</button>`).join('')}</div><footer><span>${esc(document.getElementById('userEmail')?.textContent || '')}</span><span>${esc(document.getElementById('tmsSyncBadge')?.textContent || '')}</span></footer>`;
  }
  function openModules() {
    let modal=document.getElementById('mobileModulesDialog');
    if (!modal) {
      modal=document.createElement('dialog');modal.id='mobileModulesDialog';modal.setAttribute('aria-labelledby','mobileModulesTitle');document.body.appendChild(modal);
      modal.addEventListener('click',event=>{
        const close=event.target.closest('[data-mobile-close]'),button=event.target.closest('[data-mobile-module]');
        if(close)modal.close();else if(button)navigate(button.dataset.mobileModule);
        else if(event.target===modal){const r=modal.getBoundingClientRect();if(event.clientY<r.top||event.clientY>r.bottom||event.clientX<r.left||event.clientX>r.right)modal.close();}
      });
      modal.addEventListener('close',()=>document.getElementById('mobileModulesButton')?.setAttribute('aria-expanded','false'));
    }
    renderModules(modal);if(!modal.open)modal.showModal();document.getElementById('mobileModulesButton').setAttribute('aria-expanded','true');
  }
  function navigate(id) {
    if(id==='menu'){openModules();return;}
    if(!modules().some(row=>row.id===id))return;
    document.getElementById('mobileModulesDialog')?.close();
    if(id==='auditoriaJamaica'){window.open('./auditoria-jamaica.html','_blank','noopener,noreferrer');return;}
    window.cambiarVista(id,document.getElementById('nav-'+id));updateNavigation();
  }
  function updateCalendar() {
    const container=document.getElementById('calContainer');if(!container)return;
    let controls=document.getElementById('mobileCalendarFocus');
    if(!controls){controls=document.createElement('div');controls.id='mobileCalendarFocus';controls.innerHTML='<span>Mostrar</span><button type="button" data-calendar-focus="deliveries">Con entregas</button><button type="button" data-calendar-focus="all">Todos los días</button>';container.before(controls);controls.addEventListener('click',event=>{const button=event.target.closest('[data-calendar-focus]');if(button){onlyDeliveries=button.dataset.calendarFocus==='deliveries';updateCalendar();}});}
    container.classList.toggle('mobile-deliveries-only',onlyDeliveries);
    container.classList.toggle('calendar-no-selection',!!document.getElementById('moveSelectedBtn')?.disabled);
    const queue=document.getElementById('queueSemanaSection');if(queue)queue.classList.toggle('calendar-queue-empty',!queue.querySelector('.queue-card'));
    controls.querySelectorAll('button').forEach(button=>button.setAttribute('aria-pressed',String((button.dataset.calendarFocus==='deliveries')===onlyDeliveries)));
    let total=0;
    container.querySelectorAll('.cal-month-section').forEach(week=>{
      let visible=0;
      week.querySelectorAll('.cal-dia-col').forEach(day=>{const populated=!!day.querySelector('.cal-card,.calendar-note-postit');day.classList.toggle('calendar-empty-day',!populated);if(populated){visible++;total++;}});
      week.classList.toggle('calendar-empty-week',!visible);
    });
    let empty=document.getElementById('mobileCalendarEmpty');
    if(!empty){empty=document.createElement('p');empty.id='mobileCalendarEmpty';empty.textContent='Sin entregas visibles. Selecciona «Todos los días» para ver el calendario completo.';container.after(empty);}
    empty.hidden=total>0||!onlyDeliveries||!container.querySelector('.cal-month-section');
    const toolbar=document.querySelector('#calendario .calendar-toolbar');
    if(toolbar&&!document.getElementById('mobileCalendarTools')){
      const details=document.createElement('details');details.id='mobileCalendarTools';details.innerHTML='<summary>Más opciones</summary><div></div>';toolbar.appendChild(details);
      [...toolbar.children].filter(node=>node.matches('button')&&(/JSON/.test(node.textContent)||node.id==='addTruckBtn')).forEach(node=>details.querySelector('div').appendChild(node));
    }
    const tools=document.getElementById('mobileCalendarTools'), mobile=window.matchMedia('(max-width:900px)').matches;
    if(tools&&calendarToolsMobile!==mobile){tools.open=!mobile;calendarToolsMobile=mobile;}
    container.querySelectorAll('.cal-card').forEach(card=>{
      const actions=card.querySelector('.cal-card-actions');if(!actions)return;
      let menu=card.querySelector('.calendar-card-tools');
      if(!menu){
        menu=document.createElement('details');menu.className='calendar-card-tools';
        const summary=document.createElement('summary');summary.setAttribute('aria-label','Opciones de '+(card.querySelector('.calendar-dispatch-name')?.textContent||'despacho'));summary.title='Mover, completar o añadir nota';summary.innerHTML='<span aria-hidden="true">•••</span>';
        menu.appendChild(summary);actions.before(menu);menu.appendChild(actions);
        menu.addEventListener('click',event=>{event.stopPropagation();if(event.target.closest('button'))menu.open=false;});
        menu.addEventListener('toggle',()=>{if(menu.open&&window.matchMedia('(max-width:900px)').matches)container.querySelectorAll('.calendar-card-tools[open]').forEach(other=>{if(other!==menu)other.open=false;});});
      }
      if(menu.dataset.mobile!==String(mobile)){menu.open=!mobile;menu.dataset.mobile=String(mobile);}
    });
  }
  function labelDialogTables() {
    document.querySelectorAll('.calendar-dispatch-dialog table').forEach(table=>{
      const headers=[...table.querySelectorAll('thead th')].map(th=>th.textContent.trim());
      table.querySelectorAll('tbody tr').forEach(tr=>[...tr.cells].forEach((td,index)=>td.dataset.mobileLabel=headers[index]||''));
    });
  }
  function updateSyncIndicator() { const badge=document.getElementById('tmsSyncBadge');if(badge){const label=badge.textContent;badge.setAttribute('aria-label',label);badge.dataset.mobileState=label==='Sincronizado'?'synced':/conflicto|conexión|Error|sesión/i.test(label)?'warning':'pending';} }
  function refresh() { frame=0;updateSyncIndicator();updateNavigation();updateCalendar();labelDialogTables(); }
  function schedule() { if(!frame)frame=requestAnimationFrame(refresh); }
  function init() {
    updateNavigation();updateCalendar();updateSyncIndicator();
    const select=document.getElementById('mobileModuleSelect');
    if(select)new MutationObserver(schedule).observe(select,{subtree:true,attributes:true,childList:true});
    const container=document.getElementById('calContainer');if(container)new MutationObserver(()=>{updateCalendar();}).observe(container,{childList:true});
    // Add mobile labels to newly opened/refreshed dispatch tables, without touching their data.
    new MutationObserver(records=>{if(records.some(record=>record.target.parentElement?.closest('#tmsSyncBadge')||record.target.id==='tmsSyncBadge'))updateSyncIndicator();if(records.some(record=>[...record.addedNodes].some(node=>node.nodeType===1&&(node.id==='calendarDispatchDialog'||node.closest?.('#calendarDispatchDialog')))))labelDialogTables();}).observe(document.body,{childList:true,subtree:true});
    document.addEventListener('click',event=>{if(!event.target.closest('.calendar-card-tools'))document.querySelectorAll('.calendar-card-tools[data-mobile="true"][open]').forEach(menu=>menu.open=false);});
    window.addEventListener('keydown',event=>{if(event.key==='Escape')document.querySelectorAll('.calendar-card-tools[data-mobile="true"][open]').forEach(menu=>menu.open=false);});
    window.addEventListener('tms-permissions-changed',schedule);window.addEventListener('resize',schedule);
    const oldView=window.cambiarVista;window.cambiarVista=function(...args){const result=oldView.apply(this,args);schedule();return result;};
    const oldLogout=window.cerrarSesion;window.cerrarSesion=function(...args){document.getElementById('mobileModulesDialog')?.close();return oldLogout.apply(this,args);};
  }
  window.addEventListener('DOMContentLoaded',init);
})();

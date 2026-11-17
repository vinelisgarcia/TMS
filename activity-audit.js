(function () {
 'use strict';
 const labels={auditoriaSistema:'Actividad del sistema',importar:'Importar SAP',dashboard:'Dashboard',calendario:'Calendario',comercial:'Comercial',rutas:'Rutas',reportes:'Reportes',importaciones:'Importaciones',materialesTransito:'Materiales en tránsito',prioridades:'Prioridades',incidenciasDespacho:'Incidencias',configuracion:'Configuración',solicitudesAlmacen:'Solicitudes a almacén',almacen:'Almacén',inventario:'Inventario CT / AD',showroom:'Showroom',packingMiami:'Packing list Miami',auditoriaJamaica:'Auditoría Jamaica',eur1:'Verificador EUR.1'};
 const e=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const client=()=>typeof SB!=='undefined'?SB:null;
 const admin=()=>!!(typeof APP!=='undefined'&&APP.currentUserProfile?.activo!==false&&APP.currentUserProfile?.rol==='Admin');
 const moduleId=()=>({configClientes:'configuracion'}[document.querySelector('.view.active')?.id]||document.querySelector('.view.active')?.id||'dashboard');
 const device=()=>innerWidth<600?'mobile':innerWidth<1024?'tablet':'desktop';
 const stamp=value=>value?new Date(value).toLocaleString('es-DO',{timeZone:'America/Santo_Domingo',dateStyle:'short',timeStyle:'short'}):'—';
 let user=null,starting=null,generation=0,channel=null,poll=null,heartbeat=null,cursor=0,polling=false,visit=null,filterVisit=null,older=null;
 const seen=new Set(),toastGroups=new Set();
 let pendingToasts=[],toastTimer;
 function toast(event){
  const key=event.transaction_id+'|'+event.module;if(toastGroups.has(key))return;toastGroups.add(key);if(toastGroups.size>500)toastGroups.delete(toastGroups.values().next().value);
  pendingToasts.push(event);clearTimeout(toastTimer);toastTimer=setTimeout(()=>{const rows=pendingToasts.splice(0);for(const row of rows){
   const box=document.createElement('div');box.className='activity-toast';box.setAttribute('role','status');
   box.innerHTML=`<div><strong>${e(row.actor_name)} · ${e(labels[row.module]||row.module)}</strong><p>${e(row.summary)}</p></div><button aria-label="Cerrar aviso">×</button>`;
   box.querySelector('button').onclick=()=>box.remove();document.getElementById('activityToasts')?.appendChild(box);setTimeout(()=>box.remove(),6500);
   const boxes=document.querySelectorAll('.activity-toast');if(boxes.length>3)boxes[0].remove();
  }},250);
 }
 function receive(row){if(!row||seen.has(String(row.id)))return;seen.add(String(row.id));cursor=Math.max(cursor,Number(row.id));if(seen.size>1000)seen.delete(seen.values().next().value);if(row.is_change)toast(row);if(document.getElementById('auditoriaSistema')?.classList.contains('active')&&admin())scheduleRefresh();}
 const lastAction=new Map();
 async function action(module,kind){if(!user)return null;if(kind.startsWith('exportar_')){const key=module+'|'+kind,now=Date.now();if(now-(lastAction.get(key)||0)<1000)return null;lastAction.set(key,now);}try{const {data,error}=await client().rpc('tms_audit_action',{p_module:module,p_action:kind,p_device:device()});if(error)throw error;visit=data;return data;}catch(error){console.warn('Registro de actividad pendiente:',error.message);return null;}}
 async function refreshChanges(){
  if(!user||polling||document.hidden)return;polling=true;const g=generation;
  try{let more=true;while(more&&g===generation){const {data,error}=await client().from('tms_audit_events').select('*').gt('id',cursor).order('id',{ascending:true}).limit(100);if(error)throw error;if(g!==generation)return;for(const row of data||[])receive(row);more=data?.length===100;}}
  catch(error){console.warn('Avisos de actividad:',error.message);}finally{polling=false;}
 }
 function reset(){generation++;user=null;starting=null;visit=null;cursor=0;seen.clear();toastGroups.clear();pendingToasts=[];clearTimeout(toastTimer);clearInterval(poll);clearInterval(heartbeat);if(channel)client()?.removeChannel(channel);channel=null;document.getElementById('activityToasts')?.replaceChildren();document.getElementById('activityEvents')?.replaceChildren();document.getElementById('activityVisits')?.replaceChildren();}
 async function start(){
  if(user||starting)return starting;const g=generation;
  starting=(async()=>{try{
   const {data,error}=await client().auth.getUser();if(error||!data?.user||g!==generation)return;user=data.user;
   visit=await action('auditoriaSistema','visita');if(!visit){user=null;return;}
   const latest=await client().from('tms_audit_events').select('id').order('id',{ascending:false}).limit(1);if(g!==generation)return;cursor=Number(latest.data?.[0]?.id||0);
   channel=client().channel('activity-'+user.id+'-'+Math.random().toString(36).slice(2)).on('postgres_changes',{event:'INSERT',schema:'public',table:'tms_audit_events'},()=>refreshChanges()).subscribe(()=>refreshChanges());
   poll=setInterval(refreshChanges,8000);heartbeat=setInterval(()=>{if(!document.hidden&&document.hasFocus())action('auditoriaSistema','visita');refreshUI();},60000);
   refreshUI();if(labels[moduleId()])action(moduleId(),'ver_modulo');
  }catch(error){user=null;console.warn('No se inició el registro de actividad:',error.message);}finally{starting=null;}})();return starting;
 }
 function refreshUI(){
  const allowed=admin();const nav=document.getElementById('nav-auditoriaSistema');if(nav)nav.style.display=allowed?'':'none';const option=document.querySelector('#mobileModuleSelect option[value="auditoriaSistema"]');if(option)option.hidden=!allowed;
  if(!allowed&&document.getElementById('auditoriaSistema')?.classList.contains('active')){document.getElementById('activityEvents')?.replaceChildren();document.getElementById('activityVisits')?.replaceChildren();}
 }
 let refreshTimer;function scheduleRefresh(){clearTimeout(refreshTimer);refreshTimer=setTimeout(()=>loadHistory(false),500);}
 function init(){
  const view=document.getElementById('auditoriaSistema');if(!view||view.dataset.ready)return;view.dataset.ready='1';
  view.innerHTML=`<div class="view-header"><div><h1>Actividad del sistema</h1><p>Accesos, visitas y cambios guardados, por usuario.</p></div><button class="btn btn-outline" id="activityRefresh">Actualizar</button></div>
  <div class="card activity-filters"><label>Usuario<select id="activityUser" class="form-control"><option value="">Todos los usuarios</option></select></label><label>Desde<input id="activityFrom" type="date" class="form-control"></label><label>Hasta<input id="activityTo" type="date" class="form-control"></label><label>Mostrar<select id="activityKind" class="form-control"><option value="">Toda la actividad</option><option value="changes">Cambios guardados</option><option value="access">Accesos</option></select></label></div>
  <p id="activityMessage" role="status"></p><div class="activity-layout"><section class="card"><h2>Visitas <span id="activityVisitCount"></span></h2><p class="activity-muted">Una visita nueva empieza tras 30 minutos sin actividad. Al cerrar el navegador se conserva la última actividad observada.</p><button class="btn btn-outline btn-sm" id="activityAllVisits">Ver todas las visitas</button><div id="activityVisits"></div></section><section class="card"><h2 id="activityTimelineTitle">Historial de actividad</h2><div id="activityEvents"></div><button id="activityMore" class="btn btn-outline" hidden>Ver más registros</button></section></div>`;
  for(const id of ['activityUser','activityFrom','activityTo','activityKind'])document.getElementById(id).onchange=()=>{filterVisit=null;loadHistory(false);};
  document.getElementById('activityRefresh').onclick=()=>loadHistory(false);document.getElementById('activityAllVisits').onclick=()=>{filterVisit=null;loadHistory(false);};document.getElementById('activityMore').onclick=()=>loadHistory(true);
 }
 function dateFilter(query,column){const from=document.getElementById('activityFrom').value,to=document.getElementById('activityTo').value;if(from)query=query.gte(column,from+'T00:00:00-04:00');if(to)query=query.lte(column,to+'T23:59:59.999-04:00');return query;}
 let loading=0;
 async function loadHistory(append){
  if(!admin()||!user)return;init();const token=++loading;const msg=document.getElementById('activityMessage');msg.textContent='Cargando actividad…';
  try{
   const selected=document.getElementById('activityUser').value,kind=document.getElementById('activityKind').value;
   let visitsQuery=client().from('tms_audit_visits').select('*',{count:'exact'}).order('started_at',{ascending:false}).limit(100);visitsQuery=dateFilter(visitsQuery,'started_at');if(selected)visitsQuery=visitsQuery.eq('user_id',selected);
   let eventsQuery=client().from('tms_audit_events').select('*').order('id',{ascending:false}).limit(100);eventsQuery=dateFilter(eventsQuery,'created_at');if(selected)eventsQuery=eventsQuery.eq('user_id',selected);if(filterVisit)eventsQuery=eventsQuery.eq('visit_id',filterVisit);if(kind==='changes')eventsQuery=eventsQuery.eq('is_change',true);if(kind==='access')eventsQuery=eventsQuery.eq('action','acceso');if(append&&older)eventsQuery=eventsQuery.lt('id',older);
   const [vr,er]=await Promise.all([visitsQuery,eventsQuery]);if(token!==loading||!admin())return;if(vr.error||er.error)throw vr.error||er.error;
   const profiles=typeof APP!=='undefined'?APP.userProfiles||[]:[];const users=new Map();for(const p of profiles){const id=p.userId||p.auth_user_id;if(id)users.set(id,p.nombre||p.email);}for(const v of vr.data||[])users.set(v.user_id,v.actor_name+' · '+v.actor_email);
   document.getElementById('activityUser').innerHTML='<option value="">Todos los usuarios</option>'+[...users].map(([id,name])=>`<option value="${e(id)}" ${id===selected?'selected':''}>${e(name)}</option>`).join('');
   document.getElementById('activityVisitCount').textContent='('+vr.count+')';document.getElementById('activityVisits').innerHTML=(vr.data||[]).map(v=>`<button class="activity-visit ${filterVisit===v.id?'selected':''}" data-visit="${e(v.id)}"><strong>${e(v.actor_name)}</strong><span>${e(v.actor_email)}</span><span>${e(stamp(v.started_at))} · ${e(v.device)}</span><small>Última actividad: ${e(stamp(v.last_seen_at))}${v.ended_at?' · '+e(v.end_reason):''}</small></button>`).join('')||'<p>Sin visitas en este período.</p>';
   document.querySelectorAll('[data-visit]').forEach(btn=>btn.onclick=()=>{filterVisit=btn.dataset.visit;loadHistory(false);});
   document.getElementById('activityTimelineTitle').textContent=filterVisit?'Actividad de la visita seleccionada':'Historial de actividad';
   const html=(er.data||[]).map(row=>`<article class="activity-event ${row.is_change?'changed':''}"><time>${e(stamp(row.created_at))}</time><strong>${e(row.actor_name)}</strong><p>${e(row.summary)} · ${e(labels[row.module]||row.module)}</p>${Object.keys(row.details||{}).length?`<details><summary>Ver registro</summary><dl>${Object.entries(row.details).map(([k,v])=>`<dt>${e(k.replaceAll('_',' '))}</dt><dd>${e(Array.isArray(v)?v.join(', '):typeof v==='object'?JSON.stringify(v):v)}</dd>`).join('')}</dl></details>`:''}</article>`).join('');
   if(append)document.getElementById('activityEvents').insertAdjacentHTML('beforeend',html);else document.getElementById('activityEvents').innerHTML=html||'<p>Sin actividad para estos filtros.</p>';
   older=er.data?.at(-1)?.id||null;document.getElementById('activityMore').hidden=er.data?.length!==100;msg.textContent=vr.count>100?'Se muestran las 100 visitas más recientes. Filtra por usuario o fecha para consultar otras.':'';
  }catch(error){msg.textContent='No se pudo consultar el historial: '+error.message;}
 }
 const originalJamaica=window.abrirAuditoriaJamaica;if(originalJamaica)window.abrirAuditoriaJamaica=function(...args){const result=originalJamaica.apply(this,args);if(window.TMSAuditCore?.hasPermission('auditoriaJamaica','ver'))action('auditoriaJamaica','ver_modulo');return result;};
 const originalMostrar=window.mostrarApp;window.mostrarApp=function(...args){const result=originalMostrar?.apply(this,args);start();return result;};
 const originalView=window.cambiarVista;window.cambiarVista=function(id,...args){if(id==='auditoriaSistema'&&!admin())return;const result=originalView?.call(this,id,...args);refreshUI();if(id==='auditoriaSistema'){init();loadHistory(false);}if(user&&document.querySelector('.view.active')?.id===id&&labels[moduleId()])action(moduleId(),'ver_modulo');return result;};
 const originalLogout=window.cerrarSesion;window.cerrarSesion=async function(...args){await Promise.race([action('auditoriaSistema','cerrar_sesion'),new Promise(resolve=>setTimeout(resolve,1500))]);reset();return originalLogout?.apply(this,args);};
 // Record completed download invocations, never filenames/URLs or file content.
 if(typeof XLSX!=='undefined'){for(const name of ['writeFile','writeFileXLSX']){const original=XLSX[name];if(original)XLSX[name]=function(...args){const result=original.apply(this,args);if(user)action(moduleId(),'exportar_excel');return result;};}}
 const anchorClick=HTMLAnchorElement.prototype.click;HTMLAnchorElement.prototype.click=function(...args){const result=anchorClick.apply(this,args);if(this.download&&/\.(xlsx?|csv|pdf)$/i.test(this.download)&&user)action(moduleId(),/\.pdf$/i.test(this.download)?'exportar_pdf':'exportar_excel');return result;};
 document.addEventListener('visibilitychange',()=>{if(!document.hidden){if(user){action('auditoriaSistema','visita');refreshChanges();}else start();refreshUI();}});
 window.addEventListener('online',()=>{if(user)refreshChanges();else start();});
 new MutationObserver(records=>{for(const record of records)if(record.target.open&&user&&record.target.id!=='mobileModuleDialog')action(moduleId(),'consultar_detalle');}).observe(document.body,{subtree:true,attributes:true,attributeFilter:['open']});
 client()?.auth.onAuthStateChange?.((event)=>{if(event==='SIGNED_OUT')reset();});
 const toasts=document.createElement('aside');toasts.id='activityToasts';toasts.setAttribute('aria-label','Cambios guardados');document.body.appendChild(toasts);
 window.TMSActivity={action,refreshUI,refreshChanges,start};refreshUI();
})();

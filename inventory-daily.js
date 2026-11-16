(function(){
 'use strict';
 const D=window.InventoryDailyDomain,C=window.InventoryCountDomain,X=window.InventoryDailyExcel,pending=new Map(),notices=new Map();let working=false;
 const api=kind=>kind==='showroom'?window.TMSShowroom:window.TMSInventory;
 const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
 const when=value=>{const date=new Date(value);return Number.isFinite(date.getTime())?date.toLocaleString('es-DO',{timeZone:'America/Santo_Domingo',dateStyle:'short',timeStyle:'short'}):value;};
 const modalId=kind=>'dailyCountDialog-'+kind;
 function notify(kind,text){notices.set(kind,text);const node=document.getElementById('dailyCountNotice-'+kind);if(node)node.textContent=text;}
 function close(kind){const hadDraft=pending.delete(kind);document.getElementById(modalId(kind))?.remove();if(hadDraft)api(kind).markDraft();}
 function dialog(kind,html){
  document.getElementById(modalId(kind))?.remove();const modal=document.createElement('dialog');modal.id=modalId(kind);modal.className='calendar-dispatch-dialog daily-count-dialog';modal.dataset.kind=kind;
  modal.setAttribute('aria-labelledby','dailyCountTitle-'+kind);modal.innerHTML=html;document.body.appendChild(modal);
  modal.addEventListener('cancel',event=>{event.preventDefault();close(kind);});modal.showModal();return modal;
 }
 function mount(kind,nav){
  const app=api(kind);if(!app?.can('ver'))return;
  nav.insertAdjacentHTML('beforeend',`<div class="inv-actions daily-count-toolbar">${app.can('editar')?`<button class="btn btn-outline" data-daily="template" data-kind="${kind}">Plantilla de conteo ciego</button><button class="btn btn-outline" data-daily="pick-file" data-kind="${kind}">Subir conteo diario</button><input id="dailyCountFile-${kind}" type="file" accept=".xlsx" data-daily-file="${kind}" hidden>`:''}<button class="btn btn-outline" data-daily="history" data-kind="${kind}">Conteos por día</button><span id="dailyCountNotice-${kind}" class="inv-muted" role="status">${esc(notices.get(kind)||'')}</span></div>`);
  const viewer=document.getElementById(modalId(kind));if(viewer?.dataset.historyDay)history(kind,viewer.dataset.historyDay);
 }
 function openTemplate(kind,restored){
  if(!api(kind).can('editar'))throw Error('Tu rol no permite generar planes de conteo.');
  const value={stage:'template',day:C.countDay(),scope:'all',...restored};pending.set(kind,value);
  const state=api(kind).getState(),cycle=C.active(state),fresh=!cycle||cycle.status!=='open'||cycle.sourcePair!==C.pair(state)||cycle.differenceConvention!==(kind==='showroom'?'physical_virtual':'ad_ct');const cycleNote=fresh?'Se abrirá un ciclo con el inventario actual. Los ciclos anteriores y sus conteos se conservarán.':'Se usará el ciclo de conteo abierto.';
  const modal=dialog(kind,`<h2 id="dailyCountTitle-${kind}">Plantilla de conteo ciego</h2><p>Comparte el Excel con el equipo. La plantilla no incluye existencias ni valores económicos.</p><div class="inv-form"><label>Fecha de conteo<input id="dailyPlanDay-${kind}" type="date" value="${esc(value.day)}" required></label><label>Artículos<select id="dailyPlanScope-${kind}"><option value="all" ${value.scope==='all'?'selected':''}>Todas las referencias</option><option value="guided" ${value.scope==='guided'?'selected':''}>Obligatorios del conteo guiado</option></select></label></div><p>${esc(cycleNote)} Sube las plantillas completadas antes de cerrar o archivar ese ciclo.</p><div class="inv-actions"><button class="btn btn-primary" data-daily="download" data-kind="${kind}">Descargar Excel</button><button class="btn btn-outline" data-daily="cancel" data-kind="${kind}">Cancelar</button></div><p id="dailyCountError-${kind}" role="alert" hidden></p>`);
  modal.querySelector('input').focus();api(kind).markDraft();
 }
 function readTemplate(kind){const value=pending.get(kind);if(value?.stage==='template'){const day=document.getElementById('dailyPlanDay-'+kind);if(day)value.day=day.value;value.scope=document.getElementById('dailyPlanScope-'+kind)?.value||value.scope;}return value;}
 async function download(kind){
  if(working)return;working=true;
  try{
   const app=api(kind);if(!app.can('editar'))throw Error('Tu rol no permite generar planes de conteo.');
   const value=readTemplate(kind),prepared=D.plan(app.getState(),app.getReport(),{kind,day:value.day,scope:value.scope,user:app.currentUser()?.email||''});
   const library=await X.load(),book=X.createWorkbook(library,prepared.template),bytes=await book.xlsx.writeBuffer();
   close(kind);await app.edit(state=>Object.assign(state,prepared.data));
   const url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})),link=document.createElement('a');link.href=url;link.download='TMS_Conteo_Ciego_'+(kind==='showroom'?'Showroom':'CT_AD')+'_'+prepared.template.day+'.xlsx';link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
   notify(kind,'Plantilla descargada. Completa la fecha, cantidad física y responsable; luego usa Subir conteo diario.');
  }finally{working=false;}
 }
 async function prepare(kind,file){
  if(!api(kind).can('editar'))throw Error('Tu rol no permite cargar conteos.');if(!file?.size)throw Error('El archivo está vacío.');
  const book=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false,cellNF:true,raw:true}),decoded=X.decode(book,XLSX);
  if(decoded.template.kind!==kind)throw Error('La plantilla corresponde a otro inventario. Súbela en su módulo.');
  const parsed=D.parse(decoded.template,decoded.matrix,decoded.options),value={stage:'import',fileName:file.name,template:decoded.template,rows:parsed.rows,blank:parsed.blank,days:parsed.days};
  pending.set(kind,value);renderPreview(kind);api(kind).markDraft();
 }
 function renderPreview(kind){
  const value=pending.get(kind);if(value?.stage!=='import')return;
  let result,error;try{result=D.apply(api(kind).getState(),value.template,value.rows,{kind,fileName:value.fileName,user:api(kind).currentUser()?.email||''});}catch(problem){error=problem.message;}
  dialog(kind,`<h2 id="dailyCountTitle-${kind}">Revisar conteo diario</h2><p>${esc(value.fileName)}</p><div class="inv-notice">${value.rows.length} registros con cantidad. ${value.blank} celdas vacías seguirán sin contar.<br>Fechas de conteo: ${esc(value.days.join(', '))}.${result?`<br>${result.inserted} nuevos. ${result.skipped} ya cargados, sin duplicarlos.`:''}</div>${error?`<p class="inv-error" role="alert">${esc(error)}</p>`:''}<div class="table-wrap"><table class="inv-table"><thead><tr><th>Fecha de conteo</th><th>Referencia</th><th>Ubicación contada</th><th>Cantidad física</th><th>Contado por</th></tr></thead><tbody>${value.rows.slice(0,100).map(row=>`<tr><td>${esc(row.countedOn)}</td><td>${esc(row.code)}</td><td>${esc(row.actualLocation||'Sin confirmar')}</td><td>${esc(row.quantity)}</td><td>${esc(row.countedBy)}</td></tr>`).join('')}</tbody></table></div>${value.rows.length>100?'<p>Vista previa de los primeros 100 registros. Se validan y cargan todos.</p>':''}<label class="daily-count-confirm"><input id="dailyCountConfirmed-${kind}" type="checkbox">Confirmo que son conteos físicos realizados en las fechas indicadas.</label><div class="inv-actions"><button class="btn btn-primary" data-daily="commit" data-kind="${kind}" ${error?'disabled':''}>Confirmar carga</button><button class="btn btn-outline" data-daily="cancel" data-kind="${kind}">Cancelar</button></div><p id="dailyCountError-${kind}" role="alert" hidden></p>`);
 }
 async function commit(kind){
  if(working)return;if(!document.getElementById('dailyCountConfirmed-'+kind)?.checked)throw Error('Confirma los conteos físicos y sus fechas antes de cargar.');
  const value=pending.get(kind);if(value?.stage!=='import')return;const app=api(kind);if(!app.can('editar'))throw Error('Tu rol no permite cargar conteos.');
  D.apply(app.getState(),value.template,value.rows,{kind,fileName:value.fileName,user:app.currentUser()?.email||''});
  working=true;try{close(kind);let result;await app.edit(state=>{result=D.apply(state,value.template,value.rows,{kind,fileName:value.fileName,user:app.currentUser()?.email||''});});notify(kind,`${result.inserted} conteos guardados para ${result.days.join(', ')}. ${result.skipped} ya estaban cargados.`);}finally{working=false;}
 }
 function history(kind,selectedDay){
  if(!api(kind).can('ver'))throw Error('Tu rol no permite consultar los conteos.');
  const rows=D.history(api(kind).getState()),days=[...new Set(rows.map(row=>row.countedOn))].filter(Boolean).sort().reverse(),day=selectedDay||days[0]||C.countDay(),shown=rows.filter(row=>row.countedOn===day);
  const viewer=dialog(kind,`<h2 id="dailyCountTitle-${kind}">Conteos por día</h2><label>Fecha de conteo<select data-daily-day="${kind}">${[...new Set([day,...days])].sort().reverse().map(value=>`<option value="${esc(value)}" ${value===day?'selected':''}>${esc(value)}</option>`).join('')}</select></label><p>${new Set(shown.filter(row=>!row.voided).map(row=>row.code)).size} referencias contadas. ${shown.filter(row=>!row.voided).length} registros sin anular en la bitácora de este día.</p><div class="table-wrap"><table class="inv-table"><thead><tr><th>Referencia</th><th>Ubicación</th><th>Cantidad</th><th>Contado por</th><th>Registrado en TMS</th><th>Estado</th></tr></thead><tbody>${shown.slice(0,500).map(row=>`<tr><td>${esc(row.code)}<small>${esc(row.description)}</small></td><td>${esc(row.actualLocation||'Sin confirmar')}</td><td>${esc(row.quantity)} ${esc(row.unit)}</td><td>${esc(row.countedBy||row.by)}</td><td>${esc(when(row.at))}<small>${esc(row.by)}</small></td><td>${row.voided?'Anulado':row.cycleStatus==='archived'?'Ciclo archivado':row.cycleStatus==='closed'?'Ciclo cerrado':row.current?'Vigente':'Reconteo posterior'}</td></tr>`).join('')||'<tr><td colspan="6">No hay conteos para esta fecha.</td></tr>'}</tbody></table></div>${shown.length>500?'<p>Se muestran los primeros 500 registros. El Excel del ciclo conserva el detalle completo.</p>':''}<div class="inv-actions"><button class="btn btn-outline" data-daily="cancel" data-kind="${kind}">Cerrar</button></div>`);viewer.dataset.historyDay=day;
 }
 function error(kind,problem){const message=(problem.invalidRows||[]).map(row=>'Fila '+row.row+': '+row.message).slice(0,20).join(' '),text=message||problem.message,node=document.getElementById('dailyCountError-'+kind);if(node){node.hidden=false;node.textContent=text;}else notify(kind,text);}
 document.addEventListener('click',event=>{const button=event.target.closest('[data-daily]');if(!button||button.disabled)return;const kind=button.dataset.kind;try{Promise.resolve(({'pick-file':()=>document.getElementById('dailyCountFile-'+kind)?.click(),template:()=>openTemplate(kind),download:()=>download(kind),commit:()=>commit(kind),history:()=>history(kind),cancel:()=>close(kind)})[button.dataset.daily]?.()).catch(problem=>error(kind,problem));}catch(problem){error(kind,problem);}});
 document.addEventListener('change',event=>{const input=event.target;if(input.dataset.dailyFile){const kind=input.dataset.dailyFile;prepare(kind,input.files[0]).catch(problem=>error(kind,problem));input.value='';}if(input.dataset.dailyDay)history(input.dataset.dailyDay,input.value);if(input.closest('.daily-count-dialog')){const kind=input.closest('.daily-count-dialog').dataset.kind;if(pending.has(kind)){readTemplate(kind);api(kind).markDraft();}}});
 window.InventoryDailyCounts={mount,reset:kind=>{pending.delete(kind);document.getElementById(modalId(kind))?.remove();},hasDraft:kind=>pending.has(kind),captureDrafts:kind=>pending.has(kind)?[{type:'daily-count',kind,value:JSON.parse(JSON.stringify(readTemplate(kind)))}]:[],restoreDrafts:(kind,rows)=>{const draft=rows?.find(row=>row.type==='daily-count'&&row.kind===kind);if(!draft)return;pending.set(kind,draft.value);if(draft.value.stage==='template')openTemplate(kind,draft.value);else renderPreview(kind);}};
})();

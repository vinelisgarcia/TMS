(function () {
  'use strict';
  const D=window.InventoryDomain, copy=value=>JSON.parse(JSON.stringify(value));
  const empty=()=>({version:1,sources:{ct:null,ad:null},cases:{},settings:{sameScope:true}});
  const escape=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const number=value=>value===null||value===undefined?'—':new Intl.NumberFormat('es-DO',{maximumFractionDigits:3}).format(value);
  const money=value=>value===null||value===undefined?'Sin valoración':new Intl.NumberFormat('es-DO',{style:'currency',currency:'DOP',maximumFractionDigits:2}).format(value);
  const unitMoney=value=>value===null||value===undefined?'Sin valoración':new Intl.NumberFormat('es-DO',{style:'currency',currency:'DOP',minimumFractionDigits:2,maximumFractionDigits:6}).format(value);
  const can=action=>!!(window.TMSAuditCore&&window.TMSAuditCore.hasPermission('inventario',action));
  const client=()=>typeof SB!=='undefined'?SB:window.SB;
  let data=empty(),sync=null,user=null,starting=null,pending=null,editor=null,message='',syncState={phase:'disconnected'};
  const filters={search:'',status:'differences',followup:'',page:0};
  const statuses={equal:'Coincide',difference:'Diferencia',only_ct:'Solo CT',only_ad:'Solo AD',unit_conflict:'Unidades incompatibles'};
  const workflow={pendiente:'Pendiente',en_revision:'En revisión',rectificado:'Rectificación registrada',justificado:'Diferencia justificada'};
  const pairKey=()=>[data.sources.ct?.semanticHash||'ct-pendiente',data.sources.ad?.semanticHash||'ad-pendiente'].join(':');
  const caseKey=code=>pairKey()+':'+code;
  const getCase=code=>data.cases[caseKey(code)]||{status:'pendiente',note:'',owner:'',document:''};
  const report=()=>D.compare(data.sources.ct,data.sources.ad);
  const hasDraft=()=>!!(pending||editor);
  function notify(value){message=value;const node=document.getElementById('inventoryMessage');if(node){node.textContent=message;node.hidden=!message;}}
  function syncStatus(state){
    syncState=state;
    const badge=document.getElementById('inventorySyncStatus');
    if(badge)badge.textContent=({loading:'Cargando…',synced:'Inventario sincronizado',pending:'Cambios pendientes',saving:'Guardando inventario…',offline:'Sin conexión',conflict:'Actualización pendiente'})[state.phase]||'Sincronización pendiente';
    const panel=document.getElementById('inventorySyncPanel');
    if(panel){panel.hidden=!state.conflict&&state.phase!=='offline'&&!state.error;panel.innerHTML=`<span>${escape(state.conflict?'Otra sesión actualizó el inventario. Tus cambios se conservan. Descarga tu copia antes de cargar la versión compartida.':state.error||'Se reintentará al volver la conexión.')}</span><div class="inv-actions"><button class="btn btn-outline btn-sm" data-inv="backup">Descargar mi copia</button><button class="btn btn-outline btn-sm" data-inv="reload">${state.conflict?'Cargar versión compartida':'Reintentar'}</button></div>`;}
  }
  function getSync(){
    if(!sync)sync=window.TMSSessionSync.create({
      namespace:'tms_inventory',getSnapshotRpc:'tms_inventory_get_snapshot',saveSnapshotRpc:'tms_inventory_save_snapshot',realtimeTable:'tms_inventory_sync_state',getClient:client,
      capture:()=>({tms_inventory_state:[{id:true,payload:copy(data)}]}),
      applySnapshot:snapshot=>{const row=snapshot.tables?.tms_inventory_state?.[0];if(!row?.payload)throw new Error('No se recibió el inventario compartido.');data=copy(row.payload);render();},
      onStatus:syncStatus,hasDraft,
      captureDrafts:()=>editor?[{type:'case',code:editor,values:readCaseForm()}]:[],
      restoreDrafts:drafts=>{const draft=drafts?.find(row=>row.type==='case');if(draft){openCase(draft.code);Object.entries(draft.values).forEach(([key,value])=>{const el=document.getElementById('inventoryCase_'+key);if(el)el.value=value;});}}
    });
    return sync;
  }
  async function start(){
    if(!can('ver'))throw new Error('Tu rol no tiene acceso al inventario.');
    if(starting)return starting;
    starting=(async()=>{
      const result=await client().auth.getUser();if(result.error)throw result.error;
      const next=result.data?.user;if(!next)throw new Error('Inicia sesión para usar el inventario.');
      if(user&&user.id!==next.id){sync?.stop();data=empty();pending=null;editor=null;}
      user=next;await getSync().start(user);
    })();
    try{await starting;}catch(error){notify(error.message);throw error;}finally{starting=null;}
  }
  async function open(){render();try{await start();}catch(_){} }
  async function save(){
    if(!can('editar')&&!can('importar'))throw new Error('Tu rol solo puede consultar el inventario.');
    if(!sync||sync.getState().revision===null)await start();
    getSync().markDirty();
    try{return await sync.save();}catch(error){notify(error.message);throw error;}
  }
  function sourceCard(source,label){
    const snapshot=data.sources[source];
    return `<section class="card inv-source"><div class="inv-actions" style="justify-content:space-between"><h3>${label}</h3><span class="inv-pill">${snapshot?'Cargado':'Pendiente'}</span></div><p class="inv-muted">${source==='ct'?'Stock físico de Caribetrans. Se suman las filas y los lotes de cada material.':'Stock de AD depositado en CT. Usa cantidad y valor total del stock en DOP.'}</p>${snapshot?`<div class="inv-source-meta"><strong>${escape(snapshot.fileName)}</strong><br>Corte: ${escape(snapshot.asOf)} · ${number(snapshot.acceptedRowCount)} filas · ${number(snapshot.materials.length)} materiales<br>${number(snapshot.repeatedMaterials)} materiales repetidos consolidados · ${number(snapshot.totals.quantity)} unidades${source==='ad'?`<br>Valor AD: ${money(snapshot.totals.totalValueDop)}`:''}<br>Almacenes: ${escape([...new Set(snapshot.materials.flatMap(row=>row.warehouses))].join(', ')||'No indicado')}</div>`:'<div class="inv-source-meta">Todavía no se ha cargado este inventario.</div>'}${can('importar')||can('editar')?`<label class="btn btn-outline" for="inventory${source.toUpperCase()}File">${snapshot?'Reemplazar reporte':'Seleccionar reporte'}</label><input id="inventory${source.toUpperCase()}File" type="file" accept=".xlsx,.xls,.csv,.txt" data-inv-file="${source}" style="display:none">`:''}</section>`;
  }
  function render(){
    const mount=document.getElementById('inventoryMount');if(!mount)return;
    if(!can('ver')){mount.innerHTML='<div class="inv-empty"><h3>Inventario CT / AD</h3><p>Solicita al administrador acceso a este módulo.</p></div>';return;}
    const result=report(),s=result.summary;
    mount.innerHTML=`<header class="inv-head"><div><div class="inv-eyebrow">Conciliación de almacén</div><h2>Inventario CT / AD</h2><p>Compara mercancía depositada en Caribetrans y prioriza diferencias por su impacto en DOP.</p></div><div class="inv-actions"><span id="inventorySyncStatus" class="inv-pill"></span><button class="btn btn-outline" data-inv="export" ${!result.ready?'disabled':''}>Exportar Excel</button></div></header><div id="inventoryMessage" class="inv-notice" role="status" ${!message?'hidden':''}>${escape(message)}</div><div id="inventorySyncPanel" class="inv-notice" hidden></div><div class="inv-sources">${sourceCard('ct','Caribetrans · CT')}${sourceCard('ad','Álvarez Dominicana · AD')}</div>${result.ready?`<div class="inv-grid"><div class="inv-kpi"><span>Diferencias con valoración</span><strong>${money(s.grossImpactDop)}</strong><small>Impacto absoluto · no compensa sobrantes con faltantes</small></div><div class="inv-kpi"><span>Materiales con diferencias</span><strong>${number(s.differences)}</strong><small>${number(s.equal)} coinciden · ${number(s.materials)} en el cruce</small></div><div class="inv-kpi"><span>Pendientes de valoración</span><strong>${number(s.unvalued)}</strong><small>No se consideran impacto cero</small></div><div class="inv-kpi"><span>Diferencia neta valorada</span><strong>${money(s.netImpactDop)}</strong><small>Signo de CT menos AD</small></div></div>${result.issues.map(issue=>`<div class="inv-notice">${escape(issue)}</div>`).join('')}<section class="card"><div class="card-title">Mayor impacto económico</div>${result.rows.filter(row=>row.impactDop>0).slice(0,5).map(row=>`<div class="inv-priority"><span>${escape(row.code)}</span><div class="inv-priority-bar"><i style="width:${Math.max(1,row.impactDop/(result.rows[0].impactDop||1)*100)}%"></i></div><strong style="text-align:right">${money(row.impactDop)}</strong></div>`).join('')||'<p class="inv-muted">No hay diferencias con valoración disponible.</p>'}</section><section class="card"><div class="card-title">Conciliación y rectificación</div><div class="inv-toolbar"><input id="inventorySearch" type="search" placeholder="Buscar material o descripción" value="${escape(filters.search)}"><select id="inventoryStatus"><option value="differences">Con diferencias</option><option value="all">Todos</option><option value="equal">Coinciden</option><option value="only_ct">Solo CT</option><option value="only_ad">Solo AD</option><option value="unvalued">Sin valoración</option><option value="unit_conflict">Unidades incompatibles</option></select><select id="inventoryFollowup"><option value="">Todo el seguimiento</option>${Object.entries(workflow).map(([key,label])=>`<option value="${key}">${label}</option>`).join('')}</select></div><div id="inventoryResults"></div></section>`:`<div class="inv-empty"><h3>${data.sources.ct?'Falta cargar el reporte AD':data.sources.ad?'Falta cargar el reporte CT':'Carga los dos reportes para comparar'}</h3><p class="inv-muted">El módulo no inventa un inventario ausente ni lo trata como stock cero.</p></div>`}<details class="card inv-method"><summary>Método de valoración y opciones de carga</summary><p><strong>Consolidación:</strong> cada material suma todas sus cantidades y todos sus valores AD. El costo ponderado se calcula después de consolidar: valor total AD en DOP ÷ cantidad AD. La diferencia es CT − AD; el impacto para priorizar es su valor absoluto × costo ponderado. El costo no se redondea antes del cálculo; el impacto de cada material se redondea a centavos. No se utiliza FOB de CT ni una tasa de cambio supuesta.</p><p><strong>Carga directa disponible:</strong> Excel XLSX/XLS o CSV. Revisa hoja, columnas, unidad y fecha antes de confirmar. Cada reporte sustituye la fuente seleccionada; no se acumulan las dos versiones del mismo inventario.</p><p><strong>Carpeta en la nube:</strong> puede alimentar el mismo proceso cuando se defina Drive, OneDrive u otro proveedor, la carpeta y sus permisos. Requiere una conexión y un sincronizador; esta pantalla no simula esa conexión. Por ahora, los archivos se procesan al cargarlos y el inventario consolidado se guarda de forma privada en el TMS.</p><p>El seguimiento documenta la revisión; no modifica SAP ni el sistema de CT. Un nuevo par de reportes inicia su propio seguimiento.</p></details>`;
    const status=document.getElementById('inventoryStatus');if(status)status.value=filters.status;
    const followup=document.getElementById('inventoryFollowup');if(followup)followup.value=filters.followup;
    renderTable();syncStatus(sync?sync.getState():syncState);
  }
  function filteredRows(){
    const query=filters.search.trim().toLowerCase();
    return report().rows.filter(row=>(!query||[row.code,row.description].some(value=>String(value).toLowerCase().includes(query)))&&(filters.status==='all'||filters.status==='differences'&&row.status!=='equal'||filters.status==='unvalued'&&row.difference!==0&&row.impactDop===null||row.status===filters.status)&&(!filters.followup||getCase(row.code).status===filters.followup));
  }
  function renderTable(){
    const mount=document.getElementById('inventoryResults');if(!mount)return;
    const rows=filteredRows(),pages=Math.max(1,Math.ceil(rows.length/100));filters.page=Math.min(filters.page,pages-1);
    const visible=rows.slice(filters.page*100,(filters.page+1)*100);
    mount.innerHTML=`<div class="table-wrap"><table class="inv-table"><thead><tr><th>Material / descripción</th><th class="inv-number">CT</th><th class="inv-number">AD</th><th class="inv-number">CT − AD</th><th class="inv-number">Costo AD / unidad</th><th class="inv-number">Impacto absoluto DOP</th><th>Seguimiento</th></tr></thead><tbody>${visible.map(row=>{const follow=getCase(row.code);return `<tr><td><strong>${escape(row.code)}</strong><small>${escape(row.description)}</small><span class="inv-unit">${escape(row.unit)} · ${statuses[row.status]}</span>${row.issues.length?`<small>${escape(row.issues.join(' '))}</small>`:''}</td><td class="inv-number">${number(row.ctQty)}<small>${row.ct?row.ct.rowCount+' filas':'No registrado'}</small></td><td class="inv-number">${number(row.adQty)}<small>${row.ad?row.ad.rowCount+' filas':'No registrado'}</small></td><td class="inv-number ${row.difference?'inv-delta':''}">${row.difference>0?'+':''}${number(row.difference)}</td><td class="inv-number">${unitMoney(row.unitCostDop)}</td><td class="inv-number"><strong>${money(row.impactDop)}</strong></td><td><span class="inv-pill">${workflow[follow.status]||workflow.pendiente}</span>${follow.owner?`<small>${escape(follow.owner)}</small>`:''}<button class="btn btn-outline btn-sm" data-inv="case" data-code="${escape(row.code)}">${can('editar')?'Revisar':'Ver detalle'}</button></td></tr>`;}).join('')||'<tr><td colspan="7">No hay materiales con estos filtros.</td></tr>'}</tbody></table></div><div class="inv-pagination"><span>${number(rows.length)} materiales · Página ${filters.page+1} de ${pages}</span><div class="inv-actions"><button class="btn btn-outline btn-sm" data-inv="prev" ${!filters.page?'disabled':''}>Anterior</button><button class="btn btn-outline btn-sm" data-inv="next" ${filters.page>=pages-1?'disabled':''}>Siguiente</button></div></div>`;
  }
  async function sha(value){const bytes=typeof value==='string'?new TextEncoder().encode(value):value;const hash=await crypto.subtle.digest('SHA-256',bytes);return [...new Uint8Array(hash)].map(value=>value.toString(16).padStart(2,'0')).join('');}
  async function prepareImport(source,file){
    if(!can('importar')&&!can('editar'))throw new Error('Tu rol no puede cargar inventarios.');
    if(!file||file.size===0)throw new Error('El archivo está vacío (0 bytes). Selecciona una copia válida.');
    if(file.size>30*1024*1024)throw new Error('El archivo supera 30 MB. Divide el reporte por almacén antes de cargarlo.');
    if(!sync||sync.getState().revision===null)await start();
    const buffer=await file.arrayBuffer(),fileHash=await sha(buffer);let sheets;
    if(/\.(csv|txt)$/i.test(file.name)){
      let decoded;try{decoded=new TextDecoder('utf-8',{fatal:true}).decode(buffer);}catch(_){decoded=new TextDecoder('windows-1252').decode(buffer);}
      sheets=[{name:'Reporte',matrix:D.parseCsv(decoded)}];
    }else if(/\.(xlsx|xls)$/i.test(file.name)){
      const book=XLSX.read(buffer,{type:'array',cellDates:false});
      sheets=book.SheetNames.map(name=>({name,matrix:XLSX.utils.sheet_to_json(book.Sheets[name],{header:1,defval:null,raw:true})}));
    }else throw new Error('Formato no compatible. Usa XLSX, XLS o CSV.');
    if(!sheets.length||sheets.every(sheet=>sheet.matrix.length===0))throw new Error('El archivo no contiene hojas con datos.');
    pending={source,fileName:file.name,fileHash,sheets,sheetIndex:0,preview:null};
    renderImportDialog();
  }
  function modal(id,html){document.getElementById(id)?.remove();const node=document.createElement('div');node.id=id;node.className='inv-modal';node.innerHTML=`<section class="inv-dialog" role="dialog" aria-modal="true">${html}</section>`;document.body.appendChild(node);return node;}
  function renderImportDialog(){
    if(!pending)return;const sheet=pending.sheets[pending.sheetIndex],layout=D.detectLayout(sheet.matrix,pending.source),headers=sheet.matrix[layout.headerRow]||[];
    const labels={code:'Material / código *',description:'Descripción',quantity:'Cantidad física / stock *',totalValue:'Valor TOTAL AD en DOP *',unit:'Unidad',warehouse:'Almacén',batch:'Lote',availability:'Disponibilidad (opcional)'};
    modal('inventoryImportDialog',`<h3>Revisar carga ${pending.source.toUpperCase()}</h3><p class="inv-muted">${escape(pending.fileName)} · Las referencias repetidas se suman. El reporte reemplazará solo la fuente ${pending.source.toUpperCase()}.</p><div class="inv-form"><label>Hoja<select id="inventorySheet">${pending.sheets.map((entry,index)=>`<option value="${index}" ${index===pending.sheetIndex?'selected':''}>${escape(entry.name)}</option>`).join('')}</select></label><label>Fila de encabezados<input id="inventoryHeaderRow" type="number" min="1" value="${layout.headerRow+1}"></label><label>Fecha de corte *<input id="inventoryAsOf" type="date" value="${escape(layout.asOf||'')}"></label>${Object.entries(labels).filter(([key])=>pending.source==='ad'||key!=='totalValue').map(([key,label])=>`<label>${label}<select id="inventoryMap_${key}"><option value="">No indicada</option>${headers.map((header,index)=>`<option value="${index}" ${layout.columns[key]===index?'selected':''}>${index+1}. ${escape(header||'Columna sin título')}</option>`).join('')}</select></label>`).join('')}<label>Separador decimal<select id="inventoryDecimal"><option value="comma">Coma decimal: 1.234,56</option><option value="dot">Punto decimal: 1,234.56</option><option value="auto">Detectar; rechazar cifras ambiguas</option></select></label><label>Unidad si no hay columna<input id="inventoryUnit" value="UND"></label><label>Almacén, opcional<input id="inventoryWarehouse" placeholder="Todos; separar códigos con coma"></label><label class="inv-wide"><span><input id="inventoryScope" type="checkbox" checked style="width:auto"> Este reporte corresponde a mercancía AD depositada en CT y usa unidades comparables.</span></label></div><div class="inv-actions"><button class="btn btn-outline" data-inv="preview">Calcular vista previa</button><button class="btn btn-primary" id="inventoryConfirmImport" data-inv="confirm" disabled>Confirmar carga</button><button class="btn btn-outline" data-inv="cancel-import">Cancelar</button></div><div id="inventoryImportPreview" class="inv-preview">Revisa las columnas y la fecha de corte antes de confirmar.${pending.source==='ad'?'<br>El campo de valor es el total en DOP de cada fila, nunca un precio unitario.':''}</div>`);
  }
  function importOptions(){
    const mapping={};['code','description','quantity','totalValue','unit','warehouse','batch','availability'].forEach(key=>{const value=document.getElementById('inventoryMap_'+key)?.value;if(value!==undefined&&value!=='')mapping[key]=Number(value);});
    const asOf=document.getElementById('inventoryAsOf').value;if(!asOf)throw new Error('Indica la fecha de corte del reporte.');
    if(!document.getElementById('inventoryScope').checked)throw new Error('Confirma el alcance de los reportes antes de comparar.');
    return {source:pending.source,fileName:pending.fileName,fileHash:pending.fileHash,asOf,headerRow:Number(document.getElementById('inventoryHeaderRow').value)-1,mapping,decimal:document.getElementById('inventoryDecimal').value,unitDefault:document.getElementById('inventoryUnit').value||'UND',warehouses:document.getElementById('inventoryWarehouse').value.split(',').map(value=>value.trim()).filter(Boolean)};
  }
  function previewImport(){
    if(!pending)return;
    try{
      const source=D.consolidate(pending.sheets[pending.sheetIndex].matrix,importOptions());pending.preview=source;
      document.getElementById('inventoryImportPreview').innerHTML=`<strong>${number(source.acceptedRowCount)} filas → ${number(source.materials.length)} materiales consolidados</strong><br>Cantidad total: ${number(source.totals.quantity)} · ${escape(Object.keys(source.totals.byUnit).join(', '))}<br>${number(source.repeatedMaterials)} materiales con filas repetidas${source.source==='ad'?`<br>Valor total AD: ${money(source.totals.totalValueDop)}`:''}${source.warnings.map(value=>`<div>${escape(value)}</div>`).join('')}${source.invalidRows.length?`<div class="inv-notice inv-error">${source.invalidRows.length} filas no válidas. Corrige el archivo o el mapeo. ${source.invalidRows.slice(0,6).map(row=>`Fila ${row.row}: ${escape(row.message)}`).join(' · ')}</div>`:''}<details><summary>Primeros materiales consolidados</summary><table class="inv-table"><thead><tr><th>Código</th><th>Cantidad</th><th>Filas sumadas</th><th>Valor AD</th></tr></thead><tbody>${source.materials.slice(0,8).map(row=>`<tr><td>${escape(row.code)}</td><td>${number(row.quantity)}</td><td>${row.rowCount}</td><td>${source.source==='ad'?money(row.totalValueDop):'No se usa FOB CT'}</td></tr>`).join('')}</tbody></table></details>`;
      document.getElementById('inventoryConfirmImport').disabled=!source.valid;
    }catch(error){pending.preview=null;document.getElementById('inventoryImportPreview').textContent=error.message;document.getElementById('inventoryConfirmImport').disabled=true;}
  }
  async function confirmImport(){
    previewImport();if(!pending?.preview?.valid)throw new Error('Completa una vista previa válida.');
    const snapshot=pending.preview;snapshot.semanticHash=await sha(snapshot.semanticKey);delete snapshot.semanticKey;
    snapshot.worksheet=pending.sheets[pending.sheetIndex].name;snapshot.importedBy=user?.email||'';
    const previous=data.sources[snapshot.source];
    if(previous?.semanticHash===snapshot.semanticHash){cancelImport();notify('Este reporte ya está cargado. Se reconoció por su contenido; no se sumó de nuevo.');return;}
    data.sources[snapshot.source]=snapshot;
    data.settings.importLog=[{source:snapshot.source,fileName:snapshot.fileName,asOf:snapshot.asOf,semanticHash:snapshot.semanticHash,at:snapshot.importedAt,by:snapshot.importedBy,materials:snapshot.materials.length},...(data.settings.importLog||[])].slice(0,30);
    pending=null;document.getElementById('inventoryImportDialog')?.remove();filters.page=0;render();
    notify('Reporte consolidado. Guardando el inventario compartido…');await save();notify('Reporte guardado. Las otras sesiones autorizadas recibirán la actualización.');
  }
  function cancelImport(){pending=null;document.getElementById('inventoryImportDialog')?.remove();if(sync){sync.markDirty();sync.refresh().catch(()=>{});}}
  function openCase(code){
    const row=report().rows.find(item=>item.code===code);if(!row)return;editor=code;const follow=getCase(code),editable=can('editar');
    modal('inventoryCaseDialog',`<h3>${escape(row.code)} · ${escape(row.description)}</h3><p class="inv-muted">CT ${number(row.ctQty)} · AD ${number(row.adQty)} · Diferencia ${number(row.difference)} · Impacto ${money(row.impactDop)}</p><div class="inv-form"><label>Estado<select id="inventoryCase_status" ${!editable?'disabled':''}>${Object.entries(workflow).map(([key,label])=>`<option value="${key}" ${key===follow.status?'selected':''}>${label}</option>`).join('')}</select></label><label>Responsable<input id="inventoryCase_owner" value="${escape(follow.owner)}" ${!editable?'readonly':''}></label><label>Referencia de ajuste / soporte<input id="inventoryCase_document" value="${escape(follow.document)}" ${!editable?'readonly':''}></label><label class="inv-wide">Notas de revisión<textarea id="inventoryCase_note" rows="4" ${!editable?'readonly':''}>${escape(follow.note)}</textarea></label></div><p class="inv-muted">Registrar el seguimiento no cambia las cantidades de los reportes. Un ajuste debe verificarse en una nueva carga.</p><div class="inv-actions">${editable?'<button class="btn btn-primary" data-inv="save-case">Guardar seguimiento</button>':''}<button class="btn btn-outline" data-inv="close-case">Cerrar</button></div><details><summary>Detalle de filas consolidadas CT / AD</summary>${['ct','ad'].map(source=>`<h4 style="margin-top:12px">${source.toUpperCase()}</h4><div class="table-wrap"><table class="inv-table"><thead><tr><th>Fila</th><th>Almacén</th><th>Lote</th><th>Cantidad</th><th>Valor DOP</th></tr></thead><tbody>${(row[source]?.details||[]).map(detail=>`<tr><td>${detail.row}</td><td>${escape(detail.warehouse)}</td><td>${escape(detail.batch)}</td><td>${number(detail.quantity)} ${escape(detail.unit)}</td><td>${source==='ad'?money(detail.totalValueDop):'—'}</td></tr>`).join('')||'<tr><td colspan="5">No registrado en esta fuente.</td></tr>'}</tbody></table></div>`).join('')}</details>`);
  }
  function readCaseForm(){return Object.fromEntries(['status','owner','document','note'].map(key=>[key,document.getElementById('inventoryCase_'+key)?.value||'']));}
  async function saveCase(){
    if(!can('editar')||!editor)throw new Error('No tienes permiso para cambiar el seguimiento.');
    const fields=readCaseForm();if(!workflow[fields.status])throw new Error('Estado no válido.');
    if(['rectificado','justificado'].includes(fields.status)&&(!fields.note.trim()||!fields.document.trim()))throw new Error('Indica una nota y una referencia de soporte para documentar la rectificación o justificación.');
    data.cases[caseKey(editor)]={...fields,updatedAt:new Date().toISOString(),updatedBy:user?.email||''};editor=null;document.getElementById('inventoryCaseDialog')?.remove();renderTable();await save();
  }
  function downloadBackup(){if(!sync)return;const blob=new Blob([JSON.stringify(sync.exportPending(),null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='TMS-inventario-pendiente.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
  function exportExcel(){
    if(!can('ver'))return;const result=report();if(!result.ready)throw new Error('Carga ambos reportes antes de exportar.');
    const rows=filteredRows().map(row=>{const follow=getCase(row.code);return [row.code,row.description,row.unit,row.ctQty,row.adQty,row.difference,row.ad?.totalValueDop??null,row.unitCostDop,row.impactDop,row.signedImpactDop,statuses[row.status],workflow[follow.status],follow.owner,follow.document,follow.note,follow.updatedAt||'',row.issues.join(' '),row.ct?.rowCount||0,row.ad?.rowCount||0];});
    const book=XLSX.utils.book_new(),sheet=XLSX.utils.aoa_to_sheet([['Material','Descripción','Unidad','Cantidad CT','Cantidad AD','CT menos AD','Valor total AD DOP','Costo ponderado AD DOP','Impacto absoluto DOP','Impacto neto DOP','Resultado','Seguimiento','Responsable','Soporte','Notas','Actualizado','Revisión necesaria','Filas CT','Filas AD'],...rows]);
    sheet['!cols']=[{wch:18},{wch:45},{wch:14},...Array.from({length:7},()=>({wch:23})),...Array.from({length:9},()=>({wch:28}))];
    Object.entries(sheet).forEach(([key,cell])=>{if(key[0]!=='!'&&cell.t==='n')cell.z=/^H/.test(key)?'#,##0.000000':/^[G-J]/.test(key)?'#,##0.00':'#,##0.###';});
    XLSX.utils.book_append_sheet(book,sheet,'Conciliacion');
    const sources=[['Fuente','Archivo','Fecha corte','Filas','Materiales','Cantidad','Valor total DOP','Huella de contenido'],...['ct','ad'].map(key=>{const source=data.sources[key];return [key.toUpperCase(),source.fileName,source.asOf,source.acceptedRowCount,source.materials.length,source.totals.quantity,source.totals.totalValueDop,source.semanticHash];}),[],['Método','Costo ponderado = valor total AD consolidado / cantidad AD consolidada. Impacto absoluto = abs(CT - AD) × costo ponderado.'],['Sin valoración','Una celda vacía de costo/impacto indica que no se puede valorar con AD; no representa costo cero.'],['Alcance','Mercancía de AD depositada en CT.'],['Fechas',result.issues.join(' ')||'Fechas de corte iguales.']];
    XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(sources),'Fuentes y metodo');XLSX.writeFile(book,'TMS_Conciliacion_CT_AD_'+data.sources.ct.asOf+'.xlsx');
  }
  document.addEventListener('change',event=>{
    const target=event.target;
    if(target.dataset?.invFile){prepareImport(target.dataset.invFile,target.files[0]).catch(error=>notify(error.message));target.value='';}
    if(target.id==='inventorySheet'&&pending){pending.sheetIndex=Number(target.value);renderImportDialog();}
    if(target.id==='inventoryHeaderRow'&&pending){const index=Number(target.value)-1,headers=pending.sheets[pending.sheetIndex].matrix[index]||[];document.querySelectorAll('[id^="inventoryMap_"]').forEach(select=>{select.innerHTML='<option value="">No indicada</option>'+headers.map((value,col)=>`<option value="${col}">${col+1}. ${escape(value||'Sin título')}</option>`).join('');});}
    if(target.closest('#inventoryImportDialog')&&target.id!=='inventorySheet'){pending.preview=null;document.getElementById('inventoryConfirmImport').disabled=true;}
    if(['inventoryStatus','inventoryFollowup'].includes(target.id)){filters[target.id==='inventoryStatus'?'status':'followup']=target.value;filters.page=0;renderTable();}
  });
  document.addEventListener('input',event=>{if(event.target.id==='inventorySearch'){filters.search=event.target.value;filters.page=0;renderTable();}if(event.target.closest('#inventoryCaseDialog'))sync?.markDirty();});
  document.addEventListener('click',event=>{
    const button=event.target.closest('[data-inv]');if(!button||button.disabled)return;
    const actions={preview:previewImport,confirm:confirmImport,'cancel-import':cancelImport,'case':()=>openCase(button.dataset.code),'save-case':saveCase,'close-case':()=>{editor=null;document.getElementById('inventoryCaseDialog')?.remove();sync?.markDirty();sync?.refresh().catch(()=>{});},export:exportExcel,backup:downloadBackup,prev:()=>{filters.page--;renderTable();},next:()=>{filters.page++;renderTable();},reload:async()=>{if(syncState.conflict&&!confirm('¿Cargar la versión compartida? Puedes descargar tu copia antes de continuar.'))return;pending=null;editor=null;document.getElementById('inventoryImportDialog')?.remove();document.getElementById('inventoryCaseDialog')?.remove();if(sync?.getState().revision===null)await start();else if(syncState.conflict)await sync.discardAndReload();else if(sync?.getState().dirty)await save();else await sync?.refresh();notify('Versión compartida del inventario actualizada.');}};
    try{Promise.resolve(actions[button.dataset.inv]?.()).catch(error=>notify(error.message));}catch(error){notify(error.message);}
  });
  window.addEventListener('beforeunload',event=>{if(hasDraft()||sync?.getState().dirty){event.preventDefault();event.returnValue='';}});
  window.addEventListener('tms-permissions-changed',()=>{
    if(!can('ver')&&sync){sync.stop();data=empty();user=null;pending=null;editor=null;document.getElementById('inventoryImportDialog')?.remove();document.getElementById('inventoryCaseDialog')?.remove();render();}
  });
  client()?.auth.onAuthStateChange?.((event,session)=>{if(user&&(event==='SIGNED_OUT'||session?.user&&session.user.id!==user.id)){sync?.stop();data=empty();user=null;}});
  window.guardarCambiosVistaActual=function(){if(document.querySelector('.view.active')?.id==='inventario')return save().catch(error=>notify(error.message));return window.guardarEnSupabase();};
  window.TMSInventory={open,start,render,prepareImport,confirmImport,save,getReport:report,getState:()=>copy(data),getSyncState:()=>sync?.getState()||syncState};
})();

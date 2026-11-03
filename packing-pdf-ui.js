(function(){
 'use strict';
 const P=window.TMSPackingPDFDomain,esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const state={documents:[],audit:[],bases:new Map(),busy:false,generation:0};let options;
 const read=(line,field)=>field==='FACTURA'?line.source.invoice:field==='CANTIDAD_FACTURA'?line.data.CANTIDAD:line.data[field];
 const write=(line,field,value)=>{if(field==='FACTURA')line.source.invoice=value;else line.data[field]=value;};
 const same=(a,b)=>String(a??'')===String(b??'');
 const editable=()=>options.can('editar')&&options.can('importar')&&!options.isBusy()&&!state.busy;
 const fieldKey=(id,field)=>JSON.stringify([id,field]);
 const notify=message=>{document.getElementById('packingPdfStatus').textContent=message;};
 const changed=()=>options.onChange();
 function analyze(doc){return P.analyze(doc,options.getLines());}
 function init(config){
  options=config;document.getElementById('packingInvoiceMount').innerHTML=`<section class="card"><div class="packing-section-head"><div><h2>3 · Facturas PDF</h2><p>Lee las facturas y cruza sus referencias con la planilla del consolidado.</p></div><span>Lectura local</span></div><label class="packing-upload packing-pdf-upload">Agregar facturas PDF<input id="packingInvoiceFiles" type="file" accept=".pdf,application/pdf" multiple><small>Los datos detectados se revisan antes de incorporarlos al packing list.</small></label><p id="packingPdfStatus" role="status"></p><div id="packingPdfDocuments" class="packing-pdf-documents"></div></section>`;
  document.getElementById('packingInvoiceFiles').onchange=event=>importFiles([...event.target.files]);
  document.getElementById('packingPdfDocuments').onclick=event=>{const open=event.target.closest('[data-pdf-open]'),remove=event.target.closest('[data-pdf-remove]');if(open)openReview(open.dataset.pdfOpen);if(remove)removeDocument(remove.dataset.pdfRemove);};refresh();
 }
 function refresh(){
  if(!options)return;const input=document.getElementById('packingInvoiceFiles');if(!input)return;input.disabled=!options.can('importar')||state.busy||options.isBusy();
  document.getElementById('packingPdfDocuments').innerHTML=state.documents.map(doc=>{const a=analyze(doc);return `<article><div><strong>${esc(doc.name)}</strong><p>${doc.pages.length} páginas · ${a.facts.length} datos propuestos${doc.unreadablePages.length?' · Sin texto en páginas '+doc.unreadablePages.join(', '):''}</p>${!options.getLines().length?'<p>Carga la planilla para cruzar los artículos.</p>':''}</div><div class="packing-pdf-actions"><button class="btn btn-outline" data-pdf-open="${doc.hash}" ${state.busy||options.isBusy()?'disabled':''}>Revisar lectura</button><button class="btn btn-outline" data-pdf-remove="${doc.hash}" ${!options.can('importar')||state.busy||options.isBusy()?'disabled':''}>Retirar PDF</button></div></article>`;}).join('');
 }
 async function importFiles(files){
  if(!files.length||!options.can('importar')||state.busy||options.isBusy())return;state.busy=true;const generation=state.generation,account=options.account();notify('Leyendo facturas PDF…');changed();
  try{
   const documents=[],skipped=[];
   for(const file of files){
    if(!/\.pdf$/i.test(file.name))throw new Error('Selecciona un archivo PDF: '+file.name);
    if(file.size>50*1024*1024)throw new Error(file.name+' supera 50 MB. Divide el documento.');
    const buffer=await file.arrayBuffer(),hash=[...new Uint8Array(await crypto.subtle.digest('SHA-256',buffer))].map(v=>v.toString(16).padStart(2,'0')).join('');
    if([...state.documents,...documents].some(doc=>doc.hash===hash)){skipped.push(file.name);continue;}
    if([...state.documents,...documents].some(doc=>doc.name===file.name))throw new Error('Ya existe '+file.name+'. Retira su versión anterior antes de cargarla de nuevo.');
    const doc=await TMSPackingPDFReader.read(buffer,file.name,(page,total)=>notify('Leyendo '+file.name+' · página '+page+' de '+total+'…'));
    if(generation!==state.generation||account!==options.account()||!options.can('importar'))throw new Error('El acceso o la sesión cambió. No se incorporó la selección.');
    documents.push({...doc,hash,invoiceOverride:''});
   }
   state.documents.push(...documents);notify(documents.length+' PDF leídos.'+(skipped.length?' Se omitieron '+skipped.length+' archivos ya cargados.':'')+(documents.some(doc=>doc.unreadablePages.length)?' Hay páginas sin texto; requieren OCR o una copia con texto seleccionable.':''));
  }catch(error){if(generation===state.generation)notify('No se incorporó esta selección: '+error.message);}finally{if(generation===state.generation){state.busy=false;document.getElementById('packingInvoiceFiles').value='';changed();}}
 }
 function removeDocument(hash){
  if(!options.can('importar')||state.busy||options.isBusy())return;
  if(state.audit.some(entry=>entry.docHash===hash&&entry.active)&&!options.can('editar')){notify('Necesitas permiso de edición para retirar un PDF cuyos datos ya se incorporaron.');return;}
  const keys=new Set(state.audit.filter(entry=>entry.docHash===hash&&entry.active).map(entry=>fieldKey(entry.lineId,entry.field)));
  for(const key of keys){
   const [id,field]=JSON.parse(key),line=options.getLines().find(line=>line.id===id),all=state.audit.filter(entry=>entry.active&&entry.lineId===id&&entry.field===field),remaining=all.filter(entry=>entry.docHash!==hash);
   if(line&&same(read(line,field),all.at(-1)?.value))write(line,field,remaining.length?remaining.at(-1).value:state.bases.get(key));
   if(!remaining.length)state.bases.delete(key);
  }
  state.audit.filter(entry=>entry.docHash===hash).forEach(entry=>{entry.active=false;entry.status='Documento retirado';});state.documents=state.documents.filter(doc=>doc.hash!==hash);notify('PDF retirado.'+(keys.size?' Sus datos vigentes se restauraron a los valores anteriores.':''));changed();
 }
 function invalidate(lineId,fields){
  fields.forEach(field=>{state.audit.filter(entry=>entry.active&&entry.lineId===lineId&&entry.field===field).forEach(entry=>{entry.active=false;entry.status='Modificado manualmente';});state.bases.delete(fieldKey(lineId,field));});
 }
 function prune(){const ids=new Set(options.getLines().map(line=>line.id));state.audit=state.audit.filter(entry=>ids.has(entry.lineId));for(const key of state.bases.keys())if(!ids.has(JSON.parse(key)[0]))state.bases.delete(key);}
 function clear(){state.generation++;state.documents=[];state.audit=[];state.bases.clear();state.busy=false;document.getElementById('packingPdfDialog')?.close();if(document.getElementById('packingPdfStatus'))notify('');refresh();}
 function openReview(hash){
  if(!options.can('ver')||state.busy||options.isBusy())return;const doc=state.documents.find(doc=>doc.hash===hash);if(!doc)return;
  const result=analyze(doc);let modal=document.getElementById('packingPdfDialog');if(!modal){modal=document.createElement('dialog');modal.id='packingPdfDialog';modal.className='calendar-dispatch-dialog';modal.setAttribute('aria-labelledby','packingPdfTitle');document.body.appendChild(modal);}
  modal.innerHTML=`<header><div><p>Lectura de factura PDF</p><h2 id="packingPdfTitle">${esc(doc.name)}</h2><p>${doc.pages.length} páginas · Factura detectada: ${esc(result.invoiceNos.join(' / ')||'No identificada')}</p></div><button class="btn btn-outline" data-pdf-close>Cerrar</button></header>${doc.unreadablePages.length?`<p class="packing-pdf-warning">Sin texto en páginas ${doc.unreadablePages.join(', ')}. Pueden ser escaneos o páginas en blanco; se necesita OCR o una copia con texto para leer su contenido.</p>`:''}<label class="packing-pdf-invoice-label">Número de factura para el cruce con la planilla<input id="packingPdfInvoiceOverride" class="form-control" value="${esc(doc.invoiceOverride)}" placeholder="Usar el número detectado" ${!editable()?'disabled':''}></label><p>Revisa cada valor y su línea de destino. Los pesos de bulto solo se proponen cuando el PDF los identifica como tales y especifica la unidad.</p><div class="packing-pdf-actions"><button class="btn btn-outline" data-pdf-select-new ${!editable()?'disabled':''}>Seleccionar datos nuevos</button><button class="btn btn-primary" data-pdf-apply ${!editable()?'disabled':''}>Aplicar datos seleccionados</button></div><p id="packingPdfReviewStatus" role="status"></p><div class="table-wrap"><table class="data-table packing-pdf-facts"><thead><tr><th>Usar</th><th>Artículo / página</th><th>Campo</th><th>Valor leído</th><th>Línea de destino</th><th>Valor actual</th><th>Evidencia</th></tr></thead><tbody>${result.facts.map((fact,index)=>`<tr data-pdf-fact="${index}"><td><input type="checkbox" data-pdf-selected aria-label="Aplicar ${esc(fact.field)} de página ${fact.page}"></td><td>${esc(fact.reference)}<small>Página ${fact.page}</small></td><td>${esc(fact.field==='FACTURA'?'Factura de origen':fact.field==='CANTIDAD_FACTURA'?'Cantidad en factura':fact.field)}</td><td>${esc(fact.value)}</td><td><select class="form-control" data-pdf-target aria-label="Línea de destino para ${esc(fact.field)}">${fact.targetIds.length!==1?'<option value="">'+(fact.targetIds.length?'Elige una línea':'Sin coincidencia de factura y referencia')+'</option>':''}${fact.targetIds.map(id=>{const line=options.getLines().find(line=>line.id===id);return `<option value="${esc(id)}">${esc(line.data.ARTICULO)} · ${esc(line.source.invoice||'Sin factura')} · ${esc(line.data.ALBARAN)} · ${line.data.CANTIDAD}</option>`;}).join('')}</select></td><td data-pdf-current></td><td class="packing-pdf-evidence">${esc(fact.evidence)}</td></tr>`).join('')||'<tr><td colspan="7">No se identificaron campos con referencia y etiqueta inequívocas. Puedes consultar el texto leído abajo y completar la planilla, o aportar una factura de ejemplo para adaptar su formato.</td></tr>'}</tbody></table></div><details class="packing-pdf-text"><summary>Ver texto leído por página</summary>${doc.pages.map(page=>`<section><h3>Página ${page.page}</h3><pre>${esc(page.text)||'Sin texto seleccionable.'}</pre></section>`).join('')}</details>`;
  const headings=[...modal.querySelectorAll('.packing-pdf-facts thead th')].map(th=>th.textContent.trim());
  modal.querySelectorAll('.packing-pdf-facts tbody tr').forEach(row=>{if(row.cells.length===headings.length)[...row.cells].forEach((cell,index)=>cell.dataset.mobileLabel=headings[index]);});
  const reviewStatus=modal.querySelector('#packingPdfReviewStatus');
  function rowState(row){const fact=result.facts[Number(row.dataset.pdfFact)],select=row.querySelector('[data-pdf-target]'),line=options.getLines().find(line=>line.id===select.value),checkbox=row.querySelector('[data-pdf-selected]'),current=row.querySelector('[data-pdf-current]');current.textContent=line?(read(line,fact.field)??''):'—';if(fact.readOnly){current.textContent+=line&&!same(read(line,fact.field),fact.value)?' · Difiere: revisar cantidad':' · Solo comparación';}checkbox.disabled=!!fact.readOnly||!editable()||!line||same(read(line,fact.field),fact.value);select.disabled=!editable()||!fact.targetIds.length;if(checkbox.disabled)checkbox.checked=false;}
  modal.querySelectorAll('[data-pdf-fact]').forEach(row=>{rowState(row);row.querySelector('select').onchange=()=>rowState(row);});
  modal.querySelector('[data-pdf-close]').onclick=()=>modal.close();
  modal.querySelector('#packingPdfInvoiceOverride').onchange=event=>{doc.invoiceOverride=event.target.value.trim();openReview(hash);};
  modal.querySelector('[data-pdf-select-new]').onclick=()=>{
   const values=new Map();result.facts.forEach(fact=>{if(fact.targetIds.length===1){const key=fieldKey(fact.targetIds[0],fact.field);if(!values.has(key))values.set(key,new Set());values.get(key).add(String(fact.value));}});
   modal.querySelectorAll('[data-pdf-fact]').forEach(row=>{const fact=result.facts[Number(row.dataset.pdfFact)],line=options.getLines().find(line=>line.id===row.querySelector('select').value),input=row.querySelector('[data-pdf-selected]');input.checked=!input.disabled&&!!line&&(read(line,fact.field)==null||read(line,fact.field)==='')&&values.get(fieldKey(line.id,fact.field))?.size===1;});
  };
  modal.querySelector('[data-pdf-apply]').onclick=()=>{
   if(!editable())return;const selected=[],keys=new Map();
   for(const row of modal.querySelectorAll('[data-pdf-fact]')){
    if(!row.querySelector('[data-pdf-selected]').checked)continue;const fact=result.facts[Number(row.dataset.pdfFact)],id=row.querySelector('[data-pdf-target]').value,line=options.getLines().find(line=>line.id===id);if(fact.readOnly||!line||!fact.targetIds.includes(id)){reviewStatus.textContent='Selecciona una línea de destino y un campo válido.';return;}
    const key=fieldKey(id,fact.field);if(keys.has(key)&&!same(keys.get(key),fact.value)){reviewStatus.textContent='Hay dos valores distintos seleccionados para el mismo campo y línea. Elige uno.';return;}keys.set(key,fact.value);selected.push({fact,line,key});
   }
   let count=0;for(const {fact,line,key}of selected){const before=read(line,fact.field);if(same(before,fact.value))continue;if(!state.bases.has(key))state.bases.set(key,before);write(line,fact.field,fact.value);state.audit.push({...fact,lineId:line.id,article:line.data.ARTICULO,before,at:new Date().toISOString(),active:true,status:'Aplicado'});count++;}
   changed();modal.querySelectorAll('[data-pdf-fact]').forEach(row=>rowState(row));reviewStatus.textContent=count+' datos incorporados al packing list.';
  };
  modal.onclose=()=>{if(!modal.open)modal.replaceChildren();};if(!modal.open)modal.showModal();
 }
 function addSheets(workbook,report){
  if(!state.documents.length&&!state.audit.length)return;
  const audit=workbook.addWorksheet('Datos de facturas PDF');audit.columns=[{header:'Factura PDF',key:'file',width:40},{header:'Página',key:'page',width:10},{header:'Fila PDF',key:'row',width:12},{header:'Artículo',key:'article',width:20},{header:'Campo',key:'field',width:24},{header:'Valor',key:'value',width:25},{header:'Valor anterior',key:'before',width:25},{header:'Estado',key:'status',width:32},{header:'Referencia leída',key:'reference',width:28},{header:'Fecha de revisión',key:'at',width:28}];
  const byId=new Map(report.lines.map(line=>[line.id,line])),latest=new Map();state.audit.filter(entry=>entry.active).forEach(entry=>latest.set(fieldKey(entry.lineId,entry.field),entry));state.audit.filter(entry=>byId.has(entry.lineId)).forEach(entry=>{const current=entry.active&&latest.get(fieldKey(entry.lineId,entry.field))===entry&&same(read(byId.get(entry.lineId),entry.field),entry.value);audit.addRow({...entry,status:current?'Vigente':entry.active?'Sustituido por otro PDF':entry.status});});
  audit.views=[{state:'frozen',ySplit:1}];audit.getRow(1).font={bold:true,color:{argb:'FFFFFF'}};audit.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'17384C'}};audit.eachRow((row,index)=>{row.alignment={wrapText:true,vertical:'top'};row.height=index===1?28:38;});
 }
 window.TMSPackingPDF={init,refresh,importFiles,isBusy:()=>state.busy,clear,prune,invalidate,addSheets,getDocuments:()=>state.documents.map(doc=>({name:doc.name,pages:doc.pages.length,unreadablePages:[...doc.unreadablePages]})),getIssues:()=>state.documents.filter(doc=>doc.unreadablePages.length).map(doc=>({id:doc.name,article:'',field:'Lectura PDF',message:'Páginas sin texto '+doc.unreadablePages.join(', ')+': pendientes de OCR o de verificar que estén en blanco.'}))};
})();

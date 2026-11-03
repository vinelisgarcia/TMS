(function(){
 'use strict';
 const D=window.TMSPackingDomain,escape=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const state={lines:[],files:[],template:null,shipment:{reference:'',container:'',seal:'',customer:'Álvarez Dominicana',incoterm:'',orderColumn:''},busy:false};
 let excelPromise=null;
 const can=action=>!!window.TMSAuditCore?.hasPermission('packingMiami',action);
 function status(message){document.getElementById('packingStatus').textContent=message;}
 function report(){return D.build(state.lines,state.shipment);}
 function mount(){
  const view=document.getElementById('packingMiami');if(!view)return;
  view.innerHTML=`<div class="view-title">Packing list Miami</div><div class="view-subtitle">Prepara el packing list del consolidado a partir de sus documentos.</div><section class="card packing-inputs"><div class="packing-section-head"><div><h2>Documentos del consolidado</h2><p>Planillas de entrada, archivos con las 23 columnas del packing list o CSV.</p></div><span>Documentos de esta sesión</span></div><div class="packing-upload-grid"><label class="packing-upload">1 · Plantilla de salida <small>Opcional. El formato de 23 columnas ya está disponible.</small><input id="packingTemplateFile" type="file" accept=".xlsx"><span id="packingTemplateName">Formato de packing list · 23 columnas</span></label><label class="packing-upload">2 · Documentos de mercancía <small>Puedes agregar varios archivos al mismo consolidado.</small><input id="packingSourceFiles" type="file" multiple accept=".xlsx,.xls,.csv"></label></div><div id="packingFiles" class="packing-files"></div><p id="packingStatus" role="status"></p><div class="packing-field-grid">${[['reference','Referencia del consolidado'],['container','Número de contenedor'],['seal','Número de precinto'],['customer','Cliente / destinatario'],['incoterm','Incoterm y lugar']].map(([key,label])=>`<label>${label}<input id="packingField_${key}" data-packing-field="${key}" class="form-control" value="${escape(state.shipment[key])}"></label>`).join('')}<label>PEDIDO DOM corresponde a<select id="packingOrderColumn" class="form-control"><option value="">Pendiente de definir</option><option value="PEDIDO">PEDIDO</option><option value="PEDIDO_CLIENTE">PEDIDO_CLIENTE</option><option value="NPI">NPI</option></select></label></div></section><section class="card"><div class="packing-section-head"><div><h2>Revisión del packing list</h2><p id="packingReadiness">Agrega documentos para comenzar.</p></div><div class="packing-actions"><button class="btn btn-outline" id="packingDraft" disabled>Descargar borrador Excel</button><button class="btn btn-primary" id="packingFinal" disabled>Descargar packing list Excel</button></div></div><dl id="packingMetrics" class="packing-metrics"></dl><div class="table-wrap"><table class="data-table packing-table"><thead><tr><th>Artículo</th><th>Descripción</th><th>Cantidad</th><th>Pedido DOM</th><th>Factura / albarán</th><th>Bulto</th><th>Revisión</th></tr></thead><tbody id="packingRows"></tbody></table></div><details class="packing-issues"><summary>Datos pendientes y criterios</summary><p>Pesos en kg y medidas en metros. El peso bruto y el volumen se cuentan una vez por bulto; el peso bruto de línea se prorratea por peso neto cuando hay datos suficientes. No se rellenan pesos, origen ni partidas arancelarias sin respaldo.</p><div id="packingIssues"></div></details></section>`;
  view.querySelectorAll('[data-packing-field]').forEach(input=>input.addEventListener('input',()=>{state.shipment[input.dataset.packingField]=input.value.trim();render();}));
  document.getElementById('packingOrderColumn').onchange=event=>{state.shipment.orderColumn=event.target.value;render();};
  document.getElementById('packingSourceFiles').onchange=event=>importFiles([...event.target.files]);
  document.getElementById('packingTemplateFile').onchange=event=>readTemplate(event.target.files[0]);
  document.getElementById('packingDraft').onclick=()=>exportFile(false);document.getElementById('packingFinal').onclick=()=>exportFile(true);
  document.getElementById('packingRows').onclick=event=>{const button=event.target.closest('[data-packing-edit]');if(button)editLine(button.dataset.packingEdit);};
  document.getElementById('packingFiles').onclick=event=>{const button=event.target.closest('[data-packing-remove]');if(!button||!can('importar')||state.busy)return;const hash=button.dataset.packingRemove;state.files=state.files.filter(file=>file.hash!==hash);state.lines=state.lines.filter(line=>line.hash!==hash);render();};
  render();
 }
 function render(){
  const r=report();document.getElementById('packingFiles').innerHTML=state.files.map(file=>`<div><span><strong>${escape(file.name)}</strong><small>${file.count} líneas</small></span><button class="btn btn-outline" data-packing-remove="${file.hash}" ${state.busy||!can('importar')?'disabled':''} aria-label="Retirar ${escape(file.name)}">Retirar</button></div>`).join('');
  document.getElementById('packingMetrics').innerHTML=[['Líneas',r.lines.length],['Unidades',r.quantity.toLocaleString('es-DO')],['Bultos',r.packages??'Pendiente'],['Peso bruto kg',r.gross===null?'Pendiente':r.gross.toLocaleString('es-DO',{maximumFractionDigits:3})]].map(([label,value])=>`<div><dt>${label}</dt><dd>${value}</dd></div>`).join('');
  document.getElementById('packingRows').innerHTML=r.lines.map(line=>`<tr><td>${escape(line.data.ARTICULO)}</td><td>${escape(line.data.NOMBRE)}</td><td>${line.data.CANTIDAD.toLocaleString('es-DO')}</td><td>${escape(line.source.pedidoDom)}</td><td>${escape(line.source.invoice)}<small>${escape(line.data.ALBARAN)}</small></td><td>${escape(line.data.CAJAS_ID)||'Pendiente'}</td><td><button class="btn btn-outline" data-packing-edit="${escape(line.id)}" ${!can('editar')?'disabled':''}>Completar datos</button></td></tr>`).join('')||'<tr><td colspan="7">Sin documentos cargados.</td></tr>';
  document.getElementById('packingReadiness').textContent=!r.lines.length?'Agrega documentos para comenzar.':r.complete?'Datos completos. Puedes descargar el packing list.':r.issues.length+' registros con datos pendientes. Puedes descargar un borrador o completar los datos.';
  document.getElementById('packingIssues').innerHTML=r.issues.map(issue=>`<p><strong>${escape(issue.article||issue.id)}:</strong> ${escape(issue.field)} · ${escape(issue.message)}</p>`).join('')||'<p>Sin datos pendientes.</p>';
  document.getElementById('packingDraft').disabled=state.busy||!can('ver')||!r.lines.length;document.getElementById('packingFinal').disabled=state.busy||!can('ver')||!r.complete;
  document.getElementById('packingSourceFiles').disabled=state.busy||!can('importar');document.getElementById('packingTemplateFile').disabled=state.busy||!can('importar');
  document.querySelectorAll('#packingMiami [data-packing-field],#packingOrderColumn').forEach(input=>input.disabled=!can('editar'));
 }
 function worksheetGrid(sheet){
  const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:true});
  // Keep explicitly zero-padded identifiers; numeric quantities stay numeric.
  grid.forEach((row,r)=>row.forEach((value,c)=>{const cell=sheet[XLSX.utils.encode_cell({r,c})];if(cell?.t==='n'&&/^0{2,}$/.test(cell.z||'')&&cell.w)row[c]=cell.w;}));return grid;
 }
 async function importFiles(files){
  if(!can('importar')||state.busy)return;state.busy=true;render();status('Leyendo documentos…');
  try{
   const pending=[],newFiles=[],skipped=[];
   for(const file of files){
    if(!/\.(xlsx|xls|csv)$/i.test(file.name))throw new Error('Formato no admitido: '+file.name);
    const buffer=await file.arrayBuffer(),hash=[...new Uint8Array(await crypto.subtle.digest('SHA-256',buffer))].map(v=>v.toString(16).padStart(2,'0')).join('');
    if([...state.files,...newFiles].some(item=>item.hash===hash)){skipped.push(file.name);continue;}
    if([...state.files,...newFiles].some(item=>item.name===file.name))throw new Error('Ya existe '+file.name+'. Retira su versión anterior antes de cargarla de nuevo.');
    const wb=XLSX.read(buffer,{type:'array',cellNF:true,raw:true}),parsed=[];let found=false;
    for(const name of wb.SheetNames){try{const result=D.parse(worksheetGrid(wb.Sheets[name]),file.name,name);found=true;if(result.errors.length)throw new Error(result.errors.map(error=>'Fila '+error.row+': '+error.message).join(' · '));parsed.push(...result.lines);}catch(error){if(!error.message.includes('no se identifican las columnas'))throw error;}}
    if(!found||!parsed.length)throw new Error(file.name+': no contiene materiales con artículo y cantidad válidos.');
    parsed.forEach(line=>pending.push({...line,id:hash+'|'+line.id,hash}));newFiles.push({name:file.name,hash,count:parsed.length});
   }
   if(!can('importar'))throw new Error('El permiso de importar cambió. No se incorporaron documentos.');
   state.lines.push(...pending);state.files.push(...newFiles);status(pending.length+' líneas incorporadas.'+(skipped.length?' No se duplicaron '+skipped.length+' archivos ya cargados.':''));
  }catch(error){status('No se incorporó esta selección: '+error.message);}finally{state.busy=false;document.getElementById('packingSourceFiles').value='';render();}
 }
 async function readTemplate(file){
  if(!file||!can('importar'))return;
  try{const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellStyles:true});const name=wb.SheetNames.find(name=>{const grid=worksheetGrid(wb.Sheets[name]);return grid.some(row=>D.columns.every((col,i)=>String(row[i]||'').trim()===col));});if(!name)throw new Error('La plantilla debe contener las 23 columnas del modelo, en el mismo orden.');state.template={name:file.name,sheetName:name};document.getElementById('packingTemplateName').textContent=file.name+' · 23 columnas verificadas';status('Plantilla reconocida. Sus datos de ejemplo no se incorporan como mercancía.');}catch(error){status('No se cambió la plantilla: '+error.message);}finally{document.getElementById('packingTemplateFile').value='';}
 }
 function editLine(id){
  if(!can('editar'))return;const line=state.lines.find(line=>line.id===id);if(!line)return;
  let modal=document.getElementById('packingEditDialog');if(!modal){modal=document.createElement('dialog');modal.id='packingEditDialog';modal.className='calendar-dispatch-dialog';document.body.appendChild(modal);}
  const editable=D.columns.filter(key=>!['NETO TOTAL','PESOBRUTO_LINEA','VOLUMEN'].includes(key));
  modal.innerHTML=`<header><div><p>Completar datos del material</p><h2>${escape(line.data.ARTICULO)}</h2><p>${escape(line.source.file)} · ${escape(line.source.sheet)} · fila ${line.source.row}</p></div><button class="btn btn-outline" data-packing-cancel>Cerrar</button></header><form id="packingEditForm"><div class="packing-field-grid">${editable.map(key=>`<label>${escape(key)}${['LARGO','ANCHO','ALTO'].includes(key)?' (m)':key==='KG NETO UD'||key==='PESOBRUTO'?' (kg)':''}<input class="form-control" name="${escape(key)}" ${D.numeric.has(key)?'type="number" step="any" min="0"':''} value="${escape(line.data[key]??'')}"></label>`).join('')}</div><label class="packing-copy"><input type="checkbox" id="packingCopyPackage">Aplicar peso bruto, tipo y medidas a otras líneas que ya tengan este mismo bulto.</label><p id="packingEditError" role="alert"></p><div class="modal-actions"><button type="submit" class="btn btn-primary">Guardar datos del material</button></div></form>`;
  modal.querySelector('[data-packing-cancel]').onclick=()=>modal.close();modal.querySelector('form').onsubmit=event=>{
   event.preventDefault();if(!can('editar'))return modal.close();const next={...line.data};for(const [key,value] of new FormData(event.target)){next[key]=D.numeric.has(key)?D.number(value):String(value).trim();}
   if(!next.ARTICULO||!(next.CANTIDAD>0)){document.getElementById('packingEditError').textContent='Completa artículo y cantidad positiva.';return;}
   line.data=next;if(document.getElementById('packingCopyPackage').checked&&next.CAJAS_ID)state.lines.filter(row=>row.data.CAJAS_ID===next.CAJAS_ID).forEach(row=>['PESOBRUTO','TIPO CONTENEDOR','LARGO','ANCHO','ALTO'].forEach(key=>row.data[key]=next[key]));render();modal.close();
  };modal.onclose=()=>{if(!modal.open)modal.replaceChildren();};modal.showModal();
 }
 async function excelReady(){
  if(window.ExcelJS)return;if(!excelPromise)excelPromise=new Promise((resolve,reject)=>{const script=document.createElement('script');script.src='./vendor/exceljs-4.4.0.min.js';const timer=setTimeout(fail,20000);function fail(){clearTimeout(timer);script.remove();excelPromise=null;reject(new Error('No se pudo cargar el generador Excel. Reintenta la descarga.'));}script.onload=()=>{clearTimeout(timer);if(window.ExcelJS)resolve();else fail();};script.onerror=fail;document.head.appendChild(script);});await excelPromise;
 }
 async function makeWorkbook(r,final){
  await excelReady();const wb=new ExcelJS.Workbook();wb.creator='TMS Álvarez Dominicana';wb.calcProperties.fullCalcOnLoad=true;
  const ws=wb.addWorksheet(state.template?.sheetName||'Hoja1',{views:[{state:'frozen',ySplit:6}],pageSetup:{orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0}});ws.columns=D.columns.map(key=>({key,width:key==='NOMBRE'?52:key==='CAJAS_ID'?22:key==='ARTICULO'?18:19}));
  ws.mergeCells('A1:H2');ws.getCell('A1').value=final?'PACKING LIST MIAMI':'BORRADOR · DATOS PENDIENTES';ws.getCell('A1').font={size:18,bold:true,color:{argb:'17384C'}};
  ['Nº DE BULTOS','KG.','M3'].forEach((value,i)=>ws.getCell(1,i+9).value=value);[r.packages,r.gross,r.volume].forEach((value,i)=>ws.getCell(2,i+9).value=value);
  ws.getCell('I3').value=r.shipment.reference;ws.getCell('A4').value='Nº CONTENEDOR';ws.getCell('B4').value='Nº PRECINTO';ws.getCell('A5').value=r.shipment.container;ws.getCell('B5').value=r.shipment.seal;
  ws.getRow(6).values=D.columns;ws.getRow(6).height=32;ws.getRow(6).eachCell(cell=>{cell.font={name:'Calibri',bold:true,color:{argb:'FFFFFF'},size:11};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'17384C'}};cell.alignment={wrapText:true,vertical:'middle'};});
  const groups=new Map();r.lines.forEach(line=>{const key=line.data.CAJAS_ID||'missing|'+line.id;if(!groups.has(key))groups.set(key,[]);groups.get(key).push(line);});let quantityRows=[];
  for(const items of groups.values()){
   const detailRows=[];items.forEach(line=>{const row=ws.addRow(D.columns.map(key=>line.data[key]===''?null:line.data[key]??null));detailRows.push(row.number);quantityRows.push(row.number);row.height=Math.max(30,Math.ceil(line.data.NOMBRE.length/50)*16);row.eachCell((cell,index)=>{cell.font={name:'Calibri',size:11};cell.alignment={wrapText:true,vertical:'top'};if(D.numeric.has(D.columns[index-1]))cell.numFmt=Number.isInteger(cell.value)?'#,##0':'#,##0.000';else cell.numFmt='@';});});
   if(items[0].data.CAJAS_ID){const subtotal=ws.addRow({CLIENTE:items[0].data.CLIENTE,CAJAS_ID:'Total '+items[0].data.CAJAS_ID});const first=detailRows[0],last=detailRows.at(-1);if(items.every(line=>line.data['NETO TOTAL']!==null))subtotal.getCell('NETO TOTAL').value={formula:`SUM(P${first}:P${last})`,result:items.reduce((sum,line)=>sum+line.data['NETO TOTAL'],0)};if(items.every(line=>line.data.PESOBRUTO_LINEA!==null))subtotal.getCell('PESOBRUTO_LINEA').value={formula:`SUM(R${first}:R${last})`,result:items.reduce((sum,line)=>sum+line.data.PESOBRUTO_LINEA,0)};for(const key of ['PESOBRUTO','VOLUMEN']){const values=new Set(items.map(line=>line.data[key]));subtotal.getCell(key).value=values.size===1?items[0].data[key]:null;}subtotal.font={bold:true,color:{argb:'17384C'}};subtotal.fill={type:'pattern',pattern:'solid',fgColor:{argb:'EEF3F7'}};}
  }
  const total=ws.addRow({NOMBRE:'TOTAL UNIDADES',CANTIDAD:r.quantity});total.font={bold:true};if(r.shipment.incoterm)ws.addRow({NOMBRE:'INCOTERM '+r.shipment.incoterm});ws.printTitlesRow='6:6';
  const provenance=wb.addWorksheet('Datos de origen');provenance.columns=[{header:'Archivo',key:'file',width:45},{header:'Hoja',key:'sheet',width:24},{header:'Fila',key:'row',width:10},{header:'Artículo',key:'article',width:20},{header:'Pedido DOM original',key:'pedidoDom',width:24},{header:'Referencia proveedor',key:'supplierReference',width:26},{header:'Proveedor',key:'supplier',width:30},{header:'Factura',key:'invoice',width:24},{header:'Referencia embarque',key:'shipment',width:30},{header:'Cantidad original',key:'sourceQuantity',width:18},{header:'Cantidad usada',key:'quantity',width:18}];r.lines.forEach(line=>provenance.addRow({...line.source,article:line.data.ARTICULO,sourceQuantity:line.source.quantity,quantity:line.data.CANTIDAD}));
  if(!final){const pending=wb.addWorksheet('Pendientes');pending.addRow(['Artículo / registro','Campo','Revisión']);r.issues.forEach(issue=>pending.addRow([issue.article||issue.id,issue.field,issue.message]));pending.columns=[{width:28},{width:80},{width:65}];pending.eachRow(row=>{row.alignment={wrapText:true,vertical:'top'};row.height=45;});}
  provenance.views=[{state:'frozen',ySplit:1}];provenance.autoFilter='A1:K'+Math.max(1,provenance.rowCount);return new Blob([await wb.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
 }
 async function exportFile(final){
  if(!can('ver')||state.busy)return;const r=report();if(!r.lines.length||final&&!r.complete)return;state.busy=true;render();status('Preparando Excel…');
  try{const blob=await makeWorkbook(r,final);if(!can('ver'))throw new Error('Tu acceso cambió.');const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=(final?'PACKING_LIST_MIAMI_':'BORRADOR_PACKING_MIAMI_')+(r.shipment.reference||'SIN_REFERENCIA').replace(/[^a-zA-Z0-9_-]/g,'_')+'.xlsx';document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);status(final?'Packing list preparado.':'Borrador preparado. Revisa la hoja Pendientes antes de usarlo para el despacho.');}catch(error){status(error.message);}finally{state.busy=false;render();}
 }
 window.TMSPacking={refreshPermissions(){if(!can('ver')){state.lines=[];state.files=[];document.getElementById('packingEditDialog')?.close();}if(document.getElementById('packingRows'))render();},getReport:report};
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount);else mount();
})();

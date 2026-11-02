(function (root) {
  'use strict';
  const number = value => Number.isFinite(Number(value)) ? Number(value) : 0;
  const text = value => String(value == null ? '' : value);
  const escape = value => text(value).replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const money = value => 'RD$ ' + number(value).toLocaleString('es-DO',{minimumFractionDigits:2,maximumFractionDigits:2});
  const qty = value => number(value).toLocaleString('es-DO',{maximumFractionDigits:3});
  const unique = values => [...new Set(values.filter(Boolean))];
  const dateLabel = value => new Date(value).toLocaleString('es-DO',{timeZone:'America/Santo_Domingo',dateStyle:'medium',timeStyle:'short'});
  const invoiceDisclaimer = 'El tránsito indica coincidencias de material, no una reserva para el cliente. Una misma cantidad de factura puede coincidir con varios pedidos; no se suma como disponibilidad asignada.';
  const valueDisclaimer = 'Importes en DOP. Valor en ruta = confirmado no entregado + entregado pendiente de facturar. Las cantidades se presentan por material, sin sumar unidades de artículos distintos.';

  function build(context, domain, at) {
    const all = (context.groups || []).flatMap(group => (group.items || []).map(line => ({group,line})));
    const summary = domain.calendarDispatchSummary(all.map(row => row.line),context.transitMaterials || []);
    const invoices = context.invoiceKey !== null && context.invoiceKey !== undefined ? summary.invoices.filter(row => row.key === context.invoiceKey) : summary.invoices;
    const affected = new Set(invoices.flatMap(invoice => invoice.affected.map(row => row.line)));
    const source = context.invoiceKey !== null && context.invoiceKey !== undefined ? all.filter(row => affected.has(row.line)) : all;
    if (!source.length && context.source !== 'day') throw new Error('Esta factura ya no afecta a los pedidos o el despacho ya no tiene materiales. Actualiza el calendario.');
    const scoped = domain.calendarDispatchSummary(source.map(row => row.line),[]);
    const lines = source.map(({line,group}) => ({customer:text(group.nombre),code:text(group.codigo),date:text(group.fecha),truck:text(group.camion),route:text(group.rutaNombre || group.zona),order:text(line.pedidoCliente || 'Sin pedido'),position:text(line.lineaPedidoCliente),material:text(line.articulo || 'Sin referencia'),description:text(line.descripcionArticulo),pending:number(line.cantidadPendiente),confirmed:number(line.montoPlanificado),billing:number(line.montoEntregadoNoFacturado || line.importeEntregadoNoFacturado || line.montoPendienteFacturar),total:number(line.montoPlanificado)+number(line.montoEntregadoNoFacturado || line.importeEntregadoNoFacturado || line.montoPendienteFacturar)}));
    const transit = invoices.flatMap(invoice => invoice.affected.map(row => {
      const group = all.find(item => item.line === row.line).group;
      return {invoice:text(invoice.invoice || 'Sin identificar'),eta:invoice.dates.join(' / ') || 'Por confirmar',customer:text(group.nombre),order:text(row.line.pedidoCliente || 'Sin pedido'),position:text(row.line.lineaPedidoCliente),material:text(row.line.articulo),description:text(row.line.descripcionArticulo),pending:number(row.pending),quantity:number(row.transitQuantity)};
    }));
    const groups = unique(source.map(row => row.group));
    const notes = groups.flatMap(group => [['Nota',group.nota],['Incidencia',group.incidencia],['Precinto',group.precintoDespacho],['Chofer / recibido por',group.choferRuta],['Comentario de cierre',group.comentarioRuta],['Alertas',(group.alertas || []).join(' · ')],['Evidencia',group.fotoCamion ? 'Foto adjunta al despacho: '+(group.fotoCamionNombre || 'sin nombre') : '']].filter(row => row[1]).map(([label,value]) => ({customer:text(group.nombre),label,value:text(value),status:'Guardado'})));
    (context.notes || []).forEach(row => notes.push({...row}));
    (context.drafts || []).forEach(row => notes.push({...row,status:'Borrador sin guardar'}));
    const metrics = {orders:scoped.orders,unidentified:scoped.unidentifiedOrders,confirmed:scoped.confirmed,billing:scoped.pendingBilling,total:scoped.total,customers:unique(lines.map(row => row.code || row.customer)).length,invoices:invoices.length};
    return {title:context.invoiceKey != null ? 'Materiales en tránsito' : context.source === 'day' ? 'Nota del calendario' : context.source === 'move' ? 'Programación de despachos' : context.source === 'note' ? 'Seguimiento del despacho' : 'Despacho del cliente',subtitle:unique(lines.map(row=>row.customer)).join(' · ') || context.day || 'Calendario',dates:unique(lines.map(row=>row.date)).join(' / ') || context.day,routes:unique(lines.map(row=>row.route)).join(' / ') || 'Sin ruta',trucks:unique(lines.map(row=>row.truck)).join(' / ') || 'Sin asignar',at:at || new Date().toISOString(),author:context.author || 'Usuario',metrics,lines,transit,notes,disclaimers:[valueDisclaimer,...(transit.length ? [invoiceDisclaimer] : []),...(context.source === 'move' ? ['La programación propuesta es un borrador. Este informe no mueve ni asigna pedidos.'] : [])]};
  }

  const columns = {
    lines:[['Cliente','customer'],['Código cliente','code'],['Fecha','date'],['Camión','truck'],['Ruta','route'],['Pedido','order'],['Línea','position'],['Material','material'],['Descripción','description'],['Cant. pendiente','pending','qty'],['Confirmado DOP','confirmed','money'],['Pend. facturar DOP','billing','money'],['Valor en ruta DOP','total','money']],
    transit:[['Factura','invoice'],['ETA','eta'],['Cliente','customer'],['Pedido','order'],['Línea','position'],['Material','material'],['Descripción','description'],['Sin confirmar','pending','qty'],['Coincidencia en factura','quantity','qty']],
    notes:[['Cliente','customer'],['Concepto','label'],['Detalle','value'],['Estado','status']]
  };
  const loads = new Map();
  function loadScript(file, ready) {
    if (ready()) return Promise.resolve();
    if (loads.has(file)) return loads.get(file);
    const promise = new Promise((resolve,reject) => {
      const script = document.createElement('script');script.src='./vendor/'+file;
      const timer=setTimeout(()=>failed(),20000);
      function failed(){clearTimeout(timer);script.remove();loads.delete(file);reject(new Error('No se pudo cargar el generador. Revisa la conexión y vuelve a intentarlo.'));}
      script.onload=()=>{clearTimeout(timer);if(ready())resolve();else failed();};script.onerror=failed;document.head.appendChild(script);
    });
    loads.set(file,promise);return promise;
  }
  function download(blob,name) {
    const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=name;document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
  }
  function filename(report,extension){return 'TMS_Ejecutivo_'+report.subtitle.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-zA-Z0-9]+/g,'_').slice(0,65)+'_'+report.at.slice(0,10)+'.'+extension;}

  async function excel(report) {
    await loadScript('exceljs-4.4.0.min.js',()=>!!root.ExcelJS);
    const workbook=new root.ExcelJS.Workbook();workbook.creator='TMS · Álvarez Dominicana';workbook.created=new Date(report.at);
    const ink='17384C',teal='087F8C',gold='BC892A',light='F0F5F8';
    function sheet(name,headers,rows){
      const ws=workbook.addWorksheet(name,{views:[{state:'frozen',ySplit:5}],pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0}});
      ws.properties.defaultRowHeight=24;ws.columns=headers.map(([label,key,type])=>({key,width:type?22:key==='description'||key==='value'?58:key==='customer'?38:23}));
      const end=Math.max(headers.length,2);ws.mergeCells(1,1,1,end);const title=ws.getCell(1,1);title.value='ÁLVAREZ DOMINICANA  |  '+report.title;title.font={name:'Calibri',size:18,bold:true,color:{argb:'FFFFFF'}};title.fill={type:'pattern',pattern:'solid',fgColor:{argb:ink}};ws.getRow(1).height=36;
      ws.mergeCells(2,1,2,end);ws.getCell(2,1).value=report.subtitle;ws.getCell(2,1).font={name:'Calibri',size:12,bold:true,color:{argb:ink}};ws.getCell(2,1).alignment={wrapText:true,vertical:'middle'};ws.getRow(2).height=Math.max(24,Math.min(200,18*Math.ceil(report.subtitle.length/(end*17))));
      ws.mergeCells(3,1,3,end);ws.getCell(3,1).value='Generado: '+dateLabel(report.at)+' · '+report.author;ws.getCell(3,1).font={name:'Calibri',size:11,color:{argb:'526678'}};
      ws.getRow(5).values=headers.map(row=>row[0]);ws.getRow(5).height=32;
      ws.getRow(5).eachCell(cell=>{cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:teal}};cell.font={name:'Calibri',bold:true,color:{argb:'FFFFFF'},size:11};cell.alignment={vertical:'middle',wrapText:true};});
      rows.forEach((row,index)=>{const added=ws.addRow(headers.map(([,key])=>row[key] == null ? '' : row[key]));added.height=Math.max(28,Math.min(140,20*Math.ceil(Math.max(...headers.map(([,key])=>text(row[key]).length/(key==='value'||key==='description'?55:30))))));added.eachCell((cell,col)=>{cell.font={name:'Calibri',size:11,color:{argb:ink}};cell.alignment={vertical:'top',wrapText:true};if(index%2===0)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:light}};if(headers[col-1][2])cell.numFmt=headers[col-1][2]==='money'?'"RD$ "#,##0.00;[Red]("RD$ "#,##0.00)':'#,##0.###';});});
      if(rows.length)ws.autoFilter={from:{row:5,column:1},to:{row:rows.length+5,column:headers.length}};
      ws.headerFooter.oddFooter='TMS | Informe ejecutivo &R Página &P de &N';return ws;
    }
    const m=report.metrics;
    const summary=sheet('Resumen',[['Indicador','label'],['Resultado','value']], [
      {label:'Fecha del despacho',value:report.dates},{label:'Ruta',value:report.routes},{label:'Camión',value:report.trucks},
      {label:'Clientes',value:m.customers},{label:'Pedidos identificados',value:m.orders},{label:'Líneas sin pedido identificado',value:m.unidentified},
      {label:'Confirmado no entregado',value:m.confirmed},{label:'Entregado pendiente de facturar',value:m.billing},{label:'Valor en ruta',value:m.total},{label:'Facturas en tránsito que afectan',value:m.invoices},
      ...report.disclaimers.map(value=>({label:'Criterio del informe',value}))
    ]);summary.getColumn(1).width=43;summary.getColumn(2).width=115;
    [12,13,14].forEach(row=>{summary.getCell(row,2).numFmt='"RD$ "#,##0.00;[Red]("RD$ "#,##0.00)';});
    summary.getRow(14).eachCell(cell=>{cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:gold}};cell.font={bold:true,color:{argb:'FFFFFF'},size:12};});
    const detail=sheet('Pedidos',columns.lines,report.lines);
    const total=detail.addRow({description:'TOTAL DOP',confirmed:m.confirmed,billing:m.billing,total:m.total});total.font={bold:true,color:{argb:ink}};['confirmed','billing','total'].forEach(key=>{total.getCell(key).numFmt='"RD$ "#,##0.00';});
    sheet('Tránsito',columns.transit,report.transit);sheet('Seguimiento',columns.notes,report.notes);
    return new Blob([await workbook.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }

  async function pdf(report) {
    await loadScript('jspdf-4.2.1.min.js',()=>!!root.jspdf?.jsPDF);
    await loadScript('jspdf-autotable-5.0.9.min.js',()=>!!root.jspdf?.jsPDF?.API.autoTable);
    const doc=new root.jspdf.jsPDF({orientation:'landscape',unit:'mm',format:'a4'}),width=297,height=210,ink=[23,56,76],teal=[8,127,140];
    doc.setProperties({title:'Informe ejecutivo | '+report.title,author:report.author,subject:report.subtitle,creator:'TMS Álvarez Dominicana'});
    const clean=value=>text(value).replace(/[^\x20-\x7E\xA0-\xFF\n]/g,' ');
    function header(){doc.setFillColor(...ink);doc.rect(0,0,width,19,'F');doc.setTextColor(255);doc.setFont('helvetica','bold');doc.setFontSize(11);doc.text('ÁLVAREZ DOMINICANA',14,12);doc.setFontSize(9);doc.text('TMS / INFORME EJECUTIVO',width-14,12,{align:'right'});}
    function block(value,y,font=10,bold=false,color=ink,maxLines=Infinity){doc.setFont('helvetica',bold?'bold':'normal');doc.setFontSize(font);doc.setTextColor(...color);const all=doc.splitTextToSize(clean(value),269),lines=all.slice(0,maxLines);if(all.length>maxLines)lines[lines.length-1]=lines.at(-1).slice(0,-3)+'...';doc.text(lines,14,y);return y+lines.length*font*.43+3;}
    header();let y=block(report.title,29,20,true);y=block(report.subtitle,y,11,true,ink,3);y=block(report.dates+'  |  '+report.routes+'  |  '+report.trucks,y,9,false,ink,2);
    const m=report.metrics,metrics=[['PEDIDOS',qty(m.orders)+(m.unidentified?' + sin identificar':'')],['CONFIRMADO',money(m.confirmed)],['PEND. FACTURAR',money(m.billing)],['VALOR EN RUTA',money(m.total)]];
    if(y>95){doc.addPage();header();y=30;}
    metrics.forEach(([label,value],index)=>{const x=14+index*68;doc.setFillColor(240,245,248);doc.roundedRect(x,y,65,21,2,2,'F');doc.setTextColor(...teal);doc.setFontSize(8);doc.setFont('helvetica','bold');doc.text(label,x+4,y+6);doc.setTextColor(...ink);doc.setFontSize(12);doc.text(clean(value),x+4,y+16);});y+=26;
    const base=Math.abs(m.confirmed)+Math.abs(m.billing),part=base?Math.abs(m.confirmed)/base:0;
    doc.setFillColor(...ink);doc.rect(14,y,269*part,4,'F');doc.setFillColor(188,137,42);doc.rect(14+269*part,y,269*(1-part),4,'F');y+=8;
    y=block('Composición del valor: confirmado (azul) y pendiente de facturar (dorado). '+m.customers+' cliente(s) · '+m.invoices+' factura(s) en tránsito.',y,9);
    function table(title,heads,rows){
      if(y>height-40){doc.addPage();header();y=29;}
      y=block(title,y+3,12,true);
      doc.autoTable({startY:y,margin:{left:14,right:14,top:28,bottom:16},head:[heads],body:rows.length?rows:[heads.map((_,index)=>index===0?'Sin registros':'')],styles:{font:'helvetica',fontSize:8,cellPadding:2,overflow:'linebreak',textColor:ink,lineColor:[225,232,237],lineWidth:.1},headStyles:{fillColor:teal,textColor:255,fontStyle:'bold'},alternateRowStyles:{fillColor:[245,248,250]},rowPageBreak:'avoid',didDrawPage:header});y=doc.lastAutoTable.finalY+5;
    }
    const varyingSchedule=unique(report.lines.map(row=>[row.date,row.truck,row.route].join('|'))).length>1;
    table('Pedidos y materiales',['Cliente / pedido','Material','Pendiente','Confirmado DOP','Pend. facturar DOP','Valor en ruta DOP'],report.lines.map(row=>[clean(row.customer+'\n'+row.order+(row.position?' / '+row.position:'')+(varyingSchedule?'\n'+row.date+' · '+row.truck+'\n'+row.route:'')),clean(row.material+'\n'+row.description),qty(row.pending),money(row.confirmed),money(row.billing),money(row.total)]));
    if(report.transit.length)table('Materiales afectados por tránsito',['Factura / ETA','Cliente / pedido','Material','Sin confirmar','Coincidencia en factura'],report.transit.map(row=>[clean(row.invoice+'\nETA: '+row.eta),clean(row.customer+'\n'+row.order+(row.position?' / '+row.position:'')),clean(row.material+'\n'+row.description),qty(row.pending),qty(row.quantity)]));
    if(report.notes.length)table('Seguimiento y evidencias',['Cliente','Concepto','Detalle','Estado'],report.notes.map(row=>[clean(row.customer),clean(row.label),clean(row.value),clean(row.status)]));
    report.disclaimers.forEach(value=>{doc.setFont('helvetica','normal');doc.setFontSize(8);const lines=doc.splitTextToSize(clean(value),269);if(y+lines.length*3.5+3>height-16){doc.addPage();header();y=30;}y=block(value,y,8,false,[82,102,120]);});
    const count=doc.getNumberOfPages();for(let page=1;page<=count;page++){doc.setPage(page);doc.setFont('helvetica','normal');doc.setFontSize(8);doc.setTextColor(82,102,120);doc.text(clean(dateLabel(report.at)+' · '+report.author),14,height-8);doc.text('Página '+page+' de '+count,width-14,height-8,{align:'right'});}
    return doc.output('blob');
  }

  let current=null;
  function preview(report){
    const modal=document.getElementById('calendarReportDialog') || document.body.appendChild(Object.assign(document.createElement('dialog'),{id:'calendarReportDialog',className:'calendar-report-dialog'}));
    modal.setAttribute('aria-labelledby','calendarReportTitle');current=report;
    const m=report.metrics;
    modal.innerHTML=`<div class="exec-actions"><strong>Informe ejecutivo</strong><div><button class="btn btn-outline" data-exec-format="xlsx">Descargar Excel</button><button class="btn btn-primary" data-exec-format="pdf">Descargar PDF</button><button class="btn btn-outline" data-exec-close aria-label="Cerrar informe">Cerrar</button></div></div><p role="status" class="exec-status"></p><article class="exec-report"><header class="exec-brand"><span>ÁLVAREZ DOMINICANA</span><span>TMS / CALENDARIO</span></header><div class="exec-content"><p class="exec-eyebrow">INFORME EJECUTIVO</p><h2 id="calendarReportTitle">${escape(report.title)}</h2><h3>${escape(report.subtitle)}</h3><p>${escape(report.dates)} · ${escape(report.routes)} · ${escape(report.trucks)}</p><dl class="exec-metrics">${[['Pedidos',m.orders+(m.unidentified?' + sin identificar':'')],['Confirmado',money(m.confirmed)],['Pend. facturar',money(m.billing)],['Valor en ruta',money(m.total)]].map(([label,value])=>`<div><dt>${label}</dt><dd>${escape(value)}</dd></div>`).join('')}</dl><h4>Composición del valor en ruta</h4><div class="exec-bar" role="img" aria-label="Confirmado ${escape(money(m.confirmed))}; pendiente de facturar ${escape(money(m.billing))}"><span style="width:${Math.abs(m.confirmed)+Math.abs(m.billing)?Math.abs(m.confirmed)/(Math.abs(m.confirmed)+Math.abs(m.billing))*100:0}%"></span></div><p class="exec-legend">● Confirmado <span>● Pendiente de facturar</span> · ${m.customers} cliente(s) · ${m.invoices} factura(s) en tránsito</p>${previewTable('Pedidos y materiales',columns.lines.filter(([,key])=>!['date','truck','route','code'].includes(key)),report.lines)}${report.transit.length?previewTable('Materiales afectados por tránsito',columns.transit,report.transit):''}${report.notes.length?previewTable('Seguimiento y evidencias',columns.notes,report.notes):''}<aside class="exec-criteria">${report.disclaimers.map(value=>`<p>${escape(value)}</p>`).join('')}</aside><footer>Generado ${escape(dateLabel(report.at))} · ${escape(report.author)}</footer></div></article>`;
    modal.querySelector('[data-exec-close]').onclick=()=>modal.close();
    modal.querySelectorAll('[data-exec-format]').forEach(button=>button.onclick=async()=>{
      const status=modal.querySelector('.exec-status'),buttons=modal.querySelectorAll('[data-exec-format]');buttons.forEach(b=>b.disabled=true);status.textContent='Preparando descarga…';
      const snapshot=current;
      try{if(!root.TMSAuditCore.hasPermission('calendario','ver'))throw new Error('Tu perfil ya no puede consultar el calendario.');const format=button.dataset.execFormat;const blob=await (format==='pdf'?pdf(snapshot):excel(snapshot));if(!root.TMSAuditCore.hasPermission('calendario','ver'))throw new Error('Tu perfil ya no puede consultar el calendario.');download(blob,filename(snapshot,format));status.textContent='Informe preparado. Revisa las descargas de tu dispositivo.';}catch(error){status.textContent=error.message || 'No se pudo generar el informe. Intenta nuevamente.';}finally{buttons.forEach(b=>b.disabled=false);}
    });
    modal.onclose=()=>{if(!modal.open){current=null;modal.replaceChildren();}};if(!modal.open)modal.showModal();
  }
  function previewTable(title,headers,rows){return `<section><h4>${title}</h4><div class="exec-table-wrap"><table><thead><tr>${headers.map(([label])=>`<th>${label}</th>`).join('')}</tr></thead><tbody>${rows.map(row=>`<tr>${headers.map(([,key,type])=>`<td>${escape(type==='money'?money(row[key]):type==='qty'?qty(row[key]):row[key])}</td>`).join('')}</tr>`).join('')}</tbody></table></div></section>`;}
  function open(source){
    try{
      const context=root.getCalendarExecutiveContext(source);
      if(source==='note'){
        context.drafts=[['Nota','notaInput'],['Incidencia','incidenciaInput'],['Precinto','precintoInput'],['Chofer / recibido por','choferRutaInput'],['Comentario de cierre','comentarioRutaInput'],['Salida del camión','auditTripInput']].flatMap(([label,id])=>{const value=document.getElementById(id)?.value || '';const saved={Nota:'nota',Incidencia:'incidencia',Precinto:'precintoDespacho','Chofer / recibido por':'choferRuta','Comentario de cierre':'comentarioRuta', 'Salida del camión':'tripId'}[label];return value!==text(context.groups[0][saved] || (saved==='tripId'?context.groups[0].items?.[0]?.tripId || '1':''))?[{customer:context.groups[0].nombre,label,value:value || '(Vacío)'}]:[];});
        if(APP.pendingTruckPhoto)context.drafts.push({customer:context.groups[0].nombre,label:'Evidencia',value:'Foto pendiente de guardar: '+(APP.pendingTruckPhoto.name || 'sin nombre')});
      }
      if(source==='move')context.drafts=[{customer:context.groups.map(row=>row.nombre).join(' · '),label:'Programación propuesta',value:(document.getElementById('moveFecha')?.value || 'Sin fecha')+' · '+(document.getElementById('moveCamion')?.value || 'Sin camión')}];
      preview(build(context,root.TMSDomain));
    }catch(error){alert(error.message);}
  }
  function mount(){
    [['calendarDispatchDialog','dispatch'],['moveModal','move'],['notaModal','note']].forEach(([id,source])=>{
      const modal=document.getElementById(id);if(!modal || modal.querySelector('[data-calendar-report]') || (id==='calendarDispatchDialog' && !modal.querySelector('.calendar-detail-eyebrow')))return;
      const button=document.createElement('button');button.type='button';button.className='btn btn-outline calendar-report-open';button.dataset.calendarReport=source;button.textContent='Informe ejecutivo';button.onclick=()=>open(source);
      const host=modal.querySelector('.modal-box') || modal;host.insertBefore(button,host.querySelector('.modal-actions') || host.children[1] || null);
    });
    if(document.getElementById('calendarReportDialog')?.open && !root.TMSAuditCore?.hasPermission('calendario','ver'))document.getElementById('calendarReportDialog').close();
  }
  root.TMSCalendarReports={build,excel,pdf,open,preview};
  if(typeof module!=='undefined' && module.exports)module.exports={build};
  if(typeof document!=='undefined'){
    let scheduled=false;new MutationObserver(()=>{if(!scheduled){scheduled=true;queueMicrotask(()=>{scheduled=false;mount();});}}).observe(document.body,{childList:true,subtree:true});mount();
  }
})(typeof window!=='undefined'?window:globalThis);

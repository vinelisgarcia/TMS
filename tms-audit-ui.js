(function () {
  'use strict';
  const D = window.TMSDomain;
  const core = () => window.TMSAuditCore;
  const escape = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  const money = (value, currency = 'USD') => value === null || value === undefined ? 'Pendiente' : new Intl.NumberFormat('es-DO', { style: 'currency', currency, maximumFractionDigits: 2 }).format(value);
  const localDate = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; };
  const settings = () => (APP.costSettings = APP.costSettings || {});
  function can(module, action = 'editar') { return core() && core().hasPermission(module, action); }
  function save(label) {
    window.scheduleAutoSave();
    toast(label);
  }
  function toast(message) {
    let el = document.getElementById('auditToast');
    if (!el) { el = document.createElement('div'); el.id = 'auditToast'; el.setAttribute('role', 'status'); document.body.appendChild(el); }
    el.textContent = message; el.classList.add('visible');
    clearTimeout(window._auditToastTimer); window._auditToastTimer = setTimeout(() => el.classList.remove('visible'), 4000);
  }
  function refresh() {
    window.renderMaterialesTransito(); window.renderCalendario(); window.renderComercial(); window.actualizarDashboard();
  }
  function transitView() {
    const mount = document.getElementById('materialesTransitoMount');
    if (!mount) return;
    const query = String(APP.auditTransitSearch || '').toLowerCase();
    const received = APP.auditTransitTab === 'received';
    const all = APP.transitMaterials || [];
    const filtered = all.filter(row => D.activeTransit(row) !== received && (!query || [row.factura, row.codigoMaterial, row.descripcion, row.cliente].some(value => String(value || '').toLowerCase().includes(query))));
    const invoices = new Map();
    filtered.forEach(row => {
      const id = row.factura ? D.invoiceKey(row) : `unidentified-${row.id}`;
      if (!invoices.has(id)) invoices.set(id, { invoice: row.factura || '', rows: [] });
      invoices.get(id).rows.push(row);
    });
    mount.innerHTML = `<div class="audit-module-head"><div><p class="audit-eyebrow">Disponibilidad futura</p><h2>Materiales en tránsito</h2><p>Facturas, referencias y llegada estimada en un solo lugar.</p></div><button class="btn btn-outline" onclick="exportarMaterialesTransitoExcel()">Exportar Excel</button></div>
      ${can('materialesTransito') || can('materialesTransito','importar') ? `<section class="card"><div class="card-title">Cargar materiales de una factura</div><div class="audit-form-grid">
      <label>Factura<input id="transitFactura" placeholder="Número de factura" class="form-control"></label>
      <label>Llegada estimada<input id="transitEta" type="date" class="form-control"></label>
      <label>Cliente asignado, si aplica<input id="transitCliente" placeholder="Código o nombre exacto" class="form-control"></label>
      <label>Archivo<input id="transitFile" type="file" accept=".xlsx,.xls,.xml,.csv,.txt,.pdf" class="form-control"></label></div>
      <div class="audit-actions"><button class="btn btn-primary" onclick="importarMaterialesTransito()">Importar materiales</button><span id="transitStatus" role="status"></span></div></section>` : ''}
      <div class="audit-stat-grid"><article><span>Facturas activas</span><strong>${new Set(all.filter(D.activeTransit).map(D.invoiceKey).filter(Boolean)).size}</strong></article><article><span>Referencias en tránsito</span><strong>${new Set(all.filter(D.activeTransit).map(row=>D.key(row.codigoMaterial))).size}</strong></article><article><span>Facturas recibidas</span><strong>${new Set(all.filter(row=>!D.activeTransit(row)).map(D.invoiceKey).filter(Boolean)).size}</strong></article></div>
      <div class="audit-toolbar"><div class="audit-tabs"><button data-audit="transit-tab" data-value="active" aria-pressed="${!received}">En tránsito</button><button data-audit="transit-tab" data-value="received" aria-pressed="${received}">Recibidas</button></div><label class="audit-search">Buscar<input id="auditTransitSearch" type="search" value="${escape(query)}" placeholder="Factura, referencia o cliente"></label></div>
      <div class="audit-invoices">${[...invoices.entries()].map(([id, group]) => {
        const dates = [...new Set(group.rows.map(row=>row.eta).filter(Boolean))].sort();
        const receipt = group.rows.find(row=>row.receivedAt);
        return `<article class="card audit-invoice"><header><div><span class="audit-eyebrow">${received ? 'Historial' : 'Factura de importación'}</span><h3>${escape(group.invoice || 'Factura sin identificar')}</h3><p>${group.rows.length} líneas · ${received ? `Recibida ${escape((receipt && receipt.receivedAt || '').slice(0,10))}` : `ETA ${escape(dates.join(' / ') || 'por confirmar')}`}</p></div>${can('materialesTransito') && group.invoice ? `<button class="btn ${received ? 'btn-outline' : 'btn-primary'}" data-audit="receive-invoice" data-invoice="${escape(group.invoice)}" data-value="${received ? 'reopen' : 'receive'}">${received ? 'Volver a tránsito' : 'Marcar factura recibida'}</button>` : ''}</header>
          <details><summary>Ver materiales</summary><div class="table-wrap"><table class="data-table"><thead><tr><th>Referencia</th><th>Descripción</th><th>Cantidad</th><th>Cliente</th><th>ETA</th></tr></thead><tbody>${group.rows.map(row=>`<tr><td>${escape(row.codigoMaterial)}</td><td>${escape(row.descripcion)}</td><td>${Number(row.cantidad).toLocaleString('es-DO')}</td><td>${escape(row.cliente || 'Sin asignación exclusiva')}</td><td>${escape(row.eta || 'Pendiente')}</td></tr>`).join('')}</tbody></table></div></details>${received && receipt ? `<small>Confirmada por ${escape(receipt.receivedBy || 'usuario autorizado')}. La recepción no confirma una reserva de SAP.</small>` : ''}</article>`;
      }).join('') || `<div class="audit-empty">${received ? 'Todavía no hay facturas recibidas.' : 'No hay facturas activas para esta búsqueda.'}</div>`}</div>`;
  }
  window.renderMaterialesTransito = transitView;

  function billForTrip(trip) {
    const rows = (settings().actualInvoices || []).filter(row=>row.tripKey===trip.id && (trip.orderRefs || []).includes(D.key(row.customer)+'|'+D.key(row.order)));
    if (!rows.length) return null;
    if (rows.some(row=>!['DOP','USD'].includes(row.currency) || (row.currency==='USD' && !trip.fx))) return null;
    return rows.reduce((sum,row)=>sum+row.amount*(row.currency==='USD'?trip.fx:1),0);
  }
  function costRows() {
    const from = (document.getElementById('auditCostFrom') || {}).value || '';
    const to = (document.getElementById('auditCostTo') || {}).value || '';
    return D.calculateTrips(APP).filter(row=>(!from || row.date>=from)&&(!to || row.date<=to));
  }
  function preparationRows() {
    const from=(document.getElementById('auditCostFrom') || {}).value || '';
    const to=(document.getElementById('auditCostTo') || {}).value || '';
    return D.warehouseCosts(APP).filter(row=>(!from || row.date>=from)&&(!to || (row.date && row.date<=to)));
  }
  function renderCostDashboard(targetId = 'auditCostsBody') {
    const mount = document.getElementById(targetId); if (!mount) return;
    const trips = costRows();
    const complete = trips.filter(row=>row.usd!==null);
    const totalUsd = complete.reduce((sum,row)=>sum+row.usd,0);
    const executed = trips.filter(row=>row.executed);
    const prep = preparationRows();
    const prepared = prep.filter(row=>row.prepared);
    const preparationTotal = settings().preparationRateUsd===null || settings().preparationRateUsd===undefined || settings().preparationRateUsd==='' ? null : prepared.reduce((sum,row)=>sum+(row.usd || 0),0);
    const actualMissing = !settings().actualInvoices || !settings().actualInvoices.length;
    const allTrips=D.calculateTrips(APP);
    const unmatchedBills=(settings().actualInvoices || []).filter(bill=>!allTrips.some(trip=>trip.id===bill.tripKey && (trip.orderRefs || []).includes(D.key(bill.customer)+'|'+D.key(bill.order)))).length;
    mount.innerHTML = `<div class="audit-stat-grid"><article><span>Salidas programadas</span><strong>${trips.length}</strong></article><article><span>Costo previsto USD${complete.length!==trips.length ? ' · parcial' : ''}</span><strong>${money(totalUsd)}</strong></article><article><span>Salidas confirmadas</span><strong>${executed.length}</strong></article><article><span>Solicitudes preparadas</span><strong>${prepared.length}</strong></article><article><span>Preparación USD</span><strong>${money(preparationTotal)}</strong></article></div>
      <div class="audit-notice">${actualMissing ? 'Facturación real pendiente: adjunta el reporte SAP. No se utiliza una estimación como importe facturado.' : 'La comparación utiliza importes SAP vinculados a cliente, pedido, fecha de despacho, camión y salida.'} ${unmatchedBills ? unmatchedBills+' líneas facturadas sin salida/pedido coincidente; no se incluyen en la comparación.' : ''} ${trips.some(row=>!row.fx) ? 'Falta tipo de cambio para una o más fechas.' : ''}</div>
      <div class="audit-chart-wrap"><canvas id="${targetId}Chart" aria-label="Costos diarios de transporte y facturación real"></canvas></div>
      <div class="table-wrap"><table class="data-table audit-cost-table"><thead><tr><th>Fecha / salida</th><th>Ruta / clientes</th><th>Vehículo</th><th>Transporte USD</th><th>Transporte RD$</th><th>Facturado RD$</th><th>Transporte / facturado</th><th>Estado</th></tr></thead><tbody>${trips.map(row=> {
        const billed = billForTrip(row);
        return `<tr><td>${escape(row.date)}<br>${escape(row.truck)} · ${escape(row.trip)}</td><td>${escape(row.routes.join(' + '))}<br><small>${row.clients.length} clientes · ${row.requests} solicitudes</small><details><summary>Desglose</summary>${row.clients.map((client,index)=>`<p>${escape(client.name || client.code)} · ${index ? 'Adicional' : 'Base'}: ${money(client.charge)}</p>`).join('')}</details></td><td>${escape(({large:'Grande · 12 PAL',small:'Pequeño · 4 PAL',container:'Contenedor'})[row.vehicle] || 'Por configurar')}</td><td>${money(row.usd)}</td><td>${money(row.dop,'DOP')}</td><td>${money(billed,'DOP')}</td><td>${row.dop!==null&&billed>0 ? (row.dop/billed*100).toFixed(1)+'%' : billed===0 ? 'Sin facturación' : 'Pendiente'}</td><td>${row.lockedAt ? '<span class="badge badge-ok">Salida confirmada</span>' : `<span class="badge badge-outline">Previsto</span>${(can('calendario') || can('configuracion'))&&row.usd!==null ? `<button class="btn btn-outline btn-sm" data-audit="lock-trip" data-key="${escape(row.id)}">Confirmar salida y tarifa</button>` : ''}`}<small>${escape(row.issues.join(' · '))}</small></td></tr>`;
      }).join('') || '<tr><td colspan="8">Aún no hay salidas para este período. Programa pedidos desde la cola.</td></tr>'}</tbody></table></div>
      <div class="audit-caption">Tarifario CT agosto 2026 · USD. La base corresponde al destino más costoso; los demás clientes generan su adicional. Preparación: ${settings().preparationRateUsd===undefined || settings().preparationRateUsd===null || settings().preparationRateUsd==='' ? 'tarifa pendiente de validar' : money(Number(settings().preparationRateUsd))+' por solicitud única'}. ${prep.filter(row=>row.partial).length} solicitudes parcialmente preparadas.</div>`;
    const canvas = document.getElementById(targetId+'Chart');
    if (typeof Chart === 'undefined' || !canvas) return;
    window._auditCharts = window._auditCharts || {};
    if (window._auditCharts[targetId]) window._auditCharts[targetId].destroy();
    const dates = [...new Set(trips.map(row=>row.date))];
    const daily = date => trips.filter(row=>row.date===date);
    const datasets = [
      {label:'Costo transporte previsto USD',data:dates.map(date=>{const rows=daily(date);return rows.some(row=>row.usd===null)?null:rows.reduce((sum,row)=>sum+row.usd,0);}),backgroundColor:'#596cda',borderRadius:6,yAxisID:'usd'},
      {label:'Costo transporte ejecutado USD',data:dates.map(date=>{const rows=daily(date).filter(row=>row.executed);return rows.some(row=>row.usd===null)?null:rows.reduce((sum,row)=>sum+row.usd,0);}),backgroundColor:'#29a990',borderRadius:6,yAxisID:'usd'}
    ];
    if (!actualMissing) datasets.push({type:'line',label:'Facturado real RD$',data:dates.map(date=>{const values=daily(date).map(billForTrip);return values.some(value=>value===null)?null:values.reduce((a,b)=>a+b,0);}),borderColor:'#c57925',yAxisID:'dop',spanGaps:false});
    window._auditCharts[targetId] = new Chart(canvas,{type:'bar',data:{labels:dates,datasets},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{usd:{position:'left',title:{display:true,text:'USD · transporte'},beginAtZero:true},dop:{position:'right',display:!actualMissing,title:{display:true,text:'RD$ · facturación'},grid:{drawOnChartArea:false}}}}});
  }
  function ensureDashboard() {
    const dashboard = document.getElementById('dashboard'); if (!dashboard) return;
    if (!document.getElementById('auditCostPanel')) {
      const panel = document.createElement('section');panel.className='card audit-cost-panel';panel.id='auditCostPanel';
      panel.innerHTML=`<div class="audit-module-head"><div><p class="audit-eyebrow">Control de costos</p><h2>Transporte y facturación</h2><p>De la programación a la salida real, con tarifas verificables.</p></div><button class="btn btn-outline" data-audit="export-costs">Exportar Excel</button></div><div class="audit-toolbar"><label>Desde<input type="date" id="auditCostFrom" class="form-control"></label><label>Hasta<input type="date" id="auditCostTo" class="form-control"></label></div><div id="auditCostsBody"></div>`;
      const old = document.getElementById('dashReportPanel');dashboard.insertBefore(panel,old || dashboard.firstChild);
    }
    if(!document.getElementById('auditMoreDashboard')){
      const more=document.createElement('details');more.id='auditMoreDashboard';more.className='card';more.innerHTML='<summary>Más indicadores y reportes</summary>';dashboard.appendChild(more);
      const secondary=dashboard.querySelector('.kpi-row');if(secondary)more.appendChild(secondary);
      const donut=document.getElementById('chartCamiones');if(donut)more.appendChild(donut.closest('.chart-card'));
      const report=document.getElementById('dashReportPanel');if(report)more.appendChild(report);
    }
    renderCostDashboard();
  }
  const oldDashboard = window.actualizarDashboard;
  window.actualizarDashboard = function () { oldDashboard(); ensureDashboard(); };

  function ensureConfig() {
    const config = document.getElementById('configClientes'); if (!config) return;
    const legacyCost=document.getElementById('warehouseCostInput');if(legacyCost)legacyCost.closest('.card').hidden=true;
    document.querySelectorAll('#configExtrasMount input[id^="routeCost_"]').forEach(input=>{
      input.closest('td').hidden=true;const table=input.closest('table');if(table&&table.tHead)table.tHead.rows[0].cells[2].hidden=true;
    });
    const catalog=document.querySelector('#configExtrasMount .card-title');
    if(catalog){catalog.textContent='Catálogo de rutas';const note=catalog.closest('.card').querySelector('.config-help-note');if(note)note.textContent='Configura las rutas y sus días de operación. Asigna cada cliente desde su ficha; las tarifas se definen en la configuración de costos.';}
    if (!document.getElementById('auditCreateClient')) {
      const card=document.createElement('section');card.id='auditCreateClient';card.className='card';
      card.innerHTML=`<div class="card-title">Crear cliente desde cero</div><div class="audit-form-grid"><label>Código SAP<input id="auditClientCode" class="form-control" placeholder="Código único"></label><label>Nombre<input id="auditClientName" class="form-control" placeholder="Nombre del cliente"></label></div><div class="audit-actions"><button class="btn btn-primary" data-audit="create-client">Crear y configurar ruta</button></div>`;
      const grid=config.querySelector('.config-grid');config.insertBefore(card,grid);
    }
    if (!document.getElementById('auditTariffConfig')) {
      const card=document.createElement('section');card.id='auditTariffConfig';card.className='card';config.appendChild(card);
    }
    const s=settings();
    const zones=Object.entries(D.rates).map(([id,row])=>`<option value="${id}">${escape(row.label)}</option>`).join('');
    const routes=(APP.routeConfigs || []).map(route=>route.nombre || route.zona).filter(Boolean);
    const clients=Object.values(APP.clienteConfig || {});
    document.getElementById('auditTariffConfig').innerHTML=`<div class="audit-module-head"><div><p class="audit-eyebrow">Tarifario CT · USD · Agosto 2026</p><h2>Costos de transporte y preparación</h2><p>Configura vehículos y zonas. Los importes sin confirmar permanecen pendientes.</p></div></div>
      <div class="audit-form-grid">${['CAMION 1','CAMION 2','CAMION 3'].map(truck=>`<label>${escape(truck)}<select class="form-control" data-cost-vehicle="${truck}"><option value="">Seleccionar tipo</option><option value="large">Grande · 12 paletas</option><option value="small">Pequeño · 4 paletas</option><option value="container">Contenedores</option></select></label>`).join('')}<label>Preparación por solicitud · USD<input type="number" min="0" step="0.01" id="auditPreparationRate" class="form-control" value="${escape(s.preparationRateUsd ?? '')}" placeholder="Pendiente de validar con CT"></label></div>
      <details class="audit-config-section" open><summary>Zona tarifaria de cada ruta</summary><div class="audit-form-grid">${routes.map((route,index)=>`<label>${escape(route)}<select class="form-control" data-cost-route="${index}"><option value="">Seleccionar zona tarifaria</option>${zones}</select></label>`).join('') || '<p>Crea primero las rutas en el catálogo de configuración.</p>'}</div></details>
      <details class="audit-config-section"><summary>Excepciones de paradas por cliente</summary><div class="audit-form-grid">${clients.map((client,index)=>`<label>${escape(client.nombreCliente || client.codigoCliente)}<select class="form-control" data-cost-stop="${index}"><option value="standard">Parada adicional estándar</option><option value="bahia">Bahía Príncipe La Romana</option></select></label>`).join('') || '<p>Configura los clientes para identificar excepciones.</p>'}</div></details>
      <div class="audit-actions"><button class="btn btn-primary" data-audit="save-tariffs">Guardar configuración tarifaria</button></div>
      <div class="audit-divider"></div><div class="card-title">Tipo de cambio por fecha de servicio</div><div class="audit-form-grid"><label>Fecha<input type="date" id="auditFxDate" class="form-control" value="${localDate()}"></label><label>RD$ por USD<input type="number" min="0.01" step="0.0001" id="auditFxRate" class="form-control" placeholder="Tasa aplicable al día"></label></div><div class="audit-actions"><button class="btn btn-outline" data-audit="save-fx">Guardar tipo de cambio</button></div>
      <details class="audit-config-section"><summary>Tarifas de referencia</summary><div class="table-wrap"><table class="data-table"><thead><tr><th>Destino</th><th>Grande USD</th><th>Pequeño USD</th><th>Contenedor USD</th></tr></thead><tbody>${Object.values(D.rates).map(row=>`<tr><td>${escape(row.label)}</td><td>${row.large}</td><td>${row.small}</td><td>${row.container}</td></tr>`).join('')}<tr><td>Adicional estándar</td><td>125</td><td>75</td><td>No aplica</td></tr><tr><td>Adicional Bahía Príncipe</td><td>150</td><td>175</td><td>No aplica</td></tr></tbody></table></div></details>
      <div class="audit-divider"></div><div class="card-title">Facturación efectiva de SAP</div><p class="audit-caption">Carga Excel o CSV con: Fecha despacho (AAAA-MM-DD), Camión, Salida, Cliente, Pedido, Factura, Moneda (DOP o USD) e Importe. Debe contener importes de línea, no totales de factura repetidos. Los registros sin identificación completa se rechazan.</p><label class="audit-file-label">Reporte de facturación<input type="file" id="auditBillingFile" accept=".xlsx,.xls,.csv"></label><div id="auditBillingStatus" role="status">${(s.actualInvoices || []).length} líneas cargadas.</div>`;
    document.querySelectorAll('[data-cost-vehicle]').forEach(el=>{el.value=(s.vehicles || {})[el.dataset.costVehicle] || '';});
    document.querySelectorAll('[data-cost-route]').forEach(el=>{el.value=(s.routeZones || {})[routes[Number(el.dataset.costRoute)]] || '';});
    document.querySelectorAll('[data-cost-stop]').forEach(el=>{el.value=(s.clientStops || {})[clients[Number(el.dataset.costStop)].codigoCliente] || 'standard';});
  }
  const oldConfig=window.renderConfigClientes;
  window.renderConfigClientes=function(){oldConfig();ensureConfig();};
  const oldNote=window.abrirNotaModal;
  window.abrirNotaModal=function(fecha,camion,index,options){
    oldNote(fecha,camion,index,options);
    const box=document.querySelector('#notaModal .modal-box');
    if (!box) return;
    if (!document.getElementById('auditTripInput')) {
      const label=document.createElement('label');label.className='audit-trip-field';label.innerHTML='Número de salida del camión ese día<input id="auditTripInput" class="form-control" placeholder="1" maxlength="30">';box.insertBefore(label,box.querySelector('.modal-actions'));
    }
    const group=APP.rutas[fecha]&&APP.rutas[fecha][camion]&&APP.rutas[fecha][camion][index];
    document.getElementById('auditTripInput').value=group&&group.items[0]&&group.items[0].tripId || '1';
  };
  window.generarReporte=function(){const output=document.getElementById('reporteOutput');if(!output)return;output.innerHTML='<div id="auditCostsReport"></div>';renderCostDashboard('auditCostsReport');};
  window.exportarReporteExcel=exportCosts;
  function exportCosts(){
    const rows=costRows().map(row=>({Fecha:row.date,Camion:row.truck,Salida:row.trip,Rutas:row.routes.join(', '),Clientes:row.clients.length,TransporteUSD:row.usd,TipoCambio:row.fx,TransporteDOP:row.dop,FacturadoRealDOP:billForTrip(row),Estado:row.executed?'Confirmada':'Prevista',Pendientes:row.issues.join('; ')}));
    if(!rows.length)return toast('No hay salidas para exportar.');
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Transporte');XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(preparationRows()),'Preparacion');XLSX.writeFile(wb,'TMS_Costos_Transporte.xlsx');
  }
  async function importBilling(file){
    if(!file||!can('configuracion'))return;
    const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
    const source=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
    const value=(row,label)=>{const name=Object.keys(row).find(name=>D.key(name)===D.key(label));return name?row[name]:'';};
    const rows=[];const errors=[];
    source.forEach((row,index)=>{
      const dateValue=value(row,'Fecha despacho');
      const date=dateValue instanceof Date?`${dateValue.getFullYear()}-${String(dateValue.getMonth()+1).padStart(2,'0')}-${String(dateValue.getDate()).padStart(2,'0')}`:D.isoDate(String(dateValue));
      const truck=String(value(row,'Camion')).trim().toUpperCase();const trip=String(value(row,'Salida') || '1').trim();
      const amount=value(row,'Importe');const currency=String(value(row,'Moneda')).trim().toUpperCase();
      const invoice=String(value(row,'Factura')).trim();const customer=String(value(row,'Cliente')).trim();const order=String(value(row,'Pedido')).trim();
      if(!date||!truck||!invoice||!customer||!order||!['USD','DOP'].includes(currency)||amount===''||!Number.isFinite(Number(amount))){errors.push(index+2);return;}
      const reference=String(value(row,'Referencia')).trim(),position=String(value(row,'Linea')).trim();
      rows.push({date,truck,trip,tripKey:[date,truck,trip].join('|'),amount:Number(amount),currency,invoice,customer,order,reference,position,source:file.name,sourceLine:index+2});
    });
    if(errors.length||!rows.length)throw new Error(`No se cargó el archivo. Revisa ${errors.length ? 'las filas '+errors.slice(0,12).join(', ') : 'sus columnas y datos'}.`);
    const identity=row=>JSON.stringify([row.invoice,row.customer,row.order,row.reference||'',row.position||'',row.date,row.truck,row.trip]);
    const unique=new Set(rows.map(identity));
    if(unique.size!==rows.length)throw new Error('Hay líneas idénticas. Revisa que no se repitan importes de factura.');
    core().pushUndoState('cargar facturación real');
    const merged=new Map((settings().actualInvoices || []).map(row=>[identity(row),row]));rows.forEach(row=>merged.set(identity(row),row));settings().actualInvoices=[...merged.values()];
    save(`Facturación cargada: ${rows.length} líneas.`);ensureDashboard();ensureConfig();
  }
  function retireModules(){
    ['rutas','importaciones','almacen'].forEach(id=>{
      const nav=document.getElementById('nav-'+id);if(nav)nav.remove();
      const option=document.querySelector(`#mobileModuleSelect option[value="${id}"]`);if(option)option.remove();
    });
    const imports=document.getElementById('importaciones');if(imports)imports.hidden=true;
    const routes=document.getElementById('rutas');if(routes)routes.hidden=true;
    const warehouse=document.getElementById('almacen');if(warehouse)warehouse.hidden=true;
  }
  const oldView=window.cambiarVista;
  window.cambiarVista=function(id,button){
    if(id==='importaciones')id='materialesTransito';
    if(id==='rutas')id='configClientes';
    if(id==='almacen')id='solicitudesAlmacen';
    oldView(id,document.getElementById('nav-'+id) || button);
    const content=document.getElementById('content');if(content)content.scrollTop=0;
    if(innerWidth<=900)window.scrollTo(0,0);
    if(id==='dashboard')ensureDashboard();if(id==='configClientes')ensureConfig();
  };
  document.addEventListener('click',event=>{
    const button=event.target.closest('[data-audit]');if(!button)return;
    const action=button.dataset.audit;
    if(action==='transit-tab'){APP.auditTransitTab=button.dataset.value==='received'?'received':'active';transitView();}
    if(action==='receive-invoice'){
      if(!can('materialesTransito'))return toast('Tu perfil no puede confirmar recepciones.');
      const received=button.dataset.value==='receive';
      if(!confirm(`${received?'Confirmar recepción de todas las líneas de':'Volver a tránsito todas las líneas de'} la factura ${button.dataset.invoice}?`))return;
      core().pushUndoState('actualizar recepción de factura');
      APP.transitMaterials=D.receiveInvoice(APP.transitMaterials,button.dataset.invoice,received,(APP.currentUserProfile || {}).email || 'Administrador',new Date().toISOString());
      refresh();save(received?'Factura recibida. Sus avisos activos se retiraron.':'Factura devuelta a tránsito.');
    }
    if(action==='create-client'){
      if(!can('configuracion'))return;
      const code=document.getElementById('auditClientCode').value.trim(),name=document.getElementById('auditClientName').value.trim();
      if(!code||!name)return toast('Completa el código SAP y el nombre.');
      if(!/^[a-z0-9_.-]+$/i.test(code))return toast('El código SAP debe contener letras, números, puntos o guiones.');
      if(APP.clienteConfig[code])return toast('Ese cliente ya existe. Búscalo en la configuración.');
      core().pushUndoState('crear configuración de cliente');APP.clienteConfig[code]={...core().getClienteConfigV2(code,name),nombreCliente:name,codigoCliente:code};APP.configClienteSel=code;window.renderConfigClientes();save('Cliente creado. Asigna su ruta y recepción.');
    }
    if(action==='save-tariffs'){
      if(!can('configuracion'))return;
      const preparation=document.getElementById('auditPreparationRate').value;
      if(preparation!==''&&(!Number.isFinite(Number(preparation))||Number(preparation)<0))return toast('Revisa la tarifa de preparación.');
      core().pushUndoState('configurar tarifas CT');const s=settings();s.vehicles={};s.routeZones={};s.clientStops={};
      document.querySelectorAll('[data-cost-vehicle]').forEach(el=>{s.vehicles[el.dataset.costVehicle]=el.value;});
      const routes=(APP.routeConfigs || []).map(row=>row.nombre||row.zona).filter(Boolean);document.querySelectorAll('[data-cost-route]').forEach(el=>{s.routeZones[routes[Number(el.dataset.costRoute)]]=el.value;});
      const clients=Object.values(APP.clienteConfig || {});document.querySelectorAll('[data-cost-stop]').forEach(el=>{s.clientStops[clients[Number(el.dataset.costStop)].codigoCliente]=el.value;});
      s.preparationRateUsd=preparation===''?null:Number(preparation);save('Configuración tarifaria guardada.');ensureDashboard();
    }
    if(action==='save-fx'){
      if(!can('configuracion'))return;
      const date=document.getElementById('auditFxDate').value,rate=Number(document.getElementById('auditFxRate').value);
      if(!date||!Number.isFinite(rate)||rate<=0)return toast('Indica fecha y tipo de cambio positivo.');
      core().pushUndoState('guardar tipo de cambio');settings().fxByDate=settings().fxByDate || {};settings().fxByDate[date]=rate;save('Tipo de cambio guardado para '+date);ensureDashboard();
    }
    if(action==='lock-trip'){
      if(!can('calendario')&&!can('configuracion'))return;
      const row=D.calculateTrips(APP).find(row=>row.id===button.dataset.key);
      if(!row||row.usd===null)return;
      if(!confirm('¿Confirmar que esta salida se realizó y fijar su tarifa? No modifica las cantidades facturadas.'))return;
      core().pushUndoState('confirmar salida y tarifa');const {lines,...snapshot}=row;settings().tripHistory=settings().tripHistory || {};settings().tripHistory[row.id]={...snapshot,lockedAt:new Date().toISOString()};save('Salida confirmada y tarifa conservada.');ensureDashboard();
    }
    if(action==='export-costs')exportCosts();
  });
  document.addEventListener('input',event=>{
    if(event.target.id==='auditTransitSearch'){
      const cursor=event.target.selectionStart;APP.auditTransitSearch=event.target.value;transitView();const input=document.getElementById('auditTransitSearch');input.focus();input.setSelectionRange(cursor,cursor);
    }
  });
  document.addEventListener('change',event=>{
    if(['auditCostFrom','auditCostTo'].includes(event.target.id))renderCostDashboard();
    if(event.target.id==='auditBillingFile')importBilling(event.target.files[0]).catch(error=>toast(error.message));
  });
  window.addEventListener('DOMContentLoaded',()=>{
    const glass=document.querySelector('link[href*="tms-glass.css"]');if(glass)document.head.appendChild(glass);
    retireModules();ensureConfig();ensureDashboard();
    const title=document.querySelector('#materialesTransito > .view-title');if(title)title.textContent='Disponibilidad';
    const queue=document.querySelector('#queueSemanaSection .card-title');if(queue)queue.textContent='Pedidos pendientes de asignar';
    const queueCaption=document.querySelector('#queueSemanaSection .card-title + div');if(queueCaption)queueCaption.textContent='Pedidos del plan y del control diario, agrupados por ruta para asignarlos juntos.';
    document.querySelectorAll('.nav-icon').forEach(el=>el.setAttribute('aria-hidden','true'));
    const header=document.querySelector('#appHeader .logo');if(header&&!document.getElementById('auditWorkspaceLabel'))header.insertAdjacentHTML('afterend','<div id="auditWorkspaceLabel"><strong>Operaciones</strong><span>Álvarez · TMS</span></div>');
    const login=document.querySelector('.login-card');if(login&&!login.querySelector('.audit-login-title'))login.querySelector('.login-logo').insertAdjacentHTML('afterend','<div class="audit-login-title"><span>Tu operación, conectada</span><h1>Bienvenido de nuevo.</h1><p>Accede a tu espacio de logística.</p></div>');
  });
})();

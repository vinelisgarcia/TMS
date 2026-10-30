/* Read-only calendar drilldowns. Native dialogs provide focus management and Escape. */
(function () {
  'use strict';
  const D = window.TMSDomain;
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const qty = value => new Intl.NumberFormat('es-DO', { maximumFractionDigits: 2 }).format(Number(value) || 0);
  const money = value => new Intl.NumberFormat('es-DO', { style:'currency', currency:'DOP' }).format(Number(value) || 0);
  const date = value => { const iso = D.isoDate(value); return iso ? iso.split('-').reverse().join('/') : 'Por confirmar'; };
  let opener = null;
  function transit(lines) {
    return (lines || []).map(line => ({ line, info: D.transitForLine(line, APP.transitMaterials) })).filter(row => row.info.note);
  }
  function chips(group) {
    const invoices = new Map();
    transit(group.items).forEach(({info}) => info.rows.forEach(row => {
      const id = JSON.stringify([row.factura || '', row.eta || '']);
      if (!invoices.has(id)) invoices.set(id, row);
    }));
    return [...invoices.values()].map(row => `<button type="button" class="calendar-transit-chip" data-calendar-transit aria-haspopup="dialog">${esc(row.factura || 'Sin factura')} <span>· ETA ${esc(date(row.eta))}</span></button>`).join('');
  }
  function close() { const dialog = document.getElementById('calendarDetailDialog'); if (dialog?.open) dialog.close(); }
  function show(group, onlyTransit, source) {
    if (!group) return;
    opener = source;
    let dialog = document.getElementById('calendarDetailDialog');
    if (!dialog) {
      dialog = document.createElement('dialog'); dialog.id = 'calendarDetailDialog'; dialog.className = 'calendar-detail-dialog';
      dialog.setAttribute('aria-labelledby', 'calendarDetailTitle'); document.body.appendChild(dialog);
      dialog.addEventListener('click', event => { if (event.target === dialog || event.target.closest('[data-calendar-close]')) close(); });
      dialog.addEventListener('close', () => { if (opener?.isConnected) opener.focus(); });
    }
    const lines = group.items || [];
    const matches = transit(lines);
    const actualInvoices = (APP.costSettings?.actualInvoices || []).filter(row =>
      row.date === D.isoDate(group.fecha) && row.truck === group.camion &&
      [group.codigo,group.nombre].some(value => D.key(value) === D.key(row.customer)) &&
      lines.some(line => String(line.pedidoCliente) === String(row.order) && String(line.tripId || '1') === String(row.trip || '1') &&
        (!row.reference || D.key(row.reference) === D.key(line.articulo)) && (!row.position || String(row.position) === String(line.lineaPedidoCliente)))
    );
    const billing = actualInvoices.length ? `<div class="calendar-detail-scroll"><table class="data-table"><thead><tr><th>Factura SAP</th><th>Pedido</th><th>Importe facturado</th></tr></thead><tbody>${actualInvoices.map(row => `<tr><td>${esc(row.invoice)}</td><td>${esc(row.order)}</td><td>${esc(row.currency)} ${qty(row.amount)}</td></tr>`).join('')}</tbody></table></div>` : '<p class="calendar-detail-caption">No hay importes de facturación efectiva SAP cargados para este cliente, fecha y salida. Las cantidades facturadas del control se muestran arriba.</p>';
    const sum = key => lines.reduce((total, row) => total + (Number(row[key]) || 0), 0);
    const table = `<div class="calendar-detail-scroll"><table class="data-table"><caption>Líneas incluidas en esta ficha · ${esc(group.fecha)} · ${esc(group.camion)}</caption><thead><tr><th>Pedido / referencia</th><th>Confirmado disponible</th><th>Facturado</th><th>Pendiente operativo</th><th>Sin confirmar</th><th>Confirmado no entregado</th><th>Entregado no facturado</th></tr></thead><tbody>${lines.map(row => `<tr><td><strong>${esc(row.pedidoCliente || 'Sin pedido')}</strong> · ${esc(row.lineaPedidoCliente || '—')}<br>${esc(row.articulo || '—')}<small>${esc(row.descripcionArticulo || '')}<br>Solicitado original: ${qty(row.cantidadSolicitadaOriginal ?? row.cantidadSolicitada)}</small></td><td>${qty(row.cantidadSolicitada)}</td><td>${qty(row.cantidadFacturada)}</td><td>${qty(row.cantidadPendiente)}</td><td>${qty(row.cantidadPendienteSinStock)}</td><td>${money(row.montoPlanificado)}</td><td>${money(row.montoEntregadoNoFacturado || row.importeEntregadoNoFacturado || row.montoPendienteFacturar)}</td></tr>`).join('')}</tbody></table></div>`;
    const transitHtml = matches.length ? matches.map(({line,info}) => `<article class="calendar-transit-detail"><h3>${esc(line.pedidoCliente || 'Sin pedido')} · ${esc(line.articulo)}</h3><p>${esc(line.descripcionArticulo || '')}</p><p>${esc(info.note)}</p><div class="calendar-detail-scroll"><table class="data-table"><thead><tr><th>Factura</th><th>ETA</th><th>Material</th><th>Cantidad en tránsito</th></tr></thead><tbody>${info.rows.map(row => `<tr><td>${esc(row.factura || 'Sin identificar')}</td><td>${esc(date(row.eta))}</td><td>${esc(row.codigoMaterial)}<small>${esc(row.descripcion || '')}</small></td><td>${qty(row.cantidad)}</td></tr>`).join('')}</tbody></table></div></article>`).join('') : '<p>No hay coincidencias activas de materiales en tránsito para esta ficha.</p>';
    dialog.innerHTML = `<header class="calendar-detail-head"><div><p class="calendar-detail-eyebrow">${onlyTransit ? 'Materiales en tránsito' : 'Detalle de cliente y ruta'}</p><h2 id="calendarDetailTitle">${esc(group.nombre)}</h2><p>${esc(group.codigo)} · ${esc(group.rutaNombre || group.zona || 'Sin ruta')}<br>${esc(group.fecha)} · ${esc(group.camion)}</p></div><button type="button" class="btn btn-outline" data-calendar-close aria-label="Cerrar detalle">Cerrar</button></header><div class="calendar-detail-body">${onlyTransit ? '' : `<div class="calendar-detail-summary"><div><span>Confirmado disponible · uds</span><strong>${qty(sum('cantidadSolicitada'))}</strong></div><div><span>Facturado · uds</span><strong>${qty(sum('cantidadFacturada'))}</strong></div><div><span>Pendiente operativo · uds</span><strong>${qty(sum('cantidadPendiente'))}</strong></div><div><span>Sin confirmar · uds</span><strong>${qty(sum('cantidadPendienteSinStock'))}</strong></div></div><p class="calendar-detail-caption">Cantidades e importes del plan y control cargados para esta ficha. Se muestran los saldos registrados de cada línea; no se infiere un estado de facturación a partir del tránsito. El importe facturado no se infiere de las cantidades.</p>${table}<div class="calendar-detail-amounts"><p><span>Confirmado no entregado</span><strong>${money(sum('montoPlanificado'))}</strong></p><p><span>Entregado no facturado</span><strong>${money(lines.reduce((total,row)=>total+Number(row.montoEntregadoNoFacturado || row.importeEntregadoNoFacturado || row.montoPendienteFacturar || 0),0))}</strong></p></div><h3>Facturación efectiva SAP</h3>${billing}<h3>Notas y solicitudes</h3><p>${esc(group.nota || 'Sin notas.')} ${esc(group.incidencia || '')}</p><p>Solicitudes: ${esc((group.solicitudes || []).join(', ') || 'Sin solicitudes asociadas')}</p>`}<h3>Facturas y llegada estimada</h3><p class="calendar-detail-caption">Las coincidencias de tránsito indican posible cobertura; no son reservas ni stock confirmado.</p>${transitHtml}</div>`;
    if (!dialog.open) dialog.showModal();
  }
  function openFromCard(card, onlyTransit, source) {
    const groups = Object.values(APP.rutas || {}).flatMap(day => Object.values(day).flat());
    const group = groups.find(row => row.id === card.dataset.calendarGroup);
    show(group, onlyTransit, source);
  }
  document.addEventListener('click', event => {
    const card = event.target.closest('[data-calendar-group]');
    if (!card) return;
    if (event.target.closest('[data-calendar-transit]')) { openFromCard(card, true, event.target.closest('button')); return; }
    if (event.target.closest('button,input,label,a,select,textarea')) return;
    openFromCard(card, false, card);
  });
  document.addEventListener('keydown', event => {
    const card = event.target.closest('[data-calendar-group]');
    if (card && event.target === card && ['Enter',' '].includes(event.key)) { event.preventDefault(); openFromCard(card, false, card); }
  });
  window.TMSCalendarDetails = { chips };
})();

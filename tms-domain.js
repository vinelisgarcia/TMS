/* Shared, side-effect-free business rules. Prices: CT, Agosto 2026, USD. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.TMSDomain = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const key = value => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const rates = {
    puj: { label: 'Punta Cana / Bávaro', large: 240, small: 195, container: 1250 },
    uvero: { label: 'Uvero Alto', large: 320, small: 225, container: 1325 },
    romana: { label: 'Bayahíbe / La Romana / Casa de Campo', large: 450, small: 295, container: 1860 },
    puertoPlata: { label: 'Puerto Plata', large: 955, small: 780, container: 2560 },
    santoDomingo: { label: 'Santo Domingo', large: 650, small: 520, container: 2500 },
    santiago: { label: 'Santiago', large: 850, small: 715, container: 1740 },
    samana: { label: 'Samaná', large: 1150, small: 995, container: 2250 }
  };
  const extras = { standard: { large: 125, small: 75 }, bahia: { large: 150, small: 175 } };
  const isoDate = value => {
    if (/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return value;
    const m = String(value || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return m ? `${m[3]}-${m[2]}-${m[1]}` : '';
  };
  function invoiceKey(row) { return String(row.factura || '').trim().toUpperCase(); }
  function activeTransit(row) { return !row.receivedAt && row.estado !== 'recibida'; }
  function transitForLine(line, materials) {
    const pending = Math.max(0, Number(line.cantidadPendienteSinStock) || 0);
    const result = { pending, quantity: 0, covered: 0, invoices: [], eta: '', rows: [], note: '' };
    if (!pending || !line.articulo) return result;
    const rows = (materials || []).filter(row => {
      if (!activeTransit(row) || key(row.codigoMaterial) !== key(line.articulo)) return false;
      if (row.cliente && ![line.clienteId, line.clienteNombre].some(value => key(value) === key(row.cliente))) return false;
      if (row.pedidoCliente && key(row.pedidoCliente) !== key(line.pedidoCliente)) return false;
      return Number(row.cantidad) > 0;
    }).sort((a, b) => (a.eta || '9999').localeCompare(b.eta || '9999'));
    result.rows = rows;
    result.quantity = rows.reduce((sum, row) => sum + Number(row.cantidad), 0);
    result.covered = Math.min(result.quantity, pending);
    let quantity = 0;
    for (const row of rows) {
      if (quantity >= pending) break;
      quantity += Number(row.cantidad);
      if (row.factura && !result.invoices.includes(row.factura)) result.invoices.push(row.factura);
      result.eta = row.eta || result.eta;
    }
    if (rows.length) result.note = `${pending} uds sin confirmar · ${result.covered} uds con coincidencia en tránsito · Factura ${result.invoices.join(', ') || 'sin identificar'} · ETA ${result.eta || 'por confirmar'} · Posible cobertura ${result.covered >= pending ? 'completa' : 'parcial'}, sin reserva confirmada`;
    return result;
  }
  function calendarDispatchSummary(lines, materials) {
    const number = value => Number.isFinite(Number(value)) ? Number(value) : 0;
    const orders = new Set(), invoices = new Map();
    let confirmed = 0, pendingBilling = 0, unidentifiedOrders = 0;
    (lines || []).forEach(line => {
      const order = String(line.pedidoCliente || '').trim();
      if (order && order !== 'Sin pedido') orders.add(order);
      else unidentifiedOrders++;
      confirmed += number(line.montoPlanificado);
      pendingBilling += number(line.montoEntregadoNoFacturado || line.importeEntregadoNoFacturado || line.montoPendienteFacturar || 0);
      const transit = transitForLine(line, materials);
      const matches = new Map();
      transit.rows.forEach(row => {
        const invoice = invoiceKey(row);
        if (!invoices.has(invoice)) invoices.set(invoice, { key: invoice, invoice: String(row.factura || '').trim(), dates: new Set(), affected: [] });
        const entry = invoices.get(invoice);
        if (row.eta) entry.dates.add(row.eta);
        matches.set(invoice, (matches.get(invoice) || 0) + number(row.cantidad));
      });
      matches.forEach((quantity, invoice) => invoices.get(invoice).affected.push({ line, pending: transit.pending, transitQuantity: quantity }));
    });
    return {
      orders: orders.size, unidentifiedOrders, confirmed, pendingBilling, total: confirmed + pendingBilling,
      invoices: [...invoices.values()].map(entry => ({ ...entry, dates: [...entry.dates].sort() }))
        .sort((a, b) => (a.dates[0] || '9999').localeCompare(b.dates[0] || '9999') || a.key.localeCompare(b.key))
    };
  }
  function receiveInvoice(materials, invoice, received, actor, timestamp) {
    const id = invoiceKey({ factura: invoice });
    if (!id) throw new Error('La factura debe estar identificada.');
    return materials.map(row => invoiceKey(row) !== id ? row : {
      ...row, estado: received ? 'recibida' : 'en_transito',
      receivedAt: received ? timestamp : '', receivedBy: received ? actor : '', updatedAt: timestamp
    });
  }
  function orderKey(line) { return [line.clienteId || line.codigo, line.pedidoCliente, line.lineaPedidoCliente, line.articulo].map(key).join('|'); }
  function archiveSplitResidues(items, timestamp) {
    for (const row of items) {
      if (row.archivedAt || Number(row.cantidadPendiente) > 0 || !(Number(row.cantidadFacturada) > 0)) continue;
      if ((Number(row.montoPlanificado) || 0) !== 0 || (Number(row.montoEntregadoNoFacturado) || 0) !== 0) continue;
      if (items.some(other => other !== row && orderKey(other) === orderKey(row) && Number(other.cantidadPendiente) > 0 &&
        (other.fechaPlanificada !== row.fechaPlanificada || other.camionAsignado !== row.camionAsignado))) row.archivedAt = timestamp;
    }
  }
  function latestOrders(state) {
    const map = new Map();
    for (const line of [...(state.planLineItems || []), ...(state.lineItems || []), ...(state.controlLineItems || [])]) {
      const id = orderKey(line);
      // Split deliveries must not replace the complete source order.
      if (!map.has(id) || !line.manualProgramado) map.set(id, line);
    }
    return [...map.values()];
  }
  function buildQueue(state) {
    const planned = new Set((state.lineItems || []).filter(row => row.fechaPlanificada && !row.archivedAt).map(orderKey));
    return latestOrders(state).filter(row => !planned.has(orderKey(row)) && Number(row.cantidadPendiente) > 0);
  }
  function calculateTrips(state) {
    const settings = state.costSettings || {};
    const groups = new Map();
    for (const line of state.lineItems || []) {
      if (!line.fechaPlanificada || line.archivedAt || line.camionAsignado === 'ALMACEN') continue;
      const date = isoDate(line.fechaPlanificada);
      const trip = String(line.tripId || '1');
      const id = [date, line.camionAsignado, trip].join('|');
      if (!groups.has(id)) groups.set(id, { id, date, truck: line.camionAsignado, trip, clients: new Map(), lines: [] });
      const group = groups.get(id);
      group.lines.push(line);
      const client = line.clienteId || line.clienteNombre;
      if (!group.clients.has(client)) group.clients.set(client, line);
    }
    const result = [...groups.values()].map(group => {
      const override = (settings.trips || {})[group.id] || {};
      const vehicle = override.vehicle || (settings.vehicles || {})[group.truck] || '';
      const issues = [];
      if (!vehicle) issues.push('Configurar tipo de vehículo');
      const clients = [...group.clients.values()].map(line => {
        const route = (state.routeConfigs || []).find(row => key(row.nombre || row.zona) === key(line.rutaNombre || line.zona));
        const zone = ((settings.routeZones || {})[line.rutaNombre || line.zona]) || (route && route.tariffZone) || '';
        const rate = rates[zone] && rates[zone][vehicle];
        if (rate === undefined) issues.push(`Tarifa pendiente: ${line.rutaNombre || line.zona || 'sin ruta'}`);
        const stop = (settings.clientStops || {})[line.clienteId] || 'standard';
        return { code: line.clienteId, name: line.clienteNombre, zone, rate, stop };
      }).sort((a, b) => (b.rate || 0) - (a.rate || 0) || String(a.code).localeCompare(String(b.code)));
      let usd = clients[0] && clients[0].rate;
      clients.forEach((client, index) => {
        if (!index) { client.charge = client.rate; return; }
        client.charge = (extras[client.stop] || {})[vehicle];
        if (client.charge === undefined) issues.push('Adicional no aplicable: revisar tarifa CT');
        else usd += client.charge;
      });
      const helper = override.helper ? 80 : 0;
      if (usd !== undefined) usd += helper;
      if (issues.length) usd = null;
      const fx = Number((settings.fxByDate || {})[group.date]) || null;
      const dop = usd !== null && fx ? usd * fx : null;
      const requests = new Set(group.lines.flatMap(line => String(line.solicitudAlmacen || '').split(',').map(s => s.trim()).filter(Boolean)));
      const history = (settings.tripHistory || {})[group.id];
      const locked = history && history.lockedAt;
      return { ...group, clients: locked ? history.clients : clients, vehicle: locked ? history.vehicle : vehicle, issues: locked ? [] : issues, helper,
        usd: locked ? history.usd : usd, fx: locked ? (history.fx || fx) : fx,
        dop: locked ? (history.dop ?? (fx ? history.usd * fx : null)) : dop, lockedAt: locked ? history.lockedAt : '', requests: requests.size,
        routes: locked ? history.routes : [...new Set(group.lines.map(line => line.rutaNombre || line.zona).filter(Boolean))],
        orderRefs: locked ? history.orderRefs : [...new Set(group.lines.map(line => key(line.clienteId)+'|'+key(line.pedidoCliente)))],
        executed: !!(locked || group.lines.some(line => line.fechaCierreRuta)) };
    });
    for (const [id, snapshot] of Object.entries(settings.tripHistory || {})) {
      if (snapshot.lockedAt && !result.some(row => row.id === id)) {
        const fx = snapshot.fx || Number((settings.fxByDate || {})[snapshot.date]) || null;
        result.push({ ...snapshot, id, fx, dop: snapshot.dop ?? (fx ? snapshot.usd * fx : null), lines: [], executed: true, issues: [] });
      }
    }
    return result.sort((a, b) => a.date.localeCompare(b.date) || a.id.localeCompare(b.id));
  }
  function warehouseCosts(state) {
    const settings = state.costSettings || {};
    const rate = settings.preparationRateUsd === null || settings.preparationRateUsd === undefined || settings.preparationRateUsd === '' ? null : Number(settings.preparationRateUsd);
    const requests = new Map();
    for (const row of [...(state.solicitudesPlanAlmacen || []), ...(state.solicitudesControlAlmacen || [])]) {
      if (!row.solicitud) continue;
      if (!requests.has(row.solicitud)) requests.set(row.solicitud, []);
      requests.get(row.solicitud).push(row);
    }
    return [...requests.entries()].map(([id, rows]) => {
      const latest = new Map();
      rows.forEach(row => latest.set([row.codigo, row.pedidoCliente, row.lineaPedidoCliente, row.articulo].map(key).join('|'), row));
      const lines = [...latest.values()];
      const prepared = lines.length > 0 && lines.every(row => Number(row.cantidadPendiente) <= 0 && Number(row.cantidadCompletada) > 0);
      const partial = !prepared && lines.some(row => Number(row.cantidadCompletada) > 0);
      const dates=lines.map(row=>isoDate(row.fechaFinalizada || row.fechaEnvio || row.fechaLiberacion || '')).filter(Boolean).sort();
      return { id, date: dates.at(-1) || '', prepared, partial, lines: lines.length, usd: prepared && rate !== null ? rate : null };
    });
  }
  function reminderSeeds() {
    return [
      { id: 'audit-20261023-pod', titulo: 'Validar tarifa de preparación con CT', notas: 'Confirmar si USD 15 por POD emitido equivale al cargo por cada solicitud de almacén.' },
      { id: 'audit-20261023-facturado', titulo: 'Adjuntar reporte SAP de lo realmente facturado', notas: 'Cargar importes efectivos por cliente, pedido, referencia y fecha para comparar facturación con costo de transporte.' }
    ].map(row => ({ ...row, fecha: '2026-10-23', estado: 'pendiente', createdAt: '2026-10-22T12:00:00Z', updatedAt: '2026-10-22T12:00:00Z' }));
  }
  return { key, rates, extras, isoDate, invoiceKey, activeTransit, transitForLine, calendarDispatchSummary, receiveInvoice, orderKey, archiveSplitResidues, latestOrders, buildQueue, calculateTrips, warehouseCosts, reminderSeeds };
});

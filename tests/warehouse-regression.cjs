const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync(require('node:path').join(__dirname, '../tms-v2.js'), 'utf8');
const context = {
  APP: { warehouseSettings: {}, lineItems: [], planLineItems: [], controlLineItems: [], solicitudesPlanAlmacen: [], solicitudesControlAlmacen: [] },
  normalizarFecha: x => String(x || ''), fechaToStr: () => '17/09/2026',
  getRouteConfigByName: () => null, isUuid: () => false,
  fechaStrToISO: x => x || null, safeClone: x => structuredClone(x),
  deriveItemStatus: () => 'pendiente', chooseTruckForItem: () => 'CAMION 1',
  XLSX: process.env.TMS_XLSX ? require(process.env.TMS_XLSX) : null
};
vm.createContext(context);
for (const name of ['stripAccents','normKey','text','num','normalizeRow','pickField','toDateLabel','dedupeBy','makeUniqueHeader','rowsFromWorksheet','parseWarehouseRows','getUsableRequestedQty','buildSapMatchKey','uniqueTexts','getConfirmedUndeliveredAmount','matchWarehouseToOrders','buildOrderLineDbRows']) {
  const start = source.indexOf('  function ' + name + '(');
  assert.ok(start >= 0, name);
  const end = source.indexOf('\n  }', start) + 4;
  vm.runInContext(source.slice(start, end), context);
}
const order = { baseKey:'same', clienteId:'C1', pedidoCliente:'P1', lineaPedidoCliente:'10', articulo:'A1', cantidadSolicitada:100, montoPlanificado:1000, fechaPlanificada:'2026-09-17' };
const request = { codigo:'C1', pedidoCliente:'P1', lineaPedidoCliente:'10', articulo:'A1', solicitud:'S1', cantidadSolicitada:20, estadoProceso:'Liberado' };
context.APP.solicitudesPlanAlmacen = [request];
context.matchWarehouseToOrders(); // requests before orders
context.APP.lineItems = [{...order}];
context.matchWarehouseToOrders();
assert.equal(context.APP.lineItems[0].montoAlmacenLiberado, 200);
context.APP.solicitudesControlAlmacen = [{...request}];
context.matchWarehouseToOrders();
assert.equal(context.APP.lineItems[0].montoAlmacenLiberado, 200, 'plan/control counted once');
context.APP.lineItems.push({...order});
context.matchWarehouseToOrders();
assert.equal(context.APP.lineItems.reduce((s,x)=>s+x.montoAlmacenLiberado,0), 200, 'split lines do not double amount');
context.APP.solicitudesPlanAlmacen = [{...request,codigo:'OTHER'}];
context.APP.solicitudesControlAlmacen = [];
context.matchWarehouseToOrders();
assert.equal(context.APP.lineItems[0].montoAlmacenLiberado, 0, 'no cross-client or stale match');
const saved = context.buildOrderLineDbRows([order,{...order,montoPlanificado:2000}], 'current');
assert.equal(saved.length,2);
assert.equal(new Set(saved.map(x=>x.clave_unica)).size,2);
assert.equal(saved[1].metadata.montoPlanificado,2000);
assert.notEqual(saved[0].clave_unica,context.buildOrderLineDbRows([order],'plan')[0].clave_unica);
if (process.argv[2]) {
  assert.ok(context.XLSX, 'Set TMS_XLSX to SheetJS path');
  const wb = context.XLSX.read(fs.readFileSync(process.argv[2]), {type:'buffer',cellDates:true});
  const parsed = context.parseWarehouseRows(context.rowsFromWorksheet(wb.Sheets[wb.SheetNames[0]]));
  assert.equal(parsed.length,173);
  assert.equal(new Set(parsed.map(x=>x.solicitud)).size,61);
  assert.equal(new Set(parsed.map(x=>x.codigo)).size,14);
  assert.ok(parsed.every(x=>x.estadoProceso==='Liberado'));
  console.log('User XML: 173 lines / 61 requests / 14 clients, all Liberado');
}
console.log('Warehouse matching and duplicate cloud-key regressions passed');

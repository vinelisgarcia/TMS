const {test}=require('node:test'),assert=require('node:assert/strict');
const D=require('../tms-domain.js');
const line=extra=>({clienteId:'C1',clienteNombre:'Cliente 1',pedidoCliente:'P1',lineaPedidoCliente:'1',articulo:'A',cantidadPendienteSinStock:3,montoPlanificado:100,montoEntregadoNoFacturado:20,...extra});
test('calendar counts distinct orders and sums only the supplied dispatch, without mutating it',()=>{
 const lines=[line({}),line({lineaPedidoCliente:'2',articulo:'B',montoPlanificado:50,montoEntregadoNoFacturado:5}),line({pedidoCliente:'P2',montoPlanificado:200,montoEntregadoNoFacturado:0})];
 const before=JSON.stringify(lines);const result=D.calendarDispatchSummary(lines,[]);
 assert.equal(result.orders,2);assert.equal(result.confirmed,350);assert.equal(result.pendingBilling,25);assert.equal(result.total,375);assert.deepEqual(result.invoices,[]);assert.equal(JSON.stringify(lines),before);
});
test('invoice notices group materials and preserve every ETA, excluding received and other clients/orders',()=>{
 const rows=[{factura:'F-1',codigoMaterial:'A',cantidad:2,eta:'2026-11-05'},{factura:' f-1 ',codigoMaterial:'B',cantidad:1,eta:'2026-11-06'},{factura:'OTHER-CLIENT',codigoMaterial:'A',cantidad:9,cliente:'C2'},{factura:'OTHER-ORDER',codigoMaterial:'A',cantidad:9,pedidoCliente:'P2'},{factura:'RECEIVED',codigoMaterial:'A',cantidad:9,receivedAt:'2026-10-31'},{factura:'OTHER-MATERIAL',codigoMaterial:'Z',cantidad:9}];
 const result=D.calendarDispatchSummary([line({}),line({articulo:'B',lineaPedidoCliente:'2'}),line({articulo:'Z',cantidadPendienteSinStock:0})],rows);
 assert.equal(result.invoices.length,1);assert.equal(result.invoices[0].key,'F-1');assert.deepEqual(result.invoices[0].dates,['2026-11-05','2026-11-06']);assert.equal(result.invoices[0].affected.length,2);assert.equal(result.invoices[0].affected[0].pending,3);
});
test('every relevant invoice is kept separately and transit quantity is not an allocated reservation',()=>{
 const summary=D.calendarDispatchSummary([line({})],[{factura:'F-1',codigoMaterial:'A',cantidad:20},{factura:'F1',codigoMaterial:'A',cantidad:2,eta:'2026-11-07'}]);
 assert.equal(summary.invoices.length,2);assert.equal(summary.invoices.find(row=>row.invoice==='F-1').affected[0].transitQuantity,20);assert.equal(summary.invoices[0].affected[0].pending,3);
});
test('unknown invoice/ETA and unknown order remain explicit rather than counting lines as orders',()=>{
 const summary=D.calendarDispatchSummary([line({pedidoCliente:''}),line({pedidoCliente:'',lineaPedidoCliente:'2'})],[{codigoMaterial:'A',cantidad:2}]);
 assert.equal(summary.orders,0);assert.equal(summary.unidentifiedOrders,2);assert.equal(summary.invoices.length,1);assert.equal(summary.invoices[0].invoice,'');assert.deepEqual(summary.invoices[0].dates,[]);
});
test('zero balances and pending-billing aliases keep their original accounting meaning',()=>{
 const summary=D.calendarDispatchSummary([line({montoPlanificado:0,montoEntregadoNoFacturado:0,importeEntregadoNoFacturado:12}),line({montoPlanificado:0,montoEntregadoNoFacturado:0,montoPendienteFacturar:3})],[]);
 assert.equal(summary.confirmed,0);assert.equal(summary.pendingBilling,15);assert.equal(summary.total,15);
});

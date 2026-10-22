const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const D = require('../tms-domain.js');
const source = fs.readFileSync(require('node:path').join(__dirname, '../tms-v2.js'), 'utf8');
function coreFunction(name, context = {}) {
  const start=source.indexOf('  function '+name+'(');assert.ok(start>=0,name);
  const end=source.indexOf('\n  }',start)+4;
  const sandbox=vm.createContext({...context});vm.runInContext(source.slice(start,end),sandbox);return sandbox[name];
}
const line = (client,route,extra={}) => ({clienteId:client,clienteNombre:'Cliente '+client,pedidoCliente:'P'+client,lineaPedidoCliente:'1',articulo:'A1',rutaNombre:route,zona:route,fechaPlanificada:'23/10/2026',camionAsignado:'CAMION 1',cantidadSolicitada:10,cantidadPendiente:10,...extra});
const costState = items => ({lineItems:items,routeConfigs:[{nombre:'PP'},{nombre:'SD'}],costSettings:{vehicles:{'CAMION 1':'large'},routeZones:{PP:'puertoPlata',SD:'santoDomingo'},fxByDate:{'2026-10-23':60}}});

test('same route stays together, incompatible clients remain queued',()=>{
  const context={text:x=>String(x||''),num:Number,fechaInicioPlanificacion:x=>x,fechaToStr:()=>'',appendWorkingDates:()=>['23/10/2026','26/10/2026'],getClienteConfigV2:()=>({rutaAsignada:'PP'}),getRouteConfigByName:()=>({activa:true}),chooseTruckForItem:()=> 'CAMION 1',clienteAllowsDate:()=>true,routeAllowsDate:()=>true};
  const schedule=coreFunction('schedulePlanItems',context);const rows=[line('1','PP'),line('2','PP')];schedule(rows,'23/10/2026');assert.equal(rows[0].fechaPlanificada,rows[1].fechaPlanificada);
  const incompatible=coreFunction('schedulePlanItems',{...context,clienteAllowsDate:(row,date)=>row.clienteId==='1'?date==='23/10/2026':date==='26/10/2026'});
  incompatible(rows,'23/10/2026');assert.ok(rows.every(row=>row.fechaPlanificada===''));assert.ok(rows.every(row=>row.queueReason));
});
test('unconfigured client remains in initial queue without daily control',()=>{
  const row=line('1','PP',{fechaPlanificada:''});const state={planLineItems:[row],controlLineItems:[],lineItems:[]};
  assert.equal(D.buildQueue(state).length,1);state.lineItems=[{...row,fechaPlanificada:'23/10/2026'}];assert.equal(D.buildQueue(state).length,0);
});
test('daily control with zero pending does not requeue an old pending order',()=>{
  const row=line('1','PP',{fechaPlanificada:''});assert.equal(D.buildQueue({planLineItems:[row],controlLineItems:[{...row,cantidadPendiente:0}],lineItems:[]}).length,0);
});
test('partial move archives closed residue while preserving invoiced quantities',()=>{
  const APP={lineItems:[]};const row=line('1','PP',{cantidadFacturada:4,cantidadPendiente:6,montoPlanificado:600});APP.lineItems.push(row);
  const move=coreFunction('splitAndMovePendingLine',{APP,num:Number,text:String,nowIso:()=> '2026-10-22T12:00:00Z'});move(row,'26/10/2026','CAMION 2');
  assert.equal(APP.lineItems.length,2);assert.ok(row.archivedAt);assert.equal(row.cantidadFacturada,4);assert.equal(row.cantidadPendiente,0);assert.equal(row.montoPlanificado,0);
  assert.equal(APP.lineItems[1].cantidadPendiente,6);assert.equal(APP.lineItems[1].fechaPlanificada,'26/10/2026');assert.equal(APP.lineItems[1].archivedAt,'');
});
test('numeric parsing handles US and decimal-comma inputs',()=>{
  const num=coreFunction('num');assert.equal(num('1,234.56'),1234.56);assert.equal(num('1.234,56'),1234.56);assert.equal(num('1,5'),1.5);assert.equal(num(0),0);
});
test('receiving one invoice updates all its lines, keeps history and other invoices',()=>{
  const rows=[{factura:'F1',codigoMaterial:'A1',cantidad:4},{factura:'F1',codigoMaterial:'A2',cantidad:2},{factura:'F2',codigoMaterial:'A1',cantidad:3}];
  const received=D.receiveInvoice(rows,'F1',true,'user','2026-10-23T12:00:00Z');assert.equal(received.length,3);assert.equal(received.filter(D.activeTransit).length,1);assert.equal(received[0].receivedBy,'user');
  assert.equal(D.receiveInvoice(received,'F1',false,'user','later').filter(D.activeTransit).length,3);assert.throws(()=>D.receiveInvoice(rows,'',true,'user','now'));
});
test('invoice receipt does not conflate different punctuation in invoice identifiers',()=>{
  const received=D.receiveInvoice([{factura:'F-1'},{factura:'F1'}],'F-1',true,'user','now');
  assert.equal(received[0].receivedAt,'now');assert.equal(received[1].receivedAt,undefined);
});
test('legacy split residues are archived without hiding legitimate zero-value pending rows',()=>{
  const residue=line('1','PP',{cantidadFacturada:4,cantidadPendiente:0,montoPlanificado:0});
  const pending=line('1','PP',{fechaPlanificada:'26/10/2026',cantidadPendiente:6,montoPlanificado:0});
  const complete=line('2','SD',{cantidadFacturada:5,cantidadPendiente:0,montoPlanificado:0});
  D.archiveSplitResidues([residue,pending,complete],'now');
  assert.equal(residue.archivedAt,'now');assert.equal(pending.archivedAt,undefined);assert.equal(complete.archivedAt,undefined);
});
test('commercial notices cover unconfirmed quantity, respect allocation and received invoices',()=>{
  const order=line('1','PP',{cantidadPendiente:0,cantidadPendienteSinStock:6});
  const rows=[{factura:'F1',codigoMaterial:'A1',cantidad:4,eta:'2026-11-05'},{factura:'F2',codigoMaterial:'A1',cantidad:20,cliente:'2'},{factura:'F3',codigoMaterial:'A1',cantidad:10,receivedAt:'yesterday'}];
  const info=D.transitForLine(order,rows);assert.equal(info.covered,4);assert.equal(info.pending,6);assert.ok(info.note.includes('parcial'));assert.deepEqual(info.invoices,['F1']);assert.equal(info.eta,'2026-11-05');
});
test('full order remains discoverable even with no warehouse request and no confirmed stock',()=>{
  const row=line('1','PP',{cantidadSolicitada:0,cantidadPendiente:0,cantidadSolicitadaOriginal:10,cantidadPendienteSinStock:10});
  const state={planLineItems:[row],lineItems:[],controlLineItems:[]};assert.equal(D.latestOrders(state).length,1);assert.equal(D.latestOrders(state)[0].cantidadSolicitadaOriginal,10);
});
test('CT trip uses largest destination plus per-client additional, never two full routes',()=>{
  const trips=D.calculateTrips(costState([line('1','PP'),line('2','SD')]));assert.equal(trips.length,1);assert.equal(trips[0].usd,1080);assert.equal(trips[0].dop,64800);assert.equal(trips[0].executed,false);
});
test('multiple orders of one client produce one charge; multiple departures stay separate',()=>{
  const state=costState([line('1','PP'),line('1','PP',{pedidoCliente:'other'}),line('2','SD',{tripId:'2'})]);
  const trips=D.calculateTrips(state);assert.equal(trips.length,2);assert.equal(trips[0].clients.length,1);assert.equal(trips[0].usd,955);assert.equal(trips[1].usd,650);
});
test('small truck and Bahia exception retain exact source prices',()=>{
  const state=costState([line('1','PP'),line('2','SD')]);state.costSettings.vehicles['CAMION 1']='small';state.costSettings.clientStops={'2':'bahia'};
  assert.equal(D.calculateTrips(state)[0].usd,955); // 780 + 175
});
test('missing vehicle/zone/FX and non-applicable container stop are not zero costs',()=>{
  const state=costState([line('1','PP'),line('2','SD')]);state.costSettings.vehicles={};assert.equal(D.calculateTrips(state)[0].usd,null);
  state.costSettings.vehicles={'CAMION 1':'container'};assert.equal(D.calculateTrips(state)[0].usd,null);
  state.costSettings.vehicles={'CAMION 1':'large'};state.costSettings.fxByDate={};assert.equal(D.calculateTrips(state)[0].usd,1080);assert.equal(D.calculateTrips(state)[0].dop,null);
});
test('confirmed trip snapshot survives config changes and a new monthly plan',()=>{
  const state=costState([line('1','PP'),line('2','SD')]);const trip=D.calculateTrips(state)[0];
  state.costSettings.tripHistory={[trip.id]:{...trip,lines:undefined,lockedAt:'2026-10-23T10:00:00Z'}};
  state.costSettings.vehicles['CAMION 1']='small';state.costSettings.fxByDate['2026-10-23']=80;
  assert.equal(D.calculateTrips(state)[0].usd,1080);assert.equal(D.calculateTrips(state)[0].dop,64800);
  state.lineItems=[];assert.equal(D.calculateTrips(state).length,1);assert.equal(D.calculateTrips(state)[0].usd,1080);
});
test('preparation counts a complete request once, not each line or plan/control copy',()=>{
  const rows=[{solicitud:'S1',codigo:'C1',pedidoCliente:'P1',articulo:'A1',cantidadCompletada:2,cantidadPendiente:0},{solicitud:'S1',codigo:'C1',pedidoCliente:'P1',articulo:'A2',cantidadCompletada:3,cantidadPendiente:0}];
  const state={solicitudesPlanAlmacen:rows,solicitudesControlAlmacen:rows,costSettings:{preparationRateUsd:15}};
  const costs=D.warehouseCosts(state);assert.equal(costs.length,1);assert.equal(costs[0].usd,15);assert.equal(costs[0].lines,2);
  delete state.costSettings.preparationRateUsd;assert.equal(D.warehouseCosts(state)[0].usd,null);
});
test('tomorrow reminders have stable identities and correct due date',()=>{
  const rows=D.reminderSeeds();assert.equal(rows.length,2);assert.equal(new Set(rows.map(row=>row.id)).size,2);assert.ok(rows.every(row=>row.fecha==='2026-10-23'));
});

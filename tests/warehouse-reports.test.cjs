const {test}=require('node:test'),assert=require('node:assert/strict'),D=require('../warehouse-report-domain.js');
const request=extra=>({codigo:'C1',cliente:'Hotel Costa',solicitud:'S001',pedidoCliente:'001',lineaPedidoCliente:'10',articulo:'A001',cantidadSolicitada:10,cantidadCompletada:0,fechaEnvio:'01/01/2020',...extra});
const order=extra=>({clienteId:'C1',pedidoCliente:'001',lineaPedidoCliente:'10',articulo:'A001',fechaPlanificada:'05/11/2026',camionAsignado:'CAMION 1',rutaNombre:'NORTE',...extra});
test('delivery uses exact client/order/position/material, never the input date or another customer order',()=>{
 const lines=[order({pedidoCliente:'OTHER',fechaPlanificada:'01/11/2026'}),order({lineaPedidoCliente:'20',fechaPlanificada:'02/11/2026'}),order({articulo:'OTHER',fechaPlanificada:'03/11/2026'}),order({clienteId:'C2',fechaPlanificada:'04/11/2026'}),order({})];assert.deepEqual(D.resolveSchedule(request({}),lines).dates,['05/11/2026']);assert.equal(D.resolveSchedule(request({pedidoCliente:'UNKNOWN'}),lines).label,'Sin programar en calendario');
});
test('split deliveries preserve all dates and trucks, excluding archived residues',()=>{
 const result=D.resolveSchedule(request({}),[order({fechaPlanificada:'04/11/2026',archivedAt:'old'}),order({}),order({fechaPlanificada:'2026-11-06',camionAsignado:'CAMION 2'}),order({})]);assert.deepEqual(result.dates,['05/11/2026','06/11/2026']);assert.equal(result.entries.length,2);assert.equal(result.status,'Varias entregas');
});
test('partially unassigned splits are visible, without inventing a date',()=>{
 const result=D.resolveSchedule(request({}),[order({}),order({fechaPlanificada:''})]);assert.equal(result.status,'Programación parcial');assert.equal(result.unplanned,true);assert.match(result.label,/Parte sin programar/);assert.deepEqual(D.resolveSchedule(request({}),[]).dates,[]);
});
test('plan and control join once; planned quantity and latest explicit zero are preserved',()=>{
 const context={plan:[request({cantidadCompletada:8})],control:[request({cantidadSolicitada:7,cantidadCompletada:0})],lineItems:[order({})]};const report=D.build(context);assert.equal(report.lines.length,1);assert.equal(report.lines[0].cantidadSolicitada,10);assert.equal(report.lines[0].cantidadCompletada,0);assert.equal(report.lines[0].cantidadPendiente,10);assert.equal(report.metrics.requests,1);assert.equal(report.metrics.pending,1);
});
test('request counts are unique across materials, and complete only when all lines are prepared',()=>{
 const context={plan:[request({}),request({articulo:'A002'}),request({solicitud:'S002'})],control:[request({cantidadCompletada:10})],lineItems:[order({})]};const report=D.build(context);assert.equal(report.metrics.requests,2);assert.equal(report.metrics.lines,3);assert.equal(report.metrics.completed,1);assert.equal(report.metrics.completedRequests,0);assert.equal(report.metrics.unscheduled,1);assert.equal(report.metrics.withoutReason,2);assert.equal(report.metrics.completion,1/3);
});
test('selected view and search consistently filter detail, agenda and metrics',()=>{
 const context={plan:[request({}),request({solicitud:'S002',cantidadCompletada:10}),request({codigo:'C2',cliente:'Otro cliente'})],lineItems:[order({})],mode:'pendientes',search:'Hotel Costa'};const report=D.build(context);assert.equal(report.lines.length,1);assert.equal(report.metrics.requests,1);assert.equal(report.agenda[0].requests,1);assert.equal(D.build({...context,mode:'completadas'}).lines.length,1);
});
test('calendar rescheduling is reflected on each build without mutating requests',()=>{
 const context={plan:[request({})],lineItems:[order({})]},before=JSON.stringify(context.plan);assert.deepEqual(D.build(context).lines[0].schedule.dates,['05/11/2026']);context.lineItems[0].fechaPlanificada='20/11/2026';assert.deepEqual(D.build(context).lines[0].schedule.dates,['20/11/2026']);assert.equal(JSON.stringify(context.plan),before);assert.ok(!JSON.stringify(D.build(context).lines[0].schedule).includes('2020'));
});
test('calendar dates validate and convert ISO or day-first input',()=>{assert.equal(D.iso('2026-11-05'),'2026-11-05');assert.equal(D.iso('05/11/2026'),'2026-11-05');assert.equal(D.iso('31/02/2026'),'');assert.equal(D.iso(''),'');});
test('an older control cannot add unrelated requests to the current planning report',()=>{
 const context={plan:[request({})],control:[request({solicitud:'OLD',cantidadCompletada:10})],lineItems:[order({})]};assert.deepEqual(D.build(context).lines.map(row=>row.solicitud),['S001']);assert.deepEqual(D.build({...context,mode:'completadas'}).lines.map(row=>row.solicitud),['OLD']);
});

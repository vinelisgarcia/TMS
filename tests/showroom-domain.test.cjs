const {test}=require('node:test'),assert=require('node:assert/strict');
const D=require('../showroom-domain.js'),C=require('../showroom-counts-domain.js');
const physical=rows=>D.consolidate([['Responsable','ID ','Ubicacion Showroom','','Cantidad Contada','DIFERENCIAS'],...rows],{source:'ct',asOf:'2026-10-09'});
const virtual=rows=>D.consolidate([['ID de producto','Cantidad','Valor calculado','Valor calculado (Moneda)','Cantidad (Unidad)'],...rows],{source:'ad',asOf:'2026-11-08'});
test('Conteo mapping sums repeated IDs and preserves locations, operator and source notes',()=>{
 const s=physical([['Operador','0001','CUBO','BUFFET',2,'Revisar'],['Operador','0001','PARED','',3,'']]);
 assert.equal(s.valid,true);assert.equal(s.materials[0].code,'0001');assert.equal(s.materials[0].quantity,5);
 assert.equal(s.materials[0].details[0].responsible,'Operador');assert.equal(s.materials[0].details[0].area,'BUFFET');assert.equal(s.materials[0].details[0].sourceNote,'Revisar');
 assert.equal(s.materials[0].details[1].location,'PARED');
});
test('physical minus virtual identifies actual surplus and shortage with virtual weighted cost',()=>{
 const p=physical([['O','A','C',null,12],['O','B','C',null,3]]),v=virtual([['A',4,40,'DOP','UN'],['A',6,120,'DOP','UN'],['B',5,100,'DOP','UN']]);
 const r=D.compare(p,v);assert.equal(r.rows.find(x=>x.code==='A').difference,2);assert.equal(r.rows.find(x=>x.code==='A').unitCostDop,16);assert.equal(r.rows.find(x=>x.code==='A').signedImpactDop,32);assert.equal(r.rows.find(x=>x.code==='B').signedImpactDop,-40);assert.equal(r.summary.netImpactDop,-8);assert.ok(r.issues.some(x=>x.includes('fechas')));
});
test('missing physical count is unknown; explicit zero is a real shortage',()=>{
 const p=physical([['O','ZERO','C','',0]]),v=virtual([['ZERO',5,100,'DOP','UN'],['UNKNOWN',7,70,'DOP','UN']]);const r=D.compare(p,v);
 assert.equal(r.summary.notCounted,1);assert.equal(r.summary.differences,1);assert.equal(r.summary.grossImpactDop,100);
 const unknown=r.rows.find(x=>x.code==='UNKNOWN');assert.equal(unknown.ctQty,null);assert.equal(unknown.difference,null);assert.equal(unknown.impactDop,null);assert.equal(unknown.status,'not_counted');assert.equal(r.rows.find(x=>x.code==='ZERO').difference,-5);
});
test('physical-only articles remain unvalued and appear for guided verification',()=>{
 const r=D.compare(physical([['O','NEW','C','',3]]),virtual([['A',1,10,'DOP','UN']]));const row=r.rows.find(x=>x.code==='NEW');assert.equal(row.difference,3);assert.equal(row.impactDop,null);assert.equal(C.priorities(r).unvalued[0].code,'NEW');
});
test('negative physical rows reject the import even when another row hides the negative total',()=>{
 const s=physical([['O','A','C','',-1],['O','A','C','',4]]);assert.equal(s.valid,false);assert.equal(s.invalidRows[0].row,2);
});
test('operator and location changes are included in content identity',()=>{
 assert.notEqual(physical([['Alice','A','C','',1]]).semanticKey,physical([['Bob','A','C','',1]]).semanticKey);
});
test('showroom guided route keeps location order independently of CT article ordering',()=>{
 const state={sources:{ct:physical([['O','A','Z-01','',2],['O','B','A-01','',2]]),ad:virtual([['A',1,10,'DOP','UN'],['B',1,10,'DOP','UN']])}};
 const cycle=C.createCycle(state,D.compare(state.sources.ct,state.sources.ad));
 assert.deepEqual(C.tasks(cycle).map(row=>row.code),['B','A']);
});
test('showroom count cycles reuse safely, preserve source snapshots and do not alter CT data',()=>{
 const state={sources:{ct:physical([['O','A','CUBO','',2]]),ad:virtual([['A',1,10,'DOP','UN']])}};
 state.sources.ct.semanticHash='PHYSICAL';state.sources.ad.semanticHash='VIRTUAL';
 const r=D.compare(state.sources.ct,state.sources.ad),cycle=C.createCycle(state,r,{id:'cycle'});assert.equal(cycle.differenceConvention,'physical_virtual');
 assert.equal(C.createCycle(state,r).id,'cycle');assert.equal(C.tasks(cycle)[0].expectedLocation,'CUBO');
 C.recordCount(state,{code:'A',expectedLocation:'CUBO',actualLocation:'CUBO',confirmed:true,quantity:2});C.processItem(state,'A');
 assert.equal(C.itemReport(cycle,'A').referenceSource,'Físico');C.closeCycle(state);assert.equal(cycle.status,'closed');assert.equal(state.sources.ct.materials[0].quantity,2);
});

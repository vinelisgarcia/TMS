const {test}=require('node:test'),assert=require('node:assert/strict');
const C=require('../inventory-counts-domain.js'),D=require('../inventory-domain.js');
const ct=rows=>D.consolidate([['Material','Description','Quantity','Unit of measurement','Location'],...rows],{source:'ct'});
const ad=rows=>D.consolidate([['Codigo','Descripcion','Stock','Valor total DOP','Unidad'],...rows],{source:'ad'});
function fixture(){const data={sources:{ct:ct([['A','A',5,'Piece','A01'],['A','A',5,'Piece','A02'],['B','B',2,'Piece','B01'],['C','C',4,'Piece','C01']]),ad:ad([['A','A',5,50,'UND'],['B','B',3,60,'UND']])},cases:{},settings:{}};data.sources.ct.semanticHash='CT1';data.sources.ad.semanticHash='AD1';return data;}
const report=data=>D.compare(data.sources.ct,data.sources.ad);
const record=(data,code,bin,quantity,extra={})=>C.recordCount(data,{code,expectedLocation:bin,actualLocation:bin,quantity,confirmed:true,...extra});

test('top positive 10 and negative 20 select signed financial impact, never absolute cross-sign rank',()=>{
  const rows=[];for(let i=1;i<=40;i++){rows.push({code:'P'+i,difference:i,signedImpactDop:i*10,impactDop:i*10});rows.push({code:'N'+i,difference:-i,signedImpactDop:-i*10,impactDop:i*10});}
  rows.push({code:'CT_ONLY',status:'only_ct',difference:-9999,signedImpactDop:null,impactDop:null});
  const ranks=C.priorities({ready:true,rows});assert.equal(ranks.positive.length,10);assert.equal(ranks.negative.length,20);assert.equal(ranks.positive[0].code,'P40');assert.equal(ranks.positive.at(-1).code,'P31');assert.equal(ranks.negative[0].code,'N40');assert.equal(ranks.negative.at(-1).code,'N21');assert.equal(ranks.mandatory.length,30);assert.equal(ranks.unvalued[0].code,'CT_ONLY');
});
test('source location is recognized separately from the warehouse and does not multiply stock',()=>{
  const data=fixture();const item=C.catalog(data).find(row=>row.code==='A');assert.equal(item.stockCT,10);assert.deepEqual(item.locations.map(row=>row.location),['A01','A02']);
  const noBins=D.consolidate([['Material','Quantity','Warehouse'],['Z',2,'1015']],{source:'ct'});assert.equal(C.catalog({sources:{ct:noBins}})[0].locations.length,0);
});
test('guided cycle takes both rankings once, preserves source references, and supports mandatory CT-only selection',()=>{
  const data=fixture(),cycle=C.createCycle(data,report(data),{id:'cycle'});assert.deepEqual(Object.keys(cycle.items).sort(),['A','B']);assert.ok(cycle.items.A.required);assert.equal(cycle.items.A.expectedLocations.length,2);assert.equal(C.createCycle(data,report(data)).id,'cycle');
  C.addItem(data,'C',{required:true,reason:'Solo CT'});assert.equal(cycle.items.C.adRegistered,false);assert.equal(cycle.items.C.stockAD,null);assert.equal(cycle.items.C.required,true);
  data.sources.ct.materials.find(row=>row.code==='A').quantity=999;assert.equal(cycle.items.A.stockCT,10);
});
test('zero is a valid count; empty, negative and unconfirmed locations cannot complete a task',()=>{
  const data=fixture();C.createCycle(data,report(data));assert.equal(C.itemReport(C.active(data),'B').quantity,null);
  assert.throws(()=>record(data,'B','B01',-1));assert.throws(()=>record(data,'B','B01',NaN));assert.throws(()=>record(data,'B','B01',0,{confirmed:false}));record(data,'B','B01',0);assert.equal(C.itemReport(C.active(data),'B').complete,true);assert.equal(C.itemReport(C.active(data),'B').quantity,0);
});
test('all locations must be covered before global differences and processing are available',()=>{
  const data=fixture();C.createCycle(data,report(data));record(data,'A','A01',3);let row=C.itemReport(C.active(data),'A');assert.equal(row.ctDifference,null);assert.equal(row.complete,false);assert.throws(()=>C.processItem(data,'A'));
  record(data,'A','A02',7);row=C.itemReport(C.active(data),'A');assert.equal(row.quantity,10);assert.equal(row.ctDifference,0);assert.equal(row.adDifference,5);C.processItem(data,'A');assert.equal(C.itemReport(C.active(data),'A').processed,true);
});
test('a corrected location covers its original slot and becomes the next-cycle assignment with audit evidence',()=>{
  const data=fixture();C.createCycle(data,report(data));record(data,'A','A01',5,{actualLocation:'REAL-01',note:'Estante movido'});record(data,'A','A02',5);
  const row=C.itemReport(C.active(data),'A');assert.equal(row.complete,true);assert.equal(row.records[0].locationCorrected,true);assert.equal(C.active(data).audit[0].from,'A01');assert.equal(C.active(data).audit[0].to,'REAL-01');assert.deepEqual(C.catalog(data).find(item=>item.code==='A').locations.map(bin=>bin.location),['A02','REAL-01']);
});
test('same item and real location require explicit additional-box confirmation before sums',()=>{
  const data=fixture();C.createCycle(data,report(data));record(data,'B','B01',1);assert.throws(()=>record(data,'B','B01',2),/otra caja/);record(data,'B','B01',2,{additional:true});assert.equal(C.itemReport(C.active(data),'B').quantity,3);
});
test('quantity corrections and voids require reasons, keep audit and reopen processing',()=>{
  const data=fixture();C.createCycle(data,report(data));const saved=record(data,'B','B01',2);C.processItem(data,'B');assert.throws(()=>C.correctRecord(data,saved.id,3,''));C.correctRecord(data,saved.id,3,'Reconteo');assert.equal(C.itemReport(C.active(data),'B').quantity,3);assert.equal(C.itemReport(C.active(data),'B').processed,false);C.processItem(data,'B');C.correctRecord(data,saved.id,null,'Caja duplicada');assert.equal(C.itemReport(C.active(data),'B').complete,false);assert.equal(C.active(data).audit.length,2);
});
test('mandatory pending items block cycle closure; closed cycles reject further records',()=>{
  const data=fixture();C.createCycle(data,report(data));assert.throws(()=>C.closeCycle(data),/obligatorios/);record(data,'A','A01',5);record(data,'A','A02',5);record(data,'B','B01',2);assert.throws(()=>C.closeCycle(data),/obligatorios/);C.processItem(data,'A');C.processItem(data,'B');C.closeCycle(data);assert.equal(C.active(data).status,'closed');assert.throws(()=>record(data,'B','B01',2));
});
test('original template imports barcodes/master and sums repeated location rows, preserving leading zeros',()=>{
  const data={sources:{},cases:{},settings:{}};
  C.importMaster(data,[['CODIGO','DESCRIPCION','CODIGO BARRAS','CENTRAL','COSTO PROMEDIO'],['0001','Artículo','750000000001',24,2]],{code:0,description:1,barcode:2,stock:3,cost:4},'articles',{parseNumber:D.parseNumber});
  C.importMaster(data,[['CODIGO','UBICACION','STOCK UBICACION'],['0001','A01',10],['0001','A01',2],['0001','A02',12]],{code:0,location:1,stock:2},'positions',{parseNumber:D.parseNumber});
  const item=C.resolve(data,'750000000001');assert.equal(item.code,'0001');assert.equal(item.locations[0].stock,12);assert.equal(item.masterStock,24);C.createCycle(data,{ready:false});C.addItem(data,'0001');record(data,'0001','A01',12);record(data,'0001','A02',12);const row=C.itemReport(C.active(data),'0001');assert.equal(row.referenceStock,24);assert.equal(row.referenceDifference,0);
});
test('unknown location requires actual location entry and cannot borrow the warehouse ID',()=>{
  const data={sources:{ct:D.consolidate([['Material','Quantity','Warehouse'],['A',1,'1015']],{source:'ct'})}};C.createCycle(data,{ready:false});C.addItem(data,'A');assert.throws(()=>record(data,'A','',1),/ubicación real/);record(data,'A','',1,{actualLocation:'RACK-3'});assert.equal(C.itemReport(C.active(data),'A').complete,true);
});
test('only physically complete and processed CT-only items enter the AD-registration output',()=>{
  const data=fixture();C.createCycle(data,report(data));C.addItem(data,'C');record(data,'C','C01',4);assert.equal(C.itemReport(C.active(data),'C').readyForAD,false);C.processItem(data,'C');assert.equal(C.itemReport(C.active(data),'C').readyForAD,true);assert.equal(data.sources.ad.materials.some(row=>row.code==='C'),false);
});
test('imports are atomic on error and location replacement does not remove other materials',()=>{
  const data=fixture();C.importMaster(data,[['C','L'],['A','AA'],['B','BB']],{code:0,location:1},'positions');C.importMaster(data,[['C','L'],['A','NEW']],{code:0,location:1},'positions');assert.equal(C.catalog(data).find(row=>row.code==='B').locations[0].location,'BB');
  const prior=JSON.stringify(data);assert.throws(()=>C.importMaster(data,[['C','L'],['A','BAD'],['B','']],{code:0,location:1},'positions'));assert.equal(JSON.stringify(data),prior);
});
test('new cuts archive prior cycles instead of retroactively changing theoretical or location references',()=>{
  const data=fixture();const old=C.createCycle(data,report(data),{id:'old'});record(data,'B','B01',2);data.sources.ct.semanticHash='CT2';assert.throws(()=>C.addItem(data,'C'),/otro corte/);C.createCycle(data,report(data),{id:'new'});assert.equal(old.status,'archived');assert.equal(old.records.length,1);assert.equal(C.active(data).id,'new');assert.equal(C.active(data).records.length,0);
});

test('position-only imports create countable articles without invented system stock',()=>{
  const data={sources:{}};C.importMaster(data,[['CODIGO','UBICACION'],['BIN_ONLY','A01']],{code:0,location:1},'positions');
  C.createCycle(data,{ready:false});C.addItem(data,'BIN_ONLY');record(data,'BIN_ONLY','A01',2);
  const row=C.itemReport(C.active(data),'BIN_ONLY');assert.equal(row.quantity,2);assert.equal(row.referenceStock,null);assert.equal(row.ctDifference,null);
});
test('decimal boxes sum accurately and repeated location corrections replace the previous assignment',()=>{
  const data=fixture();C.createCycle(data,report(data));record(data,'B','B01',0.1,{actualLocation:'REAL-1'});record(data,'B','B01',0.2,{actualLocation:'REAL-2'});
  assert.equal(C.itemReport(C.active(data),'B').quantity,0.3);assert.deepEqual(C.catalog(data).find(row=>row.code==='B').locations.map(bin=>bin.location),['REAL-2']);
});
test('ambiguous barcodes and incompatible units cannot silently select a material',()=>{
  const data=fixture();C.importMaster(data,[['CODIGO','BARRAS'],['A','SAME'],['B','SAME']],{code:0,barcode:1},'articles');assert.throws(()=>C.resolve(data,'SAME'),/varios artículos/);
  data.sources.ct.materials.find(row=>row.code==='C').unitConflict=true;C.createCycle(data,report(data));assert.throws(()=>C.addItem(data,'C'),/incompatibles/);
});

test('AD minus CT reverses ranked membership and old convention cycles remain archived with records',()=>{
 const data=fixture(),r=report(data),ranks=C.priorities(r);
 assert.equal(ranks.positive[0].code,'B');assert.equal(ranks.negative[0].code,'A');assert.equal(ranks.unvalued[0].code,'C');
 const old=C.createCycle(data,r,{id:'old'});record(data,'B','B01',2);delete old.differenceConvention;
 const fresh=C.createCycle(data,r,{id:'new'});assert.equal(old.status,'archived');assert.equal(old.records.length,1);assert.equal(fresh.differenceConvention,'ad_ct');
});

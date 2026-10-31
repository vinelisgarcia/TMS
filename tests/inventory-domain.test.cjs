// Synthetic fixtures only. Optional user-sample test reads local files but never
// stores their rows or publishes private inventory data in this repository.
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const D=require('../inventory-domain.js');
const physical=(rows,extra={})=>D.consolidate([['Material','Description','Quantity','Availability','Unit of measurement','Warehouse','Batchnumber'],...rows],{source:'ct',decimal:'comma',...extra});
const accounting=(rows,extra={})=>D.consolidate([['Código','Descripción','Stock','Valor total DOP','Unidad','Almacén'],...rows],{source:'ad',decimal:'comma',...extra});
const C=(code,quantity)=>[code,'Material '+code,quantity,quantity,'Piece','1015','BATCH'];
const A=(code,quantity,value,unit='UND')=>[code,'Material '+code,quantity,value,unit,'CT'];

test('CT offset header, physical stock, availability and date metadata are detected',()=>{
  const matrix=D.parseCsv('Detailed stock for AD;;;;\r\nDate: 30.10.2026;;;;\r\n;;;;\r\nMaterial;Description;Quantity;Availability;Unit of measurement\r\n00123;Plato;10;7;Piece\r\n');
  const layout=D.detectLayout(matrix,'ct');
  assert.equal(layout.headerRow,3);assert.equal(layout.columns.quantity,2);assert.equal(layout.columns.availability,3);assert.equal(layout.asOf,'2026-10-30');
  const snapshot=D.consolidate(matrix,{source:'ct',asOf:layout.asOf});
  assert.equal(snapshot.materials[0].code,'00123');assert.equal(snapshot.materials[0].quantity,10);assert.equal(snapshot.materials[0].availability,7);
});

test('CSV supports escaped quotes, delimiters inside fields, embedded newlines and CRLF',()=>{
  assert.deepEqual(D.parseCsv('Material,Description,Quantity\r\nA,"Plato, ""grande""\nsegunda línea",2\r\n'),[['Material','Description','Quantity'],['A','Plato, "grande"\nsegunda línea','2']]);
  assert.deepEqual(D.parseCsv('Material\tQuantity\nA\t0'),[['Material','Quantity'],['A','0']]);
  assert.throws(()=>D.parseCsv('Material;Quantity\nA;"2'),/comillas/);
});

test('CSV delimiter detection ignores comma characters inside quoted descriptions',()=>{
  const description=Array.from({length:40},(_,i)=>'palabra'+i).join(',');
  const matrix=D.parseCsv('Material;Description;Quantity\nA;"'+description+'";2');
  assert.equal(matrix[0].length,3);assert.equal(matrix[1][1],description);assert.equal(matrix[1][2],'2');
});

test('numbers keep blank distinct from zero and validate explicit Spanish/English formats',()=>{
  assert.equal(D.parseNumber(''),null);assert.equal(D.parseNumber(0),0);assert.equal(D.parseNumber('0,00','comma'),0);
  assert.equal(D.parseNumber('1.234,56','comma'),1234.56);assert.equal(D.parseNumber('1,234.56','dot'),1234.56);
  assert.equal(D.parseNumber('1 234,56','comma'),1234.56);assert.equal(D.parseNumber('(1234,56)','comma'),-1234.56);
  assert.throws(()=>D.parseNumber('1,234'),/ambiguo/);assert.throws(()=>D.parseNumber('1.234'),/ambiguo/);
  assert.throws(()=>D.parseNumber('12,34.56','dot'));assert.throws(()=>D.parseNumber('2 unidades'));assert.throws(()=>D.parseNumber(Infinity));
});

test('DOP accounting negatives support currency symbols inside parentheses',()=>{
  assert.equal(D.parseNumber('(RD$ 1.234,56)','comma'),-1234.56);
});

test('AD unit valuation is derived from summed total DOP divided by summed AD quantity',()=>{
  const ad=accounting([A('A',10,100),A('A',20,400)]),ct=physical([C('A',20)]);
  assert.equal(ad.materials.length,1);assert.equal(ad.materials[0].quantity,30);assert.equal(ad.materials[0].totalValueDop,500);assert.equal(ad.materials[0].rowCount,2);
  const row=D.compare(ct,ad).rows[0];assert.equal(row.unitCostDop,500/30);assert.equal(row.difference,10);
  assert.equal(row.impactDop,166.67);assert.equal(row.signedImpactDop,166.67);
});

test('batch/material repeats consolidate, including exact repeated rows which are flagged',()=>{
  const snapshot=physical([C('A',2),C('A',2),['A','Material A',3,3,'Piece','1015','SECOND']]);
  assert.equal(snapshot.materials.length,1);assert.equal(snapshot.materials[0].quantity,7);assert.equal(snapshot.materials[0].rowCount,3);
  assert.equal(snapshot.duplicateExtraRows,1);assert.equal(snapshot.repeatedMaterials,1);assert.deepEqual(snapshot.materials[0].batches,['BATCH','SECOND']);
  assert.ok(snapshot.warnings.some(message=>message.includes('idénticas')));
});

test('physical Quantity is default and explicit Availability is a separately declared basis',()=>{
  const row=['A','Material A',10,7,'Piece','1015','BATCH'];
  const total=physical([row]),available=physical([row],{quantityBasis:'availability'});
  assert.equal(total.totals.quantity,10);assert.equal(available.totals.quantity,7);assert.equal(available.quantityBasis,'availability');assert.notEqual(total.semanticKey,available.semanticKey);
});

test('CT F.O.B and mixed currencies never substitute for AD DOP total valuation',()=>{
  const source=D.consolidate([['Material','Quantity','F.O.B','Currency'],['A',10,9999,'USD']],{source:'ct'});
  assert.equal(source.materials[0].totalValueDop,null);assert.equal(source.totals.totalValueDop,null);
  assert.equal(D.detectLayout([['Código','Stock','Costo promedio'],['A',10,10]],'ad').columns.totalValue,undefined);
});

test('partial AD value remains unknown while known source value is retained for review',()=>{
  const ad=accounting([A('A',10,100),A('A',20,''),A('B',2,0)]);
  assert.equal(ad.materials[0].totalValueDop,null);assert.equal(ad.totals.totalValueDop,null);assert.equal(ad.totals.knownValueDop,100);assert.equal(ad.totals.missingValueMaterials,1);
  const result=D.compare(physical([C('A',20),C('B',2)]),ad);
  assert.equal(result.summary.unvalued,1);assert.equal(result.rows.find(row=>row.code==='B').impactDop,0);
});

test('CT-only items lack a valuation; AD-only items use the full missing stock value',()=>{
  const result=D.compare(physical([C('ONLY-CT',5)]),accounting([A('ONLY-AD',3,300)]));
  const ct=result.rows.find(row=>row.code==='ONLY-CT'),ad=result.rows.find(row=>row.code==='ONLY-AD');
  assert.equal(ct.status,'only_ct');assert.equal(ct.difference,-5);assert.equal(ct.unitCostDop,null);assert.equal(ct.impactDop,null);
  assert.equal(ad.status,'only_ad');assert.equal(ad.difference,3);assert.equal(ad.impactDop,300);
});

test('missing entire source never presents absent data as zero stock',()=>{
  assert.equal(D.compare(null,accounting([A('A',1,1)])).ready,false);assert.deepEqual(D.compare(null,null).rows,[]);
});

test('zero and negative AD stock never invent a unit cost or divide by zero',()=>{
  for(const [quantity,value] of [[0,0],[0,100],[-2,-20],[-2,20]]) {
    const row=D.compare(physical([C('A',1)]),accounting([A('A',quantity,value)])).rows[0];
    assert.equal(row.unitCostDop,null);assert.equal(row.impactDop,null);assert.ok(row.issues.length);
  }
});

test('incompatible units are isolated instead of generating false quantity and monetary differences',()=>{
  const mixed=physical([C('A',3),['A','Material A',4,4,'KG','1015','BATCH']]);
  assert.equal(mixed.materials[0].quantity,null);assert.equal(mixed.materials[0].unitConflict,true);assert.equal(mixed.totals.quantity,null);
  assert.deepEqual(mixed.materials[0].quantitiesByUnit,{UND:3,KG:4});
  const row=D.compare(mixed,accounting([A('A',3,30)])).rows[0];
  assert.equal(row.status,'unit_conflict');assert.equal(row.difference,null);assert.equal(row.impactDop,null);
  assert.equal(D.compare(physical([C('A',3)]),accounting([A('A',3,30,'KG')])).rows[0].status,'unit_conflict');
});

test('availability cannot be summed across incompatible units for the same material',()=>{
  const mixed=physical([C('A',3),['A','Material A',4,4,'KG','1015','BATCH']]);
  assert.equal(mixed.materials[0].availability,null);assert.equal(mixed.totals.availability,null);
});

test('unit synonyms match but leading zero material references stay distinct',()=>{
  const result=D.compare(physical([C('001',2),C('1',3)]),accounting([A('001',2,20,'Unidades'),A('1',3,30,'PCS')]));
  assert.equal(result.summary.materials,2);assert.equal(result.summary.equal,2);assert.equal(D.normalizeCode(' 001 '),'001');
});

test('invalid numeric rows mark source invalid and must block comparison of incomplete totals',()=>{
  const source=physical([C('A',2),C('B',''),C('C','invalid')]);
  assert.equal(source.acceptedRowCount,1);assert.equal(source.invalidRows.length,2);assert.equal(source.valid,false);
  assert.equal(D.compare(source,accounting([A('A',2,20)])).ready,false);
});

test('named total rows are not added again and selected warehouse scope is honored',()=>{
  const matrix=[['Código','Descripción','Stock','Valor total','Almacén'],['A','Material A',10,100,'CT'],['A','Material A',20,200,'CENTRAL'],['TOTAL','',30,300,'']];
  const source=D.consolidate(matrix,{source:'ad',warehouses:['CT']});
  assert.equal(source.acceptedRowCount,1);assert.equal(source.filteredRows,1);assert.equal(source.totals.quantity,10);assert.equal(source.totals.totalValueDop,100);
});

test('semantic identity ignores filename/row order but retains duplicate multiplicity and cutoff',()=>{
  const a=physical([C('A',2),C('B',3)],{fileName:'first.csv',asOf:'2026-10-30'});
  const b=physical([C('B',3),C('A',2)],{fileName:'second.xlsx',asOf:'2026-10-30'});
  assert.equal(a.semanticKey,b.semanticKey);
  assert.notEqual(a.semanticKey,physical([C('A',2),C('A',2),C('B',3)],{asOf:'2026-10-30'}).semanticKey);
  assert.notEqual(a.semanticKey,physical([C('A',2),C('B',3)],{asOf:'2026-10-31'}).semanticKey);
});

test('priority ranks financial exposure and gross impact does not net deficits against surpluses',()=>{
  const result=D.compare(physical([C('A',101),C('B',8),C('C',1000)]),accounting([A('A',100,100),A('B',10,10000),A('C',100,100)]));
  assert.equal(result.rows[0].code,'B');assert.equal(result.rows[0].impactDop,2000);assert.equal(result.summary.grossImpactDop,2901);assert.equal(result.summary.netImpactDop,1099);
});

test('different cutoff dates remain visible as a comparison issue',()=>{
  const result=D.compare(physical([C('A',2)],{asOf:'2026-10-30'}),accounting([A('A',2,20)],{asOf:'2026-10-31'}));
  assert.ok(result.issues.some(message=>message.includes('fechas de corte')));
});

test('local CT alternate exports, when present, yield one identical consolidated snapshot',t=>{
  const base='/Users/vinelis/Downloads/';
  const utf=base+'total-stock-report-2026-10-30 160119(Stock report)-2.csv',win=base+'total-stock-report-2026-10-30 160119(Stock report).csv';
  if(!fs.existsSync(utf)||!fs.existsSync(win)){t.skip('User samples are deliberately not committed.');return;}
  const a=D.consolidate(D.parseCsv(new TextDecoder('utf-8',{fatal:true}).decode(fs.readFileSync(utf))),{source:'ct',decimal:'comma'});
  const b=D.consolidate(D.parseCsv(new TextDecoder('windows-1252').decode(fs.readFileSync(win))),{source:'ct',decimal:'comma'});
  assert.equal(a.rawRowCount,2928);assert.equal(a.materials.length,2249);assert.equal(a.totals.quantity,386962);assert.equal(a.repeatedMaterials,496);
  assert.equal(a.semanticKey,b.semanticKey);assert.equal(a.duplicateExtraRows,1);assert.equal(a.totals.availability,384415);
});

 test('monetary summaries reconcile to the per-material cent amounts without rounding unit costs',()=>{
 const result=D.compare(physical([C('A',2),C('B',2),C('C',2)]),accounting([A('A',3,1),A('B',3,1),A('C',3,1)]));
 assert.equal(result.rows[0].unitCostDop,1/3);assert.equal(result.summary.grossImpactDop,0.99);assert.equal(result.summary.netImpactDop,0.99);
 });

const sapStock=[
 [null,'Resumen de stock valorado'],[],[null,'Última actualización:',46326.47296296296,'Huso horario:','UTC-4'],[],
 [null,'Tipo de stock','ID de producto','Cantidad','Cantidad (Unidad)','Descripción de producto','Valor calculado','Valor calculado (Moneda)'],
 [null,'En tránsito','A',2,'UN','Artículo A',20,'DOP'],
 [null,'Almacén externo','A',3,'UN','Artículo A',60,'DOP'],
 [null,'Almacén externo','B',1,'UN','Artículo B',100,'DOP']
];
test('actual SAP AD headings and Excel last-update serial are detected without manual mapping',()=>{
 const layout=D.detectLayout(sapStock,'ad');assert.equal(layout.headerRow,4);assert.equal(layout.columns.code,2);assert.equal(layout.columns.quantity,3);assert.equal(layout.columns.totalValue,6);assert.equal(layout.columns.unit,4);assert.equal(layout.columns.stockType,1);assert.equal(layout.columns.currency,7);assert.equal(layout.asOf,'2026-10-31');assert.deepEqual(layout.suggestedStockTypes,['Almacén externo']);
});
test('AD stock-type filter keeps in-transit stock separate from external physical inventory',()=>{
 const source=D.consolidate(sapStock,{source:'ad',stockTypes:['Almacén externo']});assert.equal(source.acceptedRowCount,2);assert.equal(source.filteredRows,1);assert.equal(source.rawRowCount,3);assert.equal(source.totals.quantity,4);assert.equal(source.totals.totalValueDop,160);assert.equal(source.materials[0].quantity,3);assert.equal(source.materials[0].details[0].stockType,'Almacén externo');assert.equal(source.materials[0].unit,'UND');
 const all=D.consolidate(sapStock,{source:'ad'});assert.equal(all.totals.quantity,6);assert.equal(all.materials[0].quantity,5);assert.ok(all.warnings.some(message=>message.includes('En tránsito')));
});
test('foreign or missing currency cannot be quietly treated as DOP when currency column exists',()=>{
 const matrix=structuredClone(sapStock);matrix[6][7]='USD';const result=D.consolidate(matrix,{source:'ad',stockTypes:['Almacén externo']});assert.equal(result.valid,false);assert.match(result.invalidRows[0].message,/DOP/);assert.equal(D.compare(physical([C('A',4)]),result).ready,false);
});

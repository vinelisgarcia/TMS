const {test}=require('node:test');
const assert=require('node:assert/strict');
const vm=require('node:vm');
const fs=require('node:fs');
const source=fs.readFileSync(require('node:path').join(__dirname,'../tms-v2.js'),'utf8');
function runtime(){
  const context={APP:{transitMaterials:[]},parseImportBulkDate:x=>x||'',XLSX:{utils:{sheet_to_json:x=>x}}};
  vm.createContext(context);
  for(const name of ['stripAccents','normKey','text','num','normalizeRow','pickField','nowIso','uniqueTexts','normalizeTransitMaterial','transitColumnAliases','parseTransitRowsFromObjects','parseTransitWorksheet','transitImportKey','consolidateTransitRows','ensureTransitMaterials','upsertTransitMaterials']){
    const start=source.indexOf('  function '+name+'(');assert.ok(start>=0,name);
    vm.runInContext(source.slice(start,source.indexOf('\n  }',start)+4),context);
  }
  return context;
}
const headers=['PEDIDO DOM','Codigo cliente del producto','Descripción de Producto','Referencia proveedor','Packslip','Cantidad solicitada','Proveedor','Factura No.','Referencia embarque'];
const detail=(qty=12,invoice='INV-TEST')=>['PO-TEST','000123','Artículo de prueba','SUP-001','PACK-TEST',qty,'Proveedor de prueba',invoice,'SHIP-TEST'];
test('packing template detects row 9 and uses internal material code, not supplier code or customer assignment',()=>{
  const c=runtime(),matrix=[['Plantilla entrada mercancía'],...Array.from({length:7},()=>[]),headers,detail()];
  const [row]=c.parseTransitWorksheet(matrix,{eta:'2026-11-20'});
  assert.equal(row.codigoMaterial,'000123');assert.equal(row.cantidad,12);assert.equal(row.factura,'INV-TEST');
  assert.equal(row.packingList,'PACK-TEST');assert.equal(row.referenciaProveedor,'SUP-001');
  assert.equal(row.pedidoOrigen,'PO-TEST');assert.equal(row.referenciaEmbarque,'SHIP-TEST');
  assert.equal(row.cliente,'');assert.equal(row.eta,'2026-11-20');
});
test('packing without invoice is identified by Packslip',()=>{
  const c=runtime(),[row]=c.parseTransitWorksheet([headers,detail(12,'')],{});
  assert.equal(row.factura,'PACK-TEST');assert.equal(row.tipoDocumento,'packing_list');
});
test('existing invoice headers and form defaults still work',()=>{
  const c=runtime(),[row]=c.parseTransitWorksheet([['Referencia','Descripción','Cantidad'],['0007','Prueba','1,5']],{factura:'F-OLD',eta:'2026-11-21',cliente:'C1'});
  assert.equal(row.codigoMaterial,'0007');assert.equal(row.cantidad,1.5);assert.equal(row.factura,'F-OLD');assert.equal(row.cliente,'C1');
});
test('repeated material lines are summed and reimport replaces the total without multiplying it',()=>{
  const c=runtime(),rows=c.parseTransitWorksheet([headers,detail(12),detail(8)],{});
  c.upsertTransitMaterials(rows);assert.equal(c.APP.transitMaterials.length,1);assert.equal(c.APP.transitMaterials[0].cantidad,20);
  const id=c.APP.transitMaterials[0].id;
  c.upsertTransitMaterials(rows);assert.equal(c.APP.transitMaterials.length,1);assert.equal(c.APP.transitMaterials[0].cantidad,20);assert.equal(c.APP.transitMaterials[0].id,id);
  c.upsertTransitMaterials(c.parseTransitWorksheet([headers,detail(7)],{}));assert.equal(c.APP.transitMaterials[0].cantidad,7);
});
test('distinct invoices retain punctuation and distinct materials remain separate',()=>{
  const c=runtime();c.upsertTransitMaterials(c.parseTransitWorksheet([headers,detail(12,'INV-1'),detail(8,'INV1')],{}));
  assert.equal(c.APP.transitMaterials.length,2);
});
test('multiple packslips for a material preserve source identifiers when consolidated',()=>{
  const c=runtime(),second=detail(8);second[4]='PACK-OTHER';second[8]='SHIP-OTHER';
  c.upsertTransitMaterials(c.parseTransitWorksheet([headers,detail(12),second],{}));
  const [row]=c.APP.transitMaterials;assert.equal(row.cantidad,20);assert.equal(row.packingList,'PACK-TEST / PACK-OTHER');assert.equal(row.referenciaEmbarque,'SHIP-TEST / SHIP-OTHER');
});
test('invalid quantities or missing material reject the whole detail before updating state',()=>{
  for(const value of ['sin cantidad',-2,'12 unidades','1.2.3']){
    const c=runtime();assert.throws(()=>c.parseTransitWorksheet([headers,detail(12),detail(value)],{}),/Revisa material y cantidad/);assert.equal(c.APP.transitMaterials.length,0);
  }
  const c=runtime(),row=detail();row[1]='';assert.throws(()=>c.parseTransitWorksheet([headers,row],{}),/Revisa material/);
});
test('blank and repeated header rows do not become materials, zero quantities are ignored',()=>{
  const c=runtime(),rows=c.parseTransitWorksheet([['Portada'],headers,detail(12),[],headers,detail(0)],{});
  assert.equal(rows.length,1);assert.equal(rows[0].cantidad,12);assert.equal(c.parseTransitWorksheet([['Sin tabla']],{}).length,0);
});

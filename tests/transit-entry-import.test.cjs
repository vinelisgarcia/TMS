// Pure parser regressions. All rows and identifiers are synthetic.
const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const source=fs.readFileSync(path.join(__dirname,'../tms-v2.js'),'utf8');
const context=vm.createContext({CURRENT_YEAR:2026,APP:{transitMaterials:[]},TMSDomain:require('../tms-domain.js')});
for(const name of ['stripAccents','normKey','text','num','nowIso','normalizeRow','pickField','parseImportBulkDate','normalizeTransitMaterial','parseTransitRowsFromObjects','transitMaterialIdentity','transitEntryHeader','parseTransitEntryRows','consolidateTransitEntryRows','parseTransitMaterialsFromText','ensureTransitMaterials','upsertTransitMaterials']){
 const start=source.indexOf('  function '+name+'(');assert.ok(start>=0,name);const end=source.indexOf('\n  }',start)+4;vm.runInContext(source.slice(start,end),context);
}
const headers=['PEDIDO DOM','Codigo cliente del producto','Descripción de Producto','Referencia proveedor','Packslip','Cantidad solicitada','VALIDAR CON EL SIGUIENTE AVISO DE ENTREGA','Proveedor','Factura No.','Referencia embarque'];
const row=(code='0001',qty=2,invoice='INV-1')=>['PO-1',code,'Material de prueba','SUP-99','PACK-1',qty,'','Proveedor',invoice,'SHIP-1'];
const matrix=(rows)=>[[],[],['Plantilla entrada mercancía'],[],[],[],[],[],headers,...rows];
const defaults={factura:'DEFAULT',eta:'2026-11-07',cliente:''},origin={sheet:'Entrada',fileName:'synthetic.xlsx'};
const parse=(rows,d=defaults)=>{const m=matrix(rows);return context.parseTransitEntryRows(m,context.transitEntryHeader(m),d,origin);};
test('reimporting an accounted invoice preserves both receipt and accounting metadata',()=>{
 context.APP.transitMaterials=[{id:'posted',factura:'INV-1',codigoMaterial:'0001',cantidad:2,eta:'2026-11-07',cliente:'',estado:'contabilizada',receivedAt:'2026-11-11',receivedBy:'almacen',accountedAt:'2026-11-12',accountedBy:'contabilidad',statusHistory:[{estado:'contabilizada',at:'2026-11-12',by:'contabilidad'}]}];
 context.upsertTransitMaterials(parse([row('0001',5)]));const saved=context.APP.transitMaterials[0];assert.equal(saved.estado,'contabilizada');assert.equal(saved.receivedAt,'2026-11-11');assert.equal(saved.accountedAt,'2026-11-12');assert.equal(saved.accountedBy,'contabilidad');assert.equal(saved.cantidad,5);assert.equal(saved.statusHistory.length,1);
});
test('detects the real template headers below title/blank rows, never generic invoice or SAP order headers',()=>{
 assert.equal(context.transitEntryHeader(matrix([])).row,8);
 assert.equal(context.transitEntryHeader([['Referencia','Cantidad','Factura']]),null);
 assert.equal(context.transitEntryHeader([['Producto','Cantidad solicitada','Pedido de cliente']]),null);
});
test('maps the AD product identifier, quantity and invoice without assigning supplier reference or purchase order as client',()=>{
 const result=parse([row()]);assert.equal(result[0].codigoMaterial,'0001');assert.equal(result[0].cantidad,2);assert.equal(result[0].factura,'INV-1');assert.equal(result[0].eta,'2026-11-07');assert.equal(result[0].cliente,'');assert.equal(result[0].pedidoCliente,undefined);
 assert.equal(result[0].entradaMercancia.lineas[0].pedidoDom,'PO-1');assert.equal(result[0].entradaMercancia.lineas[0].referenciaProveedor,'SUP-99');assert.equal(result[0].entradaMercancia.lineas[0].fila,10);
});
test('sums repeated rows only within invoice/product/ETA/client and preserves invoice punctuation',()=>{
 const result=context.consolidateTransitEntryRows(parse([row('0001',2,'F-1'),row('0001',3,'F-1'),row('0001',4,'F1')]));
 assert.equal(result.length,2);assert.equal(result[0].cantidad,5);assert.equal(result[0].entradaMercancia.lineas.length,2);assert.equal(result[1].cantidad,4);
 context.APP.transitMaterials=[];context.upsertTransitMaterials(result);assert.equal(context.APP.transitMaterials.length,2);assert.equal(new Set(context.APP.transitMaterials.map(row=>row.id)).size,2);const firstId=context.APP.transitMaterials[0].id;
 context.upsertTransitMaterials(context.consolidateTransitEntryRows(parse([row('0001',2,'F-1'),row('0001',3,'F-1'),row('0001',4,'F1')])));assert.equal(context.APP.transitMaterials.length,2);assert.equal(context.APP.transitMaterials[0].cantidad,5);assert.equal(context.APP.transitMaterials[0].id,firstId);
});
test('preserves unrelated invoices and receipt state when existing materials are updated',()=>{
 context.APP.transitMaterials=[{id:'old',factura:'INV-1',codigoMaterial:'0001',cantidad:9,eta:'2026-11-07',cliente:'',receivedAt:'2026-11-01'},{id:'other',factura:'OTHER',codigoMaterial:'0002',cantidad:8,eta:'2026-11-07',cliente:''}];
 const result=context.upsertTransitMaterials(parse([row()]));assert.equal(result.updated,1);assert.equal(context.APP.transitMaterials.length,2);assert.equal(context.APP.transitMaterials[0].receivedAt,'2026-11-01');assert.equal(context.APP.transitMaterials[1].cantidad,8);
});
test('blank ETA remains unknown; form invoice is used only if the line has none',()=>{
 const result=parse([row('0001',2,'')],{factura:'FALLBACK',eta:'',cliente:'C1'});assert.equal(result[0].eta,'');assert.equal(result[0].factura,'FALLBACK');assert.equal(result[0].cliente,'C1');
 assert.throws(()=>parse([row('0001',2,'')],{factura:'',eta:'',cliente:''}),/Factura No/);
});
test('ignores blank/footer/repeated header and zero movement rows, and blocks malformed data before upsert',()=>{
 const result=parse([row(),[],headers,row('0002',0),['TOTAL','','','','',2]]);assert.equal(result.length,1);
 for(const broken of [row('',2),row('0001','abc'),row('0001','1..2'),row('0001',-1),row('0001',Infinity),row(Number.MAX_SAFE_INTEGER+1,2)])assert.throws(()=>parse([row(),broken]),/fila 11/);
 assert.throws(()=>parse([row()],{...defaults,eta:'2026-02-30'}),/ETA no válida/);
});
test('current invoice objects and text invoices keep working',()=>{
 const result=context.parseTransitRowsFromObjects([{Referencia:'LEGACY-1',Descripcion:'Original',Cantidad:4,Factura:'OLD-INVOICE',ETA:'2026-11-09'}],defaults);
 assert.equal(result[0].codigoMaterial,'LEGACY-1');assert.equal(result[0].cantidad,4);assert.equal(result[0].factura,'OLD-INVOICE');
 assert.equal(context.parseTransitMaterialsFromText('REF-12 Descripcion con espacios 3',defaults)[0].cantidad,3);
});

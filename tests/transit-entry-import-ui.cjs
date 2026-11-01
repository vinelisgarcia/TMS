// Synthetic accounts and imports. Supabase traffic is blocked; no production writes.
const {chromium}=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright');
const assert=require('node:assert/strict'),fs=require('node:fs');
const base=process.env.TMS_TEST_URL||'http://127.0.0.1:8766';
(async()=>{
 const browser=await chromium.launch({headless:true}),errors=[],messages=[];
 try{
  const page=await browser.newPage({viewport:{width:1440,height:1000}});page.on('pageerror',e=>errors.push(e.message));page.on('dialog',d=>{messages.push(d.message());d.accept();});
  await page.route('**/*.supabase.co/**',r=>r.fulfill({status:503,contentType:'application/json',body:'{"message":"Blocked for test"}'}));
  await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>!!window.TMSAuditCore&&!!window.XLSX);
  await page.evaluate(()=>{
   const u={id:'11111111-1111-4111-8111-111111111111',email:'import@example.invalid'},p={userId:u.id,email:u.email,rol:'Admin',activo:true};APP.userProfiles=[p];APP.currentUserProfile=p;mostrarApp(u);window.scheduleAutoSave=()=>{};
   APP.transitMaterials=[{id:'existing',factura:'UNRELATED',codigoMaterial:'OLD-1',descripcion:'Preservar',cantidad:12,eta:'2026-11-07',cliente:''}];cambiarVista('materialesTransito',document.getElementById('nav-materialesTransito'));
  });
  const headers=['PEDIDO DOM','Codigo cliente del producto','Descripción de Producto','Referencia proveedor','Packslip','Cantidad solicitada','VALIDAR CON EL SIGUIENTE AVISO DE ENTREGA','Proveedor','Factura No.','Referencia embarque'];
  const line=(code,qty,invoice)=>['PO-99',code,'Producto de prueba','SUPPLIER-XYZ','PACK-1',qty,'','Proveedor',invoice,'SHIP-1'];
  async function workbook(rows,{cover=true,formatZero=false,legacy=false,bookType='xlsx',eta=false}={}){
   return Buffer.from(await page.evaluate(({rows,headers,cover,formatZero,legacy,bookType,eta})=>{
    const b=XLSX.utils.book_new();if(cover)XLSX.utils.book_append_sheet(b,XLSX.utils.aoa_to_sheet([['Instrucciones de prueba']]),'Portada');
    const matrix=legacy?[headers,...rows]:[[],[],['Plantilla entrada mercancía'],[],[],[],[],[],headers,...rows];
    const s=XLSX.utils.aoa_to_sheet(matrix);if(formatZero)s.B10.z='0000';if(eta){s.K9={t:'s',v:'ETA'};s.K10={t:'n',v:46336,z:'dd/mm/yyyy'};s['!ref']='A1:K'+matrix.length;}
    XLSX.utils.book_append_sheet(b,s,'Entrada');return Array.from(new Uint8Array(XLSX.write(b,{type:'array',bookType})));
   },{rows,headers,cover,formatZero,legacy,bookType,eta}));
  }
  async function upload(buffer,name='entrada.xlsx',{eta='2026-11-07',invoice='FORM-IGNORED'}={}){
   await page.locator('#transitFactura').fill(invoice);await page.locator('#transitEta').fill(eta);await page.locator('#transitFile').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer});const before=messages.length;await page.locator('button[onclick="importarMaterialesTransito()"]').click();await page.waitForFunction(()=>document.getElementById('transitFile').files.length===0||document.getElementById('transitStatus')?.textContent.startsWith('Error:'));assert.ok(messages.length>before);
  }
  const bytes=await workbook([line('0001',2,'F-1'),line('0001',3,'F-1'),line('0001',4,'F1')]);
  await upload(bytes);assert.match(messages.at(-1),/2 nuevas/);let state=await page.evaluate(()=>APP.transitMaterials);
  assert.equal(state.length,3);assert.equal(state.find(r=>r.factura==='F-1').cantidad,5);assert.equal(state.find(r=>r.factura==='F1').cantidad,4);assert.ok(state.every(r=>!r.cliente));assert.ok(state.every(r=>r.codigoMaterial!=='SUPPLIER-XYZ'));
  assert.equal(await page.locator('.audit-invoice').count(),3);await upload(bytes);assert.match(messages.at(-1),/2 actualizadas/);assert.equal(await page.evaluate(()=>APP.transitMaterials.length),3);
  // The original invoice format is still accepted and updates the same imported material.
  const legacy=await page.evaluate(()=>{const b=XLSX.utils.book_new();XLSX.utils.book_append_sheet(b,XLSX.utils.json_to_sheet([{Referencia:'0001',Descripcion:'Factura original',Cantidad:9,Factura:'F-1',ETA:'2026-11-07'}]),'Factura');return Array.from(new Uint8Array(XLSX.write(b,{type:'array',bookType:'xlsx'})));});
  await upload(Buffer.from(legacy),'factura-original.xlsx');assert.equal(await page.evaluate(()=>APP.transitMaterials.find(r=>r.factura==='F-1').cantidad),9);assert.equal(await page.evaluate(()=>APP.transitMaterials.length),3);
  const csv=Buffer.from('Referencia,Descripcion,Cantidad,Factura,ETA\nCSV-1,Original CSV,2,CSV-INV,2026-11-07\n');await upload(csv,'factura.csv');assert.equal(await page.evaluate(()=>APP.transitMaterials.find(r=>r.factura==='CSV-INV').cantidad),2);
  // New rows with formatted numeric identifiers retain visible leading zeroes.
  await upload(await workbook([line(7,6,'PADDED')],{formatZero:true}),'codigo-formateado.xlsx');assert.equal(await page.evaluate(()=>APP.transitMaterials.find(r=>r.factura==='PADDED').codigoMaterial),'0007');
  await upload(await workbook([line('DATED',1,'DATE-INV')],{eta:true}),'con-eta.xlsx');assert.equal(await page.evaluate(()=>APP.transitMaterials.find(r=>r.factura==='DATE-INV').eta),'2026-11-10');
  // Without an ETA in either source, it stays unknown; packslip is never interpreted as a date.
  await upload(await workbook([line('NO-ETA',1,'NO-ETA-INV')]),'sin-eta.xlsx',{eta:''});assert.match(messages.at(-1),/sin ETA/);assert.equal(await page.evaluate(()=>APP.transitMaterials.find(r=>r.factura==='NO-ETA-INV').eta),'');
  const saved=await page.evaluate(()=>JSON.stringify(APP.transitMaterials));
  await upload(await workbook([line('0001',999,'F-1'),line('',5,'BAD')]),'invalid.xlsx');assert.match(messages.at(-1),/falta el código/);assert.equal(await page.evaluate(()=>JSON.stringify(APP.transitMaterials)),saved);
  await page.evaluate(()=>{APP.lineItems=[{clienteId:'C1',clienteNombre:'Cliente',pedidoCliente:'ORDER-1',lineaPedidoCliente:'1',articulo:'0001',cantidadPendienteSinStock:2,cantidadPendienteServir:2}];});
  assert.ok(await page.evaluate(()=>TMSDomain.transitForLine(APP.lineItems[0],APP.transitMaterials).invoices.includes('F-1')));
  await page.setViewportSize({width:390,height:844});await page.screenshot({path:'/tmp/tms-transit-entry-import-mobile.png'});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),true);
  await page.evaluate(()=>{APP.currentUserProfile={...APP.currentUserProfile,rol:'Solo lectura'};renderMaterialesTransito();});assert.equal(await page.locator('#transitFile').count(),0);
  if(process.env.TMS_TRANSIT_ENTRY_XLSX){
   await page.evaluate(()=>{APP.currentUserProfile={...APP.currentUserProfile,rol:'Admin'};APP.transitMaterials=[];renderMaterialesTransito();});
   await upload(fs.readFileSync(process.env.TMS_TRANSIT_ENTRY_XLSX),'private-template.xlsx');
   const actual=await page.evaluate(()=>({rows:APP.transitMaterials.length,materials:new Set(APP.transitMaterials.map(r=>r.codigoMaterial)).size,invoices:new Set(APP.transitMaterials.map(r=>r.factura)).size,quantity:APP.transitMaterials.reduce((sum,r)=>sum+r.cantidad,0),allEta:APP.transitMaterials.every(r=>r.eta==='2026-11-07'),allUnassigned:APP.transitMaterials.every(r=>!r.cliente)}));
   if(process.env.TMS_TRANSIT_ENTRY_EXPECTED){const expected=JSON.parse(process.env.TMS_TRANSIT_ENTRY_EXPECTED);for(const [key,value] of Object.entries(expected))assert.equal(actual[key],value);}assert.ok(actual.allEta&&actual.allUnassigned);console.log('Private file, local-only reconciliation:',JSON.stringify(actual));
  }
  assert.deepEqual(errors,[]);console.log('PASS: entry-template header detection, grouped invoices, safe reimport, legacy Excel/CSV, identifiers, dates, missing ETA, atomic validation, calendar matching and mobile/read-only UI.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});

// Synthetic stock and accounts. Supabase network is blocked; no production writes.
const {chromium}=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright');
const assert=require('node:assert/strict'),fs=require('node:fs'),D=require('../inventory-domain.js');
const base=process.env.TMS_TEST_URL||'http://127.0.0.1:8766';
const ctRows=[],adRows=[];
for(let i=1;i<=12;i++){const code='P'+String(i).padStart(2,'0');ctRows.push([code,'Sobrante '+code,20,'Piece','1015']);adRows.push([code,'Sobrante '+code,10,100+i,'UND']);}
for(let i=1;i<=22;i++){const code='N'+String(i).padStart(2,'0');ctRows.push([code,'Faltante '+code,1,'Piece','1015']);adRows.push([code,'Faltante '+code,10,100+i,'UND']);}
ctRows.push(['CT_ONLY','Solo CT',7,'Piece','1015']);
const ct=D.consolidate([['Material','Description','Quantity','Unit of measurement','Warehouse'],...ctRows],{source:'ct',asOf:'2026-10-31'}),ad=D.consolidate([['Código','Descripción','Stock','Valor total DOP','Unidad'],...adRows],{source:'ad',asOf:'2026-10-31'});
ct.semanticHash='SYNTHETIC_CT';ad.semanticHash='SYNTHETIC_AD';delete ct.semanticKey;delete ad.semanticKey;
(async()=>{
 const browser=await chromium.launch({headless:true}),errors=[],db={revision:0,payload:{version:1,sources:{ct,ad},cases:{},settings:{sameScope:true}}};
 async function boot(page,role='Admin'){
  await page.evaluate(role=>{const user={id:'11111111-1111-4111-8111-111111111111',email:'counter@example.invalid'};SB.auth.getUser=async()=>({data:{user},error:null});SB.rpc=(name,args)=>window.__countRpc(name,args);SB.channel=()=>({on(type,filter,callback){window.__countRemote=callback;return this;},subscribe(){return this;}});SB.removeChannel=async()=>{};APP.userProfiles=[{userId:user.id,email:user.email,rol:role,activo:true}];APP.currentUserProfile=APP.userProfiles[0];APP.rolePermissions={[role]:{inventario:{ver:true,editar:role==='Admin',importar:role==='Admin'}}};mostrarApp(user);cambiarVista('inventario',document.getElementById('nav-inventario'));},role);
  await page.waitForFunction(()=>TMSInventory.getSyncState().revision!==null&&TMSInventory.getSyncState().revision!==undefined);
 }
 async function createPage(role='Admin'){
  const page=await browser.newPage({viewport:{width:1440,height:1000}});page.on('pageerror',error=>errors.push(error.message));
  await page.route('**/*.supabase.co/**',route=>route.fulfill({status:503,contentType:'application/json',body:'{"message":"Blocked for test"}'}));
  await page.exposeFunction('__countRpc',async(name,args)=>{
   if(name==='tms_inventory_get_snapshot')return {data:args.p_known_revision===db.revision?{revision:db.revision,unchanged:true}:{revision:db.revision,tables:{tms_inventory_state:[{id:true,payload:structuredClone(db.payload)}]}},error:null};
   assert.equal(name,'tms_inventory_save_snapshot');if(role!=='Admin')return {data:null,error:{code:'42501',message:'Read only'}};
   if(args.p_expected_revision!==db.revision)return {data:null,error:{code:'40001',message:'Concurrent update'}};
   db.payload=structuredClone(args.p_payload.tms_inventory_state[0].payload);return {data:{revision:++db.revision,changed:true},error:null};
  });
  await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>!!window.InventoryCounts);await boot(page,role);return page;
 }
 const tab=(page,name)=>page.locator(`[data-count="tab"][data-tab="${name}"]`).first().click();
 const settled=page=>page.waitForFunction(()=>TMSInventory.getSyncState().phase==='synced');
 const recordCount=()=>db.payload.counting.cycles.find(row=>row.id===db.payload.counting.activeCycleId).records.filter(row=>!row.voided).length;
 try{
  const a=await createPage();assert.equal(await a.locator('.inv-priority-columns section').first().locator('tbody tr').count(),10);assert.equal(await a.locator('.inv-priority-columns section').nth(1).locator('tbody tr').count(),20);
  assert.equal(await a.locator('.inv-priority-columns section').first().locator('tbody tr').first().locator('td').first().innerText(),'P12\nSobrante P12');
  assert.equal(await a.locator('.inv-top-surplus .inv-priority-bar i').count(),10);assert.equal(await a.locator('.inv-top-shortage .inv-priority-bar i').count(),20);
  assert.equal(await a.locator('.inv-top-surplus .inv-priority-bar i').first().evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(20, 150, 111)');
  assert.equal(await a.locator('.inv-top-shortage .inv-priority-bar i').first().evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(217, 72, 72)');
  await a.screenshot({path:'/tmp/tms-priority-bars.png',fullPage:true});
  await tab(a,'data');
  const workbook=await a.evaluate(()=>{const b=XLSX.utils.book_new();XLSX.utils.book_append_sheet(b,XLSX.utils.aoa_to_sheet([['CODIGO','DESCRIPCION','CODIGO BARRAS','CENTRAL','COSTO PROMEDIO'],['P12','Sobrante P12','750000000012',20,11.2]]),'Articulos');XLSX.utils.book_append_sheet(b,XLSX.utils.aoa_to_sheet([['CODIGO','UBICACION','STOCK UBICACION'],['P12','A01',10],['P12','A02',10]]),'Ubicaciones');return Array.from(new Uint8Array(XLSX.write(b,{type:'array',bookType:'xlsx'})));});
  const file={name:'Maestro_ubicaciones.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(workbook)};
  await a.locator('#countMasterFile').setInputFiles(file);await a.locator('[data-count="commit-master"]').click();await settled(a);
  await a.locator('#countMasterFile').setInputFiles(file);await a.locator('#countImportSheet').selectOption('1');await a.locator('[data-count="commit-master"]').click();await settled(a);
  await tab(a,'guided');await settled(a);assert.equal(await a.evaluate(()=>Object.values(InventoryCountDomain.active(TMSInventory.getState()).items).filter(row=>row.required).length),30);
  assert.equal(await a.locator('#countExpectedLocation').inputValue(),'A01');assert.equal(await a.locator('#scan-location').count(),0);
  await a.locator('#countQuantity').fill('10');await a.locator('#countSaveRecord').click();assert.match(await a.locator('#countNotice').innerText(),/Confirma/);assert.equal(recordCount(),0);
  await a.locator('#countLocationConfirmed').check();await a.locator('#countSaveRecord').click();await settled(a);assert.equal(recordCount(),1);
  assert.equal(await a.locator('#countExpectedLocation').inputValue(),'A02');await a.locator('#countLocationChoice').selectOption('correct');await a.locator('#countActualLocation').fill('REAL-02');await a.locator('#countQuantity').fill('10');await a.locator('#countLocationConfirmed').check();await a.screenshot({path:'/tmp/tms-counts-guided.png'});await a.locator('#countSaveRecord').click();await settled(a);
  assert.equal(recordCount(),2);assert.equal(db.payload.counting.cycles[0].records[1].actualLocation,'REAL-02');
  await tab(a,'results');await a.locator('[data-count="process"][data-code="P12"]').click();await settled(a);
  await tab(a,'free');await a.locator('#countScanArticle').fill('750000000012');await a.locator('#countLookupForm button').click();await settled(a);assert.match(await a.locator('#countBody h3').first().innerText(),/P12/);
  await a.locator('#countQuantity').fill('2');await a.locator('#countLocationConfirmed').check();a.once('dialog',dialog=>dialog.accept());await a.locator('#countSaveRecord').click();await settled(a);assert.equal(recordCount(),3);
  await a.locator('#countScanArticle').fill('CT_ONLY');await a.locator('#countLookupForm button').click();await settled(a);assert.equal(await a.locator('#countExpectedLocation').inputValue(),'');assert.equal(await a.locator('#countActualLocation').inputValue(),'');
  await a.locator('#countActualLocation').fill('RACK-C');await a.locator('#countQuantity').fill('7');await a.locator('#countLocationConfirmed').check();await a.locator('#countSaveRecord').click();await settled(a);
  await tab(a,'results');assert.equal(await a.locator('[data-count="close-cycle"]').isEnabled(),false);await a.locator('[data-count="process"][data-code="CT_ONLY"]').click();await settled(a);
  const third=db.payload.counting.cycles[0].records[2].id;await a.locator(`[data-count="correct"][data-record="${third}"]`).click();await a.locator('#countCorrectQuantity').fill('1');await a.locator('#countCorrectReason').fill('Reconteo de caja');await a.locator('[data-count="save-correction"]').click();await settled(a);assert.equal(db.payload.counting.cycles[0].records[2].quantity,1);assert.equal(db.payload.counting.cycles[0].processed.P12,undefined);
  const downloaded=a.waitForEvent('download');await a.locator('[data-count="export"]').click();const download=await downloaded;const bytes=fs.readFileSync(await download.path());
  const exported=await a.evaluate(bytes=>{const wb=XLSX.read(new Uint8Array(bytes),{type:'array'});return {names:wb.SheetNames,altas:XLSX.utils.sheet_to_json(wb.Sheets['Altas AD'],{header:1})};},Array.from(bytes));assert.ok(exported.names.includes('Correcciones'));assert.equal(exported.altas[1][0],'CT_ONLY');assert.equal(exported.altas[1][3],7);
  const b=await createPage();await tab(a,'free');await a.locator('#countScanArticle').fill('N22');await a.locator('#countLookupForm button').click();await settled(a);await a.locator('#countQuantity').fill('3');await a.locator('#countActualLocation').fill('DRAFT-BIN');
  await tab(b,'free');await b.locator('#countScanArticle').fill('N21');await b.locator('#countLookupForm button').click();await settled(b);await b.locator('#countQuantity').fill('1');await b.locator('#countActualLocation').fill('B-BIN');await b.locator('#countLocationConfirmed').check();await b.locator('#countSaveRecord').click();await settled(b);
  await a.evaluate(revision=>window.__countRemote({new:{revision}}),db.revision);await a.waitForFunction(()=>TMSInventory.getSyncState().conflict);assert.equal(await a.locator('#countQuantity').inputValue(),'3');assert.equal(await a.locator('#countActualLocation').inputValue(),'DRAFT-BIN');assert.equal(await a.locator('#countSyncPanel').isVisible(),true);
  await a.reload({waitUntil:'networkidle'});await boot(a);await a.waitForFunction(()=>TMSInventory.getSyncState().conflict);assert.equal(await a.locator('#countQuantity').inputValue(),'3');assert.equal(await a.locator('#countActualLocation').inputValue(),'DRAFT-BIN');
  const read=await createPage('Solo lectura');await tab(read,'data');assert.equal(await read.locator('#countMasterFile').count(),0);await tab(read,'guided');assert.equal(await read.locator('#countSaveRecord').isEnabled(),false);
  await b.setViewportSize({width:390,height:844});await b.screenshot({path:'/tmp/tms-counts-mobile.png'});assert.equal(await b.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);await b.setViewportSize({width:1440,height:1000});await b.screenshot({path:'/tmp/tms-counts-desktop.png'});
  assert.deepEqual(errors,[]);console.log('PASS: top 10/20, original master and location imports, mandatory guided queue, location confirmation/correction, free barcode and boxes, AD-registration export, audit, read-only access and concurrent draft preservation.');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

// Tests the real UI with an in-memory RPC fixture. All real Supabase calls blocked.
const {chromium}=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright');
const assert=require('node:assert/strict');
const base=process.env.TMS_TEST_URL||'http://127.0.0.1:8766';
const initial=()=>({version:1,sources:{ct:null,ad:null},cases:{},settings:{sameScope:true}});
(async()=>{
  const browser=await chromium.launch({headless:true});const db={revision:0,payload:initial()},calls=[],errors=[];
  async function pageFor(role='Admin'){
    const page=await browser.newPage({viewport:{width:1440,height:1000}});page.on('pageerror',error=>errors.push(error.message));
    await page.route('**/*.supabase.co/**',route=>route.fulfill({status:503,contentType:'application/json',body:'{"message":"Production blocked in fixture"}'}));
    await page.exposeFunction('__inventoryRpc',async(name,args)=>{
      calls.push({name,args});
      if(name==='tms_inventory_get_snapshot')return {data:args.p_known_revision===db.revision?{revision:db.revision,unchanged:true}:{revision:db.revision,tables:{tms_inventory_state:[{id:true,payload:structuredClone(db.payload)}]}},error:null};
      assert.equal(name,'tms_inventory_save_snapshot');if(role!=='Admin')return {data:null,error:{code:'42501',message:'Read only'}};
      if(args.p_expected_revision!==db.revision)return {data:null,error:{code:'40001',message:'Another session updated inventory'}};
      db.payload=structuredClone(args.p_payload.tms_inventory_state[0].payload);return {data:{revision:++db.revision,changed:true},error:null};
    });
    await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>!!window.TMSInventory&&!!window.TMSAuditCore&&!!window.XLSX);
    await page.evaluate(role=>{
      const authUser={id:'11111111-1111-4111-8111-111111111111',email:'inventory-test@example.invalid'};
      SB.auth.getUser=async()=>({data:{user:authUser},error:null});SB.rpc=(name,args)=>window.__inventoryRpc(name,args);
      SB.channel=()=>({on(type,filter,callback){window.__inventoryChange=callback;return this;},subscribe(cb){return this;}});SB.removeChannel=async()=>{};
      APP.userProfiles=[{userId:authUser.id,authUserId:authUser.id,nombre:'Fixture',email:authUser.email,rol:role,activo:true}];APP.currentUserProfile=APP.userProfiles[0];APP.rolePermissions={[role]:{inventario:{ver:true,editar:role==='Admin',importar:role==='Admin'}}};
      mostrarApp(authUser);cambiarVista('inventario',document.getElementById('nav-inventario'));
    },role);
    await page.waitForFunction(()=>TMSInventory.getSyncState().revision!==undefined&&TMSInventory.getSyncState().revision!==null);return page;
  }
  async function upload(page,source,file,date='2026-10-30'){
    await page.locator('#inventory'+source.toUpperCase()+'File').setInputFiles(file);
    await page.waitForSelector('#inventoryImportDialog');await page.locator('#inventoryAsOf').fill(date);
    await page.locator('[data-inv="preview"]').click();await assert.equal(await page.locator('#inventoryConfirmImport').isEnabled(),true);
    await page.locator('#inventoryConfirmImport').click();await page.waitForSelector('#inventoryImportDialog',{state:'detached'});await page.waitForFunction(()=>TMSInventory.getSyncState().phase==='synced');
  }
  const ct=Buffer.from('Material;Description;Quantity;Unit of measurement;Warehouse;Batchnumber\nA;Articulo A;12;Piece;1015;L1\nA;Articulo A;8;Piece;1015;L2\nB;Articulo B;5;Piece;1015;L1\nC;Articulo C;2;Piece;1015;L1');
  try{
    const a=await pageFor();assert.match(await a.locator('#inventoryMount').innerText(),/Carga los dos reportes/);
    await a.locator('#inventoryADFile').setInputFiles({name:'vacio.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.alloc(0)});await a.waitForFunction(()=>document.getElementById('inventoryMessage').textContent.includes('0 bytes'));
    if(process.env.TMS_CT_CSV){await upload(a,'ct',process.env.TMS_CT_CSV);assert.equal(await a.evaluate(()=>TMSInventory.getState().sources.ct.totals.quantity),386962);const rev=db.revision;if(process.env.TMS_CT_CSV_ALT){await upload(a,'ct',process.env.TMS_CT_CSV_ALT);assert.equal(db.revision,rev);}}
    await upload(a,'ct',{name:'CT.csv',mimeType:'text/csv',buffer:ct});assert.equal(await a.evaluate(()=>TMSInventory.getState().sources.ct.materials[0].quantity),20);
    const excel=await a.evaluate(()=>{const b=XLSX.utils.book_new();XLSX.utils.book_append_sheet(b,XLSX.utils.aoa_to_sheet([['Código','Descripción','Stock','Valor total DOP','Unidad'],['A','Articulo A',10,100,'UND'],['A','Articulo A',20,400,'UND'],['B','Articulo B',1,1000,'UND'],['D','Articulo D',3,300,'UND']]),'Stock AD');return Array.from(new Uint8Array(XLSX.write(b,{type:'array',bookType:'xlsx'})));});
    await upload(a,'ad',{name:'AD.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(excel)});
    const report=await a.evaluate(()=>TMSInventory.getReport());assert.equal(report.rows[0].code,'B');assert.equal(report.rows[0].impactDop,4000);assert.equal(report.rows.find(row=>row.code==='A').unitCostDop,500/30);assert.equal(report.rows.find(row=>row.code==='C').impactDop,null);assert.equal(report.summary.unvalued,1);
    const ctExcel=await a.evaluate(csv=>{const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(InventoryDomain.parseCsv(csv)),'Stock CT');return Array.from(new Uint8Array(XLSX.write(book,{type:'array',bookType:'xlsx'})));},ct.toString());
    const beforeExcel=db.revision;await upload(a,'ct',{name:'CT.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(ctExcel)});assert.equal(db.revision,beforeExcel,'CT Excel and CSV represent the same inventory, not additional stock');
    const rev=db.revision;await upload(a,'ct',{name:'CT-copia.csv',mimeType:'text/csv',buffer:ct});assert.equal(db.revision,rev);assert.match(await a.locator('#inventoryMessage').innerText(),/ya está cargado/);
    const download=a.waitForEvent('download');await a.locator('[data-inv="export"]').click();assert.match((await download).suggestedFilename(),/\.xlsx$/);
    const b=await pageFor();await a.locator('[data-inv="case"][data-code="B"]').click();await a.locator('#inventoryCase_owner').fill('Alice pendiente');
    await b.locator('[data-inv="case"][data-code="A"]').click();await b.locator('#inventoryCase_owner').fill('Bob guardado');await b.locator('#inventoryCase_status').selectOption('en_revision');await b.locator('[data-inv="save-case"]').click();await b.waitForFunction(()=>TMSInventory.getSyncState().phase==='synced');
    await a.evaluate(revision=>window.__inventoryChange({new:{revision}}),db.revision);await a.waitForFunction(()=>TMSInventory.getSyncState().conflict);
    assert.equal(await a.locator('#inventoryCase_owner').inputValue(),'Alice pendiente');await a.locator('[data-inv="save-case"]').click();assert.equal(await a.evaluate(()=>Object.values(TMSInventory.getState().cases).find(row=>row.owner==='Alice pendiente').owner),'Alice pendiente');
    assert.equal(Object.values(db.payload.cases).some(row=>row.owner==='Alice pendiente'),false);
    a.once('dialog',dialog=>dialog.accept());await a.locator('[data-inv="reload"]').click();await a.waitForFunction(()=>!TMSInventory.getSyncState().conflict);assert.equal(await a.evaluate(()=>Object.values(TMSInventory.getState().cases).some(row=>row.owner==='Bob guardado')),true);
    const readonly=await pageFor('Solo lectura');assert.equal(await readonly.locator('#inventoryCTFile').count(),0);await readonly.locator('[data-inv="case"]').first().click();assert.equal(await readonly.locator('[data-inv="save-case"]').count(),0);
    await a.screenshot({path:'/tmp/tms-inventory-desktop.png'});await a.setViewportSize({width:390,height:844});await a.screenshot({path:'/tmp/tms-inventory-mobile.png'});assert.equal(await a.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth),true);
    assert.deepEqual(errors,[]);console.log('PASS: inventory CSV/XLSX imports, weighted DOP valuation, consolidation, duplicate snapshots, empty AD, exports, role access, simultaneous sessions, draft preservation, responsive UI.');
  }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

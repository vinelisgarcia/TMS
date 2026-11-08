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
    await page.exposeFunction('__showroomRpc',async(name,args)=>{
      calls.push({name,args});
      if(name==='tms_showroom_get_snapshot')return {data:args.p_known_revision===db.revision?{revision:db.revision,unchanged:true}:{revision:db.revision,tables:{tms_showroom_state:[{id:true,payload:structuredClone(db.payload)}]}},error:null};
      assert.equal(name,'tms_showroom_save_snapshot');if(role!=='Admin')return {data:null,error:{code:'42501',message:'Read only'}};
      if(args.p_expected_revision!==db.revision)return {data:null,error:{code:'40001',message:'Another session updated showroom'}};
      db.payload=structuredClone(args.p_payload.tms_showroom_state[0].payload);return {data:{revision:++db.revision,changed:true},error:null};
    });
    await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>!!window.TMSShowroom&&!!window.TMSAuditCore&&!!window.XLSX);
    await page.evaluate(role=>{
      const authUser={id:'11111111-1111-4111-8111-111111111111',email:'showroom-test@example.invalid'};
      SB.auth.getUser=async()=>({data:{user:authUser},error:null});SB.rpc=(name,args)=>window.__showroomRpc(name,args);
      SB.channel=()=>({on(type,filter,callback){window.__showroomChange=callback;return this;},subscribe(cb){return this;}});SB.removeChannel=async()=>{};
      APP.userProfiles=[{userId:authUser.id,authUserId:authUser.id,nombre:'Fixture',email:authUser.email,rol:role,activo:true}];APP.currentUserProfile=APP.userProfiles[0];APP.rolePermissions={[role]:{showroom:{ver:true,editar:role==='Admin',importar:role==='Admin'}}};
      mostrarApp(authUser);cambiarVista('showroom',document.getElementById('nav-showroom'));
    },role);
    await page.waitForFunction(()=>TMSShowroom.getSyncState().revision!==undefined&&TMSShowroom.getSyncState().revision!==null);return page;
  }
  async function upload(page,source,file,date='2026-10-30'){
    await page.locator('#showroom'+source.toUpperCase()+'File').setInputFiles(file);
    await page.waitForSelector('#showroomImportDialog');
    if(source==='ct'&&typeof file==='string'&&file===process.env.TMS_SHOWROOM_PHYSICAL){assert.equal(await page.locator('#showroomSheet option:checked').textContent(),'Conteo');assert.equal(await page.locator('#showroomAsOf').inputValue(),'2026-10-09');}
    await page.locator('#showroomAsOf').fill(date);
    await page.locator('[data-show-inv="preview"]').click();await assert.equal(await page.locator('#showroomConfirmImport').isEnabled(),true);
    await page.locator('#showroomConfirmImport').click();await page.waitForSelector('#showroomImportDialog',{state:'detached'});await page.waitForFunction(()=>TMSShowroom.getSyncState().phase==='synced');
  }
  const ct=Buffer.from('Material;Description;Quantity;Unit of measurement;Warehouse;Batchnumber\nA;Articulo A;12;Piece;1015;L1\nA;Articulo A;8;Piece;1015;L2\nB;Articulo B;5;Piece;1015;L1\nC;Articulo C;2;Piece;1015;L1');
  try{
    const a=await pageFor();assert.match(await a.locator('#showroomMount').innerText(),/Carga los dos reportes/);
    await a.locator('#showroomADFile').setInputFiles({name:'vacio.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.alloc(0)});await a.waitForFunction(()=>document.getElementById('showroomMessage').textContent.includes('0 bytes'));
    await upload(a,'ct',{name:'CT.csv',mimeType:'text/csv',buffer:ct});assert.equal(await a.evaluate(()=>TMSShowroom.getState().sources.ct.materials[0].quantity),20);
    const excel=await a.evaluate(()=>{const b=XLSX.utils.book_new();XLSX.utils.book_append_sheet(b,XLSX.utils.aoa_to_sheet([['Código','Descripción','Stock','Valor total DOP','Unidad'],['A','Articulo A',10,100,'UND'],['A','Articulo A',20,400,'UND'],['B','Articulo B',1,1000,'UND'],['D','Articulo D',3,300,'UND']]),'Stock AD');return Array.from(new Uint8Array(XLSX.write(b,{type:'array',bookType:'xlsx'})));});
    await upload(a,'ad',{name:'AD.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(excel)});
    const report=await a.evaluate(()=>TMSShowroom.getReport());assert.equal(report.rows[0].code,'B');assert.equal(report.rows[0].impactDop,4000);assert.equal(report.rows.find(row=>row.code==='A').unitCostDop,500/30);assert.equal(report.rows.find(row=>row.code==='C').impactDop,null);assert.equal(report.summary.unvalued,1);assert.equal(report.rows.find(row=>row.code==='D').status,'not_counted');assert.equal(report.rows.find(row=>row.code==='D').difference,null);assert.equal(report.rows.find(row=>row.code==='B').difference,4);
    const ctExcel=await a.evaluate(csv=>{const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(ShowroomDomain.parseCsv(csv)),'Stock CT');return Array.from(new Uint8Array(XLSX.write(book,{type:'array',bookType:'xlsx'})));},ct.toString());
    const beforeExcel=db.revision;await upload(a,'ct',{name:'CT.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(ctExcel)});assert.equal(db.revision,beforeExcel,'CT Excel and CSV represent the same showroom, not additional stock');
    const rev=db.revision;await upload(a,'ct',{name:'CT-copia.csv',mimeType:'text/csv',buffer:ct});assert.equal(db.revision,rev);assert.match(await a.locator('#showroomMessage').innerText(),/ya está cargado/);
    const download=a.waitForEvent('download');await a.locator('[data-show-inv="export"]').click();assert.match((await download).suggestedFilename(),/\.xlsx$/);
    const b=await pageFor();await a.locator('[data-show-inv="case"][data-code="B"]').click();await a.locator('#showroomCase_owner').fill('Alice pendiente');
    await b.locator('[data-show-inv="case"][data-code="A"]').click();await b.locator('#showroomCase_owner').fill('Bob guardado');await b.locator('#showroomCase_status').selectOption('en_revision');await b.locator('[data-show-inv="save-case"]').click();await b.waitForFunction(()=>TMSShowroom.getSyncState().phase==='synced');
    await a.evaluate(revision=>window.__showroomChange({new:{revision}}),db.revision);await a.waitForFunction(()=>TMSShowroom.getSyncState().conflict);
    assert.equal(await a.locator('#showroomCase_owner').inputValue(),'Alice pendiente');await a.locator('[data-show-inv="save-case"]').click();assert.equal(await a.evaluate(()=>Object.values(TMSShowroom.getState().cases).find(row=>row.owner==='Alice pendiente').owner),'Alice pendiente');
    assert.equal(Object.values(db.payload.cases).some(row=>row.owner==='Alice pendiente'),false);
    a.once('dialog',dialog=>dialog.accept());await a.locator('[data-show-inv="reload"]').click();await a.waitForFunction(()=>!TMSShowroom.getSyncState().conflict);assert.equal(await a.evaluate(()=>Object.values(TMSShowroom.getState().cases).some(row=>row.owner==='Bob guardado')),true);
    const readonly=await pageFor('Solo lectura');assert.equal(await readonly.locator('#showroomCTFile').count(),0);await readonly.locator('[data-show-inv="case"]').first().click();assert.equal(await readonly.locator('[data-show-inv="save-case"]').count(),0);
    await a.screenshot({path:'/tmp/tms-showroom-desktop.png'});await a.setViewportSize({width:390,height:844});await a.screenshot({path:'/tmp/tms-showroom-mobile.png'});assert.equal(await a.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth),true);
    if(process.env.TMS_SHOWROOM_PHYSICAL&&process.env.TMS_SHOWROOM_VIRTUAL){await upload(a,'ct',process.env.TMS_SHOWROOM_PHYSICAL,'2026-10-09');await upload(a,'ad',process.env.TMS_SHOWROOM_VIRTUAL,'2026-11-08');const real=await a.evaluate(()=>TMSShowroom.getState());assert.equal(real.sources.ct.worksheet,'Conteo');assert.equal(real.sources.ct.acceptedRowCount,1690);assert.equal(real.sources.ct.materials.length,1623);assert.equal(real.sources.ct.totals.quantity,1754);assert.equal(real.sources.ad.totals.quantity,4615);assert.equal(real.sources.ad.materials.length,2807);console.log('PASS: real showroom reports: Conteo selected; 1690 physical rows, 1623 materials, 1754 units; 2807 virtual materials, 4615 units.');
      if(process.env.TMS_SHOWROOM_REPORT_OUTPUT){
        await a.locator('#showroomStatus').selectOption('all');
        const next=a.waitForEvent('download');await a.locator('[data-show-inv="export"]').click();const actual=await next;
        await actual.saveAs(process.env.TMS_SHOWROOM_REPORT_OUTPUT);
        const output=require('node:fs').readFileSync(process.env.TMS_SHOWROOM_REPORT_OUTPUT);
        const verified=await a.evaluate(bytes=>{const wb=XLSX.read(new Uint8Array(bytes),{type:'array'});return {names:wb.SheetNames,rows:XLSX.utils.sheet_to_json(wb.Sheets.Conciliacion).length,physical:XLSX.utils.sheet_to_json(wb.Sheets['Detalle fisico']).length};},[...output]);
        assert.deepEqual(verified.names,['Conciliacion','Fuentes y metodo','Detalle fisico','Detalle virtual']);
        assert.equal(verified.rows,3847);assert.equal(verified.physical,1690);
        console.log('PASS: full real comparison exported with source method, all 3847 articles and 1690 physical detail records.');
      }
    }
    assert.deepEqual(errors,[]);console.log('PASS: showroom CSV/XLSX imports, weighted DOP valuation, consolidation, duplicate snapshots, empty AD, exports, role access, simultaneous sessions, draft preservation, responsive UI.');
  }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

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
  const ct=Buffer.from('Material;Description;Quantity;Unit of measurement;Warehouse;Batchnumber\nA;Articulo A;12;Piece;1015;L1\nA;Articulo A;8;Piece;1015;L2\nB;Articulo B;5;Piece;1015;L1\nC;Articulo C;2;Piece;1015;L1\n0007;Articulo igual;4;Piece;1015;L1\nZERO;Conteo cero;0;Piece;1015;L1');
  try{
    const a=await pageFor();assert.match(await a.locator('#showroomMount').innerText(),/Carga los dos reportes/);
    await a.locator('#showroomADFile').setInputFiles({name:'vacio.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.alloc(0)});await a.waitForFunction(()=>document.getElementById('showroomMessage').textContent.includes('0 bytes'));
    await upload(a,'ct',{name:'CT.csv',mimeType:'text/csv',buffer:ct});assert.equal(await a.evaluate(()=>TMSShowroom.getState().sources.ct.materials.find(row=>row.code==='A').quantity),20);
    const excel=await a.evaluate(()=>{const b=XLSX.utils.book_new();XLSX.utils.book_append_sheet(b,XLSX.utils.aoa_to_sheet([['Código','Descripción','Stock','Valor total DOP','Unidad'],['A','Articulo A',10,100,'UND'],['A','Articulo A',20,400,'UND'],['B','Articulo B',1,1000,'UND'],['D','Articulo D',3,300,'UND'],['0007','Articulo igual',4,80,'UND'],['ZERO','Conteo cero',0,0,'UND']]),'Stock AD');return Array.from(new Uint8Array(XLSX.write(b,{type:'array',bookType:'xlsx'})));});
    await upload(a,'ad',{name:'AD.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(excel)});
    const report=await a.evaluate(()=>TMSShowroom.getReport());assert.equal(report.rows[0].code,'B');assert.equal(report.rows[0].impactDop,4000);assert.equal(report.rows.find(row=>row.code==='A').unitCostDop,500/30);assert.equal(report.rows.find(row=>row.code==='C').impactDop,null);assert.equal(report.summary.unvalued,1);assert.equal(report.rows.find(row=>row.code==='D').status,'not_counted');assert.equal(report.rows.find(row=>row.code==='D').difference,null);assert.equal(report.rows.find(row=>row.code==='B').difference,4);
    const ctExcel=await a.evaluate(csv=>{const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(ShowroomDomain.parseCsv(csv)),'Stock CT');return Array.from(new Uint8Array(XLSX.write(book,{type:'array',bookType:'xlsx'})));},ct.toString());
    const beforeExcel=db.revision;await upload(a,'ct',{name:'CT.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(ctExcel)});assert.equal(db.revision,beforeExcel,'CT Excel and CSV represent the same showroom, not additional stock');
    const rev=db.revision;await upload(a,'ct',{name:'CT-copia.csv',mimeType:'text/csv',buffer:ct});assert.equal(db.revision,rev);assert.match(await a.locator('#showroomMessage').innerText(),/ya está cargado/);
    await a.route('**/vendor/exceljs-4.4.0.min.js',route=>route.abort());await a.locator('[data-show-inv="export"]').click();
    await a.waitForFunction(()=>document.getElementById('showroomMessage').textContent.includes('No se pudo cargar el generador'));
    assert.equal(await a.locator('[data-show-inv="export"]').isEnabled(),true);await a.unroute('**/vendor/exceljs-4.4.0.min.js');
    const exportRows=async()=>{
      const download=a.waitForEvent('download');await a.locator('[data-show-inv="export"]').click();const file=await download;assert.match(file.suggestedFilename(),/\.xlsx$/);
      if(process.env.TMS_SHOWROOM_REPORT_OUTPUT)await file.saveAs(process.env.TMS_SHOWROOM_REPORT_OUTPUT);
      const bytes=require('node:fs').readFileSync(await file.path());
      return a.evaluate(async bytes=>{const book=XLSX.read(new Uint8Array(bytes),{type:'array',cellNF:true}),sheet=book.Sheets.Conciliacion,native=await new ExcelJS.Workbook().xlsx.load(new Uint8Array(bytes)),ws=native.getWorksheet('Conciliacion');return {names:book.SheetNames,range:sheet['!ref'],rows:XLSX.utils.sheet_to_json(sheet,{header:1,range:11,defval:null}),cells:Object.fromEntries(Object.entries(sheet).filter(([key])=>!key.startsWith('!')&&XLSX.utils.decode_cell(key).r>=12)),summary:[ws.getCell('A8').value,ws.getCell('C8').value,ws.getCell('E8').value],scope:ws.getCell('A5').value,notes:ws.getCell('A10').value,headerFill:ws.getCell('A12').fill.fgColor.argb,headerFont:ws.getCell('A12').font.color.argb,bandFill:ws.getCell('A13').fill.fgColor.argb,views:ws.views,filter:ws.autoFilter,printTitles:ws.pageSetup.printTitlesRow};},[...bytes]);
    };
    assert.equal(await a.locator('#showroomStatus').inputValue(),'all');
    const exported=await exportRows();assert.deepEqual(exported.names,['Conciliacion']);assert.equal(exported.range,'A1:F18');
    assert.deepEqual(exported.summary,[6,1880,3]);assert.match(exported.scope,/Todas las referencias de ambos inventarios · 6 de 6 referencias/);assert.match(exported.notes,/solo las referencias con valor virtual conocido/);
    assert.equal(exported.headerFill,'FF183F66');assert.equal(exported.headerFont,'FFFFFFFF');assert.equal(exported.bandFill,'FFF6F8FB');assert.equal(exported.views[0].ySplit,12);assert.equal(exported.views[0].showGridLines,false);assert.equal(exported.filter,'A12:F18');assert.equal(exported.printTitles,'12:12');
    assert.deepEqual(exported.rows[0],['Referencia','Descripción','Inventario','Valor económico Virtual','Conteo físico','Diferencia']);
    const exportedByCode=Object.fromEntries(exported.rows.slice(1).map(row=>[row[0],row]));
    assert.deepEqual(exportedByCode.A,['A','Articulo A',30,500,20,-10]);assert.deepEqual(exportedByCode.B,['B','Articulo B',1,1000,5,4]);assert.deepEqual(exportedByCode.C,['C','Articulo C',0,null,2,2]);
    assert.deepEqual(exportedByCode.D,['D','Articulo D',3,300,null,null]);assert.deepEqual(exportedByCode['0007'],['0007','Articulo igual',4,80,4,0]);assert.deepEqual(exportedByCode.ZERO,['ZERO','Conteo cero',0,0,0,0]);
    for(const [key,cell]of Object.entries(exported.cells)){if(/^[C-F]/.test(key)&&cell.v!==undefined)assert.equal(cell.t,'n');if(/^D/.test(key)&&cell.t==='n')assert.match(cell.z,/RD\$/);}
    await a.locator('#showroomStatus').selectOption('differences');assert.equal(await a.locator('#showroomResults tbody tr').count(),3);assert.deepEqual((await exportRows()).rows,exported.rows,'Screen status filters must not exclude zero differences or virtual-only articles from the complete report');
    await a.locator('#showroomSearch').fill('Articulo A');await a.locator('#showroomFollowup').selectOption('rectificado');assert.deepEqual((await exportRows()).rows,exported.rows,'Search and follow-up filters must not remove references from the complete report');
    await a.locator('#showroomSearch').fill('');await a.locator('#showroomFollowup').selectOption('');await a.locator('#showroomStatus').selectOption('all');assert.equal(db.revision,rev,'Exporting does not save or change inventory');
    const b=await pageFor();await a.locator('[data-show-inv="case"][data-code="B"]').click();await a.locator('#showroomCase_owner').fill('Alice pendiente');
    await b.locator('[data-show-inv="case"][data-code="A"]').click();await b.locator('#showroomCase_owner').fill('Bob guardado');await b.locator('#showroomCase_status').selectOption('en_revision');await b.locator('[data-show-inv="save-case"]').click();await b.waitForFunction(()=>TMSShowroom.getSyncState().phase==='synced');
    await a.evaluate(revision=>window.__showroomChange({new:{revision}}),db.revision);await a.waitForFunction(()=>TMSShowroom.getSyncState().conflict);
    assert.equal(await a.locator('#showroomCase_owner').inputValue(),'Alice pendiente');await a.locator('[data-show-inv="save-case"]').click();assert.equal(await a.evaluate(()=>Object.values(TMSShowroom.getState().cases).find(row=>row.owner==='Alice pendiente').owner),'Alice pendiente');
    assert.equal(Object.values(db.payload.cases).some(row=>row.owner==='Alice pendiente'),false);
    a.once('dialog',dialog=>dialog.accept());await a.locator('[data-show-inv="reload"]').click();await a.waitForFunction(()=>!TMSShowroom.getSyncState().conflict);assert.equal(await a.evaluate(()=>Object.values(TMSShowroom.getState().cases).some(row=>row.owner==='Bob guardado')),true);
    const readonly=await pageFor('Solo lectura');assert.equal(await readonly.locator('#showroomCTFile').count(),0);await readonly.locator('[data-show-inv="case"]').first().click();assert.equal(await readonly.locator('[data-show-inv="save-case"]').count(),0);
    await a.screenshot({path:'/tmp/tms-showroom-desktop.png'});await a.setViewportSize({width:390,height:844});await a.screenshot({path:'/tmp/tms-showroom-mobile.png'});assert.equal(await a.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth),true);
    const code=n=>String(n).padStart(6,'0');
    await upload(a,'ct',{name:'fisico-grande.csv',mimeType:'text/csv',buffer:Buffer.from('Material;Description;Quantity;Unit of measurement;Warehouse;Batchnumber\n'+Array.from({length:800},(_,i)=>[code(i),'Material '+code(i),1,'UND','SR','L1'].join(';')).join('\n'))});
    await upload(a,'ad',{name:'virtual-grande.csv',mimeType:'text/csv',buffer:Buffer.from('Código;Descripción;Stock;Valor total DOP;Unidad\n'+Array.from({length:800},(_,i)=>[code(i+400),'Material '+code(i+400),1,10,'UND'].join(';')).join('\n'))});
    await a.locator('#showroomSearch').fill('000005');await a.locator('#showroomStatus').selectOption('equal');
    const large=await exportRows(),largeRows=large.rows.slice(1),largeCodes=new Set(largeRows.map(row=>row[0]));assert.equal(largeRows.length,1200);assert.equal(largeCodes.size,1200);assert.deepEqual(large.summary,[1200,8000,400]);assert.equal(largeRows.filter(row=>row[5]===0).length,400);assert.equal(largeRows.filter(row=>row[4]===null).length,400);assert.equal(largeRows.filter(row=>row[3]===null).length,400);assert.ok(largeCodes.has('000000')&&largeCodes.has('001199'));
    await a.locator('#showroomSearch').fill('');await a.locator('#showroomStatus').selectOption('all');
    console.log('PASS: complete union of 1200 references above the page/API limits, including all 400 zero differences, 400 physical-only and 400 virtual-only materials, regardless of screen filters.');
    if(process.env.TMS_SHOWROOM_PHYSICAL&&process.env.TMS_SHOWROOM_VIRTUAL){await upload(a,'ct',process.env.TMS_SHOWROOM_PHYSICAL,'2026-10-09');await upload(a,'ad',process.env.TMS_SHOWROOM_VIRTUAL,'2026-11-08');const real=await a.evaluate(()=>TMSShowroom.getState());assert.equal(real.sources.ct.worksheet,'Conteo');assert.equal(real.sources.ct.acceptedRowCount,1690);assert.equal(real.sources.ct.materials.length,1623);assert.equal(real.sources.ct.totals.quantity,1754);assert.equal(real.sources.ad.totals.quantity,4615);assert.equal(real.sources.ad.materials.length,2807);console.log('PASS: real showroom reports: Conteo selected; 1690 physical rows, 1623 materials, 1754 units; 2807 virtual materials, 4615 units.');
      if(process.env.TMS_SHOWROOM_REPORT_OUTPUT){
        await a.locator('#showroomStatus').selectOption('all');
        const next=a.waitForEvent('download');await a.locator('[data-show-inv="export"]').click();const actual=await next;
        await actual.saveAs(process.env.TMS_SHOWROOM_REPORT_OUTPUT);
        const output=require('node:fs').readFileSync(process.env.TMS_SHOWROOM_REPORT_OUTPUT);
        const verified=await a.evaluate(bytes=>{const wb=XLSX.read(new Uint8Array(bytes),{type:'array'}),sheet=wb.Sheets.Conciliacion;return {names:wb.SheetNames,rows:XLSX.utils.sheet_to_json(sheet,{range:11}).length,columns:XLSX.utils.decode_range(sheet['!ref']).e.c+1};},[...output]);
        assert.deepEqual(verified.names,['Conciliacion']);
        assert.equal(verified.rows,3847);assert.equal(verified.columns,6);
        console.log('PASS: full real comparison exported with all 3847 articles and exactly six requested columns.');
      }
    }
    assert.deepEqual(errors,[]);console.log('PASS: showroom CSV/XLSX imports, weighted DOP valuation, consolidation, duplicate snapshots, empty AD, exports, role access, simultaneous sessions, draft preservation, responsive UI.');
  }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

// Synthetic uploads/accounts only. Supabase requests are blocked.
const playwright=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright');
const assert=require('node:assert/strict');
const base=process.env.TMS_TEST_URL||'http://127.0.0.1:8766';
(async()=>{
 const browser=await playwright[process.env.TMS_TEST_ENGINE||'chromium'].launch({headless:true}),errors=[],dialogs=[];
 try{
  const page=await browser.newPage({viewport:{width:390,height:844}});page.on('pageerror',e=>errors.push(e.message));page.on('dialog',d=>{dialogs.push(d.message());d.accept();});
  await page.route('**/*.supabase.co/**',r=>r.fulfill({status:503,contentType:'application/json',body:'{"message":"Blocked in test"}'}));
  await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>!!document.getElementById('mobileModulesButton'));
  assert.equal(await page.locator('#appHeader').isVisible(),false);assert.equal(await page.locator('#mobileQuickNav').isVisible(),false);
  await page.evaluate(()=>{
   const u={id:'11111111-1111-4111-8111-111111111111',email:'mobile@example.invalid'},profile={userId:u.id,authUserId:u.id,email:u.email,nombre:'Mobile',rol:'Admin',activo:true};
   APP.userProfiles=[profile];APP.currentUserProfile=profile;
   APP.clienteConfig={C1:{codigoCliente:'C1',nombreCliente:'Cliente Móvil',rutaAsignada:'SANTO DOMINGO',diasRecepcion:[1,2,3,4,5],horarioAlmacen:'8–16',configuracionCompleta:true}};
   APP.routeConfigs=[{id:'SD',nombre:'SANTO DOMINGO',zona:'SANTO DOMINGO',diasOperacion:[1,2,3,4,5],activa:true}];
   window.guardarEnSupabase=async()=>true;window.scheduleAutoSave=()=>{};mostrarApp(u);cambiarVista('importar',document.getElementById('nav-importar'));
  });
  await page.waitForFunction(()=>document.getElementById('mobileCurrentModule').textContent==='Importar SAP');
  assert.equal(await page.locator('.sap-import-tile').count(),3);assert.equal(await page.locator('.import-excel-table:visible').count(),0);
  assert.ok(await page.locator('#appHeader').evaluate(el=>el.getBoundingClientRect().height<=64));assert.equal(await page.locator('#sapControlUpload').isEnabled(),false);
  assert.equal(await page.locator('#mobileQuickNav button').count(),4);
  // Card buttons still open the original parser inputs; they do not upload to a new service.
  for(const [selector,id]of [['#sapPlanUpload','sapFile'],['[data-sap-upload="warehouse-plan"]','solicitudesPlanFileImport'],['[data-sap-upload="warehouse-control"]','solicitudesControlFileImport']]){
   const chooser=page.waitForEvent('filechooser');await page.locator(selector).click();assert.equal(await (await chooser).element().getAttribute('id'),id);
  }
  const workbook=async rows=>Buffer.from(await page.evaluate(rows=>{const b=XLSX.utils.book_new();XLSX.utils.book_append_sheet(b,XLSX.utils.json_to_sheet(rows),'02.11.2026');return Array.from(new Uint8Array(XLSX.write(b,{type:'array',bookType:'xlsx'})));},rows));
  const sapRow={'Cliente':'C1','Nombre Comercial':'Cliente Móvil','Pedido de cliente':'MOBILE-1','Posición':'1','Producto':'REF-M','Producto descripción':'Material móvil','Cantidad solicitada':10,'Cantidad pendiente de servir':10,'Cantidad facturada':0,'Cantidad pendiente sin stock':2,'Importe confirmado no entregado':1000,'Importe entregado no facturado':50};
  await page.locator('#sapFile').setInputFiles({name:'programacion-02.11.2026.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:await workbook([sapRow])});
  await page.waitForFunction(()=>APP.planLoaded&&document.getElementById('sapSuccessBanner').style.display==='flex');
  assert.equal(await page.locator('#sapSuccessBanner').isVisible(),true);assert.equal(await page.locator('#sapControlUpload').isEnabled(),true);assert.match(await page.locator('#importTablePlanStatus').innerText(),/programacion-02.11.2026.xlsx/);
  const chooser=page.waitForEvent('filechooser');await page.locator('#sapControlUpload').click();assert.equal(await (await chooser).element().getAttribute('id'),'ctrlFile');
  await page.locator('#ctrlFile').setInputFiles({name:'control-02.11.2026.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:await workbook([{...sapRow,'Cantidad facturada':2,'Importe confirmado no entregado':800}])});
  await page.waitForFunction(()=>APP.controlLoaded&&document.getElementById('ctrlSuccessBanner').style.display==='flex');assert.equal(await page.locator('#ctrlSuccessBanner').isVisible(),true);
  const warehouse={'Código':'C1','Cliente':'Cliente Móvil','Solicitud':'W-1','Pedido':'MOBILE-1','Linea':'1','Producto':'REF-M','Cantidad solicitada':5,'Cantidad completada':0,'Fecha de envío planificada':'02/11/2026'};
  for(const [id,name,row]of [['solicitudesPlanFileImport','solicitudes.xlsx',warehouse],['solicitudesControlFileImport','preparado.xlsx',{...warehouse,'Cantidad completada':2}]]){
   await page.locator('#'+id).setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:await workbook([row])});await page.waitForFunction(name=>Object.values(APP.importFiles).includes(name),name);
  }
  assert.match(await page.locator('#importTableSolStatus').innerText(),/solicitudes.xlsx/);assert.match(await page.locator('#importTableSolStatus').innerText(),/preparado.xlsx/);
  // Invalid imports retain previous rows and surface an error rather than silent success.
  const before=await page.evaluate(()=>APP.planLineItems.length);
  await page.locator('#sapFile').setInputFiles({name:'invalid.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:await workbook([{Irrelevante:'x'}])});
  await page.waitForFunction(()=>document.getElementById('sapLoading').style.display==='none');assert.equal(await page.evaluate(()=>APP.planLineItems.length),before);assert.ok(dialogs.some(d=>d.includes('No se pudo procesar')));
  await page.locator('#mobileModulesButton').click();assert.equal(await page.locator('#mobileModulesDialog').isVisible(),true);assert.equal(await page.locator('#mobileModulesDialog [data-mobile-module="rutas"]').count(),0);
  await page.locator('#mobileModulesDialog [data-mobile-module="calendario"]').click();await page.waitForFunction(()=>document.querySelector('.view.active').id==='calendario');assert.equal(await page.locator('#mobileModulesDialog').isVisible(),false);
  await page.waitForFunction(()=>document.getElementById('calContainer').classList.contains('mobile-deliveries-only'));
  const c=page.locator('.cal-card:visible').first();await c.scrollIntoViewIfNeeded();
  const fit=await c.evaluate(el=>{const b=el.getBoundingClientRect();return {left:b.left,right:b.right,width:innerWidth};});assert.ok(fit.left>=0&&fit.right<=fit.width,JSON.stringify(fit));
  assert.ok(await page.locator('.calendar-empty-day').count()>0);assert.equal(await page.locator('.calendar-empty-day:visible').count(),0);
  await page.locator('[data-calendar-focus="all"]').click();assert.ok(await page.locator('.calendar-empty-day:visible').count()>0);await page.locator('[data-calendar-focus="deliveries"]').click();
  await page.evaluate(()=>{APP.transitMaterials=[{id:'M',factura:'TRANSIT-M',codigoMaterial:'REF-M',cantidad:2,eta:'2026-11-08'}];renderCalendario();});
  await page.locator('.cal-card:visible').first().locator('input[type=checkbox]').click();await page.waitForFunction(()=>!document.getElementById('calContainer').classList.contains('calendar-no-selection'));assert.equal(await page.locator('#moveSelectedBtn').isVisible(),true);await page.locator('#clearSelectedBtn').click();await page.waitForFunction(()=>document.getElementById('calContainer').classList.contains('calendar-no-selection'));
  const compact=page.locator('.cal-card:visible').first();await compact.locator('.calendar-card-tools>summary').click();assert.equal(await page.locator('#calendarDispatchDialog').isVisible(),false);await compact.locator('.cal-card-actions button').first().click();assert.equal(await page.locator('#moveModal').evaluate(el=>el.classList.contains('show')),true);await page.evaluate(()=>cerrarMoveModal());
  await compact.scrollIntoViewIfNeeded();await page.screenshot({path:'/tmp/tms-calendar-card-phone.png'});
  await page.locator('.calendar-transit-chip:visible').first().click();await page.waitForFunction(()=>!!document.querySelector('#calendarDispatchDialog td[data-mobile-label]'));
  assert.equal(await page.locator('#calendarDispatchDialog table').evaluate(el=>el.scrollWidth<=el.clientWidth+1),true);await page.screenshot({path:'/tmp/tms-responsive-dispatch-mobile.png'});await page.keyboard.press('Escape');
  await page.locator('#mobileQuickNav [data-mobile-module="importar"]').click();await page.screenshot({path:'/tmp/tms-responsive-import-loaded.png',fullPage:true});
  // Exercise all permitted modules across small phones, tablets and desktop.
  const widths=[320,390,430,768,1024,1440],clippedAll=[];
  for(const width of widths){
   await page.setViewportSize({width,height:900});
   for(const view of ['importar','dashboard','calendario','comercial','inventario','materialesTransito','prioridades','solicitudesAlmacen','incidenciasDespacho','configClientes']){
    await page.evaluate(view=>cambiarVista(view,document.getElementById('nav-'+view)),view);
    assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),`page overflow: ${width}/${view}`);
    const clipped=await page.evaluate(()=>[...document.querySelectorAll('.view.active button,.view.active input:not([type=file]),.view.active select')].filter(el=>el.getClientRects().length&&!el.closest('.table-wrap,.commercial-table-wrap,.transit-table-wrap,.count-tabs,#calMesNav')).filter(el=>{const r=el.getBoundingClientRect();return r.width>0&&(r.left < -1 || r.right > innerWidth+1);}).map(el=>el.id||el.textContent.trim().slice(0,40)));
    if(clipped.length)clippedAll.push({width,view,controls:clipped});
    if(view==='calendario'&&width<=900){const r=await page.locator('.cal-card:visible').first().evaluate(el=>el.getBoundingClientRect().width);assert.ok(r<=width,`clipped calendar ${width}`);}
   }
   if(width<=900){assert.equal(await page.locator('#sidebar').isVisible(),false);assert.equal(await page.locator('#mobileQuickNav').isVisible(),true);}
   else{assert.equal(await page.locator('#sidebar').isVisible(),true);assert.equal(await page.locator('#mobileQuickNav').isVisible(),false);}
  }
  assert.deepEqual(clippedAll,[]);
  await page.setViewportSize({width:844,height:390});await page.evaluate(()=>cambiarVista('importar',document.getElementById('nav-importar')));await page.locator('#mobileModulesButton').click();assert.ok(await page.locator('#mobileModulesDialog').evaluate(el=>el.getBoundingClientRect().height<=innerHeight));await page.keyboard.press('Escape');
  await page.setViewportSize({width:390,height:844});
  await page.evaluate(()=>{APP.currentUserProfile={...APP.currentUserProfile,rol:'Solo lectura'};APP.rolePermissions['Solo lectura'].inventario={ver:false,editar:false,importar:false};cambiarVista('calendario',document.getElementById('nav-calendario'));});
  await page.locator('#mobileModulesButton').click();assert.equal(await page.locator('#mobileModulesDialog [data-mobile-module="inventario"]').count(),0);assert.equal(await page.locator('#mobileQuickNav [data-mobile-module="inventario"]').count(),0);await page.keyboard.press('Escape');
  // For a role allowed to view imports but not edit them, visible upload actions stay disabled.
  await page.evaluate(()=>{APP.rolePermissions['Solo lectura'].importar={ver:true,editar:false,importar:false};cambiarVista('importar',document.getElementById('nav-importar'));});assert.equal(await page.locator('#sapPlanUpload').isEnabled(),false);assert.equal(await page.locator('#sapControlUpload').isEnabled(),false);
  await page.evaluate(()=>{setVisualTheme('dark');cambiarVista('calendario',document.getElementById('nav-calendario'));});await page.screenshot({path:'/tmp/tms-responsive-calendar-dark.png',fullPage:true});
  assert.deepEqual(errors,[]);console.log('PASS: 3 SAP cards, original upload inputs/parsers, visible progress/success, invalid-file preservation, compact phone header, permission-aware module menu and quick nav, calendar fit/day focus, stacked dispatch details, 10 modules at 320/390/430/768/1024/1440px.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});

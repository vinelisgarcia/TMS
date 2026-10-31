// Synthetic data only. All Supabase network traffic is blocked.
const {chromium}=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright');
const assert=require('node:assert/strict');
const base=process.env.TMS_TEST_URL||'http://127.0.0.1:8766';
(async()=>{
 const browser=await chromium.launch({headless:true}),errors=[];
 const id='11111111-1111-4111-8111-111111111111';
 const db={revision:1,tables:Object.fromEntries(['tms_order_lines','tms_client_configs','tms_route_configs','tms_route_cost_history','tms_warehouse_plan','tms_warehouse_control','tms_warehouse_history','tms_settings','tms_user_profiles'].map(name=>[name,[]]))};
 db.tables.tms_user_profiles=[{id:'test-admin',auth_user_id:id,nombre:'Prueba',email:'test@example.invalid',rol:'Admin',activo:true,permisos_por_modulo:{}}];
 let writes=0;
 async function createPage(){
  const page=await browser.newPage({viewport:{width:1440,height:1000}});page.on('pageerror',error=>errors.push(error.message));
  await page.route('**/*.supabase.co/**',route=>route.fulfill({status:503,contentType:'application/json',body:'{"message":"Blocked for test"}'}));
  await page.exposeFunction('__calendarRpc',async(name,args)=>{
   if(name==='tms_get_snapshot')return {data:args.p_known_revision===db.revision?{revision:db.revision,unchanged:true}:{revision:db.revision,tables:structuredClone(db.tables)},error:null};
   assert.equal(name,'tms_save_snapshot');if(args.p_expected_revision!==db.revision)return {data:null,error:{code:'40001',message:'Concurrent update'}};
   for(const [key,value]of Object.entries(args.p_payload))db.tables[key]=structuredClone(value);
   writes++;return {data:{revision:++db.revision,changed:true},error:null};
  });
  await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>!!window.TMSAuditCore&&typeof XLSX!=='undefined'&&typeof Chart!=='undefined');
  await page.evaluate(async id=>{
   const user={id,email:'test@example.invalid'};SB.auth.getUser=async()=>({data:{user},error:null});SB.rpc=(name,args)=>window.__calendarRpc(name,args);SB.channel=()=>({on(type,filter,callback){window.__calendarRemote=callback;return this;},subscribe(){return this;}});SB.removeChannel=async()=>{};
   await cargarDesdeSupabase();mostrarApp(user);cambiarVista('calendario',document.getElementById('nav-calendario'));
  },id);
  return page;
 }
 const card=(page,code='C1')=>page.locator('.cal-card').filter({has:page.locator('.calendar-dispatch-code',{hasText:new RegExp('^'+code+'$')})}).first();
 try{
  const a=await createPage();
  await a.evaluate(async()=>{
   const item=extra=>({clienteId:'C1',clienteNombre:'Cliente Costa',pedidoCliente:'P1',lineaPedidoCliente:'1',articulo:'A',descripcionArticulo:'Material A',cantidadSolicitada:5,cantidadSolicitadaOriginal:8,cantidadPendiente:5,cantidadFacturada:0,cantidadPendienteSinStock:3,montoPlanificado:100,montoEntregadoNoFacturado:20,rutaNombre:'RUTA NORTE',zona:'RUTA NORTE',fechaPlanificada:'02/11/2026',camionAsignado:'CAMION 1',...extra});
   Object.assign(APP,{planLoaded:true,planFecha:'02/11/2026',clienteConfig:{C1:{codigoCliente:'C1',nombreCliente:'Cliente Costa',rutaAsignada:'RUTA NORTE',diasRecepcion:[1,2,3,4,5],horarioAlmacen:'8–16',configuracionCompleta:true},C2:{codigoCliente:'C2',nombreCliente:'Otro Cliente',rutaAsignada:'RUTA NORTE',diasRecepcion:[1,2,3,4,5],horarioAlmacen:'8–16',configuracionCompleta:true}},routeConfigs:[{id:'N',nombre:'RUTA NORTE',zona:'RUTA NORTE',diasOperacion:[1,2,3,4,5],activa:true}],lineItems:[item({}),item({lineaPedidoCliente:'2',articulo:'B',descripcionArticulo:'Material B',montoPlanificado:50,montoEntregadoNoFacturado:5}),item({pedidoCliente:'P2',articulo:'NO_TRANSIT',descripcionArticulo:'Material listo',cantidadPendienteSinStock:0,montoPlanificado:200,montoEntregadoNoFacturado:0}),item({clienteId:'C2',clienteNombre:'Otro Cliente',pedidoCliente:'OTHER_CLIENT_ORDER',articulo:'PRIVATE_C2',descripcionArticulo:'Privado otro cliente',montoPlanificado:9000}),item({pedidoCliente:'OTHER_DATE_ORDER',fechaPlanificada:'03/11/2026',articulo:'PRIVATE_DATE',descripcionArticulo:'Privado otra fecha'})],transitMaterials:[{id:'T1',factura:'F-100',codigoMaterial:'A',descripcion:'Material A',cantidad:4,eta:'2026-11-05'},{id:'T2',factura:'F-100',codigoMaterial:'B',descripcion:'Material B',cantidad:2,eta:'2026-11-05'},{id:'T3',factura:'F-200',codigoMaterial:'A',descripcion:'Material A',cantidad:1,eta:'2026-11-12'},{id:'T4',factura:'RECEIVED',codigoMaterial:'A',cantidad:5,eta:'2026-10-01',receivedAt:'2026-10-31'},{id:'T5',factura:'PRIVATE',codigoMaterial:'A',cantidad:20,cliente:'C2'},{id:'T6',factura:'PRIVATE_DATE',codigoMaterial:'PRIVATE_DATE',cantidad:2,eta:'2026-11-09'}]});
   APP.planLineItems=APP.lineItems.map(row=>({...row}));TMSAuditCore.rebuildDerivedState();renderCalendario();await guardarEnSupabase();
  });
  assert.equal(writes,1);
  const c=card(a);assert.equal(await c.locator('.calendar-dispatch-metrics div').count(),4);
  assert.deepEqual(await c.locator('dd').allTextContents(),['2','RD$ 350','RD$ 25','RD$ 375']);
  assert.equal(await c.locator('.calendar-transit-chip').count(),2);assert.match(await c.locator('.calendar-transit-chip').first().innerText(),/^Factura en tránsito: F-100 · ETA: 2026-11-05$/);
  assert.doesNotMatch(await c.innerText(),/P1|Material A|sin reserva|PRIVATE|RECEIVED/);
  await c.scrollIntoViewIfNeeded();await a.screenshot({path:'/tmp/tms-calendar-compact.png'});
  const chip=c.locator('.calendar-transit-chip').first();await chip.click();
  const modal=a.locator('#calendarDispatchDialog');assert.equal(await modal.isVisible(),true);assert.equal(await modal.locator('tbody tr').count(),2);
  assert.match(await modal.innerText(),/P1/);assert.doesNotMatch(await modal.innerText(),/P2|OTHER_CLIENT_ORDER|OTHER_DATE_ORDER|Material listo|F-200|PRIVATE_C2|PRIVATE_DATE|RECEIVED/);
  await a.screenshot({path:'/tmp/tms-calendar-transit-detail.png'});await a.keyboard.press('Escape');assert.equal(await modal.isVisible(),false);
  await c.locator('.calendar-dispatch-name').click();assert.equal(await modal.locator('.calendar-dispatch-orders tbody tr').count(),3);assert.match(await modal.innerText(),/P2/);assert.doesNotMatch(await modal.innerText(),/OTHER_CLIENT_ORDER|OTHER_DATE_ORDER|PRIVATE_C2|PRIVATE_DATE/);
  await a.screenshot({path:'/tmp/tms-calendar-dispatch-detail.png'});await modal.locator('[data-calendar-detail-close]').click();
  await c.focus();await a.keyboard.press('Enter');assert.equal(await modal.isVisible(),true);await a.keyboard.press('Escape');
  await c.locator('input[type=checkbox]').click();assert.equal(await modal.isVisible(),false);assert.equal(await c.locator('input[type=checkbox]').isChecked(),true);
  await c.locator('.cal-card-actions button').first().click();assert.equal(await a.locator('#moveModal').evaluate(el=>el.classList.contains('show')),true);assert.equal(await modal.isVisible(),false);await a.evaluate(()=>cerrarMoveModal());
  const b=await createPage();assert.deepEqual(await card(b).locator('dd').allTextContents(),['2','RD$ 350','RD$ 25','RD$ 375']);
  await card(b).locator('.calendar-transit-chip').first().click();
  await a.evaluate(async()=>{APP.transitMaterials.filter(row=>row.factura==='F-100').forEach(row=>row.eta='2026-11-07');APP.lineItems.find(row=>row.pedidoCliente==='P1').montoPlanificado=125;TMSAuditCore.rebuildDerivedState();renderCalendario();await guardarEnSupabase();});
  await b.evaluate(revision=>window.__calendarRemote({new:{revision}}),db.revision);
  await b.waitForFunction(()=>document.querySelector('#calendarDispatchDialog').textContent.includes('2026-11-07'));
  assert.match(await b.locator('#calendarDispatchDialog').innerText(),/2026-11-07/);assert.doesNotMatch(await b.locator('#calendarDispatchDialog').innerText(),/2026-11-05/);
  await b.keyboard.press('Escape');assert.equal(await card(b).locator('dd').last().innerText(),'RD$ 400');
  await card(b).locator('.calendar-transit-chip').first().click();
  await a.evaluate(async()=>{APP.transitMaterials=TMSDomain.receiveInvoice(APP.transitMaterials,'F-100',true,'tester','2026-10-31T12:00:00Z');await guardarEnSupabase();});
  await b.evaluate(revision=>window.__calendarRemote({new:{revision}}),db.revision);
  await b.waitForFunction(()=>document.querySelector('#calendarDispatchDialog').textContent.includes('ya no afecta'));
  assert.equal(await b.locator('#calendarDispatchDialog tbody tr').count(),0);await b.keyboard.press('Escape');assert.equal(await card(b).locator('.calendar-transit-chip').count(),1);
  const beforeRead=writes;await b.evaluate(()=>{APP.currentUserProfile={...APP.currentUserProfile,rol:'Solo lectura'};renderCalendario();});await card(b).locator('.calendar-dispatch-name').click();assert.equal(await b.locator('#calendarDispatchDialog').isVisible(),true);assert.equal(writes,beforeRead);await b.keyboard.press('Escape');
  for(const width of [390,768]){await b.setViewportSize({width,height:844});await card(b).scrollIntoViewIfNeeded();assert.equal(await b.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),true);await card(b).locator('.calendar-transit-chip').first().click();assert.equal(await b.locator('#calendarDispatchDialog').evaluate(el=>el.getBoundingClientRect().right<=innerWidth),true);if(width===390)await b.screenshot({path:'/tmp/tms-calendar-compact-mobile.png'});await b.keyboard.press('Escape');}
  assert.deepEqual(errors,[]);console.log('PASS: compact four-field cards, invoice/ETA grouping, scoped popups, keyboard and route controls, two synced sessions, live popup refresh, receipts, read-only and mobile.');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

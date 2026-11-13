// Real frontend and sync controller, synthetic accounts/RPC only. Production traffic is blocked.
const playwright=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright'),assert=require('node:assert/strict'),fs=require('node:fs');
const base=process.env.TMS_TEST_URL||'http://127.0.0.1:8766';
const domain=require('../tms-domain.js');
(async()=>{
 const browser=await playwright[process.env.TMS_TEST_ENGINE||'chromium'].launch({headless:true}),errors=[];
 const admin='11111111-1111-4111-8111-111111111111',reader='22222222-2222-4222-8222-222222222222';
 const db={revision:1,tables:Object.fromEntries(['tms_order_lines','tms_client_configs','tms_route_configs','tms_route_cost_history','tms_warehouse_plan','tms_warehouse_control','tms_warehouse_history','tms_settings','tms_user_profiles'].map(name=>[name,[]]))};let writes=0;
 db.tables.tms_user_profiles=[{id:admin,auth_user_id:admin,nombre:'Contabilidad',email:'accounting@example.invalid',rol:'Admin',activo:true,permisos_por_modulo:{}},{id:reader,auth_user_id:reader,nombre:'Consulta',email:'reader@example.invalid',rol:'Consulta tránsito',activo:true,permisos_por_modulo:{}}];
 db.tables.tms_settings=[{clave:'role_permissions',valor:{'Consulta tránsito':{materialesTransito:{ver:true,editar:false,importar:false}}}}];
 async function boot(page,readonly=false){await page.evaluate(async({id,email})=>{
  const user={id,email};SB.auth.getUser=async()=>({data:{user},error:null});SB.rpc=(name,args)=>window.__transitRpc(name,args);SB.channel=()=>({on(type,filter,callback){window.__transitRemote=callback;return this;},subscribe(){return this;}});SB.removeChannel=async()=>{};
  await cargarDesdeSupabase();mostrarApp(user);cambiarVista('materialesTransito',document.getElementById('nav-materialesTransito'));
 },{id:readonly?reader:admin,email:readonly?'reader@example.invalid':'accounting@example.invalid'});}
 async function pageFor(readonly=false){
  const page=await browser.newPage({viewport:{width:1440,height:1000}});page.on('pageerror',error=>errors.push(error.message));page.on('dialog',dialog=>dialog.accept());
  await page.route('**/*.supabase.co/**',route=>route.fulfill({status:503,body:'Blocked in test'}));
  await page.exposeFunction('__transitRpc',async(name,args)=>{
   if(name==='tms_get_snapshot')return {data:args.p_known_revision===db.revision?{revision:db.revision,unchanged:true}:{revision:db.revision,tables:structuredClone(db.tables)},error:null};
   assert.equal(name,'tms_save_snapshot');if(readonly)return {data:null,error:{code:'42501',message:'Read only'}};
   if(args.p_expected_revision!==db.revision)return {data:null,error:{code:'40001',message:'Concurrent update'}};
   for(const [key,value]of Object.entries(args.p_payload)){
    // The real RPC upserts supplied setting keys, preserving unchanged role settings.
    if(key==='tms_settings')for(const row of value){const index=db.tables[key].findIndex(item=>item.clave===row.clave);if(index<0)db.tables[key].push(structuredClone(row));else db.tables[key][index]=structuredClone(row);}
    else db.tables[key]=structuredClone(value);
   }writes++;return {data:{revision:++db.revision,changed:true},error:null};
  });
  await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>!!window.TMSAuditCore&&!!window.TMSDomain&&!!window.XLSX);await boot(page,readonly);
  assert.equal(await page.locator('#materialesTransito').isVisible(),true,JSON.stringify(await page.evaluate(()=>({role:APP.currentUserProfile?.rol,permissions:APP.rolePermissions[APP.currentUserProfile?.rol],view:document.querySelector('.view.active')?.id}))));return page;
 }
 const rows=()=>db.tables.tms_settings.find(row=>row.clave==='app_state').valor.transitMaterials;
 const tab=(page,value)=>page.locator('[data-audit="transit-tab"][data-value="'+value+'"]');
 const action=(page,kind,value)=>page.locator('[data-audit="'+kind+'"][data-invoice="F-100"][data-value="'+value+'"]');
 const flush=page=>page.evaluate(()=>guardarEnSupabase());
 async function remote(page){await page.evaluate(revision=>window.__transitRemote({new:{revision}}),db.revision);}
 try{
  const a=await pageFor();await a.evaluate(async()=>{
   const line={clienteId:'C1',clienteNombre:'Cliente de prueba',pedidoCliente:'ORDER-1',lineaPedidoCliente:'1',articulo:'0007',descripcionArticulo:'Material de prueba',cantidadSolicitada:5,cantidadSolicitadaOriginal:5,cantidadPendiente:5,cantidadFacturada:0,cantidadPendienteServir:5,cantidadPendienteSinStock:2,montoPlanificado:100,montoEntregadoNoFacturado:0,rutaNombre:'NORTE',zona:'NORTE',fechaPlanificada:'13/11/2026',camionAsignado:'CAMION 1'};
   Object.assign(APP,{planLoaded:true,planFecha:'13/11/2026',lineItems:[line],planLineItems:[{...line}],clienteConfig:{C1:{codigoCliente:'C1',nombreCliente:'Cliente de prueba',rutaAsignada:'NORTE',diasRecepcion:[1,2,3,4,5],configuracionCompleta:true}},routeConfigs:[{id:'N',nombre:'NORTE',zona:'NORTE',diasOperacion:[1,2,3,4,5],activa:true}],transitMaterials:[{id:'A',factura:'F-100',codigoMaterial:'0007',descripcion:'Plato de prueba',cantidad:5,eta:'2026-11-20',cliente:''},{id:'B',factura:'F-100',codigoMaterial:'B',descripcion:'Copa de prueba',cantidad:2,eta:'2026-11-20',cliente:''},{id:'OTHER',factura:'F100',codigoMaterial:'OTHER',descripcion:'Otra factura',cantidad:9,eta:'2026-11-20',cliente:''}]});
   TMSAuditCore.rebuildDerivedState();renderMaterialesTransito();renderCalendario();await guardarEnSupabase();
  });
  assert.equal(await tab(a,'accounted').count(),1);assert.equal(await a.locator('.audit-invoice').count(),2);assert.equal(await a.locator('[data-audit="account-invoice"]').count(),0);assert.equal(await a.locator('.calendar-transit-chip').count(),1);
  const b=await pageFor();await action(a,'receive-invoice','receive').click();await flush(a);const receipt=structuredClone(rows().filter(row=>row.factura==='F-100'));assert.ok(receipt.every(row=>row.receivedAt&&row.receivedBy==='accounting@example.invalid'&&!row.accountedAt));assert.equal(await a.locator('.calendar-transit-chip').count(),0);
  await remote(b);await b.waitForFunction(()=>APP.transitMaterials.filter(row=>row.factura==='F-100').every(row=>TMSDomain.transitState(row)==='received'));await tab(b,'received').click();assert.equal(await action(b,'account-invoice','account').count(),1);
  await tab(a,'received').click();await action(a,'account-invoice','account').click();await flush(a);
  const posted=structuredClone(rows().filter(row=>row.factura==='F-100'));assert.ok(posted.every((row,i)=>row.estado==='contabilizada'&&row.accountedAt&&row.accountedBy==='accounting@example.invalid'&&row.receivedAt===receipt[i].receivedAt&&row.receivedBy===receipt[i].receivedBy&&row.statusHistory.length===2));assert.equal(rows().find(row=>row.factura==='F100').estado,undefined);
  await remote(b);await b.waitForFunction(()=>APP.transitMaterials.filter(row=>row.factura==='F-100').every(row=>TMSDomain.transitState(row)==='accounted'));assert.equal(await b.locator('.audit-invoice').count(),0);await tab(b,'accounted').click();assert.match(await b.locator('.audit-invoice').innerText(),/Recepción:|Contabilización de la factura en SAP:/);assert.equal(await b.locator('.calendar-transit-chip').count(),0);
  await b.reload({waitUntil:'networkidle'});await boot(b);await tab(b,'accounted').click();assert.equal(await b.locator('.audit-invoice').count(),1);assert.equal(await b.locator('[data-audit="receive-invoice"]').count(),0);
  // Import updates quantities without undoing either confirmation or its history.
  await tab(a,'accounted').click();await a.locator('#transitFile').setInputFiles({name:'same-invoice.csv',mimeType:'text/csv',buffer:Buffer.from('Referencia,Descripcion,Cantidad,Factura,ETA\n0007,Plato actualizado,6,F-100,2026-11-20\n')});await a.locator('button[onclick="importarMaterialesTransito()"]').click();await a.waitForFunction(()=>document.getElementById('transitFile').files.length===0);await flush(a);const updated=rows().find(row=>row.id==='A');assert.equal(updated.cantidad,6);assert.equal(updated.estado,'contabilizada');assert.equal(updated.receivedAt,receipt[0].receivedAt);assert.equal(updated.accountedAt,posted[0].accountedAt);assert.deepEqual(updated.statusHistory,posted[0].statusHistory);
  const event=a.waitForEvent('download');await a.locator('button[onclick="exportarMaterialesTransitoExcel()"]').click();const file=await event;const bytes=fs.readFileSync(await file.path());const excel=await a.evaluate(bytes=>{const book=XLSX.read(new Uint8Array(bytes),{type:'array'});return XLSX.utils.sheet_to_json(book.Sheets['Materiales transito']);},[...bytes]);const exported=excel.find(row=>row.Material==='0007');assert.equal(exported.Estado,'Contabilizado');assert.equal(exported['Fecha recepción'],receipt[0].receivedAt);assert.equal(exported['Fecha contabilización SAP'],posted[0].accountedAt);assert.equal(exported.Cantidad,6);
  for(const width of [320,390,768]){await a.setViewportSize({width,height:900});assert.ok(await a.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1));assert.ok(await a.locator('.audit-tabs').evaluate(el=>el.scrollWidth<=el.clientWidth+1));if(width===390)await a.screenshot({path:'/tmp/tms-transit-accounted-mobile.png',fullPage:true});}
  const readonly=await pageFor(true);await tab(readonly,'accounted').click();assert.equal(await readonly.locator('[data-audit="account-invoice"],[data-audit="receive-invoice"]').count(),0);const before=writes;await readonly.evaluate(()=>{const button=document.createElement('button');button.dataset.audit='account-invoice';button.dataset.invoice='F-100';button.dataset.value='revert';document.body.appendChild(button);button.click();button.remove();});assert.equal(writes,before);
  await action(a,'account-invoice','revert').click();await flush(a);assert.ok(rows().filter(row=>row.factura==='F-100').every((row,i)=>domain.transitState(row)==='received'&&row.receivedAt===receipt[i].receivedAt&&!row.accountedAt&&row.statusHistory.length===3));assert.equal(await a.locator('.calendar-transit-chip').count(),0);await tab(a,'received').click();assert.equal(await action(a,'account-invoice','account').count(),1);
  assert.deepEqual(errors,[]);console.log('PASS: received then SAP-accounted invoice, distinct dates/actors, all invoice lines, exact invoice identity, calendar alert removal, two synchronized sessions/reload, safe reimport, Excel metadata, 320/390/768px, read-only and accounting reversal without reopening transit.');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

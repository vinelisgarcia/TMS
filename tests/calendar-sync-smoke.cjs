/* Synthetic two-device cloud; never contacts or writes the hosted database. */
const {chromium} = require(process.env.TMS_PLAYWRIGHT || 'playwright');
const assert = require('node:assert/strict');
const clone = value => structuredClone(value);
const user = {id:'00000000-0000-4000-8000-000000000001',email:'test@example.invalid'};
const line = (client,order) => ({id:order,cliente_id:client,cliente_nombre:client==='C1'?'Hotel Costa Norte':'Otro cliente',pedido_cliente:order,linea_pedido_cliente:'1',articulo:'REF-1',descripcion_articulo:'Material de prueba',ruta_nombre:'PP',fecha_planificada:'2026-10-23',camion_asignado:'CAMION 1',cantidad_solicitada:10,cantidad_facturada:4,cantidad_pendiente:6,metadata:{dataset:'current',cantidadPendienteSinStock:6,montoPlanificado:18000,montoEntregadoNoFacturado:1200}});
let remote = {revision:1,tables:{
  tms_order_lines:[line('C1','P-001'),line('C2','P-002')],
  tms_client_configs:[{codigo_cliente:'C1',nombre_cliente:'Hotel Costa Norte',ruta_asignada:'PP',dias_recepcion:[1,2,3,4,5],configuracion_completa:true}],
  tms_route_configs:[{nombre:'PP',dias_operacion:[1,2,3,4,5],activa:true}],
  tms_user_profiles:[{id:user.id,auth_user_id:user.id,email:user.email,nombre:'Prueba',rol:'Admin',activo:true}],
  tms_settings:[{clave:'app_state',valor:{planLoaded:true,planFecha:'23/10/2026',costSettings:{actualInvoices:[{date:'2026-10-23',truck:'CAMION 1',trip:'1',customer:'C1',order:'P-001',invoice:'SAP-1',currency:'DOP',amount:2500},{date:'2026-10-23',truck:'CAMION 1',trip:'1',customer:'C2',order:'P-002',invoice:'SAP-OTHER',currency:'DOP',amount:3000}]},transitMaterials:[{id:'T1',factura:'F-100',codigoMaterial:'REF-1',cantidad:4,eta:'2026-11-05'},{id:'T2',factura:'F-100',codigoMaterial:'REF-1',cantidad:3,eta:'2026-11-12'},{id:'T3',factura:'RECIBIDA',codigoMaterial:'REF-1',cantidad:100,eta:'2026-11-01',receivedAt:'yesterday'}]}}]
}};
let saves=0,offline=false,migrationMissing=false;
(async()=>{
  const browser=await chromium.launch({headless:true,...(process.env.TMS_BROWSER?{executablePath:process.env.TMS_BROWSER}:{})});
  try {
    async function device(expected='synced',options={}){
      const context=options.context || await browser.newContext({viewport:{width:1440,height:1000}});const page=await context.newPage();
      if(options.inheritedTabId)await page.addInitScript(id=>{if(!sessionStorage.getItem('tms_sync_tab'))sessionStorage.setItem('tms_sync_tab',id);},options.inheritedTabId);
      if(options.legacy)await page.addInitScript(snapshot=>localStorage.setItem('tms_alvarez_v2_state',JSON.stringify(snapshot)),options.legacy);
      await require('./browser-fixtures.cjs').install(page);
      await page.route('**/*.supabase.co/**',async route=>{
        const url=route.request().url();
        if(offline)return route.fulfill({status:503,contentType:'application/json',body:'{"message":"Test offline"}'});
        if(url.includes('/rpc/tms_read_snapshot'))return migrationMissing ? route.fulfill({status:404,json:{code:'PGRST202',message:'Could not find the function'}}) : route.fulfill({json:clone(remote)});
        if(url.includes('/rpc/tms_save_snapshot')){
          const body=route.request().postDataJSON();
          if(body.p_expected_revision!==remote.revision)return route.fulfill({status:409,json:{code:'40001',message:'TMS_CONFLICT: otra sesión'}});
          remote.tables={...remote.tables,...clone(body.p_tables)};
          for(const setting of body.p_settings){
            remote.tables.tms_settings=remote.tables.tms_settings.filter(row=>row.clave!==setting.clave);
            remote.tables.tms_settings.push(clone(setting));
          }
          remote.revision++;saves++;return route.fulfill({json:remote.revision});
        }
        if(url.includes('/tms_sync_state'))return route.fulfill({json:{revision:remote.revision}});
        const parsed=new URL(url),table=parsed.pathname.split('/').at(-1);
        if(migrationMissing && table in remote.tables){
          const offset=Number(parsed.searchParams.get('offset')||0),limit=Number(parsed.searchParams.get('limit')||500);
          return route.fulfill({json:clone(remote.tables[table].slice(offset,offset+limit))});
        }
        if(migrationMissing && table.startsWith('tms_'))return route.fulfill({json:[]});
        return route.fulfill({status:503,json:{message:'Unexpected synthetic request'}});
      });
      const errors=[];page.on('pageerror',e=>errors.push(e.message));
      await page.goto(process.env.TMS_TEST_URL || 'http://127.0.0.1:8765',{waitUntil:'networkidle'});
      await page.evaluate(u=>{SB.auth.getUser=async()=>({data:{user:u},error:null});mostrarApp(u);},user);
      await page.evaluate(()=>cargarDesdeSupabase());
      await page.locator(`#cloudSyncPanel[data-state="${expected}"]`).waitFor();
      assert.equal(errors.length,0,errors.join('\n'));
      return page;
    }
    const a=await device(),b=await device();assert.equal(saves,0,'Loading must not autosave');
    await a.evaluate(()=>cambiarVista('calendario'));
    const card=a.locator('[data-calendar-group]').filter({hasText:'Hotel Costa Norte'}).first();
    assert.equal(await card.locator('.calendar-transit-chip').count(),2,'Each invoice ETA is represented once');
    assert.doesNotMatch(await card.locator('.calendar-transit-chips').innerText(),/uds|cobertura|RECIBIDA/);
    await card.locator('.calendar-transit-chip').first().click();
    const dialog=a.locator('#calendarDetailDialog');
    await dialog.waitFor({state:'visible'});assert.match(await dialog.innerText(),/sin reserva confirmada/);
    assert.match(await dialog.innerText(),/05\/11\/2026/);assert.match(await dialog.innerText(),/12\/11\/2026/);
    await a.keyboard.press('Escape');assert.equal(await dialog.isVisible(),false);
    await card.locator('.cal-card-name').click();await dialog.waitFor({state:'visible'});
    assert.match(await dialog.innerText(),/P-001/);assert.doesNotMatch(await dialog.innerText(),/P-002/);
    assert.match(await dialog.innerText(),/Entregado no facturado/);
    assert.match(await dialog.innerText(),/SAP-1/);assert.doesNotMatch(await dialog.innerText(),/SAP-OTHER/);
    await a.screenshot({path:'/tmp/tms-calendar-detail-desktop.png'});
    await a.setViewportSize({width:390,height:844});
    assert.ok(await dialog.evaluate(el=>el.getBoundingClientRect().width<=innerWidth));
    await a.screenshot({path:'/tmp/tms-calendar-detail-mobile.png'});
    await dialog.locator('[data-calendar-close]').click();
    await a.setViewportSize({width:1440,height:1000});
    await card.focus();await a.keyboard.press('Enter');await dialog.waitFor({state:'visible'});await a.keyboard.press('Escape');
    await card.locator('input[type="checkbox"]').click();assert.equal(await dialog.isVisible(),false);
    await a.evaluate(async()=>{APP.calendarNotes=[{id:'cross-device',fecha:'23/10/2026',texto:'Desde dispositivo A'}];await guardarEnSupabase({silent:true});});
    assert.equal(saves,1);
    await b.evaluate(()=>window.dispatchEvent(new Event('focus')));
    await b.waitForFunction(()=>APP.calendarNotes.some(row=>row.id==='cross-device'));
    assert.equal(saves,1,'Refreshing another device must not save');
    await b.evaluate(()=>{APP.calendarNotes.push({id:'pending-b',fecha:'23/10/2026',texto:'Pendiente B'});scheduleAutoSave();clearTimeout(window._autoSaveTimer);});
    await a.evaluate(async()=>{APP.calendarNotes.push({id:'latest-a',fecha:'23/10/2026',texto:'Último A'});await guardarEnSupabase({silent:true});});
    const result=await b.evaluate(async()=>{try{await guardarEnSupabase({silent:true});return 'unexpected';}catch(e){return e.message;}});
    assert.match(result,/TMS_CONFLICT/);await b.locator('#cloudSyncPanel[data-state="conflict"]').waitFor();
    assert.ok(await b.evaluate(()=>JSON.parse(localStorage.getItem('tms_alvarez_v2_state')).pendingCloud));
    assert.ok(remote.tables.tms_settings.find(row=>row.clave==='app_state').valor.calendarNotes.some(row=>row.id==='latest-a'));
    assert.ok(!remote.tables.tms_settings.find(row=>row.clave==='app_state').valor.calendarNotes.some(row=>row.id==='pending-b'));
    offline=true;
    const failure=await a.evaluate(async()=>{APP.calendarNotes.push({id:'offline-a'});try{await guardarEnSupabase({silent:true});return 'unexpected';}catch(e){return e.message;}});
    assert.match(failure,/offline/i);assert.ok(await a.evaluate(()=>JSON.parse(localStorage.getItem('tms_alvarez_v2_state')).pendingCloud));
    offline=false;
    // A reload recovers this tab's pending version, preserving the conflict.
    b.on('dialog',dialog=>dialog.accept());
    await b.reload({waitUntil:'networkidle'});
    await b.evaluate(async u=>{SB.auth.getUser=async()=>({data:{user:u}});mostrarApp(u);await cargarDesdeSupabase();},user);
    await b.locator('#cloudSyncPanel[data-state="conflict"]').waitFor();
    assert.ok(await b.evaluate(()=>APP.calendarNotes.some(row=>row.id==='pending-b')));
    await b.reload({waitUntil:'networkidle'});
    await b.evaluate(async u=>{SB.auth.getUser=async()=>({data:{user:u}});mostrarApp(u);await cargarDesdeSupabase();},{id:'00000000-0000-4000-8000-000000000099',email:'other@example.invalid'});
    assert.equal(await b.evaluate(()=>APP.calendarNotes.some(row=>row.id==='pending-b')),false,'Another account must not inherit the pending backup');
    // Storage quota must never prevent the remote save from running.
    await a.evaluate(async()=>{
      const original=Storage.prototype.setItem;
      Storage.prototype.setItem=function(key,value){if(key.startsWith('tms_alvarez_v2_state'))throw new DOMException('Test quota','QuotaExceededError');return original.call(this,key,value);};
      try{APP.calendarNotes.push({id:'quota-a'});await guardarEnSupabase({silent:true});}finally{Storage.prototype.setItem=original;}
    });
    await a.locator('#cloudSyncPanel[data-state="synced"]').waitFor();
    assert.ok(remote.tables.tms_settings.find(row=>row.clave==='app_state').valor.calendarNotes.some(row=>row.id==='quota-a'));
    assert.match(await a.locator('#cloudSyncPanel').innerText(),/almacenamiento/);
    // Same-browser duplicate tabs inherit sessionStorage; they still need distinct backup slots.
    const parent=await device();
    const inheritedTabId=await parent.evaluate(()=>sessionStorage.getItem('tms_sync_tab'));
    const child=await device('synced',{context:parent.context(),inheritedTabId});
    assert.notEqual(await child.evaluate(()=>sessionStorage.getItem('tms_sync_tab')),inheritedTabId);
    await parent.evaluate(()=>{APP.calendarNotes.push({id:'parent-unsent',fecha:'23/10/2026',texto:'Pendiente en pestaña original'});scheduleAutoSave();clearTimeout(window._autoSaveTimer);});
    await child.evaluate(async()=>{APP.calendarNotes.push({id:'child-saved',fecha:'23/10/2026',texto:'Guardado desde la pestaña duplicada'});await guardarEnSupabase({silent:true});});
    const ownKey='tms_alvarez_v2_state:'+user.id+':'+inheritedTabId;
    assert.ok(await parent.evaluate(key=>JSON.parse(localStorage.getItem(key)).app.calendarNotes.some(row=>row.id==='parent-unsent'),ownKey));
    parent.on('dialog',dialog=>dialog.accept());
    await parent.reload({waitUntil:'networkidle'});
    await parent.evaluate(async u=>{SB.auth.getUser=async()=>({data:{user:u}});mostrarApp(u);await cargarDesdeSupabase();},user);
    await parent.locator('#cloudSyncPanel[data-state="conflict"]').waitFor();
    assert.equal(await parent.evaluate(()=>sessionStorage.getItem('tms_sync_tab')),inheritedTabId,'Reload must reclaim the same tab slot');
    assert.ok(await parent.evaluate(()=>APP.calendarNotes.some(row=>row.id==='parent-unsent')));
    // Pre-upgrade backups have no owner metadata. Archive them, never silently import or erase them.
    const oldSnapshot={savedAt:'2026-10-01T12:00:00Z',app:{calendarNotes:[{id:'legacy-unsent'}]}};
    const upgraded=await device('synced',{legacy:oldSnapshot});
    assert.equal(await upgraded.evaluate(()=>APP.calendarNotes.some(row=>row.id==='legacy-unsent')),false);
    const archived=await upgraded.evaluate(()=>localStorage.getItem(localStorage.getItem('tms_alvarez_v2_state:legacy_latest')));
    assert.deepEqual(JSON.parse(archived),oldSnapshot);
    const oldDownload=upgraded.waitForEvent('download');await upgraded.locator('#cloudLegacyBackup').click();
    const bytes=require('node:fs').readFileSync(await (await oldDownload).path());assert.deepEqual(JSON.parse(bytes),oldSnapshot);
    // Missing migration remains readable (including pagination), but cannot write.
    migrationMissing=true;
    remote.tables.tms_order_lines=Array.from({length:1205},(_,i)=>({...line('C1','BIG-'+i),id:'big-'+i}));
    const legacy=await device('migration');
    assert.equal(await legacy.evaluate(()=>APP.lineItems.length),1205);
    const before=saves;
    assert.match(await legacy.evaluate(async()=>{try{await guardarEnSupabase({silent:true});return 'unexpected';}catch(e){return e.message;}}),/0008/);
    assert.equal(saves,before);
    console.log('Calendar and sync passed: scoped details, keyboard/mobile, cross-device refresh, offline/conflict recovery, copied-tab identity isolation and legacy-backup preservation/export.');
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

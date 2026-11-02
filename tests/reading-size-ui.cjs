// Synthetic accounts only. No production calls or cloud writes.
const playwright=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright'),assert=require('node:assert/strict');
const base=process.env.TMS_TEST_URL||'http://127.0.0.1:8766';
(async()=>{
 const browser=await playwright[process.env.TMS_TEST_ENGINE||'chromium'].launch({headless:true}),context=await browser.newContext({viewport:{width:390,height:844}}),errors=[];
 async function boot(page,user='reading-user-one'){
  await page.evaluate(user=>{
   const identity={id:user,email:user+'@example.invalid'};APP.currentUserProfile={userId:user,email:identity.email,rol:'Admin',activo:true};APP.userProfiles=[APP.currentUserProfile];
   window.__readingWrites=0;window.guardarEnSupabase=async()=>{window.__readingWrites++;return true;};window.scheduleAutoSave=()=>{window.__readingWrites++;};
   APP.clienteConfig={C1:{codigoCliente:'C1',nombreCliente:'Cliente de lectura',rutaAsignada:'SANTO DOMINGO',diasRecepcion:[1,2,3,4,5],horarioAlmacen:'8–16',configuracionCompleta:true}};APP.routeConfigs=[{id:'SD',nombre:'SANTO DOMINGO',zona:'SANTO DOMINGO',diasOperacion:[1,2,3,4,5],activa:true}];
   APP.lineItems=[{clienteId:'C1',clienteNombre:'Cliente de lectura',pedidoCliente:'READ-1',lineaPedidoCliente:'1',articulo:'TEST',cantidadSolicitada:10,cantidadSolicitadaOriginal:10,cantidadFacturada:0,cantidadPendiente:10,cantidadPendienteServir:10,cantidadPendienteSinStock:0,montoPlanificado:12345678.9,montoEntregadoNoFacturado:0,fechaPlanificada:'02/11/2026',camionAsignado:'CAMION 1',rutaNombre:'SANTO DOMINGO',zona:'SANTO DOMINGO'}];APP.planLineItems=APP.lineItems.map(row=>({...row}));APP.planLoaded=true;APP.planFecha='02/11/2026';TMSAuditCore.rebuildDerivedState();mostrarApp(identity);cambiarVista('importar',document.getElementById('nav-importar'));
  },user);await page.waitForFunction(()=>!!document.getElementById('mobileReadingSizeButton'));
 }
 async function page(){const p=await context.newPage();p.on('pageerror',e=>errors.push(e.message));await p.route('**/*.supabase.co/**',r=>r.fulfill({status:503,contentType:'application/json',body:'{"message":"Blocked for test"}'}));await p.goto(base,{waitUntil:'networkidle'});await p.waitForFunction(()=>!!window.TMSAuditCore&&!!document.getElementById('mobileReadingSizeButton'));await boot(p);return p;}
 const font=(p,selector)=>p.locator(selector).first().evaluate(el=>parseFloat(getComputedStyle(el).fontSize));
 async function choose(p,size){const opener=await p.locator('#mobileReadingSizeButton').isVisible()?'#mobileReadingSizeButton':'#readingSizeButton';await p.locator(opener).click();await p.locator('#readingSizeDialog button[data-reading-size="'+size+'"]').click();await p.locator('[data-reading-close]').click();assert.equal(await p.locator('#readingSizeDialog').isVisible(),false);await p.waitForFunction(size=>document.body.dataset.readingSize===size,size);}
 try{
  const a=await page();await a.waitForFunction(()=>document.body.dataset.readingSize==='large');
  const originalData=await a.evaluate(()=>JSON.stringify(APP.lineItems));await a.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));const large=await font(a,'.sap-import-tile h2');
  await choose(a,'normal');const normal=await font(a,'.sap-import-tile h2');await choose(a,'extra');const extra=await font(a,'.sap-import-tile h2');assert.ok(normal<large&&large<extra);
  await a.evaluate(()=>{const n=document.createElement('div');n.id='readingDynamic';n.style.fontSize='10px';n.innerHTML='Padre <span id="readingChild">Texto nuevo</span>';document.querySelector('#importar').appendChild(n);});await a.waitForFunction(()=>document.getElementById('readingChild').hasAttribute('data-tms-reading'));const first=await font(a,'#readingChild');
  for(let i=0;i<4;i++){await a.evaluate(()=>{const n=document.createElement('span');n.textContent=' actualizado';document.getElementById('readingDynamic').appendChild(n);});await a.waitForFunction(()=>[...document.querySelectorAll('#readingDynamic span')].every(n=>n.hasAttribute('data-tms-reading')));assert.ok(Math.abs(await font(a,'#readingChild')-first)<0.01);}
  assert.ok(Math.abs(await font(a,'#readingDynamic')-first)<0.01);
  const b=await page();await b.waitForFunction(()=>document.body.dataset.readingSize==='extra');await choose(a,'normal');await b.waitForFunction(()=>document.body.dataset.readingSize==='normal');await choose(a,'extra');await b.waitForFunction(()=>document.body.dataset.readingSize==='extra');
  await a.reload({waitUntil:'networkidle'});await boot(a);await a.waitForFunction(()=>document.body.dataset.readingSize==='extra');
  await boot(a,'reading-user-two');await a.waitForFunction(()=>document.body.dataset.readingSize==='large');await choose(a,'normal');await boot(a);await a.waitForFunction(()=>document.body.dataset.readingSize==='extra');
  const clippedAll=[];
  for(const width of [320,390,768,1440]){
   await a.setViewportSize({width,height:900});
   for(const view of ['importar','dashboard','calendario','comercial','inventario','materialesTransito','prioridades','solicitudesAlmacen','incidenciasDespacho','configClientes']){
    await a.evaluate(view=>cambiarVista(view,document.getElementById('nav-'+view)),view);await a.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
    const clipped=await a.evaluate(()=>[...document.querySelectorAll('.view.active button,.view.active input:not([type=file]),.view.active select')].filter(el=>el.getClientRects().length&&!el.closest('.table-wrap,.commercial-table-wrap,.transit-table-wrap,.count-tabs,#calMesNav')).filter(el=>{const r=el.getBoundingClientRect();return r.width>0&&(r.left < -1||r.right>innerWidth+1);}).map(el=>el.id||el.textContent.trim().slice(0,35)));
    if(clipped.length)clippedAll.push({width,view,clipped});
    if(view==='calendario'&&width<=900)assert.equal(await a.locator('.calendar-toolbar>div:first-child button').evaluateAll(buttons=>buttons.every(button=>button.scrollWidth<=button.clientWidth+1)),true,'month arrows must fit their touch buttons');
   }
  }
  assert.deepEqual(clippedAll,[]);
  await a.setViewportSize({width:390,height:844});await a.evaluate(()=>cambiarVista('calendario',document.getElementById('nav-calendario')));await a.screenshot({path:'/tmp/tms-reading-calendar-large.png',fullPage:true});
  assert.equal(await a.evaluate(()=>JSON.stringify(APP.lineItems)),originalData);assert.equal(await a.evaluate(()=>window.__readingWrites),0);assert.equal(await b.evaluate(()=>window.__readingWrites),0);
  await a.emulateMedia({media:'print'});assert.equal(await a.locator('#mobileReadingSizeButton').isVisible(),false);
  assert.deepEqual(errors,[]);console.log('PASS: 3 text sizes, larger default, close/keyboard dialog, per-user persistence, shared-tab preference, dynamic text without compounding, no data/autosave changes and 10 views at 320/390/768/1440px.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});

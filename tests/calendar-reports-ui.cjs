// Synthetic data only. Every production backend request is blocked.
const playwright=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright'),assert=require('node:assert/strict'),fs=require('node:fs/promises');
(async()=>{
 const browser=await playwright[process.env.TMS_TEST_ENGINE||'chromium'].launch({headless:true}),page=await browser.newPage({viewport:{width:1440,height:1000}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));page.on('dialog',async dialog=>{errors.push('Unexpected dialog: '+dialog.message());await dialog.dismiss();});
 await page.route('**/*.supabase.co/**',r=>r.fulfill({status:503,contentType:'application/json',body:'{"message":"Blocked synthetic test"}'}));
 try{
  await page.goto(process.env.TMS_TEST_URL||'http://127.0.0.1:8766',{waitUntil:'networkidle'});await page.waitForFunction(()=>!!window.TMSCalendarReports&&!!window.TMSAuditCore&&typeof XLSX!=='undefined');
  await page.evaluate(()=>{
   const user={id:'11111111-1111-4111-8111-111111111111',email:'reports@example.invalid'};APP.currentUserProfile={userId:user.id,email:user.email,nombre:'Prueba ejecutiva',rol:'Admin',activo:true};APP.userProfiles=[APP.currentUserProfile];window.__writes=0;window.guardarEnSupabase=async()=>{window.__writes++;};window.scheduleAutoSave=()=>{window.__writes++;};
   const item=extra=>({clienteId:'C1',clienteNombre:'Hotel Costa',pedidoCliente:'001',lineaPedidoCliente:'01',articulo:'A001',descripcionArticulo:'Copa de cristal para restaurante',cantidadSolicitada:5,cantidadSolicitadaOriginal:5,cantidadPendiente:5,cantidadFacturada:0,cantidadPendienteSinStock:3,montoPlanificado:15000,montoEntregadoNoFacturado:2500,rutaNombre:'RUTA NORTE',zona:'RUTA NORTE',fechaPlanificada:'02/11/2026',camionAsignado:'CAMION 1',nota:'Validar recepción con el cliente',...extra});
   Object.assign(APP,{planLoaded:true,planFecha:'02/11/2026',clienteConfig:{C1:{codigoCliente:'C1',nombreCliente:'Hotel Costa',rutaAsignada:'RUTA NORTE',diasRecepcion:[1,2,3,4,5],horarioAlmacen:'8–16',configuracionCompleta:true}},routeConfigs:[{id:'N',nombre:'RUTA NORTE',zona:'RUTA NORTE',diasOperacion:[1,2,3,4,5],activa:true}],lineItems:[item({}),item({pedidoCliente:'002',articulo:'LISTO',descripcionArticulo:'Plato listo para entregar',cantidadPendienteSinStock:0,montoPlanificado:9000,montoEntregadoNoFacturado:0}),item({clienteId:'C2',clienteNombre:'PRIVADO_OTRO_CLIENTE',articulo:'PRIVADO',pedidoCliente:'SECRETO'}),item({pedidoCliente:'QUEUE',fechaPlanificada:'',articulo:'QUEUE'})],transitMaterials:[{id:'T1',factura:'F-100',codigoMaterial:'A001',cantidad:4,eta:'2026-11-05'}]});
   APP.planLineItems=APP.lineItems.map(row=>({...row}));TMSAuditCore.rebuildDerivedState();renderCalendario();mostrarApp(user);cambiarVista('calendario',document.getElementById('nav-calendario'));
  });
  const original=await page.evaluate(()=>JSON.stringify(APP.lineItems)),report=page.locator('#calendarReportDialog');
  async function open(kind='dispatch'){
   await page.evaluate(kind=>{
    const rows=APP.rutas['02/11/2026']['CAMION 1'],index=rows.findIndex(row=>row.codigo==='C1');
    if(kind==='dispatch')abrirDespachoCalendario('02/11/2026','CAMION 1',index);
    else if(kind==='invoice')abrirDespachoCalendario('02/11/2026','CAMION 1',index,0);
    else if(kind==='note')abrirNotaModal('02/11/2026','CAMION 1',index);
    else abrirMoveModal('02/11/2026','CAMION 1',index);
   },kind);
   await page.locator(kind==='note'?'#notaModal [data-calendar-report]':kind==='move'?'#moveModal [data-calendar-report]':'#calendarDispatchDialog [data-calendar-report]').click();await report.waitFor({state:'visible'});
  }
  async function close(){await report.locator('[data-exec-close]').click();await page.waitForFunction(()=>!document.querySelector('#calendarReportDialog').open&&!document.querySelector('#calendarReportDialog').children.length);}
  async function download(format,path){const event=page.waitForEvent('download');await report.locator('[data-exec-format="'+format+'"]').click();const file=await event;await file.saveAs(path);assert.equal(await file.failure(),null);await report.locator('[data-exec-format="'+format+'"]').waitFor({state:'visible'});return path;}
  await open();assert.doesNotMatch(await report.innerText(),/PRIVADO|SECRETO/);assert.match(await report.innerText(),/26,500/);
  await page.route('**/vendor/exceljs-4.4.0.min.js',r=>r.abort());await report.locator('[data-exec-format="xlsx"]').click();await report.locator('.exec-status').filter({hasText:'No se pudo cargar'}).waitFor();assert.equal(await report.locator('[data-exec-format="xlsx"]').isEnabled(),true);await page.unroute('**/vendor/exceljs-4.4.0.min.js');
  await download('xlsx','/tmp/tms-executive-report.xlsx');
  const bytes=[...await fs.readFile('/tmp/tms-executive-report.xlsx')];
  const workbook=await page.evaluate(bytes=>{const w=XLSX.read(new Uint8Array(bytes),{type:'array'});return {names:w.SheetNames,total:w.Sheets.Resumen.B14.v,rows:XLSX.utils.sheet_to_json(w.Sheets.Pedidos,{range:4,defval:''})};},bytes);
  assert.deepEqual(workbook.names,['Resumen','Pedidos','Tránsito','Seguimiento']);assert.equal(workbook.total,26500);assert.equal(workbook.rows[0].Pedido,'001');assert.equal(workbook.rows[0]['Confirmado DOP'],15000);
  await download('pdf','/tmp/tms-executive-report.pdf');await close();await page.keyboard.press('Escape');
  await open('invoice');assert.doesNotMatch(await report.innerText(),/LISTO|PRIVADO/);assert.match(await report.innerText(),/17,500/);await close();await page.keyboard.press('Escape');
  await page.evaluate(()=>{const rows=APP.rutas['02/11/2026']['CAMION 1'];abrirNotaModal('02/11/2026','CAMION 1',rows.findIndex(row=>row.codigo==='C1'));});await page.locator('#notaInput').fill('Borrador ejecutivo <script>sin ejecutar</script>');await page.locator('#notaModal [data-calendar-report]').click();assert.match(await report.innerText(),/Borrador sin guardar/);assert.match(await report.innerText(),/<script>sin ejecutar<\/script>/);await close();await page.evaluate(()=>document.getElementById('notaModal').classList.remove('show'));
  await open('move');assert.match(await report.innerText(),/Programación propuesta/);await close();await page.evaluate(()=>cerrarMoveModal());
  // Bulk selection and unassigned queue reuse the same report entry point.
  await page.evaluate(()=>{const rows=APP.rutas['02/11/2026']['CAMION 1'];toggleSeleccionCalendario('02/11/2026','CAMION 1',0);abrirMoveSelectedModal();});await page.locator('#moveModal [data-calendar-report]').click();assert.match(await report.innerText(),/Programación/);await close();await page.evaluate(()=>cerrarMoveModal());
  await page.evaluate(()=>abrirMoveQueueModal(0));await page.locator('#moveModal [data-calendar-report]').click();assert.match(await report.innerText(),/Sin programar/);await close();await page.evaluate(()=>cerrarMoveModal());
  await page.evaluate(()=>abrirMoveRouteQueueModal('RUTA NORTE'));await page.locator('#moveModal [data-calendar-report]').click();assert.match(await report.innerText(),/QUEUE/);await close();await page.evaluate(()=>cerrarMoveModal());
  await page.evaluate(()=>agregarNotaCalendario('03/11/2026'));await page.locator('#calendarDayNoteText').fill('Revisar entregas de la mañana');await page.locator('[data-day-report]').click();assert.match(await report.innerText(),/Revisar entregas/);await close();await page.locator('[data-day-cancel]').click();assert.equal(await page.evaluate(()=>APP.calendarNotes.length),0);
  assert.equal(await page.evaluate(()=>JSON.stringify(APP.lineItems)),original);assert.equal(await page.evaluate(()=>window.__writes),0);
  // Reading-only users can export, and snapshots are regenerated on every opening.
  await page.evaluate(()=>{APP.currentUserProfile.rol='Solo lectura';renderCalendario();});await open();
  for(const width of [320,390,768]){await page.setViewportSize({width,height:900});assert.equal(await report.evaluate(el=>el.getBoundingClientRect().right<=innerWidth+1),true);assert.equal(await report.locator('.exec-actions').evaluate(el=>el.scrollWidth<=el.clientWidth+1),true);if(width===390)await page.screenshot({path:'/tmp/tms-executive-mobile.png',fullPage:true});}
  await close();await page.keyboard.press('Escape');
  await page.evaluate(()=>{APP.currentUserProfile.rol='Admin';const context={source:'dispatch',groups:[{nombre:'Hotel Costa',codigo:'C1',fecha:'02/11/2026',camion:'CAMION 1',rutaNombre:'RUTA NORTE',items:Array.from({length:150},(_,i)=>({pedidoCliente:'P'+i,articulo:'MATERIAL-'+i,descripcionArticulo:'Descripción larga para validar lectura y paginación completa sin cortar las filas del detalle ejecutivo.',cantidadPendiente:12,montoPlanificado:100,montoEntregadoNoFacturado:10}))}],author:'Prueba ejecutiva'};TMSCalendarReports.preview(TMSCalendarReports.build(context,TMSDomain));});await download('pdf','/tmp/tms-executive-long.pdf');await close();
  const pdfText=(await fs.readFile('/tmp/tms-executive-long.pdf')).toString('latin1');assert.equal(new Set(pdfText.match(/MATERIAL-\d+/g)).size,150);
  // Replacing the native day-note prompt preserves creation, editing and deletion.
  await page.evaluate(()=>agregarNotaCalendario('03/11/2026'));await page.locator('#calendarDayNoteText').fill('Nota guardada');await page.locator('[data-day-save]').click();await page.waitForFunction(()=>!document.getElementById('calendarDayNoteDialog').children.length);assert.equal(await page.evaluate(()=>APP.calendarNotes[0].texto),'Nota guardada');
  await page.evaluate(()=>editarNotaCalendario(APP.calendarNotes[0].id));await page.locator('#calendarDayNoteText').fill('Nota corregida');await page.locator('[data-day-save]').click();await page.waitForFunction(()=>!document.getElementById('calendarDayNoteDialog').children.length);assert.equal(await page.evaluate(()=>APP.calendarNotes[0].texto),'Nota corregida');
  await page.evaluate(()=>editarNotaCalendario(APP.calendarNotes[0].id));await page.locator('#calendarDayNoteText').fill('');await page.locator('[data-day-save]').click();await page.waitForFunction(()=>!document.getElementById('calendarDayNoteDialog').children.length);assert.equal(await page.evaluate(()=>APP.calendarNotes.length),0);assert.equal(await page.evaluate(()=>window.__writes),3);
  assert.deepEqual(errors,[]);console.log('PASS: actual styled Excel/PDF downloads and retry, scoped dispatch/invoice totals, draft labels, move/bulk/queue, day note create/edit/delete, readonly, mobile 320/390/768, all 150 PDF rows; export makes no operational writes.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});

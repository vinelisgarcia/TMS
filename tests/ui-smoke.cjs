/* Local, synthetic fixtures only. Supabase requests are blocked: never writes production. */
const { chromium } = require(process.env.TMS_PLAYWRIGHT || 'playwright');
const assert = require('node:assert/strict');
const path = require('node:path');
const url = process.env.TMS_TEST_URL || 'http://127.0.0.1:8765';
(async()=>{
  const browser=await chromium.launch({headless:true,...(process.env.TMS_BROWSER?{executablePath:process.env.TMS_BROWSER}:{})});
  try {
    const page=await browser.newPage({viewport:{width:1440,height:1000}});
    const errors=[];
    page.on('pageerror',error=>errors.push(error.message));
    await page.route('**/*.supabase.co/**',route=>route.fulfill({status:503,contentType:'application/json',body:'{"message":"Blocked in local test"}'}));
    await page.goto(url,{waitUntil:'networkidle'});
    await page.waitForFunction(()=>typeof TMSDomain!=='undefined' && !!window.TMSAuditCore && typeof Chart!=='undefined' && typeof XLSX!=='undefined');
    await page.screenshot({path:'/tmp/tms-audit-login.png'});
    await page.evaluate(()=>{
      Chart.defaults.animation=false;
      window.guardarEnSupabase=async()=>true;
      window.scheduleAutoSave=()=>TMSAuditCore.saveLocalSnapshot();
      const profile={userId:'test-admin',nombre:'Usuario de prueba',email:'test@example.invalid',rol:'Admin',activo:true};
      Object.assign(APP,{
        userProfiles:[profile],currentUserProfile:profile,planLoaded:true,controlLoaded:false,planFecha:'23/10/2026',
        clienteConfig:{C1:{codigoCliente:'C1',nombreCliente:'Hotel Costa Norte',rutaAsignada:'PUERTO PLATA',diasRecepcion:[1,2,3,4,5],horarioAlmacen:'8:00–16:00',configuracionCompleta:true},C2:{codigoCliente:'C2',nombreCliente:'Hotel Ciudad',rutaAsignada:'SANTO DOMINGO',diasRecepcion:[1,2,3,4,5],horarioAlmacen:'8:00–16:00',configuracionCompleta:true}},
        routeConfigs:[{id:'PP',nombre:'PUERTO PLATA',zona:'PUERTO PLATA',diasOperacion:[1,2,3,4,5],activa:true},{id:'SD',nombre:'SANTO DOMINGO',zona:'SANTO DOMINGO',diasOperacion:[1,2,3,4,5],activa:true}],
        solicitudesPlanAlmacen:[],solicitudesControlAlmacen:[],solicitudesAlmacen:[],controlLineItems:[],dailyPriorities:[],transitMaterials:[
          {id:'T1',factura:'F-100',codigoMaterial:'REF-1',descripcion:'Porcelanato de prueba',cantidad:4,eta:'2026-11-05'},
          {id:'T2',factura:'F-100',codigoMaterial:'REF-2',descripcion:'Grifería de prueba',cantidad:3,eta:'2026-11-05'},
          {id:'T3',factura:'F-200',codigoMaterial:'REF-1',descripcion:'Segunda importación',cantidad:2,eta:'2026-11-12'}
        ],
        costSettings:{vehicles:{'CAMION 1':'large'},routeZones:{'PUERTO PLATA':'puertoPlata','SANTO DOMINGO':'santoDomingo'},fxByDate:{'2026-10-23':60}},
        lineItems:[
          {clienteId:'C1',clienteNombre:'Hotel Costa Norte',pedidoCliente:'P-001',lineaPedidoCliente:'1',articulo:'REF-1',descripcionArticulo:'Porcelanato de prueba',cantidadSolicitada:10,cantidadSolicitadaOriginal:16,cantidadPendiente:6,cantidadFacturada:4,cantidadPendienteSinStock:6,montoPlanificado:18000,montoEntregadoNoFacturado:0,rutaNombre:'PUERTO PLATA',zona:'PUERTO PLATA',fechaPlanificada:'23/10/2026',camionAsignado:'CAMION 1'},
          {clienteId:'C2',clienteNombre:'Hotel Ciudad',pedidoCliente:'P-002',lineaPedidoCliente:'1',articulo:'REF-2',descripcionArticulo:'Grifería de prueba',cantidadSolicitada:5,cantidadSolicitadaOriginal:5,cantidadPendiente:5,cantidadFacturada:0,cantidadPendienteSinStock:0,montoPlanificado:12000,rutaNombre:'SANTO DOMINGO',zona:'SANTO DOMINGO',fechaPlanificada:'23/10/2026',camionAsignado:'CAMION 1'}
        ]
      });
      APP.planLineItems=APP.lineItems.map(row=>({...row}));
      APP.planLineItems.push({...APP.lineItems[0],pedidoCliente:'P-003',lineaPedidoCliente:'2',articulo:'REF-3',fechaPlanificada:'',cantidadPendiente:2,queueReason:'Asignar ruta completa desde la cola'});
      mostrarApp(profile);TMSAuditCore.rebuildDerivedState();
      actualizarDashboard();renderCalendario();renderComercial();renderConfigClientes();renderPrioridades();
      cambiarVista('dashboard',document.getElementById('nav-dashboard'));
    });
    assert.equal(await page.locator('#nav-importaciones').count(),0);
    assert.equal(await page.locator('#nav-rutas').count(),0);
    assert.equal(await page.locator('#nav-almacen').count(),0);
    assert.match(await page.locator('#auditCostsBody').innerText(),/1[,.]080/);
    assert.match(await page.locator('#auditCostsBody').innerText(),/Facturación real pendiente/);
    await page.locator('#auditCostPanel').scrollIntoViewIfNeeded();
    await page.screenshot({path:'/tmp/tms-audit-dashboard.png'});
    page.once('dialog',dialog=>dialog.accept());
    await page.locator('[data-audit="lock-trip"]').click();
    assert.match(await page.locator('#auditCostsBody').innerText(),/Salida confirmada/);
    await page.locator('#nav-calendario').click();
    await page.waitForSelector('#queueSemanaSection',{state:'visible'});
    assert.match(await page.locator('#queueSemanaSection').innerText(),/P-003/);
    await page.screenshot({path:'/tmp/tms-audit-calendar.png'});
    await page.locator('#nav-comercial').click();
    assert.doesNotMatch(await page.locator('#comercialV2Mount').innerText(),/undefined|NaN/);
    assert.match(await page.locator('#comercialV2Mount').innerText(),/F-100/);
    assert.match(await page.locator('#comercialV2Mount').innerText(),/sin reserva confirmada/);
    await page.screenshot({path:'/tmp/tms-audit-commercial.png'});
    const pdfDownload=page.waitForEvent('download',{timeout:45000});
    await page.evaluate(()=>exportarComercialPDF());
    const commercialPdf=await pdfDownload;
    const pdfBytes=require('node:fs').readFileSync(await commercialPdf.path());
    assert.equal(pdfBytes.subarray(0,5).toString(),'%PDF-');
    assert.ok(pdfBytes.length>1000);
    await page.evaluate(()=>{APP.lineItems[0].nota='NOTA PRIVADA NO EXPORTAR';TMSAuditCore.rebuildDerivedState();renderCalendario();cambiarVista('calendario');});
    assert.doesNotMatch(await page.locator('.cal-card-actions').first().innerText(),/[✅🟡🔄📝✏️↩️]/u);
    const routeDownload=page.waitForEvent('download',{timeout:45000});
    await page.evaluate(()=>exportarCalendarioVisiblePDF());
    const routePdf=await routeDownload;
    await routePdf.saveAs('/tmp/tms-route-summary.pdf');
    const routeBytes=require('node:fs').readFileSync(await routePdf.path());
    assert.equal(routeBytes.subarray(0,5).toString(),'%PDF-');
    const routeText=routeBytes.toString('latin1');
    assert.ok(routeText.includes('Hotel Costa Norte'));
    assert.ok(routeText.includes('Hotel Ciudad'));
    assert.ok(routeText.includes('Monto total de ruta:'));
    assert.ok(routeText.includes('30,000'));
    assert.ok(!routeText.includes('NOTA PRIVADA'));
    assert.ok(!routeText.includes('F-100'));
    await page.evaluate(()=>cambiarVista('almacen'));
    assert.equal(await page.locator('#solicitudesAlmacen').evaluate(el=>el.classList.contains('active')),true);

    await page.locator('#nav-materialesTransito').click();
    assert.equal(await page.locator('.audit-invoice').count(),2);
    page.once('dialog',dialog=>dialog.accept());
    await page.locator('[data-audit="receive-invoice"][data-invoice="F-100"]').click();
    assert.equal(await page.locator('.audit-invoice').count(),1);
    assert.equal(await page.evaluate(()=>APP.transitMaterials.filter(row=>row.receivedAt).length),2);
    await page.locator('[data-audit="transit-tab"][data-value="received"]').click();
    assert.match(await page.locator('.audit-invoice').innerText(),/F-100/);
    await page.screenshot({path:'/tmp/tms-audit-transit.png'});
    await page.locator('#nav-comercial').click();
    assert.doesNotMatch(await page.locator('#comercialV2Mount').innerText(),/F-100/);
    assert.match(await page.locator('#comercialV2Mount').innerText(),/F-200/);
    await page.locator('#nav-prioridades').click();
    const reminders=await page.evaluate(()=>APP.dailyPriorities.filter(row=>row.id.startsWith('audit-20261023')));
    assert.equal(reminders.length,2);
    await page.locator('#nav-configClientes').click();
    await page.locator('#auditClientCode').fill('C3');await page.locator('#auditClientName').fill('Cliente desde cero');
    await page.locator('[data-audit="create-client"]').click();
    assert.equal(await page.evaluate(()=>APP.configClienteSel),'C3');
    assert.match(await page.locator('#configEditor').innerText(),/Cliente desde cero/i);
    // Actual billing is a separate report, never inferred from operational balances.
    const billing=await page.evaluate(()=>{
      const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet([{'Fecha despacho':'2026-10-23',Camion:'CAMION 1',Salida:'1',Cliente:'C1',Pedido:'P-001',Factura:'SAP-1',Moneda:'DOP',Importe:100000}]),'Facturacion');
      return Array.from(new Uint8Array(XLSX.write(wb,{type:'array',bookType:'xlsx'})));
    });
    await page.locator('#auditBillingFile').setInputFiles({name:'facturado-test.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(billing)});
    await page.waitForFunction(()=>APP.costSettings.actualInvoices&&APP.costSettings.actualInvoices.length===1);
    await page.locator('#nav-dashboard').click();
    assert.match(await page.locator('#auditCostsBody').innerText(),/64\.8%/);
    const downloadPromise=page.waitForEvent('download');await page.locator('[data-audit="export-costs"]').click();const download=await downloadPromise;
    assert.ok(require('node:fs').statSync(await download.path()).size>1000);
    await page.locator('#nav-configClientes').click();
    await page.screenshot({path:'/tmp/tms-audit-config.png'});
    for(const width of [390,768,1024]){
      await page.setViewportSize({width,height:844});
      await page.evaluate(()=>cambiarVista('materialesTransito',document.getElementById('nav-materialesTransito')));
      const size=await page.evaluate(()=>({width:innerWidth,scroll:document.documentElement.scrollWidth}));
      assert.ok(size.scroll<=size.width+1,`horizontal page overflow at ${width}: ${size.scroll}`);
      if(width===390)await page.screenshot({path:'/tmp/tms-audit-mobile.png',fullPage:true});
      if(width<=900){
        if(await page.locator('#sidebar').isVisible())console.log(await page.evaluate(()=>({display:getComputedStyle(document.getElementById('sidebar')).display,sheets:[...document.styleSheets].map(s=>s.href||s.ownerNode.id||'inline'),width:innerWidth})));
        assert.equal(await page.locator('#sidebar').isVisible(),false);
      }
    }
    await page.setViewportSize({width:390,height:844});
    for(const view of ['dashboard','calendario','comercial','configClientes','prioridades','solicitudesAlmacen','almacen','incidenciasDespacho','importar']){
      await page.evaluate(view=>cambiarVista(view,document.getElementById('nav-'+view)),view);
      const size=await page.evaluate(()=>({width:innerWidth,scroll:document.documentElement.scrollWidth}));
      assert.ok(size.scroll<=size.width+1,`mobile overflow in ${view}: ${size.scroll}`);
    }
    await page.evaluate(()=>{setVisualTheme('dark');cambiarVista('dashboard',document.getElementById('nav-dashboard'));});
    await page.screenshot({path:'/tmp/tms-audit-dark-mobile.png',fullPage:true});
    await page.evaluate(()=>setVisualTheme('light'));
    // Import a real workbook-shaped fixture through the same file input as the user.
    await page.setViewportSize({width:1440,height:1000});
    await page.evaluate(()=>cambiarVista('importar',document.getElementById('nav-importar')));
    const plan=await page.evaluate(()=>{
      const wb=XLSX.utils.book_new();
      const rows=[['C1','Hotel Costa Norte','PLAN-1'],['C1','Hotel Costa Norte','PLAN-2'],['C2','Hotel Ciudad','PLAN-3'],['C9','Cliente sin ruta','PLAN-4']].map(([Cliente,Nombre,Pedido])=>({Cliente,'Nombre Comercial':Nombre,'Pedido de cliente':Pedido,'Posición de pedido de cliente':'1',Producto:'TEST-REF','Cantidad solicitada':10,'Cantidad pendiente de servir':10,'Cantidad pendiente sin stock':0,'Importe confirmado no entregado':1000}));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Plan 23.10.2026');return Array.from(new Uint8Array(XLSX.write(wb,{type:'array',bookType:'xlsx'})));
    });
    await page.locator('#sapFile').setInputFiles({name:'Plan 23.10.2026.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(plan)});
    await page.waitForFunction(()=>APP.planLineItems.some(row=>row.pedidoCliente==='PLAN-4'));
    const imported=await page.evaluate(()=>({queue:APP.queueSemana.map(row=>row.pedidoCliente),dates:APP.lineItems.filter(row=>row.rutaNombre==='PUERTO PLATA').map(row=>row.fechaPlanificada)}));
    assert.ok(imported.queue.includes('PLAN-4'));assert.equal(new Set(imported.dates).size,1);
    // Persistence round-trip of receipt and cost settings using the actual local snapshot.
    const persisted=await page.evaluate(()=>JSON.parse(localStorage.getItem('tms_alvarez_v2_state')).app);
    assert.equal(persisted.transitMaterials.filter(row=>row.receivedAt).length,2);
    assert.equal(persisted.costSettings.vehicles['CAMION 1'],'large');
    const denied=await page.evaluate(()=>{
      APP.currentUserProfile={email:'commercial@example.invalid',rol:'Comercial'};
      renderMaterialesTransito();return document.querySelectorAll('[data-audit="receive-invoice"]').length;
    });
    assert.equal(denied,0,'Commercial profile cannot receive invoices');
    assert.equal(errors.length,0,'Browser errors: '+errors.join('\n'));
    console.log('UI smoke passed: dashboard, initial queue, commercial transit, invoice receipt/history, clients, reminders, persistence and 390/768/1024px.');
    console.log('Previews: /tmp/tms-audit-{login,dashboard,calendar,commercial,transit,config,mobile}.png');
  } finally { await browser.close(); }
})().catch(error=>{console.error(error);process.exitCode=1;});

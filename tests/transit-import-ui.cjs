// Synthetic fixtures; real Supabase requests are blocked. Private input is optional and never committed.
const {chromium}=require(process.env.TMS_PLAYWRIGHT || 'playwright');
const assert=require('node:assert/strict');
const fs=require('node:fs');
(async()=>{
  const browser=await chromium.launch({headless:true});
  try {
    const page=await browser.newPage({viewport:{width:1440,height:1000}}),errors=[],dialogs=[];
    page.on('pageerror',error=>errors.push(error.message));
    let accept=true;
    page.on('dialog',async dialog=>{dialogs.push({type:dialog.type(),message:dialog.message()});await (dialog.type()==='confirm' && !accept ? dialog.dismiss() : dialog.accept());});
    await page.route('**/*.supabase.co/**',route=>route.fulfill({status:503,contentType:'application/json',body:'{"message":"Blocked in test"}'}));
    await page.goto(process.env.TMS_TEST_URL || 'http://127.0.0.1:8766',{waitUntil:'networkidle'});
    await page.waitForFunction(()=>!!window.TMSAuditCore && typeof XLSX!=='undefined');
    await page.evaluate(async()=>{
      const u={id:'11111111-1111-4111-8111-111111111111',email:'test@example.invalid'};
      SB.auth.getUser=async()=>({data:{user:u},error:null});
      SB.rpc=async(name,args)=>{
        if(name!=='tms_get_snapshot')throw new Error('Unexpected remote write: '+name);
        if(args?.p_known_revision===1)return {data:{revision:1,unchanged:true},error:null};
        return {data:{revision:1,tables:{tms_order_lines:[],tms_client_configs:[],tms_route_configs:[],tms_route_cost_history:[],tms_warehouse_plan:[],tms_warehouse_control:[],tms_warehouse_history:[],tms_settings:[],tms_user_profiles:[{id:u.id,auth_user_id:u.id,email:u.email,nombre:'Prueba',rol:'Admin',activo:true}]}},error:null};
      };
      SB.channel=()=>({on(){return this;},subscribe(){return this;}});SB.removeChannel=async()=>{};
      await cargarDesdeSupabase();
      APP.userProfiles=[{userId:u.id,authUserId:u.id,email:u.email,nombre:'Prueba',rol:'Admin',activo:true}];
      APP.currentUserProfile=APP.userProfiles[0];APP.transitMaterials=[];
      window.guardarEnSupabase=async()=>true;window.scheduleAutoSave=()=>TMSAuditCore.saveLocalSnapshot();
      mostrarApp(u);cambiarVista('materialesTransito');
    });
    async function workbook(matrix){
      return Buffer.from(await page.evaluate(matrix=>{
        const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Portada']]),'Portada');
        XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(matrix),'Detalle');return [...new Uint8Array(XLSX.write(wb,{type:'array',bookType:'xlsx'}))];
      },matrix));
    }
    async function upload(buffer){
      await page.locator('#transitEta').fill('2026-11-20');
      await page.locator('#transitFile').setInputFiles({name:'packing-test.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer});
      await page.getByRole('button',{name:'Importar materiales',exact:true}).click();
      await page.waitForFunction(()=>!document.getElementById('transitStatus')?.textContent.includes('Procesando'));
    }
    const headers=['Codigo cliente del producto','Descripción de Producto','Referencia proveedor','Packslip','Cantidad solicitada','Factura No.','Referencia embarque'];
    const buffer=await workbook([['Plantilla'],[],headers,['000123','Prueba','SUP-1','PL-1',12,'F-1','SHIP-1'],['000123','Prueba','SUP-1','PL-2',8,'F-1','SHIP-2'],['000456','Otro','SUP-2','PL-3',4,'','SHIP-3']]);
    await upload(buffer);
    let rows=await page.evaluate(()=>APP.transitMaterials);
    assert.equal(rows.length,2);assert.equal(rows[0].codigoMaterial,'000123');assert.equal(rows[0].cantidad,20);assert.equal(rows[0].packingList,'PL-1 / PL-2');
    assert.equal(rows[1].factura,'PL-3');assert.equal(rows[1].tipoDocumento,'packing_list');assert.equal(rows[1].eta,'2026-11-20');
    assert.ok(dialogs.some(d=>d.type==='confirm' && d.message.includes('24 unidades')));
    await page.locator('.audit-invoice details').first().locator('summary').click();
    assert.match(await page.locator('.audit-invoice').first().innerText(),/Packslip: PL-1 \/ PL-2/);
    assert.match(await page.locator('.audit-invoice').nth(1).innerText(),/Packing list/i);
    await upload(buffer);assert.equal(await page.evaluate(()=>APP.transitMaterials.reduce((n,r)=>n+r.cantidad,0)),24);
    accept=false;await upload(await workbook([headers,['000123','Prueba','SUP-1','PL-1',99,'F-1','SHIP-1']]));accept=true;
    assert.equal(await page.evaluate(()=>APP.transitMaterials[0].cantidad),20);
    await upload(await workbook([headers,['000123','Prueba','SUP-1','PL-1','cantidad inválida','F-1','SHIP-1']]));
    assert.match(await page.locator('#transitStatus').innerText(),/Error/);assert.equal(await page.evaluate(()=>APP.transitMaterials[0].cantidad),20);
    await page.evaluate(()=>renderMaterialesTransito());
    const downloadEvent=page.waitForEvent('download');await page.getByRole('button',{name:'Exportar Excel',exact:true}).click();
    const download=await downloadEvent,exported=fs.readFileSync(await download.path());
    const exportedRows=await page.evaluate(bytes=>XLSX.utils.sheet_to_json(XLSX.read(new Uint8Array(bytes),{type:'array'}).Sheets['Materiales transito']),[...exported]);
    assert.equal(exportedRows[0].Packslip,'PL-1 / PL-2');assert.equal(exportedRows[0]['Referencia proveedor'],'SUP-1');
    await page.locator('[data-audit="receive-invoice"][data-invoice="PL-3"]').click();
    assert.ok(await page.evaluate(()=>APP.transitMaterials.find(r=>r.factura==='PL-3').receivedAt));
    const snapshot=await page.evaluate(()=>JSON.parse(localStorage.getItem('tms_alvarez_v2_state:11111111-1111-4111-8111-111111111111')));
    assert.equal(snapshot.app.transitMaterials[0].packingList,'PL-1 / PL-2');
    if(process.env.TMS_PACKING_FILE){
      await page.evaluate(()=>{APP.transitMaterials=[];renderMaterialesTransito();});
      await upload(fs.readFileSync(process.env.TMS_PACKING_FILE));
      rows=await page.evaluate(()=>APP.transitMaterials);
      if(process.env.TMS_PACKING_EXPECTED_LINES)assert.equal(rows.length,Number(process.env.TMS_PACKING_EXPECTED_LINES));
      if(process.env.TMS_PACKING_EXPECTED_QUANTITY)assert.equal(rows.reduce((n,r)=>n+r.cantidad,0),Number(process.env.TMS_PACKING_EXPECTED_QUANTITY));
      assert.ok(rows.every(row=>row.codigoMaterial && row.packingList && row.referenciaProveedor));
      console.log(`PASS: provided workbook imported locally: ${rows.length} lines; ${rows.reduce((n,r)=>n+r.cantidad,0)} units; no production writes.`);
    }
    assert.deepEqual(errors,[]);
    console.log('PASS: header detection, all detail sheets, packing-only and invoice import, totals, reimport, cancel, invalid input, export, receipt and saved metadata.');
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

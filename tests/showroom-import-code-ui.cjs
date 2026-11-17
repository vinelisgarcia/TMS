// Synthetic workbooks only. Every real Supabase request is blocked.
const playwright=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright'),assert=require('node:assert/strict');
const base=process.env.TMS_TEST_URL||'http://127.0.0.1:8766';
(async()=>{
 const browser=await playwright[process.env.TMS_TEST_ENGINE||'chromium'].launch({headless:true}),page=await browser.newPage({viewport:{width:1440,height:1000}}),errors=[];
 const db={revision:0,payload:{version:1,sources:{ct:null,ad:null},cases:{},settings:{sameScope:true}}};let writes=0;
 page.on('pageerror',error=>errors.push(error.message));await page.route('**/*.supabase.co/**',route=>route.fulfill({status:503,body:'Blocked in test'}));
 await page.exposeFunction('__codeRpc',async(name,args)=>{
  if(name==='tms_showroom_get_snapshot')return {data:args.p_known_revision===db.revision?{revision:db.revision,unchanged:true}:{revision:db.revision,tables:{tms_showroom_state:[{id:true,payload:structuredClone(db.payload)}]}},error:null};
  assert.equal(name,'tms_showroom_save_snapshot');assert.equal(args.p_expected_revision,db.revision);writes++;db.payload=structuredClone(args.p_payload.tms_showroom_state[0].payload);return {data:{revision:++db.revision,changed:true},error:null};
 });
 async function boot(){await page.evaluate(()=>{
  const user={id:'11111111-1111-4111-8111-111111111111',email:'code-review@example.invalid'};
  SB.auth.getUser=async()=>({data:{user},error:null});SB.rpc=(name,args)=>window.__codeRpc(name,args);SB.channel=()=>({on(){return this;},subscribe(){return this;}});SB.removeChannel=async()=>{};
  APP.currentUserProfile={userId:user.id,authUserId:user.id,email:user.email,rol:'Admin',activo:true};APP.userProfiles=[APP.currentUserProfile];mostrarApp(user);cambiarVista('showroom',document.getElementById('nav-showroom'));
 });await page.waitForFunction(()=>TMSShowroom.getSyncState().phase==='synced');}
 async function open(file){await page.locator('#showroomCTFile').setInputFiles(file);await page.locator('#showroomImportDialog').waitFor();await page.locator('#showroomAsOf').fill('2026-11-10');}
 async function preview(){await page.locator('[data-show-inv="preview"]').click();}
 async function commit(){await page.locator('#showroomConfirmImport').click();await page.locator('#showroomImportDialog').waitFor({state:'detached'});await page.waitForFunction(()=>TMSShowroom.getSyncState().phase==='synced');}
 async function excel(kind){return Buffer.from(await page.evaluate(kind=>{
  const book=XLSX.utils.book_new(),sheet={},head=['Responsable','ID','Ubicacion Showroom','','Cantidad Contada','DIFERENCIAS'];
  if(kind==='mixed'){
   XLSX.utils.sheet_add_aoa(sheet,[head,['Operador',123.5,'A01','',2,''],['Operador',Number.MAX_SAFE_INTEGER+1,'A02','',3,''],['Operador',7,'A03','',4,'']],{origin:'B3'});sheet.C6.z='000000';delete sheet.C6.w;
  }else XLSX.utils.sheet_add_aoa(sheet,[head,['Operador',kind==='invalid'?99.5:123,'A01','',1,'']],{origin:'A1'});
  XLSX.utils.book_append_sheet(book,sheet,'Conteo');return [...new Uint8Array(XLSX.write(book,{type:'array',bookType:'xlsx'}))];
 },kind));}
 const file=async kind=>({name:'conteo-'+kind+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:await excel(kind)});
 try{
  await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>!!window.TMSShowroom);await boot();
  await open({name:'anterior.csv',mimeType:'text/csv',buffer:Buffer.from('Material;Quantity;Unit of measurement\nKEEP;9;UND')});await preview();await commit();assert.equal(writes,1);const before=structuredClone(db.payload);
  await open(await file('mixed'));assert.equal(await page.locator('#showroomHeaderRow').inputValue(),'3');await preview();
  assert.equal(await page.locator('#showroomConfirmImport').isEnabled(),false);assert.equal(await page.locator('[data-show-code-row]').count(),2);assert.match(await page.locator('#showroomImportPreview').innerText(),/C4/);assert.match(await page.locator('#showroomImportPreview').innerText(),/C5/);assert.deepEqual(db.payload,before);assert.equal(writes,1);
  await page.getByLabel('Referencia exacta fila 4',{exact:true}).fill('00123.50');await preview();assert.equal(await page.getByLabel('Referencia exacta fila 4',{exact:true}).inputValue(),'00123.50');assert.equal(await page.locator('#showroomConfirmImport').isEnabled(),false);assert.equal(writes,1);
  await page.setViewportSize({width:390,height:844});await page.locator('.showroom-code-list').scrollIntoViewIfNeeded();assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1));
  for(const input of await page.locator('[data-show-code-row]').all())assert.ok(await input.evaluate(el=>{const r=el.getBoundingClientRect();return r.left>=0&&r.right<=innerWidth;}));
  await page.screenshot({path:'/tmp/tms-showroom-code-review-mobile.png'});
  await page.getByLabel('Referencia exacta fila 5',{exact:true}).fill('9007199254740993');await preview();assert.equal(await page.locator('#showroomConfirmImport').isEnabled(),true);assert.match(await page.locator('#showroomImportPreview').innerText(),/ceros iniciales/);
  await page.locator('#showroomAsOf').fill('2026-11-09');assert.equal(await page.locator('#showroomConfirmImport').isEnabled(),false);await preview();await commit();assert.equal(writes,2);
  const source=db.payload.sources.ct;assert.equal(source.totals.quantity,9);assert.equal(source.acceptedRowCount,3);assert.equal(source.worksheet,'Conteo');assert.deepEqual(source.materials.map(row=>row.code).sort(),['000007','00123.50','9007199254740993'].sort());
  assert.deepEqual(source.codeCorrections.map(row=>[row.row,row.column,row.readValue,row.code]),[[4,3,'123.5','00123.50'],[5,3,String(Number.MAX_SAFE_INTEGER+1),'9007199254740993']]);assert.deepEqual(source.materials.flatMap(row=>row.details.map(detail=>detail.row)).sort(),[4,5,6]);
  await page.reload({waitUntil:'networkidle'});await boot();assert.deepEqual(await page.evaluate(()=>TMSShowroom.getState().sources.ct.codeCorrections),source.codeCorrections);
  await open(await file('invalid'));await preview();assert.equal(await page.locator('#showroomConfirmImport').isEnabled(),false);assert.equal(await page.locator('[data-show-code-row="2"]').count(),1);assert.match(await page.locator('#showroomImportPreview').innerText(),/B2/);assert.equal(writes,2);await page.locator('[data-show-inv="cancel-import"]').click();
  await open(await file('safe'));await preview();assert.equal(await page.locator('#showroomConfirmImport').isEnabled(),true);
  await page.evaluate(()=>{const original=crypto.subtle.digest.bind(crypto.subtle);crypto.subtle.digest=async(...args)=>{await new Promise(resolve=>window.__releaseCodeHash=resolve);return original(...args);};window.__codeCommit=TMSShowroom.confirmImport().then(()=>null,error=>error.message);});
  await page.waitForFunction(()=>typeof window.__releaseCodeHash==='function');await page.locator('#showroomAsOf').fill('2026-11-08');await page.evaluate(()=>window.__releaseCodeHash());assert.match(await page.evaluate(()=>window.__codeCommit),/cambió/);assert.equal(writes,2);assert.deepEqual(db.payload.sources.ct,source);await page.locator('[data-show-inv="cancel-import"]').click();
  assert.deepEqual(errors,[]);console.log('PASS: unsafe IDs show exact Excel cells, partial/all-invalid imports stay blocked, reviewed text IDs and padded zeros load intact, source quantities remain complete, corrections persist, edits invalidate previews, changed in-flight imports cannot save and mobile fields fit.');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

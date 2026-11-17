const { chromium } = require(process.env.TMS_PLAYWRIGHT || '/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright');
const assert = require('node:assert/strict');
(async () => {
 const b = await chromium.launch();
 try {
  const p = await b.newPage({viewport:{width:1440,height:950}});
  await p.route('**/*.supabase.co/**', r => r.fulfill({status:503,body:'{}'}));
  await p.addInitScript(() => {
   const saved = {}; let calls = 0;
   window.getPickerCount = () => calls;
   function makeHandle(name) { return {name,queryPermission:async()=> 'granted',requestPermission:async()=> 'granted',getFile:async()=>new File(['file'],name)}; }
   window.showOpenFilePicker = async () => { calls++; return [makeHandle('plan.xlsx')]; };
   Object.defineProperty(window, 'indexedDB', { value: { open() {
    const req = {};
    setTimeout(() => {
     req.result = { createObjectStore() {}, transaction() {
      const tx = {objectStore() { return {
       put(handle,key) {saved[key]=handle.name;const r={};setTimeout(()=>tx.oncomplete?.(),0);return r;},
       get(key) {const r={};setTimeout(()=>{r.result=saved[key]?makeHandle(saved[key]):undefined;tx.oncomplete?.();},0);return r;}
      }; }}; return tx;
     }};
     req.onupgradeneeded?.();req.onsuccess?.();
    },0); return req;
   }}});
   window.confirm = () => true;
  });
  await p.goto((process.env.TMS_TEST_URL || 'http://127.0.0.1:8765') + '/?modulo=rutas',{waitUntil:'networkidle'});
  await p.evaluate(()=>{APP.currentUserProfile={rol:'Admin',email:'first@example.invalid',activo:true};mostrarApp(APP.currentUserProfile);window.dispatchEvent(new Event('tms-permissions-changed'));document.getElementById('nav-importar').click();window.procesarArchivo=(file,mode)=>{window.lastProcessed=[file.name,mode]};});
  await p.waitForFunction(()=>!document.querySelector('[data-sap-config="plan"]').disabled);
  const boxes=await Promise.all(['#sapPlanUpload','#sapControlUpload','[data-sap-upload="warehouse-plan"]'].map(s=>p.locator(s).boundingBox()));
  assert.ok(Math.max(...boxes.map(x=>x.y))-Math.min(...boxes.map(x=>x.y))<1);
  await p.locator('#sapPlanUpload').click();await p.waitForFunction(()=>window.lastProcessed);
  assert.deepEqual(await p.evaluate(()=>window.lastProcessed),['plan.xlsx','plan']);assert.equal(await p.evaluate(()=>getPickerCount()),1);
  await p.locator('#sapPlanUpload').click();await p.waitForTimeout(30);assert.equal(await p.evaluate(()=>getPickerCount()),1);
  await p.evaluate(()=>{APP.currentUserProfile={rol:'Admin',email:'other@example.invalid',activo:true};window.dispatchEvent(new Event('tms-permissions-changed'));});
  await p.waitForFunction(()=>!document.querySelector('[data-sap-config="plan"]').disabled);await p.locator('#sapPlanUpload').click();await p.waitForTimeout(30);assert.equal(await p.evaluate(()=>getPickerCount()),2);
  console.log('Saved file selection, reuse, account isolation and alignment passed');
 } finally {await b.close();}
})().catch(e=>{console.error(e);process.exit(1)});

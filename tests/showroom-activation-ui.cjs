// Synthetic identity and RPC only. No real Supabase requests are allowed.
const {chromium}=require(process.env.TMS_PLAYWRIGHT||'playwright');
const assert=require('node:assert/strict');
(async()=>{
 const browser=await chromium.launch({headless:true});
 try{
  let active=false,writes=0;const errors=[];
  const page=await browser.newPage();page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/*.supabase.co/**',r=>r.fulfill({status:503,body:'Blocked in test'}));
  await page.exposeFunction('__activationRpc',async(name)=>{
   if(name!=='tms_showroom_get_snapshot'){writes++;throw new Error('Unexpected write');}
   if(!active)return {data:null,error:{code:'PGRST202',message:'Function not in schema cache'}};
   return {data:{revision:0,tables:{tms_showroom_state:[{id:true,payload:{version:1,sources:{ct:null,ad:null},cases:{},settings:{sameScope:true}}}]}},error:null};
  });
  await page.goto(process.env.TMS_TEST_URL||'http://127.0.0.1:8766',{waitUntil:'networkidle'});
  await page.waitForFunction(()=>!!window.TMSShowroom);
  await page.evaluate(()=>{
   const user={id:'11111111-1111-4111-8111-111111111111',email:'test@example.invalid'};
   SB.auth.getUser=async()=>({data:{user},error:null});SB.rpc=name=>window.__activationRpc(name);
   SB.channel=()=>({on(){return this;},subscribe(){return this;}});SB.removeChannel=async()=>{};
   APP.currentUserProfile={userId:user.id,email:user.email,rol:'Admin',activo:true};APP.userProfiles=[APP.currentUserProfile];
   mostrarApp(user);cambiarVista('showroom',document.getElementById('nav-showroom'));
  });
  await page.waitForFunction(()=>document.getElementById('showroomMount').textContent.includes('Activación pendiente'));
  assert.equal(await page.locator('#showroom input[type=file]').count(),0);
  assert.equal(await page.locator('#showroomNavMount button').count(),0);assert.equal(writes,0);
  active=true;await page.getByRole('button',{name:'Comprobar activación'}).click();
  await page.waitForFunction(()=>TMSShowroom.getSyncState().phase==='synced');
  assert.equal(await page.locator('#showroomCTFile').count(),1);assert.equal(await page.locator('#showroomNavMount button').count(),5);
  assert.doesNotMatch(await page.locator('#showroomMount').innerText(),/Activación pendiente/);
  assert.equal(writes,0);assert.deepEqual(errors,[]);
  console.log('PASS: missing database shows activation notice without upload/count controls; retry restores the normal module after activation; zero writes.');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

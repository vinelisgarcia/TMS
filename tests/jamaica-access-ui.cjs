const playwright=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright'),assert=require('node:assert/strict');
(async()=>{
 const browser=await playwright[process.env.TMS_TEST_ENGINE||'chromium'].launch({headless:true});
 try{
  for(const scenario of [{role:null,allowed:false},{role:'Comercial',allowed:false},{role:'Admin',active:false,allowed:false},{role:'Admin',allowed:true},{role:'Comercial',grant:true,allowed:true}]){
   const page=await browser.newPage();await page.addInitScript(s=>window.__accessScenario=s,scenario);
   await page.route('**/npm/@supabase/supabase-js@2',r=>r.fulfill({contentType:'application/javascript',body:`window.supabase={createClient:()=>({auth:{getUser:async()=>({data:{user:window.__accessScenario.role?{id:'synthetic',user_metadata:{rol:'Admin'}}:null},error:null}),onAuthStateChange:callback=>{window.__signOut=()=>callback('SIGNED_OUT');}},from:table=>({select(){return this;},eq(){return this;},async maybeSingle(){return table==='tms_user_profiles'?{data:{rol:window.__accessScenario.role,activo:window.__accessScenario.active!==false},error:null}:{data:{valor:{Comercial:{auditoriaJamaica:{ver:!!window.__accessScenario.grant}}}},error:null};}})})};` }));
   await page.route('**/*.supabase.co/**',r=>r.abort());await page.goto((process.env.TMS_TEST_URL||'http://127.0.0.1:8766')+'/auditoria-jamaica.html',{waitUntil:'networkidle'});
   await page.evaluate(()=>TMSStandaloneAccess.ready);assert.equal(await page.evaluate(()=>TMSStandaloneAccess.allowed),scenario.allowed);assert.equal(await page.locator('main').isVisible(),scenario.allowed);assert.equal(await page.locator('#moduleAccessGate').isVisible(),!scenario.allowed);
   await page.evaluate(()=>ingest([{name:'invoice.xlsx',size:20}]));assert.equal(await page.evaluate(()=>files.size),scenario.allowed?1:0);
   if(scenario.allowed){await page.evaluate(()=>window.__signOut());assert.equal(await page.locator('main').isVisible(),false);assert.equal(await page.evaluate(()=>files.size),0);}
   await page.close();
  }
  console.log('PASS: direct Jamaica URL requires verified account, active profile and Admin or explicit view grant; user metadata cannot grant access; logout clears files.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});

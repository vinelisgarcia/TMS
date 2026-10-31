// All accounts and stock below are synthetic; this test never contacts a server.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
const {createRequire}=require('node:module');
const requireRuntime=createRequire(path.join(process.env.TMS_PGLITE_RUNTIME||'/tmp/tms-sync-validation','package.json'));
const {PGlite}=requireRuntime('@electric-sql/pglite');
const ADMIN='11111111-1111-4111-8111-111111111111',READER='22222222-2222-4222-8222-222222222222',EDITOR='33333333-3333-4333-8333-333333333333';
(async()=>{
  const db=new PGlite();await db.waitReady;const checks=[];
  const q=async(sql,args=[])=>(await db.query(sql,args)).rows[0]?.result;
  const auth=async(id,role='authenticated')=>{await db.exec('reset role');await db.query("select set_config('request.jwt.claims',$1,false)",[JSON.stringify({sub:id,role})]);await db.exec('set role '+role);};
  const snapshot=()=>q('select public.tms_inventory_get_snapshot() as result');
  const save=(rev,payload)=>q('select public.tms_inventory_save_snapshot($1,$2::jsonb) as result',[rev,JSON.stringify(payload)]);
  const check=name=>{checks.push(name);console.log('PASS:',name);};
  try{
    await db.exec(`create role anon;create role authenticated;create role service_role bypassrls;create schema auth;create table auth.users(id uuid primary key);create function auth.jwt() returns jsonb language sql stable as $$select coalesce(nullif(current_setting('request.jwt.claims',true),''),'{}')::jsonb$$;create function auth.uid() returns uuid language sql stable as $$select (auth.jwt()->>'sub')::uuid$$;grant usage on schema public,auth to anon,authenticated,service_role;grant execute on all functions in schema auth to anon,authenticated,service_role;create table public.tms_user_profiles(id uuid primary key default gen_random_uuid(),auth_user_id uuid,nombre text,email text,rol text,activo boolean,created_at timestamptz default now());create table public.tms_settings(clave text primary key,valor jsonb);create table public.tms_sync_state(id boolean primary key,revision bigint);insert into public.tms_sync_state values(true,42);create publication supabase_realtime;insert into auth.users values('${ADMIN}'),('${READER}'),('${EDITOR}');insert into public.tms_user_profiles(auth_user_id,nombre,email,rol,activo) values('${ADMIN}','Admin','admin@example.invalid','Admin',true),('${READER}','Reader','reader@example.invalid','Reader',true),('${EDITOR}','Editor','editor@example.invalid','Editor',true);insert into public.tms_settings values('role_permissions','{"Reader":{"inventario":{"ver":true,"editar":false,"importar":false}},"Editor":{"inventario":{"ver":true,"editar":true,"importar":true}}}');`);
    const migration=fs.readFileSync(path.join(__dirname,'../supabase/migrations/0009_inventory_reconciliation.sql'),'utf8');
    await db.exec(migration);await db.exec(migration);check('additive migration is idempotent and publishes only inventory revision');
    assert.equal(await q("select count(*)::int as result from pg_publication_tables where pubname='supabase_realtime' and tablename='tms_inventory_sync_state'"),1);
    assert.equal(await q("select count(*)::int as result from pg_publication_tables where tablename='tms_inventory_state'"),0);
    await auth(ADMIN);let first=await snapshot();assert.equal(first.revision,0);assert.equal(first.tables.tms_inventory_state[0].payload.sources.ad,null);
    assert.deepEqual(await save(0,{}),{revision:0,changed:false});
    assert.deepEqual(await q('select public.tms_inventory_get_snapshot(0) as result'),{revision:0,unchanged:true});check('empty state and no-op saves do not invent inventory or write revisions');
    const value=structuredClone(first.tables.tms_inventory_state[0].payload);value.sources.ct={materials:Array.from({length:2401},(_,code)=>({code:String(code),quantity:2}))};
    const payload={tms_inventory_state:[{id:true,payload:value}]};
    assert.equal((await save(0,payload)).revision,1);assert.equal((await snapshot()).tables.tms_inventory_state[0].payload.sources.ct.materials.length,2401);
    assert.equal((await save(1,payload)).changed,false);check('complete inventory snapshot above 1000 rows and content no-op detection');
    await assert.rejects(()=>save(0,payload),error=>error.code==='40001');assert.equal((await snapshot()).revision,1);check('stale simultaneous session cannot overwrite committed inventory');
    await assert.rejects(()=>save(1,{tms_settings:[]}),error=>error.code==='22023');
    await assert.rejects(()=>save(1,{tms_inventory_state:[{id:true,payload:{sources:{}}}]}),error=>error.code==='22023');
    assert.equal((await snapshot()).revision,1);check('invalid payload and cross-module table writes are rejected atomically');
    await db.exec('reset role');assert.equal(await q('select revision as result from public.tms_sync_state'),42);check('inventory saves do not change the main TMS revision');
    await auth(READER);assert.equal((await snapshot()).revision,1);await assert.rejects(()=>save(1,payload),error=>error.code==='42501');check('explicit inventory reader may view but cannot save');
    await auth(EDITOR);value.cases.A={status:'en_revision'};assert.equal((await save(1,{tms_inventory_state:[{id:true,payload:value}]})).revision,2);check('only an explicitly permitted role may edit inventory');
    await db.exec('reset role');await db.exec(`update public.tms_user_profiles set activo=false where auth_user_id='${EDITOR}'`);await auth(EDITOR);await assert.rejects(snapshot,error=>error.code==='42501');check('inactive accounts cannot read inventory');
    await auth(ADMIN);await assert.rejects(()=>db.exec("update public.tms_inventory_state set payload='{}'"),error=>error.code==='42501');check('direct unversioned authenticated writes are denied');
    await auth(null,'anon');await assert.rejects(snapshot,error=>error.code==='42501');check('anonymous callers cannot read private quantities or DOP valuation');
    console.log(JSON.stringify({checks:checks.length,productionWrites:0}));
  }finally{await db.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

/* Local disposable PostgreSQL only. Never accepts a remote database destination. */
const {Client} = require(process.env.TMS_PG_MODULE || 'pg');
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const base=process.env.TMS_PG_TEST_URL;
if(!base)throw new Error('Set TMS_PG_TEST_URL to a local disposable PostgreSQL server.');
const parsed=new URL(base);
if(!['127.0.0.1','localhost'].includes(parsed.hostname))throw new Error('Only loopback PostgreSQL test servers are accepted.');
const database='tms_sync_test_'+process.pid;
const adminId='00000000-0000-4000-8000-000000000001',warehouseId='00000000-0000-4000-8000-000000000002';
const connect=async url=>{const c=new Client({connectionString:url});await c.connect();return c;};
(async()=>{
  const root=await connect(base);let db;const peers=[];
  try {
    await root.query(`DO $$ BEGIN CREATE ROLE authenticated NOLOGIN; EXCEPTION WHEN duplicate_object THEN NULL; END $$; DO $$ BEGIN CREATE ROLE anon NOLOGIN; EXCEPTION WHEN duplicate_object THEN NULL; END $$;`);
    await root.query(`CREATE DATABASE ${database} ENCODING 'UTF8' TEMPLATE template0`);
    const url=new URL(base);url.pathname='/'+database;db=await connect(url.href);
    await db.query(`create schema auth;
      create table auth.users(id uuid primary key,email text);
      create function auth.uid() returns uuid language sql stable as $$ select nullif(current_setting('test.uid',true),'')::uuid $$;
      create function auth.jwt() returns jsonb language sql stable as $$ select jsonb_build_object('email',coalesce(current_setting('test.email',true),'')) $$;
      create function auth.role() returns text language sql stable as $$ select coalesce(nullif(current_setting('test.role',true),''),'service_role') $$;
      grant usage on schema auth to authenticated; grant execute on all functions in schema auth to authenticated;
      alter default privileges in schema public grant all on tables to authenticated;
      insert into auth.users values ('${adminId}','vinelis.garcia@hbyalvarez.com'),('${warehouseId}','warehouse@example.invalid');`);
    const directory=path.join(__dirname,'../supabase/migrations');
    for(const file of fs.readdirSync(directory).filter(x=>x.endsWith('.sql')).sort())await db.query(fs.readFileSync(path.join(directory,file),'utf8'));
    await db.query(fs.readFileSync(path.join(directory,'0008_versioned_cloud_sync.sql'),'utf8'));
    await db.query(`insert into tms_user_profiles(auth_user_id,nombre,email,rol,permisos_por_modulo) values ($1,'Almacén sintético','warehouse@example.invalid','Almacén','{"solicitudesAlmacen":{"ver":true,"editar":true},"__account":{"passwordChangeRequired":true,"passwordConfigured":false,"authNeedsConfirmation":true}}'::jsonb)`,[warehouseId]);
    async function actor(id,email){const c=await connect(url.href);peers.push(c);await c.query('set role authenticated');await c.query("select set_config('test.uid',$1,false),set_config('test.email',$2,false),set_config('test.role','authenticated',false)",[id,email]);return c;}
    const a=await actor(adminId,'vinelis.garcia@hbyalvarez.com'),b=await actor(adminId,'vinelis.garcia@hbyalvarez.com');
    const snapshot=async c=>(await c.query('select tms_read_snapshot() as snapshot')).rows[0].snapshot;
    const save=async(c,revision,tables,settings=[],account=null)=>(await c.query('select tms_save_snapshot($1,$2::jsonb,$3::jsonb,$4::jsonb) as revision',[revision,JSON.stringify(tables),JSON.stringify(settings),account===null?null:JSON.stringify(account)])).rows[0].revision;
    let initial=await snapshot(a);
    const next=await save(a,initial.revision,{tms_client_configs:[{codigo_cliente:'C1',nombre_cliente:'Cliente seguro'}]},[{clave:'app_state',valor:{calendarNotes:[{id:'A'}],costSettings:{sentinel:1}}}]);
    await assert.rejects(save(b,initial.revision,{tms_client_configs:[{codigo_cliente:'LOST'}]}),/TMS_CONFLICT/);
    assert.equal((await snapshot(a)).tables.tms_client_configs[0].codigo_cliente,'C1');
    await assert.rejects(save(a,next,{tms_client_configs:[{codigo_cliente:'PARTIAL'}],tms_order_lines:[{}]}));
    const rollback=await snapshot(a);assert.equal(rollback.revision,Number(next));assert.equal(rollback.tables.tms_client_configs[0].codigo_cliente,'C1');
    const concurrent=await Promise.allSettled([save(a,next,{},[{clave:'app_state',valor:{calendarNotes:[{id:'race-a'}]}}]),save(b,next,{},[{clave:'app_state',valor:{calendarNotes:[{id:'race-b'}]}}])]);
    assert.equal(concurrent.filter(x=>x.status==='fulfilled').length,1);assert.equal(concurrent.filter(x=>x.status==='rejected'&&/TMS_CONFLICT/.test(x.reason.message)).length,1);
    const warehouse=await actor(warehouseId,'warehouse@example.invalid');let before=await snapshot(warehouse);
    const ownBefore=before.tables.tms_user_profiles.find(row=>row.auth_user_id===warehouseId);
    const adminBefore=before.tables.tms_user_profiles.find(row=>row.auth_user_id===adminId);
    // Use the actual frontend metadata helper: stale nested flags must not
    // override explicit false values after completing the password flow.
    const source=fs.readFileSync(path.join(__dirname,'../tms-v2.js'),'utf8');
    const metaStart=source.indexOf('  function getProfileAccountMeta('),metaEnd=source.indexOf('\n  }',metaStart)+4;
    const profileMeta=new Function('text','return ('+source.slice(metaStart,metaEnd)+')')(value=>String(value??'').trim());
    const account=profileMeta({permisosPorModulo:ownBefore.permisos_por_modulo,passwordChangeRequired:false,passwordConfigured:true,authNeedsConfirmation:false,accountStatus:'Activo'});
    assert.equal(account.passwordChangeRequired,false);assert.equal(account.authNeedsConfirmation,false);
    await save(warehouse,before.revision,{},[],account);
    before=await snapshot(warehouse);
    const ownAfter=before.tables.tms_user_profiles.find(row=>row.auth_user_id===warehouseId);
    assert.equal(ownAfter.permisos_por_modulo.__account.passwordChangeRequired,false);
    assert.equal(ownAfter.permisos_por_modulo.__account.passwordConfigured,true);
    assert.equal(ownAfter.permisos_por_modulo.__account.authNeedsConfirmation,false);
    assert.deepEqual(ownAfter.permisos_por_modulo.solicitudesAlmacen,ownBefore.permisos_por_modulo.solicitudesAlmacen);
    for(const key of ['id','auth_user_id','nombre','email','rol','activo'])assert.equal(ownAfter[key],ownBefore[key]);
    assert.deepEqual(before.tables.tms_user_profiles.find(row=>row.auth_user_id===adminId),adminBefore,'Updating account flags must never target another profile');
    for(const account of [{rol:'Admin'},{id:adminBefore.id},{auth_user_id:adminId},{email:'admin@example.invalid'},{activo:false},{permisos_por_modulo:{}},{passwordChangeRequired:'false'},{accountStatus:{rol:'Admin'}}]){
      await assert.rejects(save(warehouse,before.revision,{},[],account),/cuenta|booleano/);
    }
    assert.equal((await snapshot(warehouse)).revision,before.revision,'Rejected identity or role payloads must not advance the revision');
    await warehouse.query("select set_config('tms.snapshot_write','on',false)");
    try { await assert.rejects(warehouse.query("update tms_user_profiles set rol='Admin' where auth_user_id=$1",[warehouseId]),/Solo administradores/); }
    finally { await warehouse.query("select set_config('tms.snapshot_write','off',false)"); }
    await save(warehouse,before.revision,{tms_warehouse_plan:[]},[{clave:'app_state',valor:{costSettings:{sentinel:'forbidden'},calendarNotes:[{id:'forbidden'}]}}]);
    const limited=await snapshot(a);assert.equal(limited.tables.tms_settings.find(x=>x.clave==='app_state').valor.costSettings.sentinel,1);
    await assert.rejects(save(warehouse,limited.revision,{tms_client_configs:[]}),/rol no puede/);
    const routeId='00000000-0000-4000-8000-000000000003';
    const routePayload={tms_route_configs:[{id:routeId,nombre:'PP'}],tms_order_lines:[{clave_unica:'route-line',cliente_id:'C1',ruta_id:routeId}]};
    const routeRevision=await save(a,limited.revision,routePayload);
    await save(a,routeRevision,routePayload);
    assert.equal((await snapshot(a)).tables.tms_order_lines[0].ruta_id,routeId,'Replacing routes must preserve order-line foreign keys');
    await db.query("update tms_user_profiles set rol='Operaciones' where auth_user_id=$1",[warehouseId]);
    const operations=await snapshot(warehouse);
    await save(warehouse,operations.revision,{},[{clave:'app_state',valor:{costSettings:{sentinel:'forbidden',tripHistory:{trip1:{lockedAt:'test'}}}}}]);
    const operationalSettings=(await snapshot(a)).tables.tms_settings.find(x=>x.clave==='app_state').valor.costSettings;
    assert.equal(operationalSettings.sentinel,1);assert.equal(operationalSettings.tripHistory.trip1.lockedAt,'test');
    await assert.rejects(a.query("insert into tms_order_lines(clave_unica,cliente_id) values ('old-tab','C1')"),/TMS_CLIENT_OUTDATED/);
    // Trusted maintenance writes invalidate cached versions as well.
    await db.query("insert into tms_order_lines(clave_unica,cliente_id) select 'row-'||g,'C1' from generate_series(1,1205) g");
    const big=await snapshot(a);assert.equal(big.tables.tms_order_lines.length,1206);assert.ok(big.revision>limited.revision);
    await assert.rejects(save(a,limited.revision,{}),/TMS_CONFLICT/);
    await assert.rejects(a.query('update tms_sync_state set revision=0'),/permission denied/);
    await db.query(`update tms_user_profiles set activo=false where auth_user_id=$1`,[warehouseId]);
    const last=await snapshot(a);await assert.rejects(save(warehouse,last.revision,{}),/perfil activo/);
    console.log('PostgreSQL sync passed: all migrations plus idempotent 0008, atomic rollback, concurrent CAS, role boundaries, self-account flags and guard, inactive profile rejection, old-client write rejection, maintenance revision triggers, 1206-row consistent read and route foreign keys.');
  } finally {
    await Promise.all(peers.map(c=>c.end()));if(db)await db.end();
    await root.query(`DROP DATABASE IF EXISTS ${database} WITH (FORCE)`);await root.end();
  }
})().catch(error=>{console.error(error);process.exitCode=1;});

/* Local PostgreSQL/WASM validation; all users and data below are fictional.
 * npm install --prefix /tmp/tms-sync-validation @electric-sql/pglite
 * TMS_PGLITE_RUNTIME=/tmp/tms-sync-validation node tests/session-sync-db.cjs
 * No network or production database calls are made by this test.
 */
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { createRequire } = require('node:module');
const runtime = process.env.TMS_PGLITE_RUNTIME || '/tmp/tms-sync-validation';
const runtimeRequire = createRequire(path.join(runtime, 'package.json'));
const { PGlite } = runtimeRequire('@electric-sql/pglite');
const { pgcrypto } = runtimeRequire('@electric-sql/pglite/contrib/pgcrypto');
const root = path.resolve(__dirname, '..');
const read = name => fs.readFileSync(path.join(root, name), 'utf8');
const ADMIN = '11111111-1111-4111-8111-111111111111';
const READER = '22222222-2222-4222-8222-222222222222';
const INACTIVE = '33333333-3333-4333-8333-333333333333';
const ROUTE = '44444444-4444-4444-8444-444444444444';

(async () => {
  const db = new PGlite({ extensions: { pgcrypto } });
  await db.waitReady;
  const checks = [];
  const check = name => { checks.push(name); console.log('PASS:', name); };
  const scalar = async (sql, args = []) => (await db.query(sql, args)).rows[0].result;
  const auth = async (id, email) => {
    await db.exec('reset role');
    await db.query("select set_config('request.jwt.claims', $1, false)", [JSON.stringify({sub:id,role:'authenticated',email})]);
    await db.query("select set_config('request.jwt.claim.sub', $1, false)", [id]);
    await db.query("select set_config('request.jwt.claim.role', 'authenticated', false)");
    await db.exec('set role authenticated');
  };
  const snapshot = () => scalar('select public.tms_get_snapshot() as result');
  const save = (revision, payload) => scalar('select public.tms_save_snapshot($1, $2::jsonb) as result', [revision, JSON.stringify(payload)]);
  const rejectCode = async (fn, code) => {
    await assert.rejects(fn, error => error.code === code);
  };
  try {
    await db.exec(`
      create role anon nologin;
      create role authenticated nologin;
      create role service_role nologin bypassrls;
      create schema auth;
      create table auth.users (id uuid primary key, email text not null);
      create function auth.jwt() returns jsonb language sql stable as
        $$ select coalesce(nullif(current_setting('request.jwt.claims', true), ''), '{}')::jsonb $$;
      create function auth.uid() returns uuid language sql stable as
        $$ select coalesce(nullif(current_setting('request.jwt.claim.sub', true), ''), auth.jwt()->>'sub')::uuid $$;
      create function auth.role() returns text language sql stable as
        $$ select coalesce(nullif(current_setting('request.jwt.claim.role', true), ''), auth.jwt()->>'role') $$;
      grant usage on schema auth, public to anon, authenticated, service_role;
      grant execute on all functions in schema auth to anon, authenticated, service_role;
      create publication supabase_realtime;
    `);
    // Match deployed helpers, not unapplied role-clamping migrations 0005–0007.
    for (const name of ['0001_tms_v2_schema.sql', '0002_admin_vinelis_and_client_flags.sql', '0003_role_write_policies.sql', '0004_atomic_replace_and_route_history.sql']) {
      await db.exec(read('supabase/migrations/' + name));
    }
    await db.exec(`
      select set_config('request.jwt.claims', '{"role":"service_role"}', false);
      delete from public.tms_user_profiles;
      insert into auth.users values
        ('${ADMIN}', 'admin@example.invalid'),
        ('${READER}', 'reader@example.invalid'),
        ('${INACTIVE}', 'inactive@example.invalid');
      insert into public.tms_user_profiles(auth_user_id,nombre,email,rol,activo,permisos_por_modulo) values
        ('${ADMIN}','Synthetic Admin','admin@example.invalid','Admin',true,'{}'),
        ('${READER}','Synthetic Reader','reader@example.invalid','Solo lectura',true,'{"__account":{"passwordChangeRequired":true}}'),
        ('${INACTIVE}','Synthetic Inactive','inactive@example.invalid','Admin',false,'{}');
      insert into public.tms_settings(clave,valor) values
        ('app_state','{}'),('role_permissions','{"Solo lectura":{},"Admin":{}}')
      on conflict(clave) do update set valor=excluded.valor;
      insert into public.tms_route_configs(id,nombre) values ('${ROUTE}','SYNTHETIC ROUTE');
      insert into public.tms_order_lines(clave_unica,cliente_id,ruta_id,ruta_nombre,metadata)
      select 'SYNTHETIC-'||n,'SYNTHETIC-'||n,'${ROUTE}','SYNTHETIC ROUTE','{"dataset":"current"}'
      from generate_series(1,1105) n;
      insert into public.tms_route_cost_history(fecha,ruta_id,ruta_nombre,camion)
      values ('2026-10-31','${ROUTE}','SYNTHETIC ROUTE','CAMION 1');
      grant select,insert,update,delete on all tables in schema public to authenticated,service_role;
      select set_config('request.jwt.claims','{}',false);
    `);
    await db.exec(read('supabase/migrations/0008_session_sync.sql'));
    check('migration executes on deployed schema and creates revision publication');
    assert.equal(await scalar("select count(*)::integer as result from pg_publication_tables where pubname='supabase_realtime' and tablename='tms_sync_state'"), 1);

    await auth(ADMIN, 'admin@example.invalid');
    let before = await snapshot();
    assert.equal(before.tables.tms_order_lines.length, 1105);
    check('consistent snapshot returns all 1,105 rows beyond API page limit');
    await db.exec('reset role');
    await db.exec(read('tests/session-sync-rollback.sql'));
    await db.exec('reset role');
    // The rollback test must leave both rows and revision exactly intact.
    await auth(ADMIN, 'admin@example.invalid');
    assert.deepEqual(await snapshot(), before);
    check('rollback SQL tests: CAS, legacy guard, no-op saves, atomic rollback, reader restrictions');

    // Whole-table no-op detection must ignore generated metadata and preserve
    // duplicate multiplicity, rather than relying on JSON array containment.
    let result = await save(before.revision, {tms_order_lines:before.tables.tms_order_lines});
    assert.deepEqual(result, {revision:before.revision, changed:false});
    check('identical full table does not replace rows or advance revision');

    const routes = structuredClone(before.tables.tms_route_configs);
    routes[0].observaciones = 'Synthetic concurrent-route test';
    result = await save(before.revision, {tms_route_configs:routes});
    assert.equal(result.revision, before.revision + 1);
    let after = await snapshot();
    assert.deepEqual(after.tables.tms_order_lines, before.tables.tms_order_lines);
    assert.deepEqual(after.tables.tms_route_cost_history, before.tables.tms_route_cost_history);
    check('selective route change preserves foreign keys of unchanged orders and history');

    const staleRevision = after.revision;
    const payloadA = {tms_settings:[{clave:'app_state',valor:{session:'A'}}]};
    const payloadB = {tms_settings:[{clave:'app_state',valor:{session:'B'}}]};
    const concurrent = await Promise.allSettled([save(staleRevision,payloadA),save(staleRevision,payloadB)]);
    assert.equal(concurrent.filter(x=>x.status==='fulfilled').length, 1);
    assert.equal(concurrent.filter(x=>x.status==='rejected' && x.reason.code==='40001').length, 1);
    check('two concurrent saves with one revision yield one commit and one explicit conflict');

    before = await snapshot();
    const withDuplicate = [...before.tables.tms_order_lines, structuredClone(before.tables.tms_order_lines[0])];
    result = await save(before.revision, {tms_order_lines:withDuplicate});
    assert.equal(result.changed, true);
    after = await snapshot();
    assert.equal(after.tables.tms_order_lines.length, 1106);
    result = await save(after.revision, {tms_order_lines:after.tables.tms_order_lines});
    assert.equal(result.changed, false);
    result = await save(after.revision, {tms_order_lines:before.tables.tms_order_lines});
    assert.equal(result.changed, true);
    assert.equal((await snapshot()).tables.tms_order_lines.length, 1105);
    check('no-op comparison preserves duplicate multiplicity when adding or removing split lines');

    // New synchronization still permits the existing own-account completion
    // path, while rejecting attempted role or module permission escalation.
    await auth(READER, 'reader@example.invalid');
    before = await snapshot();
    const own = structuredClone(before.tables.tms_user_profiles.find(x=>x.auth_user_id===READER));
    own.permisos_por_modulo.__account = {passwordChangeRequired:false,passwordConfigured:true};
    result = await save(before.revision, {tms_user_profiles:[own]});
    assert.equal(result.changed, true);
    after = await snapshot();
    const updated = after.tables.tms_user_profiles.find(x=>x.auth_user_id===READER);
    assert.equal(updated.rol, 'Solo lectura');
    assert.deepEqual(updated.permisos_por_modulo.__account, own.permisos_por_modulo.__account);
    own.rol = 'Admin';
    await rejectCode(()=>save(after.revision,{tms_user_profiles:[own]}),'42501');
    assert.deepEqual(await snapshot(),after);
    check('reader can update only own account status and cannot elevate permissions');

    await auth(INACTIVE, 'inactive@example.invalid');
    await rejectCode(snapshot,'42501');
    await rejectCode(()=>save(0,{}),'42501');
    check('inactive profile cannot read snapshot or save');
    await db.exec('reset role');
    await db.query("select set_config('request.jwt.claims','{}',false)");
    await db.query("select set_config('request.jwt.claim.sub','',false)");
    await db.query("select set_config('request.jwt.claim.role','',false)");
    await db.exec('set role anon');
    await rejectCode(snapshot,'42501');
    check('anonymous caller cannot execute snapshot');

    // Maintenance is intentionally available to admin/service tooling, but
    // every statement invalidates the revision seen by open sessions.
    await db.exec('reset role');
    const maintenanceBefore = await scalar('select revision as result from public.tms_sync_state where id');
    await db.exec("update public.tms_settings set valor=valor||'{\"maintenance\":true}'::jsonb where clave='app_state'");
    assert.equal(await scalar('select revision as result from public.tms_sync_state where id'), maintenanceBefore + 1);
    check('administrative maintenance statements increment revision');

    await db.query("select set_config('request.jwt.claims',$1,false)", [JSON.stringify({sub:ADMIN,role:'service_role'})]);
    await db.query("select set_config('request.jwt.claim.role','service_role',false)");
    await db.exec('set role service_role');
    await db.exec("update public.tms_settings set valor=valor||'{\"service_maintenance\":true}'::jsonb where clave='app_state'");
    await db.exec('reset role');
    assert.equal(await scalar('select revision as result from public.tms_sync_state where id'), maintenanceBefore + 2);
    check('service-role maintenance is allowed and also invalidates open sessions');

    const beforeReapply = await scalar('select revision as result from public.tms_sync_state where id');
    await db.exec(read('supabase/migrations/0008_session_sync.sql'));
    assert.equal(await scalar('select revision as result from public.tms_sync_state where id'), beforeReapply);
    assert.equal(await scalar("select count(*)::integer as result from pg_publication_tables where pubname='supabase_realtime' and tablename='tms_sync_state'"), 1);
    check('reapplying migration preserves revision, operational data and one publication entry');

    console.log(JSON.stringify({productionWrites:0,checks:checks.length,passed:checks},null,2));
  } finally {
    await db.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });

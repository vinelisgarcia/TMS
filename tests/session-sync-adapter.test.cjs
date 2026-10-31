// Real TMS snapshot/payload adapters + real sync controller, synthetic RPC/UI.
// These tests never contact a browser, Supabase, or any production endpoint.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const Sync = require('../tms-sync.js');
const source = fs.readFileSync(path.join(__dirname, '../tms-v2.js'), 'utf8');
const copy = value => JSON.parse(JSON.stringify(value));

function appFunction(name) {
  const expression = new RegExp('^  (?:async )?function ' + name + '\\(', 'm');
  const match = expression.exec(source);
  assert.ok(match, `Missing adapter function ${name}`);
  const end = source.indexOf('\n  }', match.index) + 4;
  assert.ok(end > match.index, `Unterminated adapter function ${name}`);
  return source.slice(match.index, end);
}

function windowFunction(name) {
  const start = source.indexOf('  window.' + name + ' = ');
  assert.ok(start >= 0, `Missing entry point ${name}`);
  const end = source.indexOf('\n  };', start) + 5;
  assert.ok(end > start, `Unterminated entry point ${name}`);
  return source.slice(start, end);
}

const functions = [
  'text', 'num', 'safeClone', 'nowIso', 'isUuid', 'normalizeVisualTheme',
  'getProfileAccountMeta', 'permissionsPayloadForProfile',
  'buildBaseItemKey', 'buildClientConfigDefault',
  'mapDbOrderLine', 'buildOrderLineDbRows', 'mapDbWarehousePlan', 'mapDbWarehouseControl',
  'buildWarehousePlanDbRows', 'buildWarehouseControlDbRows',
  'getProfileForAuthUser', 'setCurrentUserProfileFromAuth',
  'hasPermission', 'isAdminUser', 'userProfileDbPayload',
  'saveLocalSnapshot', 'buildCloudPayload', 'applyCloudSnapshot', 'hasUnsubmittedDraft', 'trackUnsubmittedDraft', 'captureFormDrafts', 'restoreFormDrafts', 'getCloudSync',
];
const adapterCode = functions.map(appFunction).concat([
  windowFunction('cargarDesdeSupabase'), windowFunction('guardarEnSupabase'),
]).join('\n');

const identities = {
  admin: { id: '11111111-1111-4111-8111-111111111111', email: 'admin@example.invalid' },
  readonly: { id: '22222222-2222-4222-8222-222222222222', email: 'reader@example.invalid' },
  warehouse: { id: '33333333-3333-4333-8333-333333333333', email: 'warehouse@example.invalid' },
};
const permissions = {
  Admin: {},
  'Solo lectura': { comercial: { ver: true, editar: false } },
  'Almacén': { solicitudesAlmacen: { editar: true }, almacen: { editar: true } },
};

function snapshot(size = 3) {
  const tables = {
    tms_order_lines: Array.from({ length: size }, (_, index) => ({
      id: `db-order-${index}`, clave_unica: `TEST-${index}`, cliente_id: `C${index}`,
      cliente_nombre: `Synthetic ${index}`, pedido_cliente: `P${index}`, articulo: 'A',
      linea_pedido_cliente: '1', descripcion_articulo: 'Synthetic item',
      fecha_planificada: '2026-10-31', fecha_control: null, camion_asignado: 'CAMION 1',
      ruta_nombre: 'TEST ROUTE', cantidad_solicitada: 3, cantidad_facturada: 0, cantidad_pendiente: 3,
      metadata: { dataset: index % 3 === 0 ? 'plan' : index % 3 === 1 ? 'control' : 'current', incidencia: '', montoPlanificado: 30 },
    })),
    tms_client_configs: [{ codigo_cliente: 'C0', nombre_cliente: 'Synthetic 0', ruta_asignada: 'TEST ROUTE' }],
    tms_route_configs: [{ id: '44444444-4444-4444-8444-444444444444', nombre: 'TEST ROUTE', zona: 'TEST ROUTE', dias_operacion: [1], coste_ruta: 30, activa: true }],
    tms_route_cost_history: [], tms_warehouse_plan: [], tms_warehouse_control: [], tms_warehouse_history: [],
    tms_user_profiles: Object.entries(identities).map(([name, identity]) => ({
      id: `profile-${name}`, auth_user_id: identity.id, email: identity.email, nombre: name,
      rol: name === 'admin' ? 'Admin' : name === 'warehouse' ? 'Almacén' : 'Solo lectura',
      permisos_por_modulo: {}, activo: true,
    })),
    tms_settings: [
      { clave: 'role_permissions', valor: copy(permissions) },
      { clave: 'warehouse', valor: { costoSolicitud: 10 } },
      { clave: 'app_state', valor: { planLoaded: true, planFecha: '31/10/2026', visualTheme: 'light', calendarNotes: [] } },
    ],
  };
  return { revision: 1, tables };
}

function server(data = snapshot()) {
  const db = copy(data), calls = [];
  let unavailable = false;
  return {
    db, calls,
    setUnavailable(value) { unavailable = value; },
    client: {
      from() { throw new Error('Unsafe direct-table access attempted'); },
      async rpc(name, args) {
        calls.push({ name, args: copy(args || {}) });
        if (unavailable) return { data: null, error: { code: 'PGRST202', message: 'Snapshot RPC unavailable in synthetic test' } };
        if (name === 'tms_get_snapshot') return { data: copy(db), error: null };
        assert.equal(name, 'tms_save_snapshot');
        if (Number(args.p_expected_revision) !== db.revision) return { data: null, error: { code: '40001', message: 'Synthetic stale snapshot' } };
        for (const [table, rows] of Object.entries(args.p_payload)) {
          if (table === 'tms_settings') {
            for (const row of rows) {
              const index = db.tables.tms_settings.findIndex(item => item.clave === row.clave);
              if (index < 0) db.tables.tms_settings.push(copy(row));
              else db.tables.tms_settings[index] = copy(row);
            }
          } else db.tables[table] = copy(rows);
        }
        return { data: { revision: ++db.revision, changed: true }, error: null };
      },
      channel() { return { on() { return this; }, subscribe() { return this; } }; },
      removeChannel() { return Promise.resolve('ok'); },
    },
    saves() { return calls.filter(call => call.name === 'tms_save_snapshot'); },
  };
}

function adapter(db, account = 'admin', { quota = false } = {}) {
  const identity = identities[account];
  const stored = new Map(), statuses = [], warnings = [];
  const APP = { lineItems: [], planLineItems: [], controlLineItems: [], routeConfigs: [], userProfiles: [],
    routeCostHistory: [], solicitudesPlanAlmacen: [], solicitudesControlAlmacen: [], solicitudesHistory: [],
    clienteConfig: {}, warehouseSettings: {}, rolePermissions: copy(permissions), importFiles: {}, feriadosRD: [] };
  const context = vm.createContext({
    APP, cloudUser: null, cloudSync: null, unsubmittedDrafts: new Set(), THEME_KEY: 'theme', LOCAL_KEY: 'tms_alvarez_v2_state',
    window: { TMSSessionSync: Sync },
    console: { warn: (...args) => warnings.push(args), error: (...args) => warnings.push(args) },
    document: { getElementById: () => null },
    localStorage: {
      getItem: key => stored.get(key) || null,
      setItem(key, value) {
        if (quota) throw Object.assign(new Error('Synthetic storage quota exceeded'), { name: 'QuotaExceededError' });
        stored.set(key, value);
      },
      removeItem: key => stored.delete(key),
    },
    localSnapshotKey: () => 'tms_alvarez_v2_state:' + identity.id,
    getAuthenticatedUser: async () => copy(identity), getSupabaseClient: () => db.client,
    showSyncStatus: state => statuses.push(copy(state)),
    getPermissionsForRole: role => copy(APP.rolePermissions[role] || {}),
    buildDefaultRolePermissions: () => copy(permissions), buildPermissions: () => ({}),
    getRouteConfigByName: name => APP.routeConfigs.find(route => route.nombre === name) || null,
    deriveItemStatus: () => 'pendiente', chooseTruckForItem: () => 'CAMION 1',
    fechaStrToISO: date => !date ? null : date.includes('/') ? date.split('/').reverse().join('-') : date,
    isoToFechaStr: date => !date ? '' : date.includes('-') ? date.split('-').reverse().join('/') : date,
    buildSolicitudesComparison: () => [],
    ensureAdminProfile: () => {}, ensureRolePermissions: () => {},
    enforcePasswordChangeIfNeeded: async () => {}, applyVisualTheme: () => {},
    rebuildDerivedState: () => {}, updateSolicitudesImportStatus: () => {}, actualizarEstadoBanner: () => {},
    actualizarDashboard: () => {}, renderCalendario: () => {}, renderRutas: () => {},
    renderRouteCatalog: () => {}, renderComercial: () => {}, renderConfigClientes: () => {},
    renderSolicitudesAlmacen: () => {}, renderAlmacen: () => {}, renderPrioridades: () => {},
    renderMaterialesTransito: () => {}, renderIncidenciasDespacho: () => {}, applyPermissionUi: () => {},
  });
  vm.runInContext(adapterCode, context, { timeout: 2000 });
  return { context, APP, statuses, stored, warnings,
    load: () => context.window.cargarDesdeSupabase(),
    save: () => context.window.guardarEnSupabase(),
    stop: () => context.cloudSync?.stop(),
    capture: () => copy(context.buildCloudPayload()),
  };
}

test('real admin adapter loads shared data without autosaving', async t => {
  const db = server(), app = adapter(db);
  t.after(app.stop);
  await app.load();
  assert.equal(app.APP.currentUserProfile.rol, 'Admin');
  assert.equal(app.APP.lineItems[0].clienteId, 'C2');
  assert.equal(app.APP.planLineItems[0].clienteId, 'C0');
  assert.equal(app.APP.controlLineItems[0].clienteId, 'C1');
  assert.equal(db.saves().length, 0);
  assert.equal(app.context.cloudSync.getState().revision, 1);
  assert.equal(app.context.cloudSync.getState().dirty, false);
  await app.save();
  assert.equal(db.saves().length, 0, 'admin no-op save must not publish a replacement');
  assert.ok([...app.stored.keys()].every(key => key.endsWith(identities.admin.id)));
});

test('read-only adapter can load but never captures operational tables to write', async t => {
  const db = server(), app = adapter(db, 'readonly');
  t.after(app.stop);
  await app.load();
  assert.equal(app.APP.currentUserProfile.rol, 'Solo lectura');
  assert.deepEqual(Object.keys(app.capture()), ['tms_user_profiles']);
  assert.equal(app.capture().tms_user_profiles.length, 1);
  assert.equal(app.capture().tms_user_profiles[0].email, identities.readonly.email);
  await app.save();
  assert.equal(db.saves().length, 0);
});

test('a non-admin account status edit preserves its protected stored permission matrix', async t => {
  const data = snapshot();
  const storedPermissions = { protectedCustomModule: { ver: false, editar: false }, __account: { passwordChangeRequired: true, passwordConfigured: true } };
  data.tables.tms_user_profiles.find(row => row.email === identities.readonly.email).permisos_por_modulo = copy(storedPermissions);
  const db = server(data), app = adapter(db, 'readonly');
  t.after(app.stop);
  await app.load();
  app.APP.currentUserProfile.passwordChangeRequired = false;
  const payload = app.capture().tms_user_profiles[0].permisos_por_modulo;
  assert.deepEqual(payload.protectedCustomModule, storedPermissions.protectedCustomModule);
  assert.equal(payload.comercial, undefined, 'self-account metadata must not rebuild protected permissions from role templates');
  assert.equal(payload.__account.passwordChangeRequired, false);
});

test('the real draft tracker protects unsubmitted operational and account forms', () => {
  const app = adapter(server());
  for (const id of ['auditClientName', 'cfgContacto', 'routeName_0', 'newRouteNameInput', 'newRouteCostInput',
    'newRoleNameInput', 'userNameInput', 'userEmailInput', 'userPasswordInput', 'userRoleInput', 'warehouseCostInput']) {
    app.context.unsubmittedDrafts.clear();
    app.context.trackUnsubmittedDraft({ target: { id, isConnected: true, matches: () => true, closest: selector => selector === '#mainLayout' ? {} : null, classList: { contains: () => false } } });
    assert.equal(app.context.hasUnsubmittedDraft(), true, `Unsubmitted ${id} must defer automatic re-render`);
  }
});

test('form recovery never captures a password value', () => {
  const app = adapter(server());
  app.context.unsubmittedDrafts.add({ id: 'userPasswordInput', isConnected: true, type: 'password', value: 'SYNTHETIC-PASSWORD-DO-NOT-BACKUP' });
  app.context.unsubmittedDrafts.add({ id: 'userNameInput', isConnected: true, type: 'text', value: 'Synthetic person' });
  const captured = copy(app.context.captureFormDrafts());
  assert.equal(captured.length, 1);
  assert.equal(captured[0].field, 'userNameInput');
  assert.doesNotMatch(JSON.stringify(captured), /SYNTHETIC-PASSWORD/);
});

test('warehouse adapter emits only the tables permitted for warehouse work', async t => {
  const db = server(), app = adapter(db, 'warehouse');
  t.after(app.stop);
  await app.load();
  assert.deepEqual(Object.keys(app.capture()).sort(), [
    'tms_settings', 'tms_user_profiles', 'tms_warehouse_control', 'tms_warehouse_history', 'tms_warehouse_plan',
  ]);
  assert.deepEqual(app.capture().tms_settings.map(row => row.clave).sort(), ['app_state', 'warehouse']);
});

test('real snapshot mapper retains all datasets above 1000 rows through a save', async t => {
  const db = server(snapshot(2403)), app = adapter(db);
  t.after(app.stop);
  await app.load();
  assert.equal(app.APP.lineItems.length, 801);
  assert.equal(app.APP.planLineItems.length, 801);
  assert.equal(app.APP.controlLineItems.length, 801);
  app.APP.lineItems[0].incidencia = 'SYNTHETIC CHANGE';
  app.context.cloudSync.markDirty();
  await app.save();
  assert.equal(db.saves().length, 1);
  assert.equal(db.saves()[0].args.p_payload.tms_order_lines.length, 2403);
  assert.equal(db.db.tables.tms_order_lines.filter(row => row.metadata.incidencia === 'SYNTHETIC CHANGE').length, 1);
  assert.equal(app.APP.lineItems.length + app.APP.planLineItems.length + app.APP.controlLineItems.length, 2403);
});

test('missing snapshot migration fails closed instead of falling back to legacy writes or demo data', async t => {
  const db = server(), app = adapter(db);
  t.after(app.stop);
  db.setUnavailable(true);
  await app.load();
  assert.equal(app.APP.lineItems.length, 0);
  assert.ok(app.statuses.some(state => state.phase === 'offline'));
  await assert.rejects(app.save());
  assert.equal(db.saves().length, 0);
  assert.ok(db.calls.every(call => call.name === 'tms_get_snapshot'));
});

test('two real adapters preserve the losing edit when simultaneous snapshots conflict', async t => {
  const db = server(snapshot(6)), a = adapter(db), b = adapter(db);
  t.after(() => { a.stop(); b.stop(); });
  await Promise.all([a.load(), b.load()]);
  a.APP.lineItems[0].incidencia = 'ADAPTER-A'; a.context.cloudSync.markDirty();
  b.APP.lineItems[1].incidencia = 'ADAPTER-B'; b.context.cloudSync.markDirty();
  const results = await Promise.allSettled([a.save(), b.save()]);
  assert.equal(results.filter(result => result.status === 'fulfilled').length, 1);
  assert.equal(results.filter(result => result.status === 'rejected').length, 1);
  const loser = results[0].status === 'rejected' ? a : b;
  assert.equal(loser.context.cloudSync.getState().conflict, true);
  assert.match(JSON.stringify(loser.context.cloudSync.exportPending()), /ADAPTER-[AB]/);
  assert.equal(db.db.tables.tms_order_lines.filter(row => /^ADAPTER-/.test(row.metadata.incidencia || '')).length, 1);
});

test('full local storage does not prevent the real adapter from loading or committing to cloud', async t => {
  const db = server(), app = adapter(db, 'admin', { quota: true });
  t.after(app.stop);
  await app.load();
  assert.equal(app.context.cloudSync.getState().revision, 1);
  app.APP.lineItems[0].incidencia = 'CLOUD SURVIVES FULL STORAGE';
  app.context.cloudSync.markDirty();
  await app.save();
  assert.equal(db.saves().length, 1);
  assert.equal(db.db.tables.tms_order_lines.filter(row => row.metadata.incidencia === 'CLOUD SURVIVES FULL STORAGE').length, 1);
});

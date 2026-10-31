// Synthetic sessions and an in-memory RPC server. No production network calls.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { create } = require('../tms-sync.js');

const copy = value => JSON.parse(JSON.stringify(value));
const user = id => ({ id, email: `${id}@example.invalid` });
const tick = async () => {
  for (let i = 0; i < 5; i++) await new Promise(resolve => setImmediate(resolve));
};
const initialTables = () => ({
  tms_order_lines: [{ clave_unica: 'A', metadata: { incidencia: '' } }, { clave_unica: 'B', metadata: { incidencia: '' } }],
  tms_settings: [{ clave: 'app_state', valor: { calendarNotes: [] } }, { clave: 'warehouse', valor: { costoSolicitud: 10 } }],
});

function storage() {
  const values = new Map();
  return {
    getItem: key => values.has(key) ? values.get(key) : null,
    setItem: (key, value) => values.set(key, String(value)),
    removeItem: key => values.delete(key),
    values,
  };
}

function events() {
  const listeners = new Map();
  return {
    addEventListener(name, callback) {
      if (!listeners.has(name)) listeners.set(name, new Set());
      listeners.get(name).add(callback);
    },
    removeEventListener(name, callback) { listeners.get(name)?.delete(callback); },
    emit(name) { for (const fn of listeners.get(name) || []) fn({ type: name }); },
    listenerCount() { return [...listeners.values()].reduce((total, set) => total + set.size, 0); },
    document: { visibilityState: 'visible' },
    navigator: { onLine: true },
  };
}

function server(tables = initialTables()) {
  const db = { revision: 1, tables: copy(tables) };
  const calls = [];
  const channels = new Set();
  let offline = false;
  let beforeSave = null;
  let beforeGet = null;
  function publish() {
    for (const channel of channels) {
      for (const listener of channel.listeners) {
        listener.callback({ eventType: 'UPDATE', new: { id: true, revision: db.revision }, old: {} });
      }
    }
  }
  function client() {
    return {
      async rpc(method, args) {
        calls.push({ method, args: copy(args || {}) });
        if (offline) return { data: null, error: { code: 'NETWORK', message: 'Synthetic offline transport' } };
        if (method === 'tms_get_snapshot') {
          const snapshot = copy(db);
          if (beforeGet) await beforeGet();
          return { data: snapshot, error: null };
        }
        assert.equal(method, 'tms_save_snapshot');
        if (beforeSave) await beforeSave(args);
        if (Number(args.p_expected_revision) !== db.revision) {
          return { data: null, error: { code: '40001', message: 'Another session saved first' } };
        }
        for (const [table, rows] of Object.entries(args.p_payload)) {
          if (table === 'tms_settings') {
            for (const row of rows) {
              const index = db.tables.tms_settings.findIndex(item => item.clave === row.clave);
              if (index < 0) db.tables.tms_settings.push(copy(row));
              else db.tables.tms_settings[index] = copy(row);
            }
          } else db.tables[table] = copy(rows);
        }
        const changed = Object.keys(args.p_payload).length > 0;
        if (changed) db.revision++;
        return { data: { revision: db.revision, changed }, error: null };
      },
      channel() {
        const channel = {
          listeners: [], status: null,
          on(type, filter, callback) { this.listeners.push({ type, filter, callback }); return this; },
          subscribe(callback) { this.status = callback; channels.add(this); callback?.('SUBSCRIBED'); return this; },
          unsubscribe() { channels.delete(this); return Promise.resolve('ok'); },
        };
        return channel;
      },
      removeChannel(channel) { channels.delete(channel); return Promise.resolve('ok'); },
    };
  }
  return {
    db, calls, channels, client, publish,
    setOffline(value) { offline = value; },
    setBeforeSave(value) { beforeSave = value; },
    setBeforeGet(value) { beforeGet = value; },
    saves() { return calls.filter(call => call.method === 'tms_save_snapshot'); },
  };
}

function session(db, options = {}) {
  let tables = {};
  let draft = false;
  const status = [];
  const eventTarget = events();
  const client = db.client();
  const tabStorage = options.sessionStorage || storage();
  const controller = create({
    getClient: () => client,
    capture: () => copy(tables),
    applySnapshot: async snapshot => { tables = copy(snapshot.tables); },
    onStatus: state => status.push(copy(state)),
    storage: options.storage || storage(),
    sessionStorage: tabStorage,
    sessionId: options.sessionId,
    hasDraft: () => draft,
    captureDrafts: () => draft ? [{ field: 'syntheticForm', value: 'UNSUBMITTED-DRAFT' }] : [],
    restoreDrafts: drafts => { draft = !!drafts.length; },
    eventTarget,
    pollMs: 60 * 60 * 1000,
  });
  return {
    controller, status, eventTarget, tabStorage,
    setDraft(value) { draft = value; controller.markDirty(); },
    get tables() { return tables; },
    edit(index, value) {
      tables.tms_order_lines[index].metadata.incidencia = value;
      controller.markDirty();
    },
  };
}

for (const sameAccount of [false, true]) {
  test(`simultaneous ${sameAccount ? 'same-account sessions' : 'different users'} cannot overwrite an earlier save`, async t => {
    const db = server();
    const a = session(db), b = session(db);
    t.after(() => { a.controller.stop(); b.controller.stop(); });
    await Promise.all([a.controller.start(user('a')), b.controller.start(user(sameAccount ? 'a' : 'b'))]);
    assert.equal(db.saves().length, 0, 'opening a session must never save');
    a.edit(0, 'SESSION-A'); b.edit(1, 'SESSION-B');
    const results = await Promise.allSettled([a.controller.save(), b.controller.save()]);
    assert.equal(results.filter(result => result.status === 'fulfilled').length, 1);
    assert.equal(results.filter(result => result.status === 'rejected').length, 1);
    const rejected = results.find(result => result.status === 'rejected');
    assert.equal(rejected.reason.code, '40001');
    const loser = results[0].status === 'rejected' ? a : b;
    assert.equal(loser.controller.getState().dirty, true);
    assert.equal(loser.controller.getState().conflict, true);
    const exportText = JSON.stringify(loser.controller.exportPending());
    assert.match(exportText, /SESSION-[AB]/);
    const persisted = JSON.stringify(db.db.tables.tms_order_lines);
    assert.equal((persisted.match(/SESSION-/g) || []).length, 1, 'exactly the first committed snapshot remains');
  });
}

test('changes made while saving are serialized into a second immutable save', async t => {
  const db = server();
  const a = session(db);
  t.after(() => a.controller.stop());
  await a.controller.start(user('a'));
  let release;
  let arrived;
  const atServer = new Promise(resolve => { arrived = resolve; });
  const barrier = new Promise(resolve => { release = resolve; });
  let calls = 0;
  db.setBeforeSave(async () => { if (++calls === 1) { arrived(); await barrier; } });
  a.edit(0, 'FIRST');
  const first = a.controller.save();
  await atServer;
  a.edit(1, 'SECOND');
  const second = a.controller.save();
  release();
  await Promise.all([first, second]);
  assert.equal(db.saves().length, 2);
  assert.equal(db.saves()[0].args.p_payload.tms_order_lines[1].metadata.incidencia, '', 'in-flight payload was copied');
  assert.equal(db.saves()[1].args.p_expected_revision, 2);
  assert.equal(db.db.tables.tms_order_lines[0].metadata.incidencia, 'FIRST');
  assert.equal(db.db.tables.tms_order_lines[1].metadata.incidencia, 'SECOND');
  assert.equal(a.controller.getState().dirty, false);
});

test('clean sessions receive remote changes, while dirty sessions retain their edits', async t => {
  const db = server();
  const clean = session(db), dirty = session(db);
  t.after(() => { clean.controller.stop(); dirty.controller.stop(); });
  await clean.controller.start(user('clean'));
  await dirty.controller.start(user('dirty'));
  dirty.edit(0, 'LOCAL-PENDING');
  db.db.tables.tms_order_lines[1].metadata.incidencia = 'REMOTE'; db.db.revision++;
  db.publish();
  await tick();
  assert.equal(clean.tables.tms_order_lines[1].metadata.incidencia, 'REMOTE');
  assert.equal(dirty.tables.tms_order_lines[0].metadata.incidencia, 'LOCAL-PENDING');
  assert.equal(dirty.controller.getState().dirty, true);
  assert.equal(dirty.controller.getState().conflict, true);
  assert.equal(db.saves().length, 0);
});

test('offline edits survive failed saves and reconnect detects remote conflicts', async t => {
  const db = server();
  const a = session(db);
  t.after(() => a.controller.stop());
  await a.controller.start(user('a'));
  a.edit(0, 'OFFLINE-CHANGE'); db.setOffline(true);
  await assert.rejects(a.controller.save());
  assert.equal(a.controller.getState().dirty, true);
  assert.equal(a.tables.tms_order_lines[0].metadata.incidencia, 'OFFLINE-CHANGE');
  db.setOffline(false); db.db.revision++;
  db.db.tables.tms_order_lines[1].metadata.incidencia = 'OTHER-SESSION';
  a.eventTarget.emit('online');
  await tick();
  assert.equal(a.tables.tms_order_lines[0].metadata.incidencia, 'OFFLINE-CHANGE');
  assert.equal(a.controller.getState().conflict, true);
  assert.match(JSON.stringify(a.controller.exportPending()), /OFFLINE-CHANGE/);
});

test('reconnect retries pending offline edits automatically when the server revision is unchanged', async t => {
  const db = server(), a = session(db);
  t.after(() => a.controller.stop());
  await a.controller.start(user('a'));
  await tick();
  a.edit(0, 'RETRY-WHEN-ONLINE');
  db.setOffline(true);
  await assert.rejects(a.controller.save());
  db.setOffline(false);
  a.eventTarget.emit('online');
  await tick();
  assert.equal(db.db.tables.tms_order_lines[0].metadata.incidencia, 'RETRY-WHEN-ONLINE');
  assert.equal(a.controller.getState().dirty, false);
  assert.equal(a.controller.getState().conflict, false);
});

test('a failed initial connection recovers when connectivity returns', async t => {
  const db = server(), a = session(db);
  t.after(() => a.controller.stop());
  db.setOffline(true);
  await assert.rejects(a.controller.start(user('a')));
  db.setOffline(false);
  a.eventTarget.emit('online');
  await tick();
  assert.equal(a.tables.tms_order_lines.length, 2);
  assert.equal(a.controller.getState().dirty, false);
  assert.ok(db.channels.size > 0, 'recovered session also reconnects Realtime');
});

test('an edit made during a refresh cannot be overwritten by the arriving snapshot', async t => {
  const db = server(), a = session(db);
  t.after(() => a.controller.stop());
  await a.controller.start(user('a'));
  await tick(); // Let the subscription's initial reconciliation finish first.
  db.db.revision++;
  db.db.tables.tms_order_lines[1].metadata.incidencia = 'REMOTE';
  let release, arrived;
  const barrier = new Promise(resolve => { release = resolve; });
  const inFlight = new Promise(resolve => { arrived = resolve; });
  db.setBeforeGet(async () => { arrived(); await barrier; });
  const refreshing = a.controller.refresh();
  await inFlight;
  a.edit(0, 'EDIT-DURING-REFRESH');
  release();
  await refreshing;
  assert.equal(a.tables.tms_order_lines[0].metadata.incidencia, 'EDIT-DURING-REFRESH');
  assert.equal(a.controller.getState().dirty, true);
  assert.equal(a.controller.getState().conflict, true);
});

test('pending recovery is namespaced by authenticated account', async t => {
  const db = server(), disk = storage(), tabStorage = storage();
  const original = session(db, { storage: disk, sessionStorage: tabStorage });
  await original.controller.start(user('a'));
  original.edit(0, 'RECOVER-ME');
  original.controller.stop();
  const restored = session(db, { storage: disk, sessionStorage: tabStorage });
  const other = session(db, { storage: disk });
  t.after(() => { restored.controller.stop(); other.controller.stop(); });
  await restored.controller.start(user('a'));
  assert.equal(restored.tables.tms_order_lines[0].metadata.incidencia, 'RECOVER-ME');
  assert.equal(restored.controller.getState().dirty, true);
  await other.controller.start(user('b'));
  assert.equal(other.tables.tms_order_lines[0].metadata.incidencia, '');
  assert.equal(other.controller.getState().dirty, false);
  assert.equal(db.saves().length, 0);
});

test('a clean same-account tab cannot erase another tab pending recovery', async t => {
  const db = server(), disk = storage(), tabA = storage(), tabB = storage();
  const dirty = session(db, { storage: disk, sessionStorage: tabA });
  const clean = session(db, { storage: disk, sessionStorage: tabB });
  t.after(() => { dirty.controller.stop(); clean.controller.stop(); });
  await dirty.controller.start(user('a'));
  dirty.edit(0, 'ONLY-IN-DIRTY-TAB');
  await clean.controller.start(user('a'));
  assert.equal(clean.tables.tms_order_lines[0].metadata.incidencia, '', 'a separate tab starts from the team snapshot');
  await clean.controller.refresh();
  await clean.controller.save();
  const recovered = session(db, { storage: disk, sessionStorage: tabA });
  t.after(() => recovered.controller.stop());
  await recovered.controller.start(user('a'));
  assert.equal(recovered.tables.tms_order_lines[0].metadata.incidencia, 'ONLY-IN-DIRTY-TAB');
  assert.equal(recovered.controller.getState().dirty, true);
});

test('pending edits survive consecutive reloads while offline and reconnect later', async t => {
  const db = server(), disk = storage(), tabStorage = storage();
  const original = session(db, { storage: disk, sessionStorage: tabStorage });
  t.after(() => original.controller.stop());
  await original.controller.start(user('a'));
  original.edit(0, 'OFFLINE-RELOAD-RECOVERY');
  db.setOffline(true);
  for (let attempt = 0; attempt < 2; attempt++) {
    const offline = session(db, { storage: disk, sessionStorage: tabStorage });
    t.after(() => offline.controller.stop());
    await assert.rejects(offline.controller.start(user('a')));
  }
  db.setOffline(false);
  const recovered = session(db, { storage: disk, sessionStorage: tabStorage });
  t.after(() => recovered.controller.stop());
  await recovered.controller.start(user('a'));
  assert.equal(recovered.tables.tms_order_lines[0].metadata.incidencia, 'OFFLINE-RELOAD-RECOVERY');
  assert.equal(recovered.controller.getState().dirty, true);
});

test('switching a to b to a in one browser tab retains each account recovery pointer', async t => {
  const db = server(), disk = storage(), tabStorage = storage();
  const firstA = session(db, { storage: disk, sessionStorage: tabStorage });
  t.after(() => firstA.controller.stop());
  await firstA.controller.start(user('a'));
  firstA.edit(0, 'ACCOUNT-A-RECOVERY');
  const accountB = session(db, { storage: disk, sessionStorage: tabStorage });
  t.after(() => accountB.controller.stop());
  await accountB.controller.start(user('b'));
  assert.equal(accountB.tables.tms_order_lines[0].metadata.incidencia, '');
  const finalA = session(db, { storage: disk, sessionStorage: tabStorage });
  t.after(() => finalA.controller.stop());
  await finalA.controller.start(user('a'));
  assert.equal(finalA.tables.tms_order_lines[0].metadata.incidencia, 'ACCOUNT-A-RECOVERY');
  assert.equal(finalA.controller.getState().dirty, true);
});

test('a duplicated tab clearing its copy cannot delete the predecessor backup', async t => {
  const db = server(), disk = storage(), tabA = storage();
  const original = session(db, { storage: disk, sessionStorage: tabA });
  t.after(() => original.controller.stop());
  await original.controller.start(user('a'));
  original.edit(0, 'DUPLICATE-SAFE');
  const duplicatedTab = storage();
  for (const [key, value] of tabA.values) duplicatedTab.setItem(key, value);
  const duplicate = session(db, { storage: disk, sessionStorage: duplicatedTab });
  t.after(() => duplicate.controller.stop());
  await duplicate.controller.start(user('a'));
  await duplicate.controller.discardAndReload();
  const recovered = session(db, { storage: disk, sessionStorage: tabA });
  t.after(() => recovered.controller.stop());
  await recovered.controller.start(user('a'));
  assert.equal(recovered.tables.tms_order_lines[0].metadata.incidencia, 'DUPLICATE-SAFE');
});

test('remote refresh defers while a form has unsubmitted draft input outside APP', async t => {
  const db = server(), a = session(db);
  t.after(() => a.controller.stop());
  await a.controller.start(user('a'));
  await tick();
  a.setDraft(true);
  db.db.tables.tms_order_lines[0].metadata.incidencia = 'REMOTE-AFTER-DRAFT';
  db.db.revision++;
  db.publish(); await tick();
  assert.equal(a.tables.tms_order_lines[0].metadata.incidencia, '', 'draft DOM must not be destroyed by re-render');
  assert.equal(a.controller.getState().conflict, true);
  assert.equal(db.saves().length, 0);
});

test('only changed tables and changed settings keys are sent', async t => {
  const db = server(), a = session(db);
  t.after(() => a.controller.stop());
  await a.controller.start(user('a'));
  a.tables.tms_settings.find(row => row.clave === 'warehouse').valor.costoSolicitud = 20;
  a.controller.markDirty(); await a.controller.save();
  const payload = db.saves()[0].args.p_payload;
  assert.deepEqual(Object.keys(payload), ['tms_settings']);
  assert.deepEqual(payload.tms_settings.map(row => row.clave), ['warehouse']);
  assert.equal(db.db.tables.tms_settings.find(row => row.clave === 'app_state').valor.calendarNotes.length, 0);
});

test('snapshot loads more than 1000 rows completely and does not write merely by opening', async t => {
  const tables = initialTables();
  tables.tms_order_lines = Array.from({ length: 2401 }, (_, index) => ({ clave_unica: String(index), metadata: { incidencia: '' } }));
  const db = server(tables), a = session(db);
  t.after(() => a.controller.stop());
  await a.controller.start(user('readonly'));
  assert.equal(a.tables.tms_order_lines.length, 2401);
  assert.equal(db.saves().length, 0);
  await a.controller.save();
  assert.equal(db.saves().length, 0, 'saving without edits must be a local no-op');
});

test('explicit discard loads newest state and clears pending conflict', async t => {
  const db = server(), a = session(db);
  t.after(() => a.controller.stop());
  await a.controller.start(user('a'));
  a.edit(0, 'LOCAL');
  db.db.tables.tms_order_lines[0].metadata.incidencia = 'SERVER'; db.db.revision++;
  db.publish(); await tick();
  await a.controller.discardAndReload();
  assert.equal(a.tables.tms_order_lines[0].metadata.incidencia, 'SERVER');
  assert.equal(a.controller.getState().dirty, false);
  assert.equal(a.controller.getState().conflict, false);
});

test('stop removes channels and event listeners', async () => {
  const db = server(), a = session(db);
  await a.controller.start(user('a'));
  assert.ok(db.channels.size > 0);
  a.controller.stop();
  await tick();
  assert.equal(db.channels.size, 0);
  assert.equal(a.eventTarget.listenerCount(), 0);
});

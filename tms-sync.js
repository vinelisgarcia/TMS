(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.TMSSessionSync = api;
})(typeof globalThis === 'object' ? globalThis : this, function () {
  'use strict';
  const clone = value => JSON.parse(JSON.stringify(value));
  function stable(value) {
    if (Array.isArray(value)) return '[' + value.map(stable).join(',') + ']';
    if (value && typeof value === 'object') return '{' + Object.keys(value).sort().map(key => JSON.stringify(key) + ':' + stable(value[key])).join(',') + '}';
    return JSON.stringify(value);
  }
  function difference(base, next) {
    const changes = {};
    Object.keys(next).forEach(table => {
      if (table === 'tms_settings') {
        const before = new Map((base[table] || []).map(row => [row.clave, stable(row)]));
        const rows = next[table].filter(row => before.get(row.clave) !== stable(row));
        if (rows.length) changes[table] = clone(rows);
      } else if (stable(base[table]) !== stable(next[table])) changes[table] = clone(next[table]);
    });
    return changes;
  }
  function failure(error) {
    return Object.assign(new Error(error.message || 'No se pudo sincronizar.'), error);
  }
  function create(options) {
    const events = options.eventTarget || (typeof window !== 'undefined' ? window : null);
    const storage = options.storage || (typeof localStorage !== 'undefined' ? localStorage : null);
    const tabStorage = options.sessionStorage || (typeof sessionStorage !== 'undefined' ? sessionStorage : null);
    const instanceId = options.sessionId || Date.now().toString(36) + '-' + Math.random().toString(36).slice(2);
    let previousInstance = null;
    let user = null, base = null, revision = null, remoteRevision = null;
    let dirty = false, conflict = false, saving = null, refreshing = null;
    let channel = null, timer = null, generation = 0, lifecycle = 0, stopped = true;
    let phase = 'disconnected', lastError = null, storageError = null, refreshRequested = false;
    const key = (instance) => 'tms_alvarez_sync_pending:' + (user ? user.id : '') + ':' + (instance || instanceId);
    const pointerKey = () => 'tms_sync_instance:' + (user ? user.id : '');
    const hasDraft = () => !!(options.hasDraft && options.hasDraft());
    function getState() { return {userId:user && user.id, revision, remoteRevision, dirty, draft:hasDraft(), conflict, saving:!!saving || phase === 'saving', phase, error:lastError || storageError}; }
    function status(next, error) {
      phase = next;
      lastError = error ? error.message || String(error) : null;
      if (options.onStatus) options.onStatus(getState());
    }
    function readPending() {
      try { return storage && JSON.parse(storage.getItem(key()) || (previousInstance && storage.getItem(key(previousInstance))) || 'null'); }
      catch (_) { return null; }
    }
    function persist() {
      if (!storage || !user || !base) return;
      try {
        if (!dirty && !hasDraft()) storage.removeItem(key());
        else storage.setItem(key(), JSON.stringify(exportPending()));
        if (tabStorage) tabStorage.setItem(pointerKey(),instanceId);
        storageError = null;
      } catch (error) {
        // A quota error must be visible; the in-memory copy is still preserved.
        storageError = 'No se pudo guardar el respaldo local: ' + error.message;
        if (options.onStatus) options.onStatus(getState());
      }
    }
    async function rpc(name, args) {
      const client = options.getClient();
      if (!client) throw new Error('No hay conexión de Supabase.');
      const result = await client.rpc(name, args || {});
      if (result.error) throw failure(result.error);
      return result.data;
    }
    function markDirty() {
      generation++;
      if (!base || stopped) return;
      dirty = Object.keys(difference(base, options.capture())).length > 0;
      if (!dirty && !hasDraft()) conflict = false;
      persist();
      status(conflict ? 'conflict' : (dirty ? 'pending' : 'synced'));
    }
    async function apply(snapshot, token, keepPending) {
      if (token !== lifecycle || stopped) return;
      await options.applySnapshot(snapshot);
      if (token !== lifecycle || stopped) return;
      revision = Number(snapshot.revision);
      remoteRevision = Math.max(remoteRevision || 0, revision);
      base = clone(options.capture());
      dirty = false;
      conflict = false;
      if (!keepPending) persist();
      status('synced');
    }
    async function refresh(settings) {
      if (stopped || !user) return;
      if (refreshing) return refreshing;
      const token = lifecycle;
      const atStart = generation;
      const before = base && stable(options.capture());
      refreshing = (async () => {
        try {
          const snapshot = await rpc('tms_get_snapshot', {p_known_revision:settings && settings.force ? null : revision});
          if (token !== lifecycle || stopped) return;
          remoteRevision = Math.max(remoteRevision || 0, Number(snapshot.revision));
          if (Number(snapshot.revision) < remoteRevision) {
            refreshRequested = true;
            if (dirty) { conflict = true; persist(); status('conflict'); }
            return snapshot;
          }
          if (snapshot.unchanged) {
            if (dirty && remoteRevision !== revision) conflict = true;
            status(conflict ? 'conflict' : (dirty ? 'pending' : 'synced'));
            return snapshot;
          }
          if (dirty || hasDraft() || saving || generation !== atStart || (base && stable(options.capture()) !== before)) {
            if (remoteRevision !== revision) conflict = true;
            persist();
            status(conflict ? 'conflict' : 'pending');
            return snapshot;
          }
          await apply(snapshot, token);
          return snapshot;
        } catch (error) {
          if (token === lifecycle && !stopped) { persist(); status('offline', error); }
          throw error;
        }
      })();
      try { return await refreshing; } finally {
        if (token === lifecycle) {
          refreshing = null;
          if (refreshRequested && !stopped) {
            refreshRequested = false;
            Promise.resolve().then(backgroundRefresh);
          }
        }
      }
    }
    function backgroundRefresh() {
      if (stopped || !user) return;
      if (!base) {
        if (phase !== 'loading') start({id:user.id}).catch(()=>{});
        return;
      }
      if (refreshing) { refreshRequested = true; return; }
      const retryPending = dirty && phase === 'offline';
      refresh().then(() => {
        if (retryPending && dirty && !conflict && !stopped) return save();
      }).catch(() => {});
    }
    function watchRecovery() {
      if (!timer) {
        timer = setInterval(backgroundRefresh, options.pollMs || 8000);
        if (timer && timer.unref) timer.unref();
        if (events) ['online','focus','visibilitychange'].forEach(name=>events.addEventListener(name,backgroundRefresh));
      }
    }
    function remoteChange(payload) {
      const next = Number(payload && payload.new && payload.new.revision);
      if (Number.isFinite(next)) remoteRevision = Math.max(remoteRevision || 0, next);
      // A notification can arrive before the current RPC's response. Let its
      // acknowledgement update the revision before checking for a conflict.
      if (!saving) backgroundRefresh();
    }
    async function start(nextUser) {
      if (!nextUser || !nextUser.id) throw new Error('Se requiere una sesión autenticada.');
      if (!stopped && base && user && user.id === nextUser.id) return refresh();
      stop();
      user = {id:nextUser.id}; stopped = false;
      const token = lifecycle;
      try { previousInstance = tabStorage && tabStorage.getItem(pointerKey()); }
      catch (_) { previousInstance = null; }
      const pending = readPending();
      // Copy recovery before the first network await. Several consecutive
      // offline reloads must not move the tab pointer past its only backup.
      if (pending && pending.userId === user.id && pending.base && pending.payload && storage) {
        try {
          storage.setItem(key(),JSON.stringify(pending));
          if (tabStorage) tabStorage.setItem(pointerKey(),instanceId);
        } catch (error) { storageError = 'No se pudo actualizar el respaldo local: ' + error.message; }
      }
      status('loading');
      try {
        const snapshot = await rpc('tms_get_snapshot', {p_known_revision:null});
        if (token !== lifecycle || stopped) return;
        await apply(snapshot, token,!!pending);
        if (pending && pending.userId === user.id && pending.base && pending.payload) {
          const tables = {...snapshot.tables, ...pending.payload};
          if (pending.payload.tms_settings) {
            const settings = new Map((snapshot.tables.tms_settings || []).map(row=>[row.clave,row]));
            pending.payload.tms_settings.forEach(row=>settings.set(row.clave,row));
            tables.tms_settings = [...settings.values()];
          }
          await options.applySnapshot({revision:pending.revision,tables});
          if (token !== lifecycle || stopped) return;
          if (pending.formDrafts && options.restoreDrafts) options.restoreDrafts(pending.formDrafts);
          base = clone(pending.base);
          revision = Number(pending.revision);
          remoteRevision = Number(snapshot.revision);
          dirty = Object.keys(difference(base, options.capture())).length > 0;
          conflict = (dirty || hasDraft()) && revision !== remoteRevision;
          persist(); status(conflict ? 'conflict' : (dirty ? 'pending' : 'synced'));
        }
        const client = options.getClient();
        if (client && typeof client.channel === 'function') {
          channel = client.channel('tms-sync-' + user.id + '-' + Math.random().toString(36).slice(2))
            .on('postgres_changes', {event:'UPDATE',schema:'public',table:'tms_sync_state'}, remoteChange)
            .subscribe(state => {
              if (token !== lifecycle || stopped) return;
              if (state === 'SUBSCRIBED') backgroundRefresh();
            });
        }
        watchRecovery();
        return getState();
      } catch (error) {
        if (token === lifecycle) { status('offline',error); watchRecovery(); }
        throw error;
      }
    }
    function stop() {
      if (dirty) persist();
      lifecycle++; stopped = true;
      if (timer) clearInterval(timer);
      if (events) ['online','focus','visibilitychange'].forEach(name=>events.removeEventListener(name,backgroundRefresh));
      const client = options.getClient();
      if (channel && client && typeof client.removeChannel === 'function') Promise.resolve(client.removeChannel(channel)).catch(()=>{});
      timer = null; channel = null; user = null; base = null;
      revision = null; remoteRevision = null; dirty = false; conflict = false;
      saving = null; refreshing = null; refreshRequested = false; phase = 'disconnected'; lastError = null; storageError = null;
    }
    async function save() {
      if (stopped || !user || !base) throw new Error('Espera a que termine la carga compartida antes de guardar.');
      markDirty();
      if (saving) return saving;
      if (conflict) throw Object.assign(new Error('Otra sesión guardó cambios. Tus cambios locales están conservados; revisa la versión compartida antes de continuar.'),{code:'40001'});
      if (!dirty) return true;
      const token = lifecycle;
      saving = (async () => {
        while (dirty && !stopped && token === lifecycle) {
          const captured = clone(options.capture());
          const payload = difference(base,captured);
          if (!Object.keys(payload).length) { dirty=false; break; }
          status('saving');
          try {
            const result = await rpc('tms_save_snapshot',{p_expected_revision:revision,p_payload:payload});
            if (token !== lifecycle || stopped) return;
            revision = Number(result.revision); remoteRevision = Math.max(remoteRevision || 0,revision);
            base = captured;
            dirty = Object.keys(difference(base,options.capture())).length > 0;
            persist();
          } catch (error) {
            if (token !== lifecycle || stopped) throw error;
            dirty = true; conflict = error.code === '40001';
            persist(); status(conflict ? 'conflict' : 'offline',error);
            throw error;
          }
        }
        if (token === lifecycle && !stopped) status('synced');
        return true;
      })();
      try { return await saving; }
      finally {
        if (token === lifecycle) {
          saving=null;
          if (!dirty && !stopped) {
            // Refresh generated database identifiers and server-normalized data.
            if (refreshing) await refreshing.catch(()=>{});
            await refresh({force:true}).catch(()=>{});
          }
        }
      }
    }
    async function discardAndReload() {
      if (saving) throw new Error('Espera a que termine el guardado.');
      // Retain a separate recovery copy even when the user chooses the team copy.
      if ((dirty || hasDraft()) && storage) {
        try { storage.setItem(key()+':recovery',JSON.stringify(exportPending())); } catch (_) {}
      }
      const snapshot = await rpc('tms_get_snapshot',{p_known_revision:null});
      await apply(snapshot,lifecycle);
      return getState();
    }
    function exportPending() {
      return {userId:user && user.id,revision,remoteRevision,payload:clone(options.capture()),base:base && clone(base),formDrafts:options.captureDrafts ? clone(options.captureDrafts()) : [],savedAt:new Date().toISOString()};
    }
    return {start,stop,markDirty,save,refresh,discardAndReload,exportPending,getState};
  }
  return {create,difference,stable};
});

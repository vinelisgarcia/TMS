/* Serializes writes and guards a shared cloud revision across browser sessions. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TMSSync = factory();
})(typeof globalThis === 'undefined' ? this : globalThis, function () {
  'use strict';
  function create(onChange = () => {}) {
    const state = { revision: null, dirty: false, generation: 0, saving: false, loading: false, conflict: false, error: '', checkedAt: null };
    let task = null;
    const emit = () => onChange({ ...state });
    function changed() { state.dirty = true; state.generation++; emit(); }
    function loaded(revision) {
      state.revision = revision; state.dirty = false; state.conflict = false; state.error = '';
      state.checkedAt = new Date().toISOString(); emit();
    }
    async function save(write) {
      if (task) return task;
      task = (async () => {
        if (state.loading || state.revision === null || state.conflict) {
          throw new Error(state.conflict ? 'Hay cambios en otra sesión. Descarga tu respaldo antes de recargar la nube.' : 'Primero carga la versión de nube. Si falta, aplica la migración 0008.');
        }
        state.saving = true; state.error = ''; emit();
        try {
          while (state.dirty) {
            const generation = state.generation;
            const revision = await write(state.revision);
            state.revision = revision;
            state.dirty = generation !== state.generation;
            state.checkedAt = new Date().toISOString(); emit();
          }
          return true;
        } catch (error) {
          state.error = error.message;
          state.conflict = /TMS_CONFLICT/.test(error.message);
          throw error;
        } finally { state.saving = false; emit(); }
      })();
      try { return await task; } finally { task = null; }
    }
    return { state, changed, loaded, save, emit };
  }
  return { create };
});

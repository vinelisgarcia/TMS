(function () {
  'use strict';
  const definitions = {
    plan: { label: 'Programación', process: file => window.procesarArchivo(file, 'plan') },
    control: { label: 'Control diario', process: file => window.procesarArchivo(file, 'control') },
    'warehouse-plan': { label: 'Solicitudes a almacén', process: file => window.procesarSolicitudesAlmacen(file, 'plan') },
    'warehouse-control': { label: 'Control de almacén', process: file => window.procesarSolicitudesAlmacen(file, 'control') }
  };
  const supported = typeof window.showOpenFilePicker === 'function' && !!window.indexedDB;
  const handles = new Map();
  let db, account = '', loaded = false, loading = null;
  const currentAccount = () => String(APP.currentUserProfile?.authUserId || APP.currentUserProfile?.userId || APP.currentUserProfile?.email || '').toLowerCase();
  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open('tms-sap-file-locations', 1);
      request.onupgradeneeded = () => request.result.createObjectStore('files');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }
  function transaction(mode, operation) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction('files', mode);
      const request = operation(tx.objectStore('files'));
      tx.oncomplete = () => resolve(request.result);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error || new Error('No se pudo guardar la selección.'));
    });
  }
  function render() {
    const panel = document.getElementById('sapSavedFiles');
    if (!panel) return;
    panel.hidden = !supported;
    if (!supported) return;
    panel.querySelectorAll('[data-sap-config]').forEach(button => {
      const id = button.dataset.sapConfig;
      const handle = handles.get(id);
      button.querySelector('small').textContent = handle ? handle.name : 'Sin configurar';
      button.querySelector('.sap-config-action').textContent = handle ? 'Cambiar' : 'Seleccionar';
      const upload = document.querySelector(`[data-sap-upload="${id}"]`);
      button.disabled = !loaded || !upload || upload.disabled;
    });
  }
  async function load() {
    if (!supported) return;
    const next = currentAccount();
    if (!next) return;
    if (account === next && (loaded || loading)) return loading;
    account = next; loaded = false; handles.clear(); render();
    loading = (async () => {
      try {
        db ||= await openDb();
        const entries = await Promise.all(Object.keys(definitions).map(async id => [id, await transaction('readonly', store => store.get(next + ':' + id))]));
        if (account !== next) return;
        entries.forEach(([id, handle]) => { if (handle) handles.set(id, handle); });
        loaded = true;
      } catch (error) {
        console.warn('No se pudieron recuperar las ubicaciones SAP:', error);
        loaded = false;
      } finally { loading = null; render(); }
    })();
    return loading;
  }
  function status(message, error) {
    const el = document.getElementById('sapSavedMessage');
    if (el) { el.textContent = message; el.classList.toggle('is-error', !!error); }
  }
  async function choose(id) {
    // The picker must be opened directly from the user's click.
    const [handle] = await window.showOpenFilePicker({ multiple: false, types: [{ description: 'Archivos SAP', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'], 'application/xml': ['.xml'] } }] });
    const selectedAccount = account;
    await transaction('readwrite', store => store.put(handle, selectedAccount + ':' + id));
    if (currentAccount() !== selectedAccount) return null;
    handles.set(id, handle); render();
    return handle;
  }
  async function useFile(id, configureOnly) {
    const def = definitions[id];
    if (!def || !loaded || account !== currentAccount()) { status('Espera un momento mientras se recuperan los archivos guardados.', true); load(); return; }
    const upload = document.querySelector(`[data-sap-upload="${id}"]`);
    if (!upload || upload.disabled) return;
    try {
      let handle = handles.get(id);
      if (!handle || configureOnly) {
        handle = await choose(id);
        if (!handle) return;
        status(`${def.label}: ${handle.name} guardado en este navegador.`);
        if (configureOnly) return;
      }
      const warning = id === 'plan' ? '\n\nActualizar la programación sustituye la planificación actual y reinicia el control diario.' : '';
      if (!window.confirm(`¿Actualizar ${def.label.toLowerCase()} con «${handle.name}»?${warning}`)) return;
      if (account !== currentAccount() || upload.disabled) return;
      let permission = await handle.queryPermission({ mode: 'read' });
      if (permission !== 'granted') permission = await handle.requestPermission({ mode: 'read' });
      if (permission !== 'granted') { status('No se concedió acceso al archivo. Puedes volver a intentarlo.', true); return; }
      const file = await handle.getFile();
      if (account !== currentAccount() || upload.disabled) return;
      def.process(file);
      status(`Procesando ${file.name}…`);
    } catch (error) {
      if (error.name === 'AbortError') return;
      console.warn('No se pudo acceder al archivo SAP guardado:', error);
      status('No se pudo abrir el archivo guardado. Selecciónalo de nuevo en «Configurar archivos».', true);
    }
  }
  window.addEventListener('DOMContentLoaded', () => {
    if (!supported) return;
    const workspace = document.getElementById('importExcelTableCard');
    if (!workspace) return;
    const details = document.createElement('details'); details.id = 'sapSavedFiles'; details.className = 'sap-saved-files';
    details.innerHTML = '<summary>Configurar archivos guardados <span>Opcional</span></summary><p>Elige una vez cada archivo en este navegador. Después podrás actualizarlo sin volver a buscarlo. Si lo mueves o cambias de nombre, selecciónalo de nuevo.</p><div class="sap-saved-grid">' + Object.entries(definitions).map(([id, def]) => `<button type="button" data-sap-config="${id}"><span><strong>${def.label}</strong><small>Sin configurar</small></span><span class="sap-config-action">Seleccionar</span></button>`).join('') + '</div>';
    const message = document.createElement('p'); message.id = 'sapSavedMessage'; message.className = 'sap-saved-message'; message.setAttribute('role','status');
    workspace.append(details, message);
    details.addEventListener('click', event => { const button = event.target.closest('[data-sap-config]'); if (button && !button.disabled) useFile(button.dataset.sapConfig, true); });
    workspace.addEventListener('click', event => {
      const button = event.target.closest('[data-sap-upload]');
      if (!button || button.disabled || (!loaded && !loading)) return;
      event.preventDefault(); event.stopImmediatePropagation();
      useFile(button.dataset.sapUpload, false);
    }, true);
    document.querySelectorAll('[data-sap-upload]').forEach(button => new MutationObserver(render).observe(button, { attributes: true, attributeFilter: ['disabled'] }));
    window.addEventListener('tms-permissions-changed', () => { load(); render(); });
    load(); render();
  });
})();

(function () {
  'use strict';
  const modules = {
    rutas: { title: 'Gestión de rutas', label: 'Rutas', icon: '<path d="M3 18h4l3-12h4l3 12h4M10 6l-2 8m8-8 2 8M12 12v2m0 3v2"/>', views: ['dashboard','calendario','prioridades','incidenciasDespacho','solicitudesAlmacen','configClientes','importar'] },
    inventarios: { title: 'Gestión de inventarios', label: 'Inventarios', icon: '<path d="m3 7 9-4 9 4v10l-9 4-9-4zM3 7l9 4 9-4m-9 4v10M7 5l10 5"/>', views: ['inventario','showroom'] },
    importaciones: { title: 'Gestión de importaciones', label: 'Importaciones', icon: '<path d="M3 18h18M5 15h14l2-6-9-3-9 3 2 6Zm7-9V3m-4 7h.01M12 11h.01M16 10h.01"/>', views: ['materialesTransito','packingMiami','auditoriaJamaica','eur1'] },
    comercial: { title: 'Comercial', label: 'Comercial', icon: '<path d="M4 8h16v11H4zM9 8V5h6v3M4 12h16m-9 0v2h2v-2"/>', views: ['comercial','calendario'] },
    configuracion: { title: 'Configuración', label: 'Configuración', icon: '<path d="M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8ZM12 2v3m0 14v3M2 12h3m14 0h3M5 5l2 2m10 10 2 2M5 19l2-2M17 7l2-2"/>', views: ['configClientes','auditoriaSistema'] }
  };
  const permission = {configClientes:'configuracion'};
  const canSee = id => !!window.TMSAuditCore?.hasPermission(permission[id] || id, 'ver');
  const available = module => module.views.some(canSee);
  const selected = new URL(location.href).searchParams.get('modulo');
  let current = modules[selected] ? selected : null;
  let initialized = false, previousView = '', sapNavigate;
  window.TMSModuleContains = id => !!(current && modules[current].views.includes(id));
  function render() {
    const home = document.getElementById('moduleHome');
    if (!home) return;
    home.innerHTML = '<div class="module-home-heading"><span>ÁLVAREZ · ESPACIO DE TRABAJO</span><h1>¿Qué quieres gestionar hoy?</h1><p>Elige un módulo para abrirlo en una nueva pestaña.</p></div><div class="module-cards">' + Object.entries(modules).filter(([,module]) => available(module)).map(([id,module]) => `<a class="module-card" href="?modulo=${id}" target="_blank" rel="noopener" aria-label="${module.title}, se abre en una nueva pestaña"><span class="module-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${module.icon}</svg></span><span class="module-card-body"><strong>${module.label}</strong></span></a>`).join('') + '</div>';
    if (!home.querySelector('.module-card')) home.querySelector('.module-cards').innerHTML = '<p>No tienes módulos asignados. Solicita acceso a un administrador.</p>';
  }
  function update() {
    if (!window.TMSAuditCore || typeof APP === 'undefined' || !APP.currentUserProfile || document.getElementById('mainLayout').style.display === 'none') return;
    render();
    document.body.classList.toggle('module-home', !current);
    const back = document.getElementById('moduleBack');
    back.hidden = !current;
    if (current) back.querySelector('span').textContent = modules[current].title;
    document.querySelectorAll('#sidebar .nav-btn').forEach(button => {
      const id = button.id.replace(/^nav-/, '');
      button.classList.toggle('module-outside', !window.TMSModuleContains(id));
    });
    const configLabel = document.querySelector('#nav-configClientes .nav-label');
    if (configLabel) configLabel.textContent = current === 'rutas' ? 'Clientes y rutas' : 'Configuración';
    const configTitle = document.querySelector('#configClientes .view-title');
    if (configTitle) configTitle.textContent = current === 'rutas' ? 'Clientes y rutas' : 'Configuración';
    const importButton = document.getElementById('nav-importar');
    if (importButton) importButton.querySelector('.nav-label').textContent = 'Actualizar datos SAP';
    if (current && !available(modules[current])) {
      current = null; update(); return;
    }
    if (current && !initialized) {
      initialized = true;
      const first = modules[current].views.find(id => id !== 'importar' && id !== 'auditoriaJamaica' && canSee(id));
      if (first) window.cambiarVista(first, document.getElementById('nav-' + first));
      else if (canSee('auditoriaJamaica')) {
        document.querySelectorAll('.view').forEach(view => view.classList.remove('active'));
        const note = document.createElement('div'); note.className = 'module-empty'; note.textContent = 'Selecciona Auditoría Jamaica en el menú para abrirla.'; document.getElementById('content').prepend(note);
      }
      else if (canSee('importar')) openSap();
    }
  }
  function openSap() {
    if (!canSee('importar')) return;
    const dialog = document.getElementById('sapModuleDialog');
    previousView = document.querySelector('.view.active')?.id || '';
    sapNavigate('importar', document.getElementById('nav-importar'));
    if (!dialog.open) dialog.showModal();
  }
  function closeSap() {
    const dialog = document.getElementById('sapModuleDialog');
    if (dialog.open) dialog.close();
    if (previousView && previousView !== 'importar' && canSee(previousView)) window.cambiarVista(previousView, document.getElementById('nav-' + previousView));
  }
  window.addEventListener('DOMContentLoaded', () => {
    const main = document.getElementById('mainLayout');
    const home = document.createElement('main'); home.id = 'moduleHome'; main.prepend(home);
    const mobileBack = document.createElement('a'); mobileBack.id = 'mobileHomeLink'; mobileBack.href = location.pathname; mobileBack.textContent = '← Inicio'; main.insertBefore(mobileBack, document.getElementById('mobileModuleNav'));
    const back = document.createElement('a'); back.id = 'moduleBack'; back.href = location.pathname; back.innerHTML = '← Inicio <span></span>'; document.getElementById('sidebar').prepend(back);
    const dialog = document.createElement('dialog'); dialog.id = 'sapModuleDialog'; dialog.setAttribute('aria-labelledby','sapDialogTitle');
    dialog.innerHTML = '<header class="sap-dialog-header"><div class="sap-dialog-heading"><span class="sap-dialog-mark" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none"><path d="M12 16V4m-4 4 4-4 4 4M5 15v4a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-4" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg></span><div><span class="sap-dialog-eyebrow">GESTIÓN DE RUTAS</span><h2 id="sapDialogTitle">Actualizar datos SAP</h2><p>Mantén la operación al día con tus últimos archivos.</p></div></div><button class="sap-dialog-close" type="button" aria-label="Cerrar">×</button></header><div class="sap-dialog-content"></div>';
    dialog.querySelector('.sap-dialog-content').appendChild(document.getElementById('importar')); document.body.appendChild(dialog);
    dialog.querySelector('header button').addEventListener('click', closeSap);
    dialog.addEventListener('cancel', event => { event.preventDefault(); closeSap(); });
    dialog.addEventListener('click', event => { if (event.target === dialog) closeSap(); });
    document.getElementById('nav-importar').onclick = openSap;
    const original = window.cambiarVista; sapNavigate = original;
    window.cambiarVista = function (id, button) {
      if (id === 'importar' && !dialog.open && current === 'rutas') { openSap(); return; }
      if (dialog.open && id !== 'importar') dialog.close();
      return original(id, button);
    };
    window.addEventListener('tms-permissions-changed', update);
    update();
  });
})();

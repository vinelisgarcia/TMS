(function () {
  'use strict';
  const modules = {
    rutas: { title: 'Gestión de rutas', icon: '↗', description: 'Planifica entregas y sigue la operación de despacho.', views: ['dashboard','calendario','prioridades','incidenciasDespacho','solicitudesAlmacen','configClientes','importar'] },
    inventarios: { title: 'Gestión de inventarios', icon: '▦', description: 'Consulta y concilia los inventarios de Caribetrans y showroom.', views: ['inventario','showroom'] },
    importaciones: { title: 'Gestión de importaciones', icon: '⇥', description: 'Sigue los materiales y la documentación de importación.', views: ['materialesTransito','packingMiami','auditoriaJamaica','eur1'] },
    comercial: { title: 'Comercial', icon: '◈', description: 'Consulta la actividad comercial y el calendario de entregas.', views: ['comercial','calendario'] },
    configuracion: { title: 'Configuración', icon: '⚙', description: 'Administra clientes, parámetros y accesos del sistema.', views: ['configClientes','auditoriaSistema'] }
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
    home.innerHTML = '<div class="module-home-heading"><span>ÁLVAREZ · ESPACIO DE TRABAJO</span><h1>¿Qué quieres gestionar hoy?</h1><p>Elige un módulo para abrirlo en una nueva pestaña.</p></div><div class="module-cards">' + Object.entries(modules).filter(([,module]) => available(module)).map(([id,module]) => `<a class="module-card" href="?modulo=${id}" target="_blank" rel="noopener"><span class="module-icon" aria-hidden="true">${module.icon}</span><span class="module-card-body"><strong>${module.title}</strong><small>${module.description}</small></span><span class="module-arrow" aria-hidden="true">↗</span><span class="module-card-footer">Abrir módulo <span>↗</span></span></a>`).join('') + '</div>';
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

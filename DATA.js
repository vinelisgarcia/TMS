/* TMS mobile/status UI patch
 * - Adds subtle color indicators to operational statuses.
 * - Replaces the mobile bottom module strip with a compact module selector.
 * - Tightens iPhone viewport/layout behavior for home-screen shortcut usage.
 */
(function () {
  'use strict';

  const STYLE_ID = 'tms-mobile-status-ui-patch';
  const MENU_ID = 'mobileModuleMenu';
  const SELECT_ID = 'mobileModuleSelect';

  const MODULES = [
    { id: 'importar', icon: '📥', label: 'Importar SAP' },
    { id: 'dashboard', icon: '📊', label: 'Dashboard' },
    { id: 'calendario', icon: '📅', label: 'Calendario' },
    { id: 'comercial', icon: '👔', label: 'Vista Comercial' },
    { id: 'rutas', icon: '🧭', label: 'Rutas' },
    { id: 'importaciones', icon: '↗', label: 'Importaciones' },
    { id: 'prioridades', icon: '!', label: 'Prioridades' },
    { id: 'configClientes', icon: '⚙️', label: 'Config. Clientes' },
    { id: 'solicitudesAlmacen', icon: '📦', label: 'Solicitudes Almacén' },
    { id: 'almacen', icon: '🏭', label: 'Vista Almacén' }
  ];

  function ensureViewport() {
    let viewport = document.querySelector('meta[name="viewport"]');
    if (!viewport) {
      viewport = document.createElement('meta');
      viewport.setAttribute('name', 'viewport');
      document.head.appendChild(viewport);
    }
    viewport.setAttribute('content', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }

  function injectMobileStatusStyles() {
    if (document.getElementById(STYLE_ID)) return;

    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      html,
      body {
        width: 100%;
        min-width: 0;
        overflow-x: hidden;
        -webkit-text-size-adjust: 100%;
      }

      .badge.status-pending-color {
        background: #FDE2E2 !important;
        color: #922B21 !important;
        border: 1px solid #F3B3B3 !important;
      }

      .badge.status-process-color {
        background: #FEF3C7 !important;
        color: #92400E !important;
        border: 1px solid #FCD34D !important;
      }

      .badge.status-completed-color {
        background: #D5EDD9 !important;
        color: #1D6A3A !important;
        border: 1px solid #A8D5B5 !important;
      }

      .status-dot {
        display: inline-block;
        width: 8px;
        height: 8px;
        border-radius: 999px;
        margin-right: 6px;
        vertical-align: middle;
        box-shadow: 0 0 0 2px rgba(255,255,255,.72);
      }

      .status-dot-pending { background: #922B21; }
      .status-dot-process { background: #F5C542; }
      .status-dot-completed { background: #00A676; }

      .cal-card.status-pending,
      .ruta-item.status-pending {
        border-left-color: #922B21 !important;
      }

      .cal-card.status-process,
      .ruta-item.status-process {
        border-left-color: #F5C542 !important;
      }

      .cal-card.status-completed,
      .ruta-item.status-completed {
        border-left-color: #00A676 !important;
      }

      #${MENU_ID} { display: none; }

      @media (max-width: 768px) {
        html,
        body {
          max-width: 100%;
          min-width: 0;
          overflow-x: hidden;
        }

        body {
          min-height: 100svh;
        }

        #appHeader {
          width: 100%;
          max-width: 100%;
          padding: max(8px, env(safe-area-inset-top)) 10px 8px !important;
          gap: 7px !important;
        }

        #appHeader .logo img {
          width: min(142px, 44vw) !important;
          max-width: 44vw !important;
          height: 30px !important;
        }

        #saveStateBtn,
        #undoStateBtn,
        #logoutBtn {
          min-width: 0;
          white-space: nowrap;
        }

        #mainLayout {
          display: block !important;
          width: 100%;
          max-width: 100%;
          padding-bottom: 0 !important;
          overflow-x: hidden !important;
        }

        #content {
          width: 100%;
          max-width: 100%;
          overflow-x: hidden !important;
        }

        .view {
          width: 100%;
          max-width: 100%;
          min-height: auto;
          padding: 12px 10px 18px !important;
        }

        .card,
        .chart-card,
        #reporteOutput {
          max-width: 100%;
          overflow: hidden;
        }

        #sidebar {
          display: none !important;
        }

        #${MENU_ID} {
          display: block;
          position: sticky;
          top: 0;
          z-index: 850;
          background: #0F172A;
          padding: 8px 10px;
          border-bottom: 1px solid #1E293B;
          box-shadow: 0 6px 18px rgba(15,23,42,.12);
        }

        .mobile-module-label {
          display: block;
          color: #D7DEE8;
          font-size: 11px;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: .04em;
          margin-bottom: 5px;
        }

        #${SELECT_ID} {
          width: 100%;
          min-height: 42px;
          border: none;
          border-radius: 10px;
          padding: 0 12px;
          background: #FFFFFF;
          color: #111827;
          font-size: 15px;
          font-weight: 700;
          font-family: inherit;
        }

        .stats,
        .kpi-row,
        .report-kpi-grid,
        #sapEstadoBanner > div {
          grid-template-columns: 1fr !important;
        }

        .dash-charts-grid {
          grid-template-columns: 1fr !important;
        }

        .chart-card[style*="grid-column"] {
          grid-column: auto !important;
        }

        .table-wrap {
          max-width: 100%;
          overflow-x: auto;
          -webkit-overflow-scrolling: touch;
        }

        .data-table {
          min-width: 560px;
        }

        #calContainer {
          overflow-x: hidden !important;
        }

        .cal-semana {
          min-width: 0 !important;
          width: 100%;
          display: grid !important;
          grid-template-columns: 1fr;
          gap: 10px;
        }

        .cal-dia-col {
          min-width: 0 !important;
          width: 100%;
        }

        .cal-card-actions {
          flex-wrap: wrap;
        }

        .rutas-grid,
        .config-grid,
        .config-layout {
          grid-template-columns: 1fr !important;
        }

        .modal-box {
          min-width: 0 !important;
        }
      }
    `;
    document.head.appendChild(style);
  }

  function getStatusTypeFromText(value) {
    const normalized = String(value || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase();

    if (normalized.includes('cumpl') || normalized.includes('complet') || normalized.includes('entreg')) return 'completed';
    if (normalized.includes('proceso') || normalized.includes('parcial') || normalized.includes('incidencia')) return 'process';
    if (normalized.includes('pend')) return 'pending';
    if (normalized.includes('✅')) return 'completed';
    if (normalized.includes('⚠') || normalized.includes('🟡')) return 'process';
    if (normalized.includes('🔄') || normalized.includes('⏳') || normalized.includes('🔴')) return 'pending';
    return '';
  }

  function addStatusDot(element, type) {
    if (!element || element.querySelector('.status-dot')) return;
    const dot = document.createElement('span');
    dot.className = 'status-dot status-dot-' + type;
    element.insertBefore(dot, element.firstChild);
  }

  function decorateStatusBadges(root) {
    (root || document).querySelectorAll('.badge').forEach(badge => {
      const type = getStatusTypeFromText(badge.textContent);
      if (!type) return;

      badge.classList.remove('status-pending-color', 'status-process-color', 'status-completed-color');
      badge.classList.add('status-' + type + '-color');
      addStatusDot(badge, type);
    });
  }

  function decorateStatusCards(root) {
    (root || document).querySelectorAll('.cal-card, .ruta-item').forEach(card => {
      const type = getStatusTypeFromText(card.textContent) || 'pending';
      card.classList.remove('status-pending', 'status-process', 'status-completed');
      card.classList.add('status-' + type);
    });
  }

  function decorateStatuses(root) {
    decorateStatusBadges(root);
    decorateStatusCards(root);
  }

  function activeViewId() {
    const active = document.querySelector('.view.active');
    return active ? active.id : 'importar';
  }

  function syncMobileModuleSelect(id) {
    const select = document.getElementById(SELECT_ID);
    if (select && id && select.value !== id) select.value = id;
  }

  function ensureMobileModuleMenu() {
    const layout = document.getElementById('mainLayout');
    const content = document.getElementById('content');
    if (!layout || !content || document.getElementById(MENU_ID)) return;

    const menu = document.createElement('div');
    menu.id = MENU_ID;
    menu.innerHTML = `
      <label class="mobile-module-label" for="${SELECT_ID}">Módulo</label>
      <select id="${SELECT_ID}" aria-label="Seleccionar módulo del TMS">
        ${MODULES.map(module => `<option value="${module.id}">${module.icon} ${module.label}</option>`).join('')}
      </select>
    `;

    layout.insertBefore(menu, content);
    const select = document.getElementById(SELECT_ID);
    if (select) {
      select.value = activeViewId();
      select.addEventListener('change', () => {
        const targetId = select.value;
        const navButton = document.getElementById('nav-' + targetId);
        if (typeof window.cambiarVista === 'function') {
          window.cambiarVista(targetId, navButton);
        }
      });
    }
  }

  function patchNavigationSync() {
    if (window.__tmsMobileNavigationPatched || typeof window.cambiarVista !== 'function') return;
    const originalCambiarVista = window.cambiarVista;
    window.cambiarVista = function patchedCambiarVista(id, btn) {
      const result = originalCambiarVista.apply(this, arguments);
      syncMobileModuleSelect(id);
      window.requestAnimationFrame(() => decorateStatuses(document));
      return result;
    };
    window.__tmsMobileNavigationPatched = true;
  }

  function observeDynamicRenders() {
    const target = document.getElementById('content') || document.body;
    if (!target || window.__tmsStatusObserver) return;

    let pending = false;
    const observer = new MutationObserver(() => {
      if (pending) return;
      pending = true;
      window.requestAnimationFrame(() => {
        decorateStatuses(document);
        pending = false;
      });
    });

    observer.observe(target, { childList: true, subtree: true });
    window.__tmsStatusObserver = observer;
  }

  function boot() {
    ensureViewport();
    injectMobileStatusStyles();
    ensureMobileModuleMenu();
    patchNavigationSync();
    decorateStatuses(document);
    observeDynamicRenders();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();

(function () {
  'use strict';

  if (typeof APP === 'undefined') return;

  const LOCAL_KEY = 'tms_alvarez_v2_state';
  const THEME_KEY = 'tms_alvarez_visual_theme';
  const CURRENT_YEAR = new Date().getFullYear();
  const RD_HOLIDAYS = {
    2026: [
      '01/01/2026','05/01/2026','21/01/2026','26/01/2026','27/02/2026',
      '03/04/2026','04/05/2026','04/06/2026','16/08/2026',
      '24/09/2026','09/11/2026','25/12/2026'
    ]
  };

  Object.assign(APP, {
    lineItems: [],
    planLineItems: [],
    controlLineItems: [],
    deliveryGroups: [],
    routeConfigs: APP.routeConfigs || [],
    routeCostHistory: APP.routeCostHistory || [],
    controlHistory: APP.controlHistory || [],
    solicitudesHistory: APP.solicitudesHistory || [],
    solicitudesPlanAlmacen: APP.solicitudesPlanAlmacen || [],
    solicitudesControlAlmacen: APP.solicitudesControlAlmacen || [],
    solicitudesCompare: APP.solicitudesCompare || [],
    warehouseSettings: APP.warehouseSettings || { costoSolicitud: 0 },
    userProfiles: APP.userProfiles || [],
    rolePermissions: APP.rolePermissions || buildDefaultRolePermissions(),
    currentUserProfile: APP.currentUserProfile || null,
    camionExtraEnabled: !!APP.camionExtraEnabled,
    feriadosRD: APP.feriadosRD || (RD_HOLIDAYS[CURRENT_YEAR] || []),
    importFiles: APP.importFiles || { plan: '', control: '', solicitudesPlan: '', solicitudesControl: '' },
    importShipments: APP.importShipments || [],
    dailyPriorities: APP.dailyPriorities || [],
    prioritySpainTopics: APP.prioritySpainTopics || [],
    calendarNotes: APP.calendarNotes || [],
    selectedCalendarGroups: APP.selectedCalendarGroups || [],
    visualTheme: APP.visualTheme || localStorage.getItem(THEME_KEY) || 'light',
    undoStack: APP.undoStack || [],
    pendingTruckPhoto: APP.pendingTruckPhoto || null,
    appVersion: 'v2'
  });

  function nowIso() {
    return new Date().toISOString();
  }

  function stripAccents(value) {
    return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  function normKey(value) {
    return stripAccents(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
  }

  function text(value) {
    return String(value == null ? '' : value).trim();
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function isValidLoginEmail(value) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text(value));
  }

  function getProfileAccountMeta(profile) {
    const perms = profile && profile.permisosPorModulo;
    const meta = perms && perms.__account ? perms.__account : {};
    return {
      passwordChangeRequired: !!(profile && (profile.passwordChangeRequired || meta.passwordChangeRequired)),
      authNeedsConfirmation: !!(profile && (profile.authNeedsConfirmation || meta.authNeedsConfirmation)),
      passwordConfigured: profile && profile.passwordConfigured !== undefined ? !!profile.passwordConfigured : meta.passwordConfigured !== false,
      accountStatus: text((profile && profile.accountStatus) || meta.accountStatus || '')
    };
  }

  function getBrandLogoSrc() {
    return './assets/hospitality-by-alvarez.svg';
  }

  function permissionsPayloadForProfile(profile) {
    const payload = getPermissionsForRole((profile && profile.rol) || 'Solo lectura');
    const meta = getProfileAccountMeta(profile || {});
    payload.__account = {
      passwordChangeRequired: !!meta.passwordChangeRequired,
      authNeedsConfirmation: !!meta.authNeedsConfirmation,
      passwordConfigured: !!meta.passwordConfigured,
      accountStatus: meta.accountStatus || (meta.authNeedsConfirmation ? 'Pendiente de confirmar correo' : 'Activo')
    };
    return payload;
  }

  function jsString(value) {
    return String(value == null ? '' : value)
      .replace(/\\/g, '\\\\')
      .replace(/'/g, "\\'")
      .replace(/\n/g, ' ');
  }

  function normalizeVisualTheme(theme) {
    return ['light', 'dark', 'brand'].includes(theme) ? theme : 'light';
  }

  function getThemeLabel(theme) {
    return ({ light: 'Claro', dark: 'Oscuro', brand: 'Alvarez' }[normalizeVisualTheme(theme)] || 'Claro');
  }

  const MODULE_LABELS = {
    importar: 'Importar',
    dashboard: 'Dashboard',
    calendario: 'Calendario',
    reportes: 'Reportes',
    comercial: 'Comercial',
    rutas: 'Rutas',
    importaciones: 'Importaciones',
    prioridades: 'Prioridades',
    configuracion: 'Configuración',
    solicitudesAlmacen: 'Solicitudes almacén',
    almacen: 'Almacén'
  };

  function syncThemeButtons() {
    document.querySelectorAll('[data-theme-option]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.themeOption === APP.visualTheme);
    });
    const label = document.getElementById('themeCurrentLabel');
    if (label) label.textContent = getThemeLabel(APP.visualTheme);
  }

  function ensureAppFooters() {
    const text = 'TMS para uso exclusivo de Hospitality by Alvarez. Desarrollado por MVOG.';
    document.querySelectorAll('.view').forEach(view => {
      let footer = view.querySelector(':scope > .app-view-footer');
      if (!footer) {
        footer = document.createElement('div');
        footer.className = 'app-view-footer';
        view.appendChild(footer);
      }
      footer.textContent = text;
    });
  }

  function applyBrandLogo() {
    const loginLogo = document.querySelector('.login-logo .login-icon');
    if (loginLogo && !loginLogo.querySelector('img')) {
      loginLogo.classList.add('brand-login-logo');
      loginLogo.innerHTML = `<img src="${getBrandLogoSrc()}" alt="Hospitality by Alvarez">`;
    }
    const headerLogo = document.querySelector('#appHeader .logo');
    if (headerLogo && !headerLogo.querySelector('img')) {
      headerLogo.classList.add('brand-header-logo');
      headerLogo.innerHTML = `<img src="${getBrandLogoSrc()}" alt="Hospitality by Alvarez">`;
    }
  }

  function applyVisualTheme(theme, persist) {
    const nextTheme = normalizeVisualTheme(theme);
    APP.visualTheme = nextTheme;
    if (persist) localStorage.setItem(THEME_KEY, nextTheme);
    if (!document.body) return;
    document.body.classList.remove('theme-light', 'theme-dark', 'theme-brand');
    document.body.classList.add('theme-' + nextTheme);
    document.body.dataset.theme = nextTheme;
    syncThemeButtons();
  }

  window.setVisualTheme = function setVisualTheme(theme) {
    applyVisualTheme(theme, true);
    if (typeof window.scheduleAutoSave === 'function') window.scheduleAutoSave();
  };

  function num(value) {
    if (value == null || value === '') return 0;
    if (typeof value === 'number') return isNaN(value) ? 0 : value;
    const cleaned = String(value).replace(/\s/g, '').replace(/,/g, '');
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? 0 : parsed;
  }

  function normalizeRow(row) {
    const out = {};
    Object.keys(row || {}).forEach(key => {
      out[normKey(key)] = row[key];
    });
    return out;
  }

  function pickField(row, aliases) {
    const normalized = normalizeRow(row);
    for (const alias of aliases) {
      const key = normKey(alias);
      if (Object.prototype.hasOwnProperty.call(normalized, key)) {
        return normalized[key];
      }
    }
    return '';
  }

  function toDateLabel(value) {
    return value ? normalizarFecha(value) : '';
  }

  function dedupeBy(items, keyFn) {
    const seen = new Set();
    const out = [];
    items.forEach(item => {
      const key = keyFn(item);
      if (!key || seen.has(key)) return;
      seen.add(key);
      out.push(item);
    });
    return out;
  }

  function getTruckOptions() {
    return APP.camionExtraEnabled
      ? ['CAMION 1', 'CAMION 2', 'CAMION 3', 'ALMACEN']
      : ['CAMION 1', 'CAMION 2', 'ALMACEN'];
  }

  function getTruckLabel(camion) {
    if (camion === 'CAMION 1') return 'Camión 1';
    if (camion === 'CAMION 2') return 'Camión 2';
    if (camion === 'CAMION 3') return 'Camión 3';
    if (camion === 'ALMACEN') return 'Almacén';
    return camion || 'Sin camión';
  }

  function formatLoadQty(value) {
    const qty = num(value);
    return new Intl.NumberFormat('es-DO', {
      maximumFractionDigits: qty % 1 === 0 ? 0 : 2
    }).format(qty);
  }

  function getTruckLoadSummary(groups) {
    const source = groups || [];
    const pedidos = source.reduce((sum, group) => sum + num(group.cantidadSolicitada), 0);
    const solicitudes = source.reduce((sum, group) => sum + num(group.cargaSolicitudes), 0);
    const montoSolicitudes = source.reduce((sum, group) => sum + num(group.montoSolicitudes), 0);
    const valorTransportado = source.reduce((sum, group) => sum + num(group.totalPendiente), 0);
    const rutas = uniqueTexts(source.map(group => group.rutaNombre || group.zona));
    return {
      pedidos,
      solicitudes,
      montoSolicitudes,
      valorTransportado,
      rutas,
      total: pedidos + solicitudes
    };
  }

  function getTruckBadgeClass(camion) {
    if (camion === 'CAMION 1') return 'badge-c1';
    if (camion === 'CAMION 2') return 'badge-c2';
    if (camion === 'CAMION 3') return 'badge-warn';
    return 'badge-almacen';
  }

  function buildBaseItemKey(item) {
    return [
      text(item.clienteId || item.codigo),
      text(item.pedidoCliente),
      text(item.lineaPedidoCliente),
      text(item.articulo),
      text(item.descripcionArticulo)
    ].join('|');
  }

  function buildDeliveryGroupKey(item) {
    return [
      item.fechaPlanificada || '',
      item.camionAsignado || '',
      item.rutaNombre || item.zona || '',
      text(item.clienteId || item.codigo),
      text(item.clienteNombre || item.nombre)
    ].join('|');
  }

  function buildSapMatchKey(item) {
    return [
      text(item.pedidoCliente),
      text(item.lineaPedidoCliente),
      text(item.articulo),
      text(item.clienteId || item.codigo)
    ].map(normKey).join('|');
  }

  function uniqueTexts(values) {
    return [...new Set((values || []).map(text).filter(Boolean))];
  }

  function getGroupSelectionKey(group) {
    if (!group || !Array.isArray(group.items)) return '';
    return group.items
      .map(item => item.baseKey)
      .filter(Boolean)
      .sort()
      .join('||');
  }

  function isGroupSelected(group) {
    const key = getGroupSelectionKey(group);
    return !!key && APP.selectedCalendarGroups.includes(key);
  }

  function clearCalendarSelection() {
    APP.selectedCalendarGroups = [];
    const summary = document.getElementById('calSelectionSummary');
    if (summary) summary.textContent = '0 seleccionados';
  }

  function toggleCalendarSelection(group, forceSelected) {
    const key = getGroupSelectionKey(group);
    if (!key) return;
    const current = new Set(APP.selectedCalendarGroups || []);
    const shouldSelect = typeof forceSelected === 'boolean' ? forceSelected : !current.has(key);
    if (shouldSelect) current.add(key);
    else current.delete(key);
    APP.selectedCalendarGroups = [...current];
  }

  function getSelectedGroups() {
    const keys = new Set(APP.selectedCalendarGroups || []);
    if (!keys.size) return [];
    const matches = [];
    Object.keys(APP.rutas || {}).forEach(fecha => {
      Object.keys(APP.rutas[fecha] || {}).forEach(camion => {
        (APP.rutas[fecha][camion] || []).forEach(group => {
          if (keys.has(getGroupSelectionKey(group))) matches.push({ fecha, camion, group });
        });
      });
    });
    return matches;
  }

  function getPlanningWindowDates(referenceDate) {
    const plannedDates = APP.lineItems
      .map(item => item.fechaPlanificada)
      .filter(Boolean)
      .sort((a, b) => fechaToDate(a) - fechaToDate(b));
    const fallback = referenceDate || APP.planFecha || fechaToStr(new Date());
    const minDate = plannedDates[0] || fallback;
    const maxDate = plannedDates[plannedDates.length - 1] || fallback;
    const start = fechaToDate(minDate);
    const end = fechaToDate(maxDate);
    start.setDate(start.getDate() - 21);
    end.setDate(end.getDate() + 42);
    const out = [];
    const cursor = new Date(start.getTime());
    while (cursor <= end) {
      const current = fechaToStr(cursor);
      if (isManualRouteDate(current)) out.push(current);
      cursor.setDate(cursor.getDate() + 1);
    }
    if (!out.includes(fallback) && isManualRouteDate(fallback)) out.push(fallback);
    return [...new Set(out)].sort((a, b) => fechaToDate(a) - fechaToDate(b));
  }

  function buildMoveDateOptions(referenceItem, currentDate) {
    const routeName = referenceItem && (referenceItem.rutaNombre || referenceItem.zona);
    const dates = getPlanningWindowDates(currentDate);
    const allowedDates = dates.filter(date => {
      if (!referenceItem) return true;
      const routeOk = !routeName || routeAllowsDate(routeName, date) || fechaToDate(date).getDay() === 6;
      return routeOk && (clienteAllowsDate(referenceItem, date) || fechaToDate(date).getDay() === 6);
    });
    return allowedDates.length ? allowedDates : dates;
  }

  function setMoveDateInput(dates, selectedDate) {
    const input = document.getElementById('moveFecha');
    if (!input) return;
    const sortedDates = [...new Set(dates || [])].sort((a, b) => fechaToDate(a) - fechaToDate(b));
    const fallback = selectedDate || sortedDates[0] || fechaToStr(new Date());
    input.type = 'date';
    input.className = 'form-control';
    input.min = sortedDates[0] ? fechaStrToISO(sortedDates[0]) : '';
    input.max = sortedDates[sortedDates.length - 1] ? fechaStrToISO(sortedDates[sortedDates.length - 1]) : '';
    input.value = fechaStrToISO(fallback) || '';
  }

  function getMoveDateValue() {
    const rawValue = text((document.getElementById('moveFecha') || {}).value);
    return rawValue ? isoToFechaStr(rawValue) : '';
  }

  function buildClientConfigDefault(codigo, nombre) {
    return {
      codigoCliente: codigo,
      nombreCliente: nombre || '',
      rutaAsignada: '',
      diasRecepcion: [],
      horarioAlmacen: '',
      contacto: '',
      observaciones: '',
      condicionesEspeciales: '',
      camionPermitido: 'CUALQUIERA',
      noMiercoles: false,
      noUltimaSemana: false,
      configuracionCompleta: false
    };
  }

  function getClienteConfigV2(codigo, nombre) {
    const cfg = APP.clienteConfig[codigo];
    return cfg ? { ...buildClientConfigDefault(codigo, nombre), ...cfg } : buildClientConfigDefault(codigo, nombre);
  }

  function getRouteConfigByName(routeName) {
    const name = text(routeName);
    if (!name) return null;
    const normalized = normKey(name);
    return APP.routeConfigs.find(r => {
      const names = [r.nombre, r.zona, r.id].map(text).filter(Boolean);
      return names.some(value => value === name || normKey(value) === normalized);
    }) || null;
  }

  function routeExistsByName(routeName) {
    return !!getRouteConfigByName(routeName);
  }

  function ensureRouteConfig(routeName) {
    const name = text(routeName);
    if (!name) return null;
    let route = getRouteConfigByName(name);
    if (!route) {
      route = {
        id: 'route-' + normKey(name),
        nombre: name,
        zona: name,
        diasOperacion: [],
        costeRuta: 0,
        activa: true,
        observaciones: ''
      };
      APP.routeConfigs.push(route);
    }
    return route;
  }

  function inferRouteName(clienteNombre, clienteCodigo) {
    const cfg = getClienteConfigV2(clienteCodigo, clienteNombre);
    return cfg.rutaAsignada || inferirZona(clienteNombre, clienteCodigo);
  }

  function buildAdminPermissions() {
    return {
      importar: { ver: true, editar: true, importar: true },
      dashboard: { ver: true, editar: true },
      calendario: { ver: true, editar: true },
      reportes: { ver: true, editar: true },
      comercial: { ver: true, editar: true },
      rutas: { ver: true, editar: true },
      importaciones: { ver: true, editar: true, importar: true },
      prioridades: { ver: true, editar: true, importar: true },
      configuracion: { ver: true, editar: true, importar: true },
      solicitudesAlmacen: { ver: true, editar: true },
      almacen: { ver: true, editar: true }
    };
  }

  function buildDefaultRolePermissions() {
    return {
      Admin: buildPermissions('full', 'Admin'),
      Operaciones: buildPermissions('ops', 'Operaciones'),
      Comercial: buildPermissions('commercial', 'Comercial'),
      Almacén: buildPermissions('warehouse', 'Almacén'),
      'Solo lectura': buildPermissions('read_only', 'Solo lectura')
    };
  }

  function ensureRolePermissions() {
    const defaults = buildDefaultRolePermissions();
    APP.rolePermissions = APP.rolePermissions || {};
    Object.keys(defaults).forEach(role => {
      APP.rolePermissions[role] = APP.rolePermissions[role] || defaults[role];
    });
    Object.keys(APP.rolePermissions).forEach(role => {
      const normalized = normalizePermissions(APP.rolePermissions[role]);
      const defaultPerms = defaults[role] || {};
      Object.keys(MODULE_LABELS).forEach(module => {
        const current = APP.rolePermissions[role] && APP.rolePermissions[role][module];
        if (!current && defaultPerms[module]) normalized[module] = safeClone(defaultPerms[module]);
      });
      APP.rolePermissions[role] = normalized;
    });
    APP.userProfiles.forEach(profile => {
      const account = getProfileAccountMeta(profile);
      profile.permisosPorModulo = safeClone(APP.rolePermissions[profile.rol] || APP.rolePermissions['Solo lectura'] || defaults['Solo lectura']);
      profile.passwordChangeRequired = !!account.passwordChangeRequired;
      profile.authNeedsConfirmation = !!account.authNeedsConfirmation;
      profile.passwordConfigured = !!account.passwordConfigured;
      profile.accountStatus = account.accountStatus || '';
    });
  }

  function normalizePermissions(perms) {
    const base = {};
    Object.keys(MODULE_LABELS).forEach(module => {
      const current = (perms && perms[module]) || {};
      base[module] = {
        ver: !!current.ver,
        editar: !!current.editar,
        importar: !!current.importar
      };
      if (base[module].editar || base[module].importar) base[module].ver = true;
    });
    return base;
  }

  function getPermissionsForRole(role) {
    APP.rolePermissions = APP.rolePermissions || buildDefaultRolePermissions();
    return safeClone(APP.rolePermissions[role] || APP.rolePermissions['Solo lectura'] || buildPermissions('read_only', role));
  }

  function ensureAdminProfile() {
    const { userEmail } = getSessionInfo();
    const fallbackEmail = userEmail || 'vinelis.garcia@tms-alvarez.local';
    let adminProfile = APP.userProfiles.find(profile => profile.rol === 'Admin');
    if (!adminProfile) {
      adminProfile = {
        userId: 'admin-seed',
        nombre: 'Vinelis Garcia',
        email: fallbackEmail,
        rol: 'Admin',
        permisosPorModulo: getPermissionsForRole('Admin')
      };
      APP.userProfiles.unshift(adminProfile);
    } else {
      if (!text(adminProfile.nombre) || adminProfile.nombre === 'Manuel Oñate') {
        adminProfile.nombre = 'Vinelis Garcia';
      }
      if (!text(adminProfile.email) || adminProfile.email === 'manuel.onate@tms-alvarez.local') {
        adminProfile.email = fallbackEmail;
      }
      adminProfile.permisosPorModulo = getPermissionsForRole('Admin');
    }
    ensureRolePermissions();
    APP.currentUserProfile =
      APP.userProfiles.find(profile => text(profile.email).toLowerCase() === fallbackEmail.toLowerCase()) ||
      (userEmail ? null : adminProfile) ||
      null;
  }

  function getSessionInfo() {
    const email = (document.getElementById('userEmail') || {}).textContent || '';
    return { userEmail: email.trim() };
  }

  function getSupabaseClient() {
    if (typeof SB !== 'undefined') return SB;
    return window.SB || null;
  }

  function isUuid(value) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(text(value));
  }

  function safeClone(value) {
    return JSON.parse(JSON.stringify(value == null ? null : value));
  }

  const UNDO_STATE_KEYS = [
    'lineItems',
    'planLineItems',
    'controlLineItems',
    'solicitudesAlmacen',
    'solicitudesPlanAlmacen',
    'solicitudesControlAlmacen',
    'solicitudesCompare',
    'routeConfigs',
    'routeCostHistory',
    'controlHistory',
    'solicitudesHistory',
    'warehouseSettings',
    'userProfiles',
    'rolePermissions',
    'clienteConfig',
    'planLoaded',
    'planFecha',
    'controlLoaded',
    'controlFecha',
    'solicitudesLoaded',
    'solicitudesFileName',
    'camionExtraEnabled',
    'feriadosRD',
    'importFiles',
    'importShipments',
    'dailyPriorities',
    'prioritySpainTopics',
    'calendarNotes',
    'visualTheme'
  ];

  function updateUndoButton() {
    const btn = document.getElementById('undoStateBtn');
    if (!btn) return;
    const latest = APP.undoStack && APP.undoStack[APP.undoStack.length - 1];
    btn.disabled = !(APP.undoStack || []).length;
    btn.title = latest ? 'Revertir: ' + latest.label : 'No hay cambios para revertir';
  }

  function createUndoSnapshot(label) {
    const state = {};
    UNDO_STATE_KEYS.forEach(key => {
      state[key] = safeClone(APP[key]);
    });
    return {
      label: label || 'último cambio',
      createdAt: nowIso(),
      state
    };
  }

  function pushUndoState(label) {
    if (APP.isRestoringUndo) return;
    APP.undoStack = APP.undoStack || [];
    APP.undoStack.push(createUndoSnapshot(label));
    if (APP.undoStack.length > 15) APP.undoStack.shift();
    updateUndoButton();
  }

  function restoreUndoSnapshot(snapshot) {
    if (!snapshot || !snapshot.state) return false;
    APP.isRestoringUndo = true;
    try {
      Object.assign(APP, safeClone(snapshot.state));
      ensureAdminProfile();
      applyVisualTheme(APP.visualTheme || 'light', true);
      matchWarehouseToOrders();
      rebuildDerivedState();
      updateSolicitudesImportStatus();
      if (typeof window.actualizarEstadoBanner === 'function') window.actualizarEstadoBanner();
      if (typeof window.actualizarDashboard === 'function') window.actualizarDashboard();
      renderCalendario();
      renderRutas();
      renderRouteCatalog();
      renderComercial();
      renderConfigClientes();
      renderConfigAuxSections();
      renderSolicitudesAlmacen();
      renderAlmacen();
      renderPrioridades();
      initRouteFilterOptions();
      saveLocalSnapshot();
      return true;
    } finally {
      APP.isRestoringUndo = false;
      updateUndoButton();
    }
  }

  window.revertirUltimoCambio = function revertirUltimoCambio() {
    const snapshot = (APP.undoStack || []).pop();
    if (!snapshot) {
      updateUndoButton();
      alert('No hay cambios recientes para revertir.');
      return;
    }
    if (!restoreUndoSnapshot(snapshot)) return;
    window.guardarEnSupabase().catch(e => console.warn('Error guardando reversión:', e));
    alert('Se revirtió: ' + snapshot.label);
  };

  function chunkRows(rows, size) {
    const out = [];
    for (let i = 0; i < rows.length; i += size) out.push(rows.slice(i, i + size));
    return out;
  }

  async function getAuthenticatedUser() {
    const client = getSupabaseClient();
    if (!client || !client.auth) return null;
    const { data, error } = await client.auth.getUser();
    if (error) throw error;
    return data && data.user ? data.user : null;
  }

  async function runSupabaseQuery(queryPromise, label) {
    const { data, error } = await queryPromise;
    if (error) throw new Error(label + ': ' + error.message);
    return data || [];
  }

  function describeSupabaseSaveError(error) {
    const message = text(error && error.message ? error.message : error);
    if (!message) return 'No se pudo guardar en la nube. Los cambios quedaron solo en este dispositivo.';
    if (/permission|policy|rls|denied|permis/i.test(message)) {
      return 'No se pudo guardar en la nube por permisos del rol. Los cambios quedaron solo en este dispositivo. Detalle: ' + message;
    }
    if (/network|fetch|failed|timeout|resolve|connection|conex/i.test(message)) {
      return 'No se pudo conectar con Supabase. Los cambios quedaron solo en este dispositivo. Detalle: ' + message;
    }
    return 'No se pudo guardar en la nube. Los cambios quedaron solo en este dispositivo. Detalle: ' + message;
  }

  function notifyCloudSaveError(error) {
    const detail = describeSupabaseSaveError(error);
    if (typeof actualizarEstadoBanner === 'function') {
      const banner = document.getElementById('sapEstadoBanner');
      if (banner) {
        banner.style.display = 'block';
        banner.innerHTML = '<strong>Guardado local pendiente de nube.</strong><br>' + detail;
      }
    }
    alert(detail);
  }

  function isMissingRpcError(error) {
    const message = text(error && error.message ? error.message : error);
    return /PGRST202|Could not find the function|schema cache|tms_replace_table/i.test(message);
  }

  function markCloudMigrationPending() {
    APP.cloudMigrationPending = true;
    const banner = document.getElementById('sapEstadoBanner');
    if (banner) {
      banner.style.display = 'block';
      banner.innerHTML = '<strong>Migración de nube pendiente.</strong><br>El sistema guardó en modo compatibilidad. Para guardado transaccional completo falta aplicar supabase/migrations/0004_atomic_replace_and_route_history.sql en Supabase.';
    }
  }

  async function legacyReplaceSupabaseTable(tableName, rows) {
    const client = getSupabaseClient();
    const safeRows = tableName === 'tms_route_cost_history'
      ? (rows || []).map(row => {
        const copy = { ...row };
        delete copy.camion;
        return copy;
      })
      : (rows || []);
    await runSupabaseQuery(client.from(tableName).delete().not('id', 'is', null), 'No se pudo limpiar ' + tableName);
    for (const chunk of chunkRows(safeRows, 200)) {
      await runSupabaseQuery(client.from(tableName).insert(chunk), 'No se pudo guardar ' + tableName);
    }
  }

  async function replaceSupabaseTable(tableName, rows) {
    const client = getSupabaseClient();
    if (!client || typeof client.rpc !== 'function') throw new Error('No hay cliente Supabase para guardar ' + tableName);
    try {
      await runSupabaseQuery(
        client.rpc('tms_replace_table', { p_table_name: tableName, p_rows: rows || [] }),
        'No se pudo guardar ' + tableName + ' de forma transaccional'
      );
    } catch (error) {
      if (!isMissingRpcError(error)) throw error;
      markCloudMigrationPending();
      console.warn('Migración 0004 pendiente; guardando en modo compatibilidad para', tableName);
      await legacyReplaceSupabaseTable(tableName, rows || []);
    }
  }

  function mapDbOrderLine(row) {
    const metadata = row.metadata || {};
    const item = {
      idInterno: row.id,
      clienteId: row.cliente_id,
      clienteNombre: row.cliente_nombre || '',
      rutaId: row.ruta_id || '',
      rutaNombre: row.ruta_nombre || '',
      fechaPlanificada: isoToFechaStr(row.fecha_planificada),
      fechaControl: isoToFechaStr(row.fecha_control),
      pedidoCliente: row.pedido_cliente || '',
      lineaPedidoCliente: row.linea_pedido_cliente || '',
      articulo: row.articulo || '',
      descripcionArticulo: row.descripcion_articulo || '',
      cantidadSolicitada: num(row.cantidad_solicitada),
      cantidadFacturada: num(row.cantidad_facturada),
      cantidadPendiente: num(row.cantidad_pendiente),
      cantidadSolicitadaOriginal: num(metadata.cantidadSolicitadaOriginal || row.cantidad_solicitada),
      cantidadPendienteServir: num(metadata.cantidadPendienteServir || row.cantidad_solicitada),
      cantidadPendienteSinStock: num(metadata.cantidadPendienteSinStock),
      montoPlanificado: num(metadata.montoPlanificado),
      montoEntregadoNoFacturado: num(metadata.montoEntregadoNoFacturado),
      estadoPlanificacion: row.estado_planificacion || 'planificado',
      estadoEntrega: row.estado_entrega || '',
      origen: row.origen || 'planificacion semanal',
      camionAsignado: row.camion_asignado || '',
      fechaCreacion: row.created_at || nowIso(),
      fechaActualizacion: row.updated_at || nowIso(),
      zona: metadata.zona || row.ruta_nombre || '',
      observaciones: metadata.observaciones || '',
      incidencia: metadata.incidencia || '',
      precintoDespacho: metadata.precintoDespacho || '',
      comentarioRuta: metadata.comentarioRuta || '',
      choferRuta: metadata.choferRuta || '',
      fotoCamion: metadata.fotoCamion || '',
      fotoCamionNombre: metadata.fotoCamionNombre || '',
      fechaCierreRuta: metadata.fechaCierreRuta || '',
      queueId: metadata.queueId || '',
      manualProgramado: !!row.manual_programado,
      asignado: !!metadata.asignado
    };
    item.baseKey = buildBaseItemKey(item);
    return item;
  }

  function buildOrderLineDbRows(items, dataset) {
    return items.map(item => {
      const routeCfg = getRouteConfigByName(item.rutaNombre || item.zona || '');
      const routeId = routeCfg && isUuid(routeCfg.id) ? routeCfg.id : null;
      const baseKey = item.baseKey || buildBaseItemKey(item);
      return {
        clave_unica: [dataset, baseKey, item.fechaPlanificada || '', item.fechaControl || '', item.camionAsignado || ''].join('|'),
        cliente_id: text(item.clienteId),
        cliente_nombre: text(item.clienteNombre),
        ruta_id: routeId,
        ruta_nombre: text(item.rutaNombre || item.zona),
        fecha_planificada: fechaStrToISO(item.fechaPlanificada),
        fecha_control: fechaStrToISO(item.fechaControl),
        pedido_cliente: text(item.pedidoCliente),
        linea_pedido_cliente: text(item.lineaPedidoCliente),
        articulo: text(item.articulo),
        descripcion_articulo: text(item.descripcionArticulo),
        cantidad_solicitada: num(item.cantidadSolicitada),
        cantidad_facturada: num(item.cantidadFacturada),
        cantidad_pendiente: num(item.cantidadPendiente),
        estado_planificacion: text(item.estadoPlanificacion || (dataset === 'control' ? '' : 'planificado')) || 'planificado',
        estado_entrega: text(item.estadoEntrega || deriveItemStatus(item)) || 'pendiente',
        origen: text(item.origen || (dataset === 'control' ? 'control diario' : 'planificacion semanal')),
        camion_asignado: text(item.camionAsignado || chooseTruckForItem(item)) || 'CAMION 1',
        manual_programado: !!item.manualProgramado,
        metadata: safeClone({
          dataset,
          montoPlanificado: num(item.montoPlanificado),
          montoEntregadoNoFacturado: num(item.montoEntregadoNoFacturado),
          cantidadSolicitadaOriginal: num(item.cantidadSolicitadaOriginal || item.cantidadSolicitada),
          cantidadPendienteServir: num(item.cantidadPendienteServir || item.cantidadSolicitada),
          cantidadPendienteSinStock: num(item.cantidadPendienteSinStock),
          observaciones: item.observaciones || '',
          incidencia: item.incidencia || '',
          precintoDespacho: item.precintoDespacho || '',
          comentarioRuta: item.comentarioRuta || '',
          choferRuta: item.choferRuta || '',
          fotoCamion: item.fotoCamion || '',
          fotoCamionNombre: item.fotoCamionNombre || '',
          fechaCierreRuta: item.fechaCierreRuta || '',
          queueId: item.queueId || '',
          zona: item.zona || item.rutaNombre || '',
          asignado: !!item.asignado
        })
      };
    });
  }

  function mapDbWarehousePlan(row) {
    const metadata = row.metadata || {};
    return {
      codigo: row.cliente_id,
      cliente: row.cliente_nombre || '',
      solicitud: row.solicitud_id || '',
      pedidoCliente: row.pedido_cliente || '',
      lineaPedidoCliente: metadata.lineaPedidoCliente || '',
      articulo: row.articulo || '',
      descripcionArticulo: row.descripcion_articulo || '',
      cantidadSolicitada: num(row.cantidad_solicitada),
      cantidadCompletada: num(metadata.cantidadCompletada),
      cantidadPendiente: num(metadata.cantidadPendiente),
      estadoComunicacion: metadata.estadoComunicacion || '',
      estadoProceso: metadata.estadoProceso || '',
      fechaLiberacion: isoToFechaStr(metadata.fechaLiberacion),
      fechaFinalizada: isoToFechaStr(metadata.fechaFinalizada),
      fechaEnvio: isoToFechaStr(row.fecha),
      ubicacion: metadata.ubicacion || '',
      observaciones: row.observaciones || '',
      motivoNoPreparada: metadata.motivoNoPreparada || '',
      estado: row.estado || 'Prevista',
      costeSolicitud: num(row.coste_solicitud),
      tipoSolicitudArchivo: metadata.tipoSolicitudArchivo || 'plan semanal'
    };
  }

  function mapDbWarehouseControl(row) {
    const metadata = row.metadata || {};
    return {
      codigo: row.cliente_id,
      cliente: row.cliente_nombre || '',
      solicitud: row.solicitud_id || '',
      pedidoCliente: row.pedido_cliente || '',
      lineaPedidoCliente: metadata.lineaPedidoCliente || '',
      articulo: row.articulo || '',
      descripcionArticulo: row.descripcion_articulo || '',
      cantidadSolicitada: num(row.cantidad_solicitada),
      cantidadCompletada: num(row.cantidad_completada),
      cantidadPendiente: num(row.cantidad_pendiente),
      estadoComunicacion: metadata.estadoComunicacion || '',
      estadoProceso: metadata.estadoProceso || '',
      fechaLiberacion: isoToFechaStr(metadata.fechaLiberacion),
      fechaFinalizada: isoToFechaStr(metadata.fechaFinalizada),
      fechaEnvio: isoToFechaStr(row.fecha),
      ubicacion: metadata.ubicacion || '',
      observaciones: row.observaciones || '',
      motivoNoPreparada: metadata.motivoNoPreparada || '',
      estado: row.estado || 'Pendiente',
      costeSolicitud: num(row.coste_solicitud),
      tipoSolicitudArchivo: metadata.tipoSolicitudArchivo || 'control diario'
    };
  }

  function buildWarehousePlanDbRows(items) {
    return items.map(item => ({
      clave_unica: ['plan', item.codigo, item.solicitud, item.pedidoCliente || '', item.articulo || '', item.fechaEnvio || ''].join('|'),
      fecha: fechaStrToISO(item.fechaEnvio || item.fechaLiberacion),
      cliente_id: text(item.codigo),
      cliente_nombre: text(item.cliente),
      pedido_cliente: text(item.pedidoCliente),
      solicitud_id: text(item.solicitud),
      articulo: text(item.articulo),
      descripcion_articulo: text(item.descripcionArticulo),
      cantidad_solicitada: num(item.cantidadSolicitada),
      estado: text(item.estado || 'Prevista'),
      coste_solicitud: num(item.costeSolicitud),
      observaciones: text(item.observaciones),
      metadata: safeClone({
        cantidadCompletada: num(item.cantidadCompletada),
        cantidadPendiente: num(item.cantidadPendiente),
        estadoComunicacion: item.estadoComunicacion || '',
        estadoProceso: item.estadoProceso || '',
        fechaLiberacion: fechaStrToISO(item.fechaLiberacion),
        fechaFinalizada: fechaStrToISO(item.fechaFinalizada),
        ubicacion: item.ubicacion || '',
        tipoSolicitudArchivo: item.tipoSolicitudArchivo || 'plan semanal',
        lineaPedidoCliente: item.lineaPedidoCliente || '',
        motivoNoPreparada: item.motivoNoPreparada || ''
      })
    }));
  }

  function buildWarehouseControlDbRows(items) {
    return items.map(item => ({
      clave_unica: ['control', item.codigo, item.solicitud, item.pedidoCliente || '', item.articulo || '', item.fechaEnvio || ''].join('|'),
      fecha: fechaStrToISO(item.fechaEnvio || item.fechaLiberacion),
      cliente_id: text(item.codigo),
      cliente_nombre: text(item.cliente),
      pedido_cliente: text(item.pedidoCliente),
      solicitud_id: text(item.solicitud),
      articulo: text(item.articulo),
      descripcion_articulo: text(item.descripcionArticulo),
      cantidad_solicitada: num(item.cantidadSolicitada),
      cantidad_completada: num(item.cantidadCompletada),
      cantidad_pendiente: num(item.cantidadPendiente),
      estado: text(item.estado || 'Pendiente'),
      coste_solicitud: num(item.costeSolicitud),
      observaciones: text(item.observaciones),
      metadata: safeClone({
        estadoComunicacion: item.estadoComunicacion || '',
        estadoProceso: item.estadoProceso || '',
        fechaLiberacion: fechaStrToISO(item.fechaLiberacion),
        fechaFinalizada: fechaStrToISO(item.fechaFinalizada),
        ubicacion: item.ubicacion || '',
        tipoSolicitudArchivo: item.tipoSolicitudArchivo || 'control diario',
        lineaPedidoCliente: item.lineaPedidoCliente || '',
        motivoNoPreparada: item.motivoNoPreparada || ''
      })
    }));
  }

  function moduleFromViewId(viewId) {
    return {
      importar: 'importar',
      dashboard: 'dashboard',
      calendario: 'calendario',
      comercial: 'comercial',
      rutas: 'rutas',
      importaciones: 'importaciones',
      prioridades: 'prioridades',
      configClientes: 'configuracion',
      solicitudesAlmacen: 'solicitudesAlmacen',
      almacen: 'almacen',
      reportes: 'reportes'
    }[viewId] || viewId;
  }

  function hasPermission(moduleName, action) {
    const profile = APP.currentUserProfile;
    if (profile && profile.rol === 'Admin') return true;
    const role = profile && profile.rol ? profile.rol : 'Solo lectura';
    const perms = getPermissionsForRole(role);
    return !!(perms && perms[moduleName] && perms[moduleName][action]);
  }

  function isAdminUser() {
    return !!(APP.currentUserProfile && APP.currentUserProfile.rol === 'Admin');
  }

  function requireAdminUser(actionLabel) {
    if (isAdminUser()) return true;
    alert((actionLabel || 'Esta acción') + ' solo puede realizarla un administrador.');
    return false;
  }

  function userProfileDbPayload(profile, authUserId) {
    return {
      auth_user_id: authUserId || (isUuid(profile.authUserId) ? profile.authUserId : null),
      nombre: profile.nombre || '',
      email: profile.email || '',
      rol: profile.rol || 'Solo lectura',
      permisos_por_modulo: permissionsPayloadForProfile(profile),
      activo: profile.activo !== false
    };
  }

  function getFirstAllowedView() {
    const order = ['importar', 'dashboard', 'calendario', 'comercial', 'rutas', 'importaciones', 'solicitudesAlmacen', 'almacen', 'configClientes'];
    return order.find(viewId => hasPermission(moduleFromViewId(viewId), 'ver')) || 'comercial';
  }

  function enforceCurrentViewPermission() {
    const active = document.querySelector('.view.active');
    const activeId = active && active.id;
    if (!activeId || hasPermission(moduleFromViewId(activeId), 'ver')) return;
    const nextView = getFirstAllowedView();
    const nextBtn = document.getElementById('nav-' + nextView);
    if (typeof window.cambiarVista === 'function') window.cambiarVista(nextView, nextBtn);
  }

  function applyPermissionUi() {
    const navMap = {
      importar: 'importar',
      dashboard: 'dashboard',
      calendario: 'calendario',
      comercial: 'comercial',
      rutas: 'rutas',
      importaciones: 'importaciones',
      prioridades: 'prioridades',
      configClientes: 'configuracion',
      solicitudesAlmacen: 'solicitudesAlmacen',
      almacen: 'almacen'
    };
    Object.entries(navMap).forEach(([viewId, moduleName]) => {
      const allowed = hasPermission(moduleName, 'ver');
      const btn = document.getElementById('nav-' + viewId);
      if (btn) btn.style.display = allowed ? '' : 'none';
      const mobileOption = document.querySelector(`#mobileModuleSelect option[value=\"${viewId}\"]`);
      if (mobileOption) mobileOption.hidden = !allowed;
    });
    syncMobileModuleSelect();
    const activeModule = moduleFromViewId((document.querySelector('.view.active') || {}).id);
    document.body.classList.toggle('role-readonly', !hasPermission(activeModule, 'editar'));
    const saveBtn = document.getElementById('saveStateBtn');
    if (saveBtn) saveBtn.title = isAdminUser() ? 'Guardar cambios globales' : 'Guardar cambios operativos; usuarios y roles solo Admin';
    const badge = document.getElementById('userEmail');
    if (badge && APP.currentUserProfile && APP.currentUserProfile.rol) {
      const email = text(APP.currentUserProfile.email || badge.textContent).split(' · ')[0];
      badge.textContent = `${email} · ${APP.currentUserProfile.rol}`;
    }
    enforceCurrentViewPermission();
  }

  function syncMobileModuleSelect() {
    const select = document.getElementById('mobileModuleSelect');
    const active = document.querySelector('.view.active');
    if (select && active && select.value !== active.id) select.value = active.id;
  }

  const originalCambiarVista = window.cambiarVista;
  if (typeof originalCambiarVista === 'function' && !window.__tmsPermissionNavPatched) {
    window.__tmsPermissionNavPatched = true;
    window.cambiarVista = function cambiarVistaConPermisos(id, btn) {
      const moduleName = moduleFromViewId(id);
      if (!hasPermission(moduleName, 'ver')) {
        alert('Tu rol no tiene acceso a esta vista.');
        applyPermissionUi();
        return;
      }
      originalCambiarVista(id, btn);
      syncMobileModuleSelect();
      if (id === 'rutas') renderRouteCatalog();
      if (id === 'importaciones') renderImportaciones();
      if (id === 'prioridades') renderPrioridades();
      applyPermissionUi();
    };
  }

  function isHoliday(fechaStr) {
    return APP.feriadosRD.includes(fechaStr);
  }

  function isNonWorkingDay(fechaStr) {
    const d = fechaToDate(fechaStr);
    return d.getDay() === 0 || d.getDay() === 6 || isHoliday(fechaStr);
  }

  function isManualRouteDate(fechaStr) {
    const d = fechaToDate(fechaStr);
    return d.getDay() !== 0 && !isHoliday(fechaStr);
  }

  function nextWorkingDay(fechaStr) {
    const d = fechaToDate(fechaStr);
    do {
      d.setDate(d.getDate() + 1);
    } while (isNonWorkingDay(fechaToStr(d)));
    return fechaToStr(d);
  }

  function appendWorkingDates(startDate, count) {
    const out = [];
    let current = startDate;
    while (out.length < count) {
      if (!isNonWorkingDay(current)) out.push(current);
      current = nextWorkingDay(current);
    }
    return out;
  }

  function routeAllowsDate(routeName, fechaStr) {
    const routeCfg = getRouteConfigByName(routeName);
    if (!routeCfg || !routeCfg.diasOperacion || !routeCfg.diasOperacion.length) return true;
    const dow = fechaToDate(fechaStr).getDay();
    return routeCfg.diasOperacion.includes(dow);
  }

  function clienteAllowsDate(item, fechaStr) {
    const cfg = getClienteConfigV2(item.clienteId, item.clienteNombre);
    const d = fechaToDate(fechaStr);
    const dow = d.getDay();
    if (cfg.noMiercoles && dow === 3) return false;
    if (cfg.noUltimaSemana && esUltimaSemanaDelMes(fechaStr)) return false;
    if (cfg.diasRecepcion && cfg.diasRecepcion.length > 0 && !cfg.diasRecepcion.includes(dow)) return false;
    if (!routeAllowsDate(item.rutaNombre || item.zona, fechaStr)) return false;
    if (isNonWorkingDay(fechaStr)) return false;
    return true;
  }

  function chooseTruckForItem(item) {
    const cfg = getClienteConfigV2(item.clienteId, item.clienteNombre);
    if (cfg.camionPermitido && cfg.camionPermitido !== 'CUALQUIERA') return cfg.camionPermitido;
    return asignarCamionPorZona(item.zona || item.rutaNombre || '');
  }

  function calcProgressSummary(items) {
    const requested = items.reduce((sum, item) => sum + (item.cantidadSolicitada || 0), 0);
    const invoiced = items.reduce((sum, item) => sum + (item.cantidadFacturada || 0), 0);
    const pending = items.reduce((sum, item) => sum + (item.cantidadPendiente || 0), 0);
    const pct = requested > 0 ? Math.round((invoiced / requested) * 100) : 0;
    return { requested, invoiced, pending, pct };
  }

  function getUsableRequestedQty(pendingToServe, pendingNoStock) {
    return Math.max(num(pendingToServe) - num(pendingNoStock), 0);
  }

  function getConfirmedUndeliveredAmount(item) {
    return num(item.montoPlanificado || 0);
  }

  function getDeliveredNotInvoicedAmount(item) {
    return num(item.montoEntregadoNoFacturado || item.importeEntregadoNoFacturado || item.montoPendienteFacturar || 0);
  }

  function getOperationalAmount(item) {
    return getConfirmedUndeliveredAmount(item) + getDeliveredNotInvoicedAmount(item);
  }

  function getClientConfirmedUndelivered(client) {
    if (client && Array.isArray(client.items)) {
      return client.items.reduce((sum, item) => sum + getConfirmedUndeliveredAmount(item), 0);
    }
    return num(client && client.monto);
  }

  function getClientDeliveredNotInvoiced(client) {
    if (client && Array.isArray(client.items)) {
      return client.items.reduce((sum, item) => sum + getDeliveredNotInvoicedAmount(item), 0);
    }
    return num(client && (client.pendienteFacturar || client.pf));
  }

  function getClientOperationalAmount(client) {
    return getClientConfirmedUndelivered(client) + getClientDeliveredNotInvoiced(client);
  }

  function deriveItemStatus(item) {
    if ((item.cantidadPendiente || 0) <= 0 && (item.cantidadFacturada || 0) > 0) return 'entregado';
    if ((item.cantidadFacturada || 0) > 0 && (item.cantidadPendiente || 0) > 0) return 'parcial';
    if ((item.cantidadPendiente || 0) > 0) return 'pendiente';
    return 'planificado';
  }

  function aggregateClientRows(lineItems) {
    const groups = new Map();
    lineItems.forEach(item => {
      const key = [item.clienteId, item.fechaPlanificada || '', item.camionAsignado || ''].join('|');
      if (!groups.has(key)) {
        groups.set(key, {
          id: key,
          codigo: item.clienteId,
          nombre: item.clienteNombre,
          razonSocial: item.clienteNombre,
          pedido: item.pedidoCliente || '',
          monto: 0,
          pendienteFacturar: 0,
          totalPendiente: 0,
          cantidadPedidos: 0,
          entregado: 0,
          cumplimiento: 0,
          fecha: item.fechaPlanificada || '',
          camion: item.camionAsignado || 'CAMION 1',
          zona: item.rutaNombre || item.zona || '',
          sector: item.rutaNombre || item.zona || '',
          ubicacion: '',
          lat: null,
          lon: null,
          mapeado: false,
          cumplida: false,
          nota: '',
          incidencia: '',
          configNota: getClienteConfigV2(item.clienteId, item.clienteNombre).observaciones || '',
          manualProgramado: item.origen === 'manual',
          origenQueue: item.origen === 'manual',
          queueId: item.queueId || ''
        });
      }
      const group = groups.get(key);
      group.monto += getConfirmedUndeliveredAmount(item);
      group.pendienteFacturar += getDeliveredNotInvoicedAmount(item);
      group.totalPendiente += getOperationalAmount(item);
      group.items = group.items || [];
      group.items.push(item);
      group.cantidadPedidos += 1;
      group.entregado += item.cantidadFacturada || 0;
      group.cumplida = group.cumplida || deriveItemStatus(item) === 'entregado';
    });
    return [...groups.values()].map(group => {
      group.cumplimiento = group.totalPendiente > 0 ? Math.round((group.entregado / group.totalPendiente) * 100) : 0;
      return group;
    });
  }

  function rebuildDerivedState() {
    APP.lineItems.forEach(item => {
      item.estadoEntrega = deriveItemStatus(item);
      item.estadoPlanificacion = item.manualProgramado ? 'manual' : 'planificado';
      item.rutaNombre = item.rutaNombre || inferRouteName(item.clienteNombre, item.clienteId);
      item.zona = item.zona || item.rutaNombre;
      item.camionAsignado = item.camionAsignado || chooseTruckForItem(item);
      item.baseKey = buildBaseItemKey(item);
      item.deliveryGroupKey = buildDeliveryGroupKey(item);
    });

    APP.planLineItems.forEach(item => {
      item.baseKey = buildBaseItemKey(item);
      item.estadoEntrega = deriveItemStatus(item);
    });

    APP.controlLineItems.forEach(item => {
      item.baseKey = buildBaseItemKey(item);
      item.estadoEntrega = deriveItemStatus(item);
      item.rutaNombre = item.rutaNombre || inferRouteName(item.clienteNombre, item.clienteId);
      item.zona = item.zona || item.rutaNombre;
    });

    APP.clientes = aggregateClientRows(APP.lineItems);
    APP.planificacion = aggregateClientRows(APP.planLineItems.length ? APP.planLineItems : APP.lineItems);
    APP.controlDiario = aggregateClientRows(APP.controlLineItems);
    buildCalStructure();
    construirRutas();
    construirQueueSemana();
    initRouteFilterOptions();
  }

  function makeUniqueHeader(label, previous, counts) {
    let base = text(label);
    const prev = text(previous);
    if (!base && prev) base = prev + ' descripción';
    if (!base) base = 'Columna';
    const seen = counts[base] || 0;
    counts[base] = seen + 1;
    return seen ? base + ' ' + (seen + 1) : base;
  }

  function rowsFromWorksheet(ws) {
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    const headerIndex = matrix.findIndex(row => {
      const normalized = row.map(normKey);
      return normalized.includes('pedidodecliente') ||
        normalized.includes('iddesolicituddealmacen') ||
        normalized.includes('iddesolicituddelogisticapormediodeterceros');
    });
    if (headerIndex < 0) return XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
    const headerRow = matrix[headerIndex];
    const counts = {};
    const headers = headerRow.map((cell, index) => makeUniqueHeader(cell, headerRow[index - 1], counts));
    return matrix.slice(headerIndex + 1)
      .filter(row => row.some(cell => text(cell)))
      .map(row => Object.fromEntries(headers.map((header, index) => [header, row[index] == null ? '' : row[index]])));
  }

  function parseSAPRows(arrayBuffer, fileName, mode) {
    const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellDates: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const fechaArchivo = extraerFechaDeNombreArchivo(sheetName) || extraerFechaDeNombreArchivo(fileName) || fechaToStr(new Date());
    const rows = rowsFromWorksheet(ws);

    const items = [];
    rows.forEach((row, index) => {
      const clienteId = text(pickField(row, ['Cliente', 'Codigo', 'Código', 'ID cliente', 'Cliente código', 'Destinatario de las mercancías descripción']));
      const clienteNombre = text(pickField(row, ['Nombre Comercial', 'Nombre', 'Cliente descripción', 'Cliente Nombre', 'Destinatario', 'Nombre cliente', 'Destinatario de las mercancías']));
      const pedidoCliente = text(pickField(row, ['Pedido de cliente', 'Pedido del cliente', 'Pedido cliente', 'Pedido']));
      if (!clienteId || !clienteNombre || !pedidoCliente) return;

      const lineaPedidoCliente = text(pickField(row, ['Posición de pedido de cliente', 'Posicion de pedido de cliente', 'Línea de pedido del cliente', 'Linea de pedido del cliente', 'ID de partida individual', 'Posición', 'Linea']));
      const articulo = text(pickField(row, ['Producto', 'Artículo', 'Articulo', 'Material', 'SKU']));
      const descripcionArticulo = text(pickField(row, ['Producto descripción', 'Producto descripcion', 'Artículo descripción', 'Articulo descripcion', 'Descripción del artículo', 'Descripcion del articulo', 'Descripción', 'Descripcion', 'Texto breve']));
      const cantidadSolicitadaOriginal = num(pickField(row, ['Cantidad solicitada', 'Cantidad solicitada cliente', 'Cantidad', 'Cant solicitada']));
      const cantidadFacturada = num(pickField(row, ['Cantidad facturada', 'Facturado', 'Cantidad entregada']));
      const cantidadPendienteRaw = pickField(row, ['Cantidad pendiente de servir', 'Cantidad pendiente', 'Pendiente', 'Cantidad abierta']);
      const cantidadPendienteServir = cantidadPendienteRaw !== '' ? num(cantidadPendienteRaw) : Math.max(cantidadSolicitadaOriginal - cantidadFacturada, 0);
      const cantidadPendienteSinStock = num(pickField(row, ['Cantidad pendiente sin stock', 'Cantidad sin stock', 'Pendiente sin stock', 'Cant pendiente sin stock']));
      const cantidadSolicitada = getUsableRequestedQty(cantidadPendienteServir, cantidadPendienteSinStock);
      const cantidadPendiente = Math.max(cantidadSolicitada - cantidadFacturada, 0);
      const montoPlanificado = num(pickField(row, [
        'Importe confirmado no entregado',
        'Importe Confirmado No Entregado',
        'Importe confirmado, no entregado',
        'Confirmado no entregado',
        'Confirmado No Entregado',
        'Facturado no entregado',
        'Facturado No Entregado',
        'Importe confirmado no entregado (moneda de transacción)',
        'Importe confirmado no entregado moneda de transacción',
        'Imp. confirmado no entregado'
      ]));
      const montoEntregadoNoFacturado = num(pickField(row, [
        'Importe entregado no facturado',
        'Importe Entregado No Facturado',
        'Importe entregado, no facturado',
        'Entregado no facturado',
        'Entregado No Facturado',
        'Pte de facturar',
        'Pte. de facturar',
        'Pte Facturar',
        'Importe entregado no facturado (moneda de transacción)',
        'Importe entregado no facturado moneda de transacción',
        'Imp. entregado no facturado',
        'Pendiente de facturar',
        'Importe pendiente de facturar'
      ]));
      const montoPendienteSinStock = num(pickField(row, ['Importe pendiente sin stock']));
      const fechaPrevista = toDateLabel(pickField(row, ['Fecha', 'Fecha Entrega', 'Fecha de entrega', 'Fecha de envío planificada', 'Fecha de envio planificada'])) || fechaArchivo;
      const rutaNombre = inferRouteName(clienteNombre, clienteId);
      const item = {
        idInterno: [mode, index + 1, nowIso()].join('-'),
        clienteId,
        clienteNombre,
        rutaId: text(pickField(row, ['Ruta ID', 'Id ruta'])),
        rutaNombre,
        fechaPlanificada: fechaPrevista,
        fechaControl: mode === 'control' ? fechaArchivo : '',
        pedidoCliente,
        lineaPedidoCliente,
        referenciaExterna: text(pickField(row, ['Referencia externa'])),
        articulo,
        descripcionArticulo,
        cantidadSolicitadaOriginal,
        cantidadPendienteServir,
        cantidadPendienteSinStock,
        cantidadSolicitada,
        cantidadFacturada,
        cantidadPendiente,
        montoPlanificado,
        montoEntregadoNoFacturado,
        montoPendienteSinStock,
        estadoCabecera: text(pickField(row, ['Estado (cabecera) (Pedido de cliente)', 'Estado cabecera'])),
        estadoArticulo: text(pickField(row, ['Estado de artículo', 'Estado de articulo'])),
        estadoPlanificacion: mode === 'control' ? '' : 'planificado',
        estadoEntrega: '',
        origen: mode === 'control' ? 'control diario' : 'programación mensual',
        camionAsignado: '',
        fechaCreacion: nowIso(),
        fechaActualizacion: nowIso(),
        zona: rutaNombre,
        observaciones: '',
        manualProgramado: false,
        asignado: false
      };
      item.baseKey = buildBaseItemKey(item);
      item.sapMatchKey = buildSapMatchKey(item);
      items.push(item);
    });

    return {
      fecha: fechaArchivo,
      items: dedupeBy(items, item => item.baseKey + '|' + (mode === 'control' ? item.fechaControl : item.fechaPlanificada))
    };
  }

  function schedulePlanItems(items, fechaPlanBase) {
    const startDate = fechaInicioPlanificacion(fechaPlanBase || fechaToStr(new Date()));
    const candidateDates = appendWorkingDates(startDate, 15);
    const orderGroups = new Map();
    items.forEach(item => {
      const key = [item.clienteId, item.pedidoCliente || item.baseKey].join('|');
      if (!orderGroups.has(key)) orderGroups.set(key, []);
      orderGroups.get(key).push(item);
    });

    const loadMap = {};
    orderGroups.forEach(groupItems => {
      const ref = groupItems[0];
      const preferredTruck = chooseTruckForItem(ref);
      const routeName = ref.rutaNombre || ref.zona;
      const allowedDates = candidateDates.filter(date => clienteAllowsDate(ref, date) && routeAllowsDate(routeName, date));
      const datesToUse = allowedDates.length ? allowedDates : candidateDates.filter(date => !isNonWorkingDay(date));
      let bestDate = datesToUse[0] || startDate;
      let bestLoad = Infinity;
      datesToUse.forEach(date => {
        const key = [date, preferredTruck].join('|');
        const load = loadMap[key] || 0;
        if (load < bestLoad) {
          bestLoad = load;
          bestDate = date;
        }
      });
      const key = [bestDate, preferredTruck].join('|');
      loadMap[key] = (loadMap[key] || 0) + groupItems.reduce((sum, item) => sum + (item.cantidadSolicitada || 1), 0);
      groupItems.forEach(item => {
        item.fechaPlanificada = bestDate;
        item.camionAsignado = preferredTruck;
        item.rutaNombre = routeName;
        item.zona = routeName;
      });
    });
  }

  function syncMoveOptions() {
    const moveSelect = document.getElementById('moveCamion');
    if (!moveSelect) return;
    moveSelect.innerHTML = getTruckOptions().map(camion => `<option value="${camion}">${getTruckLabel(camion)}</option>`).join('');
  }

  function syncCalendarToolbarState() {
    const summary = document.getElementById('calSelectionSummary');
    const moveBtn = document.getElementById('moveSelectedBtn');
    const clearBtn = document.getElementById('clearSelectedBtn');
    const count = (APP.selectedCalendarGroups || []).length;
    if (summary) summary.textContent = `${count} seleccionados`;
    if (moveBtn) moveBtn.disabled = count === 0;
    if (clearBtn) clearBtn.disabled = count === 0;
  }

  function saveLocalSnapshot() {
    const payload = {
      savedAt: nowIso(),
      app: {
        lineItems: APP.lineItems,
        planLineItems: APP.planLineItems,
      controlLineItems: APP.controlLineItems,
      controlHistory: APP.controlHistory,
      solicitudesAlmacen: APP.solicitudesAlmacen,
      solicitudesPlanAlmacen: APP.solicitudesPlanAlmacen,
      solicitudesControlAlmacen: APP.solicitudesControlAlmacen,
      solicitudesCompare: APP.solicitudesCompare,
      routeConfigs: APP.routeConfigs,
        routeCostHistory: APP.routeCostHistory,
        controlHistory: APP.controlHistory,
        solicitudesHistory: APP.solicitudesHistory,
        warehouseSettings: APP.warehouseSettings,
        userProfiles: APP.userProfiles,
        rolePermissions: APP.rolePermissions,
        clienteConfig: APP.clienteConfig,
        planLoaded: APP.planLoaded,
        planFecha: APP.planFecha,
        controlLoaded: APP.controlLoaded,
        controlFecha: APP.controlFecha,
        solicitudesLoaded: APP.solicitudesLoaded,
        solicitudesFileName: APP.solicitudesFileName,
        camionExtraEnabled: APP.camionExtraEnabled,
        feriadosRD: APP.feriadosRD,
        importFiles: APP.importFiles,
        importShipments: APP.importShipments,
        dailyPriorities: APP.dailyPriorities,
        prioritySpainTopics: APP.prioritySpainTopics,
        calendarNotes: APP.calendarNotes,
        visualTheme: APP.visualTheme
      }
    };
    localStorage.setItem(LOCAL_KEY, JSON.stringify(payload));
  }

  function loadEmbeddedDemoState() {
    const payload = window.ROUTECONTROL_DATA && window.ROUTECONTROL_DATA.data;
    if (!payload || !Array.isArray(payload.lineItems) || !payload.lineItems.length) return false;
    Object.assign(APP, safeClone(payload));
    APP.planLineItems = APP.planLineItems || APP.lineItems.map(item => ({ ...item }));
    APP.controlLineItems = APP.controlLineItems || [];
    APP.controlHistory = APP.controlHistory || [];
    APP.solicitudesPlanAlmacen = APP.solicitudesPlanAlmacen || [];
    APP.solicitudesControlAlmacen = APP.solicitudesControlAlmacen || [];
    APP.solicitudesAlmacen = APP.solicitudesAlmacen || APP.solicitudesPlanAlmacen;
    APP.solicitudesCompare = buildSolicitudesComparison();
    APP.visualTheme = normalizeVisualTheme(APP.visualTheme || localStorage.getItem(THEME_KEY));
    applyVisualTheme(APP.visualTheme, true);
    ensureAdminProfile();
    matchWarehouseToOrders();
    rebuildDerivedState();
    return true;
  }

  function loadLocalSnapshot() {
    try {
      const raw = localStorage.getItem(LOCAL_KEY);
      if (!raw) return false;
      const payload = JSON.parse(raw);
      Object.assign(APP, payload.app || {});
      APP.visualTheme = normalizeVisualTheme(APP.visualTheme || localStorage.getItem(THEME_KEY));
      applyVisualTheme(APP.visualTheme, true);
      ensureAdminProfile();
      rebuildDerivedState();
      return true;
    } catch (error) {
      console.warn('No se pudo leer snapshot local:', error);
      return false;
    }
  }

  window.scheduleAutoSave = function scheduleAutoSaveV2() {
    if (window._autoSaveTimer) clearTimeout(window._autoSaveTimer);
    window._autoSaveTimer = setTimeout(() => {
      saveLocalSnapshot();
      window.guardarEnSupabase({ silent: true }).catch(e => console.warn('AutoSave error:', e));
    }, 1200);
  };

  function buildRouteGroups() {
    const routes = {};
    const grouped = new Map();
    APP.lineItems.forEach(item => {
      const key = buildDeliveryGroupKey(item);
      if (!grouped.has(key)) {
        grouped.set(key, {
          id: key,
          fecha: item.fechaPlanificada,
          camion: item.camionAsignado,
          codigo: item.clienteId,
          nombre: item.clienteNombre,
          pedidoCliente: item.pedidoCliente || 'Sin pedido',
          pedidos: [],
          solicitudes: [],
          referencias: [],
          zona: item.rutaNombre || item.zona || '',
          rutaNombre: item.rutaNombre || item.zona || '',
          cantidadSolicitada: 0,
          cantidadFacturada: 0,
          cantidadPendiente: 0,
          cargaSolicitudes: 0,
          cargaTotal: 0,
          montoSolicitudes: 0,
          totalPendiente: 0,
          items: [],
          cumplida: false,
          nota: '',
          incidencia: '',
          precintoDespacho: '',
          comentarioRuta: '',
          choferRuta: '',
          fotoCamion: '',
          fotoCamionNombre: '',
          fechaCierreRuta: '',
          manualProgramado: !!item.manualProgramado,
          alertas: []
        });
      }
      const group = grouped.get(key);
      group.items.push(item);
      group.pedidos = uniqueTexts([...group.pedidos, item.pedidoCliente]);
      group.solicitudes = uniqueTexts([...group.solicitudes, item.solicitudAlmacen]);
      group.referencias = uniqueTexts([...group.referencias, item.articulo]);
      group.pedidoCliente = group.pedidos.join(', ') || 'Sin pedido';
      group.cantidadPedidos = group.pedidos.length || group.items.length;
      group.cantidadSolicitudes = group.solicitudes.length;
      group.cantidadSolicitada += item.cantidadSolicitada || 0;
      group.cantidadFacturada += item.cantidadFacturada || 0;
      group.cantidadPendiente += item.cantidadPendiente || 0;
      group.cargaSolicitudes += item.cantidadAlmacenSolicitada || 0;
      group.cargaTotal = group.cantidadSolicitada + group.cargaSolicitudes;
      if (item.solicitudAlmacen || item.cantidadAlmacenSolicitada) {
        group.montoSolicitudes += getOperationalAmount(item);
      }
      group.totalPendiente += getOperationalAmount(item);
      group.cumplida = group.items.every(current => deriveItemStatus(current) === 'entregado');
      group.nota = group.nota || item.observaciones || '';
      group.incidencia = group.incidencia || item.incidencia || '';
      group.precintoDespacho = group.precintoDespacho || item.precintoDespacho || '';
      group.comentarioRuta = group.comentarioRuta || item.comentarioRuta || '';
      group.choferRuta = group.choferRuta || item.choferRuta || '';
      group.fotoCamion = group.fotoCamion || item.fotoCamion || '';
      group.fotoCamionNombre = group.fotoCamionNombre || item.fotoCamionNombre || '';
      group.fechaCierreRuta = group.fechaCierreRuta || item.fechaCierreRuta || '';
    });

    grouped.forEach(group => {
      const cfg = getClienteConfigV2(group.codigo, group.nombre);
      if (!cfg.configuracionCompleta) group.alertas.push('Cliente sin configuración');
      if (!cfg.rutaAsignada && !group.rutaNombre) group.alertas.push('Cliente sin ruta asignada');
      if (isHoliday(group.fecha)) group.alertas.push('Pedido en feriado');
      if (!routeAllowsDate(group.rutaNombre || group.zona, group.fecha)) group.alertas.push('Ruta fuera de día operativo');
      if (!routes[group.fecha]) {
        routes[group.fecha] = {};
        getTruckOptions().forEach(camion => { routes[group.fecha][camion] = []; });
      }
      if (!routes[group.fecha][group.camion]) routes[group.fecha][group.camion] = [];
      routes[group.fecha][group.camion].push(group);
    });
    return routes;
  }

  function getCalendarMonthKey(fechaStr) {
    const d = fechaToDate(fechaStr);
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  }

  function getCalendarMonthLabel(fechaStr) {
    const d = fechaToDate(fechaStr);
    return new Intl.DateTimeFormat('es-DO', { month: 'long', year: 'numeric' })
      .format(d)
      .replace(/^\w/, char => char.toUpperCase());
  }

  function buildCalendarMonths(semanas) {
    const map = new Map();
    semanas.forEach((semana, index) => {
      const key = getCalendarMonthKey(semana.dias[0] || semana.key);
      if (!map.has(key)) {
        map.set(key, {
          key,
          label: getCalendarMonthLabel(semana.dias[0] || semana.key),
          semanas: [],
          firstSemanaIdx: index
        });
      }
      map.get(key).semanas.push({ ...semana, absoluteIndex: index });
    });
    return [...map.values()].sort((a, b) => a.key.localeCompare(b.key));
  }

  window.buildCalStructure = function buildCalStructureV2() {
    const prevMonthKey = APP.calMeses && APP.calMeses[APP.calMesIdx] ? APP.calMeses[APP.calMesIdx].key : null;
    const fechas = [...new Set(APP.lineItems.map(item => item.fechaPlanificada).filter(Boolean))].sort((a, b) => fechaToDate(a) - fechaToDate(b));
    APP.calSemanas = construirSemanasCalendario().filter(semana => semana.dias.some(fecha => fechas.includes(fecha)));
    fechas.forEach(fecha => {
      const weekKey = fechaToStr(lunesDeSemana(fecha));
      let week = APP.calSemanas.find(semana => semana.key === weekKey);
      if (!week) {
        const lunes = fechaToDate(weekKey);
        const dias = semanaLaborableDesde(lunes);
        week = { key: weekKey, mesNombre: getCalendarMonthLabel(fecha), dias };
        APP.calSemanas.push(week);
      }
      if (!week.dias.includes(fecha)) week.dias.push(fecha);
      week.dias = [...new Set(week.dias)].sort((a, b) => fechaToDate(a) - fechaToDate(b));
      week.nombre = fechaLabel(week.dias[0]) + ' ' + week.dias[0] + ' - ' + fechaLabel(week.dias[week.dias.length - 1]) + ' ' + week.dias[week.dias.length - 1];
    });
    APP.calSemanas = APP.calSemanas.sort((a, b) => fechaToDate(a.key) - fechaToDate(b.key));
    if (!APP.calSemanas.length && fechas.length) {
      const inicio = lunesDeSemana(fechas[0]);
      APP.calSemanas = [{
        key: fechaToStr(inicio),
        nombre: fechas[0] + ' - ' + (fechas[fechas.length - 1] || fechas[0]),
        mesNombre: getCalendarMonthLabel(fechas[0]),
        dias: fechas
      }];
    }
    APP.calMeses = buildCalendarMonths(APP.calSemanas || []);
    if (APP.calMeses.length) {
      const preservedIdx = prevMonthKey ? APP.calMeses.findIndex(month => month.key === prevMonthKey) : -1;
      APP.calMesIdx = preservedIdx >= 0 ? preservedIdx : Math.min(APP.calMesIdx || 0, APP.calMeses.length - 1);
      APP.calSemanaIdx = APP.calMeses[APP.calMesIdx].firstSemanaIdx || 0;
    } else {
      APP.calMesIdx = 0;
      APP.calSemanaIdx = 0;
    }
  };

  window.construirRutas = function construirRutasV2() {
    APP.rutas = buildRouteGroups();
  };

  window.construirQueueSemana = function construirQueueSemanaV2() {
    const plannedKeys = new Set(APP.lineItems.map(item => item.baseKey));
    const queue = APP.controlLineItems
      .filter(item => (item.cantidadPendiente || 0) > 0 && !plannedKeys.has(item.baseKey))
      .map(item => {
        const cfg = getClienteConfigV2(item.clienteId, item.clienteNombre);
        const routeName = item.rutaNombre || item.zona || inferRouteName(item.clienteNombre, item.clienteId);
        return {
          queueId: item.baseKey,
          codigo: item.clienteId,
          nombre: item.clienteNombre,
          pedidoCliente: item.pedidoCliente || 'Sin pedido',
          lineaPedidoCliente: item.lineaPedidoCliente || '',
          articulo: item.articulo || '',
          descripcionArticulo: item.descripcionArticulo || '',
          cantidadSolicitada: item.cantidadSolicitada || 0,
          cantidadFacturada: item.cantidadFacturada || 0,
          cantidadPendiente: item.cantidadPendiente || 0,
          monto: getOperationalAmount(item),
          pf: getDeliveredNotInvoicedAmount(item),
          totalPendiente: getOperationalAmount(item),
          planificado: false,
          fechaControl: item.fechaControl || APP.controlFecha || '',
          zona: routeName,
          rutaNombre: routeName,
          camion: chooseTruckForItem(item),
          asignado: false,
          fechaAsignada: '',
          camionAsignado: '',
          item
        };
      });
    APP.queueSemana = dedupeBy(queue, item => item.queueId + '|' + item.fechaControl);
  };

  function queueByRouteName(item) {
    return item.rutaNombre || item.zona || 'Sin ruta / Configuración pendiente';
  }

  window.renderQueueSemana = function renderQueueSemanaV2() {
    const section = document.getElementById('queueSemanaSection');
    const cont = document.getElementById('queueSemanaContainer');
    const resumen = document.getElementById('queueSemanaResumen');
    if (!section || !cont || !resumen) return;

    if (!APP.controlLoaded) {
      section.style.display = 'none';
      cont.innerHTML = '';
      return;
    }

    const search = text((document.getElementById('calSearch') || {}).value).toLowerCase();
    const visibles = APP.queueSemana.filter(item => {
      if (!search) return true;
      return [item.nombre, item.codigo, item.pedidoCliente, item.articulo, item.descripcionArticulo, item.rutaNombre]
        .some(value => text(value).toLowerCase().includes(search));
    });

    section.style.display = 'block';
    resumen.textContent = `${visibles.length} líneas pendientes`;

    if (!visibles.length) {
      cont.innerHTML = '<div class="empty-state"><div class="empty-icon">📥</div><p>No hay pedidos pendientes en QUEUE para este filtro.</p></div>';
      return;
    }

    const grouped = {};
    visibles.forEach(item => {
      const index = APP.queueSemana.indexOf(item);
      const routeName = queueByRouteName(item);
      const clientKey = [routeName, item.codigo, item.nombre].join('|');
      if (!grouped[routeName]) grouped[routeName] = {};
      if (!grouped[routeName][clientKey]) grouped[routeName][clientKey] = { codigo: item.codigo, nombre: item.nombre, routeName, items: [], indexes: [] };
      grouped[routeName][clientKey].items.push(item);
      grouped[routeName][clientKey].indexes.push(index);
    });

    cont.innerHTML = `<div class="rutas-grid">${Object.keys(grouped).sort().map(routeName => {
      const clients = Object.values(grouped[routeName]);
      const warn = routeName.toLowerCase().includes('sin ruta');
      const lineCount = clients.reduce((sum, client) => sum + client.items.length, 0);
      return `<div class="ruta-col">
        <div class="ruta-col-header ${warn ? 'ch-alm' : 'ch-c1'} queue-route-header">
          <span>${routeName} (${clients.length} clientes · ${lineCount} líneas)</span>
          <button class="btn btn-outline btn-sm" onclick="abrirMoveRouteQueueModal('${jsString(routeName)}')">Asignar ruta completa</button>
        </div>
        <div class="ruta-col-body">
          ${clients.map(client => {
            const cfg = getClienteConfigV2(client.codigo, client.nombre);
            const alerts = [];
            if (!cfg.configuracionCompleta) alerts.push('Cliente sin configuración');
            if (!cfg.rutaAsignada && !client.routeName) alerts.push('Cliente sin ruta asignada');
            const pedidos = uniqueTexts(client.items.map(item => item.pedidoCliente));
            const referencias = uniqueTexts(client.items.map(item => item.articulo));
            const solicitado = client.items.reduce((sum, item) => sum + (item.cantidadSolicitada || 0), 0);
            const facturado = client.items.reduce((sum, item) => sum + (item.cantidadFacturada || 0), 0);
            const pendiente = client.items.reduce((sum, item) => sum + (item.cantidadPendiente || 0), 0);
            const monto = client.items.reduce((sum, item) => sum + (item.monto || getOperationalAmount(item.item || item) || 0), 0);
            return `<div class="queue-card ${warn ? 'queue-card-new' : 'queue-card-plan'}" draggable="true"
              ondragstart="onDragStartQueue(event,${client.indexes[0]})"
              ondragend="APP.dragQueue=null;this.classList.remove('dragging')">
              <div class="queue-card-title">${client.nombre}</div>
              <div class="queue-card-meta">${client.codigo} · ${pedidos.length} pedidos: ${pedidos.join(', ')}</div>
              <div class="queue-card-meta" style="margin-top:6px;">Referencias: <strong>${referencias.slice(0, 6).join(', ') || 'Sin referencia'}</strong>${referencias.length > 6 ? ' +' + (referencias.length - 6) : ''}</div>
              <div class="queue-card-meta" style="margin-top:6px;">Solicitado: <strong>${solicitado}</strong> · Facturado: <strong>${facturado}</strong> · Pendiente: <strong>${pendiente}</strong> · Monto: <strong>${formatMonto(monto)}</strong></div>
              ${alerts.length ? `<div style="font-size:11px;color:var(--danger);margin-top:6px;">${alerts.join(' · ')}</div>` : ''}
              <button class="btn btn-outline btn-sm" style="margin-top:8px;" onclick="abrirMoveQueueModal(${client.indexes[0]})">Asignar fecha/camión</button>
            </div>`;
          }).join('')}
        </div>
      </div>`;
    }).join('')}</div>`;
  };

  function getCalendarNoteAuthor() {
    const profile = APP.currentUserProfile || {};
    return text(profile.nombre || profile.email || 'Usuario');
  }

  function ensureCalendarNotes() {
    APP.calendarNotes = Array.isArray(APP.calendarNotes) ? APP.calendarNotes : [];
    APP.calendarNotes.forEach(note => {
      note.id = note.id || ('calnote-' + Date.now() + '-' + Math.random().toString(16).slice(2));
      note.fecha = text(note.fecha);
      note.texto = text(note.texto || note.nota || '');
      note.author = text(note.author || note.autor || getCalendarNoteAuthor());
      note.createdAt = note.createdAt || nowIso();
      note.updatedAt = note.updatedAt || note.createdAt;
    });
    APP.calendarNotes = APP.calendarNotes.filter(note => note.fecha && note.texto);
  }

  function getCalendarNotesForDate(fecha) {
    ensureCalendarNotes();
    return APP.calendarNotes
      .filter(note => note.fecha === fecha)
      .sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
  }

  function renderCalendarNotesHtml(fecha) {
    const notes = getCalendarNotesForDate(fecha);
    return `<div class="calendar-notes-wrap">
      <div class="calendar-notes-head">
        <span>${notes.length ? notes.length + ' nota' + (notes.length === 1 ? '' : 's') : 'Notas'}</span>
        <button class="calendar-note-add" onclick="agregarNotaCalendario('${jsString(fecha)}')">+ Nota</button>
      </div>
      ${notes.length ? `<div class="calendar-note-list">${notes.map(note => `<button class="calendar-note-postit" onclick="editarNotaCalendario('${jsString(note.id)}')" title="Editar nota">
        <span>${escapeHtml(note.texto)}</span>
        <small>${escapeHtml(note.author)} · ${new Date(note.updatedAt || note.createdAt || nowIso()).toLocaleTimeString('es-DO', { hour: '2-digit', minute: '2-digit' })}</small>
      </button>`).join('')}</div>` : ''}
    </div>`;
  }

  window.agregarNotaCalendario = function agregarNotaCalendario(fecha) {
    if (!fecha) return;
    const texto = text(prompt('Nota pública para ' + fecha, ''));
    if (!texto) return;
    pushUndoState('agregar nota de calendario');
    ensureCalendarNotes();
    APP.calendarNotes.push({
      id: 'calnote-' + Date.now() + '-' + Math.random().toString(16).slice(2),
      fecha,
      texto,
      author: getCalendarNoteAuthor(),
      createdAt: nowIso(),
      updatedAt: nowIso()
    });
    renderCalendario();
    scheduleAutoSave();
  };

  window.editarNotaCalendario = function editarNotaCalendario(id) {
    ensureCalendarNotes();
    const note = APP.calendarNotes.find(item => item.id === id);
    if (!note) return;
    const next = prompt('Editar nota pública. Deja vacío para eliminarla.', note.texto);
    if (next === null) return;
    pushUndoState('editar nota de calendario');
    const texto = text(next);
    if (!texto) APP.calendarNotes = APP.calendarNotes.filter(item => item.id !== id);
    else {
      note.texto = texto;
      note.author = getCalendarNoteAuthor();
      note.updatedAt = nowIso();
    }
    renderCalendario();
    scheduleAutoSave();
  };

  window.renderCalendario = function renderCalendarioV2() {
    const calContainer = document.getElementById('calContainer');
    const calMesNav = document.getElementById('calMesNav');
    const calMesLabel = document.getElementById('calMesActualLabel');
    if (!calContainer || !calMesNav || !calMesLabel) return;

    if (!APP.calMeses || !APP.calMeses.length) APP.calMeses = buildCalendarMonths(APP.calSemanas || []);
    if (!APP.calSemanas || !APP.calSemanas.length || !APP.calMeses.length) {
      calContainer.innerHTML = '<div class="empty-state"><div class="empty-icon">📅</div><p>Importa una planificación para comenzar.</p></div>';
      calMesNav.innerHTML = '';
      calMesLabel.textContent = '—';
      window.renderQueueSemana();
      return;
    }

    if (APP.calMesIdx >= APP.calMeses.length) APP.calMesIdx = Math.max(APP.calMeses.length - 1, 0);
    const currentMonth = APP.calMeses[APP.calMesIdx];
    APP.calSemanaIdx = currentMonth.firstSemanaIdx || 0;
    calMesLabel.textContent = currentMonth.label;
    calMesNav.innerHTML = '<span class="calendar-month-label">Mes:</span>' + APP.calMeses.map((month, index) =>
      `<button class="filter-btn month-filter ${index === APP.calMesIdx ? 'active' : ''}" onclick="APP.calMesIdx=${index};renderCalendario()">${month.label}</button>`
    ).join('');

    const search = text((document.getElementById('calSearch') || {}).value).toLowerCase();
    const selectedCount = (APP.selectedCalendarGroups || []).length;
    const html = [`
      <div class="calendar-topbar">
        <div style="font-size:13px;color:var(--muted);font-weight:700;">Mes visible: ${currentMonth.label}</div>
        <div class="calendar-selection-bar">
          <span id="calSelectionSummary" class="badge badge-warn">${selectedCount} seleccionados</span>
          <button id="moveSelectedBtn" class="btn btn-outline btn-sm" onclick="abrirMoveSelectedModal()" ${selectedCount ? '' : 'disabled'}>Mover seleccionados</button>
          <button id="clearSelectedBtn" class="btn btn-outline btn-sm" onclick="limpiarSeleccionCalendario()" ${selectedCount ? '' : 'disabled'}>Limpiar</button>
        </div>
      </div>
    `];

    currentMonth.semanas.forEach((semana, weekIndex) => {
      html.push(`<section class="cal-month-section">
        <div class="cal-week-title">Semana ${weekIndex + 1}<span>${semana.nombre}</span></div>
        <div class="cal-semana">`);
      semana.dias.forEach(fecha => {
        const isHol = isHoliday(fecha);
        const dayRoutes = APP.rutas[fecha] || {};
        html.push(`<div class="cal-dia-col">
          <div class="cal-dia-header" style="${isHol ? 'background:#fde8e8;color:var(--danger);' : ''}">
            <strong>${fechaLabel(fecha)}</strong><small>${fecha}${isHol ? ' · Feriado RD' : ''}</small>
          </div>
          ${renderCalendarNotesHtml(fecha)}`);

        const activeTrucks = getTruckOptions().filter(camion => (dayRoutes[camion] || []).length || search);
        if (!activeTrucks.length) activeTrucks.push('CAMION 1');
        activeTrucks.forEach(camion => {
          const truckClass = camion === 'CAMION 1' ? 'ch-c1' : camion === 'CAMION 2' ? 'ch-c2' : camion === 'CAMION 3' ? 'ch-alm' : 'ch-alm';
          const cardClass = camion === 'CAMION 1' ? 'cal-card-c1' : camion === 'CAMION 2' ? 'cal-card-c2' : 'cal-card-alm';
          const groups = (dayRoutes[camion] || []).filter(group => {
            if (!search) return true;
            return [group.nombre, group.codigo, group.pedidoCliente, group.zona]
              .some(value => text(value).toLowerCase().includes(search));
          });
          const load = getTruckLoadSummary(groups);
          const loadText = `Carga: ${formatLoadQty(load.total)} uds · Valor transportado: ${formatMonto(load.valorTransportado)}${load.solicitudes ? ` · Sol: ${formatLoadQty(load.solicitudes)} uds · Valor sol: ${formatMonto(load.montoSolicitudes)}` : ''}`;
          html.push(`<div class="cal-carril"
            ondragover="event.preventDefault();this.classList.add('drag-over')"
            ondragleave="this.classList.remove('drag-over')"
            ondrop="onDropCal(event,'${fecha}','${camion}',this)">
            <span class="cal-carril-header ${truckClass}">
              <span>${getTruckLabel(camion)} ${groups.length ? '(' + groups.length + ')' : ''}${load.rutas.length ? ` · ${load.rutas.join(', ')}` : ''}</span>
              <small>${loadText}</small>
            </span>`);

          groups.forEach(group => {
            const realIdx = (dayRoutes[camion] || []).indexOf(group);
            const status = group.cumplida ? '✅' : group.cantidadFacturada > 0 ? '🟡' : '🔄';
            const selected = isGroupSelected(group);
            const alertHtml = group.alertas.length
              ? `<div style="font-size:10px;color:var(--danger);margin-top:4px;">${group.alertas.join(' · ')}</div>`
              : '';
            const evidenceBits = [
              group.precintoDespacho ? 'Precinto: ' + group.precintoDespacho : '',
              group.choferRuta ? 'Chofer/recibido: ' + group.choferRuta : '',
              group.fotoCamion ? 'Foto adjunta' : ''
            ].filter(Boolean);
            const evidenceHtml = evidenceBits.length ? `<div class="route-evidence-line">${evidenceBits.join(' · ')}</div>` : '';
            html.push(`<div class="cal-card ${cardClass} ${selected ? 'cal-card-selected' : ''}" draggable="true"
              ondragstart="onDragStartCal(event,'${fecha}','${camion}',${realIdx})"
              ondragend="this.classList.remove('dragging')"
              style="${group.alertas.length ? 'border:1px solid #ef4444;' : ''}">
              <div class="cal-card-topline">
                <label class="cal-card-check">
                  <input type="checkbox" ${selected ? 'checked' : ''} onclick="toggleSeleccionCalendario('${fecha}','${camion}',${realIdx}, event)">
                  <span>Selec.</span>
                </label>
                <span class="badge ${selected ? 'badge-primary' : 'badge-outline'}">${selected ? 'Lote' : 'Individual'}</span>
              </div>
              <div class="cal-card-name">${group.nombre}</div>
              <div class="queue-card-meta">${group.codigo} · ${group.cantidadPedidos || 1} pedidos: ${group.pedidoCliente}</div>
              <div class="cal-card-monto">${group.items.length} líneas · ${group.cantidadSolicitudes || 0} solicitudes · Carga ${formatLoadQty(group.cargaTotal)} uds · Valor ${formatMonto(group.totalPendiente)}${group.montoSolicitudes ? ` · Valor sol ${formatMonto(group.montoSolicitudes)}` : ''}</div>
              ${evidenceHtml}
              ${alertHtml}
              <div class="cal-card-actions">
                <button class="btn btn-outline btn-sm" style="padding:2px 6px;font-size:10px;" onclick="abrirMoveModal('${fecha}','${camion}',${realIdx})">✏️ Mover</button>
                <button class="btn btn-outline btn-sm" style="padding:2px 6px;font-size:10px;" onclick="toggleCumplida('${fecha}','${camion}',${realIdx})">${group.cumplida ? '↩️' : '✅'}</button>
                <button class="btn btn-outline btn-sm" style="padding:2px 6px;font-size:10px;" onclick="abrirNotaModal('${fecha}','${camion}',${realIdx})">📝</button>
                <span>${status}</span>
              </div>
            </div>`);
          });

          if (!groups.length) html.push('<div style="font-size:11px;color:var(--muted);padding:8px 0;">Sin pedidos asignados.</div>');
          html.push('</div>');
        });
        html.push('</div>');
      });
      html.push('</div></section>');
    });

    calContainer.innerHTML = html.join('');
    syncCalendarToolbarState();
    window.renderQueueSemana();
  };

  window.navCalMes = function navCalMesV2(delta) {
    if (!APP.calMeses || !APP.calMeses.length) return;
    APP.calMesIdx = Math.max(0, Math.min(APP.calMeses.length - 1, (APP.calMesIdx || 0) + delta));
    renderCalendario();
  };


  function applyClientRouteToLineItems(codigo, routeName) {
    const code = text(codigo);
    const route = text(routeName);
    if (!code) return;
    [APP.lineItems, APP.planLineItems, APP.controlLineItems].forEach(list => {
      (list || []).forEach(item => {
        if (item.clienteId !== code) return;
        item.rutaNombre = route;
        item.zona = route;
        item.fechaActualizacion = nowIso();
      });
    });
  }

  function setClienteRutaAsignada(codigo, nombre, routeName) {
    const code = text(codigo);
    const route = text(routeName);
    if (!code) return null;
    const cfg = getClienteConfigV2(code, nombre);
    cfg.codigoCliente = code;
    cfg.nombreCliente = text(nombre) || cfg.nombreCliente || code;
    cfg.rutaAsignada = route;
    cfg.configuracionCompleta = !!(cfg.rutaAsignada && (cfg.diasRecepcion || []).length && text(cfg.horarioAlmacen));
    APP.clienteConfig[code] = cfg;
    if (route) ensureRouteConfig(route);
    applyClientRouteToLineItems(code, route);
    return cfg;
  }

  function getRouteCatalogClients() {
    const map = new Map();
    const add = (codigo, nombre) => {
      const code = text(codigo);
      if (!code) return;
      const current = map.get(code) || { codigo: code, nombre: text(nombre) || code };
      if (!current.nombre || current.nombre === code) current.nombre = text(nombre) || current.nombre;
      map.set(code, current);
    };
    Object.values(APP.clienteConfig || {}).forEach(cfg => add(cfg.codigoCliente || cfg.codigo, cfg.nombreCliente));
    [...(APP.lineItems || []), ...(APP.planLineItems || []), ...(APP.controlLineItems || [])].forEach(item => add(item.clienteId, item.clienteNombre));
    [...(APP.solicitudesAlmacen || []), ...(APP.solicitudesPlanAlmacen || []), ...(APP.solicitudesControlAlmacen || [])].forEach(item => add(item.codigo, item.cliente));
    return [...map.values()].map(client => {
      const cfg = getClienteConfigV2(client.codigo, client.nombre);
      const route = text(cfg.rutaAsignada);
      const complete = !!(cfg.configuracionCompleta && cfg.rutaAsignada && (cfg.diasRecepcion || []).length && text(cfg.horarioAlmacen));
      const orders = (APP.lineItems || []).filter(item => item.clienteId === client.codigo);
      return {
        ...client,
        ruta: route || 'Pendiente de configurar',
        cfg,
        complete,
        orders,
        monto: orders.reduce((sum, item) => sum + getOperationalAmount(item), 0),
        lineas: orders.length
      };
    }).sort((a, b) => text(a.ruta).localeCompare(text(b.ruta), 'es') || text(a.nombre).localeCompare(text(b.nombre), 'es'));
  }

  function getRouteAssignmentOptions() {
    const names = uniqueTexts([
      ...(APP.routeConfigs || []).map(route => route.nombre || route.zona),
      ...Object.values(APP.clienteConfig || {}).map(cfg => cfg.rutaAsignada),
      ...(APP.lineItems || []).map(item => item.rutaNombre || item.zona),
      ...(APP.planLineItems || []).map(item => item.rutaNombre || item.zona),
      ...(APP.controlLineItems || []).map(item => item.rutaNombre || item.zona)
    ]).filter(route => route && route !== 'Pendiente de configurar');
    return names.sort((a, b) => text(a).localeCompare(text(b), 'es'));
  }

  function routeAssignmentSelectHtml(client) {
    const options = getRouteAssignmentOptions();
    const current = text(client && client.cfg && client.cfg.rutaAsignada);
    const optionHtml = options.map(route => `<option value="${jsString(route)}" ${route === current ? 'selected' : ''}>${route}</option>`).join('');
    return `<select class="route-assign-select" onclick="event.stopPropagation()" onchange="asignarClienteRutaRapida('${jsString(client.codigo)}', this.value)">
      <option value="">Sin ruta</option>
      ${optionHtml}
    </select>`;
  }

  function routeCatalogGroups() {
    const groups = {};
    const routes = uniqueTexts([...(APP.routeConfigs || []).map(route => route.nombre || route.zona), 'Pendiente de configurar']);
    routes.forEach(route => { groups[route || 'Pendiente de configurar'] = []; });
    getRouteCatalogClients().forEach(client => {
      const key = client.cfg.rutaAsignada ? client.cfg.rutaAsignada : 'Pendiente de configurar';
      if (!groups[key]) groups[key] = [];
      groups[key].push(client);
    });
    return groups;
  }

  window.renderRouteCatalog = function renderRouteCatalog() {
    const cont = document.getElementById('routeCatalogContainer');
    if (!cont) return;
    const summary = document.getElementById('routeCatalogSummary');
    const q = text((document.getElementById('routeCatalogSearch') || {}).value).toLowerCase();
    const groups = routeCatalogGroups();
    const routeNames = Object.keys(groups).sort((a, b) => (a === 'Pendiente de configurar' ? -1 : b === 'Pendiente de configurar' ? 1 : a.localeCompare(b, 'es')));
    const total = routeNames.reduce((sum, route) => sum + groups[route].length, 0);
    const pending = (groups['Pendiente de configurar'] || []).length;
    if (summary) summary.textContent = `${total} clientes · ${pending} pendientes`;
    if (!total) {
      cont.innerHTML = '<div class="empty-state"><div class="empty-icon">🧭</div><p>Carga una planificación o configuración para ver clientes por ruta.</p></div>';
      return;
    }
    cont.innerHTML = `<div class="route-catalog-board">${routeNames.map(routeName => {
      const allClients = groups[routeName] || [];
      const clients = allClients.filter(client => {
        if (!q) return true;
        return [client.codigo, client.nombre, client.ruta, client.cfg.observaciones, client.cfg.condicionesEspeciales].some(value => text(value).toLowerCase().includes(q));
      });
      const isPending = routeName === 'Pendiente de configurar';
      const countLabel = q && clients.length !== allClients.length ? `${clients.length}/${allClients.length} clientes` : `${allClients.length} clientes`;
      return `<section class="route-catalog-col ${isPending ? 'pending' : ''}" ondragover="event.preventDefault();this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="moverClienteCatalogoRuta(event,'${jsString(routeName)}')">
        <div class="route-catalog-head"><strong>${routeName}</strong><span>${countLabel}</span></div>
        <div class="route-catalog-body">
          ${clients.length ? clients.map(client => `<article class="route-client-card ${client.complete ? '' : 'pending'}" draggable="true" ondragstart="onDragStartRouteCatalog(event,'${jsString(client.codigo)}')" onclick="abrirDetalleRutaCliente('${jsString(client.codigo)}')">
            <span><strong>${client.nombre}</strong><small>${client.codigo} · ${client.lineas} líneas · ${formatMonto(client.monto)}</small></span>
            <div class="route-card-actions" onclick="event.stopPropagation()">
              <em>${client.complete ? 'Config.' : 'Pendiente'}</em>
              ${routeAssignmentSelectHtml(client)}
            </div>
          </article>`).join('') : '<div class="route-catalog-empty">Sin clientes.</div>'}
        </div>
      </section>`;
    }).join('')}</div>`;
  };

  function commitClienteRutaAssignment(codigo, routeName, label) {
    const route = text(routeName);
    const client = getRouteCatalogClients().find(item => item.codigo === text(codigo));
    if (!client) return;
    pushUndoState(label || 'asignar cliente a ruta');
    setClienteRutaAsignada(client.codigo, client.nombre, route);
    rebuildDerivedState();
    renderRouteCatalog();
    renderConfigClientes();
    renderCalendario();
    renderRutas();
    renderComercial();
    scheduleAutoSave();
  }

  window.onDragStartRouteCatalog = function onDragStartRouteCatalog(event, codigo) {
    APP.routeCatalogDragCode = text(codigo);
    if (event && event.dataTransfer) {
      event.dataTransfer.setData('text/plain', APP.routeCatalogDragCode);
      event.dataTransfer.effectAllowed = 'move';
    }
  };

  window.moverClienteCatalogoRuta = function moverClienteCatalogoRuta(event, routeName) {
    if (event && event.preventDefault) event.preventDefault();
    const codigo = (event && event.dataTransfer && event.dataTransfer.getData('text/plain')) || APP.routeCatalogDragCode;
    APP.routeCatalogDragCode = '';
    if (!codigo || routeName === 'Pendiente de configurar') return;
    commitClienteRutaAssignment(codigo, routeName, 'mover cliente de ruta');
  };

  window.asignarClienteRutaRapida = function asignarClienteRutaRapida(codigo, routeName) {
    commitClienteRutaAssignment(codigo, routeName, 'asignar ruta desde tablero');
  };

  window.abrirDetalleRutaCliente = function abrirDetalleRutaCliente(codigo) {
    const client = getRouteCatalogClients().find(item => item.codigo === codigo);
    if (!client) return;
    const cfg = client.cfg || {};
    const lines = [
      `Cliente: ${client.nombre}`,
      `Código: ${client.codigo}`,
      `Ruta asignada: ${cfg.rutaAsignada || 'Pendiente de configurar'}`,
      `Días recepción: ${(cfg.diasRecepcion || []).join(', ') || 'No definidos'}`,
      `Horario almacén: ${cfg.horarioAlmacen || 'No definido'}`,
      `Contacto: ${cfg.contacto || 'No definido'}`,
      `Observaciones: ${cfg.observaciones || 'Sin observaciones'}`,
      `Condiciones especiales: ${cfg.condicionesEspeciales || 'Sin condiciones especiales'}`
    ];
    alert(lines.join('\n'));
  };

  window.exportarRouteCatalogExcel = function exportarRouteCatalogExcel() {
    const rows = getRouteCatalogClients().map(client => ({
      'Ruta': client.cfg.rutaAsignada || 'Pendiente de configurar',
      'Código': client.codigo,
      'Cliente': client.nombre,
      'Configuración completa': client.complete ? 'Sí' : 'No',
      'Días recepción': (client.cfg.diasRecepcion || []).join(', '),
      'Horario almacén': client.cfg.horarioAlmacen || '',
      'Contacto': client.cfg.contacto || '',
      'Observaciones': client.cfg.observaciones || '',
      'Condiciones especiales': client.cfg.condicionesEspeciales || ''
    }));
    if (!rows.length) return alert('No hay clientes para exportar.');
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Rutas clientes');
    XLSX.writeFile(wb, 'TMS_Rutas_Clientes.xlsx');
  };

  window.renderRutas = function renderRutasV2() {
    const sel = document.getElementById('rutasDiaSelect');
    const cont = document.getElementById('rutasContainer');
    if (!sel || !cont) return;
    const fechas = Object.keys(APP.rutas).sort((a, b) => fechaToDate(a) - fechaToDate(b));
    const prevVal = sel.value || APP.rutaFecha || '';
    sel.innerHTML = '<option value="">Seleccionar fecha...</option>' + fechas.map(fecha => `<option value="${fecha}" ${fecha === prevVal ? 'selected' : ''}>${fecha} — ${fechaLabel(fecha)}</option>`).join('');
    APP.rutaFecha = sel.value || APP.rutaFecha || fechas[0] || '';
    if (!APP.rutaFecha || !APP.rutas[APP.rutaFecha]) {
      cont.innerHTML = '<div class="empty-state"><div class="empty-icon">🚚</div><p>Selecciona una fecha para ver las rutas.</p></div>';
      return;
    }

    const search = text((document.getElementById('rutasSearch') || {}).value).toLowerCase();
    const rutasDia = APP.rutas[APP.rutaFecha];
    cont.innerHTML = `<div class="rutas-grid">${getTruckOptions().map(camion => {
      const groups = (rutasDia[camion] || []).filter(group => {
        if (!search) return true;
        return [group.nombre, group.codigo, group.pedidoCliente, group.zona].some(value => text(value).toLowerCase().includes(search));
      });
      return `<div class="ruta-col">
        <div class="ruta-col-header ${camion === 'CAMION 1' ? 'ch-c1' : camion === 'CAMION 2' ? 'ch-c2' : 'ch-alm'}">${getTruckLabel(camion)} (${groups.length})</div>
        <div class="ruta-col-body">
          ${groups.length ? groups.map((group, idx) => `<div class="ruta-item" style="${group.alertas.length ? 'border-left:3px solid #ef4444;' : ''}">
            <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;">
              <div>
                <div class="ruta-item-name">${idx + 1}. ${group.nombre}</div>
                <div class="ruta-item-detail">${group.codigo} · ${group.cantidadPedidos || 1} pedidos · ${group.zona}</div>
                <div class="ruta-item-detail">Pedidos: ${group.pedidoCliente}</div>
                <div class="ruta-item-detail">${group.items.length} líneas · ${group.cantidadSolicitudes || 0} solicitudes · ${group.cantidadPendiente} pendientes</div>
                ${group.alertas.length ? `<div style="font-size:11px;color:var(--danger);margin-top:4px;">${group.alertas.join(' · ')}</div>` : ''}
              </div>
              <div style="text-align:right;">
                <div class="ruta-item-monto">${group.cantidadSolicitada} uds.</div>
                <div class="badge ${getTruckBadgeClass(group.camion)}">${group.cumplida ? 'Cumplida' : 'Activa'}</div>
              </div>
            </div>
          </div>`).join('') : '<div class="empty-state" style="padding:20px 10px;"><p>Sin pedidos en esta ruta.</p></div>'}
        </div>
      </div>`;
    }).join('')}</div>`;
  };

  function getGroupByRouteRef(fecha, camion, idx) {
    return APP.rutas[fecha] && APP.rutas[fecha][camion] ? APP.rutas[fecha][camion][idx] : null;
  }

  window.asignarQueueACalendario = function asignarQueueACalendarioV2(idx, fecha, camion) {
    const queueItem = APP.queueSemana[idx];
    if (!queueItem) return;
    pushUndoState('asignar pedido al calendario');
    const item = { ...queueItem.item };
    item.fechaPlanificada = fecha;
    item.camionAsignado = camion;
    item.origen = 'manual';
    item.manualProgramado = true;
    item.fechaActualizacion = nowIso();
    APP.lineItems.push(item);
    APP.queueSemana.splice(idx, 1);
    rebuildDerivedState();
    actualizarDashboard();
    renderCalendario();
    renderRutas();
    renderComercial();
    renderAlmacen();
    scheduleAutoSave();
  };

  window.asignarRutaQueueACalendario = function asignarRutaQueueACalendarioV2(routeName, fecha, camion) {
    const selected = APP.queueSemana.filter(item => queueByRouteName(item) === routeName);
    if (!selected.length) return;
    pushUndoState('asignar ruta al calendario');
    const selectedIds = new Set(selected.map(item => item.queueId + '|' + item.fechaControl));
    selected.forEach(queueItem => {
      const item = { ...queueItem.item };
      item.fechaPlanificada = fecha;
      item.camionAsignado = camion;
      item.rutaNombre = routeName;
      item.zona = routeName;
      item.origen = 'manual';
      item.manualProgramado = true;
      item.fechaActualizacion = nowIso();
      APP.lineItems.push(item);
    });
    APP.queueSemana = APP.queueSemana.filter(item => !selectedIds.has(item.queueId + '|' + item.fechaControl));
    rebuildDerivedState();
    actualizarDashboard();
    renderCalendario();
    renderRutas();
    renderComercial();
    renderAlmacen();
    scheduleAutoSave();
  };

  window.moverCliente = function moverClienteV2(fecha, camion, idx, nuevaFecha, nuevoCamion) {
    const group = getGroupByRouteRef(fecha, camion, idx);
    if (!group) return;
    pushUndoState('mover pedido de ruta');
    const selectedKeys = new Set(APP.selectedCalendarGroups || []);
    const groupKey = getGroupSelectionKey(group);
    const multiMove = selectedKeys.has(groupKey) && selectedKeys.size > 1;
    const keys = new Set();
    if (multiMove) {
      getSelectedGroups().forEach(({ group: selectedGroup }) => {
        selectedGroup.items.forEach(item => keys.add(item.baseKey));
      });
    } else {
      group.items.forEach(item => keys.add(item.baseKey));
    }
    APP.lineItems.forEach(item => {
      if (keys.has(item.baseKey)) {
        item.fechaPlanificada = nuevaFecha;
        item.camionAsignado = nuevoCamion;
        item.manualProgramado = true;
        item.origen = item.origen === 'planificacion semanal' ? 'manual' : item.origen;
        item.fechaActualizacion = nowIso();
      }
    });
    rebuildDerivedState();
    clearCalendarSelection();
    renderCalendario();
    renderRutas();
    renderComercial();
    renderAlmacen();
    scheduleAutoSave();
  };

  window.toggleCumplida = function toggleCumplidaV2(fecha, camion, idx) {
    const group = getGroupByRouteRef(fecha, camion, idx);
    if (!group) return;
    if (!group.cumplida) {
      window.abrirNotaModal(fecha, camion, idx, { completeOnSave: true });
      return;
    }
    pushUndoState('deshacer ruta cumplida');
    const newQty = group.cumplida ? 0 : null;
    const keys = new Set(group.items.map(item => item.baseKey));
    APP.lineItems.forEach(item => {
      if (!keys.has(item.baseKey)) return;
      item.cantidadFacturada = newQty === null ? (item.cantidadSolicitada || item.cantidadFacturada || 0) : 0;
      item.cantidadPendiente = Math.max((item.cantidadSolicitada || 0) - (item.cantidadFacturada || 0), 0);
      item.fechaActualizacion = nowIso();
    });
    rebuildDerivedState();
    actualizarDashboard();
    renderCalendario();
    renderRutas();
    renderComercial();
    scheduleAutoSave();
  };

  window.abrirMoveModal = function abrirMoveModalV2(fecha, camion, idx) {
    const group = getGroupByRouteRef(fecha, camion, idx);
    if (!group) return;
    APP.moveCliente = { fecha, cam: camion, idx };
    document.getElementById('moveModalClienteNombre').textContent = `${group.nombre} · Pedido ${group.pedidoCliente}`;
    const dates = buildMoveDateOptions(group.items[0], fecha);
    setMoveDateInput(dates, fecha);
    syncMoveOptions();
    document.getElementById('moveCamion').value = camion;
    document.getElementById('moveModal').classList.add('show');
  };

  window.abrirMoveSelectedModal = function abrirMoveSelectedModalV2() {
    const selectedGroups = getSelectedGroups();
    if (!selectedGroups.length) return alert('Selecciona al menos un pedido del calendario.');
    const [{ fecha, camion, group }] = selectedGroups;
    APP.moveCliente = {
      selectedKeys: [...new Set(APP.selectedCalendarGroups || [])],
      fecha,
      cam: camion,
      idx: (APP.rutas[fecha] && APP.rutas[fecha][camion] ? APP.rutas[fecha][camion].indexOf(group) : 0)
    };
    document.getElementById('moveModalClienteNombre').textContent = `${selectedGroups.length} pedidos seleccionados`;
    const dates = buildMoveDateOptions(group.items[0], fecha);
    setMoveDateInput(dates, fecha);
    syncMoveOptions();
    document.getElementById('moveCamion').value = camion;
    document.getElementById('moveModal').classList.add('show');
  };

  window.abrirMoveRouteQueueModal = function abrirMoveRouteQueueModalV2(routeName) {
    APP.moveCliente = { queueRouteName: routeName };
    const routeItems = APP.queueSemana.filter(item => queueByRouteName(item) === routeName);
    const clientCount = new Set(routeItems.map(item => item.codigo)).size;
    document.getElementById('moveModalClienteNombre').textContent = `${routeName} · ${clientCount} clientes · ${routeItems.length} líneas`;
    const sample = routeItems[0] && routeItems[0].item;
    const dates = buildMoveDateOptions(sample || {}, APP.planFecha || fechaToStr(new Date()));
    setMoveDateInput(dates, dates[0]);
    syncMoveOptions();
    document.getElementById('moveCamion').value = getTruckOptions()[0];
    document.getElementById('moveModal').classList.add('show');
  };

  window.abrirMoveQueueModal = function abrirMoveQueueModalV2(idx) {
    const item = APP.queueSemana[idx];
    if (!item) return;
    APP.moveCliente = { queueIdx: idx };
    document.getElementById('moveModalClienteNombre').textContent = `${item.nombre} · Pedido ${item.pedidoCliente}`;
    const dates = buildMoveDateOptions(item.item, APP.planFecha || fechaToStr(new Date()));
    setMoveDateInput(dates, dates[0]);
    syncMoveOptions();
    document.getElementById('moveCamion').value = item.camion || getTruckOptions()[0];
    document.getElementById('moveModal').classList.add('show');
  };

  window.toggleSeleccionCalendario = function toggleSeleccionCalendarioV2(fecha, camion, idx, event) {
    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }
    const group = getGroupByRouteRef(fecha, camion, idx);
    if (!group) return;
    toggleCalendarSelection(group);
    renderCalendario();
  };

  window.limpiarSeleccionCalendario = function limpiarSeleccionCalendarioV2() {
    clearCalendarSelection();
    renderCalendario();
  };

  window.cerrarMoveModal = function cerrarMoveModalV2() {
    const modal = document.getElementById('moveModal');
    if (modal) modal.classList.remove('show');
    APP.moveCliente = null;
  };

  window.confirmarMover = function confirmarMoverV2() {
    if (!APP.moveCliente) return;
    const nuevaFecha = getMoveDateValue();
    if (!nuevaFecha) return alert('Selecciona una fecha para mover el cliente.');
    if (!isManualRouteDate(nuevaFecha)) return alert('Para movimientos manuales puedes usar lunes a sábado. Domingos y feriados no están disponibles.');
    const nuevoCamion = document.getElementById('moveCamion').value;
    const moveState = APP.moveCliente;
    window.cerrarMoveModal();
    if (moveState.queueRouteName) {
      asignarRutaQueueACalendario(moveState.queueRouteName, nuevaFecha, nuevoCamion);
      return;
    }
    if (moveState.queueIdx != null) {
      asignarQueueACalendario(moveState.queueIdx, nuevaFecha, nuevoCamion);
      return;
    }
    moverCliente(moveState.fecha, moveState.cam, moveState.idx, nuevaFecha, nuevoCamion);
  };

  function loadScriptOnce(src) {
    return new Promise((resolve, reject) => {
      const existing = document.querySelector(`script[src="${src}"]`);
      if (existing) {
        if (existing.dataset.loaded === 'true') resolve();
        else {
          existing.addEventListener('load', resolve, { once: true });
          existing.addEventListener('error', reject, { once: true });
        }
        return;
      }
      const script = document.createElement('script');
      script.src = src;
      script.async = true;
      script.onload = () => {
        script.dataset.loaded = 'true';
        resolve();
      };
      script.onerror = () => reject(new Error('No se pudo cargar ' + src));
      document.head.appendChild(script);
    });
  }

  async function ensureHtml2Pdf() {
    if (typeof window.html2pdf !== 'undefined') return window.html2pdf;
    const sources = [
      'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js',
      'https://cdn.jsdelivr.net/npm/html2pdf.js@0.10.1/dist/html2pdf.bundle.min.js'
    ];
    for (const src of sources) {
      try {
        await loadScriptOnce(src);
        if (typeof window.html2pdf !== 'undefined') return window.html2pdf;
      } catch (error) {
        console.warn('No se pudo cargar html2pdf desde', src, error);
      }
    }
    throw new Error('La librería de PDF no se pudo cargar. Revisa la conexión e intenta de nuevo.');
  }

  async function ensurePdfTools() {
    if (typeof window.html2canvas === 'undefined') {
      const canvasSources = [
        'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
        'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js'
      ];
      for (const src of canvasSources) {
        try {
          await loadScriptOnce(src);
          if (typeof window.html2canvas !== 'undefined') break;
        } catch (error) {
          console.warn('No se pudo cargar html2canvas desde', src, error);
        }
      }
    }
    let JsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
    if (!JsPDF) {
      const pdfSources = [
        'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
        'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js'
      ];
      for (const src of pdfSources) {
        try {
          await loadScriptOnce(src);
          JsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
          if (JsPDF) break;
        } catch (error) {
          console.warn('No se pudo cargar jsPDF desde', src, error);
        }
      }
    }
    if (typeof window.html2canvas === 'undefined' || !JsPDF) {
      throw new Error('La librería de PDF no se pudo cargar. Revisa la conexión e intenta de nuevo.');
    }
    return { html2canvas: window.html2canvas, JsPDF };
  }

  function canvasHasVisibleContent(canvas) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    if (!ctx || !canvas.width || !canvas.height) return false;
    const stepX = Math.max(1, Math.floor(canvas.width / 80));
    const stepY = Math.max(1, Math.floor(canvas.height / 80));
    for (let y = 0; y < canvas.height; y += stepY) {
      for (let x = 0; x < canvas.width; x += stepX) {
        const data = ctx.getImageData(x, y, 1, 1).data;
        if (data[3] > 0 && (data[0] < 245 || data[1] < 245 || data[2] < 245)) return true;
      }
    }
    return false;
  }

  function downloadPdf(pdf, filename) {
    try {
      const blob = pdf.output('blob');
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      link.style.display = 'none';
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
    } catch (error) {
      console.warn('Descarga PDF por Blob falló; usando save().', error);
      pdf.save(filename);
    }
  }

  function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error || new Error('No se pudo leer la imagen.'));
      reader.readAsDataURL(file);
    });
  }

  function loadImageFromDataUrl(dataUrl) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('No se pudo procesar la imagen.'));
      img.src = dataUrl;
    });
  }

  async function compressTruckPhoto(file) {
    const raw = await readFileAsDataUrl(file);
    const img = await loadImageFromDataUrl(raw);
    const maxSide = 1200;
    const scale = Math.min(1, maxSide / Math.max(img.width || maxSide, img.height || maxSide));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round((img.width || maxSide) * scale));
    canvas.height = Math.max(1, Math.round((img.height || maxSide) * scale));
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    return {
      dataUrl: canvas.toDataURL('image/jpeg', 0.76),
      name: file.name || 'foto-camion.jpg'
    };
  }

  function setTruckPhotoPreview(photo) {
    const wrap = document.getElementById('fotoCamionPreviewWrap');
    const img = document.getElementById('fotoCamionPreview');
    const name = document.getElementById('fotoCamionNombre');
    if (!wrap || !img) return;
    if (photo && photo.dataUrl) {
      wrap.style.display = 'block';
      img.src = photo.dataUrl;
      if (name) name.textContent = photo.name || 'Foto adjunta';
    } else {
      wrap.style.display = 'none';
      img.removeAttribute('src');
      if (name) name.textContent = '';
    }
  }

  window.previsualizarFotoCamion = async function previsualizarFotoCamion(file) {
    if (!file) {
      APP.pendingTruckPhoto = null;
      setTruckPhotoPreview(null);
      return;
    }
    if (!/^image\//.test(file.type || '')) {
      alert('Selecciona un archivo de imagen.');
      return;
    }
    try {
      APP.pendingTruckPhoto = await compressTruckPhoto(file);
      setTruckPhotoPreview(APP.pendingTruckPhoto);
    } catch (error) {
      console.error('Error procesando foto:', error);
      alert('No se pudo preparar la foto del camión.');
    }
  };

  window.exportarCalendarioVisiblePDF = async function exportarCalendarioVisiblePDFV2() {
    const el = document.getElementById('calExportArea');
    if (!el || !el.innerHTML.trim()) {
      alert('No hay calendario visible para exportar.');
      return;
    }
    let pdfTools;
    try {
      pdfTools = await ensurePdfTools();
    } catch (error) {
      console.error('Error cargando PDF:', error);
      alert(error.message || 'La librería de PDF no está disponible en este momento.');
      return;
    }
    const month = APP.calMeses && APP.calMeses[APP.calMesIdx];
    const filename = 'RouteControl_Calendario_' + (month ? month.key : 'visible') + '.pdf';
    const wrapper = document.createElement('div');
    wrapper.className = 'calendar-pdf-stage';
    const clone = el.cloneNode(true);
    clone.classList.add('calendar-pdf-export');
    clone.querySelectorAll('.cal-card-topline, .cal-card-actions, .calendar-selection-bar, button, input').forEach(node => node.remove());
    wrapper.innerHTML = `
      <div class="calendar-pdf-title">
        <strong>Calendario de Entregas</strong>
        <span>${month ? month.label : 'Vista actual'}</span>
      </div>
    `;
    wrapper.appendChild(clone);
    document.body.appendChild(wrapper);
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    if (!wrapper.scrollWidth || !wrapper.scrollHeight || !wrapper.textContent.trim()) {
      wrapper.remove();
      alert('No se pudo preparar el contenido del calendario para PDF.');
      return;
    }
    try {
      const canvas = await pdfTools.html2canvas(wrapper, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        width: wrapper.scrollWidth,
        height: wrapper.scrollHeight,
        windowWidth: wrapper.scrollWidth,
        windowHeight: wrapper.scrollHeight,
        scrollX: 0,
        scrollY: 0
      });
      if (!canvasHasVisibleContent(canvas)) {
        throw new Error('La captura del calendario salió vacía.');
      }
      const pdf = new pdfTools.JsPDF({ unit: 'mm', format: 'a4', orientation: 'landscape' });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 4;
      const maxW = pageW - margin * 2;
      const maxH = pageH - margin * 2;
      const ratio = canvas.width / canvas.height;
      let imgW = maxW;
      let imgH = imgW / ratio;
      if (imgH > maxH) {
        imgH = maxH;
        imgW = imgH * ratio;
      }
      const x = (pageW - imgW) / 2;
      const y = (pageH - imgH) / 2;
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.95), 'JPEG', x, y, imgW, imgH);
      downloadPdf(pdf, filename);
    } catch (error) {
      console.error('Error exportando calendario:', error);
      alert('No se pudo exportar el calendario: ' + error.message);
    } finally {
      wrapper.remove();
    }
  };

  window.abrirNotaModal = function abrirNotaModalV2(fecha, camion, idx, options) {
    const group = getGroupByRouteRef(fecha, camion, idx);
    if (!group) return;
    APP.moveCliente = { fecha, cam: camion, idx, completeOnSave: !!(options && options.completeOnSave) };
    APP.pendingTruckPhoto = null;
    document.getElementById('notaModalNombre').textContent = `${group.nombre} · Pedido ${group.pedidoCliente}`;
    document.getElementById('notaInput').value = group.nota || '';
    document.getElementById('incidenciaInput').value = group.incidencia || '';
    const precintoInput = document.getElementById('precintoInput');
    const comentarioInput = document.getElementById('comentarioRutaInput');
    const choferInput = document.getElementById('choferRutaInput');
    const fotoInput = document.getElementById('fotoCamionInput');
    if (precintoInput) precintoInput.value = group.precintoDespacho || '';
    if (choferInput) choferInput.value = group.choferRuta || '';
    if (comentarioInput) comentarioInput.value = group.comentarioRuta || '';
    if (fotoInput) fotoInput.value = '';
    setTruckPhotoPreview(group.fotoCamion ? { dataUrl: group.fotoCamion, name: group.fotoCamionNombre } : null);
    document.getElementById('notaModal').classList.add('show');
  };

  window.guardarNota = function guardarNotaV2() {
    if (!APP.moveCliente) return;
    const group = getGroupByRouteRef(APP.moveCliente.fecha, APP.moveCliente.cam, APP.moveCliente.idx);
    if (!group) return;
    pushUndoState(APP.moveCliente.completeOnSave ? 'completar ruta con evidencia' : 'guardar nota de ruta');
    const nota = text(document.getElementById('notaInput').value);
    const incidencia = text(document.getElementById('incidenciaInput').value);
    const precinto = text((document.getElementById('precintoInput') || {}).value);
    const comentarioRuta = text((document.getElementById('comentarioRutaInput') || {}).value);
    const choferRuta = text((document.getElementById('choferRutaInput') || {}).value);
    const currentPhoto = APP.pendingTruckPhoto || (group.fotoCamion ? { dataUrl: group.fotoCamion, name: group.fotoCamionNombre } : null);
    const keys = new Set(group.items.map(item => item.baseKey));
    let photoStored = false;
    APP.lineItems.forEach(item => {
      if (!keys.has(item.baseKey)) return;
      item.observaciones = nota;
      item.incidencia = incidencia;
      item.precintoDespacho = precinto;
      item.comentarioRuta = comentarioRuta;
      item.choferRuta = choferRuta;
      item.fotoCamion = currentPhoto && !photoStored ? currentPhoto.dataUrl : '';
      item.fotoCamionNombre = currentPhoto ? currentPhoto.name : '';
      if (currentPhoto && !photoStored) photoStored = true;
      if (APP.moveCliente.completeOnSave) {
        item.cantidadFacturada = item.cantidadSolicitada || item.cantidadFacturada || 0;
        item.cantidadPendiente = Math.max((item.cantidadSolicitada || 0) - (item.cantidadFacturada || 0), 0);
        item.fechaCierreRuta = nowIso();
      }
      item.fechaActualizacion = nowIso();
    });
    document.getElementById('notaModal').classList.remove('show');
    APP.moveCliente = null;
    APP.pendingTruckPhoto = null;
    rebuildDerivedState();
    actualizarDashboard();
    renderCalendario();
    renderRutas();
    renderComercial();
    renderAlmacen();
    scheduleAutoSave();
  };

  function summarizeImport(mode, items, fileName, fecha) {
    const label = mode === 'plan' ? 'Planificación' : 'Control';
    const uniqueClients = new Set(items.map(item => item.clienteId)).size;
    const uniqueOrders = new Set(items.map(item => [item.clienteId, item.pedidoCliente].join('|'))).size;
    return `${label} cargada: ${fileName} · ${uniqueClients} clientes · ${uniqueOrders} pedidos · ${fecha}`;
  }

  function registerRouteCostSnapshot(controlDate) {
    const rows = buildRouteCostRowsFromSchedule((APP.lineItems || []).filter(item => item.fechaPlanificada === controlDate));
    const merged = [...APP.routeCostHistory.filter(row => row.fecha !== controlDate), ...rows];
    APP.routeCostHistory = dedupeBy(merged, row => [row.fecha, row.rutaNombre].join('|'));
  }


  function registerControlHistorySnapshot(controlDate, fileName) {
    const fecha = controlDate || APP.controlFecha || fechaToStr(new Date());
    const items = APP.controlLineItems || [];
    if (!fecha || !items.length) return;
    const row = {
      fecha,
      archivo: fileName || APP.importFiles.control || '',
      lineas: items.length,
      clientes: new Set(items.map(item => item.clienteId).filter(Boolean)).size,
      pedidos: new Set(items.map(item => [item.clienteId, item.pedidoCliente].join('|')).filter(Boolean)).size,
      cantidadSolicitada: items.reduce((sum, item) => sum + num(item.cantidadSolicitada), 0),
      cantidadFacturada: items.reduce((sum, item) => sum + num(item.cantidadFacturada), 0),
      cantidadPendiente: items.reduce((sum, item) => sum + num(item.cantidadPendiente), 0),
      confirmadoNoEntregado: items.reduce((sum, item) => sum + getConfirmedUndeliveredAmount(item), 0),
      entregadoNoFacturado: items.reduce((sum, item) => sum + getDeliveredNotInvoicedAmount(item), 0),
      montoOperativo: items.reduce((sum, item) => sum + getOperationalAmount(item), 0),
      creadoEn: nowIso()
    };
    APP.controlHistory = [
      ...(APP.controlHistory || []).filter(item => item.fecha !== fecha),
      row
    ].sort((a, b) => fechaToDate(a.fecha) - fechaToDate(b.fecha));
  }

  function initControlHistoryPanel() {
    const grid = document.querySelector('#dashboard .dash-charts-grid');
    if (!grid || document.getElementById('controlHistoryPanel')) return;
    grid.insertAdjacentHTML('afterend', `
      <div class="chart-card control-history-panel" id="controlHistoryPanel">
        <div class="control-history-head">
          <div>
            <div class="chart-card-title">Histórico del control diario</div>
          </div>
          <button class="btn btn-outline btn-sm" onclick="exportarControlHistoryExcel()">Excel</button>
        </div>
        <div class="chart-canvas-wrap" style="height:220px;"><canvas id="chartControlHistory"></canvas></div>
        <div id="controlHistoryTable"></div>
      </div>
    `);
  }

  function renderControlHistoryPanel() {
    initControlHistoryPanel();
    const table = document.getElementById('controlHistoryTable');
    const canvas = document.getElementById('chartControlHistory');
    const rows = [...(APP.controlHistory || [])].sort((a, b) => fechaToDate(a.fecha) - fechaToDate(b.fecha));
    if (table) {
      table.innerHTML = rows.length ? `<div class="table-wrap"><table class="data-table control-history-table">
        <thead><tr><th>Fecha</th><th>Clientes</th><th>Líneas</th><th>Conf. no entr.</th><th>Ent. no fact.</th><th>Total</th></tr></thead>
        <tbody>${rows.slice(-8).reverse().map(row => `<tr>
          <td>${row.fecha}</td><td>${row.clientes}</td><td>${row.lineas}</td><td>${formatMonto(row.confirmadoNoEntregado)}</td><td>${formatMonto(row.entregadoNoFacturado)}</td><td><strong>${formatMonto(row.montoOperativo)}</strong></td>
        </tr>`).join('')}</tbody>
      </table></div>` : '<div class="empty-state" style="padding:18px;"><p>Sube controles diarios para ver la evolución histórica.</p></div>';
    }
    if (!canvas || typeof Chart === 'undefined') return;
    const labels = rows.map(row => row.fecha);
    const datasets = [
      { label: 'Total operativo', data: rows.map(row => num(row.montoOperativo)), borderColor: '#111827', backgroundColor: 'rgba(17,24,39,.08)', tension: .28, fill: true },
      { label: 'Conf. no entregado', data: rows.map(row => num(row.confirmadoNoEntregado)), borderColor: '#00A676', backgroundColor: 'rgba(0,166,118,.08)', tension: .28 },
      { label: 'Ent. no facturado', data: rows.map(row => num(row.entregadoNoFacturado)), borderColor: '#F5C542', backgroundColor: 'rgba(245,197,66,.12)', tension: .28 }
    ];
    if (APP.controlHistoryChart) APP.controlHistoryChart.destroy();
    APP.controlHistoryChart = new Chart(canvas.getContext('2d'), {
      type: 'line',
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' } },
        scales: { y: { ticks: { callback: value => formatMonto(value).replace('RD$ ', '') } } }
      }
    });
  }

  window.exportarControlHistoryExcel = function exportarControlHistoryExcel() {
    const rows = APP.controlHistory || [];
    if (!rows.length) return alert('No hay histórico de control diario para exportar.');
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Histórico control');
    XLSX.writeFile(wb, 'TMS_Historico_Control_Diario.xlsx');
  };

  function matchWarehouseToOrders() {
    const warehouseItems = [
      ...(APP.solicitudesPlanAlmacen || []),
      ...(APP.solicitudesControlAlmacen || [])
    ];
    const byFullKey = new Map();
    const byLooseKey = new Map();
    warehouseItems.forEach(item => {
      const fullKey = buildSapMatchKey(item);
      const looseKey = [item.pedidoCliente, item.lineaPedidoCliente, item.articulo].map(normKey).join('|');
      if (fullKey) byFullKey.set(fullKey, item);
      if (looseKey) byLooseKey.set(looseKey, item);
    });
    [...(APP.lineItems || []), ...(APP.planLineItems || []), ...(APP.controlLineItems || [])].forEach(item => {
      const fullKey = buildSapMatchKey(item);
      const looseKey = [item.pedidoCliente, item.lineaPedidoCliente, item.articulo].map(normKey).join('|');
      const match = byFullKey.get(fullKey) || byLooseKey.get(looseKey);
      item.sapMatchKey = fullKey;
      if (!match) return;
      item.solicitudAlmacen = match.solicitud || '';
      item.cantidadAlmacenSolicitada = match.cantidadSolicitada || 0;
      item.cantidadAlmacenProcesada = match.cantidadCompletada || 0;
      item.cantidadAlmacenPendiente = match.cantidadPendiente || 0;
      item.estadoAlmacen = match.estado || '';
    });
  }

  function registerSolicitudesSnapshot(fileName) {
    const rows = APP.solicitudesAlmacen.map(item => ({
      fecha: item.fechaLiberacion || item.fechaEnvio || fechaToStr(new Date()),
      clienteId: item.codigo,
      clienteNombre: item.cliente,
      pedidoCliente: item.pedidoCliente || '',
      articulo: item.articulo || '',
      descripcionArticulo: item.descripcionArticulo || '',
      cantidadSolicitada: item.cantidadSolicitada || 0,
      cantidadSolicitadaOriginal: item.cantidadSolicitadaOriginal || item.cantidadSolicitada || 0,
      cantidadPendienteServir: item.cantidadPendienteServir || item.cantidadSolicitada || 0,
      cantidadPendienteSinStock: item.cantidadPendienteSinStock || 0,
      cantidadCompletada: item.cantidadCompletada || 0,
      cantidadPendiente: item.cantidadPendiente || 0,
      estado: item.estado || '',
      costeSolicitud: num(APP.warehouseSettings.costoSolicitud || 0),
      observaciones: item.observaciones || '',
      fuente: fileName || '',
      tipo: item.tipoSolicitudArchivo || ''
    }));
    const keys = new Set(rows.map(row => [row.fecha, row.clienteId, row.pedidoCliente, row.articulo].join('|')));
    APP.solicitudesHistory = [
      ...APP.solicitudesHistory.filter(row => !keys.has([row.fecha, row.clienteId, row.pedidoCliente, row.articulo].join('|'))),
      ...rows
    ];
  }

  window.procesarArchivo = function procesarArchivoV2(file, mode) {
    if (!file) return;
    mode = mode || 'plan';
    const isPlan = mode === 'plan';
    const uploadZone = document.getElementById(isPlan ? 'uploadZone' : 'ctrlUploadZone');
    const loadingEl = document.getElementById(isPlan ? 'sapLoading' : 'ctrlLoading');
    const progressEl = document.getElementById(isPlan ? 'sapProgressBar' : 'ctrlProgressBar');
    const msgEl = document.getElementById(isPlan ? 'sapLoadingMsg' : 'ctrlLoadingMsg');
    const bannerEl = document.getElementById(isPlan ? 'sapSuccessBanner' : 'ctrlSuccessBanner');
    const successMsgEl = document.getElementById(isPlan ? 'sapSuccessMsg' : 'ctrlSuccessMsg');
    const fileInput = document.getElementById(isPlan ? 'sapFile' : 'ctrlFile');

    if (fileInput) fileInput.value = '';
    if (!isPlan && !APP.planLoaded) {
      document.getElementById('ctrlSinPlan').style.display = 'block';
      return;
    }

    uploadZone.style.display = 'none';
    bannerEl.style.display = 'none';
    loadingEl.style.display = 'block';
    progressEl.style.width = '12%';
    msgEl.textContent = 'Leyendo archivo...';

    const reader = new FileReader();
    reader.onerror = function () {
      loadingEl.style.display = 'none';
      uploadZone.style.display = 'block';
      alert('No se pudo leer el archivo.');
    };

    reader.onload = function (event) {
      try {
        msgEl.textContent = 'Procesando estructura y columnas...';
        const parsed = parseSAPRows(event.target.result, file.name, mode);
        const items = parsed.items;
        if (!items.length) throw new Error('No se encontraron filas válidas para importar.');
        pushUndoState(isPlan ? 'importar planificación semanal' : 'importar control diario');

        if (isPlan) {
          APP.importFiles.plan = file.name;
          APP.planFecha = fechaInicioPlanificacion(parsed.fecha);
          APP.planLoaded = true;
          APP.controlLoaded = false;
          APP.controlFecha = null;
          APP.controlLineItems = [];
          APP.queueSemana = [];
          items.forEach(item => {
            item.fechaPlanificada = APP.planFecha;
            item.camionAsignado = chooseTruckForItem(item);
          });
          schedulePlanItems(items, APP.planFecha);
          APP.planLineItems = items.map(item => ({ ...item }));
          APP.lineItems = items.map(item => ({ ...item }));
        } else {
          APP.importFiles.control = file.name;
          APP.controlFecha = parsed.fecha;
          APP.controlLoaded = true;
          APP.controlLineItems = items.map(item => ({
            ...item,
            fechaControl: parsed.fecha,
            origen: 'control diario'
          }));
          APP.lineItems.forEach(item => {
            const match = APP.controlLineItems.find(controlItem => controlItem.baseKey === buildBaseItemKey(item));
            if (!match) return;
            item.cantidadFacturada = match.cantidadFacturada || item.cantidadFacturada || 0;
            item.cantidadPendiente = match.cantidadPendiente || Math.max((item.cantidadSolicitada || 0) - (item.cantidadFacturada || 0), 0);
            item.montoPlanificado = match.montoPlanificado || item.montoPlanificado || 0;
            item.montoEntregadoNoFacturado = match.montoEntregadoNoFacturado || item.montoEntregadoNoFacturado || 0;
            item.montoPendienteSinStock = match.montoPendienteSinStock || item.montoPendienteSinStock || 0;
            item.fechaActualizacion = nowIso();
          });
        }

        matchWarehouseToOrders();
        rebuildDerivedState();
        if (!isPlan) {
          registerRouteCostSnapshot(APP.controlFecha);
          registerControlHistorySnapshot(APP.controlFecha, file.name);
        }
        actualizarDashboard();
        renderCalendario();
        renderRutas();
        renderRouteCatalog();
        renderComercial();
        renderAlmacen();
        renderSolicitudesAlmacen();
        progressEl.style.width = '100%';

        setTimeout(() => {
          loadingEl.style.display = 'none';
          uploadZone.style.display = 'block';
          successMsgEl.textContent = summarizeImport(mode, items, file.name, parsed.fecha);
          bannerEl.style.display = 'flex';
          actualizarEstadoBanner();
          saveLocalSnapshot();
          window.guardarEnSupabase().catch(e => console.warn('Error guardando datos:', e));
        }, 220);
      } catch (error) {
        console.error('Error importando archivo:', error);
        loadingEl.style.display = 'none';
        uploadZone.style.display = 'block';
        alert('No se pudo procesar el archivo: ' + error.message);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  window.actualizarEstadoBanner = function actualizarEstadoBannerV2() {
    const banner = document.getElementById('sapEstadoBanner');
    const planTxt = document.getElementById('planEstadoTxt');
    const ctrlTxt = document.getElementById('ctrlEstadoTxt');
    if (!banner || !planTxt || !ctrlTxt) return;
    banner.style.display = (APP.planLoaded || APP.controlLoaded || APP.solicitudesLoaded) ? 'block' : 'none';
    planTxt.textContent = APP.planLoaded ? `${new Set(APP.planLineItems.map(item => item.clienteId)).size} clientes · ${APP.planLineItems.length} líneas · ${APP.planFecha}` : 'No cargada';
    ctrlTxt.textContent = APP.controlLoaded ? `${new Set(APP.controlLineItems.map(item => item.clienteId)).size} clientes · ${APP.controlLineItems.length} líneas · ${APP.controlFecha}` : 'Sin control diario';

    let solCard = document.getElementById('solEstadoCard');
    if (!solCard) {
      solCard = document.createElement('div');
      solCard.id = 'solEstadoCard';
      solCard.style.cssText = 'background:#FFF7ED;border:1px solid #FDBA74;border-radius:10px;padding:14px 16px;display:flex;align-items:center;gap:10px;margin-top:12px;';
      solCard.innerHTML = '<span style="font-size:22px;">📦</span><div><div style="font-size:11px;font-weight:700;color:#C2410C;text-transform:uppercase;">Solicitudes a Almacén</div><div id="solEstadoTxt" style="font-size:13px;color:var(--text);margin-top:2px;">Sin archivo</div></div>';
      banner.appendChild(solCard);
    }
    const solTxt = document.getElementById('solEstadoTxt');
    if (solTxt) solTxt.textContent = APP.solicitudesLoaded ? `${APP.solicitudesAlmacen.length} solicitudes · ${APP.solicitudesFileName || 'archivo actual'}` : 'Sin archivo';
    updateImportExcelTableStatus();
  };

  function parseWarehouseRows(rows) {
    return dedupeBy(rows.map(row => {
      const codigo = text(pickField(row, ['Destinatario de las mercancías descripción', 'ID de destinatario de mercancías', 'ID destinatario', 'Código', 'Codigo']));
      const cliente = text(pickField(row, ['Destinatario de las mercancías', 'Nombre de destinatario de mercancías', 'Cliente', 'Nombre destinatario']));
      const solicitud = text(pickField(row, ['ID de solicitud de almacén', 'ID de solicitud de logística por medio de terceros', 'ID de solicitud', 'Solicitud']));
      const pedidoCliente = text(pickField(row, ['Pedido de cliente (documento precedente)', 'Pedido del cliente', 'Pedido de cliente', 'Pedido']));
      if (!codigo || !solicitud || !pedidoCliente) return null;
      const lineaPedidoCliente = text(pickField(row, ['ID de partida individual', 'Posición de pedido de cliente', 'Posicion de pedido de cliente', 'Línea', 'Linea']));
      const articulo = text(pickField(row, ['Producto', 'Artículo', 'Articulo', 'Material', 'SKU']));
      const descripcionArticulo = text(pickField(row, ['Producto descripción', 'Producto descripcion', 'Artículo descripción', 'Articulo descripcion', 'Descripción del artículo', 'Descripcion del articulo', 'Descripción', 'Descripcion']));
      const cantidadSolicitadaOriginal = num(pickField(row, ['Cantidad solicitada', 'Cantidad', 'Cant solicitada']));
      const cantidadCompletada = num(pickField(row, ['Cantidad procesada', 'Cantidad completada', 'Cantidad facturada', 'Cantidad entregada']));
      const cantidadPendienteServirRaw = pickField(row, ['Cantidad pendiente de servir', 'Cantidad pendiente', 'Pendiente']);
      const cantidadPendienteServir = cantidadPendienteServirRaw !== '' ? num(cantidadPendienteServirRaw) : Math.max(cantidadSolicitadaOriginal - cantidadCompletada, 0);
      const cantidadPendienteSinStock = num(pickField(row, ['Cantidad pendiente sin stock', 'Cantidad sin stock', 'Pendiente sin stock', 'Cant pendiente sin stock']));
      const cantidadSolicitada = getUsableRequestedQty(cantidadPendienteServir, cantidadPendienteSinStock);
      const cantidadPendiente = Math.max(cantidadSolicitada - cantidadCompletada, 0);
      const fecha = toDateLabel(pickField(row, ['Fecha de entrega planificada', 'Fecha de envío planificada', 'Fecha de envio planificada', 'Fecha de liberación', 'Fecha de liberacion'])) || fechaToStr(new Date());
      return {
        codigo,
        cliente,
        solicitud,
        pedidoCliente,
        lineaPedidoCliente,
        articulo,
        descripcionArticulo,
        cantidadSolicitadaOriginal,
        cantidadPendienteServir,
        cantidadPendienteSinStock,
        cantidadSolicitada,
        cantidadCompletada,
        cantidadPendiente,
        estadoComunicacion: text(pickField(row, ['Estado de comunicación', 'Estado de comunicacion'])),
        estadoProceso: text(pickField(row, ['Estado de procesamiento', 'Estado proceso'])),
        fechaLiberacion: toDateLabel(pickField(row, ['Fecha de creación', 'Fecha de creacion', 'Fecha de liberación', 'Fecha de liberacion'])),
        fechaFinalizada: toDateLabel(pickField(row, ['Fecha finalizada'])),
        fechaEnvio: fecha,
        ubicacion: text(pickField(row, ['Ubicación de procedencia', 'Ubicacion de procedencia'])),
        observaciones: text(pickField(row, ['Observaciones', 'Nota'])),
        motivoNoPreparada: text(pickField(row, ['Motivo no preparada', 'Motivo no preparación', 'Motivo no preparacion', 'Razón no preparada', 'Razon no preparada'])),
        estado: cantidadPendiente <= 0 ? 'Completada' : cantidadCompletada > 0 ? 'Parcial' : 'Pendiente',
        costeSolicitud: num(APP.warehouseSettings.costoSolicitud || 0),
        sapMatchKey: [pedidoCliente, lineaPedidoCliente, articulo, codigo].map(normKey).join('|')
      };
    }).filter(Boolean), item => [item.codigo, item.solicitud, item.pedidoCliente, item.lineaPedidoCliente, item.articulo, item.fechaEnvio].join('|'));
  }

  window.parsearSolicitudesRows = parseWarehouseRows;

  window.procesarSolicitudesAlmacen = function procesarSolicitudesAlmacenV2(file, mode) {
    if (!file) return;
    mode = mode || 'plan';
    const reader = new FileReader();
    reader.onload = function (event) {
      try {
        const wb = XLSX.read(new Uint8Array(event.target.result), { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = rowsFromWorksheet(ws);
        const parsed = parseWarehouseRows(rows).map(item => ({
          ...item,
          tipoSolicitudArchivo: mode === 'control' ? 'control diario' : 'plan semanal'
        }));
        pushUndoState(mode === 'control' ? 'importar solicitudes de control' : 'importar solicitudes de planificación');
        if (mode === 'control') {
          APP.solicitudesControlAlmacen = parsed;
          APP.importFiles.solicitudesControl = file.name;
        } else {
          APP.solicitudesPlanAlmacen = parsed;
          APP.importFiles.solicitudesPlan = file.name;
        }
        APP.solicitudesAlmacen = parsed;
        APP.solicitudesLoaded = true;
        APP.solicitudesFileName = file.name;
        APP.solicitudesCompare = buildSolicitudesComparison();
        matchWarehouseToOrders();
        rebuildDerivedState();
        registerSolicitudesSnapshot(file.name);
        updateSolicitudesImportStatus();
        renderSolicitudesAlmacen();
        renderAlmacen();
        actualizarEstadoBanner();
        saveLocalSnapshot();
        window.guardarEnSupabase().catch(e => console.warn('Error guardando solicitudes:', e));
        alert(`Solicitudes ${mode === 'control' ? 'de control diario' : 'de planificación'} cargadas: ${parsed.length}`);
      } catch (error) {
        console.error('Error procesando solicitudes:', error);
        alert('No se pudo procesar el archivo de solicitudes: ' + error.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  function buildSolicitudesComparison() {
    const planMap = new Map(APP.solicitudesPlanAlmacen.map(item => [[item.codigo, item.solicitud, item.pedidoCliente, item.lineaPedidoCliente || '', item.articulo].join('|'), item]).entries());
    const controlMap = new Map(APP.solicitudesControlAlmacen.map(item => [[item.codigo, item.solicitud, item.pedidoCliente, item.lineaPedidoCliente || '', item.articulo].join('|'), item]).entries());
    const keys = new Set([...planMap.keys(), ...controlMap.keys()]);
    return [...keys].map(key => {
      const plan = planMap.get(key);
      const control = controlMap.get(key);
      const current = control || plan;
      const solicitada = (plan && plan.cantidadSolicitada) || (control && control.cantidadSolicitada) || 0;
      const completada = (control && control.cantidadCompletada) || 0;
      const pendiente = Math.max(solicitada - completada, 0);
      return {
        key,
        fecha: (control && control.fechaEnvio) || (plan && plan.fechaEnvio) || fechaToStr(new Date()),
        codigo: current ? current.codigo : '',
        cliente: current ? current.cliente : '',
        solicitud: current ? current.solicitud : '',
        pedidoCliente: current ? current.pedidoCliente : '',
        articulo: current ? current.articulo : '',
        descripcionArticulo: current ? current.descripcionArticulo : '',
        cantidadSolicitada: solicitada,
        cantidadCompletada: completada,
        cantidadPendiente: pendiente,
        estado: pendiente <= 0 ? 'Completada' : completada > 0 ? 'Parcial' : 'Pendiente',
        plan,
        control
      };
    }).sort((a, b) => fechaToDate(a.fecha) - fechaToDate(b.fecha));
  }

  function renderWarehouseCard(item) {
    const badgeClass = item.estado === 'Completada' ? 'badge-ok' : item.estado === 'Parcial' ? 'badge-warn' : 'badge-pend';
    return `<div class="sol-card" style="${item.cantidadPendiente > 0 ? 'border-left:3px solid #f59e0b;' : ''}">
      <div class="sol-card-head">
        <span>Solicitud ${item.solicitud}</span>
        <span class="badge ${badgeClass}">${item.estado}</span>
      </div>
      <div class="sol-card-meta">
        Pedido: <strong>${item.pedidoCliente || '—'}</strong>${item.lineaPedidoCliente ? ' · Línea ' + item.lineaPedidoCliente : ''} · Artículo: <strong>${item.articulo || '—'}</strong><br>
        ${item.descripcionArticulo || 'Sin descripción'}<br>
        Solicitada: ${item.cantidadSolicitada} · Completada: ${item.cantidadCompletada} · Pendiente: ${item.cantidadPendiente}<br>
        ${item.observaciones ? 'Obs.: ' + item.observaciones + '<br>' : ''}
        Procedencia: ${item.ubicacion || '—'}
      </div>
    </div>`;
  }

  window.renderSolicitudCard = renderWarehouseCard;

  function getWarehouseScheduledDate(item) {
    const exact = (APP.lineItems || []).find(line =>
      text(line.clienteId) === text(item.codigo) &&
      (!item.pedidoCliente || text(line.pedidoCliente) === text(item.pedidoCliente))
    );
    const anyClientLine = exact || (APP.lineItems || []).find(line => text(line.clienteId) === text(item.codigo));
    if (anyClientLine && anyClientLine.fechaPlanificada) return anyClientLine.fechaPlanificada;
    const legacyClient = (APP.clientes || []).find(client => text(client.codigo) === text(item.codigo));
    return (legacyClient && legacyClient.fecha) || item.fechaEnvio || item.fechaLiberacion || fechaToStr(new Date());
  }

  function getWarehouseClientName(item) {
    const line = (APP.lineItems || []).find(row => text(row.clienteId) === text(item.codigo));
    const legacyClient = (APP.clientes || []).find(client => text(client.codigo) === text(item.codigo));
    return item.cliente || (line && line.clienteNombre) || (legacyClient && legacyClient.nombre) || item.codigo || 'Cliente sin nombre';
  }

  function getWarehouseRequestAmount(item) {
    const exact = (APP.lineItems || []).filter(line =>
      text(line.clienteId) === text(item.codigo) &&
      (!item.pedidoCliente || text(line.pedidoCliente) === text(item.pedidoCliente)) &&
      (!item.articulo || text(line.articulo) === text(item.articulo))
    );
    const source = exact.length ? exact : (APP.lineItems || []).filter(line => text(line.clienteId) === text(item.codigo));
    const amount = source.reduce((sum, line) => sum + getOperationalAmount(line), 0);
    return amount || num(item.costeSolicitud || APP.warehouseSettings.costoSolicitud || 0);
  }

  function getWarehouseRowsForView(mode) {
    const source = mode === 'completadas'
      ? (APP.solicitudesControlAlmacen.length ? APP.solicitudesControlAlmacen : APP.solicitudesAlmacen)
      : (APP.solicitudesPlanAlmacen.length ? APP.solicitudesPlanAlmacen : APP.solicitudesAlmacen);
    return (source || []).map(item => ({
      ...item,
      fechaProgramada: getWarehouseScheduledDate(item),
      clienteNombreVista: getWarehouseClientName(item),
      montoProgramado: getWarehouseRequestAmount(item),
      estadoVista: item.cantidadPendiente <= 0 ? 'Completada' : item.cantidadCompletada > 0 ? 'Parcial' : 'Pendiente'
    }));
  }

  function renderWarehouseProgrammedCard(item, showReason) {
    const badgeClass = item.estadoVista === 'Completada' ? 'badge-ok' : item.estadoVista === 'Parcial' ? 'badge-warn' : 'badge-pend';
    return `<div class="sol-card" style="${item.estadoVista !== 'Completada' ? 'border-left:3px solid #f59e0b;' : ''}">
      <div class="sol-card-head"><span>Solicitud ${escapeHtml(item.solicitud || '—')}</span><span class="badge ${badgeClass}">${escapeHtml(item.estadoVista)}</span></div>
      <div class="sol-card-meta">
        Pedido: <strong>${escapeHtml(item.pedidoCliente || '—')}</strong>${item.lineaPedidoCliente ? ' · Línea ' + escapeHtml(item.lineaPedidoCliente) : ''} · Artículo: <strong>${escapeHtml(item.articulo || '—')}</strong><br>
        ${escapeHtml(item.descripcionArticulo || 'Sin descripción')}<br>
        Solicitada: ${num(item.cantidadSolicitada)} · Completada: ${num(item.cantidadCompletada)} · Pendiente: ${num(item.cantidadPendiente)} · Monto: ${formatMonto(item.montoProgramado)}<br>
        En calendario: <strong>${escapeHtml(item.fechaProgramada)}</strong> · Archivo: ${escapeHtml(item.tipoSolicitudArchivo || '—')}
        ${showReason ? `<div style="margin-top:8px;"><label style="display:block;font-size:11px;font-weight:800;color:var(--muted);margin-bottom:4px;">Motivo de no preparación</label><input class="form-control" value="${escapeHtml(item.motivoNoPreparada || '')}" placeholder="Ej. falta stock, documento pendiente, no liberada..." onchange="actualizarMotivoSolicitud('${jsString(item.solicitud)}','${jsString(item.pedidoCliente)}','${jsString(item.articulo)}','${jsString(item.codigo)}',this.value)"></div>` : ''}
      </div>
    </div>`;
  }

  function renderWarehouseGroupedByDate(rows, mode) {
    const search = text((document.getElementById('solSearch') || {}).value).toLowerCase();
    const filtered = rows.filter(item => {
      if (!search) return true;
      return [item.codigo, item.clienteNombreVista, item.solicitud, item.pedidoCliente, item.articulo, item.descripcionArticulo]
        .some(value => text(value).toLowerCase().includes(search));
    });
    if (!filtered.length) return '<div class="empty-state"><div class="empty-icon">🔍</div><p>Sin resultados para esta vista.</p></div>';
    const byDate = new Map();
    filtered.forEach(item => {
      const date = item.fechaProgramada || 'Sin fecha';
      if (!byDate.has(date)) byDate.set(date, new Map());
      const clients = byDate.get(date);
      const code = item.codigo || 'sin-codigo';
      if (!clients.has(code)) clients.set(code, []);
      clients.get(code).push(item);
    });
    return [...byDate.entries()].sort((a, b) => fechaToDate(a[0]) - fechaToDate(b[0])).map(([date, clients]) => {
      const allItems = [...clients.values()].flat();
      const amount = allItems.reduce((sum, item) => sum + num(item.montoProgramado), 0);
      return `<div class="card" style="margin-bottom:14px;">
        <div class="card-title" style="display:flex;justify-content:space-between;gap:10px;align-items:center;"><span>Fecha ${escapeHtml(date)} · ${escapeHtml(fechaLabel(date))}</span><span class="badge badge-pend">${allItems.length} solicitudes · ${formatMonto(amount)}</span></div>
        ${[...clients.entries()].sort((a, b) => text(a[1][0].clienteNombreVista).localeCompare(text(b[1][0].clienteNombreVista))).map(([code, items]) => {
          const clientAmount = items.reduce((sum, item) => sum + num(item.montoProgramado), 0);
          return `<details style="margin-bottom:10px;" open><summary style="cursor:pointer;font-weight:800;color:var(--primary);margin-bottom:8px;">${escapeHtml(items[0].clienteNombreVista)} · ${escapeHtml(code)} · ${items.length} solicitudes · ${formatMonto(clientAmount)}</summary>${items.map(item => renderWarehouseProgrammedCard(item, mode === 'pendientes')).join('')}</details>`;
        }).join('')}
      </div>`;
    }).join('');
  }

  function buildWarehouseThursdaySummary(rows) {
    const pending = rows.filter(item => item.estadoVista !== 'Completada');
    if (!pending.length) return 'Resumen jueves: no hay solicitudes pendientes de preparar.';
    const withoutReason = pending.filter(item => !text(item.motivoNoPreparada)).length;
    return `Resumen jueves: ${pending.length} solicitudes pendientes de preparar. ${withoutReason} sin motivo registrado.`;
  }

  window.actualizarMotivoSolicitud = function actualizarMotivoSolicitud(solicitud, pedidoCliente, articulo, codigo, value) {
    if (!hasPermission('solicitudesAlmacen', 'editar')) return alert('Tu rol no puede editar solicitudes.');
    const apply = item => {
      if (text(item.solicitud) === text(solicitud) && text(item.pedidoCliente) === text(pedidoCliente) && text(item.articulo) === text(articulo) && text(item.codigo) === text(codigo)) item.motivoNoPreparada = text(value);
    };
    pushUndoState('actualizar motivo solicitud');
    [...(APP.solicitudesAlmacen || []), ...(APP.solicitudesPlanAlmacen || []), ...(APP.solicitudesControlAlmacen || [])].forEach(apply);
    saveLocalSnapshot();
    scheduleAutoSave();
  };

  window.cambiarVistaSolicitudes = function cambiarVistaSolicitudes(mode) {
    APP.solicitudesVista = mode || 'programadas';
    renderSolicitudesAlmacen();
  };

  window.renderSolicitudCard = renderWarehouseCard;

  window.renderSolicitudesAlmacen = function renderSolicitudesAlmacenV2() {
    const cont = document.getElementById('solicitudesContainer');
    const resumen = document.getElementById('solicitudesResumen');
    if (!cont || !resumen) return;
    const baseRows = APP.solicitudesPlanAlmacen.length || APP.solicitudesControlAlmacen.length ? [...APP.solicitudesPlanAlmacen, ...APP.solicitudesControlAlmacen] : APP.solicitudesAlmacen;
    resumen.textContent = APP.solicitudesLoaded ? `${baseRows.length} solicitudes` : 'Sin archivo';
    if (!APP.solicitudesLoaded || !baseRows.length) {
      cont.innerHTML = '<div class="empty-state"><div class="empty-icon">📦</div><p>No hay solicitudes registradas.</p></div>';
      return;
    }
    const mode = APP.solicitudesVista || 'programadas';
    const planRows = getWarehouseRowsForView('programadas');
    const pendingRows = planRows.filter(item => item.estadoVista !== 'Completada');
    const completedRows = getWarehouseRowsForView('completadas').filter(item => item.estadoVista === 'Completada');
    const rowsByMode = { programadas: planRows, pendientes: pendingRows, completadas: completedRows };
    const activeRows = rowsByMode[mode] || planRows;
    const totalAmount = activeRows.reduce((sum, item) => sum + num(item.montoProgramado), 0);
    cont.innerHTML = `
      <div class="card" style="border-left:4px solid var(--secondary);margin-bottom:14px;">
        <div class="card-title">Solicitudes programadas</div>
        <div style="font-size:12px;color:var(--muted);line-height:1.45;margin-bottom:10px;">Agrupadas por la fecha de entrega definida en Calendario. Usa pendientes para registrar por qué no se prepararon antes del resumen de jueves.</div>
        <div class="view-toolbar" style="gap:8px;">
          <button class="filter-btn ${mode === 'programadas' ? 'active' : ''}" onclick="cambiarVistaSolicitudes('programadas')">Solicitudes planificadas</button>
          <button class="filter-btn ${mode === 'pendientes' ? 'active' : ''}" onclick="cambiarVistaSolicitudes('pendientes')">Pendientes de preparar</button>
          <button class="filter-btn ${mode === 'completadas' ? 'active' : ''}" onclick="cambiarVistaSolicitudes('completadas')">Completadas</button>
          <span class="badge badge-c1">${activeRows.length} solicitudes</span>
          <span class="badge badge-ok">${formatMonto(totalAmount)}</span>
        </div>
        <div class="import-report" style="margin-top:8px;">${escapeHtml(buildWarehouseThursdaySummary(pendingRows))}</div>
      </div>
      ${renderWarehouseGroupedByDate(activeRows, mode)}
    `;
  };

  window.renderAlmacen = function renderAlmacenV2() {
    const cont = document.getElementById('almContainer');
    if (!cont) return;
    const search = text((document.getElementById('almSearch') || {}).value).toLowerCase();
    const list = APP.clientes.filter(item => {
      if (!search) return true;
      return [item.codigo, item.nombre, item.zona].some(value => text(value).toLowerCase().includes(search));
    });
    document.getElementById('ak1').textContent = APP.lineItems.length;
    document.getElementById('ak2').textContent = APP.lineItems.filter(item => item.camionAsignado === 'CAMION 1').length;
    document.getElementById('ak3').textContent = APP.lineItems.filter(item => item.camionAsignado === 'CAMION 2').length;
    const nextDate = [...new Set(APP.lineItems.map(item => item.fechaPlanificada).filter(Boolean))].sort((a, b) => fechaToDate(a) - fechaToDate(b))[0];
    document.getElementById('ak4').textContent = nextDate ? fechaLabel(nextDate) : '—';
    document.getElementById('ak4b').textContent = nextDate || '';
    if (!list.length) {
      cont.innerHTML = '<div class="empty-state"><div class="empty-icon">🏭</div><p>Sin datos cargados todavía.</p></div>';
      return;
    }
    const byDate = {};
    APP.lineItems.forEach(item => {
      const key = item.fechaPlanificada || 'Sin fecha';
      if (!byDate[key]) byDate[key] = [];
      byDate[key].push(item);
    });
    cont.innerHTML = Object.keys(byDate).sort((a, b) => fechaToDate(a) - fechaToDate(b)).map(date => {
      const items = byDate[date];
      return `<div style="margin-bottom:18px;">
        <div style="background:var(--primary);color:#fff;padding:8px 12px;border-radius:8px;font-weight:700;margin-bottom:10px;">📅 ${date}</div>
        ${getTruckOptions().map(camion => {
          const truckItems = items.filter(item => item.camionAsignado === camion);
          if (!truckItems.length) return '';
          return `<div class="card">
            <div class="card-title">${getTruckLabel(camion)} · ${truckItems.length} líneas</div>
            ${truckItems.map(item => {
              const solicitudes = APP.solicitudesAlmacen.filter(sol => sol.codigo === item.clienteId);
              return `<div class="ruta-item">
                <div style="display:flex;justify-content:space-between;gap:12px;">
                  <div>
                    <div class="ruta-item-name">${item.clienteNombre} · Pedido ${item.pedidoCliente || '—'}</div>
                    <div class="ruta-item-detail">${item.articulo || 'Sin artículo'} · ${item.descripcionArticulo || 'Sin descripción'} · ${item.rutaNombre || 'Sin ruta'}</div>
                  </div>
                  <div style="text-align:right;">
                    <div class="badge ${getTruckBadgeClass(camion)}">${deriveItemStatus(item)}</div>
                    <div style="font-size:11px;color:var(--muted);margin-top:4px;">${solicitudes.length} solicitudes</div>
                  </div>
                </div>
              </div>`;
            }).join('')}
          </div>`;
        }).join('')}
      </div>`;
    }).join('');
  };

  function getCommercialWarehouseSource() {
    const planRequests = APP.solicitudesPlanAlmacen || [];
    const activeRequests = APP.solicitudesAlmacen || [];
    return planRequests.length ? planRequests : activeRequests;
  }

  function findOrderLineForWarehouseItem(item) {
    const targetKey = [
      item.pedidoCliente || '',
      item.lineaPedidoCliente || '',
      item.articulo || '',
      item.codigo || item.clienteId || ''
    ].map(normKey).join('|');
    const candidates = [...(APP.lineItems || []), ...(APP.planLineItems || []), ...(APP.controlLineItems || [])];
    return candidates.find(line => {
      const lineKey = [
        line.pedidoCliente || '',
        line.lineaPedidoCliente || '',
        line.articulo || '',
        line.clienteId || ''
      ].map(normKey).join('|');
      return lineKey === targetKey;
    }) || candidates.find(line =>
      normKey(line.pedidoCliente) === normKey(item.pedidoCliente) &&
      normKey(line.articulo) === normKey(item.articulo) &&
      normKey(line.clienteId) === normKey(item.codigo || item.clienteId)
    ) || null;
  }

  function getCommercialWarehouseQty(item) {
    const requested = num(item.cantidadAlmacenSolicitada !== undefined ? item.cantidadAlmacenSolicitada : item.cantidadSolicitada);
    const processed = num(item.cantidadAlmacenProcesada !== undefined ? item.cantidadAlmacenProcesada : item.cantidadCompletada);
    return requested > 0 ? requested : processed;
  }

  function getCommercialLineAmount(item) {
    const orderLine = item.clienteId ? item : findOrderLineForWarehouseItem(item);
    if (!orderLine) return 0;
    const orderQty = num(orderLine.cantidadSolicitada);
    const warehouseQty = getCommercialWarehouseQty(item);
    const amount = getOperationalAmount(orderLine);
    if (warehouseQty <= 0) return 0;
    if (orderQty > 0 && warehouseQty < orderQty) return amount * (warehouseQty / orderQty);
    return amount;
  }

  function getCommercialPlannedQty(item) {
    return getCommercialWarehouseQty(item);
  }

  function getCommercialDateRange() {
    const dates = getCommercialWarehouseSource().map(item => item.fechaEnvio || item.fechaLiberacion).filter(Boolean).sort((a, b) => fechaToDate(a) - fechaToDate(b));
    return { first: dates[0] || '', last: dates[dates.length - 1] || '' };
  }

  function getCommercialRouteOptions() {
    const warehouseRouteNames = getCommercialWarehouseSource().map(item => {
      const orderLine = item.clienteId ? item : findOrderLineForWarehouseItem(item);
      const cfg = getClienteConfigV2(item.codigo || item.clienteId, item.cliente || item.clienteNombre || '');
      return (orderLine && (orderLine.rutaNombre || orderLine.zona || orderLine.camionAsignado)) || cfg.rutaAsignada || cfg.camionPermitido || '';
    });
    return [...new Set([
      ...warehouseRouteNames,
      ...APP.lineItems.map(item => item.rutaNombre || item.zona),
      ...APP.planLineItems.map(item => item.rutaNombre || item.zona),
      ...APP.controlLineItems.map(item => item.rutaNombre || item.zona),
      ...APP.lineItems.map(item => item.camionAsignado),
      ...Object.values(APP.clienteConfig || {}).map(cfg => cfg.rutaAsignada),
      ...APP.routeConfigs.map(route => route.nombre || route.zona)
    ].filter(Boolean))].sort();
  }

  function syncCommercialRouteFilterOptions() {
    const select = document.getElementById('cfComRuta');
    if (!select) return;
    const current = select.value;
    const options = getCommercialRouteOptions();
    select.innerHTML = '<option value="">Todas las rutas/camiones</option>' + options.map(route => `<option value="${route}">${route}</option>`).join('');
    if (current && options.includes(current)) select.value = current;
  }

  function isoToFechaInput(value) {
    if (!value) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
    const d = fechaToDate(value);
    if (Number.isNaN(d.getTime())) return '';
    return d.toISOString().slice(0, 10);
  }

  function commercialFilterDateToDate(value) {
    if (!value) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      const parts = value.split('-').map(Number);
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    return fechaToDate(value);
  }

  function commercialRowsFromItems(items) {
    return items.map(item => {
      const orderLine = item.clienteId ? item : findOrderLineForWarehouseItem(item);
      const cfg = getClienteConfigV2(item.codigo || item.clienteId, item.cliente || item.clienteNombre || '');
      const warehouseQty = getCommercialWarehouseQty(item);
      const section = warehouseQty > 0 ? 'en_ruta' : 'no_considerado';
      const routeName = (orderLine && (orderLine.rutaNombre || orderLine.zona)) || cfg.rutaAsignada || '';
      const truckName = (orderLine && orderLine.camionAsignado) || cfg.camionPermitido || '';
      return {
        fecha: item.fechaEnvio || item.fechaLiberacion || (orderLine && orderLine.fechaPlanificada) || '',
        cliente: item.cliente || item.clienteNombre || '',
        codigoCliente: item.codigo || item.clienteId || '',
        ruta: routeName,
        camion: getTruckLabel(truckName || ''),
        pedido: item.pedidoCliente || '',
        referencia: item.articulo || '',
        descripcion: item.descripcionArticulo || '',
        solicitud: item.solicitud || item.solicitudAlmacen || '',
        cantidadPedida: orderLine ? num(orderLine.cantidadSolicitada) : num(item.cantidadSolicitadaOriginal || item.cantidadSolicitada),
        cantidadPlanificada: warehouseQty,
        cantidadFacturada: orderLine ? num(orderLine.cantidadFacturada) : num(item.cantidadCompletada),
        monto: getCommercialLineAmount(item),
        estado: section === 'en_ruta' ? text(item.estado || item.estadoAlmacen || deriveItemStatus(item)) : 'No considerado en ruta',
        section
      };
    });
  }

  function renderCommercialRowsTable(rows, extraClass) {
    return `<div class="commercial-table-wrap ${extraClass || ''}">
      <table class="commercial-detail-table">
        <thead><tr><th>Pedido</th><th>Solicitud</th><th>Ref.</th><th>Descripción</th><th>Pedida</th><th>En ruta</th><th>Fact.</th><th>Monto</th></tr></thead>
        <tbody>${rows.map(row => `<tr>
          <td>${row.pedido || '—'}</td>
          <td>${row.solicitud || '—'}</td>
          <td>${row.referencia || '—'}</td>
          <td>${row.descripcion || '—'}</td>
          <td>${row.cantidadPedida}</td>
          <td>${row.cantidadPlanificada}</td>
          <td>${row.cantidadFacturada}</td>
          <td>${formatMonto(row.monto)}</td>
        </tr>`).join('')}</tbody>
      </table>
    </div>`;
  }

  function getCommercialFilteredItems() {
    const clientFilter = text((document.getElementById('cfComCliente') || {}).value).toLowerCase();
    const quickDate = text((document.getElementById('cfcol_fecha') || {}).value);
    const dateStart = text((document.getElementById('cfComFechaInicio') || {}).value);
    const dateEnd = text((document.getElementById('cfComFechaFin') || {}).value);
    const routeFilter = text((document.getElementById('cfComRuta') || {}).value).toLowerCase();
    const stateFilter = text((document.getElementById('cfComEstado') || {}).value).toLowerCase();
    return getCommercialWarehouseSource().filter(item => {
      const orderLine = item.clienteId ? item : findOrderLineForWarehouseItem(item);
      const cfg = getClienteConfigV2(item.codigo || item.clienteId, item.cliente || item.clienteNombre || '');
      const itemDateText = item.fechaEnvio || item.fechaLiberacion || (orderLine && orderLine.fechaPlanificada) || '';
      const itemDate = fechaToDate(itemDateText || '');
      const routeName = (orderLine && (orderLine.rutaNombre || orderLine.zona || orderLine.camionAsignado)) || cfg.rutaAsignada || cfg.camionPermitido || '';
      const startDate = commercialFilterDateToDate(dateStart);
      const endDate = commercialFilterDateToDate(dateEnd);
      if (clientFilter && ![item.cliente || item.clienteNombre, item.codigo || item.clienteId, item.pedidoCliente, item.solicitud, item.articulo, item.descripcionArticulo].some(value => text(value).toLowerCase().includes(clientFilter))) return false;
      if (quickDate && !text(itemDateText).includes(quickDate)) return false;
      if (startDate && itemDate < startDate) return false;
      if (endDate && itemDate > endDate) return false;
      if (routeFilter && !text(routeName).toLowerCase().includes(routeFilter)) return false;
      const warehouseQty = getCommercialWarehouseQty(item);
      const commercialState = warehouseQty > 0 ? text(item.estado || item.estadoAlmacen || deriveItemStatus(item)) : 'No considerado en ruta';
      if (stateFilter && !commercialState.toLowerCase().includes(stateFilter)) return false;
      return true;
    }).sort((a, b) => fechaToDate(a.fechaEnvio || a.fechaLiberacion) - fechaToDate(b.fechaEnvio || b.fechaLiberacion) || text(a.cliente || a.clienteNombre).localeCompare(text(b.cliente || b.clienteNombre)) || text(a.pedidoCliente).localeCompare(text(b.pedidoCliente)));
  }

  function setCommercialFilterDates(start, end) {
    const startEl = document.getElementById('cfComFechaInicio');
    const endEl = document.getElementById('cfComFechaFin');
    if (startEl) startEl.value = isoToFechaInput(start);
    if (endEl) endEl.value = isoToFechaInput(end || start);
    renderComercial();
  }

  window.filtrarCom = function filtrarComV2(mode) {
    document.querySelectorAll('#comercial .filter-btn').forEach(btn => btn.classList.remove('active'));
    const active = document.getElementById('cf_' + mode);
    if (active) active.classList.add('active');
    const range = getCommercialDateRange();
    const dates = [...new Set(getCommercialWarehouseSource().map(item => item.fechaEnvio || item.fechaLiberacion).filter(Boolean))].sort((a, b) => fechaToDate(a) - fechaToDate(b));
    const routeEl = document.getElementById('cfComRuta');
    if (routeEl && !['c1', 'c2'].includes(mode)) routeEl.value = '';
    if (mode === 'todas') setCommercialFilterDates('', '');
    if (mode === 'sem1') setCommercialFilterDates(dates[0] || range.first, dates[Math.min(4, dates.length - 1)] || range.last);
    if (mode === 'sem2') setCommercialFilterDates(dates[5] || range.first, dates[Math.min(9, dates.length - 1)] || range.last);
    if (mode === 'sem3') setCommercialFilterDates(dates[10] || range.first, dates[Math.min(14, dates.length - 1)] || range.last);
    if (mode === 'c1' || mode === 'c2') {
      if (routeEl) routeEl.value = mode === 'c1' ? 'CAMION 1' : 'CAMION 2';
      renderComercial();
    }
  };

  window.limpiarFiltrosCom = function limpiarFiltrosComV2() {
    ['cfComCliente', 'cfComFechaInicio', 'cfComFechaFin', 'cfComRuta', 'cfComEstado', 'cfcol_fecha'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    document.querySelectorAll('#comercial .filter-btn').forEach(btn => btn.classList.remove('active'));
    const all = document.getElementById('cf_todas');
    if (all) all.classList.add('active');
    renderComercial();
  };

  window.renderComercial = function renderComercialV2() {
    const mount = document.getElementById('comercialV2Mount');
    if (!mount) return;
    syncCommercialRouteFilterOptions();
    const items = getCommercialFilteredItems();
    const rows = commercialRowsFromItems(items);
    const deliveryRows = rows.filter(row => row.section === 'en_ruta');
    const totalMonto = deliveryRows.reduce((sum, row) => sum + row.monto, 0);
    const totalPlanificada = deliveryRows.reduce((sum, row) => sum + row.cantidadPlanificada, 0);
    const totalFacturada = deliveryRows.reduce((sum, row) => sum + row.cantidadFacturada, 0);
    const pedidos = new Set(deliveryRows.map(row => row.pedido).filter(Boolean));

    const ck1 = document.getElementById('ck1');
    const ck2 = document.getElementById('ck2');
    const ck3 = document.getElementById('ck3');
    const ck4 = document.getElementById('ck4');
    if (ck1) ck1.textContent = new Set(deliveryRows.map(row => row.codigoCliente)).size;
    if (ck2) ck2.textContent = pedidos.size;
    if (ck3) ck3.textContent = totalPlanificada;
    if (ck4) ck4.textContent = formatMonto(totalMonto);

    if (!rows.length || !deliveryRows.length) {
      const message = APP.solicitudesLoaded
        ? 'No hay productos con cantidad solicitada al almacén para los filtros aplicados.'
        : 'Carga primero las solicitudes realizadas al almacén para construir la vista comercial.';
      mount.innerHTML = `<div class="empty-state"><div class="empty-icon">👔</div><p>${message}</p></div>`;
      return;
    }

    const byDate = {};
    rows.forEach(row => {
      const date = row.fecha || 'Sin fecha';
      if (!byDate[date]) byDate[date] = [];
      byDate[date].push(row);
    });

    mount.innerHTML = `
      <div id="comercialExportArea" class="commercial-calendar">
        ${Object.keys(byDate).sort((a, b) => fechaToDate(a) - fechaToDate(b)).map(date => {
          const dayRows = byDate[date];
          const dayMonto = dayRows.reduce((sum, row) => sum + row.monto, 0);
          const dayPedidos = new Set(dayRows.map(row => row.pedido).filter(Boolean)).size;
          const byClient = {};
          dayRows.forEach(row => {
            const key = row.codigoCliente + '|' + row.cliente;
            if (!byClient[key]) byClient[key] = [];
            byClient[key].push(row);
          });
          return `<section class="commercial-day">
            <div class="commercial-day-head">
              <div><strong>${fechaLabel(date)}</strong><span>${date}</span></div>
              <div>${Object.keys(byClient).length} clientes · ${dayPedidos} pedidos · ${formatMonto(dayMonto)}</div>
            </div>
            <div class="commercial-client-grid">
              ${Object.keys(byClient).sort().map(key => {
                const clientRows = byClient[key];
                const routeRows = clientRows.filter(row => row.section === 'en_ruta');
                const excludedRows = clientRows.filter(row => row.section !== 'en_ruta');
                const visibleRows = routeRows.length ? routeRows : clientRows;
                const clientMonto = routeRows.reduce((sum, row) => sum + row.monto, 0);
                const clientPedidos = new Set(routeRows.map(row => row.pedido).filter(Boolean)).size;
                return `<article class="commercial-client-card">
                  <div class="commercial-client-head">
                    <div><strong>${visibleRows[0].cliente}</strong><span>${visibleRows[0].codigoCliente} · ${visibleRows[0].ruta || 'Sin ruta'} · ${visibleRows[0].camion || 'Sin camión'}</span></div>
                    <div>${clientPedidos} pedidos<br>${formatMonto(clientMonto)}</div>
                  </div>
                  <div class="commercial-section-label">En ruta según solicitudes a almacén</div>
                  ${renderCommercialRowsTable(routeRows, '')}
                  ${excludedRows.length ? `<div class="commercial-section-label muted">No considerado en ruta</div>${renderCommercialRowsTable(excludedRows, 'commercial-muted-table')}` : ''}
                </article>`;
              }).join('')}
            </div>
          </section>`;
        }).join('')}
      </div>
    `;
  };

  window.exportarComercialExcel = function exportarComercialExcelV2() {
    const data = commercialRowsFromItems(getCommercialFilteredItems()).map(row => ({
      'Fecha entrega': row.fecha,
      'Cliente': row.cliente,
      'Código cliente': row.codigoCliente,
      'Ruta': row.ruta,
      'Camión': row.camion,
      'Pedido': row.pedido,
      'Referencia': row.referencia,
      'Descripción': row.descripcion,
      'Cantidad pedida': row.cantidadPedida,
      'Solicitud almacén': row.solicitud,
      'Cantidad planificada entrega': row.cantidadPlanificada,
      'Cantidad facturada': row.cantidadFacturada,
      'Monto': row.monto,
      'Estado': row.estado,
      'Sección': row.section === 'en_ruta' ? 'En ruta' : 'No considerado en ruta'
    }));
    if (!data.length) {
      alert('No hay datos comerciales para exportar.');
      return;
    }
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [
      { wch: 14 }, { wch: 28 }, { wch: 14 }, { wch: 18 }, { wch: 12 }, { wch: 12 },
      { wch: 14 }, { wch: 42 }, { wch: 14 }, { wch: 22 }, { wch: 16 }, { wch: 14 }, { wch: 12 }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Vista Comercial');
    XLSX.writeFile(wb, 'TMS_Vista_Comercial.xlsx');
  };

  window.exportarComercialPDF = async function exportarComercialPDFV2() {
    const area = document.getElementById('comercialExportArea');
    if (!area || !area.innerHTML.trim()) {
      alert('No hay vista comercial para exportar.');
      return;
    }
    let pdfTools;
    try {
      pdfTools = await ensurePdfTools();
    } catch (error) {
      console.error('Error cargando PDF comercial:', error);
      alert(error.message || 'La librería de PDF no está disponible en este momento.');
      return;
    }
    const wrapper = document.createElement('div');
    wrapper.className = 'commercial-pdf-stage';
    wrapper.innerHTML = `<div class="commercial-pdf-title"><strong>Vista Comercial</strong><span>${new Date().toLocaleDateString('es-DO')}</span></div>`;
    const clone = area.cloneNode(true);
    clone.classList.add('commercial-pdf-export');
    wrapper.appendChild(clone);
    document.body.appendChild(wrapper);
    try {
      await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const canvas = await pdfTools.html2canvas(wrapper, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        width: wrapper.scrollWidth,
        height: wrapper.scrollHeight,
        windowWidth: wrapper.scrollWidth,
        windowHeight: wrapper.scrollHeight,
        scrollX: 0,
        scrollY: 0
      });
      if (!canvasHasVisibleContent(canvas)) {
        throw new Error('La captura de la vista comercial salió vacía.');
      }
      const pdf = new pdfTools.JsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 8;
      const imgW = pageW - margin * 2;
      const imgH = canvas.height * imgW / canvas.width;
      const pageImgH = pageH - margin * 2;
      let remaining = imgH;
      let position = margin;
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.95), 'JPEG', margin, position, imgW, imgH);
      remaining -= pageImgH;
      while (remaining > 0) {
        pdf.addPage('a4', 'portrait');
        position = margin - (imgH - remaining);
        pdf.addImage(canvas.toDataURL('image/jpeg', 0.95), 'JPEG', margin, position, imgW, imgH);
        remaining -= pageImgH;
      }
      downloadPdf(pdf, 'TMS_Vista_Comercial.pdf');
    } catch (error) {
      console.error('Error exportando vista comercial:', error);
      alert('No se pudo exportar la vista comercial: ' + error.message);
    } finally {
      wrapper.remove();
    }
  };

  function parseReceptionDays(raw) {
    const map = { lun: 1, lunes: 1, mar: 2, martes: 2, mie: 3, miercoles: 3, miercoles: 3, mié: 3, jue: 4, jueves: 4, vie: 5, viernes: 5 };
    return text(raw).split(/[,;/]+/).map(part => map[normKey(part)]).filter(Boolean);
  }

  function clienteTieneConfigV2(codigo) {
    const cfg = getClienteConfigV2(codigo);
    return !!(cfg.rutaAsignada || cfg.contacto || cfg.observaciones || cfg.condicionesEspeciales || cfg.diasRecepcion.length || cfg.horarioAlmacen);
  }

  window.clienteTieneConfig = clienteTieneConfigV2;

  window.renderConfigClientes = function renderConfigClientesV2() {
    const list = document.getElementById('configClientList');
    const resumen = document.getElementById('configResumen');
    if (!list || !resumen) return;
    const search = text((document.getElementById('configSearch') || {}).value).toLowerCase();
    const baseClients = dedupeBy(APP.planificacion.length ? APP.planificacion : APP.clientes, item => item.codigo);
    const filtered = baseClients.filter(item => !search || [item.codigo, item.nombre].some(value => text(value).toLowerCase().includes(search)));
    resumen.textContent = `${filtered.filter(item => clienteTieneConfigV2(item.codigo)).length} configurados`;
    if (!filtered.length) {
      list.innerHTML = '<div class="empty-state"><div class="empty-icon">⚙️</div><p>Sin clientes para configurar.</p></div>';
      return;
    }
    if (!APP.configClienteSel || !filtered.some(item => item.codigo === APP.configClienteSel)) APP.configClienteSel = filtered[0].codigo;
    const pendingHelp = '<div class="config-help-note"><strong>Pendiente</strong> significa que falta guardar la configuración completa del cliente: ruta asignada, días de recepción y horario de almacén.</div>';
    list.innerHTML = filtered.map(item => {
      const cfg = getClienteConfigV2(item.codigo, item.nombre);
      return `<button class="config-client-item ${APP.configClienteSel === item.codigo ? 'active' : ''}" onclick="seleccionarConfigCliente('${item.codigo}')">
        <strong>${item.nombre}</strong>${cfg.configuracionCompleta ? '<span class="badge badge-ok" style="float:right;">Completa</span>' : '<span class="badge badge-warn" style="float:right;">Pendiente</span>'}<br>
        <span style="font-size:11px;color:var(--muted);">${item.codigo} · ${cfg.rutaAsignada || item.zona || 'Sin ruta'}</span>
      </button>`;
    }).join('') + pendingHelp;
    renderConfigEditor();
    renderConfigAuxSections();
  };

  window.renderConfigEditor = function renderConfigEditorV2() {
    const editor = document.getElementById('configEditor');
    if (!editor) return;
    const client = (APP.planificacion.find(item => item.codigo === APP.configClienteSel) || APP.clientes.find(item => item.codigo === APP.configClienteSel));
    if (!client) {
      editor.innerHTML = '<div class="empty-state"><div class="empty-icon">👈</div><p>Selecciona un cliente.</p></div>';
      return;
    }
    const cfg = getClienteConfigV2(client.codigo, client.nombre);
    const routeOptions = dedupeBy([
      ...APP.routeConfigs,
      { id: '', nombre: client.zona || inferRouteName(client.nombre, client.codigo), zona: client.zona || '', activa: true }
    ], item => item.nombre || item.zona);
    editor.innerHTML = `
      <div class="card-title" style="margin-bottom:4px;">${client.nombre}</div>
      <div style="font-size:12px;color:var(--muted);margin-bottom:16px;">${client.codigo} · ${cfg.rutaAsignada || client.zona || 'Sin ruta'}</div>
      <div class="form-row"><label>Ruta / zona asignada</label>
        <select id="cfgRuta" class="form-control">
          <option value="">Sin asignar</option>
          ${routeOptions.map(route => `<option value="${route.nombre || route.zona}" ${cfg.rutaAsignada === (route.nombre || route.zona) ? 'selected' : ''}>${route.nombre || route.zona}</option>`).join('')}
        </select>
      </div>
      <div class="form-row"><label>Días de recepción</label><input id="cfgDias" class="form-control" value="${(cfg.diasRecepcion || []).join(',')}" placeholder="1,2,4,5 = Lun,Mar,Jue,Vie"></div>
      <div class="form-row"><label>Horario almacén</label><input id="cfgHorario" class="form-control" value="${cfg.horarioAlmacen || ''}" placeholder="08:00 - 14:00"></div>
      <div class="form-row"><label>Contacto</label><input id="cfgContacto" class="form-control" value="${cfg.contacto || ''}" placeholder="Nombre / teléfono / correo"></div>
      <div class="form-row"><label>Condiciones especiales</label><textarea id="cfgCondiciones" class="form-control" style="min-height:70px;">${cfg.condicionesEspeciales || ''}</textarea></div>
      <div class="form-row"><label>Observaciones</label><textarea id="cfgNota" class="form-control" style="min-height:90px;">${cfg.observaciones || ''}</textarea></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;">
        <button class="btn btn-primary" onclick="guardarConfigCliente()">Guardar configuración</button>
        <button class="btn btn-outline" onclick="limpiarConfigCliente()">Limpiar</button>
      </div>
    `;
  };

  window.guardarConfigCliente = function guardarConfigClienteV2() {
    const codigo = APP.configClienteSel;
    if (!codigo) return;
    pushUndoState('guardar configuración de cliente');
    const client = APP.planificacion.find(item => item.codigo === codigo) || APP.clientes.find(item => item.codigo === codigo) || {};
    const rutaAsignada = text(document.getElementById('cfgRuta').value);
    const diasRecepcion = text(document.getElementById('cfgDias').value).split(',').map(part => parseInt(part, 10)).filter(Boolean);
    APP.clienteConfig[codigo] = {
      ...getClienteConfigV2(codigo, client.nombre),
      codigoCliente: codigo,
      nombreCliente: client.nombre || '',
      rutaAsignada,
      diasRecepcion,
      horarioAlmacen: text(document.getElementById('cfgHorario').value),
      contacto: text(document.getElementById('cfgContacto').value),
      observaciones: text(document.getElementById('cfgNota').value),
      condicionesEspeciales: text(document.getElementById('cfgCondiciones').value),
      configuracionCompleta: !!(rutaAsignada && diasRecepcion.length && text(document.getElementById('cfgHorario').value))
    };
    if (rutaAsignada) ensureRouteConfig(rutaAsignada);
    applyClientRouteToLineItems(codigo, rutaAsignada);
    rebuildDerivedState();
    refreshRouteCostHistoryFromSchedule();
    renderConfigClientes();
    renderRouteCatalog();
    renderCalendario();
    renderRutas();
    renderComercial();
    scheduleAutoSave();
  };

  window.limpiarConfigCliente = function limpiarConfigClienteV2() {
    if (!APP.configClienteSel) return;
    pushUndoState('limpiar configuración de cliente');
    const code = APP.configClienteSel;
    applyClientRouteToLineItems(code, '');
    delete APP.clienteConfig[code];
    rebuildDerivedState();
    renderConfigClientes();
    renderRouteCatalog();
    renderCalendario();
    renderRutas();
    renderComercial();
    scheduleAutoSave();
  };

  window.recalcularPlanConConfig = function recalcularPlanConConfigV2() {
    if (!APP.planLineItems.length) {
      alert('Importa una planificación semanal antes de recalcular.');
      return;
    }
    pushUndoState('recalcular plan con configuración');
    APP.lineItems = APP.planLineItems.map(item => ({ ...item }));
    schedulePlanItems(APP.lineItems, APP.planFecha || fechaToStr(new Date()));
    rebuildDerivedState();
    actualizarDashboard();
    renderCalendario();
    renderRutas();
    renderComercial();
    renderConfigClientes();
    scheduleAutoSave();
  };

  window.exportarConfigClientes = function exportarConfigClientesV2() {
    const rows = Object.values(APP.clienteConfig).map(cfg => ({
      'Código cliente': cfg.codigoCliente || cfg.codigo,
      'Cliente': cfg.nombreCliente || '',
      'Ruta': cfg.rutaAsignada || '',
      'Días de recepción': (cfg.diasRecepcion || []).join(','),
      'Horario almacén': cfg.horarioAlmacen || '',
      'Contacto': cfg.contacto || '',
      'Observaciones': cfg.observaciones || '',
      'Condiciones especiales': cfg.condicionesEspeciales || ''
    }));
    if (!rows.length) return alert('No hay configuraciones para exportar.');
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Configuracion');
    XLSX.writeFile(wb, 'TMS_Configuracion.xlsx');
  };

  window.importarConfigClientes = function importarConfigClientesV2(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (event) {
      try {
        let rows = [];
        if (file.name.toLowerCase().endsWith('.json')) {
          const parsed = JSON.parse(event.target.result);
          rows = Array.isArray(parsed) ? parsed : Object.values(parsed);
        } else {
          const wb = XLSX.read(new Uint8Array(event.target.result), { type: 'array', cellDates: true });
          rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
        }
        let created = 0;
        let updated = 0;
        const errors = [];
        pushUndoState('importar configuración de clientes');
        rows.forEach((row, index) => {
          const codigo = text(pickField(row, ['Código cliente', 'Codigo cliente', 'Código', 'Codigo']));
          if (!codigo) {
            errors.push(index + 1);
            return;
          }
          const exists = !!APP.clienteConfig[codigo];
          APP.clienteConfig[codigo] = {
            ...getClienteConfigV2(codigo, text(pickField(row, ['Cliente', 'Nombre cliente']))),
            codigoCliente: codigo,
            nombreCliente: text(pickField(row, ['Cliente', 'Nombre cliente'])),
            rutaAsignada: text(pickField(row, ['Ruta', 'Zona'])),
            diasRecepcion: parseReceptionDays(pickField(row, ['Días de recepción', 'Dias de recepcion'])),
            horarioAlmacen: text(pickField(row, ['Horario almacén', 'Horario almacen'])),
            contacto: text(pickField(row, ['Contacto'])),
            observaciones: text(pickField(row, ['Observaciones'])),
            condicionesEspeciales: text(pickField(row, ['Condiciones especiales'])),
            configuracionCompleta: true
          };
          APP.clienteConfig[codigo].configuracionCompleta = !!(
            APP.clienteConfig[codigo].rutaAsignada &&
            (APP.clienteConfig[codigo].diasRecepcion || []).length &&
            text(APP.clienteConfig[codigo].horarioAlmacen)
          );
          if (APP.clienteConfig[codigo].rutaAsignada) ensureRouteConfig(APP.clienteConfig[codigo].rutaAsignada);
          applyClientRouteToLineItems(codigo, APP.clienteConfig[codigo].rutaAsignada);
          if (exists) updated += 1;
          else created += 1;
        });
        rebuildDerivedState();
        renderConfigClientes();
        renderRouteCatalog();
        renderCalendario();
        renderRutas();
        renderComercial();
        alert(`Importación completada.\nNuevos: ${created}\nActualizados: ${updated}\nFilas con error: ${errors.length}`);
        scheduleAutoSave();
      } catch (error) {
        console.error('Error importando configuración:', error);
        alert('No se pudo importar la configuración: ' + error.message);
      }
    };
    if (file.name.toLowerCase().endsWith('.json')) reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
  };

  function renderConfigAuxSections() {
    const mount = document.getElementById('configExtrasMount');
    if (!mount) return;
    const days = [
      { value: 1, label: 'Lun' },
      { value: 2, label: 'Mar' },
      { value: 3, label: 'Mié' },
      { value: 4, label: 'Jue' },
      { value: 5, label: 'Vie' },
      { value: 6, label: 'Sáb' }
    ];
    ensureRolePermissions();
    const roleOptions = Object.keys(APP.rolePermissions);
    if (!APP.selectedRolePermission || !APP.rolePermissions[APP.selectedRolePermission]) {
      APP.selectedRolePermission = roleOptions[0] || 'Operaciones';
    }
    mount.innerHTML = `
      <div class="card" style="margin-top:16px;">
        <div class="view-toolbar" style="justify-content:space-between;">
          <div class="card-title" style="margin-bottom:0;">Catálogo de rutas y costos</div>
          <button class="btn btn-primary btn-sm" onclick="guardarConfigRutas()">Guardar cambios de rutas</button>
        </div>
        <div class="config-help-note" style="margin-bottom:12px;">Aquí se crean rutas, días de operación y costos. La asignación diaria de clientes a rutas se hace más rápido en la vista Rutas.</div>
        <div class="route-add-form">
          <input id="newRouteNameInput" class="form-control" placeholder="Nueva ruta / zona">
          <input id="newRouteCostInput" class="form-control" type="number" step="0.01" min="0" placeholder="Coste">
          <button class="btn btn-success btn-sm" onclick="agregarRutaConfig()">Agregar ruta</button>
        </div>
        <div class="table-wrap"><table class="data-table">
          <thead><tr><th>Nombre</th><th>Días operación</th><th>Coste</th><th>Activa</th><th>Observaciones</th></tr></thead>
          <tbody>${APP.routeConfigs.length ? APP.routeConfigs.map((route, index) => `<tr>
            <td><input class="form-control" id="routeName_${index}" value="${route.nombre || route.zona || ''}" style="min-width:150px;"></td>
            <td><div style="display:flex;gap:6px;flex-wrap:wrap;min-width:220px;">${days.map(day => `<label style="font-size:11px;display:flex;align-items:center;gap:3px;"><input type="checkbox" class="routeDay_${index}" value="${day.value}" ${(route.diasOperacion || []).includes(day.value) ? 'checked' : ''}>${day.label}</label>`).join('')}</div></td>
            <td><input class="form-control" id="routeCost_${index}" type="number" step="0.01" value="${route.costeRuta || 0}" style="min-width:95px;"></td>
            <td><input id="routeActive_${index}" type="checkbox" ${route.activa === false ? '' : 'checked'}></td>
            <td><input class="form-control" id="routeObs_${index}" value="${route.observaciones || ''}" style="min-width:180px;"></td>
          </tr>`).join('') : '<tr><td colspan="5">Sin rutas configuradas aún. Se crearán automáticamente al importar o asignar clientes.</td></tr>'}</tbody>
        </table></div>
      </div>
      <div class="card">
        <div class="card-title">Permisos por rol</div>
        <div class="config-help-note" style="margin-bottom:12px;">Primero configura qué puede ver, editar o importar cada rol. Al crear un usuario solo tendrás que asignarle un rol.</div>
        <div class="role-permission-toolbar">
          <label class="form-row" style="margin-bottom:0;">Rol a configurar
            <select id="rolePermissionSelect" class="form-control" onchange="seleccionarRolPermisos(this.value)">
              ${roleOptions.map(role => `<option value="${role}" ${role === APP.selectedRolePermission ? 'selected' : ''}>${role}</option>`).join('')}
            </select>
          </label>
          <label class="form-row" style="margin-bottom:0;">Nuevo rol
            <input id="newRoleNameInput" class="form-control" placeholder="Ej. Supervisor">
          </label>
          <button class="btn btn-outline btn-sm" onclick="agregarRolPermisos()">Crear rol</button>
          <button class="btn btn-primary btn-sm" onclick="guardarPermisosRol()">Guardar permisos del rol</button>
        </div>
        <div class="table-wrap"><table class="data-table role-permission-table">
          <thead><tr><th>Módulo</th><th>Ver</th><th>Editar</th><th>Importar</th></tr></thead>
          <tbody>${renderRolePermissionRows(APP.selectedRolePermission)}</tbody>
        </table></div>
      </div>
      <div class="card">
        <div class="card-title">Usuarios</div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-bottom:14px;">
          <label class="form-row">Nombre
            <input id="userNameInput" class="form-control" placeholder="Nombre completo">
          </label>
          <label class="form-row">Email
            <input id="userEmailInput" class="form-control" type="email" placeholder="correo@empresa.com">
          </label>
          <label class="form-row">Contraseña inicial
            <span class="password-field-wrap">
              <input id="userPasswordInput" class="form-control" type="password" minlength="6" placeholder="Mínimo 6 caracteres" autocomplete="new-password">
              <button type="button" class="btn btn-outline btn-sm" onclick="toggleUserPasswordVisibility()">Ver</button>
            </span>
          </label>
          <label class="form-row">Rol
            <select id="userRoleInput" class="form-control">
              ${roleOptions.map(role => `<option value="${role}">${role}</option>`).join('')}
            </select>
          </label>
        </div>
        <div class="config-help-note" style="margin-bottom:12px;">La contraseña inicial se envía a Supabase Auth al crear el acceso. Para usuarios existentes, usa reset de contraseña; por seguridad Supabase no permite editar contraseñas de otros usuarios desde esta pantalla sin backend administrativo.</div>
        <div style="margin-bottom:12px;">
          <button class="btn btn-primary btn-sm" onclick="crearUsuarioPerfil()">Crear usuario</button>
        </div>
        <div class="table-wrap"><table class="data-table">
          <thead><tr><th>Nombre</th><th>Email</th><th>Rol</th><th>Estado acceso</th><th>Permisos efectivos</th><th>Detalle</th><th>Acciones</th></tr></thead>
          <tbody>${APP.userProfiles.map((profile, index) => {
            const account = getProfileAccountMeta(profile);
            const statusLabel = account.authNeedsConfirmation ? 'Pendiente correo' : account.passwordChangeRequired ? 'Cambio requerido' : 'Activo';
            const statusClass = account.authNeedsConfirmation || account.passwordChangeRequired ? 'badge-warn' : 'badge-ok';
            return `<tr>
            <td>${profile.nombre}</td>
            <td>${profile.email || '—'}</td>
            <td>${profile.rol}</td>
            <td><span class="badge ${statusClass}">${statusLabel}</span></td>
            <td>${profile.rol === 'Admin' ? 'Acceso total' : summarizePermissions(getPermissionsForRole(profile.rol))}</td>
            <td>${renderPermissionChips(getPermissionsForRole(profile.rol))}</td>
            <td>${profile.rol === 'Admin' ? '<span class="badge badge-ok">Base</span>' : `<button class="btn btn-outline btn-sm" onclick="enviarResetPasswordUsuario(${index})">Enviar reset</button> <button class="btn btn-outline btn-sm" onclick="eliminarUsuarioPerfil(${index})">Eliminar</button>`}</td>
          </tr>`;
          }).join('')}</tbody>
        </table></div>
      </div>
      <div class="card">
        <div class="card-title">Costes operativos</div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;">
          <label class="form-row">Coste por solicitud de almacén
            <input id="warehouseCostInput" class="form-control" type="number" step="0.01" value="${APP.warehouseSettings.costoSolicitud || 0}">
          </label>
        </div>
        <div style="margin-top:10px;">
          <button class="btn btn-primary btn-sm" onclick="guardarCostesOperativos()">Guardar costes</button>
        </div>
      </div>
    `;
  }

  window.sincronizarPermisosPorRol = function sincronizarPermisosPorRol() {
    const role = text((document.getElementById('userRoleInput') || {}).value);
    return getPermissionsForRole(role);
  };


  window.toggleUserPasswordVisibility = function toggleUserPasswordVisibility() {
    const input = document.getElementById('userPasswordInput');
    if (!input) return;
    input.type = input.type === 'password' ? 'text' : 'password';
  };

  function renderRolePermissionRows(role) {
    const perms = getPermissionsForRole(role);
    return Object.entries(MODULE_LABELS).map(([module, label]) => {
      const actions = perms[module] || {};
      const canImport = module === 'importar' || module === 'configuracion';
      return `<tr>
        <td><strong>${label}</strong></td>
        <td><input type="checkbox" class="rolePerm" data-module="${module}" data-action="ver" ${actions.ver ? 'checked' : ''}></td>
        <td><input type="checkbox" class="rolePerm" data-module="${module}" data-action="editar" ${actions.editar ? 'checked' : ''}></td>
        <td>${canImport ? `<input type="checkbox" class="rolePerm" data-module="${module}" data-action="importar" ${actions.importar ? 'checked' : ''}>` : '<span style="color:var(--muted);">—</span>'}</td>
      </tr>`;
    }).join('');
  }

  window.seleccionarRolPermisos = function seleccionarRolPermisos(role) {
    APP.selectedRolePermission = role;
    renderConfigAuxSections();
  };

  window.agregarRolPermisos = function agregarRolPermisos() {
    const input = document.getElementById('newRoleNameInput');
    const role = text(input && input.value);
    if (!role) return alert('Escribe el nombre del nuevo rol.');
    if (APP.rolePermissions[role]) return alert('Ese rol ya existe.');
    pushUndoState('crear rol');
    APP.rolePermissions[role] = buildPermissions('read_only', role);
    APP.selectedRolePermission = role;
    renderConfigAuxSections();
    scheduleAutoSave();
  };

  window.guardarPermisosRol = function guardarPermisosRol() {
    if (!requireAdminUser('Guardar permisos')) return;
    const role = text((document.getElementById('rolePermissionSelect') || {}).value) || APP.selectedRolePermission;
    if (!role) return;
    pushUndoState('guardar permisos de rol');
    const perms = {};
    Object.keys(MODULE_LABELS).forEach(module => {
      perms[module] = { ver: false, editar: false, importar: false };
    });
    document.querySelectorAll('.rolePerm').forEach(input => {
      const module = input.dataset.module;
      const action = input.dataset.action;
      if (!module || !action) return;
      perms[module][action] = !!input.checked;
    });
    Object.keys(perms).forEach(module => {
      if (perms[module].editar || perms[module].importar) perms[module].ver = true;
    });
    APP.rolePermissions[role] = normalizePermissions(perms);
    APP.userProfiles.forEach(profile => {
      if (profile.rol === role) profile.permisosPorModulo = getPermissionsForRole(role);
    });
    renderConfigAuxSections();
    scheduleAutoSave();
    alert('Permisos del rol actualizados.');
  };

  function summarizePermissions(perms) {
    const entries = Object.entries(perms || {}).filter(([module]) => MODULE_LABELS[module]);
    if (!entries.length) return 'Sin permisos definidos';
    const editable = entries.filter(([, actions]) => actions && (actions.editar || actions.importar)).map(([module]) => MODULE_LABELS[module] || module);
    if (editable.length === entries.length) return 'Ver y editar todo';
    if (!editable.length) return 'Solo lectura';
    return 'Edita: ' + editable.join(', ');
  }

  function renderPermissionChips(perms) {
    const entries = Object.entries(perms || {}).filter(([module]) => MODULE_LABELS[module]);
    if (!entries.length) return '<span class="permission-chip muted">Sin permisos</span>';
    return entries.map(([module, actions]) => {
      const label = MODULE_LABELS[module] || module;
      const canImport = !!(actions && actions.importar);
      const canEdit = !!(actions && actions.editar);
      const canView = !!(actions && actions.ver);
      const status = canImport ? 'Importa' : (canEdit ? 'Edita' : (canView ? 'Ver' : 'Sin acceso'));
      const cls = canImport || canEdit ? 'edit' : (canView ? 'view' : 'off');
      return `<span class="permission-chip ${cls}">${label}: ${status}</span>`;
    }).join('');
  }

  window.agregarRutaConfig = function agregarRutaConfig() {
    const nameInput = document.getElementById('newRouteNameInput');
    const costInput = document.getElementById('newRouteCostInput');
    const routeName = text(nameInput && nameInput.value);
    if (!routeName) return alert('Escribe el nombre de la ruta.');
    if (routeExistsByName(routeName)) return alert('Esa ruta ya existe en Configuración.');
    pushUndoState('agregar ruta');
    APP.routeConfigs.push({
      id: 'route-' + normKey(routeName),
      nombre: routeName,
      zona: routeName,
      diasOperacion: [],
      costeRuta: num(costInput && costInput.value),
      activa: true,
      observaciones: 'Creada manualmente'
    });
    if (nameInput) nameInput.value = '';
    if (costInput) costInput.value = '';
    renderConfigAuxSections();
    renderConfigEditor();
    initRouteFilterOptions();
    scheduleAutoSave();
  };

  window.guardarConfigRutas = function guardarConfigRutas() {
    pushUndoState('guardar configuración de rutas');
    APP.routeConfigs = APP.routeConfigs.map((route, index) => {
      const nombre = text((document.getElementById('routeName_' + index) || {}).value) || route.nombre || route.zona;
      const diasOperacion = [...document.querySelectorAll('.routeDay_' + index + ':checked')].map(input => parseInt(input.value, 10)).filter(Boolean);
      return {
        ...route,
        nombre,
        zona: nombre,
        diasOperacion,
        costeRuta: num((document.getElementById('routeCost_' + index) || {}).value),
        activa: !!((document.getElementById('routeActive_' + index) || {}).checked),
        observaciones: text((document.getElementById('routeObs_' + index) || {}).value)
      };
    });
    APP.lineItems.forEach(item => {
      const cfg = getClienteConfigV2(item.clienteId, item.clienteNombre);
      if (cfg.rutaAsignada) {
        item.rutaNombre = cfg.rutaAsignada;
        item.zona = cfg.rutaAsignada;
      }
    });
    rebuildDerivedState();
    renderConfigClientes();
    renderRouteCatalog();
    renderCalendario();
    renderRutas();
    scheduleAutoSave();
    alert('Rutas actualizadas. Si una ruta se usa fuera de sus días de operación, se marcará en rojo en el calendario.');
  };

  window.guardarCostesOperativos = function guardarCostesOperativos() {
    pushUndoState('guardar costes operativos');
    APP.warehouseSettings.costoSolicitud = num(document.getElementById('warehouseCostInput').value);
    registerSolicitudesSnapshot(APP.importFiles.solicitudes);
    scheduleAutoSave();
    renderConfigAuxSections();
  };

  function buildPermissions(scope, role) {
    const all = {
      importar: { ver: true, editar: true, importar: true },
      dashboard: { ver: true, editar: true },
      calendario: { ver: true, editar: true },
      reportes: { ver: true, editar: true },
      comercial: { ver: true, editar: true },
      rutas: { ver: true, editar: true },
      importaciones: { ver: true, editar: true, importar: true },
      prioridades: { ver: role === 'Admin' || scope === 'full', editar: role === 'Admin' || scope === 'full', importar: false },
      configuracion: { ver: true, editar: true, importar: true },
      solicitudesAlmacen: { ver: true, editar: true },
      almacen: { ver: true, editar: true }
    };
    if (role === 'Admin' || scope === 'full') return all;
    if (scope === 'read_only') {
      return Object.fromEntries(Object.keys(all).map(key => [
        key,
        key === 'prioridades' ? { ver: false, editar: false, importar: false } : { ver: true, editar: false, importar: false }
      ]));
    }
    if (scope === 'commercial') {
      return {
        importar: { ver: false, editar: false, importar: false },
        dashboard: { ver: false, editar: false },
        calendario: { ver: false, editar: false },
        reportes: { ver: false, editar: false },
        comercial: { ver: true, editar: false },
        rutas: { ver: false, editar: false },
        importaciones: { ver: false, editar: false, importar: false },
        prioridades: { ver: false, editar: false, importar: false },
        configuracion: { ver: false, editar: false, importar: false },
        solicitudesAlmacen: { ver: false, editar: false },
        almacen: { ver: false, editar: false }
      };
    }
    if (scope === 'warehouse') {
      return {
        importar: { ver: false, editar: false, importar: false },
        dashboard: { ver: true, editar: false },
        calendario: { ver: true, editar: false },
        reportes: { ver: true, editar: false },
        comercial: { ver: false, editar: false },
        rutas: { ver: false, editar: false },
        importaciones: { ver: false, editar: false, importar: false },
        prioridades: { ver: false, editar: false, importar: false },
        configuracion: { ver: false, editar: false, importar: false },
        solicitudesAlmacen: { ver: true, editar: true },
        almacen: { ver: true, editar: true }
      };
    }
    return {
      importar: { ver: true, editar: true, importar: true },
      dashboard: { ver: true, editar: false },
      calendario: { ver: true, editar: true },
      reportes: { ver: true, editar: true },
      comercial: { ver: true, editar: false },
      rutas: { ver: true, editar: true },
      importaciones: { ver: true, editar: true, importar: true },
      prioridades: { ver: false, editar: false, importar: false },
      configuracion: { ver: false, editar: false, importar: false },
      solicitudesAlmacen: { ver: true, editar: true },
      almacen: { ver: true, editar: true }
    };
  }

  async function createAuthUserForProfile(email, password, nombre, rol) {
    const client = getSupabaseClient();
    if (!client || !client.functions) {
      throw new Error('No hay función administrativa disponible para crear usuarios. Despliega supabase/functions/admin-create-user.');
    }
    const { data, error } = await client.functions.invoke('admin-create-user', {
      body: {
        email,
        password,
        nombre,
        rol,
        permisosPorModulo: getPermissionsForRole(rol)
      }
    });
    if (error) {
      let detail = error.message || String(error);
      const response = error.context || error.response;
      if (response && typeof response.clone === 'function') {
        try {
          const body = await response.clone().json();
          detail = body.error || body.message || detail;
        } catch (_) {
          try { detail = await response.clone().text() || detail; } catch (_) {}
        }
      }
      throw new Error(detail);
    }
    if (data && data.error) throw new Error(data.error);
    return {
      userId: (data && data.userId) || '',
      needsConfirmation: !!(data && data.needsConfirmation)
    };
  }

  window.crearUsuarioPerfil = async function crearUsuarioPerfil() {
    if (!requireAdminUser('Crear usuarios')) return;
    const nombre = text(document.getElementById('userNameInput').value);
    const email = text(document.getElementById('userEmailInput').value);
    const password = String((document.getElementById('userPasswordInput') || {}).value || '');
    const rol = text(document.getElementById('userRoleInput').value) || 'Operaciones';
    if (!nombre || !email || !password) {
      alert('Completa nombre, email y contraseña para crear el usuario.');
      return;
    }
    if (!isValidLoginEmail(email)) {
      alert('Escribe un correo completo y válido. Ejemplo: usuario@hbyalvarez.com');
      return;
    }
    if (password.length < 6) {
      alert('La contraseña debe tener al menos 6 caracteres.');
      return;
    }
    if (APP.userProfiles.some(profile => text(profile.email).toLowerCase() === email.toLowerCase())) {
      alert('Ya existe un usuario con ese email.');
      return;
    }
    let authUserId = '';
    let authNeedsConfirmation = false;
    try {
      const authResult = await createAuthUserForProfile(email, password, nombre, rol);
      authUserId = authResult.userId || '';
      authNeedsConfirmation = !!authResult.needsConfirmation;
    } catch (error) {
      console.error('No se pudo crear el acceso del usuario:', error);
      alert('No se pudo crear el acceso real en Supabase Auth. No se guardó un perfil incompleto. Detalle: ' + (error.message || error));
      return;
    }
    pushUndoState('crear usuario');
    APP.userProfiles.push({
      userId: authUserId || 'user-' + Date.now(),
      authUserId,
      nombre,
      email,
      rol,
      permisosPorModulo: getPermissionsForRole(rol),
      passwordConfigured: !authNeedsConfirmation,
      authNeedsConfirmation,
      passwordChangeRequired: !authNeedsConfirmation,
      accountStatus: authNeedsConfirmation ? 'Pendiente de confirmar correo' : 'Contraseña temporal'
    });
    const passwordInput = document.getElementById('userPasswordInput');
    if (passwordInput) passwordInput.value = '';
    renderConfigAuxSections();
    scheduleAutoSave();
    alert(authNeedsConfirmation
      ? 'Usuario creado. Debe confirmar el correo antes de acceder con la contraseña temporal.'
      : 'Usuario creado. Al primer acceso tendrá que cambiar la contraseña temporal.');
  };

  window.enviarResetPasswordUsuario = async function enviarResetPasswordUsuario(index) {
    if (!requireAdminUser('Enviar reset de contraseña')) return;
    const profile = APP.userProfiles[index];
    if (!profile || !profile.email) return;
    const client = getSupabaseClient();
    if (!client || !client.auth) return alert('No hay conexión de autenticación disponible.');
    try {
      const { error } = await client.auth.resetPasswordForEmail(profile.email, {
        redirectTo: window.location.origin + window.location.pathname
      });
      if (error) throw error;
      alert('Se envió un correo de recuperación a ' + profile.email + '.');
    } catch (error) {
      console.error('No se pudo enviar recuperación:', error);
      alert('No se pudo enviar el correo de recuperación: ' + (error.message || error));
    }
  };

  window.eliminarUsuarioPerfil = function eliminarUsuarioPerfil(index) {
    if (!requireAdminUser('Eliminar usuarios')) return;
    const profile = APP.userProfiles[index];
    if (!profile || profile.rol === 'Admin') return;
    pushUndoState('eliminar usuario');
    APP.userProfiles.splice(index, 1);
    renderConfigAuxSections();
    scheduleAutoSave();
  };

  function buildRouteCostRowsFromSchedule(items) {
    const grouped = new Map();
    items.forEach(item => {
      const fecha = item.fechaPlanificada || '';
      const routeName = item.rutaNombre || item.zona || inferRouteName(item.clienteNombre, item.clienteId);
      const camion = item.camionAsignado || chooseTruckForItem(item);
      if (!fecha || !routeName) return;
      const key = [fecha, routeName, camion].join('|');
      if (!grouped.has(key)) grouped.set(key, { fecha, rutaNombre: routeName, camion, items: [] });
      grouped.get(key).items.push(item);
    });
    return [...grouped.values()].map(row => {
      const routeCfg = getRouteConfigByName(row.rutaNombre);
      const summary = calcProgressSummary(row.items);
      const evidence = getRouteEvidenceFromItems(row.items);
      const importePlanificado = row.items.reduce((sum, item) => sum + getOperationalAmount(item), 0);
      const importeReal = row.items.reduce((sum, item) => {
        const requested = item.cantidadSolicitada || 0;
        const deliveredRatio = requested > 0 ? Math.min((item.cantidadFacturada || 0) / requested, 1) : 0;
        return sum + getOperationalAmount(item) * deliveredRatio;
      }, 0);
      const coste = routeCfg ? num(routeCfg.costeRuta) : 0;
      return {
        fecha: row.fecha,
        rutaId: (routeCfg && routeCfg.id) || row.rutaNombre,
        rutaNombre: row.rutaNombre,
        camion: row.camion,
        coste,
        costePendienteConfig: !routeCfg || coste <= 0,
        costeAlerta: !routeCfg ? 'Ruta no existe en Configuración' : (coste <= 0 ? 'Coste de ruta en cero' : ''),
        fuenteControlDiario: APP.importFiles.control || '',
        pedidosAsociados: uniqueTexts(row.items.map(item => item.pedidoCliente)),
        cumplimiento: summary.pct,
        precintoDespacho: evidence.precintoDespacho,
        comentarioRuta: evidence.comentarioRuta,
        choferRuta: evidence.choferRuta,
        fotoCamion: evidence.fotoCamion,
        fotoCamionAdjunta: !!evidence.fotoCamion,
        fotoCamionNombre: evidence.fotoCamionNombre,
        importePlanificado,
        importeReal,
        diferencia: importePlanificado - importeReal
      };
    });
  }

  function getRouteEvidenceFromItems(items) {
    const first = (items || []).find(item => item.precintoDespacho || item.comentarioRuta || item.choferRuta || item.fotoCamion || item.fotoCamionNombre) || {};
    return {
      precintoDespacho: first.precintoDespacho || '',
      comentarioRuta: first.comentarioRuta || '',
      choferRuta: first.choferRuta || '',
      fotoCamion: first.fotoCamion || '',
      fotoCamionNombre: first.fotoCamionNombre || '',
      fechaCierreRuta: first.fechaCierreRuta || ''
    };
  }

  function normalizeRouteCostRow(row) {
    const routeCfg = getRouteConfigByName(row.rutaNombre || row.ruta || '') || null;
    const importePlanificado = num(row.importePlanificado);
    const importeReal = num(row.importeReal);
    const coste = routeCfg ? num(routeCfg.costeRuta) : num(row.coste);
    return {
      ...row,
      rutaId: (routeCfg && routeCfg.id) || row.rutaId || row.rutaNombre || '',
      camion: row.camion || '',
      coste,
      costePendienteConfig: !routeCfg || coste <= 0,
      costeAlerta: !routeCfg ? 'Ruta no existe en Configuración' : (coste <= 0 ? 'Coste de ruta en cero' : ''),
      importePlanificado,
      importeReal,
      diferencia: importePlanificado - importeReal
    };
  }

  function refreshRouteCostHistoryFromSchedule() {
    const rows = buildRouteCostRowsFromSchedule(APP.lineItems || []);
    const keys = new Set(rows.map(row => [row.fecha, row.rutaNombre, row.camion || ''].join('|')));
    const routeDateKeys = new Set(rows.map(row => [row.fecha, row.rutaNombre].join('|')));
    APP.routeCostHistory = [
      ...(APP.routeCostHistory || []).map(normalizeRouteCostRow).filter(row => {
        const exactKey = [row.fecha, row.rutaNombre, row.camion || ''].join('|');
        const legacyKey = [row.fecha, row.rutaNombre].join('|');
        return !keys.has(exactKey) && !routeDateKeys.has(legacyKey);
      }),
      ...rows
    ];
    return rows;
  }

  function buildWarehouseRowsForReport() {
    if (!(APP.solicitudesAlmacen || []).length) return APP.solicitudesHistory || [];
    return (APP.solicitudesAlmacen || []).map(item => ({
      fecha: item.fechaLiberacion || item.fechaEnvio || fechaToStr(new Date()),
      clienteId: item.codigo,
      clienteNombre: item.cliente,
      pedidoCliente: item.pedidoCliente || '',
      articulo: item.articulo || '',
      descripcionArticulo: item.descripcionArticulo || '',
      cantidadSolicitada: item.cantidadSolicitada || 0,
      cantidadSolicitadaOriginal: item.cantidadSolicitadaOriginal || item.cantidadSolicitada || 0,
      cantidadPendienteServir: item.cantidadPendienteServir || item.cantidadSolicitada || 0,
      cantidadPendienteSinStock: item.cantidadPendienteSinStock || 0,
      cantidadCompletada: item.cantidadCompletada || 0,
      cantidadPendiente: item.cantidadPendiente || 0,
      estado: item.estado || '',
      costeSolicitud: num(item.costeSolicitud || APP.warehouseSettings.costoSolicitud || 0),
      fuente: APP.importFiles.solicitudesControl || APP.importFiles.solicitudesPlan || APP.importFiles.solicitudes || '',
      tipo: item.tipoSolicitudArchivo || ''
    }));
  }

  function summarizeRouteCostIssues(rows) {
    const grouped = new Map();
    rows.forEach(row => {
      const key = [row.rutaNombre, row.costeAlerta].join('|');
      if (!grouped.has(key)) grouped.set(key, { ...row, fechas: [], count: 0 });
      const current = grouped.get(key);
      current.count += 1;
      current.fechas = uniqueTexts([...current.fechas, row.fecha]);
    });
    return [...grouped.values()];
  }

  window.generarReporte = function generarReporteV2() {
    const output = document.getElementById('reporteOutput');
    if (!output) return;
    const clientQuery = text((document.getElementById('repSearch') || {}).value).toLowerCase();
    const routeFilter = text((document.getElementById('repRouteFilter') || {}).value);
    const start = text((document.getElementById('repDateStart') || {}).value);
    const end = text((document.getElementById('repDateEnd') || {}).value);

    const items = APP.lineItems.filter(item => {
      if (clientQuery && ![item.clienteId, item.clienteNombre, item.pedidoCliente].some(value => text(value).toLowerCase().includes(clientQuery))) return false;
      if (routeFilter && text(item.rutaNombre || item.zona) !== routeFilter) return false;
      const date = item.fechaPlanificada ? fechaStrToISO(item.fechaPlanificada) : '';
      if (start && date && date < start) return false;
      if (end && date && date > end) return false;
      return true;
    });

    refreshRouteCostHistoryFromSchedule();
    const routeHistory = (APP.routeCostHistory || []).filter(row => {
      const iso = fechaStrToISO(row.fecha);
      if (routeFilter && row.rutaNombre !== routeFilter) return false;
      if (start && iso < start) return false;
      if (end && iso > end) return false;
      return true;
    });
    const warehouseHistory = buildWarehouseRowsForReport().filter(row => {
      const iso = fechaStrToISO(row.fecha);
      if (clientQuery && ![row.clienteId, row.clienteNombre].some(value => text(value).toLowerCase().includes(clientQuery))) return false;
      if (start && iso < start) return false;
      if (end && iso > end) return false;
      return true;
    });

    const progress = calcProgressSummary(items);
    const routeCostTotal = routeHistory.reduce((sum, row) => sum + num(row.coste), 0);
    const warehouseCostTotal = warehouseHistory.reduce((sum, row) => sum + num(row.costeSolicitud), 0);
    const routeCostIssues = summarizeRouteCostIssues(routeHistory.filter(row => row.costePendienteConfig));
    const routeEvidenceRows = routeHistory.filter(row => row.precintoDespacho || row.comentarioRuta || row.choferRuta || row.fotoCamion);
    const routeEvidenceByDate = routeEvidenceRows.reduce((map, row) => {
      if (!map.has(row.fecha)) map.set(row.fecha, []);
      map.get(row.fecha).push(row);
      return map;
    }, new Map());
    const warehouseSummary = {
      solicitudes: warehouseHistory.length,
      clientes: new Set(warehouseHistory.map(row => row.clienteId)).size,
      pendientes: warehouseHistory.reduce((sum, row) => sum + num(row.cantidadPendiente), 0),
      completadas: warehouseHistory.reduce((sum, row) => sum + num(row.cantidadCompletada), 0)
    };
    output.innerHTML = `
      <div class="card">
        <div class="card-title">Resumen operativo</div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;">
          <div class="kpi-card"><div class="kpi-label">Clientes</div><div class="kpi-val">${new Set(items.map(item => item.clienteId)).size}</div></div>
          <div class="kpi-card"><div class="kpi-label">Pedidos / líneas</div><div class="kpi-val">${items.length}</div></div>
          <div class="kpi-card"><div class="kpi-label">% cumplimiento</div><div class="kpi-val">${progress.pct}%</div></div>
          <div class="kpi-card"><div class="kpi-label">Gasto rutas</div><div class="kpi-val">${formatMonto(routeCostTotal)}</div></div>
          <div class="kpi-card"><div class="kpi-label">Gasto almacén</div><div class="kpi-val">${formatMonto(warehouseCostTotal)}</div></div>
        </div>
      </div>
      ${routeCostIssues.length ? `<div class="card route-cost-alert">
        <div class="card-title">Rutas con coste pendiente</div>
        <div style="font-size:12px;color:var(--muted);line-height:1.45;margin-bottom:10px;">Si aparece <strong>BAVARO</strong>, normalmente significa que el sistema infirió esa ruta desde el nombre del cliente, pero no existe una ruta configurada con ese nombre exacto y coste. Asígnala en Configuración a la ruta correcta, por ejemplo PUNTA CANA / BAVARO, BARCELO- BAVARO o la que aplique.</div>
        <div class="route-cost-alert-list">
          ${routeCostIssues.map(row => `<div class="route-cost-alert-item">
            <strong>${row.rutaNombre}</strong>
            <span>${row.count} día(s): ${row.fechas.join(', ')} · ${row.costeAlerta}</span>
          </div>`).join('')}
        </div>
      </div>` : ''}
      ${routeEvidenceRows.length ? `<div class="card">
        <div class="card-title">Evidencias de despacho por día</div>
        <div class="route-evidence-report">
          ${[...routeEvidenceByDate.entries()].sort((a, b) => fechaToDate(a[0]) - fechaToDate(b[0])).map(([date, rows]) => `<div class="route-evidence-day">
            <div class="route-evidence-day-title"><strong>${fechaLabel(date)} ${date}</strong><span>${rows.length} evidencia(s)</span></div>
            <div class="route-evidence-grid">
              ${rows.map(row => `<div class="route-evidence-card">
                ${row.fotoCamion ? `<img src="${row.fotoCamion}" alt="Evidencia ${row.rutaNombre}">` : '<div class="route-evidence-placeholder">Sin foto</div>'}
                <div class="route-evidence-info">
                  <strong>${row.rutaNombre}</strong>
                  <span>${getTruckLabel(row.camion)} · Coste ${formatMonto(row.coste)}</span>
                  <span>Precinto: ${row.precintoDespacho || '—'}</span>
                  <span>Chofer/recibido: ${row.choferRuta || '—'}</span>
                  <span>${row.comentarioRuta || 'Sin observación'}</span>
                </div>
              </div>`).join('')}
            </div>
          </div>`).join('')}
        </div>
      </div>` : ''}
      <div class="card">
        <div class="card-title">Gasto teórico / real de rutas</div>
        <div style="font-size:12px;color:var(--muted);line-height:1.45;margin-bottom:10px;">Planificado usa el monto operativo: confirmado no entregado + entregado no facturado. Real estima la porción entregada usando ese mismo monto y el avance facturado por línea.</div>
        ${routeHistory.length ? `<div class="table-wrap"><table class="data-table">
          <thead><tr><th>Fecha</th><th>Ruta</th><th>Camión</th><th>Coste</th><th>Estado coste</th><th>Cumplimiento</th><th>Precinto</th><th>Chofer/recibido</th><th>Comentario</th><th>Foto</th><th>Planificado</th><th>Real</th><th>Diferencia</th></tr></thead>
          <tbody>${routeHistory.map(row => `<tr>
            <td>${row.fecha}</td>
            <td>${row.rutaNombre}</td>
            <td>${getTruckLabel(row.camion)}</td>
            <td>${formatMonto(row.coste)}</td>
            <td>${row.costePendienteConfig ? `<span class="badge badge-warn">${row.costeAlerta}</span>` : '<span class="badge badge-ok">Configurado</span>'}</td>
            <td>${row.cumplimiento}%</td>
            <td>${row.precintoDespacho || '—'}</td>
            <td>${row.choferRuta || '—'}</td>
            <td>${row.comentarioRuta || '—'}</td>
            <td>${row.fotoCamion ? `<img src="${row.fotoCamion}" alt="Foto ${row.rutaNombre}" class="route-evidence-thumb">` : (row.fotoCamionAdjunta ? (row.fotoCamionNombre || 'Adjunta') : '—')}</td>
            <td>${formatMonto(row.importePlanificado)}</td>
            <td>${formatMonto(row.importeReal)}</td>
            <td>${formatMonto(row.diferencia)}</td>
          </tr>`).join('')}</tbody>
        </table></div>` : '<div class="empty-state" style="padding:20px;"><p>No hay histórico de rutas para ese rango.</p></div>'}
      </div>
      <div class="card">
        <div class="card-title">Solicitudes a almacén por fecha de liberación</div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin-bottom:12px;">
          <div class="kpi-card"><div class="kpi-label">Solicitudes / líneas</div><div class="kpi-val">${warehouseSummary.solicitudes}</div></div>
          <div class="kpi-card"><div class="kpi-label">Clientes</div><div class="kpi-val">${warehouseSummary.clientes}</div></div>
          <div class="kpi-card"><div class="kpi-label">Pendientes</div><div class="kpi-val">${warehouseSummary.pendientes}</div></div>
          <div class="kpi-card"><div class="kpi-label">Completadas</div><div class="kpi-val">${warehouseSummary.completadas}</div></div>
        </div>
        ${warehouseHistory.length ? `<div class="table-wrap"><table class="data-table">
          <thead><tr><th>Fecha liberación</th><th>Cliente</th><th>Artículo</th><th>Solicitada</th><th>Completada</th><th>Pendiente</th><th>Estado</th><th>Coste</th></tr></thead>
          <tbody>${warehouseHistory.map(row => `<tr>
            <td>${row.fecha}</td>
            <td>${row.clienteNombre || row.clienteId}</td>
            <td>${row.articulo || row.descripcionArticulo || '—'}</td>
            <td>${row.cantidadSolicitada}</td>
            <td>${row.cantidadCompletada}</td>
            <td>${row.cantidadPendiente}</td>
            <td>${row.estado}</td>
            <td>${formatMonto(row.costeSolicitud)}</td>
          </tr>`).join('')}</tbody>
        </table></div>` : '<div class="empty-state" style="padding:20px;"><p>No hay histórico de solicitudes para ese rango.</p></div>'}
      </div>
    `;
  };

  window.exportarReporteExcel = function exportarReporteExcelV2() {
    const wb = XLSX.utils.book_new();
    const wsRoutes = XLSX.utils.json_to_sheet((APP.routeCostHistory || []).map(row => ({
      Fecha: row.fecha,
      Ruta: row.rutaNombre,
      Camion: getTruckLabel(row.camion),
      Coste: row.coste,
      EstadoCoste: row.costeAlerta || 'Configurado',
      Cumplimiento: row.cumplimiento,
      Precinto: row.precintoDespacho || '',
      ChoferRecibidoPor: row.choferRuta || '',
      ComentarioRuta: row.comentarioRuta || '',
      FotoCamion: row.fotoCamionAdjunta ? (row.fotoCamionNombre || 'Adjunta') : '',
      Planificado: row.importePlanificado,
      Real: row.importeReal,
      Diferencia: row.diferencia
    })));
    const wsWarehouse = XLSX.utils.json_to_sheet(APP.solicitudesHistory);
    XLSX.utils.book_append_sheet(wb, wsRoutes, 'Rutas');
    XLSX.utils.book_append_sheet(wb, wsWarehouse, 'Almacen');
    XLSX.writeFile(wb, 'TMS_Reportes.xlsx');
  };

  window.exportarSolicitudesAlmacenExcel = function exportarSolicitudesAlmacenExcelV2() {
    const rows = getWarehouseRowsForView(APP.solicitudesVista || 'programadas');
    if (!rows.length) return alert('No hay solicitudes para exportar.');
    const ws = XLSX.utils.json_to_sheet(rows.map(item => ({
      FechaProgramada: item.fechaProgramada,
      Codigo: item.codigo,
      Cliente: item.clienteNombreVista,
      Solicitud: item.solicitud,
      Pedido: item.pedidoCliente,
      Articulo: item.articulo,
      Descripcion: item.descripcionArticulo,
      Solicitada: item.cantidadSolicitada,
      Completada: item.cantidadCompletada,
      Pendiente: item.cantidadPendiente,
      Estado: item.estadoVista,
      Monto: item.montoProgramado,
      MotivoNoPreparada: item.motivoNoPreparada || ''
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Solicitudes');
    XLSX.writeFile(wb, 'TMS_Solicitudes_Almacen.xlsx');
  };

  window.exportarExcelRutas = function exportarExcelRutasV2() {
    if (!APP.rutaFecha || !APP.rutas[APP.rutaFecha]) return alert('Selecciona una fecha para exportar.');
    const wb = XLSX.utils.book_new();
    const summaryRows = [];
    getTruckOptions().forEach(camion => {
      const sourceRows = APP.rutas[APP.rutaFecha][camion] || [];
      const rows = sourceRows.map(group => ({
        'Cliente': group.nombre,
        'Código': group.codigo,
        'Pedido': group.pedidoCliente,
        'Ruta': group.rutaNombre,
        'Líneas': group.items.length,
        'Solicitada': group.cantidadSolicitada,
        'Facturada': group.cantidadFacturada,
        'Pendiente': group.cantidadPendiente,
        'Precinto': group.precintoDespacho || '',
        'Chofer / recibido por': group.choferRuta || '',
        'Comentario cierre': group.comentarioRuta || '',
        'Foto camión': group.fotoCamionNombre || (group.fotoCamion ? 'Adjunta' : '')
      }));
      if (!rows.length) return;
      summaryRows.push({
        'Fecha': APP.rutaFecha,
        'Carril': getTruckLabel(camion),
        'Pedidos': sourceRows.length,
        'Líneas': sourceRows.reduce((sum, group) => sum + group.items.length, 0),
        'Pendiente': sourceRows.reduce((sum, group) => sum + (group.cantidadPendiente || 0), 0)
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), camion.replace(/\s+/g, ''));
    });
    if (summaryRows.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), 'Resumen');
    }
    XLSX.writeFile(wb, 'TMS_Rutas_' + APP.rutaFecha.replace(/\//g, '-') + '.xlsx');
  };

  window.exportarAlmacenExcel = function exportarAlmacenExcelV2() {
    const ws = XLSX.utils.json_to_sheet(APP.solicitudesHistory.length ? APP.solicitudesHistory : APP.solicitudesAlmacen);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Almacen');
    XLSX.writeFile(wb, 'TMS_Almacen.xlsx');
  };

  window.exportarPlanVsRealExcel = function exportarPlanVsRealExcelV2() {
    const data = APP.lineItems.map(item => ({
      'Cliente': item.clienteNombre,
      'Pedido': item.pedidoCliente,
      'Línea': item.lineaPedidoCliente,
      'Artículo': item.articulo,
      'Solicitada': item.cantidadSolicitada,
      'Facturada': item.cantidadFacturada,
      'Pendiente': item.cantidadPendiente,
      'Estado': deriveItemStatus(item),
      'Fecha': item.fechaPlanificada,
      'Camión': item.camionAsignado
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PlanVsReal');
    XLSX.writeFile(wb, 'TMS_Plan_vs_Real.xlsx');
  };

  const originalActualizarDashboard = window.actualizarDashboard;
  window.actualizarDashboard = function actualizarDashboardV2() {
    if (typeof originalActualizarDashboard === 'function') originalActualizarDashboard();
    const clients = APP.clientes || [];
    const confirmedTotal = clients.reduce((sum, client) => sum + getClientConfirmedUndelivered(client), 0);
    const deliveredPendingTotal = clients.reduce((sum, client) => sum + getClientDeliveredNotInvoiced(client), 0);
    const operationalTotal = clients.reduce((sum, client) => sum + getClientOperationalAmount(client), 0);
    const confirmedEl = document.getElementById('dashConfirmado');
    const pendingEl = document.getElementById('dashPendiente');
    const totalEl = document.getElementById('dashTotal');
    if (confirmedEl) confirmedEl.textContent = formatMonto(confirmedTotal);
    if (pendingEl) pendingEl.textContent = formatMonto(deliveredPendingTotal);
    if (totalEl) totalEl.textContent = formatMonto(operationalTotal);
    renderControlHistoryPanel();
    const topDiv = document.getElementById('dashTopClientes');
    if (topDiv) {
      const top10 = [...clients]
        .map(client => ({ ...client, montoOperativo: getClientOperationalAmount(client) }))
        .filter(client => client.montoOperativo > 0)
        .sort((a, b) => b.montoOperativo - a.montoOperativo)
        .slice(0, 10);
      const max = top10[0] ? top10[0].montoOperativo : 1;
      topDiv.innerHTML = top10.length ? top10.map(client => {
        const pct = Math.max(2, Math.round(client.montoOperativo / max * 100));
        const camColor = client.camion === 'CAMION 1' ? 'var(--c1)' : client.camion === 'CAMION 2' ? 'var(--c2)' : 'var(--calmacen)';
        return `<div class="prog-row">
          <div class="prog-label" title="${client.nombre}">${client.nombre}</div>
          <div class="prog-bar-bg"><div class="prog-bar-fill" style="width:${pct}%;background:${camColor};"></div></div>
          <div class="prog-val">${formatMonto(client.montoOperativo)}</div>
        </div>`;
      }).join('') : '<div class="empty-state" style="padding:16px;"><p>Sin montos pendientes para mostrar.</p></div>';
    }
    const empty = !APP.lineItems.length;
    const dashContent = document.getElementById('dashContent');
    if (empty && dashContent) {
      dashContent.innerHTML = '<div class="empty-state"><div class="empty-icon">📊</div><p>No hay datos cargados todavía.</p></div>';
    }
  };


  async function enforcePasswordChangeIfNeeded() {
    const profile = APP.currentUserProfile;
    if (!profile || !getProfileAccountMeta(profile).passwordChangeRequired) return;
    const client = getSupabaseClient();
    if (!client || !client.auth) return;
    let nextPassword = '';
    let confirmPassword = '';
    while (true) {
      nextPassword = window.prompt('Debes cambiar tu contraseña temporal. Escribe una nueva contraseña de al menos 6 caracteres:') || '';
      if (!nextPassword) {
        await client.auth.signOut();
        alert('Debes cambiar la contraseña para acceder al sistema.');
        location.reload();
        return;
      }
      confirmPassword = window.prompt('Confirma la nueva contraseña:') || '';
      if (nextPassword.length < 6) alert('La contraseña debe tener al menos 6 caracteres.');
      else if (nextPassword !== confirmPassword) alert('Las contraseñas no coinciden.');
      else break;
    }
    const { error } = await client.auth.updateUser({ password: nextPassword });
    if (error) {
      alert('No se pudo cambiar la contraseña: ' + error.message);
      await client.auth.signOut();
      location.reload();
      return;
    }
    profile.passwordChangeRequired = false;
    profile.passwordConfigured = true;
    profile.authNeedsConfirmation = false;
    profile.accountStatus = 'Activo';
    alert('Contraseña actualizada. Ya puedes usar el sistema.');
    await window.guardarEnSupabase();
  }

  window.cargarDesdeSupabase = async function cargarDesdeSupabaseV2() {
    try {
      const client = getSupabaseClient();
      const user = await getAuthenticatedUser();
      ensureAdminProfile();
      if (!client || !user) {
        if (loadLocalSnapshot() || loadEmbeddedDemoState()) {
          APP.currentUserProfile = {
            userId: 'local-readonly',
            nombre: 'Usuario local',
            email: (user && user.email) || '',
            rol: 'Solo lectura',
            permisosPorModulo: getPermissionsForRole('Solo lectura'),
            activo: true
          };
          updateSolicitudesImportStatus();
          actualizarEstadoBanner();
          actualizarDashboard();
          renderCalendario();
          renderRutas();
          renderRouteCatalog();
          renderComercial();
          renderConfigClientes();
          renderSolicitudesAlmacen();
          renderAlmacen();
          return;
        }
        throw new Error('No hay sesión autenticada para cargar desde Supabase.');
      }

      const [
        lineRows,
        clientConfigRows,
        routeConfigRows,
        routeCostRows,
        warehousePlanRows,
        warehouseControlRows,
        warehouseHistoryRows,
        userProfileRows,
        settingsRows
      ] = await Promise.all([
        runSupabaseQuery(client.from('tms_order_lines').select('*').order('created_at', { ascending: true }), 'No se pudieron cargar las líneas'),
        runSupabaseQuery(client.from('tms_client_configs').select('*').order('codigo_cliente', { ascending: true }), 'No se pudo cargar la configuración de clientes'),
        runSupabaseQuery(client.from('tms_route_configs').select('*').order('nombre', { ascending: true }), 'No se pudo cargar la configuración de rutas'),
        runSupabaseQuery(client.from('tms_route_cost_history').select('*').order('fecha', { ascending: true }), 'No se pudo cargar el histórico de rutas'),
        runSupabaseQuery(client.from('tms_warehouse_plan').select('*').order('fecha', { ascending: true }), 'No se pudo cargar la planificación de solicitudes'),
        runSupabaseQuery(client.from('tms_warehouse_control').select('*').order('fecha', { ascending: true }), 'No se pudo cargar el control de solicitudes'),
        runSupabaseQuery(client.from('tms_warehouse_history').select('*').order('fecha', { ascending: true }), 'No se pudo cargar el histórico de solicitudes'),
        runSupabaseQuery(client.from('tms_user_profiles').select('*').order('created_at', { ascending: true }), 'No se pudieron cargar los usuarios'),
        runSupabaseQuery(client.from('tms_settings').select('*'), 'No se pudieron cargar los ajustes')
      ]);

      APP.clienteConfig = {};
      clientConfigRows.forEach(row => {
        APP.clienteConfig[row.codigo_cliente] = {
          ...buildClientConfigDefault(row.codigo_cliente, row.nombre_cliente || ''),
          codigoCliente: row.codigo_cliente,
          nombreCliente: row.nombre_cliente || '',
          rutaAsignada: row.ruta_asignada || '',
          diasRecepcion: row.dias_recepcion || [],
          horarioAlmacen: row.horario_almacen || '',
          contacto: row.contacto || '',
          observaciones: row.observaciones || '',
          condicionesEspeciales: row.condiciones_especiales || '',
          camionPermitido: row.camion_permitido || 'CUALQUIERA',
          noMiercoles: !!row.no_miercoles,
          noUltimaSemana: !!row.no_ultima_semana,
          configuracionCompleta: !!row.configuracion_completa
        };
      });

      APP.routeConfigs = routeConfigRows.map(row => ({
        id: row.id,
        nombre: row.nombre,
        zona: row.zona || '',
        diasOperacion: row.dias_operacion || [],
        costeRuta: num(row.coste_ruta),
        activa: row.activa !== false,
        observaciones: row.observaciones || ''
      }));

      APP.userProfiles = userProfileRows.map(row => {
        const account = row.permisos_por_modulo && row.permisos_por_modulo.__account ? row.permisos_por_modulo.__account : {};
        return {
          userId: row.id,
          authUserId: row.auth_user_id || '',
          nombre: row.nombre,
          email: row.email,
          rol: row.rol,
          permisosPorModulo: row.permisos_por_modulo || buildPermissions('read_only', row.rol),
          activo: row.activo !== false,
          passwordChangeRequired: !!account.passwordChangeRequired,
          authNeedsConfirmation: !!account.authNeedsConfirmation,
          passwordConfigured: account.passwordConfigured !== false,
          accountStatus: account.accountStatus || ''
        };
      });

      const settingsMap = Object.fromEntries(settingsRows.map(row => [row.clave, row.valor || {}]));
      APP.warehouseSettings = settingsMap.warehouse || { costoSolicitud: 0 };
      APP.rolePermissions = settingsMap.role_permissions || APP.rolePermissions || buildDefaultRolePermissions();
      ensureRolePermissions();
      if (settingsMap.app_state) {
        APP.planLoaded = !!settingsMap.app_state.planLoaded;
        APP.planFecha = settingsMap.app_state.planFecha || null;
        APP.controlLoaded = !!settingsMap.app_state.controlLoaded;
        APP.controlFecha = settingsMap.app_state.controlFecha || null;
        APP.solicitudesLoaded = !!settingsMap.app_state.solicitudesLoaded;
        APP.solicitudesFileName = settingsMap.app_state.solicitudesFileName || '';
        APP.camionExtraEnabled = !!settingsMap.app_state.camionExtraEnabled;
        APP.feriadosRD = settingsMap.app_state.feriadosRD || APP.feriadosRD;
        APP.importFiles = { ...APP.importFiles, ...(settingsMap.app_state.importFiles || {}) };
        APP.importShipments = Array.isArray(settingsMap.app_state.importShipments) ? settingsMap.app_state.importShipments : (APP.importShipments || []);
        APP.dailyPriorities = Array.isArray(settingsMap.app_state.dailyPriorities) ? settingsMap.app_state.dailyPriorities : (APP.dailyPriorities || []);
        APP.prioritySpainTopics = Array.isArray(settingsMap.app_state.prioritySpainTopics) ? settingsMap.app_state.prioritySpainTopics : (APP.prioritySpainTopics || []);
        APP.calendarNotes = Array.isArray(settingsMap.app_state.calendarNotes) ? settingsMap.app_state.calendarNotes : (APP.calendarNotes || []);
        APP.controlHistory = settingsMap.app_state.controlHistory || APP.controlHistory || [];
        APP.visualTheme = normalizeVisualTheme(settingsMap.app_state.visualTheme || APP.visualTheme || localStorage.getItem(THEME_KEY));
        applyVisualTheme(APP.visualTheme, true);
      }

      const currentItems = [];
      const planItems = [];
      const controlItems = [];
      lineRows.forEach(row => {
        const item = mapDbOrderLine(row);
        const dataset = text((row.metadata || {}).dataset) || (row.fecha_control ? 'control' : 'current');
        if (dataset === 'plan') planItems.push(item);
        else if (dataset === 'control') controlItems.push(item);
        else currentItems.push(item);
      });

      APP.planLineItems = planItems;
      APP.controlLineItems = controlItems;
      APP.lineItems = currentItems.length ? currentItems : (planItems.length ? planItems.map(item => ({ ...item })) : []);
      APP.planLoaded = APP.planLoaded || APP.planLineItems.length > 0;
      APP.controlLoaded = APP.controlLoaded || APP.controlLineItems.length > 0;
      APP.planFecha = APP.planFecha || (APP.planLineItems[0] && APP.planLineItems[0].fechaPlanificada) || null;
      APP.controlFecha = APP.controlFecha || (APP.controlLineItems[0] && APP.controlLineItems[0].fechaControl) || null;

      APP.routeCostHistory = routeCostRows.map(row => ({
        fecha: isoToFechaStr(row.fecha),
        rutaId: row.ruta_id || '',
        rutaNombre: row.ruta_nombre || '',
        camion: row.camion || '',
        coste: num(row.coste),
        fuenteControlDiario: row.fuente_control_diario || '',
        pedidosAsociados: row.pedidos_asociados || [],
        cumplimiento: num(row.cumplimiento),
        importePlanificado: num(row.importe_planificado),
        importeReal: num(row.importe_real),
        diferencia: num(row.diferencia)
      }));

      APP.solicitudesPlanAlmacen = warehousePlanRows.map(mapDbWarehousePlan);
      APP.solicitudesControlAlmacen = warehouseControlRows.map(mapDbWarehouseControl);
      APP.solicitudesAlmacen = APP.solicitudesControlAlmacen.length ? APP.solicitudesControlAlmacen : APP.solicitudesPlanAlmacen;
      APP.solicitudesHistory = warehouseHistoryRows.map(row => ({
        fecha: isoToFechaStr(row.fecha),
        clienteId: row.cliente_id,
        clienteNombre: row.cliente_nombre || '',
        pedidoCliente: row.pedido_cliente || '',
        solicitud: row.solicitud_id || '',
        articulo: row.articulo || '',
        descripcionArticulo: row.descripcion_articulo || '',
        cantidadSolicitada: num(row.cantidad_solicitada),
        cantidadCompletada: num(row.cantidad_completada),
        cantidadPendiente: num(row.cantidad_pendiente),
        estado: row.estado || '',
        costeSolicitud: num(row.coste_solicitud),
        observaciones: row.observaciones || '',
        fuente: row.fuente || ''
      }));
      APP.solicitudesCompare = buildSolicitudesComparison();
      APP.solicitudesLoaded = APP.solicitudesLoaded || !!(APP.solicitudesPlanAlmacen.length || APP.solicitudesControlAlmacen.length);

      ensureAdminProfile();
      APP.currentUserProfile =
        APP.userProfiles.find(profile => text(profile.email).toLowerCase() === text(user.email).toLowerCase()) ||
        {
          userId: user.id || 'unprofiled-user',
          authUserId: user.id || '',
          nombre: user.email || 'Usuario sin perfil',
          email: (user && user.email) || '',
          rol: 'Solo lectura',
          permisosPorModulo: getPermissionsForRole('Solo lectura'),
          activo: true
        };

      rebuildDerivedState();
      await enforcePasswordChangeIfNeeded();
      updateSolicitudesImportStatus();
      actualizarEstadoBanner();
      actualizarDashboard();
      renderCalendario();
      renderRutas();
      renderRouteCatalog();
      renderComercial();
      renderConfigClientes();
      renderSolicitudesAlmacen();
      renderAlmacen();
      renderPrioridades();
      applyPermissionUi();
      saveLocalSnapshot();
    } catch (error) {
      console.warn('Fallo la carga remota, usando respaldo local si existe:', error);
      if (loadLocalSnapshot() || loadEmbeddedDemoState()) {
        updateSolicitudesImportStatus();
        actualizarEstadoBanner();
        actualizarDashboard();
        renderCalendario();
        renderRutas();
        renderRouteCatalog();
        renderComercial();
        renderConfigClientes();
        renderSolicitudesAlmacen();
        renderAlmacen();
        renderPrioridades();
        return;
      }
      APP.lineItems = [];
      APP.planLineItems = [];
      APP.controlLineItems = [];
      APP.planLoaded = false;
      APP.controlLoaded = false;
      APP.solicitudesLoaded = false;
      rebuildDerivedState();
      updateSolicitudesImportStatus();
      actualizarEstadoBanner();
      actualizarDashboard();
      renderCalendario();
      renderRutas();
      renderRouteCatalog();
      renderComercial();
      renderConfigClientes();
      renderSolicitudesAlmacen();
      renderAlmacen();
      renderPrioridades();
    }
  };

  window.guardarEnSupabase = async function guardarEnSupabaseV2(options) {
    const silent = !!(options && options.silent);
    saveLocalSnapshot();
    const client = getSupabaseClient();
    const user = await getAuthenticatedUser();
    const btn = document.getElementById('saveStateBtn');
    const originalText = btn ? btn.textContent : '';
    if (btn) {
      btn.disabled = true;
      btn.textContent = '⏳ Guardando...';
    }
    try {
      if (!client || !user) throw new Error('No hay sesión autenticada para guardar en Supabase.');
      ensureAdminProfile();
      APP.currentUserProfile =
        APP.userProfiles.find(profile => text(profile.email).toLowerCase() === text(user.email).toLowerCase()) ||
        APP.currentUserProfile ||
        {
          userId: user.id || 'unprofiled-user',
          authUserId: user.id || '',
          nombre: user.email || 'Usuario sin perfil',
          email: (user && user.email) || '',
          rol: 'Solo lectura',
          permisosPorModulo: getPermissionsForRole('Solo lectura'),
          activo: true
        };

      const adminCanManageUsers = isAdminUser();
      const canWriteImports = adminCanManageUsers || hasPermission('importar', 'editar') || hasPermission('importar', 'importar');
      const canWriteRoutes = adminCanManageUsers || hasPermission('rutas', 'editar');
      const canWriteConfig = adminCanManageUsers || hasPermission('configuracion', 'editar') || hasPermission('configuracion', 'importar');
      const canWriteWarehousePlan = adminCanManageUsers || canWriteImports || hasPermission('solicitudesAlmacen', 'editar');
      const canWriteWarehouseControl = adminCanManageUsers || canWriteImports || hasPermission('almacen', 'editar');
      const canWriteWarehouseHistory = adminCanManageUsers || hasPermission('solicitudesAlmacen', 'editar') || hasPermission('almacen', 'editar');
      const canWriteCalendarNotes = hasPermission('calendario', 'ver');
      const canWriteAppState = adminCanManageUsers || canWriteImports || canWriteWarehousePlan || canWriteWarehouseControl || canWriteConfig || canWriteRoutes || canWriteCalendarNotes;
      const settingsRows = [
        ...((adminCanManageUsers || canWriteWarehouseHistory) ? [{ clave: 'warehouse', valor: safeClone(APP.warehouseSettings || { costoSolicitud: 0 }) }] : []),
        ...(adminCanManageUsers ? [{ clave: 'role_permissions', valor: safeClone(APP.rolePermissions || buildDefaultRolePermissions()) }] : []),
        ...(canWriteAppState ? [{
          clave: 'app_state',
          valor: safeClone({
            planLoaded: !!APP.planLoaded,
            planFecha: APP.planFecha || null,
            controlLoaded: !!APP.controlLoaded,
            controlFecha: APP.controlFecha || null,
            solicitudesLoaded: !!APP.solicitudesLoaded,
            solicitudesFileName: APP.solicitudesFileName || '',
            camionExtraEnabled: !!APP.camionExtraEnabled,
            feriadosRD: APP.feriadosRD || [],
            importFiles: APP.importFiles || {},
            importShipments: APP.importShipments || [],
            dailyPriorities: APP.dailyPriorities || [],
            prioritySpainTopics: APP.prioritySpainTopics || [],
            calendarNotes: APP.calendarNotes || [],
            controlHistory: APP.controlHistory || [],
            visualTheme: APP.visualTheme || 'light'
          })
        }] : [])
      ];

      const userRows = APP.userProfiles.map(profile => userProfileDbPayload(
        profile,
        text(profile.email).toLowerCase() === text(user.email).toLowerCase() ? user.id : null
      ));
      const currentUserProfileRow = APP.currentUserProfile && APP.currentUserProfile.email
        ? userProfileDbPayload(APP.currentUserProfile, user.id)
        : null;

      const routeRows = APP.routeConfigs.map(route => {
        const payload = {
          nombre: route.nombre || route.zona || '',
          zona: route.zona || route.nombre || '',
          dias_operacion: route.diasOperacion || [],
          coste_ruta: num(route.costeRuta),
          activa: route.activa !== false,
          observaciones: route.observaciones || ''
        };
        if (isUuid(route.id)) payload.id = route.id;
        return payload;
      }).filter(route => route.nombre);

      const clientRows = Object.values(APP.clienteConfig || {}).map(cfg => ({
        codigo_cliente: cfg.codigoCliente || cfg.codigo || '',
        nombre_cliente: cfg.nombreCliente || '',
        ruta_asignada: cfg.rutaAsignada || '',
        dias_recepcion: cfg.diasRecepcion || [],
        horario_almacen: cfg.horarioAlmacen || '',
        contacto: cfg.contacto || '',
        observaciones: cfg.observaciones || '',
        condiciones_especiales: cfg.condicionesEspeciales || '',
        camion_permitido: cfg.camionPermitido || 'CUALQUIERA',
        configuracion_completa: !!cfg.configuracionCompleta,
        no_miercoles: !!cfg.noMiercoles,
        no_ultima_semana: !!cfg.noUltimaSemana
      })).filter(cfg => cfg.codigo_cliente);

      const orderRows = [
        ...buildOrderLineDbRows(APP.lineItems, 'current'),
        ...buildOrderLineDbRows(APP.planLineItems, 'plan'),
        ...buildOrderLineDbRows(APP.controlLineItems, 'control')
      ];

      const routeHistoryRows = APP.routeCostHistory.map(row => ({
        fecha: fechaStrToISO(row.fecha),
        ruta_id: isUuid(row.rutaId) ? row.rutaId : null,
        ruta_nombre: row.rutaNombre || '',
        camion: row.camion || '',
        coste: num(row.coste),
        fuente_control_diario: row.fuenteControlDiario || '',
        pedidos_asociados: safeClone(row.pedidosAsociados || []),
        cumplimiento: num(row.cumplimiento),
        importe_planificado: num(row.importePlanificado),
        importe_real: num(row.importeReal),
        diferencia: num(row.diferencia)
      })).filter(row => row.fecha && row.ruta_nombre);

      const warehousePlanRows = buildWarehousePlanDbRows(APP.solicitudesPlanAlmacen);
      const warehouseControlRows = buildWarehouseControlDbRows(APP.solicitudesControlAlmacen);
      const warehouseHistoryRows = APP.solicitudesHistory.map(row => ({
        fecha: fechaStrToISO(row.fecha),
        cliente_id: row.clienteId || '',
        cliente_nombre: row.clienteNombre || '',
        pedido_cliente: row.pedidoCliente || '',
        solicitud_id: row.solicitud || '',
        articulo: row.articulo || '',
        descripcion_articulo: row.descripcionArticulo || '',
        cantidad_solicitada: num(row.cantidadSolicitada),
        cantidad_completada: num(row.cantidadCompletada),
        cantidad_pendiente: num(row.cantidadPendiente),
        estado: row.estado || '',
        coste_solicitud: num(row.costeSolicitud),
        observaciones: row.observaciones || '',
        fuente: row.fuente || ''
      })).filter(row => row.fecha && row.cliente_id);

      if (settingsRows.length) {
        await runSupabaseQuery(client.from('tms_settings').upsert(settingsRows, { onConflict: 'clave' }), 'No se pudieron guardar los ajustes');
      }
      if (adminCanManageUsers) {
        await replaceSupabaseTable('tms_user_profiles', userRows);
      } else if (currentUserProfileRow) {
        await runSupabaseQuery(client.from('tms_user_profiles').upsert(currentUserProfileRow, { onConflict: 'email' }), 'No se pudo actualizar tu estado de acceso');
      }
      if (adminCanManageUsers || canWriteConfig || canWriteRoutes) await replaceSupabaseTable('tms_route_configs', routeRows);
      if (adminCanManageUsers || canWriteConfig || canWriteRoutes) await replaceSupabaseTable('tms_client_configs', clientRows);
      if (adminCanManageUsers || canWriteImports || hasPermission('calendario', 'editar') || canWriteRoutes) await replaceSupabaseTable('tms_order_lines', orderRows);
      if (adminCanManageUsers || canWriteRoutes) await replaceSupabaseTable('tms_route_cost_history', routeHistoryRows);
      if (canWriteWarehousePlan) await replaceSupabaseTable('tms_warehouse_plan', warehousePlanRows);
      if (canWriteWarehouseControl) await replaceSupabaseTable('tms_warehouse_control', warehouseControlRows);
      if (canWriteWarehouseHistory) await replaceSupabaseTable('tms_warehouse_history', warehouseHistoryRows);

      if (btn) {
        btn.textContent = APP.cloudMigrationPending ? 'Nube modo compat.' : 'Guardado en nube';
      }
      if (APP.cloudMigrationPending && !silent && !window._tmsMigrationPendingNoticeShown) {
        window._tmsMigrationPendingNoticeShown = true;
        alert('Cambios guardados en la nube en modo compatibilidad. Falta aplicar la migración 0004 en Supabase para activar guardado transaccional completo.');
      }
      return true;
    } catch (error) {
      console.error('Error guardando en Supabase:', error);
      if (btn) btn.textContent = '⚠️ Guardado local';
      if (!silent) notifyCloudSaveError(error);
      throw error;
    } finally {
      if (btn) {
        setTimeout(() => {
          btn.disabled = false;
          btn.textContent = originalText || 'Guardar cambios';
        }, 1400);
      }
    }
  };

  function initImportSolicitudesCard() {
    const importView = document.getElementById('importar');
    if (!importView || document.getElementById('solicitudesImportCard')) return;
    const card = document.createElement('div');
    card.id = 'solicitudesImportCard';
    card.className = 'card';
    card.style.borderLeft = '4px solid #C2410C';
    card.innerHTML = `
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px;">
        <span style="font-size:20px;">📦</span>
        <div>
          <div class="card-title" style="margin-bottom:0;color:#C2410C;">Solicitudes a Almacén</div>
          <div style="font-size:12px;color:var(--muted);">Sube las solicitudes realizadas al almacén en XLSX o XML y compara lo pendiente vs. preparado.</div>
        </div>
      </div>
      <div class="view-toolbar" style="margin-top:14px;align-items:flex-start;">
        <div style="display:flex;flex-direction:column;gap:8px;">
          <input type="file" id="solicitudesPlanFileImport" accept=".xlsx,.xls,.xml" style="display:none" onchange="procesarSolicitudesAlmacen(this.files[0],'plan')">
          <button class="btn btn-warning" onclick="document.getElementById('solicitudesPlanFileImport').click()">Solicitudes realizadas al almacén</button>
          <div id="solPlanImportStatus" style="font-size:12px;color:var(--muted);">Sin archivo de planificación</div>
        </div>
        <div style="display:flex;flex-direction:column;gap:8px;">
          <input type="file" id="solicitudesControlFileImport" accept=".xlsx,.xls,.xml" style="display:none" onchange="procesarSolicitudesAlmacen(this.files[0],'control')">
          <button class="btn btn-outline" onclick="document.getElementById('solicitudesControlFileImport').click()">Control diario de almacén</button>
          <div id="solCtrlImportStatus" style="font-size:12px;color:var(--muted);">Sin archivo de control</div>
        </div>
        <span style="font-size:12px;color:var(--muted);" id="solicitudesImportHint"><strong id="solNewRequestsCount">0</strong> nuevas solicitudes detectadas</span>
      </div>`;
    importView.appendChild(card);
  }

  function updateSolicitudesImportStatus() {
    const planStatus = document.getElementById('solPlanImportStatus');
    const ctrlStatus = document.getElementById('solCtrlImportStatus');
    if (planStatus) planStatus.textContent = APP.importFiles.solicitudesPlan || 'Sin archivo de planificación';
    if (ctrlStatus) ctrlStatus.textContent = APP.importFiles.solicitudesControl || 'Sin archivo de control';
    const newCount = document.getElementById('solNewRequestsCount');
    if (newCount) newCount.textContent = getSolicitudesNewCount();
    updateImportExcelTableStatus();
  }


  function getSolicitudesNewCount() {
    const controlKeys = new Set((APP.solicitudesControlAlmacen || []).map(item => [item.codigo, item.solicitud, item.pedidoCliente, item.lineaPedidoCliente || '', item.articulo].join('|')));
    const planKeys = new Set((APP.solicitudesPlanAlmacen || []).map(item => [item.codigo, item.solicitud, item.pedidoCliente, item.lineaPedidoCliente || '', item.articulo].join('|')));
    const source = controlKeys.size ? controlKeys : planKeys;
    const base = controlKeys.size ? planKeys : new Set();
    return [...source].filter(key => !base.has(key)).length;
  }

  function statusBadgeHtml(status, tone) {
    return `<span class="import-table-badge ${tone || ''}">${status}</span>`;
  }

  function initImportExcelTable() {
    const importView = document.getElementById('importar');
    if (!importView || document.getElementById('importExcelTableCard')) return;
    importView.classList.add('import-table-mode');
    const title = importView.querySelector('.view-title');
    const tableCard = document.createElement('div');
    tableCard.id = 'importExcelTableCard';
    tableCard.className = 'card import-excel-card';
    tableCard.innerHTML = `
      <div class="import-excel-header">
        <div>
          <div class="card-title">Carga de archivos SAP</div>
          <div class="import-excel-subtitle">Sube o reemplaza los archivos base del TMS desde una tabla sencilla.</div>
        </div>
        <div class="import-excel-meta">XLSX / XML</div>
      </div>
      <div class="import-excel-wrap">
        <table class="import-excel-table">
          <thead>
            <tr>
              <th>Archivo requerido</th>
              <th>Uso en el TMS</th>
              <th>Estado actual</th>
              <th>Acción</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Programación mensual</strong><span>Pedidos de cliente / planificación comercial</span></td>
              <td>Define calendario, rutas iniciales y montos pedidos.</td>
              <td id="importTablePlanStatus">No cargada</td>
              <td><button class="btn btn-primary btn-sm" onclick="document.getElementById('sapFile').click()">Subir</button></td>
            </tr>
            <tr>
              <td><strong>Control diario</strong><span>Cierre del día</span></td>
              <td>Actualiza cantidades facturadas, pendientes y avance real.</td>
              <td id="importTableCtrlStatus">No cargado</td>
              <td><button class="btn btn-success btn-sm" onclick="abrirControlDiario()">Subir</button></td>
            </tr>
            <tr>
              <td><strong>Solicitudes a almacén</strong><span>Pendientes de preparación / control almacén</span></td>
              <td>Concilia pedidos cliente contra solicitudes realizadas al almacén.</td>
              <td id="importTableSolStatus">Sin archivo</td>
              <td>
                <div class="import-table-actions">
                  <button class="btn btn-warning btn-sm" onclick="document.getElementById('solicitudesPlanFileImport').click()">Plan</button>
                  <button class="btn btn-outline btn-sm" onclick="document.getElementById('solicitudesControlFileImport').click()">Control</button>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="import-excel-foot">
        <span id="importTableNewCount">0 nuevas solicitudes</span>
        <span id="importTableLastFile">Esperando archivos de carga</span>
      </div>`;
    if (title && title.nextSibling) importView.insertBefore(tableCard, title.nextSibling);
    else importView.prepend(tableCard);
    updateImportExcelTableStatus();
  }

  function updateImportExcelTableStatus() {
    const planStatus = document.getElementById('importTablePlanStatus');
    const ctrlStatus = document.getElementById('importTableCtrlStatus');
    const solStatus = document.getElementById('importTableSolStatus');
    const newCount = document.getElementById('importTableNewCount');
    const lastFile = document.getElementById('importTableLastFile');
    if (!planStatus || !ctrlStatus || !solStatus) return;

    planStatus.innerHTML = APP.planLoaded
      ? `${statusBadgeHtml('Cargada', 'ok')}<span>${new Set(APP.planLineItems.map(item => item.clienteId)).size} clientes · ${APP.planLineItems.length} líneas · ${APP.importFiles.plan || APP.planFecha}</span>`
      : `${statusBadgeHtml('Pendiente', 'wait')}<span>No cargada</span>`;
    ctrlStatus.innerHTML = APP.controlLoaded
      ? `${statusBadgeHtml('Cargado', 'ok')}<span>${new Set(APP.controlLineItems.map(item => item.clienteId)).size} clientes · ${APP.controlLineItems.length} líneas · ${APP.importFiles.control || APP.controlFecha}</span>`
      : `${statusBadgeHtml('Opcional', 'info')}<span>Sin control diario</span>`;
    solStatus.innerHTML = APP.solicitudesLoaded
      ? `${statusBadgeHtml('Cargadas', 'ok')}<span>${APP.solicitudesAlmacen.length} solicitudes · ${APP.solicitudesFileName || 'archivo actual'}</span>`
      : `${statusBadgeHtml('Pendiente', 'wait')}<span>Sin archivo</span>`;

    const count = getSolicitudesNewCount();
    if (newCount) newCount.textContent = `${count} nuevas solicitudes detectadas`;
    if (lastFile) {
      const file = APP.importFiles.solicitudesControl || APP.importFiles.solicitudesPlan || APP.importFiles.control || APP.importFiles.plan || '';
      lastFile.textContent = file ? `Último archivo: ${file}` : 'Esperando archivos de carga';
    }
  }

  function initCalendarActions() {
    const toolbar = document.querySelector('#calendario .view-toolbar');
    if (!toolbar || document.getElementById('addTruckBtn')) return;
    toolbar.classList.add('calendar-toolbar');
    const button = document.createElement('button');
    button.id = 'addTruckBtn';
    button.className = 'btn btn-outline btn-sm';
    const syncButton = () => { button.textContent = APP.camionExtraEnabled ? 'Quitar camión adicional' : 'Agregar camión adicional'; };
    button.onclick = function () {
      pushUndoState(APP.camionExtraEnabled ? 'quitar camión adicional' : 'agregar camión adicional');
      APP.camionExtraEnabled = !APP.camionExtraEnabled;
      if (!APP.camionExtraEnabled) {
        APP.lineItems.forEach(item => {
          if (item.camionAsignado === 'CAMION 3') item.camionAsignado = chooseTruckForItem(item);
        });
      }
      syncMoveOptions();
      rebuildDerivedState();
      renderCalendario();
      renderRutas();
      scheduleAutoSave();
      syncButton();
    };
    toolbar.appendChild(button);
    syncButton();
  }

  function initCommercialMount() {
    const view = document.getElementById('comercial');
    const card = document.querySelector('#comercial .card');
    if (!view || !card || document.getElementById('comercialV2Mount')) return;
    const kpis = view.querySelectorAll('.kpi-label');
    if (kpis[0]) kpis[0].textContent = 'Clientes';
    if (kpis[1]) kpis[1].textContent = 'Pedidos';
    if (kpis[2]) kpis[2].textContent = 'Cant. planificada';
    if (kpis[3]) kpis[3].textContent = 'Monto';

    const routeOptions = getCommercialRouteOptions();
    card.innerHTML = `
      <div class="commercial-toolbar">
        <div class="commercial-quick">
          <span>Acceso rápido:</span>
          <button id="cf_todas" class="filter-btn active" onclick="filtrarCom('todas')">Todas</button>
          <button id="cf_sem1" class="filter-btn" onclick="filtrarCom('sem1')">Semana 1</button>
          <button id="cf_sem2" class="filter-btn" onclick="filtrarCom('sem2')">Semana 2</button>
          <button id="cf_sem3" class="filter-btn" onclick="filtrarCom('sem3')">Semana 3</button>
          <button id="cf_c1" class="filter-btn" onclick="filtrarCom('c1')">C1</button>
          <button id="cf_c2" class="filter-btn" onclick="filtrarCom('c2')">C2</button>
        </div>
        <div class="commercial-actions">
          <button class="btn btn-outline btn-sm" onclick="limpiarFiltrosCom()">Limpiar</button>
          <button class="btn btn-outline btn-sm" onclick="exportarComercialPDF()">PDF vertical</button>
          <button class="btn btn-success btn-sm" onclick="exportarComercialExcel()">Excel XLSX</button>
        </div>
      </div>
      <div class="commercial-filter-grid">
        <input id="cfComCliente" class="search-input" placeholder="Cliente, pedido o referencia" oninput="renderComercial()">
        <input id="cfComFechaInicio" class="search-input" type="date" onchange="renderComercial()">
        <input id="cfComFechaFin" class="search-input" type="date" onchange="renderComercial()">
        <select id="cfComRuta" class="search-input" onchange="renderComercial()">
          <option value="">Todas las rutas/camiones</option>
          ${routeOptions.map(route => `<option value="${route}">${route}</option>`).join('')}
        </select>
        <select id="cfComEstado" class="search-input" onchange="renderComercial()">
          <option value="">Todos los estados</option>
          <option value="pendiente">Pendiente</option>
          <option value="parcial">Parcial</option>
          <option value="entregado">Entregado</option>
          <option value="no considerado">No considerado</option>
        </select>
      </div>
      <input id="cfcol_fecha" type="hidden">
      <div id="comercialV2Mount" style="margin-top:12px;"></div>
    `;
  }

  function initReportFilters() {
    const toolbar = document.querySelector('#dashReportPanel .view-toolbar');
    if (!toolbar || document.getElementById('repDateStart')) return;
    toolbar.insertAdjacentHTML('beforeend', `
      <input type="date" id="repDateStart" class="search-input" style="max-width:160px;">
      <input type="date" id="repDateEnd" class="search-input" style="max-width:160px;">
      <select id="repRouteFilter" class="search-input" style="max-width:180px;"><option value="">Todas las rutas</option></select>
    `);
  }

  function initThemeSettings() {
    const view = document.getElementById('configClientes');
    if (!view || document.getElementById('visualThemeCard')) return;
    const card = document.createElement('div');
    card.id = 'visualThemeCard';
    card.className = 'card theme-settings-card';
    card.innerHTML = `
      <div class="theme-settings-head">
        <div>
          <div class="card-title">Tema visual</div>
          <div class="theme-settings-sub">Simplifica colores, iconos y encabezados para cada forma de trabajo.</div>
        </div>
        <span id="themeCurrentLabel" class="badge badge-outline">${getThemeLabel(APP.visualTheme)}</span>
      </div>
      <div class="theme-option-grid">
        <button type="button" class="theme-option" data-theme-option="light" onclick="setVisualTheme('light')">
          <span class="theme-preview theme-preview-light"></span>
          <strong>Claro operativo</strong>
          <small>Blanco, gris y azul sobrio.</small>
        </button>
        <button type="button" class="theme-option" data-theme-option="dark" onclick="setVisualTheme('dark')">
          <span class="theme-preview theme-preview-dark"></span>
          <strong>Oscuro</strong>
          <small>Menos brillo para uso continuo.</small>
        </button>
        <button type="button" class="theme-option" data-theme-option="brand" onclick="setVisualTheme('brand')">
          <span class="theme-preview theme-preview-brand"></span>
          <strong>Alvarez color</strong>
          <small>Marca simple con acento amarillo.</small>
        </button>
      </div>
    `;
    const toolbarCard = view.querySelector('.card');
    view.insertBefore(card, toolbarCard || view.firstChild);
    syncThemeButtons();
  }

  function initConfigExtras() {
    const view = document.getElementById('configClientes');
    if (!view || document.getElementById('configExtrasMount')) return;
    const mount = document.createElement('div');
    mount.id = 'configExtrasMount';
    view.appendChild(mount);
  }

  function initUiLabels() {
    const navConfig = document.querySelector('#nav-configClientes .nav-label');
    if (navConfig) navConfig.textContent = 'Configuración';
    const titleConfig = document.querySelector('#configClientes .view-title');
    if (titleConfig) titleConfig.textContent = 'Configuración';
    const titleImport = document.querySelector('#importar .view-title');
    if (titleImport) titleImport.textContent = 'Importar Datos';
    const planTitle = document.querySelector('#uploadZone')?.closest('.card')?.querySelector('.card-title');
    if (planTitle) planTitle.textContent = 'Programación mensual / pedidos de cliente';
    const planSub = document.querySelector('#uploadZone .upload-sub');
    if (planSub) planSub.textContent = 'Archivo SAP CRMSLOIB01_Q0001 en XLSX o XML';
    const controlSub = document.querySelector('#ctrlUploadZone .upload-sub');
    if (controlSub) controlSub.textContent = 'Cierre diario SAP en XLSX o XML';
  }

  function enhanceImportUi() {
    const importView = document.getElementById('importar');
    if (!importView) return;
    importView.classList.add('import-view-shell');
    const statusBanner = document.getElementById('sapEstadoBanner');
    if (statusBanner) statusBanner.classList.add('import-status-banner');
    ['planEstadoCard', 'ctrlEstadoCard'].forEach(id => {
      const card = document.getElementById(id);
      if (card) card.classList.add('import-status-card');
    });
    ['uploadZone', 'ctrlUploadZone'].forEach(id => {
      const zone = document.getElementById(id);
      if (zone) zone.classList.add('upload-zone-premium');
    });
    const importCards = importView.querySelectorAll('.card');
    importCards.forEach(card => card.classList.add('import-card-shell'));
    const rutasView = document.getElementById('rutas');
    if (rutasView) rutasView.classList.add('routes-view-shell');
  }

  function initRouteFilterOptions() {
    const select = document.getElementById('repRouteFilter');
    if (!select) return;
    const options = [...new Set([
      ...APP.routeConfigs.map(route => route.nombre || route.zona),
      ...APP.lineItems.map(item => item.rutaNombre || item.zona)
    ].filter(Boolean))].sort();
    select.innerHTML = '<option value="">Todas las rutas</option>' + options.map(route => `<option value="${route}">${route}</option>`).join('');
  }

  function injectV2Styles() {
    if (document.getElementById('tmsV2Styles')) return;
    const style = document.createElement('style');
    style.id = 'tmsV2Styles';
    style.textContent = `
      .form-row { display:flex; flex-direction:column; gap:6px; margin-bottom:12px; }
      .form-control { padding:8px 10px; border:1px solid var(--border); border-radius:8px; background:var(--surface, #fff); }
      #undoStateBtn { border:1px solid rgba(255,255,255,0.32); background:rgba(255,255,255,0.14); color:#fff; border-radius:8px; padding:8px 12px; font-size:12px; font-weight:800; cursor:pointer; }
      #undoStateBtn:disabled { opacity:.45; cursor:not-allowed; }
      .queue-card { border:1px solid var(--border); border-radius:8px; padding:12px; background:var(--surface, #fff); margin-bottom:10px; box-shadow:var(--shadow-sm, 0 1px 2px rgba(15,23,42,0.05)); }
      .queue-route-header { display:flex; align-items:center; justify-content:space-between; gap:8px; }
      .queue-route-header .btn { background:rgba(255,255,255,0.92); color:var(--primary); border-color:rgba(255,255,255,0.7); }
      .queue-card-title { font-weight:700; color:var(--primary); font-size:13px; }
      .queue-card-meta { font-size:11px; color:var(--muted); }
      .route-evidence-line { font-size:10px; color:#166534; font-weight:800; margin-top:5px; overflow-wrap:anywhere; }
      .badge-primary { background: var(--primary); color: #fff; }
      .badge-outline { background: #fff; color: var(--muted); border: 1px solid var(--border); }
      .config-help-note { color:var(--muted); font-size:12px; line-height:1.45; border:1px solid var(--border); border-radius:8px; padding:10px 12px; background:var(--surface-2, #F8FAFC); }
      .route-add-form { display:grid; grid-template-columns:minmax(180px, 1fr) minmax(110px, 160px) auto; gap:8px; align-items:center; margin-bottom:12px; }
      .role-permission-toolbar { display:grid; grid-template-columns:minmax(180px, 240px) minmax(180px, 1fr) auto auto; gap:10px; align-items:end; margin-bottom:12px; }
      .role-permission-table th,
      .role-permission-table td { text-align:center; }
      .role-permission-table th:first-child,
      .role-permission-table td:first-child { text-align:left; }
      .role-permission-table input[type="checkbox"] { width:16px; height:16px; }
      .permission-guide { display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:8px; margin-top:12px; }
      .permission-guide-item { border:1px solid var(--border); border-radius:8px; padding:10px 12px; background:var(--surface-2, #F8FAFC); }
      .permission-guide-item strong { display:block; font-size:13px; color:var(--text); }
      .permission-guide-item span { display:block; font-size:11px; font-weight:800; color:var(--primary); text-transform:uppercase; margin-top:2px; }
      .permission-guide-item p { color:var(--muted); font-size:12px; line-height:1.35; margin-top:6px; }
      .permission-chip { display:inline-flex; align-items:center; border:1px solid var(--border); border-radius:999px; padding:3px 7px; margin:2px; font-size:10px; font-weight:800; white-space:nowrap; }
      .password-field-wrap { display:flex; gap:8px; align-items:center; }
      .password-field-wrap .form-control { min-width:0; flex:1; }
      .password-field-wrap .btn { flex:0 0 auto; }
      .permission-chip.edit { background:#EAF7EF; color:#166534; border-color:#B7E4C7; }
      .permission-chip.view { background:#EFF6FF; color:#1D4ED8; border-color:#BFDBFE; }
      .permission-chip.off,
      .permission-chip.muted { background:#F3F4F6; color:#6B7280; border-color:#E5E7EB; }
      .brand-login-logo {
        display:block;
        margin:0 auto 12px;
        font-size:0 !important;
      }
      .brand-login-logo img {
        display:block;
        width:min(300px, 82vw);
        height:auto;
        margin:0 auto;
      }
      .brand-header-logo {
        display:inline-flex !important;
        align-items:center;
        justify-content:center;
        width:176px;
        min-width:142px;
        height:38px;
        padding:4px 8px;
        border-radius:6px;
        background:#fff;
        overflow:hidden;
      }
      .brand-header-logo img {
        display:block;
        width:100%;
        height:100%;
        object-fit:contain;
      }
      .app-view-footer {
        margin:28px 0 4px;
        padding:12px 4px 0;
        border-top:1px solid var(--border);
        color:var(--muted);
        font-size:11px;
        line-height:1.35;
        text-align:center;
      }
      .view-title { letter-spacing:0; }
      .card, .kpi-card, .stat-box, .chart-card { box-shadow:0 10px 24px rgba(15,23,42,.06); }
      .import-module { display:grid; gap:14px; }
      .import-hero { display:flex; justify-content:space-between; gap:16px; align-items:flex-start; padding:18px; border:1px solid var(--border); border-radius:10px; background:linear-gradient(135deg, var(--surface), var(--surface-2)); box-shadow:var(--shadow-sm); }
      .import-hero h2 { margin:0; font-size:18px; color:var(--text); }
      .import-hero p { margin:4px 0 0; color:var(--muted); font-size:12px; line-height:1.45; }
      .import-kpis { display:grid; grid-template-columns:repeat(auto-fit, minmax(150px, 1fr)); gap:10px; }
      .import-kpi { border:1px solid var(--border); border-radius:10px; padding:12px; background:var(--surface); }
      .import-kpi strong { display:block; font-size:22px; color:var(--primary); }
      .import-kpi span { display:block; margin-top:2px; color:var(--muted); font-size:11px; text-transform:uppercase; font-weight:800; }
      .import-form-grid { display:grid; grid-template-columns:repeat(auto-fit, minmax(170px, 1fr)); gap:10px; }
      .import-stage-board { display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:12px; align-items:start; }
      .import-stage { border:1px solid var(--border); border-radius:10px; background:var(--surface); overflow:hidden; min-height:160px; }
      .import-stage-head { padding:10px 12px; background:var(--primary); color:#fff; display:flex; justify-content:space-between; gap:8px; font-size:12px; font-weight:800; }
      .import-stage-body { display:grid; gap:9px; padding:10px; }
      .import-card { border:1px solid var(--border); border-radius:9px; padding:10px; background:var(--surface-2); display:grid; gap:7px; }
      .import-card.alert { border-color:#F59E0B; background:#FFFBEB; }
      .import-card.overdue { border-color:#DC2626; background:#FEF2F2; }
      .import-card-title { display:flex; justify-content:space-between; gap:8px; font-size:13px; font-weight:800; color:var(--text); }
      .import-card-meta { color:var(--muted); font-size:11px; line-height:1.35; }
      .import-docs { display:flex; flex-wrap:wrap; gap:5px; }
      .import-doc-chip { border:1px solid var(--border); border-radius:999px; padding:3px 7px; font-size:10px; font-weight:800; background:#fff; color:var(--muted); }
      .import-doc-chip.ok { color:#166534; background:#F0FDF4; border-color:#BBF7D0; }
      .import-bulk-box { border:1px dashed var(--border); border-radius:10px; background:var(--surface-2); padding:12px; display:grid; gap:10px; }
      .import-bulk-box strong { font-size:13px; color:var(--text); }
      .import-bulk-box small { color:var(--muted); font-size:11px; line-height:1.4; }
      .import-bulk-actions { display:flex; flex-wrap:wrap; gap:8px; align-items:center; }
      .import-bulk-text { width:100%; min-height:82px; resize:vertical; }
      .import-bulk-status { color:var(--muted); font-size:12px; min-height:18px; }
      .import-actions { display:flex; flex-wrap:wrap; gap:6px; }
      .import-report { border:1px solid var(--border); border-radius:10px; background:var(--surface); padding:12px; white-space:pre-wrap; color:var(--text); font-size:12px; line-height:1.5; }
      .priorities-module { display:grid; gap:14px; }
      .priority-hero { display:flex; justify-content:space-between; gap:16px; align-items:flex-start; padding:18px; border:1px solid var(--border); border-radius:10px; background:linear-gradient(135deg, var(--surface), var(--surface-2)); box-shadow:var(--shadow-sm); }
      .priority-hero h2 { margin:0; font-size:18px; color:var(--text); }
      .priority-hero p { margin:4px 0 0; color:var(--muted); font-size:12px; line-height:1.45; }
      .priority-top-actions, .priority-form, .priority-actions { display:flex; flex-wrap:wrap; gap:8px; align-items:center; }
      .priority-form .form-control { min-width:180px; flex:1; }
      .priority-voice-status { margin-top:8px; font-size:12px; color:var(--muted); }
      .priority-kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:10px; }
      .priority-kpi { border:1px solid var(--border); border-radius:10px; background:var(--surface); padding:12px; }
      .priority-kpi strong { display:block; font-size:22px; color:var(--primary); }
      .priority-kpi span { color:var(--muted); font-size:11px; text-transform:uppercase; font-weight:800; }
      .priority-board { display:grid; grid-template-columns:repeat(auto-fit,minmax(250px,1fr)); gap:12px; align-items:start; }
      .priority-column { border:1px solid var(--border); border-radius:10px; background:var(--surface); overflow:hidden; min-height:180px; }
      .priority-column.drag-over { outline:2px solid var(--warning); outline-offset:2px; }
      .priority-column-head { padding:10px 12px; background:var(--primary); color:#fff; display:flex; justify-content:space-between; gap:8px; font-size:12px; font-weight:800; }
      .priority-column.status-pendiente .priority-column-head { background:#7F1D1D; }
      .priority-column.status-proceso .priority-column-head { background:#92400E; }
      .priority-column.status-completada .priority-column-head { background:#166534; }
      .priority-column-body { display:grid; gap:9px; padding:10px; }
      .priority-card { border:1px solid var(--border); border-left:4px solid var(--primary); border-radius:8px; background:var(--surface-2); padding:10px; display:grid; gap:7px; cursor:grab; }
      .priority-card.status-pendiente { border-left-color:#DC2626; }
      .priority-card.status-proceso { border-left-color:#F59E0B; }
      .priority-card.status-completada { border-left-color:#16A34A; }
      .priority-status-dot { width:9px; height:9px; border-radius:999px; display:inline-block; margin-right:6px; vertical-align:-1px; background:var(--primary); }
      .priority-status-dot.status-pendiente { background:#DC2626; }
      .priority-status-dot.status-proceso { background:#F59E0B; }
      .priority-status-dot.status-completada { background:#16A34A; }
      .priority-card-title { font-size:13px; font-weight:800; color:var(--text); line-height:1.25; }
      .priority-card-note { font-size:12px; color:var(--muted); line-height:1.35; }
      .priority-card-meta { display:flex; justify-content:space-between; gap:8px; color:var(--muted); font-size:10px; font-weight:800; text-transform:uppercase; }
      .priority-report { border:1px solid var(--border); border-radius:10px; background:var(--surface); padding:12px; }
      .priority-report pre { margin:0; white-space:pre-wrap; color:var(--text); font-size:12px; line-height:1.5; font-family:inherit; }
      .priority-insights-grid { display:grid; grid-template-columns:minmax(280px,1.25fr) minmax(260px,.85fr); gap:12px; align-items:stretch; }
      .priority-chart { border:1px solid var(--border); border-radius:10px; background:var(--surface); padding:12px; display:grid; gap:10px; }
      .priority-chart-head { display:flex; justify-content:space-between; gap:10px; align-items:flex-start; }
      .priority-chart-head strong { font-size:13px; color:var(--text); }
      .priority-chart-head span { font-size:11px; color:var(--muted); }
      .priority-chart-row { display:grid; grid-template-columns:74px 1fr; gap:8px; align-items:center; font-size:11px; color:var(--muted); }
      .priority-bars { display:grid; gap:3px; }
      .priority-bar-track { height:8px; border-radius:999px; background:#EEF2F7; overflow:hidden; }
      .priority-bar-fill { height:100%; min-width:4px; border-radius:999px; }
      .priority-bar-fill.started { background:#64748B; }
      .priority-bar-fill.carried { background:#F59E0B; }
      .priority-bar-fill.done { background:#16A34A; }
      .priority-bar-labels { display:flex; gap:8px; flex-wrap:wrap; font-size:10px; color:var(--muted); }
      .priority-spain-panel { border:1px solid var(--border); border-radius:10px; background:var(--surface); padding:12px; display:grid; gap:10px; }
      .priority-spain-head { display:flex; justify-content:space-between; gap:10px; align-items:flex-start; }
      .priority-spain-head strong { font-size:13px; color:var(--text); }
      .priority-spain-head span { font-size:11px; color:var(--muted); }
      .priority-spain-form { display:grid; grid-template-columns:1fr auto; gap:8px; }
      .priority-spain-list { display:grid; gap:7px; max-height:260px; overflow:auto; }
      .priority-spain-item { border:1px solid var(--border); border-radius:8px; background:var(--surface-2); padding:8px; display:grid; grid-template-columns:auto 1fr auto; gap:8px; align-items:start; }
      .priority-spain-item.done { opacity:.62; }
      .priority-spain-item strong { display:block; font-size:12px; color:var(--text); overflow-wrap:anywhere; }
      .priority-spain-item small { display:block; margin-top:2px; font-size:10px; color:var(--muted); }
      .priority-spain-item input { margin-top:2px; accent-color:var(--secondary); }
      @media (max-width: 760px) { .priority-insights-grid { grid-template-columns:1fr; } .priority-spain-form { grid-template-columns:1fr; } }
      #appHeader {
        min-height:62px;
      }
      #appHeader h1 { display:none !important; }
      .route-cost-alert { border-left:4px solid var(--warning); }
      .route-cost-alert-list { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:8px; }
      .route-cost-alert-item { border:1px solid var(--border); border-radius:8px; padding:10px 12px; background:#FFF8EC; }
      .route-cost-alert-item strong { display:block; color:var(--text); font-size:13px; margin-bottom:3px; }
      .route-cost-alert-item span { display:block; color:var(--muted); font-size:12px; }
      .route-evidence-report { display:flex; flex-direction:column; gap:14px; }
      .route-evidence-day { border:1px solid var(--border); border-radius:8px; overflow:hidden; background:var(--surface, #fff); }
      .route-evidence-day-title { display:flex; align-items:center; justify-content:space-between; gap:10px; padding:10px 12px; background:var(--surface-2, #F8FAFC); border-bottom:1px solid var(--border); color:var(--text); }
      .route-evidence-day-title span { color:var(--muted); font-size:12px; font-weight:700; }
      .route-evidence-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:10px; padding:12px; }
      .route-evidence-card { border:1px solid var(--border); border-radius:8px; overflow:hidden; background:var(--surface, #fff); display:flex; flex-direction:column; min-height:100%; }
      .route-evidence-card img, .route-evidence-placeholder { width:100%; height:150px; object-fit:cover; background:var(--surface-2, #F8FAFC); border-bottom:1px solid var(--border); }
      .route-evidence-placeholder { display:flex; align-items:center; justify-content:center; color:var(--muted); font-size:12px; font-weight:800; }
      .route-evidence-info { display:flex; flex-direction:column; gap:4px; padding:10px 12px; font-size:12px; color:var(--muted); line-height:1.35; }
      .route-evidence-info strong { color:var(--text); font-size:13px; }
      .route-evidence-thumb { width:74px; height:52px; object-fit:cover; border-radius:6px; border:1px solid var(--border); display:block; }
      .import-view-shell { align-content:start; }
      .import-table-mode > #sapEstadoBanner,
      .import-table-mode > .card:not(#importExcelTableCard) { display:none !important; }
      .import-table-mode > #importExcelTableCard { display:block !important; width:100%; }
      .import-excel-card { border:1px solid var(--border) !important; box-shadow:var(--shadow-sm, 0 1px 2px rgba(15,23,42,0.05)); padding:0; overflow:hidden; background:var(--surface, #fff); }
      .import-excel-header { display:flex; align-items:center; justify-content:space-between; gap:12px; padding:14px 16px; border-bottom:1px solid var(--border); background:var(--surface, #fff); }
      .import-excel-subtitle { font-size:12px; color:var(--muted); margin-top:2px; }
      .import-excel-meta { font-size:11px; font-weight:800; color:var(--primary); border:1px solid var(--border); border-radius:999px; padding:5px 10px; background:var(--surface-2, #F8FAFC); white-space:nowrap; }
      .import-excel-wrap { overflow-x:auto; }
      .import-excel-table { width:100%; border-collapse:collapse; table-layout:fixed; font-size:12px; }
      .import-excel-table th { text-align:left; background:var(--surface-2, #F8FAFC); color:#334155; font-size:10px; letter-spacing:0; text-transform:uppercase; padding:9px 10px; border-bottom:1px solid var(--border); }
      .import-excel-table th:nth-child(1), .import-excel-table td:nth-child(1) { width:28%; }
      .import-excel-table th:nth-child(2), .import-excel-table td:nth-child(2) { width:28%; }
      .import-excel-table th:nth-child(3), .import-excel-table td:nth-child(3) { width:28%; }
      .import-excel-table th:nth-child(4), .import-excel-table td:nth-child(4) { width:16%; }
      .import-excel-table td { padding:10px; border-bottom:1px solid var(--border); vertical-align:middle; color:var(--text); overflow-wrap:anywhere; }
      .import-excel-table tr:last-child td { border-bottom:none; }
      .import-excel-table td:first-child strong { display:block; font-size:12px; color:var(--primary); margin-bottom:3px; }
      .import-excel-table td:first-child span,
      .import-excel-table td[id^=importTable] span:not(.import-table-badge) { display:block; color:var(--muted); font-size:11px; line-height:1.3; }
      .import-table-badge { display:inline-flex; align-items:center; border-radius:999px; padding:3px 7px; font-size:10px; font-weight:800; margin-bottom:5px; border:1px solid var(--border); color:var(--muted); background:#fff; }
      .import-table-badge.ok { color:#166534; background:#F0FDF4; border-color:#BBF7D0; }
      .import-table-badge.wait { color:#92400E; background:#FFFBEB; border-color:#FDE68A; }
      .import-table-badge.info { color:#1D4ED8; background:#EFF6FF; border-color:#BFDBFE; }
      .import-table-actions { display:flex; gap:5px; flex-wrap:wrap; }
      .import-table-actions .btn, .import-excel-table td:nth-child(4) .btn { width:100%; justify-content:center; padding:6px 8px; font-size:11px; }
      .import-excel-foot { display:flex; align-items:center; justify-content:space-between; gap:10px; padding:10px 16px; background:var(--surface-2, #F8FAFC); border-top:1px solid var(--border); font-size:12px; color:var(--muted); flex-wrap:wrap; }
      .import-excel-foot span:first-child { color:#C2410C; font-weight:800; }
      @media (max-width: 900px) {
        .import-excel-header { align-items:flex-start; flex-direction:column; }
      }
      .calendar-toolbar { flex-wrap: wrap; gap: 10px; }
      .calendar-topbar {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:12px;
        margin:0 0 10px 2px;
        flex-wrap:wrap;
      }
      .calendar-selection-bar { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
      .calendar-month-label { font-size:13px; font-weight:700; color:var(--muted); margin-right:4px; align-self:center; }
      .month-filter { border-radius:12px; padding:8px 16px; font-size:13px; }
      .cal-month-section { margin-bottom:18px; border:1px solid var(--border); border-radius:8px; overflow:hidden; background:var(--surface, #fff); box-shadow:var(--shadow-sm, 0 1px 2px rgba(15,23,42,0.05)); }
      .cal-week-title { background:var(--primary); color:#fff; font-size:14px; font-weight:800; padding:9px 12px; display:flex; align-items:center; justify-content:space-between; gap:10px; }
      .cal-week-title span { font-size:12px; font-weight:700; opacity:.9; }
      .cal-month-section .cal-semana { margin:0; padding:0; gap:0; }
      .cal-month-section .cal-dia-col { border-right:1px solid var(--border); padding:0 6px 10px; }
      .cal-month-section .cal-dia-header { background:var(--surface-2, #F8FAFC); color:var(--text); border-bottom:1px solid var(--border); padding:7px 6px; display:flex; align-items:center; justify-content:space-between; gap:6px; min-height:34px; }
      .cal-month-section .cal-dia-header strong { font-size:12px; }
      .cal-month-section .cal-dia-header small { font-size:10px; color:var(--muted); font-weight:600; }
      .cal-month-section .cal-dia-col:last-child { border-right:none; }
      .calendar-pdf-stage {
        position:fixed;
        left:0;
        top:0;
        width:1800px;
        max-width:none;
        background:#fff;
        color:#111827;
        padding:18px;
        z-index:2147483647;
        pointer-events:none;
      }
      .calendar-pdf-title {
        display:flex;
        align-items:flex-end;
        justify-content:space-between;
        gap:16px;
        margin-bottom:10px;
        color:#1B4F72;
        font-size:24px;
      }
      .calendar-pdf-title span { color:#475569; font-size:16px; font-weight:700; }
      .calendar-pdf-export .calendar-topbar { margin-bottom:6px; }
      .calendar-pdf-export .calendar-topbar > div:first-child { font-size:11px !important; }
      .calendar-pdf-export .cal-month-section { margin-bottom:8px; border-radius:6px; break-inside:avoid; }
      .calendar-pdf-export .cal-week-title { font-size:13px; padding:7px 10px; }
      .calendar-pdf-export .cal-dia-header { font-size:11px; padding:6px 4px; }
      .calendar-pdf-export .cal-dia-header small { font-size:9px; }
      .calendar-pdf-export .cal-dia-col { padding:0 4px 6px; }
      .calendar-pdf-export .cal-carril { margin-top:5px; }
      .calendar-pdf-export .cal-carril-header { font-size:9px; padding:3px 5px; }
      .calendar-pdf-export .cal-card {
        padding:5px;
        margin-top:4px;
        border-radius:5px;
        box-shadow:none;
        min-height:0;
      }
      .calendar-pdf-export .cal-card-name { font-size:9px; line-height:1.15; }
      .calendar-pdf-export .queue-card-meta,
      .calendar-pdf-export .cal-card-monto { font-size:8px; line-height:1.2; }
      .calendar-pdf-export .cal-card-topline,
      .calendar-pdf-export .cal-card-actions,
      .calendar-pdf-export .calendar-selection-bar,
      .calendar-pdf-export button,
      .calendar-pdf-export input { display:none !important; }
      .cal-card-topline {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:8px;
        margin-bottom:4px;
      }
      .cal-card-check {
        display:flex;
        align-items:center;
        gap:6px;
        font-size:10px;
        color:var(--muted);
      }
      .cal-card-check input { accent-color: var(--primary); }
      .cal-card-selected {
        background: #FFF8D8;
        box-shadow: 0 0 0 2px rgba(245,197,66,0.55);
      }
      .calendar-notes-wrap {
        margin:8px 0 10px;
        display:grid;
        gap:6px;
      }
      .calendar-notes-head {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:6px;
        font-size:10px;
        color:var(--muted);
        font-weight:800;
        text-transform:uppercase;
      }
      .calendar-note-add {
        border:1px solid #FACC15;
        background:#FEF3C7;
        color:#713F12;
        border-radius:999px;
        padding:3px 7px;
        font-size:10px;
        font-weight:800;
        cursor:pointer;
      }
      .calendar-note-list {
        display:grid;
        gap:5px;
      }
      .calendar-note-postit {
        text-align:left;
        border:1px solid #FDE68A;
        border-left:4px solid #F59E0B;
        border-radius:7px;
        background:#FFFBEB;
        color:#3F2B08;
        padding:7px 8px;
        box-shadow:0 1px 2px rgba(146,64,14,.08);
        cursor:pointer;
        width:100%;
      }
      .calendar-note-postit span {
        display:block;
        font-size:11px;
        line-height:1.35;
        font-weight:700;
        white-space:normal;
        overflow-wrap:anywhere;
      }
      .calendar-note-postit small {
        display:block;
        margin-top:4px;
        font-size:9px;
        color:#92400E;
        font-weight:800;
        text-transform:uppercase;
      }
      .commercial-toolbar {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:10px;
        flex-wrap:wrap;
        margin-bottom:12px;
      }
      .commercial-quick,
      .commercial-actions { display:flex; align-items:center; gap:6px; flex-wrap:wrap; }
      .commercial-quick span { font-size:11px; font-weight:700; color:var(--muted); }
      .commercial-filter-grid {
        display:grid;
        grid-template-columns: minmax(220px, 1.4fr) repeat(4, minmax(140px, 1fr));
        gap:8px;
        margin-bottom:12px;
      }
      .commercial-calendar { display:flex; flex-direction:column; gap:14px; }
      .commercial-day {
        border:1px solid var(--border);
        border-radius:10px;
        overflow:hidden;
        background:var(--surface, #fff);
        box-shadow:var(--shadow-sm, 0 1px 2px rgba(15,23,42,0.05));
      }
      .commercial-day-head {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:12px;
        padding:10px 12px;
        background:var(--primary);
        color:#fff;
        font-size:13px;
        font-weight:700;
      }
      .commercial-day-head span { display:block; font-size:11px; font-weight:500; opacity:.86; margin-top:2px; }
      .commercial-client-grid { display:grid; grid-template-columns:repeat(auto-fit, minmax(420px, 1fr)); gap:10px; padding:10px; }
      .commercial-client-card { border:1px solid var(--border); border-radius:8px; background:var(--surface, #fff); overflow:hidden; box-shadow:var(--shadow-sm, 0 1px 2px rgba(15,23,42,0.05)); }
      .commercial-client-head {
        display:flex;
        align-items:flex-start;
        justify-content:space-between;
        gap:10px;
        padding:10px 12px;
        background:var(--surface-2, #F8FAFC);
        border-bottom:1px solid var(--border);
        font-size:12px;
      }
      .commercial-client-head strong { color:var(--primary); font-size:13px; }
      .commercial-client-head span { display:block; color:var(--muted); font-size:11px; margin-top:2px; }
      .commercial-client-head > div:last-child { text-align:right; font-weight:800; color:var(--text); white-space:nowrap; }
      .commercial-table-wrap { overflow:auto; }
      .commercial-detail-table { width:100%; border-collapse:collapse; font-size:11px; min-width:760px; }
      .commercial-detail-table th { background:#F8FAFC; color:#334155; text-align:left; padding:7px 8px; border-bottom:1px solid var(--border); font-size:10px; text-transform:uppercase; }
      .commercial-detail-table td { padding:7px 8px; border-bottom:1px solid var(--border); vertical-align:top; }
      .commercial-detail-table tr:last-child td { border-bottom:none; }
      .commercial-detail-table th:nth-child(4),
      .commercial-detail-table td:nth-child(4) { min-width:260px; }
      .commercial-section-label { padding:8px 12px; font-size:11px; font-weight:800; text-transform:uppercase; color:#166534; background:#F0FDF4; border-bottom:1px solid var(--border); }
      .commercial-section-label.muted { color:#6B7280; background:#F8FAFC; border-top:1px solid var(--border); }
      .commercial-muted-table { opacity:.74; }
      .commercial-muted-table .commercial-detail-table td { color:var(--muted); }
      .commercial-pdf-stage {
        position:fixed;
        left:-10000px;
        top:0;
        width:900px;
        background:#fff;
        color:#111827;
        padding:20px;
        z-index:-1;
      }
      .commercial-pdf-title {
        display:flex;
        align-items:flex-end;
        justify-content:space-between;
        gap:16px;
        margin-bottom:12px;
        color:#1B4F72;
        font-size:22px;
      }
      .commercial-pdf-title span { color:#475569; font-size:12px; font-weight:700; }
      .commercial-pdf-export .commercial-day { margin-bottom:12px; break-inside:avoid; }
      .commercial-pdf-export .commercial-client-grid { display:block; padding:8px; }
      .commercial-pdf-export .commercial-client-card { margin-bottom:8px; break-inside:avoid; }
      .commercial-pdf-export .commercial-detail-table { min-width:0; font-size:8px; }
      .commercial-pdf-export .commercial-detail-table th,
      .commercial-pdf-export .commercial-detail-table td { padding:4px; }
      .commercial-pdf-export .commercial-detail-table th:nth-child(3),
      .commercial-pdf-export .commercial-detail-table td:nth-child(3) { min-width:180px; }
      @media (max-width: 980px) {
        .commercial-filter-grid { grid-template-columns:1fr 1fr; }
        .commercial-client-grid { grid-template-columns:1fr; }
      }
      @media (max-width: 620px) {
        .commercial-filter-grid { grid-template-columns:1fr; }
        .commercial-day-head,
        .commercial-client-head { align-items:flex-start; flex-direction:column; }
        .commercial-client-head > div:last-child { text-align:left; }
      }

      .route-catalog-board { display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:12px; align-items:start; }
      .route-catalog-col { border:1px solid var(--border); border-radius:8px; background:var(--surface, #fff); overflow:hidden; min-height:140px; box-shadow:var(--shadow-sm); }
      .route-catalog-col.drag-over { box-shadow:0 0 0 2px rgba(0,166,118,.25); border-color:var(--secondary); }
      .route-catalog-col.pending { border-color:#FDE68A; }
      .route-catalog-head { display:flex; justify-content:space-between; gap:10px; align-items:flex-start; padding:10px 12px; background:var(--primary); color:#fff; }
      .route-catalog-head strong { font-size:13px; overflow-wrap:anywhere; }
      .route-catalog-head span { font-size:11px; font-weight:700; opacity:.86; white-space:nowrap; }
      .route-catalog-col.pending .route-catalog-head { background:#92400E; }
      .route-catalog-body { display:grid; gap:8px; padding:10px; }
      .route-client-card { width:100%; border:1px solid var(--border); background:var(--surface, #fff); border-radius:8px; padding:9px 10px; display:flex; justify-content:space-between; gap:10px; text-align:left; cursor:pointer; color:var(--text); }
      .route-client-card:hover { border-color:var(--secondary); box-shadow:0 1px 4px rgba(15,23,42,.10); }
      .route-client-card span { min-width:0; }
      .route-client-card strong { display:block; font-size:12px; overflow-wrap:anywhere; }
      .route-client-card small { display:block; color:var(--muted); font-size:11px; margin-top:3px; }
      .route-client-card em { align-self:flex-start; border-radius:999px; padding:3px 7px; font-size:10px; font-style:normal; font-weight:800; background:#F0FDF4; color:#166534; white-space:nowrap; }
      .route-card-actions { display:flex; flex-direction:column; align-items:flex-end; gap:6px; flex:0 0 150px; }
      .route-assign-select { width:150px; max-width:100%; border:1px solid var(--border); border-radius:8px; padding:6px 8px; font-size:11px; background:var(--surface, #fff); color:var(--text); }
      .route-client-card.pending { background:#FFFBEB; border-color:#FDE68A; }
      .route-client-card.pending em { background:#FEF3C7; color:#92400E; }
      .route-catalog-empty { color:var(--muted); font-size:12px; padding:8px; text-align:center; }
      .control-history-head { display:flex; justify-content:space-between; align-items:center; gap:10px; margin-bottom:8px; }
      .control-history-panel { margin-bottom:16px; }
      .control-history-table { min-width:720px; }
      .routes-view-shell #rutasContainer::before {
        content: 'Exporta la fecha activa a Excel con un resumen por carril y el detalle por ruta.';
        display:block;
        margin-bottom:12px;
        padding:12px 14px;
        border:1px solid var(--border);
        border-radius:12px;
        background:#fff;
        color:var(--muted);
        font-size:12px;
      }
      @media (max-width: 760px) {
        #calMesNav {
          display:flex;
          gap:6px;
          overflow-x:auto;
          flex-wrap:nowrap;
          padding-bottom:6px;
          -webkit-overflow-scrolling:touch;
        }
        #calMesNav .month-filter,
        #calMesNav .filter-btn {
          flex:0 0 auto;
          white-space:nowrap;
          padding:7px 10px;
          font-size:12px;
        }
        .calendar-topbar {
          align-items:stretch;
          flex-direction:column;
          gap:8px;
          margin-left:0;
        }
        .calendar-selection-bar {
          display:grid;
          grid-template-columns:1fr 1fr;
          gap:8px;
        }
        .calendar-selection-bar .badge {
          grid-column:1 / -1;
          text-align:center;
          padding:6px 8px;
        }
        .calendar-selection-bar .btn {
          width:100%;
          justify-content:center;
        }
        .cal-month-section {
          border-radius:8px;
          overflow:hidden;
          margin-bottom:14px;
        }
        .cal-week-title {
          align-items:flex-start;
          flex-direction:column;
          gap:2px;
          font-size:13px;
        }
        .cal-month-section .cal-semana {
          display:grid !important;
          grid-template-columns:1fr !important;
        }
        .cal-month-section .cal-dia-col {
          border-right:none !important;
          border-bottom:1px solid var(--border);
          padding:0 8px 10px;
        }
        .cal-month-section .cal-dia-col:last-child { border-bottom:none; }
        .cal-month-section .cal-dia-header {
          position:sticky;
          top:0;
          z-index:2;
          margin:0 -8px 8px;
          padding:8px 10px;
        }
        .cal-carril {
          margin-top:8px;
          border:1px solid var(--border);
          border-radius:8px;
          padding:8px;
          background:#fff;
        }
        .cal-carril-header {
          display:block;
          text-align:center;
          margin:-8px -8px 8px;
          border-radius:8px 8px 0 0;
          padding:7px 8px;
        }
        .cal-card {
          min-height:0;
          padding:10px !important;
          margin-top:8px;
          border-radius:8px;
        }
        .cal-card-name {
          font-size:13px;
          line-height:1.25;
          overflow-wrap:anywhere;
        }
        .cal-card-actions {
          display:grid;
          grid-template-columns:1fr 1fr 44px;
          gap:6px;
          align-items:center;
        }
        .cal-card-actions .btn {
          width:100%;
          justify-content:center;
        }
        .queue-route-header,
        .import-excel-header,
        .commercial-toolbar,
        .commercial-quick,
        .commercial-actions,
        .theme-settings-head {
          align-items:stretch;
          flex-direction:column;
        }
        .queue-route-header .btn,
        .commercial-actions .btn,
        .commercial-quick .filter-btn {
          width:100%;
          justify-content:center;
        }
        .queue-card {
          border-radius:8px;
          padding:10px;
        }
        .route-add-form,
        .role-permission-toolbar,
        .commercial-filter-grid {
          grid-template-columns:1fr !important;
        }
        .data-table input.form-control,
        .data-table select.form-control {
          min-width:120px !important;
        }
        .role-permission-table {
          min-width:520px;
        }
        .permission-chip {
          white-space:normal;
          overflow-wrap:anywhere;
        }
        .import-excel-table {
          min-width:0;
          table-layout:auto;
        }
        .import-excel-table thead { display:none; }
        .import-excel-table,
        .import-excel-table tbody,
        .import-excel-table tr,
        .import-excel-table td {
          display:block;
          width:100% !important;
        }
        .import-excel-table tr {
          border-bottom:1px solid var(--border);
          padding:8px 0;
        }
        .import-excel-table td {
          border-bottom:none;
          padding:7px 10px;
        }
        .import-excel-table td:nth-child(4) .btn,
        .import-table-actions .btn {
          width:100%;
        }
        .commercial-client-grid {
          grid-template-columns:1fr;
          padding:8px;
        }
        .commercial-day-head,
        .commercial-client-head {
          align-items:flex-start;
          flex-direction:column;
        }
        .commercial-client-head > div:last-child {
          text-align:left;
          white-space:normal;
        }
        .commercial-table-wrap {
          margin-left:-8px;
          margin-right:-8px;
          -webkit-overflow-scrolling:touch;
        }
        .commercial-detail-table {
          min-width:760px;
        }
        html,
        body {
          width:100%;
          max-width:100%;
          overflow-x:hidden;
        }
        #mainLayout,
        #content,
        .view,
        .card {
          min-width:0;
          max-width:100%;
        }
        #sidebar {
          position:fixed !important;
          left:0 !important;
          right:0 !important;
          bottom:0 !important;
          width:100% !important;
          height:92px !important;
          min-height:92px !important;
          max-height:92px !important;
          display:flex !important;
          flex-direction:row !important;
          gap:8px !important;
          padding:8px 10px calc(8px + env(safe-area-inset-bottom)) !important;
          overflow-x:auto !important;
          overflow-y:hidden !important;
          -webkit-overflow-scrolling:touch;
        }
        .nav-btn {
          flex:0 0 118px !important;
          min-width:118px !important;
          width:118px !important;
          min-height:70px !important;
          height:70px !important;
          padding:8px 8px !important;
          font-size:12px !important;
          flex-direction:column;
          gap:4px;
          text-align:center;
          white-space:normal;
        }
        .nav-btn span.nav-label {
          display:block !important;
          line-height:1.12;
        }
        .nav-icon::before {
          font-size:20px !important;
        }
        .brand-header-logo {
          width:154px;
          min-width:132px;
          height:34px;
          order:0;
        }
        #appHeader h1 { display:none !important; }
        #mainLayout {
          padding-bottom:104px !important;
        }
        #content,
        .view.active {
          min-height:calc(100vh - 170px);
        }
        .password-field-wrap {
          display:grid;
          grid-template-columns:1fr auto;
        }
        .route-cost-alert-list {
          grid-template-columns:1fr;
        }
        .theme-option-grid {
          grid-template-columns:1fr !important;
        }
      }
      @media (max-width: 420px) {
        .calendar-selection-bar,
        .cal-card-actions {
          grid-template-columns:1fr;
        }
        .cal-card-actions span {
          text-align:center;
        }
        .commercial-detail-table,
        .data-table {
          min-width:680px;
        }
      }
      body.theme-light {
        --primary:#111827;
        --secondary:#00A676;
        --warning:#F5C542;
        --danger:#B42318;
        --bg:#EEF2F6;
        --card:#FFFFFF;
        --border:#D6DEE8;
        --text:#17212B;
        --muted:#667085;
        --soft:#F8FAFC;
        --header:#0B1220;
        --icon:#17212B;
        --surface:#FFFFFF;
        --surface-2:#F8FAFC;
        --shadow-sm:0 1px 2px rgba(15,23,42,0.05);
      }
      body.theme-dark {
        --primary:#D7E7F5;
        --secondary:#7FD49B;
        --warning:#F7C948;
        --danger:#FF9B8D;
        --bg:#0F141A;
        --card:#171E26;
        --border:#2A3440;
        --text:#E7EDF3;
        --muted:#A5B1BE;
        --soft:#111820;
        --header:#101820;
        --icon:#F8FAFC;
        --surface:#171E26;
        --surface-2:#111820;
        --shadow-sm:0 1px 2px rgba(0,0,0,0.22);
      }
      body.theme-brand {
        --primary:#161616;
        --secondary:#1F7A3A;
        --warning:#C28A00;
        --danger:#B42318;
        --bg:#F7F6F1;
        --card:#FFFFFF;
        --border:#DDD8C8;
        --text:#181818;
        --muted:#68635B;
        --soft:#FFF7D6;
        --header:#F4C900;
        --icon:#111111;
        --surface:#FFFFFF;
        --surface-2:#FFFDF3;
        --shadow-sm:0 1px 2px rgba(85,70,0,0.08);
      }
      body.theme-light {
        background:var(--bg);
      }
      body.theme-dark {
        background:var(--bg);
      }
      body.theme-brand {
        background:var(--bg);
      }
      body.theme-dark,
      body.theme-dark #content { color:var(--text); }
      body.theme-dark .card,
      body.theme-dark .stat-box,
      body.theme-dark .kpi-card,
      body.theme-dark .chart-card,
      body.theme-dark .queue-card,
      body.theme-dark .sol-card,
      body.theme-dark .commercial-day,
      body.theme-dark .commercial-client-card,
      body.theme-dark .config-client-list,
      body.theme-dark .config-client-item,
      body.theme-dark .import-excel-card,
      body.theme-dark .import-excel-header { background:var(--surface) !important; color:var(--text); border-color:var(--border) !important; }
      body.theme-dark .search-input,
      body.theme-dark .form-control,
      body.theme-dark .nota-input,
      body.theme-dark input,
      body.theme-dark select,
      body.theme-dark textarea { background:#111820 !important; color:var(--text) !important; border-color:var(--border) !important; }
      body.theme-dark .data-table td,
      body.theme-dark .commercial-detail-table td,
      body.theme-dark .import-excel-table td { background:rgba(23,30,38,0.62); color:var(--text); border-color:var(--border); }
      body.theme-dark .data-table th,
      body.theme-dark .commercial-detail-table th,
      body.theme-dark .import-excel-table th,
      body.theme-dark .com-th { background:#222C36 !important; color:#F8FAFC !important; }
      body.theme-dark .cal-month-section,
      body.theme-dark .cal-dia-col,
      body.theme-dark .cal-card { background:var(--surface) !important; border-color:var(--border) !important; color:var(--text); }
      body.theme-dark .cal-dia-header,
      body.theme-dark .commercial-day-head { background:#111820 !important; color:var(--text); border-color:var(--border); }
      body.theme-brand #appHeader { background:var(--header); color:#111; border-bottom:1px solid #D7B400; }
      body.theme-brand #saveStateBtn,
      body.theme-brand #sapHeaderInfo { color:#111; border-color:rgba(0,0,0,.24); background:rgba(255,255,255,.36); }
      #appHeader .logo { font-size:20px; color:currentColor; filter:grayscale(1); }
      body.theme-brand .btn-primary,
      body.theme-brand .data-table th,
      body.theme-brand .cal-week-title,
      body.theme-brand .com-th { background:#161616 !important; color:#fff !important; }
      body.theme-brand .nav-btn.active { background:#161616; color:#fff; }
      body.theme-brand .stat-value,
      body.theme-brand .kpi-val,
      body.theme-brand .card-title,
      body.theme-brand .view-title { color:#161616; }
      body.theme-dark #appHeader { background:var(--header); color:var(--text); border-color:var(--border); }
      body.theme-dark #sidebar { background:var(--surface); color:var(--text); border-color:var(--border); }
      body.theme-dark .nav-btn { color:var(--text); }
      body.theme-dark .nav-btn:hover { background:#1D2630; }
      body.theme-dark .nav-btn.active { background:#E7EDF3; color:#101820; }
      body.theme-dark .btn-outline { background:#111820; color:var(--text); border-color:var(--border); }
      body.theme-light #appHeader { background:var(--header); }
      body.theme-light #sidebar { background:#0F172A; border-color:#1E293B; }
      body.theme-light .nav-btn { color:#D7DEE8; }
      body.theme-light .nav-btn:hover { background:#182235; color:#fff; }
      body.theme-light .nav-btn.active { background:var(--warning); color:#111827; }
      body.theme-brand #sidebar { background:var(--surface); }
      .nav-icon {
        width:22px;
        display:inline-flex;
        justify-content:center;
        color:currentColor;
        font-size:0 !important;
        line-height:1;
      }
      .nav-icon::before { font-size:17px; font-weight:800; color:currentColor; }
      #nav-importar .nav-icon::before { content:'↓'; }
      #nav-dashboard .nav-icon::before { content:'▥'; }
      #nav-calendario .nav-icon::before { content:'□'; }
      #nav-comercial .nav-icon::before { content:'$'; }
      #nav-rutas .nav-icon::before { content:'⇄'; }
      #nav-importaciones .nav-icon::before { content:'↗'; }
      #nav-prioridades .nav-icon::before { content:'!'; }
      #nav-configClientes .nav-icon::before { content:'⚙'; }
      #nav-solicitudesAlmacen .nav-icon::before { content:'▤'; }
      #nav-almacen .nav-icon::before { content:'▦'; }
      .theme-settings-card { padding:16px; }
      .theme-settings-head { display:flex; align-items:flex-start; justify-content:space-between; gap:12px; margin-bottom:12px; }
      .theme-settings-sub { color:var(--muted); font-size:12px; margin-top:-6px; }
      .theme-option-grid { display:grid; grid-template-columns:repeat(3, minmax(0, 1fr)); gap:10px; }
      .theme-option {
        background:var(--surface);
        border:1px solid var(--border);
        color:var(--text);
        border-radius:10px;
        padding:10px;
        text-align:left;
        cursor:pointer;
        display:grid;
        gap:5px;
        box-shadow:var(--shadow-sm);
      }
      .theme-option:hover, .theme-option.active { border-color:var(--primary); box-shadow:0 0 0 2px color-mix(in srgb, var(--primary) 16%, transparent); }
      .theme-option strong { font-size:13px; }
      .theme-option small { color:var(--muted); font-size:11px; line-height:1.35; }
      .theme-preview { height:34px; border-radius:7px; border:1px solid var(--border); display:block; }
      .theme-preview-light { background:linear-gradient(90deg,#123F5D 0 24%,#fff 24% 66%,#F5F6F8 66%); }
      .theme-preview-dark { background:linear-gradient(90deg,#101820 0 24%,#171E26 24% 66%,#0F141A 66%); }
      .theme-preview-brand { background:linear-gradient(90deg,#F4C900 0 24%,#fff 24% 66%,#161616 66%); }
      @media (max-width: 760px) { .route-add-form, .role-permission-toolbar { grid-template-columns:1fr; } }
      @media (max-width: 760px) { .theme-option-grid { grid-template-columns:1fr; } }
      details summary { list-style:none; }
      details summary::-webkit-details-marker { display:none; }
    `;
    document.head.appendChild(style);
  }


  const IMPORT_STAGES = [
    { key: 'pedido', label: 'Pedido / fabricación' },
    { key: 'docs_origen', label: 'Docs origen' },
    { key: 'miami_consolidacion', label: 'Miami / consolidación' },
    { key: 'transito', label: 'Tránsito a RD' },
    { key: 'aduana_rd', label: 'Aduana RD' },
    { key: 'transporte_almacen', label: 'Ruta a almacén' },
    { key: 'recepcion_sap', label: 'Recepción / SAP' }
  ];

  const IMPORT_DOCS = {
    espana: [
      ['eur1', 'EUR1'], ['seguro', 'Seguro'], ['waybill', 'Waybill / BL'], ['facturaProveedor', 'Factura proveedor'], ['packingList', 'Packing list'], ['preliquidacion', 'Preliquidación'], ['pin', 'PIN pagado'], ['liquidacion', 'Liquidación pagada'], ['facturaSap', 'Factura proveedor SAP'], ['gastosSap', 'Gastos SAP']
    ],
    miami: [
      ['po', 'PO'], ['trackingUsa', 'Tracking proveedor-Miami'], ['reciboMiami', 'Recibido en Miami'], ['consolidacionMiami', 'Consolidación Miami'], ['facturaProveedor', 'Factura proveedor'], ['packingList', 'Packing list'], ['bl', 'BL / guía marítima'], ['preliquidacion', 'Preliquidación'], ['pin', 'PIN pagado'], ['liquidacion', 'Liquidación pagada'], ['facturaSap', 'Factura proveedor SAP'], ['gastosSap', 'Gastos SAP']
    ]
  };


  const IMPORT_BULK_HEADERS = [
    'EXPEDIENTE', 'Nº BOOKING', 'PARTNER', 'DESTINO', 'POA', 'VESSEL', 'ETD', 'ETA', 'EQUIPO', 'FECHA CARGA', 'HORA CARGA', 'SITUACION', 'POD', 'Nº CONTENEDOR'
  ];

  const IMPORT_BULK_ALIASES = {
    expediente: ['EXPEDIENTE', 'Expediente', 'REF', 'Referencia'],
    booking: ['Nº BOOKING', 'N° BOOKING', 'NO BOOKING', 'BOOKING', 'N BOOKING', 'Numero Booking'],
    partner: ['PARTNER', 'Transitario', 'Agente'],
    destino: ['DESTINO', 'Destino'],
    poa: ['POA'],
    vessel: ['VESSEL', 'Buque', 'Barco'],
    etd: ['ETD'],
    eta: ['ETA'],
    equipo: ['EQUIPO', 'Equipo'],
    fechaCarga: ['FECHA CARGA', 'Fecha carga', 'Fecha de carga'],
    horaCarga: ['HORA CARGA', 'Hora carga', 'Hora de carga'],
    situacion: ['SITUACION', 'SITUACIÓN', 'Estado', 'Status'],
    pod: ['POD'],
    contenedor: ['Nº CONTENEDOR', 'N° CONTENEDOR', 'NO CONTENEDOR', 'CONTENEDOR', 'Container']
  };

  function getBulkAliasValue(row, field) {
    const aliases = IMPORT_BULK_ALIASES[field] || [field];
    for (const alias of aliases) {
      const direct = row && row[alias];
      if (direct !== undefined && direct !== null && text(direct)) return text(direct);
      const normalized = normalizeRow(row || {});
      const key = normKey(alias);
      if (normalized[key] !== undefined && normalized[key] !== null && text(normalized[key])) return text(normalized[key]);
    }
    return '';
  }

  function parseImportBulkDate(value) {
    const raw = text(value).replace(/\.$/, '');
    if (!raw || raw === '-') return '';
    const slash = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
    if (slash) {
      const year = slash[3].length === 2 ? '20' + slash[3] : slash[3];
      return `${year}-${String(slash[2]).padStart(2, '0')}-${String(slash[1]).padStart(2, '0')}`;
    }
    const monthMap = {
      ene: 1, enero: 1, jan: 1,
      feb: 2, febrero: 2,
      mar: 3, marzo: 3,
      abr: 4, abril: 4, apr: 4,
      may: 5, mayo: 5,
      jun: 6, junio: 6,
      jul: 7, julio: 7,
      ago: 8, agosto: 8, aug: 8,
      sep: 9, sept: 9, septiembre: 9,
      oct: 10, octubre: 10,
      nov: 11, noviembre: 11,
      dic: 12, diciembre: 12, dec: 12
    };
    const monthText = stripAccents(raw.toLowerCase()).match(/^(\d{1,2})\s*[-/ ]\s*([a-z]+)(?:\s*[-/ ]\s*(\d{2,4}))?$/);
    if (monthText) {
      const month = monthMap[monthText[2]];
      if (!month) return raw;
      const year = monthText[3] ? (monthText[3].length === 2 ? '20' + monthText[3] : monthText[3]) : String(CURRENT_YEAR);
      return `${year}-${String(month).padStart(2, '0')}-${String(monthText[1]).padStart(2, '0')}`;
    }
    return raw;
  }

  function normalizeImportSituation(value) {
    const status = stripAccents(text(value).toLowerCase());
    if (status.includes('anulado')) return 'Anulado';
    if (status.includes('pdte embarque')) return 'Pdte embarque';
    if (status.includes('pdte carga')) return 'Pdte Carga';
    if (status.includes('cargado')) return 'Cargado';
    if (status.includes('embarcado')) return 'Embarcado';
    return text(value);
  }

  function importStageFromSituation(value) {
    const status = stripAccents(text(value).toLowerCase());
    if (status.includes('embarcado')) return 'transito';
    if (status.includes('cargado') || status.includes('pdte embarque') || status.includes('pdte carga')) return 'docs_origen';
    if (status.includes('anulado')) return 'pedido';
    return 'docs_origen';
  }

  function normalizeBulkShipment(raw) {
    const expediente = text(raw.expediente || raw.referencia);
    const booking = text(raw.booking || raw.po);
    const contenedor = text(raw.contenedor || raw.bl);
    if (!expediente && !booking && !contenedor) return null;
    const situacion = normalizeImportSituation(raw.situacion);
    const eta = parseImportBulkDate(raw.eta);
    const etd = parseImportBulkDate(raw.etd);
    const fechaCarga = parseImportBulkDate(raw.fechaCarga);
    const id = 'imp-bulk-' + normKey([expediente, contenedor || booking].filter(Boolean).join('-'));
    const detailParts = [
      raw.destino ? 'Destino: ' + raw.destino : '',
      raw.poa ? 'POA: ' + raw.poa : '',
      raw.vessel ? 'Vessel: ' + raw.vessel : '',
      etd ? 'ETD: ' + etd : '',
      raw.equipo ? 'Equipo: ' + raw.equipo : '',
      fechaCarga ? 'Carga: ' + fechaCarga + (raw.horaCarga ? ' ' + raw.horaCarga : '') : '',
      situacion ? 'Situación: ' + situacion : '',
      raw.pod ? 'POD: ' + raw.pod : '',
      contenedor ? 'Contenedor: ' + contenedor : ''
    ].filter(Boolean);
    return {
      id,
      origen: 'espana',
      etapa: importStageFromSituation(situacion),
      referencia: expediente || booking || contenedor,
      proveedor: text(raw.partner) || 'RAMINATRANS',
      po: booking,
      factura: '',
      bl: contenedor || booking,
      tracking: text(raw.vessel || raw.partner),
      eta,
      fechaMiami: '',
      fechaAduana: '',
      fechaSalidaAduana: '',
      fechaAlmacen: '',
      valor: 0,
      notas: detailParts.join(' · '),
      booking,
      partner: text(raw.partner),
      destino: text(raw.destino),
      poa: text(raw.poa),
      vessel: text(raw.vessel),
      etd,
      equipo: text(raw.equipo),
      fechaCarga,
      horaCarga: text(raw.horaCarga),
      situacion,
      pod: text(raw.pod),
      contenedor,
      docs: { waybill: !!(booking || contenedor), bl: !!(booking || contenedor) },
      source: 'bulk-import',
      updatedAt: nowIso()
    };
  }

  function shipmentKey(item) {
    return normKey([item.referencia, item.contenedor || item.bl, item.booking || item.po].filter(Boolean).join('|'));
  }

  function parseBulkRowsFromObjects(rows) {
    return (rows || []).map(row => normalizeBulkShipment({
      expediente: getBulkAliasValue(row, 'expediente'),
      booking: getBulkAliasValue(row, 'booking'),
      partner: getBulkAliasValue(row, 'partner'),
      destino: getBulkAliasValue(row, 'destino'),
      poa: getBulkAliasValue(row, 'poa'),
      vessel: getBulkAliasValue(row, 'vessel'),
      etd: getBulkAliasValue(row, 'etd'),
      eta: getBulkAliasValue(row, 'eta'),
      equipo: getBulkAliasValue(row, 'equipo'),
      fechaCarga: getBulkAliasValue(row, 'fechaCarga'),
      horaCarga: getBulkAliasValue(row, 'horaCarga'),
      situacion: getBulkAliasValue(row, 'situacion'),
      pod: getBulkAliasValue(row, 'pod'),
      contenedor: getBulkAliasValue(row, 'contenedor')
    })).filter(Boolean);
  }

  function parseCompactShipmentLine(line) {
    const original = text(line).replace(/\s+/g, ' ');
    if (!/^H\d{5,}/i.test(original)) return null;
    const expediente = (original.match(/^H\d{5,}/i) || [''])[0];
    let rest = text(original.slice(expediente.length));
    const bookingMatch = rest.match(/^([A-Z0-9]{5,}|\d{5,})\b/i);
    const booking = bookingMatch ? bookingMatch[1] : '';
    if (booking) rest = text(rest.slice(booking.length));
    const partnerMatch = rest.match(/^(RAMINATRANS|[A-Z][A-Z0-9&. -]{2,}?)(?=\s+ALVAREZ|\s+CLIENTE|\s+[A-Z]+\s+[A-Z]+)/i);
    const partner = partnerMatch ? partnerMatch[1] : 'RAMINATRANS';
    if (partnerMatch) rest = text(rest.slice(partnerMatch[0].length));

    const dateMatches = [...rest.matchAll(/\b\d{1,2}\s*[-/]\s*(?:ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic|jan|apr|aug|dec|\d{1,2})\b/ig)];
    const etdRaw = dateMatches[0] ? dateMatches[0][0] : '';
    const etaRaw = dateMatches[1] ? dateMatches[1][0] : '';
    const beforeDates = etdRaw ? text(rest.slice(0, dateMatches[0].index)) : rest;
    const afterEtaStart = dateMatches[1] ? dateMatches[1].index + dateMatches[1][0].length : (dateMatches[0] ? dateMatches[0].index + dateMatches[0][0].length : 0);
    let tail = afterEtaStart ? text(rest.slice(afterEtaStart)) : '';

    const poaOptions = ['SANTO DOMINGO', 'CAUCEDO', 'MORELOS', 'KINGSTON'];
    let poa = '';
    let poaIndex = -1;
    poaOptions.forEach(option => {
      const idx = stripAccents(beforeDates.toUpperCase()).indexOf(stripAccents(option));
      if (idx >= 0 && (poaIndex < 0 || idx < poaIndex)) {
        poa = option;
        poaIndex = idx;
      }
    });
    const destino = poaIndex >= 0 ? text(beforeDates.slice(0, poaIndex)) : '';
    const vessel = poaIndex >= 0 ? text(beforeDates.slice(poaIndex + poa.length)) : beforeDates;

    const equipoMatch = tail.match(/^(LCL|20|40\s*HC|40HC|40|45\s*HC)\b/i);
    const equipo = equipoMatch ? equipoMatch[1].replace(/\s+/g, ' ').toUpperCase() : '';
    if (equipoMatch) tail = text(tail.slice(equipoMatch[0].length));
    const fechaCargaMatch = tail.match(/^(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}|-)\b/);
    const fechaCarga = fechaCargaMatch ? fechaCargaMatch[1] : '';
    if (fechaCargaMatch) tail = text(tail.slice(fechaCargaMatch[0].length));
    const horaMatch = tail.match(/^(\d{1,2}:\d{2}|-)\b/);
    const horaCarga = horaMatch ? horaMatch[1] : '';
    if (horaMatch) tail = text(tail.slice(horaMatch[0].length));
    const situacionMatch = tail.match(/^(Pdte\s+embarque|Pdte\s+Carga|Embarcado|Cargado|Anulado|Pote\s+Carga)\b/i);
    const situacion = situacionMatch ? situacionMatch[1].replace(/^Pote/i, 'Pdte') : '';
    if (situacionMatch) tail = text(tail.slice(situacionMatch[0].length));
    const podMatch = tail.match(/^(VALENCIA|CAUCEDO|SANTO DOMINGO|KINGSTON|MORELOS)\b/i);
    const pod = podMatch ? podMatch[1].toUpperCase() : '';
    if (podMatch) tail = text(tail.slice(podMatch[0].length));
    const contenedor = tail || ((original.match(/[A-Z]{4}\d{7}/i) || [''])[0]);

    return normalizeBulkShipment({
      expediente,
      booking,
      partner,
      destino,
      poa,
      vessel,
      etd: etdRaw,
      eta: etaRaw,
      equipo,
      fechaCarga,
      horaCarga,
      situacion,
      pod,
      contenedor
    });
  }

  function parseBulkShipmentsFromText(rawText) {
    const lines = String(rawText || '')
      .replace(/\r/g, '\n')
      .split('\n')
      .map(line => text(line).replace(/\s+/g, ' '))
      .filter(Boolean)
      .filter(line => !IMPORT_BULK_HEADERS.some(header => normKey(line) === normKey(header)));
    const compact = lines.map(parseCompactShipmentLine).filter(Boolean);
    if (compact.length >= Math.max(1, Math.floor(lines.filter(line => /^H\d{5,}/i.test(line)).length * 0.6))) return compact;
    const records = [];
    let current = [];
    lines.forEach(line => {
      if (/^H\d{5,}/i.test(line) && current.length) {
        records.push(current);
        current = [];
      }
      current.push(line);
    });
    if (current.length) records.push(current);
    const sequential = records
      .filter(parts => parts.length > 1)
      .map(parts => normalizeBulkShipment({
        expediente: parts[0],
        booking: parts[1],
        partner: parts[2],
        destino: parts[3],
        poa: parts[4],
        vessel: parts[5],
        etd: parts[6],
        eta: parts[7],
        equipo: parts[8],
        fechaCarga: parts[9],
        horaCarga: parts[10],
        situacion: parts[11],
        pod: parts[12],
        contenedor: parts[13]
      })).filter(Boolean);
    if (sequential.length) return sequential;
    return lines.filter(line => /^H\d{5,}/i.test(line)).map(line => {
      const expediente = (line.match(/H\d{5,}/i) || [''])[0];
      const contenedor = (line.match(/[A-Z]{4}\d{7}/i) || [''])[0];
      const situacion = (line.match(/Pdte\s+Carga|Pdte\s+embarque|Embarcado|Cargado|Anulado/i) || [''])[0];
      const eta = (line.match(/\b\d{1,2}[-/](?:ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic|\d{1,2})\b/i) || [''])[0];
      return normalizeBulkShipment({ expediente, eta, situacion, contenedor, partner: 'RAMINATRANS', vessel: line.replace(expediente, '').replace(contenedor, '').trim() });
    }).filter(Boolean);
  }

  function upsertBulkImportShipments(shipments) {
    ensureImportShipments();
    let created = 0;
    let updated = 0;
    shipments.forEach(next => {
      const key = shipmentKey(next);
      const idx = APP.importShipments.findIndex(item => item.id === next.id || (key && shipmentKey(item) === key));
      if (idx >= 0) {
        APP.importShipments[idx] = {
          ...APP.importShipments[idx],
          ...next,
          docs: { ...(APP.importShipments[idx].docs || {}), ...(next.docs || {}) },
          createdAt: APP.importShipments[idx].createdAt || nowIso(),
          updatedAt: nowIso()
        };
        updated += 1;
      } else {
        APP.importShipments.push({ ...next, createdAt: nowIso(), updatedAt: nowIso() });
        created += 1;
      }
    });
    return { created, updated };
  }

  function loadTesseractScript() {
    if (window.Tesseract) return Promise.resolve();
    if (window._tmsTesseractLoading) return window._tmsTesseractLoading;
    window._tmsTesseractLoading = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js';
      script.onload = () => resolve();
      script.onerror = () => reject(new Error('No se pudo cargar el lector OCR. Revisa la conexión e intenta de nuevo.'));
      document.head.appendChild(script);
    });
    return window._tmsTesseractLoading;
  }

  function readTextFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = event => resolve(event.target.result || '');
      reader.onerror = () => reject(reader.error || new Error('No se pudo leer el archivo.'));
      reader.readAsText(file);
    });
  }

  function readArrayBufferFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = event => resolve(event.target.result);
      reader.onerror = () => reject(reader.error || new Error('No se pudo leer el archivo.'));
      reader.readAsArrayBuffer(file);
    });
  }

  function updateBulkImportStatus(message) {
    const status = document.getElementById('impBulkStatus');
    if (status) status.textContent = message || '';
  }

  async function parseImportShipmentsFile(file) {
    const name = (file.name || '').toLowerCase();
    if (/\.(xlsx|xls|xml)$/.test(name)) {
      const buffer = await readArrayBufferFile(file);
      const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      return parseBulkRowsFromObjects(XLSX.utils.sheet_to_json(sheet, { defval: '' }));
    }
    if (/\.(csv|txt)$/.test(name) || /^text\//.test(file.type || '')) {
      return parseBulkShipmentsFromText(await readTextFile(file));
    }
    if (/^image\//.test(file.type || '')) {
      updateBulkImportStatus('Leyendo imagen con OCR... puede tardar unos segundos.');
      await loadTesseractScript();
      const result = await window.Tesseract.recognize(file, 'spa+eng', {
        logger: event => {
          if (event && event.status) {
            const pct = event.progress ? ` ${Math.round(event.progress * 100)}%` : '';
            updateBulkImportStatus('OCR: ' + event.status + pct);
          }
        }
      });
      return parseBulkShipmentsFromText(result && result.data ? result.data.text : '');
    }
    return parseBulkShipmentsFromText(await readTextFile(file));
  }

  window.importarEmbarquesDesdeArchivo = async function importarEmbarquesDesdeArchivo(file) {
    if (!file) return;
    if (!hasPermission('importaciones', 'editar')) return alert('Tu rol no puede editar importaciones.');
    try {
      updateBulkImportStatus('Procesando archivo...');
      const shipments = await parseImportShipmentsFile(file);
      if (!shipments.length) throw new Error('No pude detectar embarques en el archivo. Prueba con Excel/CSV o una captura más nítida.');
      pushUndoState('importar embarques');
      const result = upsertBulkImportShipments(shipments);
      updateBulkImportStatus(`Importación lista: ${result.created} nuevos · ${result.updated} actualizados.`);
      renderImportaciones();
      scheduleAutoSave();
    } catch (error) {
      console.error('Error importando embarques:', error);
      updateBulkImportStatus('Error: ' + error.message);
      alert('No se pudieron importar los embarques: ' + error.message);
    } finally {
      const input = document.getElementById('impBulkFile');
      if (input) input.value = '';
    }
  };

  window.importarEmbarquesDesdeTexto = function importarEmbarquesDesdeTexto() {
    if (!hasPermission('importaciones', 'editar')) return alert('Tu rol no puede editar importaciones.');
    const textarea = document.getElementById('impBulkText');
    const raw = textarea ? textarea.value : '';
    const shipments = parseBulkShipmentsFromText(raw);
    if (!shipments.length) return alert('No pude detectar embarques en el texto pegado.');
    pushUndoState('importar embarques pegados');
    const result = upsertBulkImportShipments(shipments);
    if (textarea) textarea.value = '';
    updateBulkImportStatus(`Importación lista: ${result.created} nuevos · ${result.updated} actualizados.`);
    renderImportaciones();
    scheduleAutoSave();
  };

  function ensureImportShipments() {
    APP.importShipments = Array.isArray(APP.importShipments) ? APP.importShipments : [];
    APP.importShipments.forEach(item => {
      item.docs = item.docs || {};
      item.origen = item.origen || 'espana';
      item.etapa = item.etapa || 'pedido';
    });
  }

  function importStageLabel(stage) {
    const found = IMPORT_STAGES.find(item => item.key === stage);
    return found ? found.label : 'Sin etapa';
  }

  function importOriginLabel(origin) {
    return origin === 'miami' ? 'EE. UU. / Miami' : 'España';
  }

  function importEtaState(item) {
    if (!item.eta) return { cls: '', label: 'Sin ETA', days: null };
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const eta = commercialFilterDateToDate(item.eta);
    if (!eta || Number.isNaN(eta.getTime())) return { cls: '', label: 'ETA inválida', days: null };
    eta.setHours(0, 0, 0, 0);
    const days = Math.round((eta - today) / 86400000);
    if (days < 0 && !['recepcion_sap'].includes(item.etapa)) return { cls: 'overdue', label: `ETA vencida hace ${Math.abs(days)} días`, days };
    if (days <= 5 && !['recepcion_sap'].includes(item.etapa)) return { cls: 'alert', label: days === 0 ? 'ETA hoy' : `ETA en ${days} días`, days };
    return { cls: '', label: `ETA ${item.eta}`, days };
  }

  function importRequiredDocs(origin) {
    return IMPORT_DOCS[origin] || IMPORT_DOCS.espana;
  }

  function importDocsProgress(item) {
    const docs = importRequiredDocs(item.origen);
    const done = docs.filter(([key]) => !!(item.docs && item.docs[key])).length;
    return { done, total: docs.length, pct: docs.length ? Math.round((done / docs.length) * 100) : 0 };
  }

  function getImportFormValue(id) {
    return text((document.getElementById(id) || {}).value);
  }

  function importDocsFormHtml(origin, current) {
    return importRequiredDocs(origin).map(([key, label]) => `<label class="import-doc-chip ${current && current[key] ? 'ok' : ''}" style="cursor:pointer;"><input type="checkbox" class="importDocInput" value="${key}" ${current && current[key] ? 'checked' : ''}> ${label}</label>`).join('');
  }

  window.cambiarDocsImportacionOrigen = function cambiarDocsImportacionOrigen() {
    const origin = getImportFormValue('impOrigen') || 'espana';
    const wrap = document.getElementById('impDocsWrap');
    if (wrap) wrap.innerHTML = importDocsFormHtml(origin, {});
  };

  function clearImportForm() {
    APP.importEditingId = '';
    ['impReferencia', 'impProveedor', 'impPO', 'impFactura', 'impBL', 'impTracking', 'impETA', 'impMiami', 'impAduana', 'impSalidaAduana', 'impAlmacen', 'impValor', 'impNotas'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    const origin = document.getElementById('impOrigen');
    if (origin) origin.value = 'espana';
    const stage = document.getElementById('impEtapa');
    if (stage) stage.value = 'pedido';
    cambiarDocsImportacionOrigen();
    const btn = document.getElementById('impSaveBtn');
    if (btn) btn.textContent = 'Crear embarque';
  }

  window.clearImportForm = clearImportForm;

  window.editarImportacion = function editarImportacion(id) {
    if (!hasPermission('importaciones', 'editar')) return alert('Tu rol no puede editar importaciones.');
    ensureImportShipments();
    const item = APP.importShipments.find(row => row.id === id);
    if (!item) return;
    APP.importEditingId = id;
    const set = (fieldId, value) => { const el = document.getElementById(fieldId); if (el) el.value = value || ''; };
    set('impOrigen', item.origen || 'espana');
    set('impEtapa', item.etapa || 'pedido');
    set('impReferencia', item.referencia);
    set('impProveedor', item.proveedor);
    set('impPO', item.po);
    set('impFactura', item.factura);
    set('impBL', item.bl);
    set('impTracking', item.tracking);
    set('impETA', item.eta);
    set('impMiami', item.fechaMiami);
    set('impAduana', item.fechaAduana);
    set('impSalidaAduana', item.fechaSalidaAduana);
    set('impAlmacen', item.fechaAlmacen);
    set('impValor', item.valor);
    set('impNotas', item.notas);
    const wrap = document.getElementById('impDocsWrap');
    if (wrap) wrap.innerHTML = importDocsFormHtml(item.origen || 'espana', item.docs || {});
    const btn = document.getElementById('impSaveBtn');
    if (btn) btn.textContent = 'Guardar cambios';
    document.getElementById('importacionesMount')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  window.eliminarImportacion = function eliminarImportacion(id) {
    if (!hasPermission('importaciones', 'editar')) return alert('Tu rol no puede editar importaciones.');
    if (!confirm('¿Eliminar este embarque del seguimiento?')) return;
    pushUndoState('eliminar importación');
    APP.importShipments = (APP.importShipments || []).filter(item => item.id !== id);
    renderImportaciones();
    scheduleAutoSave();
  };

  window.avanzarImportacion = function avanzarImportacion(id) {
    if (!hasPermission('importaciones', 'editar')) return alert('Tu rol no puede editar importaciones.');
    const item = (APP.importShipments || []).find(row => row.id === id);
    if (!item) return;
    const idx = IMPORT_STAGES.findIndex(stage => stage.key === item.etapa);
    if (idx < IMPORT_STAGES.length - 1) {
      pushUndoState('avanzar importación');
      item.etapa = IMPORT_STAGES[idx + 1].key;
      item.updatedAt = nowIso();
      renderImportaciones();
      scheduleAutoSave();
    }
  };

  window.guardarImportacion = function guardarImportacion() {
    if (!hasPermission('importaciones', 'editar')) return alert('Tu rol no puede editar importaciones.');
    ensureImportShipments();
    const docs = {};
    document.querySelectorAll('.importDocInput').forEach(input => { docs[input.value] = !!input.checked; });
    const id = APP.importEditingId || ('imp-' + Date.now());
    const item = {
      id,
      origen: getImportFormValue('impOrigen') || 'espana',
      etapa: getImportFormValue('impEtapa') || 'pedido',
      referencia: getImportFormValue('impReferencia') || id,
      proveedor: getImportFormValue('impProveedor'),
      po: getImportFormValue('impPO'),
      factura: getImportFormValue('impFactura'),
      bl: getImportFormValue('impBL'),
      tracking: getImportFormValue('impTracking'),
      eta: getImportFormValue('impETA'),
      fechaMiami: getImportFormValue('impMiami'),
      fechaAduana: getImportFormValue('impAduana'),
      fechaSalidaAduana: getImportFormValue('impSalidaAduana'),
      fechaAlmacen: getImportFormValue('impAlmacen'),
      valor: num(getImportFormValue('impValor')),
      notas: getImportFormValue('impNotas'),
      docs,
      updatedAt: nowIso()
    };
    pushUndoState(APP.importEditingId ? 'editar importación' : 'crear importación');
    const idx = APP.importShipments.findIndex(row => row.id === id);
    if (idx >= 0) APP.importShipments[idx] = item;
    else APP.importShipments.push(item);
    clearImportForm();
    renderImportaciones();
    scheduleAutoSave();
  };

  function importCardHtml(item) {
    const eta = importEtaState(item);
    const progress = importDocsProgress(item);
    const docsHtml = importRequiredDocs(item.origen).slice(0, 6).map(([key, label]) => `<span class="import-doc-chip ${item.docs && item.docs[key] ? 'ok' : ''}">${item.docs && item.docs[key] ? '✓' : '·'} ${label}</span>`).join('');
    return `<article class="import-card ${eta.cls}">
      <div class="import-card-title"><span>${item.referencia || 'Sin ref.'}</span><span>${progress.done}/${progress.total}</span></div>
      <div class="import-card-meta"><strong>${importOriginLabel(item.origen)}</strong> · ${item.proveedor || item.partner || 'Proveedor pendiente'}<br>Booking/PO ${item.booking || item.po || '—'} · Vessel ${item.vessel || '—'}<br>Contenedor/BL ${item.contenedor || item.bl || '—'} · ${eta.label}</div>
      <div class="import-docs">${docsHtml}</div>
      <div class="import-card-meta">ETD: ${item.etd || '—'} · Carga: ${item.fechaCarga || '—'} ${item.horaCarga || ''} · Situación: ${item.situacion || '—'}<br>Destino: ${item.destino || '—'} · POA: ${item.poa || '—'} · POD: ${item.pod || '—'}</div>
      <div class="import-card-meta">Miami: ${item.fechaMiami || '—'} · Aduana: ${item.fechaAduana || '—'} · Salida: ${item.fechaSalidaAduana || '—'} · Almacén: ${item.fechaAlmacen || '—'}</div>
      ${item.notas ? `<div class="import-card-meta">${item.notas}</div>` : ''}
      <div class="import-actions">
        <button class="btn btn-outline btn-sm" onclick="editarImportacion('${jsString(item.id)}')">Editar</button>
        <button class="btn btn-primary btn-sm" onclick="avanzarImportacion('${jsString(item.id)}')">Avanzar</button>
        <button class="btn btn-outline btn-sm" onclick="eliminarImportacion('${jsString(item.id)}')">Eliminar</button>
      </div>
    </article>`;
  }

  function buildImportDailyReport() {
    ensureImportShipments();
    const active = APP.importShipments.filter(item => item.etapa !== 'recepcion_sap');
    const overdue = active.filter(item => importEtaState(item).cls === 'overdue');
    const alerts = active.filter(item => importEtaState(item).cls === 'alert');
    const missingSap = APP.importShipments.filter(item => item.etapa === 'recepcion_sap' && !(item.docs && item.docs.facturaSap && item.docs.gastosSap));
    const byStage = IMPORT_STAGES.map(stage => `${stage.label}: ${APP.importShipments.filter(item => item.etapa === stage.key).length}`).join('\n');
    return `Reporte diario de importaciones\n\nEmbarques activos: ${active.length}\nETA vencida: ${overdue.length}\nETA próxima: ${alerts.length}\nPendientes de registro SAP: ${missingSap.length}\n\nPor etapa:\n${byStage}\n\nSeguimiento prioritario:\n${[...overdue, ...alerts, ...missingSap].slice(0, 10).map(item => `- ${item.referencia || item.id}: ${importStageLabel(item.etapa)} · ${importEtaState(item).label} · ${item.notas || 'Sin nota'}`).join('\n') || '- Sin alertas críticas.'}`;
  }

  window.exportarImportacionesExcel = function exportarImportacionesExcel() {
    ensureImportShipments();
    const rows = APP.importShipments.map(item => ({
      Referencia: item.referencia,
      Origen: importOriginLabel(item.origen),
      Etapa: importStageLabel(item.etapa),
      Proveedor: item.proveedor,
      Booking: item.booking || item.po,
      PO: item.po,
      Factura: item.factura,
      BL_Waybill: item.bl,
      Contenedor: item.contenedor || '',
      Vessel: item.vessel || '',
      ETD: item.etd || '',
      Equipo: item.equipo || '',
      Fecha_Carga: item.fechaCarga || '',
      Hora_Carga: item.horaCarga || '',
      Situacion: item.situacion || '',
      Destino: item.destino || '',
      POA: item.poa || '',
      POD: item.pod || '',
      Tracking: item.tracking,
      ETA: item.eta,
      Estado_ETA: importEtaState(item).label,
      Llegada_Miami: item.fechaMiami,
      Fecha_Aduana: item.fechaAduana,
      Salida_Aduana: item.fechaSalidaAduana,
      Llegada_Almacen: item.fechaAlmacen,
      Docs: `${importDocsProgress(item).done}/${importDocsProgress(item).total}`,
      Valor: item.valor,
      Notas: item.notas
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Importaciones');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ Reporte: buildImportDailyReport() }]), 'Reporte diario');
    XLSX.writeFile(wb, 'TMS_Importaciones.xlsx');
  };


  const PRIORITY_STATUSES = [
    { key: 'pendiente', label: 'Pendiente' },
    { key: 'proceso', label: 'En proceso' },
    { key: 'completada', label: 'Completada' }
  ];

  function todayIsoDate() {
    return new Date().toISOString().slice(0, 10);
  }

  function priorityStatusLabel(status) {
    return (PRIORITY_STATUSES.find(item => item.key === status) || PRIORITY_STATUSES[0]).label;
  }

  function selectedPriorityDate() {
    const date = text((document.getElementById('prioDate') || {}).value) || APP.prioritiesDate || todayIsoDate();
    APP.prioritiesDate = date;
    return date;
  }

  function isoDateFromAny(value, fallback) {
    const raw = text(value);
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
    if (raw) {
      const d = new Date(raw);
      if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    }
    return fallback || todayIsoDate();
  }

  function ensureDailyPriorities() {
    APP.dailyPriorities = Array.isArray(APP.dailyPriorities) ? APP.dailyPriorities : [];
    APP.dailyPriorities.forEach(item => {
      item.id = item.id || ('prio-' + Date.now() + '-' + Math.random().toString(16).slice(2));
      item.fecha = isoDateFromAny(item.fecha || item.createdAt, todayIsoDate());
      item.estado = PRIORITY_STATUSES.some(status => status.key === item.estado) ? item.estado : 'pendiente';
      item.titulo = text(item.titulo || item.texto || 'Tarea sin título');
      item.notas = text(item.notas || '');
      item.createdAt = item.createdAt || nowIso();
      item.updatedAt = item.updatedAt || item.createdAt;
      if (item.estado === 'completada') {
        item.completedDate = isoDateFromAny(item.completedDate || item.completedAt || item.updatedAt, item.fecha);
        item.completedAt = item.completedAt || item.updatedAt || nowIso();
      }
    });
  }

  function ensurePrioritySpainTopics() {
    APP.prioritySpainTopics = Array.isArray(APP.prioritySpainTopics) ? APP.prioritySpainTopics : [];
    APP.prioritySpainTopics.forEach(item => {
      item.id = item.id || ('spain-' + Date.now() + '-' + Math.random().toString(16).slice(2));
      item.titulo = text(item.titulo || item.texto || 'Tema pendiente');
      item.notas = text(item.notas || '');
      item.done = !!item.done;
      item.createdAt = item.createdAt || nowIso();
      item.updatedAt = item.updatedAt || item.createdAt;
      if (item.done) item.completedAt = item.completedAt || item.updatedAt;
    });
  }

  function taskIsVisibleOnPriorityDate(item, date) {
    if (item.fecha > date) return false;
    if (item.estado === 'completada') return item.completedDate === date;
    return true;
  }

  function getPrioritiesForSelectedDate() {
    ensureDailyPriorities();
    const date = selectedPriorityDate();
    return APP.dailyPriorities
      .filter(item => taskIsVisibleOnPriorityDate(item, date))
      .sort((a, b) => {
        const stateOrder = { pendiente: 1, proceso: 2, completada: 3 };
        return (stateOrder[a.estado] || 9) - (stateOrder[b.estado] || 9) || (a.orden || 0) - (b.orden || 0) || text(a.titulo).localeCompare(text(b.titulo));
      });
  }

  function getPriorityMetricsForDate(date) {
    ensureDailyPriorities();
    const started = APP.dailyPriorities.filter(item => item.fecha === date).length;
    const completed = APP.dailyPriorities.filter(item => item.estado === 'completada' && item.completedDate === date).length;
    const pending = APP.dailyPriorities.filter(item => item.estado === 'pendiente' && item.fecha <= date).length;
    const progress = APP.dailyPriorities.filter(item => item.estado === 'proceso' && item.fecha <= date).length;
    const carried = APP.dailyPriorities.filter(item => item.estado !== 'completada' && item.fecha < date).length;
    return { date, started, completed, pending, progress, carried };
  }

  function priorityDateRange(centerDate) {
    const base = new Date(centerDate + 'T00:00:00');
    return Array.from({ length: 7 }).map((_, index) => {
      const d = new Date(base);
      d.setDate(base.getDate() - 6 + index);
      return d.toISOString().slice(0, 10);
    });
  }

  function renderPriorityChart(selectedDate) {
    const metrics = priorityDateRange(selectedDate).map(getPriorityMetricsForDate);
    const max = Math.max(1, ...metrics.flatMap(row => [row.started, row.carried, row.completed]));
    return `<section class="priority-chart">
      <div class="priority-chart-head"><div><strong>Movimiento semanal</strong><span>Iniciadas, arrastradas y completadas por día</span></div></div>
      ${metrics.map(row => `<div class="priority-chart-row">
        <strong>${row.date.slice(5)}</strong>
        <div class="priority-bars">
          <div class="priority-bar-track"><div class="priority-bar-fill started" style="width:${Math.max(3, row.started / max * 100)}%;"></div></div>
          <div class="priority-bar-track"><div class="priority-bar-fill carried" style="width:${Math.max(3, row.carried / max * 100)}%;"></div></div>
          <div class="priority-bar-track"><div class="priority-bar-fill done" style="width:${Math.max(3, row.completed / max * 100)}%;"></div></div>
          <div class="priority-bar-labels"><span>Ini ${row.started}</span><span>Arr ${row.carried}</span><span>Comp ${row.completed}</span><span>Proc ${row.progress}</span></div>
        </div>
      </div>`).join('')}
    </section>`;
  }

  function priorityCardHtml(item) {
    const statusClass = 'status-' + (item.estado || 'pendiente');
    const carried = item.estado !== 'completada' && item.fecha < (APP.prioritiesDate || todayIsoDate());
    return `<div class="priority-card ${statusClass}" draggable="true" ondragstart="dragPriorityTask(event,'${item.id}')">
      <div class="priority-card-title">${escapeHtml(item.titulo)}</div>
      ${item.notas ? `<div class="priority-card-note">${escapeHtml(item.notas)}</div>` : ''}
      <div class="priority-card-meta"><span><i class="priority-status-dot ${statusClass}"></i>${priorityStatusLabel(item.estado)}${carried ? ' · Arrastrada' : ''}</span><span>${item.estado === 'completada' ? 'Cerrada ' + (item.completedDate || '') : 'Inicio ' + item.fecha}</span></div>
      <div class="priority-actions">
        ${PRIORITY_STATUSES.map(status => `<button class="btn btn-outline btn-sm" onclick="moverPrioridad('${item.id}','${status.key}')" ${item.estado === status.key ? 'disabled' : ''}>${status.label}</button>`).join('')}
        <button class="btn btn-outline btn-sm" onclick="editarPrioridad('${item.id}')">Editar</button>
        <button class="btn btn-outline btn-sm" onclick="eliminarPrioridad('${item.id}')">Eliminar</button>
      </div>
    </div>`;
  }

  function renderSpainTopicsHtml() {
    ensurePrioritySpainTopics();
    const open = APP.prioritySpainTopics.filter(item => !item.done).length;
    const done = APP.prioritySpainTopics.filter(item => item.done).length;
    return `<section class="priority-spain-panel">
      <div class="priority-spain-head"><div><strong>Temas pendientes de verificar con España</strong><span>${open} abiertos · ${done} cerrados</span></div></div>
      <div class="priority-spain-form">
        <input id="spainTopicInput" class="form-control" placeholder="Tema a hablar o cerrar">
        <button class="btn btn-primary btn-sm" onclick="agregarTemaEspana()">Agregar</button>
      </div>
      <div class="priority-spain-list">
        ${APP.prioritySpainTopics.length ? APP.prioritySpainTopics.map(item => `<div class="priority-spain-item ${item.done ? 'done' : ''}">
          <input type="checkbox" ${item.done ? 'checked' : ''} onchange="toggleTemaEspana('${item.id}')">
          <span><strong>${escapeHtml(item.titulo)}</strong>${item.notas ? `<small>${escapeHtml(item.notas)}</small>` : ''}<small>${item.done ? 'Cerrado' : 'Pendiente'} · ${new Date(item.updatedAt || item.createdAt || nowIso()).toLocaleDateString('es-DO')}</small></span>
          <button class="btn btn-outline btn-sm" onclick="editarTemaEspana('${item.id}')">Editar</button>
        </div>`).join('') : '<div class="route-catalog-empty">Sin temas pendientes.</div>'}
      </div>
    </section>`;
  }

  function buildPriorityReport(rows) {
    ensurePrioritySpainTopics();
    const pending = rows.filter(item => item.estado === 'pendiente').length;
    const progress = rows.filter(item => item.estado === 'proceso').length;
    const done = rows.filter(item => item.estado === 'completada').length;
    const carried = rows.filter(item => item.estado !== 'completada' && item.fecha < (APP.prioritiesDate || todayIsoDate())).length;
    const lines = PRIORITY_STATUSES.map(status => {
      const items = rows.filter(item => item.estado === status.key);
      return `${status.label}: ${items.length}\n${items.map(item => '- ' + item.titulo + (item.notas ? ' · ' + item.notas : '')).join('\n') || '- Sin tareas'}`;
    }).join('\n\n');
    const spain = APP.prioritySpainTopics.map(item => `- [${item.done ? 'x' : ' '}] ${item.titulo}${item.notas ? ' · ' + item.notas : ''}`).join('\n') || '- Sin temas';
    return `Reporte diario de prioridades\nFecha: ${APP.prioritiesDate || todayIsoDate()}\n\nResumen:\nPendientes: ${pending}\nEn proceso: ${progress}\nCompletadas del día: ${done}\nArrastradas de días previos: ${carried}\nTotal visible: ${rows.length}\n\n${lines}\n\nTemas pendientes de verificar con España:\n${spain}`;
  }

  function parsePriorityDictation(transcript) {
    let value = text(transcript)
      .replace(/^agregar\s+(tarea|prioridad)\s+/i, '')
      .replace(/^crear\s+(tarea|prioridad)\s+/i, '')
      .replace(/^tarea\s+/i, '');
    const parts = value.split(/\s+(?:nota|notas|observacion|observación)\s+/i);
    return { titulo: text(parts[0]), notas: text(parts.slice(1).join(' ')) };
  }

  window.renderPrioridades = function renderPrioridades() {
    const mount = document.getElementById('prioridadesMount');
    if (!mount) return;
    if (!isAdminUser()) {
      mount.innerHTML = '<div class="empty-state"><div class="empty-icon">!</div><p>Este módulo es solo para administradores.</p></div>';
      return;
    }
    ensureDailyPriorities();
    ensurePrioritySpainTopics();
    APP.prioritiesDate = APP.prioritiesDate || todayIsoDate();
    const rows = getPrioritiesForSelectedDate();
    const selectedDate = APP.prioritiesDate;
    const counts = Object.fromEntries(PRIORITY_STATUSES.map(status => [status.key, rows.filter(item => item.estado === status.key).length]));
    const carried = rows.filter(item => item.estado !== 'completada' && item.fecha < selectedDate).length;
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    mount.innerHTML = `<div class="priorities-module">
      <section class="priority-hero">
        <div><h2>Prioridades del día</h2><p>Las tareas pendientes o en proceso se arrastran automáticamente hasta que se marquen como completadas. El histórico muestra las completadas por día.</p></div>
        <div class="priority-top-actions">
          <input id="prioDate" class="form-control" type="date" value="${APP.prioritiesDate}" onchange="APP.prioritiesDate=this.value;renderPrioridades()">
          <button class="btn btn-success btn-sm" onclick="exportarPrioridadesExcel()">Exportar Excel</button>
          <button class="btn btn-outline btn-sm" onclick="exportarPrioridadesPdf()">Exportar PDF</button>
        </div>
      </section>
      <section class="priority-insights-grid">
        ${renderPriorityChart(selectedDate)}
        ${renderSpainTopicsHtml()}
      </section>
      <section class="card">
        <div class="priority-form">
          <input id="prioTitle" class="form-control" placeholder="Nueva prioridad del día">
          <input id="prioNotes" class="form-control" placeholder="Nota opcional">
          <button class="btn btn-primary btn-sm" onclick="agregarPrioridad()">Agregar</button>
          <button id="prioVoiceBtn" class="btn btn-outline btn-sm" onclick="iniciarDictadoPrioridades()" ${SpeechRecognition ? '' : 'disabled'}>${SpeechRecognition ? 'Dictar' : 'Dictado no disponible'}</button>
        </div>
        <div id="prioVoiceStatus" class="priority-voice-status">${SpeechRecognition ? 'Puedes dictar una tarea y se agregará automáticamente.' : 'Este navegador no expone dictado por voz.'}</div>
      </section>
      <section class="priority-kpis">
        <div class="priority-kpi"><strong>${rows.length}</strong><span>Total visible</span></div>
        <div class="priority-kpi"><strong>${counts.pendiente || 0}</strong><span>Pendientes</span></div>
        <div class="priority-kpi"><strong>${counts.proceso || 0}</strong><span>En proceso</span></div>
        <div class="priority-kpi"><strong>${counts.completada || 0}</strong><span>Completadas del día</span></div>
        <div class="priority-kpi"><strong>${carried}</strong><span>Arrastradas</span></div>
      </section>
      <section class="priority-board">
        ${PRIORITY_STATUSES.map(status => {
          const items = rows.filter(item => item.estado === status.key);
          return `<div class="priority-column status-${status.key}" ondragover="event.preventDefault();this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="dropPriorityTask(event,'${status.key}',this)">
            <div class="priority-column-head"><span>${status.label}</span><span>${items.length}</span></div>
            <div class="priority-column-body">${items.length ? items.map(priorityCardHtml).join('') : '<div class="route-catalog-empty">Sin tareas.</div>'}</div>
          </div>`;
        }).join('')}
      </section>
      <section class="priority-report"><pre>${escapeHtml(buildPriorityReport(rows))}</pre></section>
    </div>`;
  };

  window.agregarPrioridad = function agregarPrioridad() {
    if (!requireAdminUser('Crear prioridades')) return;
    const titulo = text((document.getElementById('prioTitle') || {}).value);
    const notas = text((document.getElementById('prioNotes') || {}).value);
    if (!titulo) return alert('Escribe la prioridad que quieres agregar.');
    ensureDailyPriorities();
    pushUndoState('agregar prioridad');
    const fecha = selectedPriorityDate();
    APP.dailyPriorities.push({ id: 'prio-' + Date.now(), fecha, titulo, notas, estado: 'pendiente', orden: APP.dailyPriorities.length + 1, createdAt: nowIso(), updatedAt: nowIso() });
    APP.prioritiesDate = fecha;
    renderPrioridades();
    scheduleAutoSave();
  };

  window.moverPrioridad = function moverPrioridad(id, estado) {
    if (!requireAdminUser('Mover prioridades')) return;
    const item = (APP.dailyPriorities || []).find(row => row.id === id);
    if (!item || item.estado === estado) return;
    pushUndoState('mover prioridad');
    item.estado = estado;
    if (estado === 'completada') {
      item.completedDate = selectedPriorityDate();
      item.completedAt = nowIso();
    } else {
      item.completedDate = '';
      item.completedAt = '';
    }
    item.updatedAt = nowIso();
    renderPrioridades();
    scheduleAutoSave();
  };

  window.dragPriorityTask = function dragPriorityTask(event, id) {
    event.dataTransfer.setData('text/plain', id);
  };

  window.dropPriorityTask = function dropPriorityTask(event, estado, el) {
    event.preventDefault();
    if (el) el.classList.remove('drag-over');
    const id = event.dataTransfer.getData('text/plain');
    moverPrioridad(id, estado);
  };

  window.editarPrioridad = function editarPrioridad(id) {
    if (!requireAdminUser('Editar prioridades')) return;
    const item = (APP.dailyPriorities || []).find(row => row.id === id);
    if (!item) return;
    const titulo = prompt('Editar prioridad', item.titulo);
    if (titulo === null) return;
    const notas = prompt('Nota opcional', item.notas || '');
    pushUndoState('editar prioridad');
    item.titulo = text(titulo) || item.titulo;
    item.notas = text(notas || '');
    item.updatedAt = nowIso();
    renderPrioridades();
    scheduleAutoSave();
  };

  window.eliminarPrioridad = function eliminarPrioridad(id) {
    if (!requireAdminUser('Eliminar prioridades')) return;
    const item = (APP.dailyPriorities || []).find(row => row.id === id);
    if (!item || !confirm('¿Eliminar esta prioridad?')) return;
    pushUndoState('eliminar prioridad');
    APP.dailyPriorities = APP.dailyPriorities.filter(row => row.id !== id);
    renderPrioridades();
    scheduleAutoSave();
  };

  window.agregarTemaEspana = function agregarTemaEspana() {
    if (!requireAdminUser('Agregar tema España')) return;
    const input = document.getElementById('spainTopicInput');
    const titulo = text(input && input.value);
    if (!titulo) return alert('Escribe el tema pendiente.');
    ensurePrioritySpainTopics();
    pushUndoState('agregar tema España');
    APP.prioritySpainTopics.push({ id: 'spain-' + Date.now(), titulo, notas: '', done: false, createdAt: nowIso(), updatedAt: nowIso() });
    renderPrioridades();
    scheduleAutoSave();
  };

  window.toggleTemaEspana = function toggleTemaEspana(id) {
    if (!requireAdminUser('Cerrar tema España')) return;
    const item = (APP.prioritySpainTopics || []).find(row => row.id === id);
    if (!item) return;
    pushUndoState('actualizar tema España');
    item.done = !item.done;
    item.completedAt = item.done ? nowIso() : '';
    item.updatedAt = nowIso();
    renderPrioridades();
    scheduleAutoSave();
  };

  window.editarTemaEspana = function editarTemaEspana(id) {
    if (!requireAdminUser('Editar tema España')) return;
    const item = (APP.prioritySpainTopics || []).find(row => row.id === id);
    if (!item) return;
    const titulo = prompt('Editar tema pendiente. Deja vacío para eliminarlo.', item.titulo);
    if (titulo === null) return;
    pushUndoState('editar tema España');
    if (!text(titulo)) APP.prioritySpainTopics = APP.prioritySpainTopics.filter(row => row.id !== id);
    else {
      const notas = prompt('Nota opcional', item.notas || '');
      item.titulo = text(titulo);
      item.notas = text(notas || '');
      item.updatedAt = nowIso();
    }
    renderPrioridades();
    scheduleAutoSave();
  };

  window.iniciarDictadoPrioridades = function iniciarDictadoPrioridades() {
    if (!requireAdminUser('Dictar prioridades')) return;
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    const status = document.getElementById('prioVoiceStatus');
    if (!SpeechRecognition) return alert('El dictado no está disponible en este navegador.');
    const recognition = new SpeechRecognition();
    recognition.lang = 'es-DO';
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;
    if (status) status.textContent = 'Escuchando...';
    recognition.onresult = event => {
      const transcript = event.results && event.results[0] && event.results[0][0] ? event.results[0][0].transcript : '';
      const parsed = parsePriorityDictation(transcript);
      if (!parsed.titulo) {
        if (status) status.textContent = 'No pude detectar una tarea clara. Intenta otra vez.';
        return;
      }
      const titleEl = document.getElementById('prioTitle');
      const notesEl = document.getElementById('prioNotes');
      if (titleEl) titleEl.value = parsed.titulo;
      if (notesEl) notesEl.value = parsed.notas;
      if (status) status.textContent = 'Detectado: ' + parsed.titulo;
      agregarPrioridad();
    };
    recognition.onerror = event => {
      if (status) status.textContent = 'No se pudo usar el dictado: ' + (event.error || 'error desconocido');
    };
    recognition.onend = () => {
      const current = document.getElementById('prioVoiceStatus');
      if (current && current.textContent === 'Escuchando...') current.textContent = 'Dictado detenido.';
    };
    recognition.start();
  };

  window.exportarPrioridadesExcel = function exportarPrioridadesExcel() {
    const rows = getPrioritiesForSelectedDate();
    ensurePrioritySpainTopics();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.map(item => ({
      Fecha_inicio: item.fecha,
      Fecha_completada: item.completedDate || '',
      Estado: priorityStatusLabel(item.estado),
      Prioridad: item.titulo,
      Notas: item.notas || '',
      Creada: item.createdAt || '',
      Actualizada: item.updatedAt || ''
    }))), 'Prioridades');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(priorityDateRange(APP.prioritiesDate || todayIsoDate()).map(getPriorityMetricsForDate)), 'Resumen');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(APP.prioritySpainTopics.map(item => ({ Tema: item.titulo, Notas: item.notas || '', Estado: item.done ? 'Cerrado' : 'Pendiente', Creado: item.createdAt || '', Cerrado: item.completedAt || '' }))), 'Temas Espana');
    XLSX.writeFile(wb, 'TMS_Prioridades_' + (APP.prioritiesDate || todayIsoDate()) + '.xlsx');
  };

  window.exportarPrioridadesPdf = function exportarPrioridadesPdf() {
    const rows = getPrioritiesForSelectedDate();
    const report = buildPriorityReport(rows);
    const stage = document.createElement('div');
    stage.style.cssText = 'position:fixed;left:-10000px;top:0;width:760px;background:#fff;color:#111827;padding:24px;font-family:Arial,sans-serif;';
    stage.innerHTML = `<h1 style="font-size:20px;margin:0 0 12px;">Prioridades ${escapeHtml(APP.prioritiesDate || todayIsoDate())}</h1><pre style="white-space:pre-wrap;font-family:Arial,sans-serif;font-size:12px;line-height:1.45;">${escapeHtml(report)}</pre>`;
    document.body.appendChild(stage);
    const done = () => setTimeout(() => stage.remove(), 500);
    if (window.html2pdf) {
      window.html2pdf().set({ margin: 10, filename: 'TMS_Prioridades_' + (APP.prioritiesDate || todayIsoDate()) + '.pdf', html2canvas: { scale: 2 }, jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' } }).from(stage).save().then(done).catch(done);
    } else {
      alert('No se pudo cargar el exportador PDF. Usa Exportar Excel por ahora.');
      done();
    }
  };

  window.renderImportaciones = function renderImportaciones() {
    const mount = document.getElementById('importacionesMount');
    if (!mount) return;
    ensureImportShipments();
    const canEdit = hasPermission('importaciones', 'editar');
    const q = text((document.getElementById('impSearch') || {}).value).toLowerCase();
    const originFilter = text((document.getElementById('impFilterOrigen') || {}).value);
    const rows = APP.importShipments.filter(item => {
      if (originFilter && item.origen !== originFilter) return false;
      if (!q) return true;
      return [item.referencia, item.proveedor, item.po, item.factura, item.bl, item.tracking, item.notas].some(value => text(value).toLowerCase().includes(q));
    });
    const active = APP.importShipments.filter(item => item.etapa !== 'recepcion_sap');
    const overdue = APP.importShipments.filter(item => importEtaState(item).cls === 'overdue').length;
    const upcoming = APP.importShipments.filter(item => importEtaState(item).cls === 'alert').length;
    const sapPending = APP.importShipments.filter(item => item.etapa === 'recepcion_sap' && !(item.docs && item.docs.facturaSap && item.docs.gastosSap)).length;
    const currentOrigin = getImportFormValue('impOrigen') || 'espana';
    mount.innerHTML = `<div class="import-module">
      <section class="import-hero">
        <div><h2>Importaciones España y EE. UU. / Miami</h2><p>Seguimiento desde PO, documentación de origen, llegada a Miami, consolidación/envío a RD, aduana, transporte al almacén, recepción y registro de facturas/gastos en SAP.</p></div>
        <div class="import-actions"><button class="btn btn-success btn-sm" onclick="exportarImportacionesExcel()">Exportar reporte</button></div>
      </section>
      <div class="import-kpis">
        <div class="import-kpi"><strong>${active.length}</strong><span>Activos</span></div>
        <div class="import-kpi"><strong>${overdue}</strong><span>ETA vencida</span></div>
        <div class="import-kpi"><strong>${upcoming}</strong><span>ETA próxima</span></div>
        <div class="import-kpi"><strong>${sapPending}</strong><span>Pendiente SAP</span></div>
      </div>
      <section class="card" style="display:${canEdit ? 'block' : 'none'};">
        <div class="card-title">Carga automática de embarques</div>
        <div class="import-bulk-box">
          <strong>Sube Excel/CSV/TXT o una captura de pantalla</strong>
          <small>El TMS buscará expediente, booking, partner, destino, POA, vessel, ETD, ETA, equipo, fecha/hora de carga, situación, POD y contenedor. Si el embarque ya existe, lo actualiza; si no existe, lo crea.</small>
          <div class="import-bulk-actions">
            <input id="impBulkFile" class="form-control" type="file" accept=".xlsx,.xls,.csv,.txt,image/*" onchange="importarEmbarquesDesdeArchivo(this.files[0])">
          </div>
          <textarea id="impBulkText" class="form-control import-bulk-text" placeholder="También puedes pegar aquí la tabla o el texto OCR de los embarques..."></textarea>
          <div class="import-bulk-actions"><button class="btn btn-outline btn-sm" onclick="importarEmbarquesDesdeTexto()">Registrar texto pegado</button><span id="impBulkStatus" class="import-bulk-status"></span></div>
        </div>
      </section>
      <section class="card" style="display:${canEdit ? 'block' : 'none'};">
        <div class="card-title">Nuevo / editar embarque</div>
        <div class="import-form-grid">
          <select id="impOrigen" class="form-control" onchange="cambiarDocsImportacionOrigen()"><option value="espana">España</option><option value="miami">EE. UU. / Miami</option></select>
          <select id="impEtapa" class="form-control">${IMPORT_STAGES.map(stage => `<option value="${stage.key}">${stage.label}</option>`).join('')}</select>
          <input id="impReferencia" class="form-control" placeholder="Referencia / expediente">
          <input id="impProveedor" class="form-control" placeholder="Proveedor / fabricante">
          <input id="impPO" class="form-control" placeholder="PO">
          <input id="impFactura" class="form-control" placeholder="Factura proveedor">
          <input id="impBL" class="form-control" placeholder="BL / Waybill">
          <input id="impTracking" class="form-control" placeholder="Tracking / transitario">
          <label class="form-row">ETA RD<input id="impETA" class="form-control" type="date"></label>
          <label class="form-row">Llegada Miami<input id="impMiami" class="form-control" type="date"></label>
          <label class="form-row">Llegada aduana<input id="impAduana" class="form-control" type="date"></label>
          <label class="form-row">Salida aduana<input id="impSalidaAduana" class="form-control" type="date"></label>
          <label class="form-row">Llegada almacén<input id="impAlmacen" class="form-control" type="date"></label>
          <input id="impValor" class="form-control" type="number" step="0.01" placeholder="Valor estimado">
          <input id="impNotas" class="form-control" placeholder="Notas / bloqueo / próximo paso">
        </div>
        <div style="margin-top:10px;"><div class="card-title" style="font-size:12px;margin-bottom:8px;">Documentos y pagos</div><div id="impDocsWrap" class="import-docs">${importDocsFormHtml(currentOrigin, {})}</div></div>
        <div class="import-actions" style="margin-top:12px;"><button id="impSaveBtn" class="btn btn-primary btn-sm" onclick="guardarImportacion()">Crear embarque</button><button class="btn btn-outline btn-sm" onclick="clearImportForm()">Limpiar</button></div>
      </section>
      <section class="card">
        <div class="view-toolbar" style="justify-content:space-between;"><input id="impSearch" class="search-input" placeholder="Buscar referencia, proveedor, PO, factura o BL" oninput="renderImportaciones()" style="max-width:340px;"><select id="impFilterOrigen" class="search-input" onchange="renderImportaciones()" style="max-width:180px;"><option value="">Todos los orígenes</option><option value="espana">España</option><option value="miami">EE. UU. / Miami</option></select></div>
        <div class="import-stage-board">${IMPORT_STAGES.map(stage => {
          const stageRows = rows.filter(item => item.etapa === stage.key);
          return `<div class="import-stage"><div class="import-stage-head"><span>${stage.label}</span><span>${stageRows.length}</span></div><div class="import-stage-body">${stageRows.length ? stageRows.map(importCardHtml).join('') : '<div class="route-catalog-empty">Sin embarques.</div>'}</div></div>`;
        }).join('')}</div>
      </section>
      <section class="import-report">${buildImportDailyReport()}</section>
    </div>`;
    const origin = document.getElementById('impOrigen');
    if (origin) origin.value = currentOrigin;
    const searchEl = document.getElementById('impSearch');
    if (searchEl) searchEl.value = q;
    const filterEl = document.getElementById('impFilterOrigen');
    if (filterEl) filterEl.value = originFilter;
  };

  function initV2UI() {
    ensureAdminProfile();
    injectV2Styles();
    applyBrandLogo();
    ensureAppFooters();
    applyVisualTheme(APP.visualTheme || localStorage.getItem(THEME_KEY) || 'light', true);
    initUiLabels();
    enhanceImportUi();
    initImportSolicitudesCard();
    initImportExcelTable();
    initCalendarActions();
    initCommercialMount();
    initReportFilters();
    initControlHistoryPanel();
    initThemeSettings();
    initConfigExtras();
    renderImportaciones();
    syncMoveOptions();
    initRouteFilterOptions();
    renderRouteCatalog();
    renderControlHistoryPanel();
    updateSolicitudesImportStatus();
    actualizarEstadoBanner();
    updateUndoButton();
    applyPermissionUi();
  }

  window.addEventListener('DOMContentLoaded', function () {
    initV2UI();
    if (!APP.lineItems.length && !APP.planLoaded) {
      renderCalendario();
      renderComercial();
      renderSolicitudesAlmacen();
      renderAlmacen();
      renderConfigClientes();
      renderRouteCatalog();
    }
  });

  window.iniciarConDemo = function iniciarConDemoDisabled() {
    alert('Los datos demo fueron deshabilitados. Importa una planificación real para comenzar.');
  };
})();

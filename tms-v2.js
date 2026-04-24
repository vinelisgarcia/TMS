(function () {
  'use strict';

  if (typeof APP === 'undefined') return;

  const LOCAL_KEY = 'tms_alvarez_v2_state';
  const CURRENT_YEAR = new Date().getFullYear();
  const RD_HOLIDAYS = {
    2026: [
      '01/01/2026','06/01/2026','21/01/2026','26/01/2026','27/02/2026',
      '03/04/2026','04/04/2026','01/05/2026','04/06/2026','16/08/2026',
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
    solicitudesHistory: APP.solicitudesHistory || [],
    solicitudesPlanAlmacen: APP.solicitudesPlanAlmacen || [],
    solicitudesControlAlmacen: APP.solicitudesControlAlmacen || [],
    solicitudesCompare: APP.solicitudesCompare || [],
    warehouseSettings: APP.warehouseSettings || { costoSolicitud: 0 },
    userProfiles: APP.userProfiles || [],
    currentUserProfile: APP.currentUserProfile || null,
    camionExtraEnabled: !!APP.camionExtraEnabled,
    feriadosRD: APP.feriadosRD || (RD_HOLIDAYS[CURRENT_YEAR] || []),
    importFiles: APP.importFiles || { plan: '', control: '', solicitudesPlan: '', solicitudesControl: '' },
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
      buildBaseItemKey(item)
    ].join('|');
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
    return APP.routeConfigs.find(r => text(r.nombre) === name || text(r.zona) === name) || null;
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
      rutas: { ver: true, editar: true },
      reportes: { ver: true, editar: true },
      comercial: { ver: true, editar: true },
      configuracion: { ver: true, editar: true, importar: true },
      solicitudesAlmacen: { ver: true, editar: true },
      almacen: { ver: true, editar: true }
    };
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
        permisosPorModulo: buildAdminPermissions()
      };
      APP.userProfiles.unshift(adminProfile);
    } else {
      if (!text(adminProfile.nombre) || adminProfile.nombre === 'Manuel Oñate') {
        adminProfile.nombre = 'Vinelis Garcia';
      }
      if (!text(adminProfile.email) || adminProfile.email === 'manuel.onate@tms-alvarez.local') {
        adminProfile.email = fallbackEmail;
      }
      adminProfile.permisosPorModulo = adminProfile.permisosPorModulo || buildAdminPermissions();
    }
    APP.currentUserProfile =
      APP.userProfiles.find(profile => text(profile.email).toLowerCase() === fallbackEmail.toLowerCase()) ||
      adminProfile ||
      APP.userProfiles[0];
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

  async function replaceSupabaseTable(tableName, rows) {
    const client = getSupabaseClient();
    await runSupabaseQuery(client.from(tableName).delete().not('id', 'is', null), 'No se pudo limpiar ' + tableName);
    for (const chunk of chunkRows(rows, 200)) {
      await runSupabaseQuery(client.from(tableName).insert(chunk), 'No se pudo guardar ' + tableName);
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
      montoPlanificado: num(metadata.montoPlanificado),
      estadoPlanificacion: row.estado_planificacion || 'planificado',
      estadoEntrega: row.estado_entrega || '',
      origen: row.origen || 'planificacion semanal',
      camionAsignado: row.camion_asignado || '',
      fechaCreacion: row.created_at || nowIso(),
      fechaActualizacion: row.updated_at || nowIso(),
      zona: metadata.zona || row.ruta_nombre || '',
      observaciones: metadata.observaciones || '',
      incidencia: metadata.incidencia || '',
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
          observaciones: item.observaciones || '',
          incidencia: item.incidencia || '',
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
        tipoSolicitudArchivo: item.tipoSolicitudArchivo || 'plan semanal'
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
        tipoSolicitudArchivo: item.tipoSolicitudArchivo || 'control diario'
      })
    }));
  }

  function hasPermission(moduleName, action) {
    const profile = APP.currentUserProfile;
    if (!profile || profile.rol === 'Admin') return true;
    return !!(profile.permisosPorModulo && profile.permisosPorModulo[moduleName] && profile.permisosPorModulo[moduleName][action]);
  }

  function isHoliday(fechaStr) {
    return APP.feriadosRD.includes(fechaStr);
  }

  function isNonWorkingDay(fechaStr) {
    const d = fechaToDate(fechaStr);
    return d.getDay() === 0 || d.getDay() === 6 || isHoliday(fechaStr);
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
      group.monto += item.montoPlanificado || 0;
      group.pendienteFacturar += item.cantidadPendiente || 0;
      group.totalPendiente += (item.montoPlanificado || item.cantidadSolicitada || 0);
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

  function parseSAPRows(arrayBuffer, fileName, mode) {
    const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellDates: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const fechaArchivo = extraerFechaDeNombreArchivo(sheetName) || extraerFechaDeNombreArchivo(fileName) || fechaToStr(new Date());
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const items = [];
    rows.forEach((row, index) => {
      const clienteId = text(pickField(row, ['Cliente', 'Codigo', 'Código', 'ID cliente', 'Cliente código']));
      const clienteNombre = text(pickField(row, ['Nombre Comercial', 'Nombre', 'Cliente Nombre', 'Destinatario', 'Nombre cliente']));
      if (!clienteId || !clienteNombre) return;

      const pedidoCliente = text(pickField(row, ['Pedido del cliente', 'Pedido de cliente', 'Pedido cliente', 'Pedido']));
      const lineaPedidoCliente = text(pickField(row, ['Línea de pedido del cliente', 'Linea de pedido del cliente', 'Linea pedido cliente', 'Posición', 'Linea']));
      const articulo = text(pickField(row, ['Artículo', 'Articulo', 'Material', 'SKU']));
      const descripcionArticulo = text(pickField(row, ['Descripción del artículo', 'Descripcion del articulo', 'Descripción', 'Descripcion', 'Texto breve']));
      const cantidadSolicitada = num(pickField(row, ['Cantidad solicitada', 'Cantidad solicitada cliente', 'Cantidad', 'Cant solicitada']));
      const cantidadFacturada = num(pickField(row, ['Cantidad facturada', 'Facturado', 'Cantidad entregada']));
      const cantidadPendienteRaw = pickField(row, ['Cantidad pendiente', 'Pendiente', 'Cantidad abierta']);
      const cantidadPendiente = cantidadPendienteRaw !== '' ? num(cantidadPendienteRaw) : Math.max(cantidadSolicitada - cantidadFacturada, 0);
      const montoPlanificado = num(pickField(row, ['Importe confirmado no entregado', 'Importe Confirmado No Entregado', 'Confirmado No Entregado', 'Monto']));
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
        articulo,
        descripcionArticulo,
        cantidadSolicitada,
        cantidadFacturada,
        cantidadPendiente,
        montoPlanificado,
        estadoPlanificacion: mode === 'control' ? '' : 'planificado',
        estadoEntrega: '',
        origen: mode === 'control' ? 'control diario' : 'planificacion semanal',
        camionAsignado: '',
        fechaCreacion: nowIso(),
        fechaActualizacion: nowIso(),
        zona: rutaNombre,
        observaciones: '',
        manualProgramado: false,
        asignado: false
      };
      item.baseKey = buildBaseItemKey(item);
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

  function saveLocalSnapshot() {
    const payload = {
      savedAt: nowIso(),
      app: {
        lineItems: APP.lineItems,
        planLineItems: APP.planLineItems,
      controlLineItems: APP.controlLineItems,
      solicitudesAlmacen: APP.solicitudesAlmacen,
      solicitudesPlanAlmacen: APP.solicitudesPlanAlmacen,
      solicitudesControlAlmacen: APP.solicitudesControlAlmacen,
      solicitudesCompare: APP.solicitudesCompare,
      routeConfigs: APP.routeConfigs,
        routeCostHistory: APP.routeCostHistory,
        solicitudesHistory: APP.solicitudesHistory,
        warehouseSettings: APP.warehouseSettings,
        userProfiles: APP.userProfiles,
        clienteConfig: APP.clienteConfig,
        planLoaded: APP.planLoaded,
        planFecha: APP.planFecha,
        controlLoaded: APP.controlLoaded,
        controlFecha: APP.controlFecha,
        solicitudesLoaded: APP.solicitudesLoaded,
        solicitudesFileName: APP.solicitudesFileName,
        camionExtraEnabled: APP.camionExtraEnabled,
        feriadosRD: APP.feriadosRD,
        importFiles: APP.importFiles
      }
    };
    localStorage.setItem(LOCAL_KEY, JSON.stringify(payload));
  }

  function loadLocalSnapshot() {
    try {
      const raw = localStorage.getItem(LOCAL_KEY);
      if (!raw) return false;
      const payload = JSON.parse(raw);
      Object.assign(APP, payload.app || {});
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
      window.guardarEnSupabase().catch(e => console.warn('AutoSave error:', e));
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
          zona: item.rutaNombre || item.zona || '',
          rutaNombre: item.rutaNombre || item.zona || '',
          cantidadSolicitada: 0,
          cantidadFacturada: 0,
          cantidadPendiente: 0,
          totalPendiente: 0,
          items: [],
          cumplida: false,
          nota: '',
          incidencia: '',
          manualProgramado: !!item.manualProgramado,
          alertas: []
        });
      }
      const group = grouped.get(key);
      group.items.push(item);
      group.cantidadSolicitada += item.cantidadSolicitada || 0;
      group.cantidadFacturada += item.cantidadFacturada || 0;
      group.cantidadPendiente += item.cantidadPendiente || 0;
      group.totalPendiente += item.montoPlanificado || item.cantidadSolicitada || 0;
      group.cumplida = group.items.every(current => deriveItemStatus(current) === 'entregado');
      group.nota = group.nota || item.observaciones || '';
      group.incidencia = group.incidencia || item.incidencia || '';
    });

    grouped.forEach(group => {
      const cfg = getClienteConfigV2(group.codigo, group.nombre);
      if (!cfg.configuracionCompleta) group.alertas.push('Cliente sin configuración');
      if (!cfg.rutaAsignada && !group.rutaNombre) group.alertas.push('Cliente sin ruta asignada');
      if (isHoliday(group.fecha)) group.alertas.push('Pedido en feriado');
      if (!routes[group.fecha]) {
        routes[group.fecha] = {};
        getTruckOptions().forEach(camion => { routes[group.fecha][camion] = []; });
      }
      if (!routes[group.fecha][group.camion]) routes[group.fecha][group.camion] = [];
      routes[group.fecha][group.camion].push(group);
    });
    return routes;
  }

  window.buildCalStructure = function buildCalStructureV2() {
    const fechas = [...new Set(APP.lineItems.map(item => item.fechaPlanificada).filter(Boolean))].sort((a, b) => fechaToDate(a) - fechaToDate(b));
    APP.calSemanas = construirSemanasCalendario().filter(semana => semana.dias.some(fecha => fechas.includes(fecha)));
    if (!APP.calSemanas.length && fechas.length) {
      const inicio = lunesDeSemana(fechas[0]);
      APP.calSemanas = [{
        key: fechaToStr(inicio),
        nombre: fechas[0] + ' - ' + (fechas[fechas.length - 1] || fechas[0]),
        mesNombre: new Intl.DateTimeFormat('es-DO', { month: 'long', year: 'numeric' }).format(fechaToDate(fechas[0])),
        dias: fechas
      }];
    }
    if (APP.calSemanaIdx >= APP.calSemanas.length) APP.calSemanaIdx = Math.max(APP.calSemanas.length - 1, 0);
    APP.calMesIdx = APP.calSemanaIdx;
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
          monto: item.montoPlanificado || item.cantidadSolicitada || 0,
          pf: item.cantidadPendiente || 0,
          totalPendiente: item.montoPlanificado || item.cantidadPendiente || item.cantidadSolicitada || 0,
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
      return [
        item.nombre, item.codigo, item.pedidoCliente, item.articulo, item.descripcionArticulo, item.rutaNombre
      ].some(value => text(value).toLowerCase().includes(search));
    });

    section.style.display = 'block';
    resumen.textContent = visibles.length + ' pedidos pendientes';

    if (!visibles.length) {
      cont.innerHTML = '<div class="empty-state"><div class="empty-icon">📥</div><p>No hay pedidos pendientes en QUEUE para este filtro.</p></div>';
      return;
    }

    const grouped = {};
    visibles.forEach((item, index) => {
      const routeName = queueByRouteName(item);
      if (!grouped[routeName]) grouped[routeName] = [];
      grouped[routeName].push({ item, index });
    });

    cont.innerHTML = `<div class="rutas-grid">${Object.keys(grouped).sort().map(routeName => {
      const items = grouped[routeName];
      const warn = routeName.toLowerCase().includes('sin ruta');
      return `<div class="ruta-col">
        <div class="ruta-col-header ${warn ? 'ch-alm' : 'ch-c1'}">
          ${warn ? '⚠️' : '📍'} ${routeName} (${items.length})
        </div>
        <div class="ruta-col-body">
          ${items.map(({ item, index }) => {
            const cfg = getClienteConfigV2(item.codigo, item.nombre);
            const alerts = [];
            if (!cfg.configuracionCompleta) alerts.push('Cliente sin configuración');
            if (!cfg.rutaAsignada && !item.rutaNombre) alerts.push('Cliente sin ruta asignada');
            return `<div class="queue-card ${warn ? 'queue-card-new' : 'queue-card-plan'}" draggable="true"
              ondragstart="onDragStartQueue(event,${index})"
              ondragend="APP.dragQueue=null;this.classList.remove('dragging')">
              <div class="queue-card-title">${item.nombre}</div>
              <div class="queue-card-meta">${item.codigo} · Pedido ${item.pedidoCliente}${item.lineaPedidoCliente ? ' · Línea ' + item.lineaPedidoCliente : ''}</div>
              <div class="queue-card-meta" style="margin-top:6px;"><strong>${item.articulo || 'Sin artículo'}</strong> ${item.descripcionArticulo ? '· ' + item.descripcionArticulo : ''}</div>
              <div class="queue-card-meta" style="margin-top:6px;">Solicitado: <strong>${item.cantidadSolicitada}</strong> · Facturado: <strong>${item.cantidadFacturada}</strong> · Pendiente: <strong>${item.cantidadPendiente}</strong></div>
              ${alerts.length ? `<div style="font-size:11px;color:var(--danger);margin-top:6px;">${alerts.join(' · ')}</div>` : ''}
              <button class="btn btn-outline btn-sm" style="margin-top:8px;" onclick="abrirMoveQueueModal(${index})">Asignar fecha/camión</button>
            </div>`;
          }).join('')}
        </div>
      </div>`;
    }).join('')}</div>`;
  };

  window.renderCalendario = function renderCalendarioV2() {
    const calContainer = document.getElementById('calContainer');
    const calMesNav = document.getElementById('calMesNav');
    const calMesLabel = document.getElementById('calMesActualLabel');
    if (!calContainer || !calMesNav || !calMesLabel) return;

    if (!APP.calSemanas || !APP.calSemanas.length) {
      calContainer.innerHTML = '<div class="empty-state"><div class="empty-icon">📅</div><p>Importa una planificación semanal para comenzar.</p></div>';
      calMesNav.innerHTML = '';
      calMesLabel.textContent = '—';
      window.renderQueueSemana();
      return;
    }

    calMesNav.innerHTML = APP.calSemanas.map((semana, index) =>
      `<button class="filter-btn ${index === APP.calSemanaIdx ? 'active' : ''}" onclick="APP.calSemanaIdx=${index};APP.calMesIdx=${index};renderCalendario()">Semana ${index + 1}</button>`
    ).join('');

    const semana = APP.calSemanas[APP.calSemanaIdx];
    calMesLabel.textContent = semana.mesNombre.replace(/^\w/, char => char.toUpperCase());
    const search = text((document.getElementById('calSearch') || {}).value).toLowerCase();

    const html = [`<div style="font-size:13px;color:var(--muted);font-weight:700;margin:0 0 8px 2px;">Semana visible: ${semana.nombre}</div><div class="cal-semana">`];
    semana.dias.forEach(fecha => {
      const isHol = isHoliday(fecha);
      const dayRoutes = APP.rutas[fecha] || {};
      html.push(`<div class="cal-dia-col">
        <div class="cal-dia-header" style="${isHol ? 'background:#fde8e8;color:var(--danger);' : ''}">
          ${fechaLabel(fecha)}<br><small style="font-weight:400;opacity:0.8;">${fecha}${isHol ? ' · Feriado RD' : ''}</small>
        </div>`);

      getTruckOptions().forEach(camion => {
        const truckClass = camion === 'CAMION 1' ? 'ch-c1' : camion === 'CAMION 2' ? 'ch-c2' : camion === 'CAMION 3' ? 'ch-alm' : 'ch-alm';
        const cardClass = camion === 'CAMION 1' ? 'cal-card-c1' : camion === 'CAMION 2' ? 'cal-card-c2' : 'cal-card-alm';
        const groups = (dayRoutes[camion] || []).filter(group => {
          if (!search) return true;
          return [group.nombre, group.codigo, group.pedidoCliente, group.zona]
            .some(value => text(value).toLowerCase().includes(search));
        });
        html.push(`<div class="cal-carril"
          ondragover="event.preventDefault();this.classList.add('drag-over')"
          ondragleave="this.classList.remove('drag-over')"
          ondrop="onDropCal(event,'${fecha}','${camion}',this)">
          <span class="cal-carril-header ${truckClass}">${getTruckLabel(camion)} ${groups.length ? '(' + groups.length + ')' : ''}</span>`);

        groups.forEach((group, index) => {
          const realIdx = (dayRoutes[camion] || []).indexOf(group);
          const status = group.cumplida ? '✅' : group.cantidadFacturada > 0 ? '🟡' : '🔄';
          const alertHtml = group.alertas.length
            ? `<div style="font-size:10px;color:var(--danger);margin-top:4px;">${group.alertas.join(' · ')}</div>`
            : '';
          html.push(`<div class="cal-card ${cardClass}" draggable="true"
            ondragstart="onDragStartCal(event,'${fecha}','${camion}',${realIdx})"
            ondragend="this.classList.remove('dragging')"
            style="${group.alertas.length ? 'border:1px solid #ef4444;' : ''}">
            <div class="cal-card-name">${group.nombre}</div>
            <div class="queue-card-meta">${group.codigo} · Pedido ${group.pedidoCliente}</div>
            <div class="cal-card-monto">${group.items.length} líneas · ${group.cantidadPendiente} pendientes</div>
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
    html.push('</div>');
    calContainer.innerHTML = html.join('');
    window.renderQueueSemana();
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
                <div class="ruta-item-detail">${group.codigo} · Pedido ${group.pedidoCliente} · ${group.zona}</div>
                <div class="ruta-item-detail">${group.items.length} líneas · ${group.cantidadPendiente} pendientes</div>
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

  window.moverCliente = function moverClienteV2(fecha, camion, idx, nuevaFecha, nuevoCamion) {
    const group = getGroupByRouteRef(fecha, camion, idx);
    if (!group) return;
    const keys = new Set(group.items.map(item => item.baseKey));
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
    renderCalendario();
    renderRutas();
    renderComercial();
    renderAlmacen();
    scheduleAutoSave();
  };

  window.toggleCumplida = function toggleCumplidaV2(fecha, camion, idx) {
    const group = getGroupByRouteRef(fecha, camion, idx);
    if (!group) return;
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
    const dates = appendWorkingDates(fechaToStr(lunesDeSemana(fecha)), 15);
    document.getElementById('moveFecha').innerHTML = dates.map(date => `<option value="${date}" ${date === fecha ? 'selected' : ''}>${date} (${fechaLabel(date)})</option>`).join('');
    syncMoveOptions();
    document.getElementById('moveCamion').value = camion;
    document.getElementById('moveModal').classList.add('show');
  };

  window.abrirMoveQueueModal = function abrirMoveQueueModalV2(idx) {
    const item = APP.queueSemana[idx];
    if (!item) return;
    APP.moveCliente = { queueIdx: idx };
    document.getElementById('moveModalClienteNombre').textContent = `${item.nombre} · Pedido ${item.pedidoCliente}`;
    const dates = appendWorkingDates(APP.planFecha || fechaToStr(new Date()), 15);
    document.getElementById('moveFecha').innerHTML = dates.map((date, index) => `<option value="${date}" ${index === 0 ? 'selected' : ''}>${date} (${fechaLabel(date)})</option>`).join('');
    syncMoveOptions();
    document.getElementById('moveCamion').value = item.camion || getTruckOptions()[0];
    document.getElementById('moveModal').classList.add('show');
  };

  window.abrirNotaModal = function abrirNotaModalV2(fecha, camion, idx) {
    const group = getGroupByRouteRef(fecha, camion, idx);
    if (!group) return;
    APP.moveCliente = { fecha, cam: camion, idx };
    document.getElementById('notaModalNombre').textContent = `${group.nombre} · Pedido ${group.pedidoCliente}`;
    document.getElementById('notaInput').value = group.nota || '';
    document.getElementById('incidenciaInput').value = group.incidencia || '';
    document.getElementById('notaModal').classList.add('show');
  };

  window.guardarNota = function guardarNotaV2() {
    if (!APP.moveCliente) return;
    const group = getGroupByRouteRef(APP.moveCliente.fecha, APP.moveCliente.cam, APP.moveCliente.idx);
    if (!group) return;
    const nota = text(document.getElementById('notaInput').value);
    const incidencia = text(document.getElementById('incidenciaInput').value);
    const keys = new Set(group.items.map(item => item.baseKey));
    APP.lineItems.forEach(item => {
      if (!keys.has(item.baseKey)) return;
      item.observaciones = nota;
      item.incidencia = incidencia;
      item.fechaActualizacion = nowIso();
    });
    document.getElementById('notaModal').classList.remove('show');
    APP.moveCliente = null;
    rebuildDerivedState();
    renderCalendario();
    renderRutas();
    renderComercial();
    scheduleAutoSave();
  };

  function summarizeImport(mode, items, fileName, fecha) {
    const label = mode === 'plan' ? 'Planificación' : 'Control';
    const uniqueClients = new Set(items.map(item => item.clienteId)).size;
    const uniqueOrders = new Set(items.map(item => [item.clienteId, item.pedidoCliente].join('|'))).size;
    return `${label} cargada: ${fileName} · ${uniqueClients} clientes · ${uniqueOrders} pedidos · ${fecha}`;
  }

  function registerRouteCostSnapshot(controlDate) {
    const rows = [];
    Object.keys(APP.rutas).forEach(fecha => {
      if (fecha !== controlDate) return;
      Object.keys(APP.rutas[fecha] || {}).forEach(camion => {
        (APP.rutas[fecha][camion] || []).forEach(group => {
          const routeCfg = getRouteConfigByName(group.rutaNombre) || { nombre: group.rutaNombre, costeRuta: 0 };
          const summary = calcProgressSummary(group.items);
          rows.push({
            fecha,
            rutaId: routeCfg.id || group.rutaNombre,
            rutaNombre: group.rutaNombre,
            coste: num(routeCfg.costeRuta),
            fuenteControlDiario: APP.importFiles.control || '',
            pedidosAsociados: group.items.map(item => item.pedidoCliente).filter(Boolean),
            cumplimiento: summary.pct,
            importePlanificado: group.items.reduce((sum, item) => sum + (item.montoPlanificado || item.cantidadSolicitada || 0), 0),
            importeReal: group.items.reduce((sum, item) => sum + (item.cantidadFacturada || 0), 0),
            diferencia: group.items.reduce((sum, item) => sum + (item.cantidadPendiente || 0), 0)
          });
        });
      });
    });
    const merged = [...APP.routeCostHistory.filter(row => row.fecha !== controlDate), ...rows];
    APP.routeCostHistory = dedupeBy(merged, row => [row.fecha, row.rutaNombre].join('|'));
  }

  function registerSolicitudesSnapshot(fileName) {
    const fecha = fechaToStr(new Date());
    const rows = APP.solicitudesAlmacen.map(item => ({
      fecha,
      clienteId: item.codigo,
      clienteNombre: item.cliente,
      pedidoCliente: item.pedidoCliente || '',
      articulo: item.articulo || '',
      descripcionArticulo: item.descripcionArticulo || '',
      cantidadSolicitada: item.cantidadSolicitada || 0,
      cantidadCompletada: item.cantidadCompletada || 0,
      cantidadPendiente: item.cantidadPendiente || 0,
      estado: item.estado || '',
      costeSolicitud: num(APP.warehouseSettings.costoSolicitud || 0),
      observaciones: item.observaciones || '',
      fuente: fileName || '',
      tipo: item.tipoSolicitudArchivo || ''
    }));
    APP.solicitudesHistory = [...APP.solicitudesHistory.filter(row => row.fecha !== fecha), ...rows];
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
            item.fechaActualizacion = nowIso();
          });
        }

        rebuildDerivedState();
        if (!isPlan) registerRouteCostSnapshot(APP.controlFecha);
        actualizarDashboard();
        renderCalendario();
        renderRutas();
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
  };

  function parseWarehouseRows(rows) {
    return dedupeBy(rows.map(row => {
      const codigo = text(pickField(row, ['ID de destinatario de mercancías', 'ID destinatario', 'Código', 'Codigo']));
      const solicitud = text(pickField(row, ['ID de solicitud de logística por medio de terceros', 'ID de solicitud', 'Solicitud']));
      if (!codigo || !solicitud) return null;
      const cliente = text(pickField(row, ['Nombre de destinatario de mercancías', 'Cliente', 'Nombre destinatario']));
      const pedidoCliente = text(pickField(row, ['Pedido del cliente', 'Pedido de cliente', 'Pedido']));
      const articulo = text(pickField(row, ['Artículo', 'Articulo', 'Material', 'SKU']));
      const descripcionArticulo = text(pickField(row, ['Descripción del artículo', 'Descripcion del articulo', 'Descripción', 'Descripcion']));
      const cantidadSolicitada = num(pickField(row, ['Cantidad solicitada', 'Cantidad', 'Cant solicitada']));
      const cantidadCompletada = num(pickField(row, ['Cantidad completada', 'Cantidad facturada', 'Cantidad entregada']));
      const cantidadPendiente = Math.max(num(pickField(row, ['Cantidad pendiente', 'Pendiente'])) || (cantidadSolicitada - cantidadCompletada), 0);
      const fecha = toDateLabel(pickField(row, ['Fecha de envío planificada', 'Fecha de envio planificada', 'Fecha de liberación', 'Fecha de liberacion'])) || fechaToStr(new Date());
      return {
        codigo,
        cliente,
        solicitud,
        pedidoCliente,
        articulo,
        descripcionArticulo,
        cantidadSolicitada,
        cantidadCompletada,
        cantidadPendiente,
        estadoComunicacion: text(pickField(row, ['Estado de comunicación', 'Estado de comunicacion'])),
        estadoProceso: text(pickField(row, ['Estado de procesamiento', 'Estado proceso'])),
        fechaLiberacion: toDateLabel(pickField(row, ['Fecha de liberación', 'Fecha de liberacion'])),
        fechaFinalizada: toDateLabel(pickField(row, ['Fecha finalizada'])),
        fechaEnvio: fecha,
        ubicacion: text(pickField(row, ['Ubicación de procedencia', 'Ubicacion de procedencia'])),
        observaciones: text(pickField(row, ['Observaciones', 'Nota'])),
        estado: cantidadPendiente <= 0 ? 'Completada' : cantidadCompletada > 0 ? 'Parcial' : 'Pendiente',
        costeSolicitud: num(APP.warehouseSettings.costoSolicitud || 0)
      };
    }).filter(Boolean), item => [item.codigo, item.solicitud, item.pedidoCliente, item.articulo, item.fechaEnvio].join('|'));
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
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '', range: 4 });
        const parsed = parseWarehouseRows(rows).map(item => ({
          ...item,
          tipoSolicitudArchivo: mode === 'control' ? 'control diario' : 'plan semanal'
        }));
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
    const planMap = new Map(APP.solicitudesPlanAlmacen.map(item => [[item.codigo, item.solicitud, item.pedidoCliente, item.articulo].join('|'), item]).entries());
    const controlMap = new Map(APP.solicitudesControlAlmacen.map(item => [[item.codigo, item.solicitud, item.pedidoCliente, item.articulo].join('|'), item]).entries());
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
        Pedido: <strong>${item.pedidoCliente || '—'}</strong> · Artículo: <strong>${item.articulo || '—'}</strong><br>
        ${item.descripcionArticulo || 'Sin descripción'}<br>
        Solicitada: ${item.cantidadSolicitada} · Completada: ${item.cantidadCompletada} · Pendiente: ${item.cantidadPendiente}<br>
        ${item.observaciones ? 'Obs.: ' + item.observaciones + '<br>' : ''}
        Procedencia: ${item.ubicacion || '—'}
      </div>
    </div>`;
  }

  window.renderSolicitudCard = renderWarehouseCard;

  window.renderSolicitudesAlmacen = function renderSolicitudesAlmacenV2() {
    const cont = document.getElementById('solicitudesContainer');
    const resumen = document.getElementById('solicitudesResumen');
    if (!cont || !resumen) return;
    const visibleBase = APP.solicitudesControlAlmacen.length ? APP.solicitudesControlAlmacen : APP.solicitudesPlanAlmacen;
    resumen.textContent = APP.solicitudesLoaded ? `${visibleBase.length} solicitudes` : 'Sin archivo';
    if (!APP.solicitudesLoaded || !visibleBase.length) {
      cont.innerHTML = '<div class="empty-state"><div class="empty-icon">📦</div><p>No hay solicitudes registradas.</p></div>';
      return;
    }

    const search = text((document.getElementById('solSearch') || {}).value).toLowerCase();
    const filtered = visibleBase.filter(item => {
      if (!search) return true;
      return [item.codigo, item.cliente, item.solicitud, item.pedidoCliente, item.articulo]
        .some(value => text(value).toLowerCase().includes(search));
    });
    const byDate = {};
    filtered.forEach(item => {
      const date = item.fechaEnvio || item.fechaLiberacion || fechaToStr(new Date());
      if (!byDate[date]) byDate[date] = {};
      if (!byDate[date][item.codigo]) byDate[date][item.codigo] = [];
      byDate[date][item.codigo].push(item);
    });
    const comparison = (APP.solicitudesCompare || []).filter(item => {
      if (!search) return true;
      return [item.codigo, item.cliente, item.solicitud, item.pedidoCliente, item.articulo]
        .some(value => text(value).toLowerCase().includes(search));
    });
    const incomplete = comparison.filter(item => item.cantidadPendiente > 0);

    cont.innerHTML = Object.keys(byDate).sort((a, b) => fechaToDate(a) - fechaToDate(b)).map(date => {
      const clients = byDate[date];
      return `<div class="card" style="margin-bottom:14px;">
        <div class="card-title" style="display:flex;justify-content:space-between;align-items:center;">
          <span>📅 ${date} · ${fechaLabel(date)}</span>
          <span class="badge badge-pend">${Object.keys(clients).length} clientes</span>
        </div>
        ${Object.keys(clients).sort().map(code => {
          const items = clients[code];
          return `<details style="margin-bottom:10px;" open>
            <summary style="cursor:pointer;font-weight:700;color:var(--primary);margin-bottom:8px;">${items[0].cliente || code} · ${code} (${items.length} solicitudes)</summary>
            ${items.map(renderWarehouseCard).join('')}
          </details>`;
        }).join('')}
      </div>`;
    }).join('') + `
      <div class="card" style="border-left:4px solid #6366f1;">
        <div class="card-title" style="color:#4f46e5;">Comparativo plan semanal vs control diario</div>
        ${comparison.length ? `<div class="table-wrap"><table class="data-table">
          <thead><tr><th>Fecha</th><th>Cliente</th><th>Solicitud</th><th>Artículo</th><th>Plan</th><th>Control</th><th>Pendiente</th><th>Estado</th></tr></thead>
          <tbody>${comparison.map(item => `<tr>
            <td>${item.fecha}</td>
            <td>${item.cliente || item.codigo}</td>
            <td>${item.solicitud}</td>
            <td>${item.articulo || item.descripcionArticulo || '—'}</td>
            <td>${item.cantidadSolicitada}</td>
            <td>${item.cantidadCompletada}</td>
            <td>${item.cantidadPendiente}</td>
            <td>${item.estado}</td>
          </tr>`).join('')}</tbody>
        </table></div>` : '<div class="empty-state" style="padding:20px;"><p>Carga planificación y control diario de solicitudes para ver el comparativo.</p></div>'}
      </div>
      <div class="card" style="border-left:4px solid #f59e0b;">
        <div class="card-title" style="color:#b45309;">Solicitudes incompletas de la semana</div>
        ${incomplete.length ? `<div class="table-wrap"><table class="data-table">
          <thead><tr><th>Fecha</th><th>Cliente</th><th>Solicitud</th><th>Artículo</th><th>Solicitada</th><th>Completada</th><th>Pendiente</th><th>Estado</th></tr></thead>
          <tbody>${incomplete.map(item => `<tr>
            <td>${item.fechaEnvio || item.fechaLiberacion || '—'}</td>
            <td>${item.cliente || item.codigo}</td>
            <td>${item.solicitud}</td>
            <td>${item.articulo || item.descripcionArticulo || '—'}</td>
            <td>${item.cantidadSolicitada}</td>
            <td>${item.cantidadCompletada}</td>
            <td>${item.cantidadPendiente}</td>
            <td>${item.estado}</td>
          </tr>`).join('')}</tbody>
        </table></div>` : '<div class="empty-state" style="padding:20px;"><p>Sin solicitudes incompletas.</p></div>'}
      </div>`;
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

  window.renderComercial = function renderComercialV2() {
    const mount = document.getElementById('comercialV2Mount');
    if (!mount) return;
    const clientFilter = text((document.getElementById('cfcol_nombre') || {}).value).toLowerCase();
    const dateFilter = text((document.getElementById('cfcol_fecha') || {}).value).toLowerCase();
    const zoneFilter = text((document.getElementById('cfcol_zona') || {}).value).toLowerCase();
    const stateFilter = text((document.getElementById('cfcol_estado') || {}).value).toLowerCase();

    let items = [...APP.lineItems];
    if (clientFilter) items = items.filter(item => text(item.clienteNombre).toLowerCase().includes(clientFilter) || text(item.clienteId).toLowerCase().includes(clientFilter));
    if (dateFilter) items = items.filter(item => text(item.fechaPlanificada).toLowerCase().includes(dateFilter));
    if (zoneFilter) items = items.filter(item => text(item.rutaNombre || item.zona).toLowerCase().includes(zoneFilter));
    if (stateFilter) items = items.filter(item => text(deriveItemStatus(item)).toLowerCase().includes(stateFilter));

    const summary = calcProgressSummary(items);
    document.getElementById('ck1').textContent = new Set(items.map(item => item.clienteId)).size;
    document.getElementById('ck2').textContent = summary.requested;
    document.getElementById('ck3').textContent = summary.pending;
    document.getElementById('ck4').textContent = summary.invoiced;

    if (!items.length) {
      mount.innerHTML = '<div class="empty-state"><div class="empty-icon">👔</div><p>No hay datos comerciales para los filtros aplicados.</p></div>';
      return;
    }

    const byDate = {};
    items.forEach(item => {
      const date = item.fechaPlanificada || 'Sin fecha';
      if (!byDate[date]) byDate[date] = {};
      if (!byDate[date][item.clienteId]) byDate[date][item.clienteId] = [];
      byDate[date][item.clienteId].push(item);
    });

    mount.innerHTML = `
      <div class="card" style="margin-bottom:16px;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
          <div>
            <div class="card-title" style="margin-bottom:4px;">Progreso de entrega</div>
            <div style="font-size:12px;color:var(--muted);">Entregado: ${summary.pct}% — ${summary.invoiced} de ${summary.requested} unidades facturadas / solicitadas</div>
          </div>
          <span class="badge badge-ok">${summary.pct}%</span>
        </div>
        <div style="margin-top:10px;background:#e5e7eb;border-radius:999px;height:12px;overflow:hidden;">
          <div style="height:100%;width:${Math.min(summary.pct, 100)}%;background:linear-gradient(90deg,var(--secondary),#22c55e);"></div>
        </div>
      </div>
      ${Object.keys(byDate).sort((a, b) => fechaToDate(a) - fechaToDate(b)).map(date => {
        const clients = byDate[date];
        return `<div class="card" style="margin-bottom:14px;">
          <div class="card-title">📅 ${date} · ${fechaLabel(date)}</div>
          ${Object.keys(clients).sort().map(code => {
            const clientItems = clients[code];
            const clientSummary = calcProgressSummary(clientItems);
            const routeName = clientItems[0].rutaNombre || clientItems[0].zona || 'Sin ruta';
            const byOrder = {};
            clientItems.forEach(item => {
              const orderKey = item.pedidoCliente || item.baseKey;
              if (!byOrder[orderKey]) byOrder[orderKey] = [];
              byOrder[orderKey].push(item);
            });
            return `<details style="margin-bottom:12px;" open>
              <summary style="cursor:pointer;font-weight:700;color:var(--primary);">${clientItems[0].clienteNombre} · ${code} · ${routeName} · ${clientSummary.pct}%</summary>
              <div style="margin-top:8px;">
                ${Object.keys(byOrder).map(orderKey => {
                  const orderItems = byOrder[orderKey];
                  return `<div class="card" style="padding:14px;margin-bottom:10px;">
                    <div style="display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:8px;">
                      <div><strong>Pedido ${orderKey}</strong></div>
                      <span class="badge ${orderItems.every(item => deriveItemStatus(item) === 'entregado') ? 'badge-ok' : orderItems.some(item => deriveItemStatus(item) === 'parcial') ? 'badge-warn' : 'badge-pend'}">${orderItems.every(item => deriveItemStatus(item) === 'entregado') ? 'Entregado' : orderItems.some(item => deriveItemStatus(item) === 'parcial') ? 'Parcial' : 'Pendiente'}</span>
                    </div>
                    <div class="table-wrap"><table class="data-table">
                      <thead><tr><th>Artículo</th><th>Descripción</th><th>Solicitada</th><th>Facturada</th><th>Pendiente</th><th>Estado</th></tr></thead>
                      <tbody>${orderItems.map(item => `<tr>
                        <td>${item.articulo || '—'}</td>
                        <td>${item.descripcionArticulo || '—'}</td>
                        <td>${item.cantidadSolicitada}</td>
                        <td>${item.cantidadFacturada}</td>
                        <td>${item.cantidadPendiente}</td>
                        <td>${deriveItemStatus(item)}</td>
                      </tr>`).join('')}</tbody>
                    </table></div>
                  </div>`;
                }).join('')}
              </div>
            </details>`;
          }).join('')}
        </div>`;
      }).join('')}
    `;
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
    list.innerHTML = filtered.map(item => {
      const cfg = getClienteConfigV2(item.codigo, item.nombre);
      return `<button class="config-client-item ${APP.configClienteSel === item.codigo ? 'active' : ''}" onclick="seleccionarConfigCliente('${item.codigo}')">
        <strong>${item.nombre}</strong>${cfg.configuracionCompleta ? '<span class="badge badge-ok" style="float:right;">Completa</span>' : '<span class="badge badge-warn" style="float:right;">Pendiente</span>'}<br>
        <span style="font-size:11px;color:var(--muted);">${item.codigo} · ${cfg.rutaAsignada || item.zona || 'Sin ruta'}</span>
      </button>`;
    }).join('');
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
        <button class="btn btn-primary" onclick="guardarConfigCliente()">💾 Guardar configuración</button>
        <button class="btn btn-outline" onclick="limpiarConfigCliente()">Limpiar</button>
      </div>
    `;
  };

  window.guardarConfigCliente = function guardarConfigClienteV2() {
    const codigo = APP.configClienteSel;
    if (!codigo) return;
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
    ensureRouteConfig(rutaAsignada);
    APP.lineItems.forEach(item => {
      if (item.clienteId !== codigo) return;
      item.rutaNombre = rutaAsignada || item.rutaNombre;
      item.zona = rutaAsignada || item.zona;
      item.fechaActualizacion = nowIso();
    });
    rebuildDerivedState();
    renderConfigClientes();
    renderCalendario();
    renderComercial();
    scheduleAutoSave();
  };

  window.limpiarConfigCliente = function limpiarConfigClienteV2() {
    if (!APP.configClienteSel) return;
    delete APP.clienteConfig[APP.configClienteSel];
    renderConfigClientes();
    scheduleAutoSave();
  };

  window.recalcularPlanConConfig = function recalcularPlanConConfigV2() {
    if (!APP.planLineItems.length) {
      alert('Importa una planificación semanal antes de recalcular.');
      return;
    }
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
          ensureRouteConfig(APP.clienteConfig[codigo].rutaAsignada);
          if (exists) updated += 1;
          else created += 1;
        });
        renderConfigClientes();
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
    mount.innerHTML = `
      <div class="card" style="margin-top:16px;">
        <div class="card-title">Rutas / zonas</div>
        <div class="table-wrap"><table class="data-table">
          <thead><tr><th>Nombre</th><th>Días operación</th><th>Coste</th><th>Activa</th><th>Observaciones</th></tr></thead>
          <tbody>${APP.routeConfigs.length ? APP.routeConfigs.map(route => `<tr>
            <td>${route.nombre || route.zona}</td>
            <td>${(route.diasOperacion || []).join(', ') || 'Todos'}</td>
            <td>${route.costeRuta || 0}</td>
            <td>${route.activa === false ? 'No' : 'Sí'}</td>
            <td>${route.observaciones || ''}</td>
          </tr>`).join('') : '<tr><td colspan="5">Sin rutas configuradas aún.</td></tr>'}</tbody>
        </table></div>
      </div>
      <div class="card">
        <div class="card-title">Roles y usuarios</div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-bottom:14px;">
          <label class="form-row">Nombre
            <input id="userNameInput" class="form-control" placeholder="Nombre completo">
          </label>
          <label class="form-row">Email
            <input id="userEmailInput" class="form-control" type="email" placeholder="correo@empresa.com">
          </label>
          <label class="form-row">Rol
            <select id="userRoleInput" class="form-control">
              <option value="Admin">Admin</option>
              <option value="Operaciones">Operaciones</option>
              <option value="Comercial">Comercial</option>
              <option value="Almacén">Almacén</option>
              <option value="Solo lectura">Solo lectura</option>
            </select>
          </label>
          <label class="form-row">Permisos
            <select id="userScopeInput" class="form-control">
              <option value="full">Acceso completo</option>
              <option value="read_only">Solo lectura</option>
              <option value="ops">Operación</option>
              <option value="commercial">Comercial</option>
            </select>
          </label>
        </div>
        <div style="margin-bottom:12px;">
          <button class="btn btn-primary btn-sm" onclick="crearUsuarioPerfil()">➕ Crear usuario</button>
        </div>
        <div class="table-wrap"><table class="data-table">
          <thead><tr><th>Nombre</th><th>Email</th><th>Rol</th><th>Permisos</th><th>Acciones</th></tr></thead>
          <tbody>${APP.userProfiles.map((profile, index) => `<tr>
            <td>${profile.nombre}</td>
            <td>${profile.email || '—'}</td>
            <td>${profile.rol}</td>
            <td>${profile.rol === 'Admin' ? 'Acceso total' : 'Personalizado'}</td>
            <td>${profile.rol === 'Admin' ? '<span class="badge badge-ok">Base</span>' : `<button class="btn btn-outline btn-sm" onclick="eliminarUsuarioPerfil(${index})">Eliminar</button>`}</td>
          </tr>`).join('')}</tbody>
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

  window.guardarCostesOperativos = function guardarCostesOperativos() {
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
      rutas: { ver: true, editar: true },
      reportes: { ver: true, editar: true },
      comercial: { ver: true, editar: true },
      configuracion: { ver: true, editar: true, importar: true },
      solicitudesAlmacen: { ver: true, editar: true },
      almacen: { ver: true, editar: true }
    };
    if (role === 'Admin' || scope === 'full') return all;
    if (scope === 'read_only') {
      return Object.fromEntries(Object.keys(all).map(key => [key, { ver: true, editar: false, importar: false }]));
    }
    if (scope === 'commercial') {
      return {
        importar: { ver: false, editar: false, importar: false },
        dashboard: { ver: true, editar: false },
        calendario: { ver: true, editar: false },
        rutas: { ver: false, editar: false },
        reportes: { ver: true, editar: false },
        comercial: { ver: true, editar: true },
        configuracion: { ver: false, editar: false, importar: false },
        solicitudesAlmacen: { ver: true, editar: false },
        almacen: { ver: false, editar: false }
      };
    }
    return {
      importar: { ver: true, editar: true, importar: true },
      dashboard: { ver: true, editar: false },
      calendario: { ver: true, editar: true },
      rutas: { ver: true, editar: true },
      reportes: { ver: true, editar: true },
      comercial: { ver: true, editar: false },
      configuracion: { ver: false, editar: false, importar: false },
      solicitudesAlmacen: { ver: true, editar: true },
      almacen: { ver: true, editar: true }
    };
  }

  window.crearUsuarioPerfil = function crearUsuarioPerfil() {
    const nombre = text(document.getElementById('userNameInput').value);
    const email = text(document.getElementById('userEmailInput').value);
    const rol = text(document.getElementById('userRoleInput').value) || 'Operaciones';
    const scope = text(document.getElementById('userScopeInput').value) || 'ops';
    if (!nombre || !email) {
      alert('Completa nombre y email para crear el usuario.');
      return;
    }
    if (APP.userProfiles.some(profile => text(profile.email).toLowerCase() === email.toLowerCase())) {
      alert('Ya existe un usuario con ese email.');
      return;
    }
    APP.userProfiles.push({
      userId: 'user-' + Date.now(),
      nombre,
      email,
      rol,
      permisosPorModulo: buildPermissions(scope, rol)
    });
    renderConfigAuxSections();
    scheduleAutoSave();
  };

  window.eliminarUsuarioPerfil = function eliminarUsuarioPerfil(index) {
    const profile = APP.userProfiles[index];
    if (!profile || profile.rol === 'Admin') return;
    APP.userProfiles.splice(index, 1);
    renderConfigAuxSections();
    scheduleAutoSave();
  };

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

    const routeHistory = APP.routeCostHistory.filter(row => {
      const iso = fechaStrToISO(row.fecha);
      if (routeFilter && row.rutaNombre !== routeFilter) return false;
      if (start && iso < start) return false;
      if (end && iso > end) return false;
      return true;
    });
    const warehouseHistory = APP.solicitudesHistory.filter(row => {
      const iso = fechaStrToISO(row.fecha);
      if (clientQuery && ![row.clienteId, row.clienteNombre].some(value => text(value).toLowerCase().includes(clientQuery))) return false;
      if (start && iso < start) return false;
      if (end && iso > end) return false;
      return true;
    });

    const progress = calcProgressSummary(items);
    const routeCostTotal = routeHistory.reduce((sum, row) => sum + num(row.coste), 0);
    const warehouseCostTotal = warehouseHistory.reduce((sum, row) => sum + num(row.costeSolicitud), 0);
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
      <div class="card">
        <div class="card-title">Gasto teórico / real de rutas</div>
        ${routeHistory.length ? `<div class="table-wrap"><table class="data-table">
          <thead><tr><th>Fecha</th><th>Ruta</th><th>Coste</th><th>Cumplimiento</th><th>Planificado</th><th>Real</th><th>Diferencia</th></tr></thead>
          <tbody>${routeHistory.map(row => `<tr>
            <td>${row.fecha}</td>
            <td>${row.rutaNombre}</td>
            <td>${formatMonto(row.coste)}</td>
            <td>${row.cumplimiento}%</td>
            <td>${formatMonto(row.importePlanificado)}</td>
            <td>${formatMonto(row.importeReal)}</td>
            <td>${formatMonto(row.diferencia)}</td>
          </tr>`).join('')}</tbody>
        </table></div>` : '<div class="empty-state" style="padding:20px;"><p>No hay histórico de rutas para ese rango.</p></div>'}
      </div>
      <div class="card">
        <div class="card-title">Solicitudes a almacén</div>
        ${warehouseHistory.length ? `<div class="table-wrap"><table class="data-table">
          <thead><tr><th>Fecha</th><th>Cliente</th><th>Artículo</th><th>Solicitada</th><th>Completada</th><th>Pendiente</th><th>Estado</th><th>Coste</th></tr></thead>
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
    const wsRoutes = XLSX.utils.json_to_sheet(APP.routeCostHistory);
    const wsWarehouse = XLSX.utils.json_to_sheet(APP.solicitudesHistory);
    XLSX.utils.book_append_sheet(wb, wsRoutes, 'Rutas');
    XLSX.utils.book_append_sheet(wb, wsWarehouse, 'Almacen');
    XLSX.writeFile(wb, 'TMS_Reportes.xlsx');
  };

  window.exportarComercialExcel = function exportarComercialExcelV2() {
    const data = APP.lineItems.map(item => ({
      'Cliente': item.clienteNombre,
      'Código cliente': item.clienteId,
      'Fecha': item.fechaPlanificada,
      'Ruta': item.rutaNombre,
      'Pedido': item.pedidoCliente,
      'Línea': item.lineaPedidoCliente,
      'Artículo': item.articulo,
      'Descripción': item.descripcionArticulo,
      'Cant. solicitada': item.cantidadSolicitada,
      'Cant. facturada': item.cantidadFacturada,
      'Cant. pendiente': item.cantidadPendiente,
      'Estado': deriveItemStatus(item)
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Comercial');
    XLSX.writeFile(wb, 'TMS_Comercial.xlsx');
  };

  window.exportarSolicitudesAlmacenExcel = function exportarSolicitudesAlmacenExcelV2() {
    const ws = XLSX.utils.json_to_sheet(APP.solicitudesAlmacen);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Solicitudes');
    XLSX.writeFile(wb, 'TMS_Solicitudes_Almacen.xlsx');
  };

  window.exportarExcelRutas = function exportarExcelRutasV2() {
    if (!APP.rutaFecha || !APP.rutas[APP.rutaFecha]) return alert('Selecciona una fecha para exportar.');
    const wb = XLSX.utils.book_new();
    getTruckOptions().forEach(camion => {
      const rows = (APP.rutas[APP.rutaFecha][camion] || []).map(group => ({
        'Cliente': group.nombre,
        'Código': group.codigo,
        'Pedido': group.pedidoCliente,
        'Ruta': group.rutaNombre,
        'Líneas': group.items.length,
        'Solicitada': group.cantidadSolicitada,
        'Facturada': group.cantidadFacturada,
        'Pendiente': group.cantidadPendiente
      }));
      if (!rows.length) return;
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), camion.replace(/\s+/g, ''));
    });
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
    const empty = !APP.lineItems.length;
    const dashContent = document.getElementById('dashContent');
    if (empty && dashContent) {
      dashContent.innerHTML = '<div class="empty-state"><div class="empty-icon">📊</div><p>No hay datos cargados todavía.</p></div>';
    }
  };

  window.cargarDesdeSupabase = async function cargarDesdeSupabaseV2() {
    try {
      const client = getSupabaseClient();
      const user = await getAuthenticatedUser();
      ensureAdminProfile();
      if (!client || !user) {
        if (loadLocalSnapshot()) {
          updateSolicitudesImportStatus();
          actualizarEstadoBanner();
          actualizarDashboard();
          renderCalendario();
          renderRutas();
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

      APP.userProfiles = userProfileRows.map(row => ({
        userId: row.id,
        authUserId: row.auth_user_id || '',
        nombre: row.nombre,
        email: row.email,
        rol: row.rol,
        permisosPorModulo: row.permisos_por_modulo || buildPermissions('read_only', row.rol),
        activo: row.activo !== false
      }));

      const settingsMap = Object.fromEntries(settingsRows.map(row => [row.clave, row.valor || {}]));
      APP.warehouseSettings = settingsMap.warehouse || { costoSolicitud: 0 };
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
        APP.userProfiles.find(profile => profile.rol === 'Admin') ||
        APP.userProfiles[0] ||
        null;

      rebuildDerivedState();
      updateSolicitudesImportStatus();
      actualizarEstadoBanner();
      actualizarDashboard();
      renderCalendario();
      renderRutas();
      renderComercial();
      renderConfigClientes();
      renderSolicitudesAlmacen();
      renderAlmacen();
      saveLocalSnapshot();
    } catch (error) {
      console.warn('Fallo la carga remota, usando respaldo local si existe:', error);
      if (loadLocalSnapshot()) {
        updateSolicitudesImportStatus();
        actualizarEstadoBanner();
        actualizarDashboard();
        renderCalendario();
        renderRutas();
        renderComercial();
        renderConfigClientes();
        renderSolicitudesAlmacen();
        renderAlmacen();
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
      renderComercial();
      renderConfigClientes();
      renderSolicitudesAlmacen();
      renderAlmacen();
    }
  };

  window.guardarEnSupabase = async function guardarEnSupabaseV2() {
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
        APP.currentUserProfile ||
        APP.userProfiles.find(profile => text(profile.email).toLowerCase() === text(user.email).toLowerCase()) ||
        APP.userProfiles.find(profile => profile.rol === 'Admin') ||
        null;

      const settingsRows = [
        { clave: 'warehouse', valor: safeClone(APP.warehouseSettings || { costoSolicitud: 0 }) },
        {
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
            importFiles: APP.importFiles || {}
          })
        }
      ];

      const userRows = APP.userProfiles.map(profile => ({
        auth_user_id: text(profile.email).toLowerCase() === text(user.email).toLowerCase() ? user.id : (isUuid(profile.authUserId) ? profile.authUserId : null),
        nombre: profile.nombre || '',
        email: profile.email || '',
        rol: profile.rol || 'Solo lectura',
        permisos_por_modulo: safeClone(profile.permisosPorModulo || buildPermissions('read_only', profile.rol)),
        activo: profile.activo !== false
      }));

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

      await runSupabaseQuery(client.from('tms_settings').upsert(settingsRows, { onConflict: 'clave' }), 'No se pudieron guardar los ajustes');
      await replaceSupabaseTable('tms_user_profiles', userRows);
      await replaceSupabaseTable('tms_route_configs', routeRows);
      await replaceSupabaseTable('tms_client_configs', clientRows);
      await replaceSupabaseTable('tms_order_lines', orderRows);
      await replaceSupabaseTable('tms_route_cost_history', routeHistoryRows);
      await replaceSupabaseTable('tms_warehouse_plan', warehousePlanRows);
      await replaceSupabaseTable('tms_warehouse_control', warehouseControlRows);
      await replaceSupabaseTable('tms_warehouse_history', warehouseHistoryRows);

      if (btn) {
        btn.textContent = '✅ Guardado en nube';
      }
      return true;
    } catch (error) {
      console.error('Error guardando en Supabase:', error);
      if (btn) btn.textContent = '⚠️ Guardado local';
      throw error;
    } finally {
      if (btn) {
        setTimeout(() => {
          btn.disabled = false;
          btn.textContent = originalText || '💾 Guardar cambios';
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
          <div style="font-size:12px;color:var(--muted);">Importa aquí la planificación semanal de solicitudes y luego el control diario para poder comparar cumplimiento.</div>
        </div>
      </div>
      <div class="view-toolbar" style="margin-top:14px;align-items:flex-start;">
        <div style="display:flex;flex-direction:column;gap:8px;">
          <input type="file" id="solicitudesPlanFileImport" accept=".xlsx,.xls" style="display:none" onchange="procesarSolicitudesAlmacen(this.files[0],'plan')">
          <button class="btn btn-warning" onclick="document.getElementById('solicitudesPlanFileImport').click()">📂 Cargar planificación de solicitudes</button>
          <div id="solPlanImportStatus" style="font-size:12px;color:var(--muted);">Sin archivo de planificación</div>
        </div>
        <div style="display:flex;flex-direction:column;gap:8px;">
          <input type="file" id="solicitudesControlFileImport" accept=".xlsx,.xls" style="display:none" onchange="procesarSolicitudesAlmacen(this.files[0],'control')">
          <button class="btn btn-outline" onclick="document.getElementById('solicitudesControlFileImport').click()">📂 Cargar control diario de solicitudes</button>
          <div id="solCtrlImportStatus" style="font-size:12px;color:var(--muted);">Sin archivo de control</div>
        </div>
        <span style="font-size:12px;color:var(--muted);" id="solicitudesImportHint">Campos soportados: pedido, artículo, descripción, cantidades y estados.</span>
      </div>`;
    importView.appendChild(card);
  }

  function updateSolicitudesImportStatus() {
    const planStatus = document.getElementById('solPlanImportStatus');
    const ctrlStatus = document.getElementById('solCtrlImportStatus');
    if (planStatus) planStatus.textContent = APP.importFiles.solicitudesPlan || 'Sin archivo de planificación';
    if (ctrlStatus) ctrlStatus.textContent = APP.importFiles.solicitudesControl || 'Sin archivo de control';
  }

  function initCalendarActions() {
    const toolbar = document.querySelector('#calendario .view-toolbar');
    if (!toolbar || document.getElementById('addTruckBtn')) return;
    const button = document.createElement('button');
    button.id = 'addTruckBtn';
    button.className = 'btn btn-outline btn-sm';
    button.textContent = '➕ Agregar tercer camión';
    button.onclick = function () {
      APP.camionExtraEnabled = true;
      syncMoveOptions();
      rebuildDerivedState();
      renderCalendario();
      renderRutas();
      scheduleAutoSave();
      this.disabled = true;
      this.textContent = 'Camión 3 activo';
    };
    toolbar.appendChild(button);
    if (APP.camionExtraEnabled) {
      button.disabled = true;
      button.textContent = 'Camión 3 activo';
    }
  }

  function initCommercialMount() {
    const card = document.querySelector('#comercial .card');
    if (!card || document.getElementById('comercialV2Mount')) return;
    const legacyTable = document.getElementById('comTable');
    if (legacyTable) legacyTable.closest('.table-wrap').style.display = 'none';
    const filterBar = document.createElement('div');
    filterBar.style.cssText = 'display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-top:8px;';
    filterBar.innerHTML = `
      <input id="cfcol_nombre" class="search-input" placeholder="Cliente / código" oninput="renderComercial()">
      <input id="cfcol_fecha" class="search-input" placeholder="Fecha DD/MM/AAAA" oninput="renderComercial()">
      <input id="cfcol_zona" class="search-input" placeholder="Ruta / zona" oninput="renderComercial()">
      <input id="cfcol_estado" class="search-input" placeholder="Estado: entregado, parcial..." oninput="renderComercial()">
    `;
    card.appendChild(filterBar);
    const mount = document.createElement('div');
    mount.id = 'comercialV2Mount';
    mount.style.marginTop = '12px';
    card.appendChild(mount);
  }

  function initReportFilters() {
    const toolbar = document.querySelector('#reportes .view-toolbar');
    if (!toolbar || document.getElementById('repDateStart')) return;
    toolbar.insertAdjacentHTML('beforeend', `
      <input type="date" id="repDateStart" class="search-input" style="max-width:160px;">
      <input type="date" id="repDateEnd" class="search-input" style="max-width:160px;">
      <select id="repRouteFilter" class="search-input" style="max-width:180px;"><option value="">Todas las rutas</option></select>
    `);
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
    if (titleConfig) titleConfig.textContent = '⚙️ Configuración';
    const titleImport = document.querySelector('#importar .view-title');
    if (titleImport) titleImport.textContent = '📥 Importar Datos';
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
      .form-control { padding:8px 10px; border:1px solid var(--border); border-radius:8px; }
      .queue-card { border:1px solid var(--border); border-radius:10px; padding:12px; background:#fff; margin-bottom:10px; }
      .queue-card-title { font-weight:700; color:var(--primary); font-size:13px; }
      .queue-card-meta { font-size:11px; color:var(--muted); }
      details summary { list-style:none; }
      details summary::-webkit-details-marker { display:none; }
    `;
    document.head.appendChild(style);
  }

  function initV2UI() {
    ensureAdminProfile();
    injectV2Styles();
    initUiLabels();
    initImportSolicitudesCard();
    initCalendarActions();
    initCommercialMount();
    initReportFilters();
    initConfigExtras();
    syncMoveOptions();
    initRouteFilterOptions();
    updateSolicitudesImportStatus();
    actualizarEstadoBanner();
  }

  window.addEventListener('DOMContentLoaded', function () {
    initV2UI();
    if (!APP.lineItems.length && !APP.planLoaded) {
      renderCalendario();
      renderComercial();
      renderSolicitudesAlmacen();
      renderAlmacen();
      renderConfigClientes();
    }
  });

  window.iniciarConDemo = function iniciarConDemoDisabled() {
    alert('Los datos demo fueron deshabilitados. Importa una planificación real para comenzar.');
  };
})();

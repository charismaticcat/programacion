/**
 * CRMService.gs — Pipeline, Kanban, actividades, seguimientos y bandeja de solicitudes.
 * Esta es información INTERNA: nunca se mezcla con los datos públicos del cliente
 * ni aparece en el reporte para el diseño de la página.
 */
const OUTGOING_TYPES = ['LLAMADA', 'WHATSAPP', 'EMAIL', 'REUNION', 'VISITA'];

const CRMService = {
  // ================= Oportunidades =================
  _clientMap() {
    const map = {};
    SearchIndex.records().forEach(o => { map[o.CLIENTE_ID] = o; });
    return map;
  },

  _requireOpp(user, id) {
    const o = SheetService.getById('OPORTUNIDADES', id);
    if (!o) throw new AppError('NOT_FOUND', 'Oportunidad no encontrada.');
    ClientService.requireAccessible(user, o.CLIENTE_ID);
    return o;
  },

  /** Oportunidades para el Kanban/listado (abiertas + cerradas en los últimos N días). */
  listOpportunities(user, filters) {
    filters = filters || {};
    const days = Math.min(Math.max(Number(filters.closedDays) || 60, 7), 730);
    const since = Utils.addDays(null, -days);
    const clients = CRMService._clientMap();
    const q = Utils.normalizeText(filters.q || '');
    const items = SheetService.readAll('OPORTUNIDADES').filter(o => {
      const c = clients[o.CLIENTE_ID];
      if (!c) return false;
      if (!Auth.canAccessClient(user, c)) return false;
      if (!o.ABIERTA && String(o.FECHA_CIERRE_REAL || o.ACTUALIZADO_EN).slice(0, 10) < since) return false;
      if (filters.responsable === 'me' && o.RESPONSABLE_EMAIL !== user.email) return false;
      if (filters.responsable && filters.responsable !== 'me' && o.RESPONSABLE_EMAIL !== filters.responsable) return false;
      if (filters.etapa && o.ETAPA !== filters.etapa) return false;
      if (filters.productoWebId && o.PRODUCTO_WEB_ID !== filters.productoWebId) return false;
      if (q && Utils.normalizeText([o.TITULO, c.RAZON_SOCIAL, c.NOMBRE_COMERCIAL].join(' ')).indexOf(q) < 0) return false;
      return true;
    }).map(o => {
      const c = clients[o.CLIENTE_ID];
      delete o._row;
      o.CLIENTE_NOMBRE = c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL;
      o.CLIENTE_TELEFONO = c.TELEFONO;
      o.CLIENTE_MUNICIPIO = c.MUNICIPIO;
      o.DIAS_EN_ETAPA = Math.floor(Utils.minutesBetween(o.FECHA_CAMBIO_ETAPA || o.CREADO_EN, Utils.nowIso()) / 1440);
      return o;
    });
    // Próximo seguimiento pendiente por oportunidad.
    const nextFollow = {};
    SheetService.readColumns('SEGUIMIENTOS', ['OPORTUNIDAD_ID', 'FECHA_VENCIMIENTO', 'ESTADO']).forEach(s => {
      if (s.ESTADO !== 'PENDIENTE' || !s.OPORTUNIDAD_ID) return;
      if (!nextFollow[s.OPORTUNIDAD_ID] || s.FECHA_VENCIMIENTO < nextFollow[s.OPORTUNIDAD_ID]) nextFollow[s.OPORTUNIDAD_ID] = s.FECHA_VENCIMIENTO;
    });
    items.forEach(o => { o.PROXIMO_SEGUIMIENTO = nextFollow[o.OPORTUNIDAD_ID] || ''; });
    const stages = ENUMS.ETAPA.map(e => ({ key: e, label: STAGES[e].label, prob: STAGES[e].prob, closed: STAGES[e].closed }));
    const totals = {};
    ENUMS.ETAPA.forEach(e => {
      const list = items.filter(o => o.ETAPA === e);
      totals[e] = { count: list.length, value: list.reduce((s, o) => s + Number(o.VALOR_ESTIMADO || 0), 0) };
    });
    return { items: items, stages: stages, totals: totals };
  },

  createOpportunity(user, data) {
    data = data || {};
    const c = ClientService.requireAccessible(user, String(data.CLIENTE_ID || ''));
    const etapa = ENUMS.ETAPA.indexOf(data.ETAPA) >= 0 ? data.ETAPA : 'POR_CONTACTAR';
    if (STAGES[etapa].closed) throw new AppError('VALIDATION', 'Una oportunidad nueva no puede crearse cerrada.');
    const clean = Validation.cleanRecord('OPORTUNIDADES', data);
    let producto = null;
    if (clean.PRODUCTO_WEB_ID) {
      producto = SheetService.getById('PRODUCTOS_WEB', clean.PRODUCTO_WEB_ID);
      if (!producto) throw new AppError('VALIDATION', 'Producto web no válido.');
    }
    const rec = Object.assign({
      CLIENTE_ID: c.CLIENTE_ID,
      TITULO: 'Página web · ' + ClientService.displayName(c),
      ORIGEN_LEAD: c.ORIGEN === 'BASE_CCH' ? 'BASE' : c.ORIGEN,
      FECHA_CIERRE_ESTIMADA: Utils.addDays(null, 30)
    }, clean, {
      ETAPA: etapa,
      PROBABILIDAD: clean.PROBABILIDAD || STAGES[etapa].prob,
      RESPONSABLE_EMAIL: c.RESPONSABLE_EMAIL || user.email,
      FECHA_CAMBIO_ETAPA: Utils.nowIso(),
      ABIERTA: true
    });
    if (!rec.VALOR_ESTIMADO && producto) rec.VALOR_ESTIMADO = producto.PRECIO_SUGERIDO || 0;
    SheetService.insert('OPORTUNIDADES', rec);
    Audit.log({ accion: 'CREAR', entidad: 'OPORTUNIDADES', entidadId: rec.OPORTUNIDAD_ID, clienteId: c.CLIENTE_ID, despues: { TITULO: rec.TITULO, ETAPA: etapa, VALOR_ESTIMADO: rec.VALOR_ESTIMADO } });
    if (!c.RESPONSABLE_EMAIL) {
      const r = SheetService.update('CLIENTES', c.CLIENTE_ID, { RESPONSABLE_EMAIL: rec.RESPONSABLE_EMAIL });
      SearchIndex.upsert(r.record);
    }
    ClientService.syncEstadoCrm(c.CLIENTE_ID);
    return ClientService._strip(rec);
  },

  updateOpportunity(user, id, patch, expectedVersion) {
    const o = CRMService._requireOpp(user, id);
    const clean = Validation.cleanRecord('OPORTUNIDADES', patch || {});
    if (clean.PRODUCTO_WEB_ID && !SheetService.getById('PRODUCTOS_WEB', clean.PRODUCTO_WEB_ID)) {
      throw new AppError('VALIDATION', 'Producto web no válido.');
    }
    const res = SheetService.update('OPORTUNIDADES', id, clean, { expectedVersion: expectedVersion });
    Audit.logChange('EDITAR', 'OPORTUNIDADES', id, o.CLIENTE_ID, res);
    return ClientService._strip(res.record);
  },

  /**
   * Cambia la etapa (Kanban). GANADO exige venta registrada; PERDIDO y
   * NO_INTERESADO exigen motivo.
   */
  moveOpportunity(user, id, etapa, expectedVersion, extra) {
    extra = extra || {};
    const o = CRMService._requireOpp(user, id);
    Validation.requireEnum(etapa, 'ETAPA', 'Etapa');
    if (o.ETAPA === etapa) return ClientService._strip(o);
    const st = STAGES[etapa];
    const patch = {
      ETAPA: etapa, PROBABILIDAD: st.prob, ABIERTA: !st.closed, FECHA_CAMBIO_ETAPA: Utils.nowIso(),
      FECHA_CIERRE_REAL: st.closed ? Utils.today() : ''
    };
    if (etapa === 'GANADO') {
      const ventas = SheetService.findAllBy('VENTAS', 'OPORTUNIDAD_ID', id).filter(v => v.ESTADO_PAGO !== 'ANULADA');
      if (!ventas.length) throw new AppError('SALE_REQUIRED', 'Para marcar como Ganado registre primero la venta.');
    }
    if (etapa === 'PERDIDO' || etapa === 'NO_INTERESADO') {
      const motivo = Validation.long(extra.motivo, 1000);
      if (!motivo) throw new AppError('REASON_REQUIRED', 'Escriba el motivo del cierre.');
      patch.MOTIVO_CIERRE = motivo;
    }
    const res = SheetService.update('OPORTUNIDADES', id, patch, { expectedVersion: expectedVersion });
    SheetService.insert('ACTIVIDADES', {
      CLIENTE_ID: o.CLIENTE_ID, OPORTUNIDAD_ID: id, TIPO: 'CAMBIO_ETAPA', DIRECCION: 'INTERNA', RESULTADO: '',
      ASUNTO: STAGES[o.ETAPA].label + ' → ' + st.label, DETALLE: patch.MOTIVO_CIERRE || '', DURACION_MIN: 0,
      FECHA_HORA: Utils.nowIso(), REALIZADA_POR: user.email
    });
    Audit.log({ accion: 'CAMBIO_ETAPA', entidad: 'OPORTUNIDADES', entidadId: id, clienteId: o.CLIENTE_ID, campos: res.changed, antes: res.before, despues: res.after });
    ClientService.syncEstadoCrm(o.CLIENTE_ID);
    return ClientService._strip(res.record);
  },

  /** "Cambiar etapa" desde la ficha del cliente: mueve su oportunidad abierta o crea una. */
  setClientStage(user, clienteId, etapa, extra) {
    const c = ClientService.requireAccessible(user, clienteId);
    Validation.requireEnum(etapa, 'ETAPA', 'Etapa');
    const open = SheetService.findAllBy('OPORTUNIDADES', 'CLIENTE_ID', c.CLIENTE_ID).filter(o => o.ABIERTA)
      .sort((a, b) => String(b.FECHA_CAMBIO_ETAPA).localeCompare(String(a.FECHA_CAMBIO_ETAPA)));
    if (open.length) return CRMService.moveOpportunity(user, open[0].OPORTUNIDAD_ID, etapa, null, extra);
    if (STAGES[etapa].closed) {
      if (etapa === 'GANADO') throw new AppError('SALE_REQUIRED', 'Cree una oportunidad y registre la venta.');
      const opp = CRMService.createOpportunity(user, { CLIENTE_ID: c.CLIENTE_ID, ETAPA: 'POR_CONTACTAR' });
      return CRMService.moveOpportunity(user, opp.OPORTUNIDAD_ID, etapa, null, extra);
    }
    return CRMService.createOpportunity(user, { CLIENTE_ID: c.CLIENTE_ID, ETAPA: etapa });
  },

  // ================= Actividades =================
  logActivity(user, data) {
    data = data || {};
    const c = ClientService.requireAccessible(user, String(data.CLIENTE_ID || ''));
    const tipo = Validation.requireEnum(data.TIPO, 'ACT_TIPO', 'Tipo');
    if (tipo === 'CAMBIO_ETAPA' || tipo === 'SISTEMA') throw new AppError('VALIDATION', 'Tipo reservado.');
    const clean = Validation.cleanRecord('ACTIVIDADES', data);
    const direccion = clean.DIRECCION || (tipo === 'NOTA' ? 'INTERNA' : 'SALIENTE');
    if (c.NO_CONTACTAR && direccion === 'SALIENTE' && OUTGOING_TYPES.indexOf(tipo) >= 0) {
      throw new AppError('DO_NOT_CONTACT', 'Este cliente pidió no ser contactado (NO CONTACTAR).');
    }
    let oppId = String(data.OPORTUNIDAD_ID || '');
    if (oppId) {
      const o = SheetService.getById('OPORTUNIDADES', oppId);
      if (!o || o.CLIENTE_ID !== c.CLIENTE_ID) throw new AppError('VALIDATION', 'La oportunidad no pertenece a este cliente.');
    }
    if (clean.RESULTADO === 'VOLVER_A_LLAMAR' && !(data.seguimiento && data.seguimiento.FECHA_VENCIMIENTO)) {
      throw new AppError('FOLLOWUP_REQUIRED', 'Indique la fecha para volver a llamar (seguimiento).');
    }
    if (!clean.ASUNTO && !clean.DETALLE) throw new AppError('VALIDATION', 'Escriba el asunto o el detalle.');

    // Regla: primera actividad saliente mueve Nuevo/Por contactar → Contactado.
    if (direccion === 'SALIENTE' && OUTGOING_TYPES.indexOf(tipo) >= 0) {
      const opps = SheetService.findAllBy('OPORTUNIDADES', 'CLIENTE_ID', c.CLIENTE_ID);
      const open = opps.filter(o => o.ABIERTA);
      if (!oppId && open.length) oppId = open[0].OPORTUNIDAD_ID;
      const target = open.find(o => o.OPORTUNIDAD_ID === oppId);
      if (target && (target.ETAPA === 'NUEVO' || target.ETAPA === 'POR_CONTACTAR')) {
        CRMService.moveOpportunity(user, target.OPORTUNIDAD_ID, 'CONTACTADO');
      } else if (!opps.length) {
        const created = CRMService.createOpportunity(user, { CLIENTE_ID: c.CLIENTE_ID, ETAPA: 'CONTACTADO' });
        oppId = created.OPORTUNIDAD_ID;
      }
    }
    const rec = Object.assign({}, clean, {
      CLIENTE_ID: c.CLIENTE_ID, OPORTUNIDAD_ID: oppId, TIPO: tipo, DIRECCION: direccion,
      FECHA_HORA: clean.FECHA_HORA || Utils.nowIso(), REALIZADA_POR: user.email
    });
    SheetService.insert('ACTIVIDADES', rec);
    Audit.log({ accion: 'ACTIVIDAD', entidad: 'ACTIVIDADES', entidadId: rec.ACTIVIDAD_ID, clienteId: c.CLIENTE_ID, despues: { TIPO: tipo, RESULTADO: rec.RESULTADO, ASUNTO: rec.ASUNTO } });
    let seguimiento = null;
    if (data.seguimiento && data.seguimiento.FECHA_VENCIMIENTO) {
      seguimiento = CRMService.saveFollowup(user, Object.assign({
        CLIENTE_ID: c.CLIENTE_ID, OPORTUNIDAD_ID: oppId, ACTIVIDAD_ORIGEN_ID: rec.ACTIVIDAD_ID,
        TIPO: tipo === 'WHATSAPP' ? 'WHATSAPP' : (tipo === 'EMAIL' ? 'EMAIL' : 'LLAMAR')
      }, data.seguimiento));
    }
    return { actividad: ClientService._strip(rec), seguimiento: seguimiento };
  },

  listActivities(user, filters) {
    filters = filters || {};
    const clients = CRMService._clientMap();
    const q = Utils.normalizeText(filters.q || '');
    return SheetService.readAll('ACTIVIDADES').filter(a => {
      const c = clients[a.CLIENTE_ID];
      if (!c || !Auth.canAccessClient(user, c)) return false;
      if (filters.clienteId && a.CLIENTE_ID !== filters.clienteId) return false;
      if (filters.tipo && a.TIPO !== filters.tipo) return false;
      if (filters.realizadaPor === 'me' && a.REALIZADA_POR !== user.email) return false;
      if (filters.desde && String(a.FECHA_HORA).slice(0, 10) < filters.desde) return false;
      if (filters.hasta && String(a.FECHA_HORA).slice(0, 10) > filters.hasta) return false;
      if (q && Utils.normalizeText([a.ASUNTO, a.DETALLE, c.RAZON_SOCIAL].join(' ')).indexOf(q) < 0) return false;
      return true;
    }).sort((a, b) => String(b.FECHA_HORA).localeCompare(String(a.FECHA_HORA))).slice(0, 300).map(a => {
      delete a._row;
      const c = clients[a.CLIENTE_ID];
      a.CLIENTE_NOMBRE = c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL;
      return a;
    });
  },

  // ================= Seguimientos =================
  saveFollowup(user, data) {
    data = data || {};
    const key = data.SEGUIMIENTO_ID;
    if (key) {
      const s = SheetService.getById('SEGUIMIENTOS', key);
      if (!s) throw new AppError('NOT_FOUND', 'Seguimiento no encontrado.');
      ClientService.requireAccessible(user, s.CLIENTE_ID);
      const clean = Validation.cleanRecord('SEGUIMIENTOS', data);
      if (clean.ASIGNADO_A) CRMService._requireActiveUserEmail(clean.ASIGNADO_A);
      const res = SheetService.update('SEGUIMIENTOS', key, clean, { expectedVersion: data.VERSION });
      Audit.logChange('SEGUIMIENTO', 'SEGUIMIENTOS', key, s.CLIENTE_ID, res);
      return ClientService._strip(res.record);
    }
    const c = ClientService.requireAccessible(user, String(data.CLIENTE_ID || ''));
    const clean = Validation.cleanRecord('SEGUIMIENTOS', data);
    if (!clean.FECHA_VENCIMIENTO) throw new AppError('VALIDATION', 'Indique la fecha de vencimiento.');
    if (!clean.TIPO) clean.TIPO = 'LLAMAR';
    const asignado = clean.ASIGNADO_A || c.RESPONSABLE_EMAIL || user.email;
    CRMService._requireActiveUserEmail(asignado);
    const oppId = String(data.OPORTUNIDAD_ID || '');
    if (oppId) {
      const o = SheetService.getById('OPORTUNIDADES', oppId);
      if (!o || o.CLIENTE_ID !== c.CLIENTE_ID) throw new AppError('VALIDATION', 'La oportunidad no pertenece a este cliente.');
    }
    const rec = Object.assign({ PRIORIDAD: 'MEDIA' }, clean, {
      CLIENTE_ID: c.CLIENTE_ID, OPORTUNIDAD_ID: oppId, ACTIVIDAD_ORIGEN_ID: String(data.ACTIVIDAD_ORIGEN_ID || ''),
      ESTADO: 'PENDIENTE', ASIGNADO_A: asignado
    });
    SheetService.insert('SEGUIMIENTOS', rec);
    Audit.log({ accion: 'SEGUIMIENTO', entidad: 'SEGUIMIENTOS', entidadId: rec.SEGUIMIENTO_ID, clienteId: c.CLIENTE_ID, despues: { TIPO: rec.TIPO, FECHA_VENCIMIENTO: rec.FECHA_VENCIMIENTO, ASIGNADO_A: asignado } });
    return ClientService._strip(rec);
  },

  _requireActiveUserEmail(email) {
    const u = Auth.findUserByEmail(email);
    if (!u || !u.ACTIVO) throw new AppError('VALIDATION', 'El usuario asignado no existe o está inactivo.');
  },

  completeFollowup(user, id, resultado, estado) {
    const s = SheetService.getById('SEGUIMIENTOS', id);
    if (!s) throw new AppError('NOT_FOUND', 'Seguimiento no encontrado.');
    ClientService.requireAccessible(user, s.CLIENTE_ID);
    const nuevo = estado === 'CANCELADO' ? 'CANCELADO' : 'COMPLETADO';
    const res = SheetService.update('SEGUIMIENTOS', id, {
      ESTADO: nuevo, COMPLETADO_EN: Utils.nowIso(), RESULTADO: Validation.long(resultado, 2000)
    });
    Audit.logChange('SEGUIMIENTO_' + nuevo, 'SEGUIMIENTOS', id, s.CLIENTE_ID, res);
    return ClientService._strip(res.record);
  },

  /** bucket: vencidos | hoy | proximos | completados | todos. asignado: me | all | email */
  listFollowups(user, bucket, asignado) {
    const today = Utils.today();
    const clients = CRMService._clientMap();
    const items = SheetService.readAll('SEGUIMIENTOS').filter(s => {
      const c = clients[s.CLIENTE_ID];
      if (!c || !Auth.canAccessClient(user, c)) return false;
      if (asignado === 'me' && s.ASIGNADO_A !== user.email) return false;
      if (asignado && asignado !== 'me' && asignado !== 'all' && s.ASIGNADO_A !== asignado) return false;
      const pend = s.ESTADO === 'PENDIENTE';
      switch (bucket) {
        case 'vencidos': return pend && s.FECHA_VENCIMIENTO < today;
        case 'hoy': return pend && s.FECHA_VENCIMIENTO === today;
        case 'proximos': return pend && s.FECHA_VENCIMIENTO > today;
        case 'completados': return !pend;
        default: return true;
      }
    }).map(s => {
      delete s._row;
      const c = clients[s.CLIENTE_ID];
      s.CLIENTE_NOMBRE = c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL;
      s.CLIENTE_TELEFONO = c.TELEFONO;
      s.VENCIDO = s.ESTADO === 'PENDIENTE' && s.FECHA_VENCIMIENTO < today;
      return s;
    });
    items.sort((a, b) => bucket === 'completados'
      ? String(b.COMPLETADO_EN).localeCompare(String(a.COMPLETADO_EN))
      : String(a.FECHA_VENCIMIENTO).localeCompare(String(b.FECHA_VENCIMIENTO)));
    return items.slice(0, 500);
  },

  // ================= Formulario enviado =================
  onFormSubmitted(cliente, sol, isNew, created) {
    const open = SheetService.findAllBy('OPORTUNIDADES', 'CLIENTE_ID', cliente.CLIENTE_ID).filter(o => o.ABIERTA);
    let oppId;
    if (open.length) {
      oppId = open[0].OPORTUNIDAD_ID;
    } else {
      const rec = {
        CLIENTE_ID: cliente.CLIENTE_ID, TITULO: 'Página web · ' + ClientService.displayName(cliente), ETAPA: 'NUEVO',
        PROBABILIDAD: STAGES.NUEVO.prob, VALOR_ESTIMADO: 0, FECHA_CIERRE_ESTIMADA: Utils.addDays(null, 30),
        ORIGEN_LEAD: 'FORMULARIO', RESPONSABLE_EMAIL: cliente.RESPONSABLE_EMAIL || '', FECHA_CAMBIO_ETAPA: Utils.nowIso(), ABIERTA: true
      };
      SheetService.insert('OPORTUNIDADES', rec);
      created.push({ table: 'OPORTUNIDADES', id: rec.OPORTUNIDAD_ID });
      oppId = rec.OPORTUNIDAD_ID;
      Audit.log({ accion: 'CREAR', entidad: 'OPORTUNIDADES', entidadId: oppId, clienteId: cliente.CLIENTE_ID, detalle: 'Automática por formulario ' + sol.RADICADO });
    }
    const act = {
      CLIENTE_ID: cliente.CLIENTE_ID, OPORTUNIDAD_ID: oppId, TIPO: 'SISTEMA', DIRECCION: 'ENTRANTE', RESULTADO: '',
      ASUNTO: 'Formulario recibido ' + sol.RADICADO, DETALLE: isNew ? 'Cliente nuevo desde el formulario.' : 'Cliente de la base (' + sol.TIPO_VINCULACION + ').',
      DURACION_MIN: 0, FECHA_HORA: Utils.nowIso(), REALIZADA_POR: 'SISTEMA'
    };
    SheetService.insert('ACTIVIDADES', act);
    created.push({ table: 'ACTIVIDADES', id: act.ACTIVIDAD_ID });
    const assignee = cliente.RESPONSABLE_EMAIL || Settings.get('DEFAULT_ASSIGNEE') || '';
    const seg = {
      CLIENTE_ID: cliente.CLIENTE_ID, OPORTUNIDAD_ID: oppId, ACTIVIDAD_ORIGEN_ID: act.ACTIVIDAD_ID, TIPO: 'REVISAR_SOLICITUD',
      DESCRIPCION: 'Revisar formulario ' + sol.RADICADO + ' y el reporte generado.', FECHA_VENCIMIENTO: Utils.addDays(null, 1),
      PRIORIDAD: 'ALTA', ESTADO: 'PENDIENTE', ASIGNADO_A: assignee
    };
    SheetService.insert('SEGUIMIENTOS', seg);
    created.push({ table: 'SEGUIMIENTOS', id: seg.SEGUIMIENTO_ID });
    ClientService.syncEstadoCrm(cliente.CLIENTE_ID);
  },

  // ================= Bandeja de solicitudes =================
  listSolicitudes(user, filters) {
    filters = filters || {};
    const clients = CRMService._clientMap();
    const cols = ['SOLICITUD_ID', 'RADICADO', 'CLIENTE_ID', 'CLIENTE_BASE_ID', 'TIPO_VINCULACION', 'ESTADO', 'PORCENTAJE',
      'CAMBIOS_ESTADO', 'REPORTE_ESTADO', 'INICIADA_EN', 'ENVIADA_EN', 'REVISADO_POR', 'ACTUALIZADO_EN'];
    const estado = filters.estado || 'ENVIADAS';
    return SheetService.readColumns('SOLICITUDES', cols).filter(s => {
      if (estado === 'ENVIADAS' && s.ESTADO === 'BORRADOR') return false;
      if (estado !== 'ENVIADAS' && estado !== 'TODAS' && s.ESTADO !== estado) return false;
      const cid = s.CLIENTE_ID || s.CLIENTE_BASE_ID;
      if (Auth.isScopedToOwn(user)) {
        const c = clients[cid];
        if (!c || !Auth.canAccessClient(user, c)) return false;
      }
      if (filters.cambiosPendientes && s.CAMBIOS_ESTADO !== 'PENDIENTE') return false;
      return true;
    }).map(s => {
      delete s._row;
      const c = clients[s.CLIENTE_ID || s.CLIENTE_BASE_ID];
      s.CLIENTE_NOMBRE = c ? (c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL) : '';
      return s;
    }).sort((a, b) => String(b.ENVIADA_EN || b.ACTUALIZADO_EN).localeCompare(String(a.ENVIADA_EN || a.ACTUALIZADO_EN))).slice(0, 500);
  },

  getSolicitud(user, id) {
    const s = SheetService.getById('SOLICITUDES', id);
    if (!s) throw new AppError('NOT_FOUND', 'Solicitud no encontrada.');
    const cid = s.CLIENTE_ID || s.CLIENTE_BASE_ID;
    if (cid) ClientService.requireAccessible(user, cid);
    else if (Auth.isScopedToOwn(user)) throw new AppError('FORBIDDEN', 'Solicitud no asignada.');
    const data = s.DATOS_FILE_ID ? DriveService.readJson(s.DATOS_FILE_ID) : s.DATOS_JSON;
    const labels = {
      NOMBRE_PROPIETARIO: 'Nombre del propietario', RAZON_SOCIAL: 'Razón social', NOMBRE_COMERCIAL: 'Nombre comercial',
      TIPO_ESTABLECIMIENTO: 'Tipo de establecimiento', ACTIVIDAD: 'Actividad', TELEFONO: 'Teléfono', EMAIL: 'Email',
      MUNICIPIO: 'Municipio', DIRECCION: 'Dirección', BARRIO: 'Barrio'
    };
    const cambios = (s.CAMBIOS_PROPUESTOS_JSON || []).map(c => Object.assign({ etiqueta: labels[c.campo] || c.campo }, c));
    return {
      SOLICITUD_ID: s.SOLICITUD_ID, RADICADO: s.RADICADO, CLIENTE_ID: s.CLIENTE_ID, CLIENTE_BASE_ID: s.CLIENTE_BASE_ID,
      TIPO_VINCULACION: s.TIPO_VINCULACION, ESTADO: s.ESTADO, PORCENTAJE: s.PORCENTAJE, INICIADA_EN: s.INICIADA_EN,
      ENVIADA_EN: s.ENVIADA_EN, CAMBIOS_ESTADO: s.CAMBIOS_ESTADO, REPORTE_ESTADO: s.REPORTE_ESTADO,
      OBSERVACIONES_ADMIN: s.OBSERVACIONES_ADMIN, REVISADO_POR: s.REVISADO_POR, REVISADO_EN: s.REVISADO_EN,
      VERSION: s.VERSION, cambios: cambios, datos: data || {},
      snapshotUrl: DriveService.fileUrl(s.SNAPSHOT_FILE_ID), carpetaUrl: DriveService.folderUrl(s.DRIVE_FOLDER_ID)
    };
  },

  /** Aplica (o descarta) las correcciones propuestas por un cliente de la base. */
  applyCorrections(user, solicitudId, campos, descartar) {
    const s = SheetService.getById('SOLICITUDES', solicitudId);
    if (!s || s.CAMBIOS_ESTADO !== 'PENDIENTE') throw new AppError('NOT_FOUND', 'No hay correcciones pendientes.');
    const c = ClientService.requireAccessible(user, s.CLIENTE_ID);
    if (descartar) {
      const r = SheetService.update('SOLICITUDES', solicitudId, { CAMBIOS_ESTADO: 'DESCARTADO', REVISADO_POR: user.email, REVISADO_EN: Utils.nowIso() });
      Audit.logChange('DESCARTAR_CORRECCIONES', 'SOLICITUDES', solicitudId, c.CLIENTE_ID, r);
      return { applied: [] };
    }
    const wanted = Array.isArray(campos) ? campos : [];
    const patch = {};
    (s.CAMBIOS_PROPUESTOS_JSON || []).forEach(ch => {
      if (wanted.indexOf(ch.campo) >= 0) patch[ch.campo] = ch.propuesto;
    });
    let result = { changed: [] };
    if (Object.keys(patch).length) result = ClientService.update(user, c.CLIENTE_ID, patch);
    const r = SheetService.update('SOLICITUDES', solicitudId, { CAMBIOS_ESTADO: 'APLICADO', REVISADO_POR: user.email, REVISADO_EN: Utils.nowIso() });
    Audit.logChange('APLICAR_CORRECCIONES', 'SOLICITUDES', solicitudId, c.CLIENTE_ID, r, 'Campos: ' + Object.keys(patch).join(','));
    return { applied: result.changed };
  },

  setSolicitudEstado(user, solicitudId, estado, observaciones) {
    const s = SheetService.getById('SOLICITUDES', solicitudId);
    if (!s) throw new AppError('NOT_FOUND', 'Solicitud no encontrada.');
    if (s.ESTADO === 'BORRADOR') throw new AppError('VALIDATION', 'La solicitud aún no ha sido enviada.');
    if (s.CLIENTE_ID) ClientService.requireAccessible(user, s.CLIENTE_ID);
    if (['ENVIADA', 'EN_REVISION', 'APROBADA', 'DESCARTADA'].indexOf(estado) < 0) throw new AppError('VALIDATION', 'Estado no válido.');
    const patch = { ESTADO: estado, REVISADO_POR: user.email, REVISADO_EN: Utils.nowIso() };
    if (observaciones !== undefined) patch.OBSERVACIONES_ADMIN = Validation.long(observaciones, 3000);
    const res = SheetService.update('SOLICITUDES', solicitudId, patch);
    Audit.logChange('CAMBIO_ESTADO', 'SOLICITUDES', solicitudId, s.CLIENTE_ID, res);
    const map = { EN_REVISION: 'EN_REVISION', APROBADA: 'APROBADO', ENVIADA: 'ENVIADO' };
    if (s.CLIENTE_ID && map[estado]) {
      const r = SheetService.update('CLIENTES', s.CLIENTE_ID, { ESTADO_FORMULARIO: map[estado] });
      SearchIndex.upsert(r.record);
      Audit.logChange('CAMBIO_ESTADO', 'CLIENTES', s.CLIENTE_ID, s.CLIENTE_ID, r);
    }
    return { ESTADO: estado };
  },

  /** Enlace de WhatsApp con la plantilla de invitación. */
  whatsappLink(user, clienteId) {
    const c = ClientService.requireAccessible(user, clienteId);
    if (!c.TELEFONO) throw new AppError('VALIDATION', 'El cliente no tiene teléfono.');
    if (c.NO_CONTACTAR) throw new AppError('DO_NOT_CONTACT', 'Este cliente pidió no ser contactado.');
    let url = Settings.get('PUBLIC_FORM_URL');
    if (!url) {
      try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
    }
    const msg = Settings.get('WHATSAPP_PLANTILLA')
      .replace(/\{NOMBRE\}/g, c.NOMBRE_PROPIETARIO || ClientService.displayName(c))
      .replace(/\{URL\}/g, url);
    return { url: 'https://wa.me/57' + c.TELEFONO + '?text=' + encodeURIComponent(msg), mensaje: msg };
  }
};

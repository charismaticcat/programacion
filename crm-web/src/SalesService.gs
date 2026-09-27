/**
 * SalesService.gs — Propuestas, ventas (cierres), pagos y proyectos web.
 * Todos los valores son editables (D14); los totales se recalculan en el servidor.
 */
const SalesService = {
  _requireOpp(user, id) {
    const o = SheetService.getById('OPORTUNIDADES', id);
    if (!o) throw new AppError('NOT_FOUND', 'Oportunidad no encontrada.');
    ClientService.requireAccessible(user, o.CLIENTE_ID);
    return o;
  },

  /** Limpia ítems [{descripcion, cantidad, precioUnitario, productoWebId}] y calcula subtotal. */
  _items(raw) {
    const list = Array.isArray(raw) ? raw.slice(0, 50) : [];
    const items = [];
    list.forEach(it => {
      if (!Utils.isPlainObject(it)) return;
      const descripcion = Validation.short(it.descripcion, 300);
      if (!descripcion) return;
      const cantidad = Math.max(1, Math.min(1000, Math.round(Number(it.cantidad) || 1)));
      const precio = Utils.parseMoney(it.precioUnitario);
      items.push({
        descripcion: descripcion, cantidad: cantidad, precioUnitario: precio,
        productoWebId: /^PWB-[A-Z0-9]+$/.test(String(it.productoWebId || '')) ? it.productoWebId : '',
        total: cantidad * precio
      });
    });
    if (!items.length) throw new AppError('VALIDATION', 'Agregue al menos un ítem con descripción.');
    return { items: items, subtotal: items.reduce((s, i) => s + i.total, 0) };
  },

  // ================= Propuestas =================
  saveProposal(user, data) {
    data = data || {};
    const calc = SalesService._items(data.items);
    const descuento = Math.min(Utils.parseMoney(data.DESCUENTO), calc.subtotal);
    const fields = {
      ITEMS_JSON: calc.items, SUBTOTAL: calc.subtotal, DESCUENTO: descuento, TOTAL: calc.subtotal - descuento,
      VALIDEZ_HASTA: Utils.isIsoDate(data.VALIDEZ_HASTA) ? data.VALIDEZ_HASTA : Utils.addDays(null, 15),
      NOTAS: Validation.long(data.NOTAS, 3000)
    };
    if (data.PROPUESTA_ID) {
      const p = SheetService.getById('PROPUESTAS', data.PROPUESTA_ID);
      if (!p) throw new AppError('NOT_FOUND', 'Propuesta no encontrada.');
      ClientService.requireAccessible(user, p.CLIENTE_ID);
      if (p.ESTADO === 'ACEPTADA') throw new AppError('VALIDATION', 'Una propuesta aceptada no se puede editar; cree una nueva.');
      const res = SheetService.update('PROPUESTAS', p.PROPUESTA_ID, fields, { expectedVersion: data.VERSION });
      Audit.logChange('EDITAR', 'PROPUESTAS', p.PROPUESTA_ID, p.CLIENTE_ID, res);
      return ClientService._strip(res.record);
    }
    const o = SalesService._requireOpp(user, String(data.OPORTUNIDAD_ID || ''));
    const rec = Object.assign({
      OPORTUNIDAD_ID: o.OPORTUNIDAD_ID, CLIENTE_ID: o.CLIENTE_ID,
      NUMERO: 'PRO-' + Utils.year() + '-' + Seq.pad(Seq.next('PROPUESTA_' + Utils.year()), 4),
      ESTADO: 'BORRADOR'
    }, fields);
    SheetService.insert('PROPUESTAS', rec);
    Audit.log({ accion: 'CREAR', entidad: 'PROPUESTAS', entidadId: rec.PROPUESTA_ID, clienteId: o.CLIENTE_ID, despues: { NUMERO: rec.NUMERO, TOTAL: rec.TOTAL } });
    if (!o.VALOR_ESTIMADO || o.ETAPA === 'NUEVO') {
      const r = SheetService.update('OPORTUNIDADES', o.OPORTUNIDAD_ID, { VALOR_ESTIMADO: rec.TOTAL });
      Audit.logChange('EDITAR', 'OPORTUNIDADES', o.OPORTUNIDAD_ID, o.CLIENTE_ID, r, 'Valor desde propuesta ' + rec.NUMERO);
    }
    return ClientService._strip(rec);
  },

  setProposalStatus(user, id, estado) {
    const p = SheetService.getById('PROPUESTAS', id);
    if (!p) throw new AppError('NOT_FOUND', 'Propuesta no encontrada.');
    ClientService.requireAccessible(user, p.CLIENTE_ID);
    Validation.requireEnum(estado, 'PROP_ESTADO', 'Estado');
    const patch = { ESTADO: estado };
    if (estado === 'ENVIADA') patch.ENVIADA_EN = Utils.nowIso();
    if (estado === 'ACEPTADA' || estado === 'RECHAZADA') patch.RESPUESTA_EN = Utils.nowIso();
    const res = SheetService.update('PROPUESTAS', id, patch);
    Audit.logChange('CAMBIO_ESTADO', 'PROPUESTAS', id, p.CLIENTE_ID, res);
    const o = SheetService.getById('OPORTUNIDADES', p.OPORTUNIDAD_ID);
    if (o && o.ABIERTA) {
      if (estado === 'ENVIADA' && ['NUEVO', 'POR_CONTACTAR', 'CONTACTADO', 'INTERESADO', 'REUNION'].indexOf(o.ETAPA) >= 0) {
        CRMService.moveOpportunity(user, o.OPORTUNIDAD_ID, 'PROPUESTA_ENVIADA');
      }
      const r2 = SheetService.update('OPORTUNIDADES', o.OPORTUNIDAD_ID, { VALOR_ESTIMADO: p.TOTAL });
      Audit.logChange('EDITAR', 'OPORTUNIDADES', o.OPORTUNIDAD_ID, o.CLIENTE_ID, r2);
    }
    return ClientService._strip(res.record);
  },

  listProposals(user, filters) {
    filters = filters || {};
    const clients = CRMService._clientMap();
    return SheetService.readAll('PROPUESTAS').filter(p => {
      const c = clients[p.CLIENTE_ID];
      if (!c || !Auth.canAccessClient(user, c)) return false;
      if (filters.estado && p.ESTADO !== filters.estado) return false;
      if (filters.clienteId && p.CLIENTE_ID !== filters.clienteId) return false;
      return true;
    }).map(p => {
      delete p._row;
      const c = clients[p.CLIENTE_ID];
      p.CLIENTE_NOMBRE = c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL;
      p.VENCIDA_CALC = p.ESTADO === 'ENVIADA' && p.VALIDEZ_HASTA && p.VALIDEZ_HASTA < Utils.today();
      return p;
    }).sort((a, b) => String(b.CREADO_EN).localeCompare(String(a.CREADO_EN)));
  },

  // ================= Ventas =================
  /** Registra el cierre: crea la venta, marca GANADO y crea el proyecto web. */
  registerSale(user, data) {
    data = data || {};
    const o = SalesService._requireOpp(user, String(data.OPORTUNIDAD_ID || ''));
    const existing = SheetService.findAllBy('VENTAS', 'OPORTUNIDAD_ID', o.OPORTUNIDAD_ID).filter(v => v.ESTADO_PAGO !== 'ANULADA');
    if (existing.length) throw new AppError('VALIDATION', 'Esta oportunidad ya tiene una venta registrada (' + existing[0].NUMERO + ').');
    let propuesta = null;
    if (data.PROPUESTA_ID) {
      propuesta = SheetService.getById('PROPUESTAS', data.PROPUESTA_ID);
      if (!propuesta || propuesta.OPORTUNIDAD_ID !== o.OPORTUNIDAD_ID) throw new AppError('VALIDATION', 'La propuesta no corresponde a esta oportunidad.');
    }
    const valor = Utils.parseMoney(data.VALOR_TOTAL) || (propuesta ? propuesta.TOTAL : 0) || Number(o.VALOR_ESTIMADO || 0);
    if (!valor) throw new AppError('VALIDATION', 'Escriba el valor de la venta.');
    const fecha = Utils.isIsoDate(data.FECHA_VENTA) ? data.FECHA_VENTA : Utils.today();
    const formaPago = ENUMS.FORMA_PAGO.indexOf(data.FORMA_PAGO) >= 0 ? data.FORMA_PAGO : 'CONTADO';
    const venta = {
      OPORTUNIDAD_ID: o.OPORTUNIDAD_ID, PROPUESTA_ID: propuesta ? propuesta.PROPUESTA_ID : '', CLIENTE_ID: o.CLIENTE_ID,
      NUMERO: 'VEN-' + Utils.year() + '-' + Seq.pad(Seq.next('VENTA_' + Utils.year()), 4),
      FECHA_VENTA: fecha, VALOR_TOTAL: valor, FORMA_PAGO: formaPago, ESTADO_PAGO: 'PENDIENTE', VALOR_PAGADO: 0,
      VENDEDOR_EMAIL: o.RESPONSABLE_EMAIL || user.email, NOTAS: Validation.long(data.NOTAS, 2000)
    };
    SheetService.insert('VENTAS', venta);
    Audit.log({ accion: 'VENTA', entidad: 'VENTAS', entidadId: venta.VENTA_ID, clienteId: o.CLIENTE_ID, despues: { NUMERO: venta.NUMERO, VALOR_TOTAL: valor, FORMA_PAGO: formaPago } });
    if (propuesta && propuesta.ESTADO !== 'ACEPTADA') {
      const r = SheetService.update('PROPUESTAS', propuesta.PROPUESTA_ID, { ESTADO: 'ACEPTADA', RESPUESTA_EN: Utils.nowIso() });
      Audit.logChange('CAMBIO_ESTADO', 'PROPUESTAS', propuesta.PROPUESTA_ID, o.CLIENTE_ID, r);
    }
    SheetService.update('OPORTUNIDADES', o.OPORTUNIDAD_ID, { VALOR_ESTIMADO: valor });
    CRMService.moveOpportunity(user, o.OPORTUNIDAD_ID, 'GANADO');
    // Proyecto web: siguiente eslabón de la cadena (BASE → … → VENTA → PROYECTO WEB).
    const reportes = SheetService.findAllBy('REPORTES', 'CLIENTE_ID', o.CLIENTE_ID).filter(r => r.ESTADO === 'GENERADO')
      .sort((a, b) => b.VERSION_REPORTE - a.VERSION_REPORTE);
    const proyecto = {
      VENTA_ID: venta.VENTA_ID, CLIENTE_ID: o.CLIENTE_ID, ESTADO: 'BRIEF', REPORTE_ID: reportes.length ? reportes[0].REPORTE_ID : '',
      DOMINIO: '', URL_PUBLICADA: '', FECHA_ENTREGA_ESTIMADA: Utils.addDays(fecha, 30), RESPONSABLE_EMAIL: venta.VENDEDOR_EMAIL, NOTAS: ''
    };
    SheetService.insert('PROYECTOS_WEB', proyecto);
    Audit.log({ accion: 'CREAR', entidad: 'PROYECTOS_WEB', entidadId: proyecto.PROYECTO_ID, clienteId: o.CLIENTE_ID, detalle: 'Por venta ' + venta.NUMERO });
    const c = ClientService.get(o.CLIENTE_ID);
    if (c && c.ESTADO_FORMULARIO === 'NO_INICIADO') {
      // El cliente puede no haber llenado el formulario: se crea el seguimiento para pedirle la información.
      CRMService.saveFollowup(user, {
        CLIENTE_ID: o.CLIENTE_ID, OPORTUNIDAD_ID: o.OPORTUNIDAD_ID, TIPO: 'WHATSAPP', PRIORIDAD: 'ALTA',
        DESCRIPCION: 'Enviar el enlace del formulario para recoger la información de la página.', FECHA_VENCIMIENTO: Utils.addDays(null, 1)
      });
    }
    if (Utils.parseMoney(data.PAGO_INICIAL) > 0) {
      SalesService.registerPayment(user, {
        VENTA_ID: venta.VENTA_ID, VALOR: data.PAGO_INICIAL, FECHA_PAGO: fecha, METODO: data.METODO || 'TRANSFERENCIA', REFERENCIA: data.REFERENCIA
      });
    }
    return { venta: ClientService._strip(SheetService.getById('VENTAS', venta.VENTA_ID)), proyecto: ClientService._strip(proyecto) };
  },

  updateSale(user, id, patch, expectedVersion) {
    const v = SheetService.getById('VENTAS', id);
    if (!v) throw new AppError('NOT_FOUND', 'Venta no encontrada.');
    ClientService.requireAccessible(user, v.CLIENTE_ID);
    const clean = Validation.cleanRecord('VENTAS', patch || {});
    const res = SheetService.update('VENTAS', id, clean, { expectedVersion: expectedVersion });
    Audit.logChange('EDITAR', 'VENTAS', id, v.CLIENTE_ID, res);
    SalesService._recalcSale(id);
    return ClientService._strip(SheetService.getById('VENTAS', id));
  },

  cancelSale(user, id, motivo) {
    const v = SheetService.getById('VENTAS', id);
    if (!v) throw new AppError('NOT_FOUND', 'Venta no encontrada.');
    ClientService.requireAccessible(user, v.CLIENTE_ID);
    const m = Validation.long(motivo, 1000);
    if (!m) throw new AppError('REASON_REQUIRED', 'Escriba el motivo de la anulación.');
    const res = SheetService.update('VENTAS', id, { ESTADO_PAGO: 'ANULADA', NOTAS: (v.NOTAS ? v.NOTAS + '\n' : '') + 'ANULADA: ' + m });
    Audit.logChange('ANULAR_VENTA', 'VENTAS', id, v.CLIENTE_ID, res);
    const o = SheetService.getById('OPORTUNIDADES', v.OPORTUNIDAD_ID);
    if (o && o.ETAPA === 'GANADO') CRMService.moveOpportunity(user, o.OPORTUNIDAD_ID, 'NEGOCIACION');
    SheetService.findAllBy('PROYECTOS_WEB', 'VENTA_ID', id).forEach(p => {
      const r = SheetService.update('PROYECTOS_WEB', p.PROYECTO_ID, { ESTADO: 'CANCELADO' });
      Audit.logChange('CAMBIO_ESTADO', 'PROYECTOS_WEB', p.PROYECTO_ID, v.CLIENTE_ID, r);
    });
    return ClientService._strip(res.record);
  },

  registerPayment(user, data) {
    data = data || {};
    const v = SheetService.getById('VENTAS', String(data.VENTA_ID || ''));
    if (!v) throw new AppError('NOT_FOUND', 'Venta no encontrada.');
    ClientService.requireAccessible(user, v.CLIENTE_ID);
    if (v.ESTADO_PAGO === 'ANULADA') throw new AppError('VALIDATION', 'La venta está anulada.');
    const valor = Utils.parseMoney(data.VALOR);
    if (!valor) throw new AppError('VALIDATION', 'Escriba el valor del pago.');
    const pago = {
      VENTA_ID: v.VENTA_ID, CLIENTE_ID: v.CLIENTE_ID,
      FECHA_PAGO: Utils.isIsoDate(data.FECHA_PAGO) ? data.FECHA_PAGO : Utils.today(),
      VALOR: valor, METODO: ENUMS.METODO_PAGO.indexOf(data.METODO) >= 0 ? data.METODO : 'OTRO',
      REFERENCIA: Validation.short(data.REFERENCIA, 120), COMPROBANTE_ARCHIVO_ID: '', NOTAS: Validation.long(data.NOTAS, 1000)
    };
    SheetService.insert('PAGOS', pago);
    Audit.log({ accion: 'PAGO', entidad: 'PAGOS', entidadId: pago.PAGO_ID, clienteId: v.CLIENTE_ID, despues: { VENTA: v.NUMERO, VALOR: valor, METODO: pago.METODO } });
    SalesService._recalcSale(v.VENTA_ID);
    return ClientService._strip(pago);
  },

  deletePayment(user, pagoId) {
    const p = SheetService.getById('PAGOS', pagoId);
    if (!p) throw new AppError('NOT_FOUND', 'Pago no encontrado.');
    ClientService.requireAccessible(user, p.CLIENTE_ID);
    SheetService.softDelete('PAGOS', pagoId);
    Audit.log({ accion: 'ELIMINAR', entidad: 'PAGOS', entidadId: pagoId, clienteId: p.CLIENTE_ID, antes: { VALOR: p.VALOR, FECHA: p.FECHA_PAGO } });
    SalesService._recalcSale(p.VENTA_ID);
    return true;
  },

  /** Recalcula VALOR_PAGADO y ESTADO_PAGO desde los pagos. */
  _recalcSale(ventaId) {
    const v = SheetService.getById('VENTAS', ventaId);
    if (!v || v.ESTADO_PAGO === 'ANULADA') return;
    const pagado = SheetService.findAllBy('PAGOS', 'VENTA_ID', ventaId).reduce((s, p) => s + Number(p.VALOR || 0), 0);
    const estado = pagado <= 0 ? 'PENDIENTE' : (pagado >= Number(v.VALOR_TOTAL || 0) ? 'PAGADA' : 'PARCIAL');
    const res = SheetService.update('VENTAS', ventaId, { VALOR_PAGADO: pagado, ESTADO_PAGO: estado });
    if (res.changed.length) Audit.logChange('RECALCULO_PAGOS', 'VENTAS', ventaId, v.CLIENTE_ID, res);
  },

  listSales(user, filters) {
    filters = filters || {};
    const clients = CRMService._clientMap();
    return SheetService.readAll('VENTAS').filter(v => {
      const c = clients[v.CLIENTE_ID];
      if (!c || !Auth.canAccessClient(user, c)) return false;
      if (filters.estadoPago && v.ESTADO_PAGO !== filters.estadoPago) return false;
      if (filters.desde && v.FECHA_VENTA < filters.desde) return false;
      if (filters.hasta && v.FECHA_VENTA > filters.hasta) return false;
      return true;
    }).map(v => {
      delete v._row;
      const c = clients[v.CLIENTE_ID];
      v.CLIENTE_NOMBRE = c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL;
      v.SALDO = Math.max(0, Number(v.VALOR_TOTAL) - Number(v.VALOR_PAGADO));
      return v;
    }).sort((a, b) => String(b.FECHA_VENTA).localeCompare(String(a.FECHA_VENTA)));
  },

  listPayments(user, filters) {
    filters = filters || {};
    const clients = CRMService._clientMap();
    const ventas = {};
    SheetService.readColumns('VENTAS', ['VENTA_ID', 'NUMERO']).forEach(v => { ventas[v.VENTA_ID] = v.NUMERO; });
    return SheetService.readAll('PAGOS').filter(p => {
      const c = clients[p.CLIENTE_ID];
      if (!c || !Auth.canAccessClient(user, c)) return false;
      if (filters.ventaId && p.VENTA_ID !== filters.ventaId) return false;
      if (filters.desde && p.FECHA_PAGO < filters.desde) return false;
      if (filters.hasta && p.FECHA_PAGO > filters.hasta) return false;
      return true;
    }).map(p => {
      delete p._row;
      const c = clients[p.CLIENTE_ID];
      p.CLIENTE_NOMBRE = c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL;
      p.VENTA_NUMERO = ventas[p.VENTA_ID] || '';
      return p;
    }).sort((a, b) => String(b.FECHA_PAGO).localeCompare(String(a.FECHA_PAGO)));
  },

  // ================= Proyectos web =================
  listProjects(user, filters) {
    filters = filters || {};
    const clients = CRMService._clientMap();
    return SheetService.readAll('PROYECTOS_WEB').filter(p => {
      const c = clients[p.CLIENTE_ID];
      if (!c || !Auth.canAccessClient(user, c)) return false;
      if (filters.estado && p.ESTADO !== filters.estado) return false;
      return true;
    }).map(p => {
      delete p._row;
      const c = clients[p.CLIENTE_ID];
      p.CLIENTE_NOMBRE = c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL;
      return p;
    }).sort((a, b) => String(a.FECHA_ENTREGA_ESTIMADA).localeCompare(String(b.FECHA_ENTREGA_ESTIMADA)));
  },

  updateProject(user, id, patch, expectedVersion) {
    const p = SheetService.getById('PROYECTOS_WEB', id);
    if (!p) throw new AppError('NOT_FOUND', 'Proyecto no encontrado.');
    ClientService.requireAccessible(user, p.CLIENTE_ID);
    const clean = Validation.cleanRecord('PROYECTOS_WEB', patch || {});
    if (clean.RESPONSABLE_EMAIL) CRMService._requireActiveUserEmail(clean.RESPONSABLE_EMAIL);
    if (clean.URL_PUBLICADA && Validation.url(clean.URL_PUBLICADA) === null) throw new AppError('VALIDATION', 'La URL debe empezar por https://');
    const res = SheetService.update('PROYECTOS_WEB', id, clean, { expectedVersion: expectedVersion });
    Audit.logChange(clean.ESTADO && clean.ESTADO !== p.ESTADO ? 'CAMBIO_ESTADO' : 'EDITAR', 'PROYECTOS_WEB', id, p.CLIENTE_ID, res);
    return ClientService._strip(res.record);
  }
};

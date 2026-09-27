/**
 * AuditService.gs — Historial inmutable (Parte J).
 * Las entradas se acumulan durante la ejecución y se escriben juntas al final
 * (Audit.flush) en el spreadsheet de auditoría, que solo recibe filas nuevas.
 */
const Audit = {
  _buffer: [],
  MAX_JSON: 8000,

  /**
   * @param {Object} e {accion, entidad, entidadId, clienteId, campos, antes, despues, resultado, detalle}
   */
  log(e) {
    const clip = (v) => {
      if (v === undefined || v === null || v === '') return '';
      const s = typeof v === 'string' ? v : JSON.stringify(Audit._redact(v));
      return s.length > Audit.MAX_JSON ? s.slice(0, Audit.MAX_JSON) + '…[recortado]' : s;
    };
    Audit._buffer.push({
      AUDIT_ID: Utils.newId('AUD'),
      FECHA_HORA: Utils.nowIso(),
      ACTOR: Ctx.actor(),
      ROL: Ctx.rol(),
      APP: Ctx.app(),
      ACCION: e.accion || '',
      ENTIDAD: e.entidad || '',
      ENTIDAD_ID: e.entidadId || '',
      CLIENTE_ID: e.clienteId || '',
      CAMPOS: Array.isArray(e.campos) ? e.campos.join(',') : (e.campos || ''),
      ANTES_JSON: clip(e.antes),
      DESPUES_JSON: clip(e.despues),
      RESULTADO: e.resultado || 'OK',
      DETALLE: Utils.truncate(e.detalle || '', 2000)
    });
  },

  /** Registra un cambio de update() con sus valores antes/después. */
  logChange(accion, table, id, clienteId, result, detalle) {
    if (!result || !result.changed || !result.changed.length) return;
    Audit.log({
      accion: accion, entidad: table, entidadId: id, clienteId: clienteId,
      campos: result.changed, antes: result.before, despues: result.after, detalle: detalle
    });
  },

  /** Quita secretos antes de guardar. */
  _redact(obj) {
    if (!Utils.isPlainObject(obj)) return obj;
    const out = {};
    Object.keys(obj).forEach(k => {
      if (/PASSWORD|TOKEN|SALT/i.test(k)) out[k] = '[protegido]';
      else if (k === '_row') return;
      else out[k] = obj[k];
    });
    return out;
  },

  flush() {
    if (!Audit._buffer.length) return;
    const entries = Audit._buffer;
    Audit._buffer = [];
    try {
      Lock.run(() => {
        const sh = SheetService.sheet('AUDITORIA');
        const rows = entries.map(o => SheetService.objToRow('AUDITORIA', o));
        const start = sh.getLastRow() + 1;
        SheetService._ensureCapacity(sh, start + rows.length - 1);
        const range = sh.getRange(start, 1, rows.length, rows[0].length);
        range.setNumberFormat('@');
        range.setValues(rows);
      });
    } catch (err) {
      // La auditoría nunca debe romper la operación del usuario, pero sí quedar en el log.
      console.error('AUDIT_FLUSH_FAILED', err && err.message, JSON.stringify(entries).slice(0, 4000));
    }
  },

  /** Consulta paginada (más recientes primero). */
  query(filters, page, pageSize) {
    filters = filters || {};
    pageSize = Math.min(Math.max(Number(pageSize) || 50, 10), 200);
    page = Math.max(Number(page) || 1, 1);
    const sh = SheetService.sheet('AUDITORIA');
    const last = sh.getLastRow();
    if (last < 2) return { items: [], total: 0, page: 1, pages: 1 };
    // Se leen como máximo las últimas 20.000 filas para mantener la respuesta rápida.
    const window = Math.min(last - 1, 20000);
    const startRow = last - window + 1;
    const width = SheetService.headers('AUDITORIA').__width;
    const values = sh.getRange(startRow, 1, window, width).getValues();
    const q = Utils.normalizeText(filters.q || '');
    const items = [];
    for (let i = values.length - 1; i >= 0; i--) {
      const o = SheetService.rowToObj('AUDITORIA', values[i], startRow + i);
      if (filters.clienteId && o.CLIENTE_ID !== filters.clienteId) continue;
      if (filters.entidad && o.ENTIDAD !== filters.entidad) continue;
      if (filters.accion && o.ACCION !== filters.accion) continue;
      if (filters.actor && o.ACTOR.indexOf(filters.actor) < 0) continue;
      if (filters.resultado && o.RESULTADO !== filters.resultado) continue;
      if (filters.desde && o.FECHA_HORA.slice(0, 10) < filters.desde) continue;
      if (filters.hasta && o.FECHA_HORA.slice(0, 10) > filters.hasta) continue;
      if (q && Utils.normalizeText([o.ENTIDAD_ID, o.CLIENTE_ID, o.DETALLE, o.CAMPOS].join(' ')).indexOf(q) < 0) continue;
      delete o._row;
      items.push(o);
    }
    const total = items.length;
    return {
      items: items.slice((page - 1) * pageSize, page * pageSize),
      total: total,
      page: page,
      pages: Math.max(1, Math.ceil(total / pageSize)),
      truncated: last - 1 > window
    };
  }
};

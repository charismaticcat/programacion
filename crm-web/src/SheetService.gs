/**
 * SheetService.gs — Repositorio genérico sobre Google Sheets.
 *
 * - Mapea columnas por el NOMBRE del encabezado (nunca por posición fija).
 * - Localiza filas por ID con TextFinder (coincidencia exacta).
 * - Nunca borra ni reordena filas: el borrado es lógico (ELIMINADO = SI).
 *   IMPORTANTE: no ordene ni filtre-y-borre filas manualmente en las hojas.
 * - Todas las escrituras van dentro de Lock.run() para evitar colisiones.
 */

/** Bloqueo de script reentrante dentro de la misma ejecución. */
const Lock = {
  _depth: 0,
  _lock: null,
  run(fn) {
    if (Lock._depth > 0) {
      Lock._depth++;
      try { return fn(); } finally { Lock._depth--; }
    }
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) {
      throw new AppError('BUSY', 'El sistema está atendiendo muchas solicitudes. Intente de nuevo en unos segundos.');
    }
    Lock._lock = lock;
    Lock._depth = 1;
    try {
      return fn();
    } finally {
      Lock._depth = 0;
      Lock._lock = null;
      try { SpreadsheetApp.flush(); } catch (e) { /* sin cambios pendientes */ }
      lock.releaseLock();
    }
  }
};

const SheetService = {
  _db: null,
  _audit: null,
  _sheets: {},
  _headers: {},

  /** Spreadsheet principal. */
  db() {
    if (!SheetService._db) SheetService._db = SpreadsheetApp.openById(Props.require('DB_SPREADSHEET_ID'));
    return SheetService._db;
  },
  auditDb() {
    if (!SheetService._audit) SheetService._audit = SpreadsheetApp.openById(Props.require('AUDIT_SPREADSHEET_ID'));
    return SheetService._audit;
  },
  sheet(table) {
    if (SheetService._sheets[table]) return SheetService._sheets[table];
    const t = SchemaUtil.table(table);
    const ss = t.auditDb ? SheetService.auditDb() : SheetService.db();
    const sh = ss.getSheetByName(table);
    if (!sh) throw new AppError('NOT_CONFIGURED', 'Falta la hoja ' + table + '. Ejecute setup().');
    SheetService._sheets[table] = sh;
    return sh;
  },
  /** Mapa NOMBRE_COLUMNA -> índice (base 1). */
  headers(table) {
    if (SheetService._headers[table]) return SheetService._headers[table];
    const sh = SheetService.sheet(table);
    const lastCol = sh.getLastColumn();
    const row = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const map = {};
    row.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
    SchemaUtil.columnNames(table).forEach(c => {
      if (!map[c]) throw new AppError('SCHEMA_MISMATCH', 'La hoja ' + table + ' no tiene la columna ' + c + '. Ejecute setup() para actualizarla.');
    });
    map.__width = lastCol;
    SheetService._headers[table] = map;
    return map;
  },
  resetCache() {
    SheetService._db = null;
    SheetService._audit = null;
    SheetService._sheets = {};
    SheetService._headers = {};
  },

  // ---------- Conversión fila <-> objeto ----------
  _fromCell(type, v) {
    if (v === null || v === undefined) v = '';
    switch (type) {
      case 'n': return v === '' ? 0 : Number(v) || 0;
      case 'm': return v === '' ? 0 : Math.round(Number(String(v).replace(/[^\d.-]/g, '')) || 0);
      case 'b': return Utils.toBool(v);
      case 'j': return Utils.safeJsonParse(v, null);
      default:
        if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
        return String(v);
    }
  },
  _toCell(type, v) {
    if (v === null || v === undefined) return '';
    switch (type) {
      case 'n':
      case 'm': return v === '' ? '' : String(Number(v) || 0);
      case 'b': return Utils.toBool(v) ? 'SI' : 'NO';
      case 'j': return typeof v === 'string' ? v : JSON.stringify(v);
      default: {
        const s = String(v);
        if (s.length > 49000) throw new AppError('TOO_LONG', 'Un texto supera el tamaño máximo permitido.');
        return Utils.sanitizeCell(s);
      }
    }
  },
  rowToObj(table, values, rowNumber) {
    const h = SheetService.headers(table);
    const obj = {};
    SchemaUtil.columns(table).forEach(c => {
      obj[c.name] = SheetService._fromCell(c.type, values[h[c.name] - 1]);
    });
    obj._row = rowNumber;
    return obj;
  },
  objToRow(table, obj) {
    const h = SheetService.headers(table);
    const row = new Array(h.__width).fill('');
    SchemaUtil.columns(table).forEach(c => {
      row[h[c.name] - 1] = SheetService._toCell(c.type, obj[c.name]);
    });
    return row;
  },

  // ---------- Lectura ----------
  lastRow(table) {
    return SheetService.sheet(table).getLastRow();
  },
  /** Número de fila del ID (o -1). */
  findRow(table, id) {
    if (!id) return -1;
    const t = SchemaUtil.table(table);
    const sh = SheetService.sheet(table);
    const last = sh.getLastRow();
    if (last < 2) return -1;
    const col = SheetService.headers(table)[t.key];
    const cell = sh.getRange(2, col, last - 1, 1)
      .createTextFinder(String(id)).matchEntireCell(true).matchCase(true).findNext();
    return cell ? cell.getRow() : -1;
  },
  getById(table, id, opts) {
    const row = SheetService.findRow(table, id);
    if (row < 0) return null;
    const sh = SheetService.sheet(table);
    const values = sh.getRange(row, 1, 1, SheetService.headers(table).__width).getValues()[0];
    const obj = SheetService.rowToObj(table, values, row);
    if (obj.ELIMINADO && !(opts && opts.includeDeleted)) return null;
    return obj;
  },
  /** Filas cuyo campo `col` es exactamente `value` (TextFinder, rápido en tablas grandes). */
  findAllBy(table, col, value, opts) {
    if (value === '' || value === null || value === undefined) return [];
    const sh = SheetService.sheet(table);
    const last = sh.getLastRow();
    if (last < 2) return [];
    const h = SheetService.headers(table);
    const cells = sh.getRange(2, h[col], last - 1, 1)
      .createTextFinder(String(value)).matchEntireCell(true).matchCase(true).findAll();
    const width = h.__width;
    const out = [];
    cells.forEach(cell => {
      const r = cell.getRow();
      const values = sh.getRange(r, 1, 1, width).getValues()[0];
      const obj = SheetService.rowToObj(table, values, r);
      if (!obj.ELIMINADO || (opts && opts.includeDeleted)) out.push(obj);
    });
    return out;
  },
  /** Toda la tabla (usar solo en tablas medianas; CLIENTES usa SearchIndex). */
  readAll(table, opts) {
    const sh = SheetService.sheet(table);
    const last = sh.getLastRow();
    if (last < 2) return [];
    const width = SheetService.headers(table).__width;
    const values = sh.getRange(2, 1, last - 1, width).getValues();
    const out = [];
    for (let i = 0; i < values.length; i++) {
      const obj = SheetService.rowToObj(table, values[i], i + 2);
      if (!obj.ELIMINADO || (opts && opts.includeDeleted)) out.push(obj);
    }
    return out;
  },
  /** Lee solo algunas columnas (eficiente en tablas grandes). */
  readColumns(table, cols, opts) {
    const sh = SheetService.sheet(table);
    const last = sh.getLastRow();
    if (last < 2) return [];
    const h = SheetService.headers(table);
    const need = cols.slice();
    if (need.indexOf('ELIMINADO') < 0 && h.ELIMINADO) need.push('ELIMINADO');
    const data = {};
    need.forEach(c => {
      if (!h[c]) throw new AppError('SCHEMA_MISMATCH', 'Columna desconocida ' + c);
      data[c] = sh.getRange(2, h[c], last - 1, 1).getValues();
    });
    const types = {};
    need.forEach(c => { types[c] = (SchemaUtil.column(table, c) || { type: 's' }).type; });
    const out = [];
    for (let i = 0; i < last - 1; i++) {
      const o = { _row: i + 2 };
      need.forEach(c => { o[c] = SheetService._fromCell(types[c], data[c][i][0]); });
      if (o.ELIMINADO && !(opts && opts.includeDeleted)) continue;
      out.push(o);
    }
    return out;
  },

  // ---------- Escritura ----------
  _ensureCapacity(sh, neededLastRow) {
    const max = sh.getMaxRows();
    if (neededLastRow > max) sh.insertRowsAfter(max, neededLastRow - max + 200);
  },
  _stampNew(table, obj, actor) {
    const t = SchemaUtil.table(table);
    if (t.key && !obj[t.key] && t.prefix) obj[t.key] = Utils.newId(t.prefix);
    if (t.control) {
      const now = Utils.nowIso();
      obj.CREADO_EN = obj.CREADO_EN || now;
      obj.CREADO_POR = obj.CREADO_POR || actor;
      obj.ACTUALIZADO_EN = now;
      obj.ACTUALIZADO_POR = actor;
      obj.VERSION = 1;
      obj.ELIMINADO = false;
    }
    return obj;
  },
  insert(table, obj) {
    return SheetService.insertMany(table, [obj])[0];
  },
  insertMany(table, objs) {
    if (!objs.length) return [];
    return Lock.run(() => {
      const actor = Ctx.actor();
      const sh = SheetService.sheet(table);
      const rows = objs.map(o => SheetService.objToRow(table, SheetService._stampNew(table, o, actor)));
      const start = sh.getLastRow() + 1;
      SheetService._ensureCapacity(sh, start + rows.length - 1);
      const range = sh.getRange(start, 1, rows.length, rows[0].length);
      range.setNumberFormat('@');
      range.setValues(rows);
      objs.forEach((o, i) => { o._row = start + i; });
      return objs;
    });
  },
  /**
   * Actualiza campos. expectedVersion (opcional) detecta ediciones simultáneas.
   * Devuelve {before, after, changed:[campos]}.
   */
  update(table, id, patch, opts) {
    opts = opts || {};
    return Lock.run(() => {
      const current = SheetService.getById(table, id, { includeDeleted: !!opts.includeDeleted });
      if (!current) throw new AppError('NOT_FOUND', 'Registro no encontrado (' + table + ').');
      const t = SchemaUtil.table(table);
      if (t.control && opts.expectedVersion !== undefined && opts.expectedVersion !== null &&
          Number(opts.expectedVersion) !== Number(current.VERSION)) {
        throw new AppError('CONFLICT', 'Otra persona modificó este registro. Recargue para ver la versión actual.');
      }
      const cols = SchemaUtil.columnNames(table);
      const changed = [];
      const next = Object.assign({}, current);
      Object.keys(patch).forEach(k => {
        if (cols.indexOf(k) < 0 || k === t.key) return;
        const colType = SchemaUtil.column(table, k).type;
        const a = SheetService._toCell(colType, current[k]);
        const b = SheetService._toCell(colType, patch[k]);
        if (a !== b) {
          changed.push(k);
          next[k] = patch[k];
        }
      });
      if (!changed.length) return { before: current, after: current, changed: [] };
      if (t.control) {
        next.ACTUALIZADO_EN = Utils.nowIso();
        next.ACTUALIZADO_POR = Ctx.actor();
        next.VERSION = Number(current.VERSION || 0) + 1;
      }
      const sh = SheetService.sheet(table);
      const row = SheetService.objToRow(table, next);
      const range = sh.getRange(current._row, 1, 1, row.length);
      range.setNumberFormat('@');
      range.setValues([row]);
      const before = {};
      const after = {};
      changed.forEach(k => { before[k] = current[k]; after[k] = next[k]; });
      next._row = current._row;
      return { before: before, after: after, changed: changed, record: next };
    });
  },
  softDelete(table, id, opts) {
    return SheetService.update(table, id, { ELIMINADO: true }, opts);
  },
  /** Cambia CLIENTE_ID de `fromId` a `toId` en una tabla (fusiones). Devuelve cuántas filas. */
  reassignClient(table, fromId, toId) {
    return Lock.run(() => {
      const sh = SheetService.sheet(table);
      const last = sh.getLastRow();
      if (last < 2) return 0;
      const h = SheetService.headers(table);
      const cells = sh.getRange(2, h.CLIENTE_ID, last - 1, 1)
        .createTextFinder(String(fromId)).matchEntireCell(true).matchCase(true).findAll();
      cells.forEach(c => c.setValue(toId));
      return cells.length;
    });
  }
};

/** Contadores consecutivos (radicados, números de propuesta y venta). */
const Seq = {
  next(name) {
    return Lock.run(() => {
      const sh = SheetService.sheet('SECUENCIAS');
      const row = SheetService.findRow('SECUENCIAS', name);
      const h = SheetService.headers('SECUENCIAS');
      if (row < 0) {
        const start = sh.getLastRow() + 1;
        SheetService._ensureCapacity(sh, start);
        const r = new Array(h.__width).fill('');
        r[h.NOMBRE - 1] = name;
        r[h.ULTIMO_VALOR - 1] = '1';
        sh.getRange(start, 1, 1, r.length).setNumberFormat('@').setValues([r]);
        return 1;
      }
      const cell = sh.getRange(row, h.ULTIMO_VALOR);
      const v = (Number(cell.getValue()) || 0) + 1;
      cell.setNumberFormat('@').setValue(String(v));
      return v;
    });
  },
  /** Fija el valor mínimo (usado tras importar la base). */
  ensureAtLeast(name, value) {
    return Lock.run(() => {
      const row = SheetService.findRow('SECUENCIAS', name);
      if (row < 0) {
        const sh = SheetService.sheet('SECUENCIAS');
        const h = SheetService.headers('SECUENCIAS');
        const start = sh.getLastRow() + 1;
        SheetService._ensureCapacity(sh, start);
        const r = new Array(h.__width).fill('');
        r[h.NOMBRE - 1] = name;
        r[h.ULTIMO_VALOR - 1] = String(value);
        sh.getRange(start, 1, 1, r.length).setNumberFormat('@').setValues([r]);
        return;
      }
      const h = SheetService.headers('SECUENCIAS');
      const cell = SheetService.sheet('SECUENCIAS').getRange(row, h.ULTIMO_VALOR);
      if ((Number(cell.getValue()) || 0) < value) cell.setNumberFormat('@').setValue(String(value));
    });
  },
  pad(n, len) {
    return ('000000000' + n).slice(-len);
  }
};

/** Contexto de la ejecución actual (quién actúa). */
const Ctx = {
  _actor: 'SISTEMA:interno',
  _rol: 'SISTEMA',
  _app: 'SISTEMA',
  _user: null,
  set(actor, rol, app, user) {
    Ctx._actor = actor;
    Ctx._rol = rol || '';
    Ctx._app = app || 'SISTEMA';
    Ctx._user = user || null;
  },
  actor() { return Ctx._actor; },
  rol() { return Ctx._rol; },
  app() { return Ctx._app; },
  user() { return Ctx._user; }
};

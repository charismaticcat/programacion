/**
 * Simulación en memoria de los servicios de Google Apps Script usados por el proyecto.
 * Permite ejecutar el código .gs real en Node para pruebas automáticas.
 * No pretende ser completa: cubre exactamente las APIs que usa src/.
 */
'use strict';
const crypto = require('crypto');
const zlib = require('zlib');

const toSigned = (buf) => Array.from(buf, b => (b > 127 ? b - 256 : b));
const toBuf = (v) => {
  if (Buffer.isBuffer(v)) return v;
  if (Array.isArray(v)) return Buffer.from(v.map(b => b & 0xff));
  return Buffer.from(String(v), 'utf8');
};

let idCounter = 0;
const newId = (p) => p + '_' + (++idCounter).toString(36) + crypto.randomBytes(6).toString('hex');

// ------------------------------------------------------------------ Blob
class Blob {
  constructor(data, contentType, name) {
    this._bytes = toBuf(data == null ? '' : data);
    this._type = contentType || 'application/octet-stream';
    this._name = name || '';
  }
  getBytes() { return toSigned(this._bytes); }
  getDataAsString() { return this._bytes.toString('utf8'); }
  getContentType() { return this._type; }
  setContentType(t) { this._type = t; return this; }
  getName() { return this._name; }
  setName(n) { this._name = n; return this; }
  copyBlob() { return new Blob(Buffer.from(this._bytes), this._type, this._name); }
}

// ------------------------------------------------------------------ Utilities
function pad(n, l) { return String(n).padStart(l, '0'); }
function formatDate(date, tz, pattern) {
  const offset = tz === 'UTC' || tz === 'GMT' ? 0 : -5;
  const d = new Date(date.getTime() + offset * 3600000);
  const map = {
    yyyy: d.getUTCFullYear(), MM: pad(d.getUTCMonth() + 1, 2), dd: pad(d.getUTCDate(), 2),
    HH: pad(d.getUTCHours(), 2), mm: pad(d.getUTCMinutes(), 2), ss: pad(d.getUTCSeconds(), 2)
  };
  let out = '';
  let i = 0;
  while (i < pattern.length) {
    if (pattern[i] === "'") {
      const j = pattern.indexOf("'", i + 1);
      out += pattern.slice(i + 1, j);
      i = j + 1;
      continue;
    }
    const tok = ['yyyy', 'MM', 'dd', 'HH', 'mm', 'ss'].find(t => pattern.startsWith(t, i));
    if (tok) { out += map[tok]; i += tok.length; } else { out += pattern[i]; i++; }
  }
  return out;
}
const Utilities = {
  DigestAlgorithm: { SHA_256: 'sha256', MD5: 'md5' },
  Charset: { UTF_8: 'utf8' },
  computeDigest(alg, value) {
    return toSigned(crypto.createHash(alg).update(toBuf(value)).digest());
  },
  computeHmacSha256Signature(value, key) {
    return toSigned(crypto.createHmac('sha256', toBuf(key)).update(toBuf(value)).digest());
  },
  base64Encode(v) { return toBuf(v).toString('base64'); },
  base64Decode(s) {
    if (/[^A-Za-z0-9+/=]/.test(s)) throw new Error('Could not decode string.');
    return toSigned(Buffer.from(s, 'base64'));
  },
  getUuid() { return crypto.randomUUID(); },
  formatDate: formatDate,
  newBlob(data, contentType, name) { return new Blob(data, contentType, name); },
  gzip(blob) { return new Blob(zlib.gzipSync(blob._bytes), 'application/x-gzip', blob.getName() + '.gz'); },
  ungzip(blob) { return new Blob(zlib.gunzipSync(blob._bytes), 'application/octet-stream'); },
  sleep() {}
};

// ------------------------------------------------------------------ Spreadsheets
class Range {
  constructor(sheet, row, col, nr, nc) {
    if (row < 1 || col < 1 || nr < 1 || nc < 1) throw new Error('Rango inválido ' + [row, col, nr, nc]);
    if (row + nr - 1 > sheet._maxRows || col + nc - 1 > sheet._maxCols) {
      throw new Error('Those rows/columns are out of bounds: ' + [row, col, nr, nc, sheet._maxRows, sheet._maxCols].join(','));
    }
    this.s = sheet; this.r = row; this.c = col; this.nr = nr; this.nc = nc;
  }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(this.s._get(this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  }
  getValue() { return this.s._get(this.r, this.c); }
  setValues(vals) {
    if (vals.length !== this.nr) throw new Error('setValues: filas ' + vals.length + ' != ' + this.nr);
    vals.forEach((row, i) => {
      if (row.length !== this.nc) throw new Error('setValues: columnas ' + row.length + ' != ' + this.nc);
      row.forEach((v, j) => this.s._set(this.r + i, this.c + j, v));
    });
    return this;
  }
  setValue(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.s._set(this.r + i, this.c + j, v); return this; }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  createTextFinder(text) { return new TextFinder(this, String(text)); }
}
class TextFinder {
  constructor(range, text) { this.range = range; this.text = text; this.entire = false; this.caseSens = false; this._pos = 0; }
  matchEntireCell(b) { this.entire = b; return this; }
  matchCase(b) { this.caseSens = b; return this; }
  _matches() {
    const out = [];
    const r = this.range;
    for (let i = 0; i < r.nr; i++) {
      for (let j = 0; j < r.nc; j++) {
        let v = String(r.s._get(r.r + i, r.c + j));
        let t = this.text;
        if (!this.caseSens) { v = v.toLowerCase(); t = t.toLowerCase(); }
        if (this.entire ? v === t : v.indexOf(t) >= 0) out.push(new Range(r.s, r.r + i, r.c + j, 1, 1));
      }
    }
    return out;
  }
  findNext() { const m = this._matches(); return m[this._pos++] || null; }
  findAll() { return this._matches(); }
}
class Sheet {
  constructor(ss, name) {
    this.ss = ss; this.name = name; this.rows = []; this._maxRows = 1000; this._maxCols = 26; this.frozen = 0;
  }
  _get(r, c) { const row = this.rows[r - 1]; const v = row ? row[c - 1] : undefined; return v === undefined ? '' : v; }
  _set(r, c, v) {
    if (typeof v === 'string' && v.charAt(0) === '=') throw new Error('Fórmula escrita sin sanear: ' + v);
    if (typeof v === 'string' && v.charAt(0) === "'") v = v.slice(1); // prefijo de texto de Sheets: se guarda como texto
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v === null || v === undefined ? '' : v;
    this.ss._writes++;
  }
  getName() { return this.name; }
  getLastRow() {
    for (let i = this.rows.length - 1; i >= 0; i--) if (this.rows[i].some(v => v !== '' && v !== undefined)) return i + 1;
    return 0;
  }
  getLastColumn() {
    let m = 0;
    this.rows.forEach(r => { for (let j = r.length - 1; j >= 0; j--) if (r[j] !== '') { m = Math.max(m, j + 1); break; } });
    return m;
  }
  getMaxRows() { return this._maxRows; }
  getMaxColumns() { return this._maxCols; }
  insertRowsAfter(after, n) { this._maxRows += n; return this; }
  insertColumnsAfter(after, n) { this._maxCols += n; return this; }
  getRange(row, col, nr, nc) { return new Range(this, row, col, nr || 1, nc || 1); }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  setFrozenRows(n) { this.frozen = n; return this; }
}
class Spreadsheet {
  constructor(name, id) { this.name = name; this.id = id || newId('ss'); this.sheets = [new Sheet(this, 'Hoja 1')]; this._writes = 0; }
  getId() { return this.id; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id; }
  getName() { return this.name; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  getSheets() { return this.sheets.slice(); }
  insertSheet(n) { const s = new Sheet(this, n); this.sheets.push(s); return s; }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
  setSpreadsheetTimeZone() {}
}

// ------------------------------------------------------------------ Drive
class DFile {
  constructor(drive, blob, parent) {
    this.drive = drive; this.id = newId('file'); this.blob = blob; this.name = blob.getName(); this.parent = parent;
    this.trashed = false; this.updated = new Date(); this.editors = []; this.access = 'PRIVATE';
  }
  getId() { return this.id; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getBlob() { return this.blob.copyBlob(); }
  getAs(mime) { return new Blob(Buffer.from('%PDF-1.4 mock ' + this.name), mime, this.name + '.pdf'); }
  setContent(s) { this.blob = new Blob(s, this.blob.getContentType(), this.name); this.updated = new Date(); return this; }
  setTrashed(b) { this.trashed = b; return this; }
  isTrashed() { return this.trashed; }
  moveTo(folder) { this.parent = folder; return this; }
  getLastUpdated() { return this.updated; }
  getThumbnail() { return null; }
  addEditor(e) { this.editors.push(e); return this; }
  getSharingAccess() { return this.access; }
  getUrl() { return 'https://drive.google.com/file/d/' + this.id + '/view'; }
  getSize() { return this.blob._bytes.length; }
}
class DFolder {
  constructor(drive, name, parent) { this.drive = drive; this.id = newId('folder'); this.name = name; this.parent = parent; this.trashed = false; this.access = 'PRIVATE'; }
  getId() { return this.id; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  isTrashed() { return this.trashed; }
  createFolder(n) { const f = new DFolder(this.drive, n, this); this.drive.folders[f.id] = f; return f; }
  getFoldersByName(n) {
    const list = Object.values(this.drive.folders).filter(f => f.parent === this && f.name === n && !f.trashed);
    let i = 0;
    return { hasNext: () => i < list.length, next: () => list[i++] };
  }
  createFile(blob) { const f = new DFile(this.drive, blob, this); this.drive.files[f.id] = f; return f; }
  getSharingAccess() { return this.access; }
  children() { return Object.values(this.drive.folders).filter(f => f.parent === this); }
  files() { return Object.values(this.drive.files).filter(f => f.parent === this); }
}

// ------------------------------------------------------------------ Documents
class DText {
  setBold() { return this; } setItalic() { return this; } setFontSize() { return this; }
}
class DImage {
  constructor() { this.w = 800; this.h = 600; }
  getWidth() { return this.w; } getHeight() { return this.h; }
  setWidth(w) { this.w = w; return this; } setHeight(h) { this.h = h; return this; }
}
class DParagraph {
  constructor(body, text) { this.body = body; this.text = text; this.heading = null; this.link = null; this.images = 0; }
  setHeading(h) { this.heading = h; return this; }
  editAsText() { return new DText(); }
  setLinkUrl(u) { this.link = u; return this; }
  appendInlineImage() { this.images++; this.body.images++; return new DImage(); }
}
class DBody {
  constructor() { this.paragraphs = []; this.images = 0; }
  clear() { this.paragraphs = []; return this; }
  setMarginTop() { return this; } setMarginBottom() { return this; } setMarginLeft() { return this; } setMarginRight() { return this; }
  appendParagraph(t) { const p = new DParagraph(this, t); this.paragraphs.push(p); return p; }
  getText() { return this.paragraphs.map(p => p.text).join('\n'); }
}

// ------------------------------------------------------------------ Entorno completo
function createEnvironment(options) {
  options = options || {};
  const state = {
    spreadsheets: {}, drive: { folders: {}, files: {} }, cache: new Map(), props: {}, triggers: [], mails: [], docs: {},
    activeUser: options.activeUser || '', effectiveUser: options.effectiveUser || 'owner@example.com', logs: []
  };
  const driveRoot = new DFolder(state.drive, 'Mi unidad', null);
  state.drive.folders[driveRoot.id] = driveRoot;

  const SpreadsheetApp = {
    create(name) { const ss = new Spreadsheet(name); state.spreadsheets[ss.id] = ss; const f = driveRoot.createFile(new Blob('', 'application/vnd.google-apps.spreadsheet', name)); f.id = ss.id; state.drive.files[ss.id] = f; return ss; },
    openById(id) { const ss = state.spreadsheets[id]; if (!ss) throw new Error('No se encontró el spreadsheet ' + id); return ss; },
    flush() {}
  };
  const DriveApp = {
    Access: { ANYONE: 'ANYONE', ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', PRIVATE: 'PRIVATE' },
    createFolder(name) { return driveRoot.createFolder(name); },
    getFolderById(id) { const f = state.drive.folders[id]; if (!f) throw new Error('No item with the given ID could be found: ' + id); return f; },
    getFileById(id) { const f = state.drive.files[id]; if (!f) throw new Error('No item with the given ID could be found: ' + id); return f; },
    getRootFolder() { return driveRoot; }
  };
  const cacheObj = {
    get(k) {
      const e = state.cache.get(k);
      if (!e) return null;
      if (e.exp < Date.now()) { state.cache.delete(k); return null; }
      return e.v;
    },
    put(k, v, ttl) {
      if (typeof v !== 'string') throw new Error('Cache: valor no string');
      if (k.length > 250) throw new Error('Cache: clave demasiado larga');
      if (Buffer.byteLength(v) > 100 * 1024) throw new Error('Argument too large: value (' + k + ')');
      state.cache.set(k, { v: v, exp: Date.now() + (ttl || 600) * 1000 });
    },
    getAll(keys) { const o = {}; keys.forEach(k => { const v = cacheObj.get(k); if (v !== null) o[k] = v; }); return o; },
    putAll(map, ttl) { Object.keys(map).forEach(k => cacheObj.put(k, map[k], ttl)); },
    remove(k) { state.cache.delete(k); },
    removeAll(keys) { keys.forEach(k => state.cache.delete(k)); }
  };
  const CacheService = { getScriptCache: () => cacheObj };
  const propsObj = {
    getProperty: (k) => (Object.prototype.hasOwnProperty.call(state.props, k) ? state.props[k] : null),
    setProperty: (k, v) => {
      if (String(v).length > 9000) throw new Error('Propiedad demasiado grande: ' + k);
      state.props[k] = String(v);
      return propsObj;
    },
    deleteProperty: (k) => { delete state.props[k]; return propsObj; },
    getProperties: () => Object.assign({}, state.props)
  };
  const PropertiesService = { getScriptProperties: () => propsObj };
  const LockService = { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {}, hasLock: () => true }) };
  const Session = {
    getActiveUser: () => ({ getEmail: () => state.activeUser }),
    getEffectiveUser: () => ({ getEmail: () => state.effectiveUser })
  };
  const makeTrigger = (fn) => {
    const t = { fn: fn, uid: 'trg_' + crypto.randomBytes(4).toString('hex') };
    const api = {
      timeBased: () => api, everyMinutes: () => api, everyHours: () => api, everyDays: () => api, atHour: () => api, after: () => api,
      create: () => { state.triggers.push(t); return { getUniqueId: () => t.uid, getHandlerFunction: () => fn }; }
    };
    return api;
  };
  const ScriptApp = {
    newTrigger: makeTrigger,
    getProjectTriggers: () => state.triggers.map(t => ({ getUniqueId: () => t.uid, getHandlerFunction: () => t.fn, _t: t })),
    deleteTrigger: (tr) => { state.triggers = state.triggers.filter(t => t !== tr._t); },
    getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' })
  };
  const MailApp = { sendEmail: (o) => { state.mails.push(o); }, getRemainingDailyQuota: () => 100 };
  const DocumentApp = {
    ParagraphHeading: { TITLE: 'TITLE', SUBTITLE: 'SUBTITLE', HEADING1: 'H1', HEADING2: 'H2', NORMAL: 'N' },
    create(name) {
      const body = new DBody();
      const file = driveRoot.createFile(new Blob('', 'application/vnd.google-apps.document', name));
      let closed = false;
      const guard = (fn) => () => { if (closed) throw new Error('Document is closed, its contents cannot be updated.'); return fn(); };
      const doc = {
        getId: guard(() => file.id), getUrl: guard(() => 'https://docs.google.com/document/d/' + file.id + '/edit'),
        getBody: guard(() => body), saveAndClose: guard(() => { closed = true; file.updated = new Date(Date.now() - 1000); }), _body: body, _name: name
      };
      state.docs[file.id] = doc;
      return doc;
    }
  };
  const Logger = { log: (m) => state.logs.push(String(m)) };
  const consoleProxy = {
    log: (...a) => state.logs.push(a.join(' ')),
    warn: (...a) => state.logs.push('WARN ' + a.join(' ')),
    error: (...a) => { state.logs.push('ERROR ' + a.join(' ')); if (options.echoErrors) process.stderr.write(a.join(' ') + '\n'); },
    info: (...a) => state.logs.push(a.join(' '))
  };
  return {
    state: state,
    globals: {
      Utilities, SpreadsheetApp, DriveApp, CacheService, PropertiesService, LockService, Session, ScriptApp,
      MailApp, DocumentApp, Logger, console: consoleProxy
    },
    helpers: { Blob, Spreadsheet, formatDate, toSigned }
  };
}

module.exports = { createEnvironment, Blob, formatDate };

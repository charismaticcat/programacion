/**
 * Utils.gs — Utilidades puras (sin acceso a Sheets/Drive).
 */

/** Error controlado: su mensaje se puede mostrar al usuario. */
class AppError extends Error {
  constructor(code, message, details) {
    super(message);
    this.name = 'AppError';
    this.code = code;
    this.details = details || null;
  }
}

const Utils = {
  // ---------- Fechas ----------
  nowIso() {
    return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  },
  today() {
    return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  },
  year() {
    return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy');
  },
  /** Suma días a una fecha yyyy-MM-dd (o a hoy) y devuelve yyyy-MM-dd. */
  addDays(isoDate, days) {
    const base = isoDate ? Utils.parseIsoDate(isoDate) : new Date();
    const d = new Date(base.getTime() + days * 86400000);
    return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  },
  /** Convierte 'yyyy-MM-dd' o 'yyyy-MM-ddTHH:mm:ss' (hora Bogotá, UTC-5) a Date. */
  parseIsoDate(s) {
    const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    if (!m) return null;
    // Colombia no tiene horario de verano: UTC-5 fijo.
    return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], (+(m[4] || 0)) + 5, +(m[5] || 0), +(m[6] || 0)));
  },
  isIsoDate(s) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) && Utils.parseIsoDate(s) !== null;
  },
  isIsoDateTime(s) {
    return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/.test(String(s || ''));
  },
  /** 20260303 -> '2026-03-03' */
  yyyymmddToIso(v) {
    const s = String(v == null ? '' : v).replace(/\D/g, '');
    if (s.length !== 8) return '';
    return s.slice(0, 4) + '-' + s.slice(4, 6) + '-' + s.slice(6, 8);
  },
  minutesBetween(isoA, isoB) {
    const a = Utils.parseIsoDate(isoA);
    const b = Utils.parseIsoDate(isoB);
    if (!a || !b) return 0;
    return (b.getTime() - a.getTime()) / 60000;
  },

  // ---------- IDs y criptografía ----------
  /** ID no secuencial: PREFIJO-<timestamp base36><6 hex aleatorios>. */
  newId(prefix) {
    const t = Date.now().toString(36).toUpperCase();
    const r = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
    return prefix + '-' + t + r;
  },
  /** Token aleatorio de 64 hex (2 UUID v4 = 244 bits de azar, pasados por SHA-256). */
  randomToken() {
    return Utils.sha256Hex(Utilities.getUuid() + Utilities.getUuid() + Date.now());
  },
  bytesToHex(bytes) {
    return bytes.map(b => ('0' + ((b < 0 ? b + 256 : b)).toString(16)).slice(-2)).join('');
  },
  sha256Hex(input) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
    return Utils.bytesToHex(bytes);
  },
  sha256HexBytes(bytes) {
    return Utils.bytesToHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes));
  },
  /** Comparación en tiempo constante de dos strings. */
  safeEqual(a, b) {
    a = String(a || '');
    b = String(b || '');
    let diff = a.length ^ b.length;
    const n = Math.max(a.length, b.length);
    for (let i = 0; i < n; i++) {
      diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
    }
    return diff === 0;
  },

  // ---------- Texto ----------
  cleanSpaces(s) {
    return String(s == null ? '' : s).replace(/[   ]/g, ' ').replace(/\s+/g, ' ').trim();
  },
  /** Mayúsculas, sin tildes ni signos, espacios simples. */
  normalizeText(s) {
    return String(s == null ? '' : s)
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase()
      .replace(/[^A-Z0-9]+/g, ' ')
      .trim();
  },
  /** Quita sufijos societarios para comparar nombres. */
  stripLegalSuffixes(normalized) {
    return String(normalized || '')
      .replace(/\b(S A S|SAS|S A|SA|LTDA|LIMITADA|E U|EU|S EN C|SENC|Y CIA|CIA|E S P|ESP|EN LIQUIDACION|SOCIEDAD|COMPANIA)\b/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  },
  slug(s, max) {
    const out = Utils.normalizeText(s).replace(/\s+/g, '-');
    return (out || 'SIN-NOMBRE').slice(0, max || 40).replace(/-+$/, '');
  },
  truncate(s, n) {
    s = String(s == null ? '' : s);
    return s.length > n ? s.slice(0, n) : s;
  },
  wordCount(s) {
    const t = Utils.cleanSpaces(s);
    return t ? t.split(' ').length : 0;
  },
  /** '41001 - NEIVA' -> {code:'41001', name:'NEIVA'} */
  splitCodeName(v) {
    const s = Utils.cleanSpaces(v);
    const m = s.match(/^(\d+)\s*-\s*(.*)$/);
    if (m) return { code: m[1], name: Utils.cleanSpaces(m[2]) };
    return { code: '', name: s };
  },
  /** 'G4721 ** Comercio al por menor…' -> {code:'G4721', desc:'Comercio al por menor…'} */
  parseCiiu(v) {
    const s = Utils.cleanSpaces(v);
    const m = s.match(/^([A-Z]\d{4})\s*\*\*\s*(.*)$/);
    if (m) return { code: m[1], desc: Utils.cleanSpaces(m[2]) };
    return { code: '', desc: s };
  },

  // ---------- Teléfonos y emails ----------
  digits(s) {
    return String(s == null ? '' : s).replace(/\D/g, '');
  },
  /**
   * Regla D7: 7 dígitos -> '608' + número; 10 dígitos -> igual; otro -> vacío.
   * Devuelve {value, note}.
   */
  normalizePhone(raw) {
    const d = Utils.digits(raw);
    if (d.length === 10) return { value: d, note: '' };
    if (d.length === 7) return { value: '608' + d, note: 'TEL_608' };
    if (d.length === 12 && d.indexOf('57') === 0) return { value: d.slice(2), note: '' };
    if (!d) return { value: '', note: '' };
    return { value: '', note: 'TEL_INVALIDO' };
  },
  normalizeEmail(raw) {
    return Utils.cleanSpaces(raw).toLowerCase();
  },
  isValidEmail(e) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(e || ''));
  },
  maskPhone(p) {
    const d = Utils.digits(p);
    if (d.length < 6) return d ? '***' : '';
    return d.slice(0, 3) + '****' + d.slice(-3);
  },
  maskEmail(e) {
    const s = String(e || '');
    const at = s.indexOf('@');
    if (at < 1) return '';
    return s.slice(0, Math.min(2, at)) + '*****' + s.slice(at);
  },

  // ---------- Números y dinero ----------
  toNumber(v, def) {
    if (v === '' || v === null || v === undefined) return def === undefined ? 0 : def;
    const n = Number(String(v).replace(/[^\d.-]/g, ''));
    return isFinite(n) ? n : (def === undefined ? 0 : def);
  },
  /** '$ 25.000' / '25000' -> 25000 (entero COP). */
  parseMoney(v) {
    const d = String(v == null ? '' : v).replace(/[^\d]/g, '');
    return d ? parseInt(d, 10) : 0;
  },
  formatMoney(n) {
    const v = Math.round(Number(n) || 0);
    const s = String(Math.abs(v)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return (v < 0 ? '-$ ' : '$ ') + s;
  },
  /** Precio libre (D14): si es solo número se formatea, si es texto se respeta. */
  formatPriceText(v) {
    const s = Utils.cleanSpaces(v);
    if (!s) return '';
    if (/^\$?\s*[\d.,]+$/.test(s)) return Utils.formatMoney(Utils.parseMoney(s));
    return s;
  },
  toBool(v) {
    if (v === true) return true;
    const s = String(v == null ? '' : v).trim().toUpperCase();
    return s === 'TRUE' || s === 'SI' || s === 'SÍ' || s === '1' || s === 'YES';
  },

  // ---------- Similitud de nombres ----------
  trigrams(s) {
    const t = '  ' + s + ' ';
    const set = {};
    for (let i = 0; i < t.length - 2; i++) set[t.slice(i, i + 3)] = true;
    return set;
  },
  /** Coeficiente de Dice entre dos textos normalizados (0..1). */
  similarity(a, b) {
    a = Utils.stripLegalSuffixes(Utils.normalizeText(a));
    b = Utils.stripLegalSuffixes(Utils.normalizeText(b));
    if (!a || !b) return 0;
    if (a === b) return 1;
    const ta = Utils.trigrams(a);
    const tb = Utils.trigrams(b);
    const ka = Object.keys(ta);
    const kb = Object.keys(tb);
    let inter = 0;
    ka.forEach(k => { if (tb[k]) inter++; });
    return (2 * inter) / (ka.length + kb.length);
  },

  // ---------- JSON / objetos ----------
  safeJsonParse(s, def) {
    if (s === null || s === undefined || s === '') return def;
    try { return JSON.parse(s); } catch (e) { return def; }
  },
  pick(obj, keys) {
    const out = {};
    keys.forEach(k => { if (obj && Object.prototype.hasOwnProperty.call(obj, k)) out[k] = obj[k]; });
    return out;
  },
  isPlainObject(o) {
    return o !== null && typeof o === 'object' && !Array.isArray(o);
  },

  // ---------- Hojas / CSV ----------
  /** Evita inyección de fórmulas: el apóstrofo inicial fuerza texto en Sheets y no se muestra. */
  sanitizeCell(v) {
    if (typeof v !== 'string') return v;
    return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
  },
  csvEscape(v) {
    let s = String(v == null ? '' : v);
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  },
  chunkString(s, size) {
    const out = [];
    for (let i = 0; i < s.length; i += size) out.push(s.slice(i, i + size));
    return out;
  }
};

/**
 * SearchIndex.gs — Índice compacto de CLIENTES en CacheService.
 *
 * Leer 31.000 filas × 38 columnas en cada búsqueda sería lento. El índice guarda
 * solo las columnas necesarias para buscar, filtrar y listar, comprimido (gzip)
 * y partido en trozos de 90 KB. Se reconstruye si falta y se actualiza fila a
 * fila en cada creación o edición de clientes.
 */
const SearchIndex = {
  FIELDS: ['CLIENTE_ID', 'RAZON_SOCIAL', 'NOMBRE_COMERCIAL', 'NOMBRE_PROPIETARIO', 'TELEFONO', 'EMAIL',
    'MUNICIPIO', 'BARRIO', 'ESTADO_CRM', 'ESTADO_FORMULARIO', 'RESPONSABLE_EMAIL', 'ORIGEN', 'CIIU_CODIGO',
    'TAMANO_EMPRESA', 'ULT_ANO_RENOVACION', 'PRIORIDAD', 'NO_CONTACTAR', 'TELEFONO_COMPARTIDO',
    'FUSIONADO_EN', 'ACTUALIZADO_EN', 'TIPO_ESTABLECIMIENTO'],
  _mem: null,

  _encode(rows) {
    const json = JSON.stringify({ f: SearchIndex.FIELDS, r: rows });
    const gz = Utilities.gzip(Utilities.newBlob(json, 'application/json'));
    return Utilities.base64Encode(gz.getBytes());
  },
  _decode(b64) {
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip');
    const json = Utilities.ungzip(blob).getDataAsString('UTF-8');
    const obj = JSON.parse(json);
    if (JSON.stringify(obj.f) !== JSON.stringify(SearchIndex.FIELDS)) return null;
    return obj.r;
  },

  build() {
    const rows = SheetService.readColumns('CLIENTES', SearchIndex.FIELDS);
    const out = rows.map(o => SearchIndex.FIELDS.map(f => {
      const v = o[f];
      return typeof v === 'boolean' ? (v ? 1 : 0) : v;
    }));
    SearchIndex._save(out);
    return out;
  },

  _save(rows) {
    SearchIndex._mem = rows;
    const b64 = SearchIndex._encode(rows);
    const chunks = Utils.chunkString(b64, CONFIG.INDEX_CHUNK_CHARS);
    const version = String(Date.now());
    const cache = CacheService.getScriptCache();
    const map = {};
    chunks.forEach((c, i) => { map['idx:' + version + ':' + i] = c; });
    map['idx:meta'] = JSON.stringify({ v: version, n: chunks.length, count: rows.length });
    cache.putAll(map, CONFIG.INDEX_TTL_SEC);
  },

  load() {
    if (SearchIndex._mem) return SearchIndex._mem;
    const cache = CacheService.getScriptCache();
    const meta = Utils.safeJsonParse(cache.get('idx:meta'), null);
    if (meta) {
      const keys = [];
      for (let i = 0; i < meta.n; i++) keys.push('idx:' + meta.v + ':' + i);
      const got = cache.getAll(keys);
      if (keys.every(k => got[k])) {
        try {
          const rows = SearchIndex._decode(keys.map(k => got[k]).join(''));
          if (rows) {
            SearchIndex._mem = rows;
            return rows;
          }
        } catch (e) {
          console.warn('Índice dañado, se reconstruye', e && e.message);
        }
      }
    }
    return Lock.run(() => SearchIndex.build());
  },

  /** Índice como objetos. */
  records() {
    const F = SearchIndex.FIELDS;
    return SearchIndex.load().map(r => {
      const o = {};
      F.forEach((f, i) => { o[f] = r[i]; });
      o.NO_CONTACTAR = !!o.NO_CONTACTAR;
      o.TELEFONO_COMPARTIDO = !!o.TELEFONO_COMPARTIDO;
      return o;
    });
  },

  /** Inserta o reemplaza la entrada de un cliente (llamar tras crear/editar). */
  upsert(cliente) {
    if (!cliente || !cliente.CLIENTE_ID) return;
    Lock.run(() => {
      // Se relee dentro del bloqueo para no pisar cambios de otra ejecución.
      SearchIndex._mem = null;
      const rows = SearchIndex.load().slice();
      const entry = SearchIndex.FIELDS.map(f => {
        const v = cliente[f];
        return typeof v === 'boolean' ? (v ? 1 : 0) : (v === undefined || v === null ? '' : v);
      });
      const i = rows.findIndex(r => r[0] === cliente.CLIENTE_ID);
      if (cliente.ELIMINADO) {
        if (i >= 0) rows.splice(i, 1);
      } else if (i >= 0) {
        rows[i] = entry;
      } else {
        rows.push(entry);
      }
      SearchIndex._save(rows);
    });
  },

  /** Reconstruye el índice si no está en caché (lo llama el activador cada 10 min). */
  warm() {
    if (CacheService.getScriptCache().get('idx:meta')) return false;
    Lock.run(() => SearchIndex.build());
    return true;
  },

  invalidate() {
    SearchIndex._mem = null;
    CacheService.getScriptCache().remove('idx:meta');
  },

  /** Texto normalizado para buscar por nombre. */
  nameKey(o) {
    return Utils.normalizeText([o.RAZON_SOCIAL, o.NOMBRE_COMERCIAL, o.NOMBRE_PROPIETARIO].join(' '));
  },

  /**
   * Búsqueda del formulario público (Parte F).
   * type: 'telefono' | 'email' | 'nombre'. Solo coincidencias exactas para teléfono/email.
   */
  publicSearch(type, value, municipio) {
    const recs = SearchIndex.records().filter(o => !o.FUSIONADO_EN);
    if (type === 'telefono') {
      const p = Utils.normalizePhone(value).value;
      if (!p) throw new AppError('VALIDATION', 'Escriba un teléfono de 7 o 10 dígitos.');
      return recs.filter(o => o.TELEFONO === p);
    }
    if (type === 'email') {
      const e = Utils.normalizeEmail(value);
      if (!Utils.isValidEmail(e)) throw new AppError('VALIDATION', 'Escriba un email válido.');
      return recs.filter(o => o.EMAIL === e);
    }
    if (type === 'nombre') {
      const q = Utils.normalizeText(value);
      if (q.replace(/\s/g, '').length < CONFIG.SEARCH_MIN_NAME_CHARS) {
        throw new AppError('VALIDATION', 'Escriba al menos ' + CONFIG.SEARCH_MIN_NAME_CHARS + ' letras del nombre.');
      }
      const tokens = q.split(' ').filter(t => t.length >= 2);
      const mun = municipio ? Validation.municipioName(municipio) : '';
      return recs.filter(o => {
        if (mun && o.MUNICIPIO !== mun) return false;
        const key = SearchIndex.nameKey(o);
        return tokens.every(t => key.indexOf(t) >= 0);
      });
    }
    throw new AppError('VALIDATION', 'Tipo de búsqueda no válido.');
  },

  /** Listado del admin con filtros, orden y paginación. */
  adminQuery(user, params) {
    params = params || {};
    const f = params.filters || {};
    let recs = SearchIndex.records();
    if (!f.incluirFusionados) recs = recs.filter(o => !o.FUSIONADO_EN);
    if (Auth.isScopedToOwn(user)) recs = recs.filter(o => String(o.RESPONSABLE_EMAIL).toLowerCase() === user.email);

    const q = Utils.normalizeText(params.q || '');
    const qDigits = Utils.digits(params.q || '');
    const qEmail = String(params.q || '').trim().toLowerCase();
    if (q) {
      const tokens = q.split(' ');
      recs = recs.filter(o => {
        if (o.CLIENTE_ID === String(params.q).trim().toUpperCase()) return true;
        if (qDigits.length >= 5 && String(o.TELEFONO).indexOf(qDigits) >= 0) return true;
        if (qEmail.indexOf('@') >= 0 && String(o.EMAIL).indexOf(qEmail) >= 0) return true;
        const key = SearchIndex.nameKey(o) + ' ' + Utils.normalizeText(o.EMAIL);
        return tokens.every(t => key.indexOf(t) >= 0);
      });
    }
    const eq = (field, val) => { if (val) recs = recs.filter(o => String(o[field]) === String(val)); };
    eq('MUNICIPIO', f.municipio);
    eq('ESTADO_CRM', f.estadoCrm);
    eq('ESTADO_FORMULARIO', f.estadoFormulario);
    eq('ORIGEN', f.origen);
    eq('PRIORIDAD', f.prioridad);
    eq('TAMANO_EMPRESA', f.tamano);
    eq('ULT_ANO_RENOVACION', f.ultAno);
    if (f.responsable === '__SIN__') recs = recs.filter(o => !o.RESPONSABLE_EMAIL);
    else eq('RESPONSABLE_EMAIL', f.responsable);
    if (f.ciiu) recs = recs.filter(o => String(o.CIIU_CODIGO).indexOf(String(f.ciiu).toUpperCase()) === 0);
    if (f.sinEmail) recs = recs.filter(o => !o.EMAIL);
    if (f.sinTelefono) recs = recs.filter(o => !o.TELEFONO);
    if (f.telefonoCompartido) recs = recs.filter(o => o.TELEFONO_COMPARTIDO);
    if (f.noContactar === 'SI') recs = recs.filter(o => o.NO_CONTACTAR);
    if (f.noContactar === 'NO') recs = recs.filter(o => !o.NO_CONTACTAR);

    const sort = params.sort || 'actualizado';
    if (sort === 'nombre') recs.sort((a, b) => String(a.RAZON_SOCIAL).localeCompare(String(b.RAZON_SOCIAL)));
    else if (sort === 'id') recs.sort((a, b) => String(a.CLIENTE_ID).localeCompare(String(b.CLIENTE_ID)));
    else recs.sort((a, b) => String(b.ACTUALIZADO_EN).localeCompare(String(a.ACTUALIZADO_EN)));

    const pageSize = Math.min(Math.max(Number(params.pageSize) || 25, 10), 100);
    const total = recs.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const page = Math.min(Math.max(Number(params.page) || 1, 1), pages);
    return {
      items: recs.slice((page - 1) * pageSize, page * pageSize),
      total: total, page: page, pages: pages, pageSize: pageSize,
      all: params.returnAll ? recs : undefined
    };
  },

  /** Mapa CLIENTE_ID -> nombre visible (para tableros y listados). */
  nameMap() {
    const map = {};
    SearchIndex.records().forEach(o => {
      map[o.CLIENTE_ID] = o.NOMBRE_COMERCIAL || o.RAZON_SOCIAL || o.CLIENTE_ID;
    });
    return map;
  }
};

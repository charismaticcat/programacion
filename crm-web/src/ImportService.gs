/**
 * ImportService.gs — Importa la base existente (CCH@E26-9316) a CLIENTES.
 *
 * Reglas confirmadas por el cliente:
 *  D1 NOMBRE_PROPIETARIO en blanco · D3 no se lee la columna NIT · D4 sin tipo de persona
 *  D5 emails repetidos o mal escritos → en blanco · D6 ACTIVIDAD vacía → descripción CIIU
 *  D7 teléfonos de 7 dígitos → 608 + número · D9 solo Neiva, Pitalito, Palermo, Rivera
 *
 * Por el límite de 6 minutos se procesa por lotes y se reanuda con un activador.
 * Es idempotente: una fila ya importada (ID_REGISTRO_ORIGEN) no se duplica.
 */
const IMPORT_COLUMNS = ['RAZON SOCIAL', 'FEC-MATRICULA', 'FEC-RENOVACION', 'ULT-ANO_REN', 'DIR-COMERCIAL',
  'BARRIO-COMERCIAL', 'MUN-COMERCIAL', 'TEL-COM-1', 'EMAIL-COMERCIAL', 'CIIU-1', 'ACTIVIDAD', 'TAM-EMPRESA'];

const ImportService = {
  status() {
    return {
      sourceConfigured: !!Props.get('IMPORT_SOURCE_SPREADSHEET_ID'),
      state: Props.get('IMPORT_STATE') || 'NO_INICIADA',
      cursor: Number(Props.get('IMPORT_CURSOR') || 0),
      total: Number(Props.get('IMPORT_TOTAL') || 0),
      imported: Number(Props.get('IMPORT_IMPORTED') || 0),
      lastMessage: Props.get('IMPORT_MESSAGE') || ''
    };
  },

  /** Conjunto de valores que se repiten en una columna (en minúsculas). */
  _repeated(values, normalizer) {
    const counts = {};
    values.forEach(v => {
      const k = normalizer(v);
      if (k) counts[k] = (counts[k] || 0) + 1;
    });
    const rep = {};
    Object.keys(counts).forEach(k => { if (counts[k] > 1) rep[k] = true; });
    return rep;
  },

  _source() {
    const id = Props.require('IMPORT_SOURCE_SPREADSHEET_ID');
    const ss = SpreadsheetApp.openById(id);
    const sh = ss.getSheetByName(CONFIG.IMPORT_SOURCE_SHEET) || ss.getSheets()[0];
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => Utils.cleanSpaces(h).toUpperCase());
    const idx = {};
    IMPORT_COLUMNS.forEach(c => {
      const i = header.indexOf(c);
      if (i < 0) throw new AppError('IMPORT_BAD_SOURCE', 'La base no tiene la columna "' + c + '".');
      idx[c] = i + 1;
    });
    return { sheet: sh, idx: idx, total: sh.getLastRow() - 1 };
  },

  /**
   * Transforma una fila de la base en un registro de CLIENTES (función pura, probada en tests).
   * @param {Object} r fila {COLUMNA: valor}
   * @param {number} rowNumber número de registro (1 = primera fila de datos)
   * @param {Object} repeatedEmails / repeatedPhones conjuntos de valores repetidos
   */
  transformRow(r, rowNumber, repeatedEmails, repeatedPhones) {
    const notes = [];
    const phone = Utils.normalizePhone(r['TEL-COM-1']);
    if (phone.note) notes.push(phone.note);
    let email = Utils.normalizeEmail(r['EMAIL-COMERCIAL']);
    if (email && !Utils.isValidEmail(email)) { notes.push('EMAIL_INVALIDO_DESCARTADO'); email = ''; }
    if (email && repeatedEmails[email]) { notes.push('EMAIL_REPETIDO_DESCARTADO'); email = ''; }
    const ciiu = Utils.parseCiiu(r['CIIU-1']);
    let actividad = Utils.cleanSpaces(r['ACTIVIDAD']);
    let fuente = 'BASE';
    if (!actividad && ciiu.desc) { actividad = ciiu.desc; fuente = 'CIIU'; notes.push('ACTIVIDAD_DESDE_CIIU'); }
    if (!actividad) fuente = '';
    const mun = Utils.splitCodeName(r['MUN-COMERCIAL']);
    let municipio = Validation.municipioName(mun.name) || Validation.municipioName(mun.code);
    if (!municipio) notes.push('MUNICIPIO_FUERA_DE_LISTA:' + mun.name);
    const barrio = Utils.splitCodeName(r['BARRIO-COMERCIAL']);
    const tam = Utils.cleanSpaces(r['TAM-EMPRESA']).toUpperCase();
    return {
      CLIENTE_ID: 'CLI-' + Seq.pad(rowNumber, 6),
      ORIGEN: 'BASE_CCH',
      ID_REGISTRO_ORIGEN: String(rowNumber),
      NOMBRE_PROPIETARIO: '',
      RAZON_SOCIAL: Utils.cleanSpaces(r['RAZON SOCIAL']),
      NOMBRE_COMERCIAL: '',
      TIPO_ESTABLECIMIENTO: '',
      ACTIVIDAD: actividad,
      ACTIVIDAD_FUENTE: fuente,
      CIIU_CODIGO: ciiu.code,
      CIIU_DESCRIPCION: ciiu.desc,
      TAMANO_EMPRESA: ENUMS.TAMANO.indexOf(tam) >= 0 ? tam : '',
      TELEFONO: phone.value,
      TELEFONO_ORIGINAL: Utils.cleanSpaces(r['TEL-COM-1']),
      TELEFONO_COMPARTIDO: !!(phone.value && repeatedPhones[phone.value]),
      EMAIL: email,
      MUNICIPIO_COD: municipio ? Validation.municipioCode(municipio) : mun.code,
      MUNICIPIO: municipio,
      DIRECCION: Utils.cleanSpaces(r['DIR-COMERCIAL']),
      BARRIO_COD: barrio.code,
      BARRIO: barrio.name,
      FEC_MATRICULA: Utils.yyyymmddToIso(r['FEC-MATRICULA']),
      FEC_RENOVACION: Utils.yyyymmddToIso(r['FEC-RENOVACION']),
      ULT_ANO_RENOVACION: Utils.digits(r['ULT-ANO_REN']).slice(0, 4),
      ESTADO_CRM: 'NUEVO',
      ESTADO_FORMULARIO: 'NO_INICIADO',
      RESPONSABLE_EMAIL: '',
      PRIORIDAD: Utils.digits(r['ULT-ANO_REN']).slice(0, 4) === Utils.year() ? 'MEDIA' : 'BAJA',
      ETIQUETAS: '',
      FUSIONADO_EN: '',
      DRIVE_FOLDER_ID: '',
      NO_CONTACTAR: false,
      OBSERVACION_IMPORTACION: notes.join('; ')
    };
  },

  /** Inicia (o reinicia desde cero el cursor) la importación. */
  start() {
    const src = ImportService._source();
    Props.set('IMPORT_STATE', 'EN_CURSO');
    Props.set('IMPORT_CURSOR', '0');
    Props.set('IMPORT_TOTAL', String(src.total));
    Props.set('IMPORT_IMPORTED', '0');
    Props.set('IMPORT_MESSAGE', 'Preparando…');
    // Precalcula emails y teléfonos repetidos sobre TODA la base (regla D5).
    const emails = src.sheet.getRange(2, src.idx['EMAIL-COMERCIAL'], src.total, 1).getValues().map(r => r[0]);
    const phones = src.sheet.getRange(2, src.idx['TEL-COM-1'], src.total, 1).getValues().map(r => r[0]);
    const repEmails = Object.keys(ImportService._repeated(emails, v => Utils.normalizeEmail(v)));
    const repPhones = Object.keys(ImportService._repeated(phones, v => Utils.normalizePhone(v).value));
    ImportService._saveSet('IMPORT_REP_EMAILS', repEmails);
    ImportService._saveSet('IMPORT_REP_PHONES', repPhones);
    Audit.log({ accion: 'IMPORTAR_INICIO', entidad: 'CLIENTES', detalle: src.total + ' filas · ' + repEmails.length + ' emails repetidos · ' + repPhones.length + ' teléfonos repetidos' });
    return ImportService.continueImport();
  },

  /** Conjuntos grandes se guardan en un archivo JSON de Drive (Script Properties tiene límite de 9 KB). */
  _saveSet(key, list) {
    const folder = DriveApp.getFolderById(Props.require('SYSTEM_FOLDER_ID'));
    const oldId = Props.get(key);
    if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) { /* ya no existe */ } }
    const file = folder.createFile(Utilities.newBlob(JSON.stringify(list), 'application/json', key + '.json'));
    Props.set(key, file.getId());
  },
  _loadSet(key) {
    const id = Props.get(key);
    const set = {};
    if (!id) return set;
    (Utils.safeJsonParse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'), []) || []).forEach(v => { set[v] = true; });
    return set;
  },

  /** Procesa lotes hasta ~4,5 minutos y programa la continuación si falta. */
  continueImport() {
    const started = Date.now();
    const src = ImportService._source();
    const repEmails = ImportService._loadSet('IMPORT_REP_EMAILS');
    const repPhones = ImportService._loadSet('IMPORT_REP_PHONES');
    let cursor = Number(Props.get('IMPORT_CURSOR') || 0);
    let imported = Number(Props.get('IMPORT_IMPORTED') || 0);
    const total = src.total;
    // IDs ya existentes para no duplicar si se repite un lote.
    const existing = {};
    SheetService.readColumns('CLIENTES', ['ID_REGISTRO_ORIGEN'], { includeDeleted: true }).forEach(o => {
      if (o.ID_REGISTRO_ORIGEN) existing[o.ID_REGISTRO_ORIGEN] = true;
    });
    const width = src.sheet.getLastColumn();
    while (cursor < total && Date.now() - started < CONFIG.IMPORT_MAX_MS) {
      const n = Math.min(CONFIG.IMPORT_BATCH_SIZE, total - cursor);
      const values = src.sheet.getRange(2 + cursor, 1, n, width).getValues();
      const recs = [];
      values.forEach((row, i) => {
        const rowNumber = cursor + i + 1;
        if (existing[String(rowNumber)]) return;
        const r = {};
        IMPORT_COLUMNS.forEach(c => { r[c] = row[src.idx[c] - 1]; });
        if (!Utils.cleanSpaces(r['RAZON SOCIAL'])) return;
        recs.push(ImportService.transformRow(r, rowNumber, repEmails, repPhones));
      });
      if (recs.length) SheetService.insertMany('CLIENTES', recs);
      imported += recs.length;
      cursor += n;
      Props.set('IMPORT_CURSOR', String(cursor));
      Props.set('IMPORT_IMPORTED', String(imported));
      Props.set('IMPORT_MESSAGE', 'Importadas ' + imported + ' de ' + total);
    }
    if (cursor >= total) {
      Props.set('IMPORT_STATE', 'COMPLETADA');
      Props.set('IMPORT_MESSAGE', 'Importación completa: ' + imported + ' registros.');
      Seq.ensureAtLeast('CLIENTE', total);
      ImportService._removeContinuationTrigger();
      SearchIndex.invalidate();
      SearchIndex.build();
      Audit.log({ accion: 'IMPORTAR_FIN', entidad: 'CLIENTES', detalle: imported + ' registros importados' });
    } else {
      ImportService._scheduleContinuation();
    }
    return ImportService.status();
  },

  _scheduleContinuation() {
    ImportService._removeContinuationTrigger();
    ScriptApp.newTrigger('cronContinueImport').timeBased().after(60 * 1000).create();
  },
  _removeContinuationTrigger() {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'cronContinueImport') ScriptApp.deleteTrigger(t);
    });
  }
};

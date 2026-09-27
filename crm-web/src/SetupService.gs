/**
 * SetupService.gs — Instalación idempotente (se ejecuta desde el editor de Apps Script).
 *
 * setup() crea, si no existen:
 *  - Carpeta CRM_WEB_SISTEMA (hojas de cálculo y archivos internos)
 *  - Carpeta CLIENTES_WEB (carpetas por cliente)
 *  - Spreadsheet CRM_WEB_DB con todas las hojas y encabezados
 *  - Spreadsheet CRM_WEB_AUDITORIA
 *  - Catálogos iniciales, productos web y ajustes por defecto
 *  - Activadores (reportes cada 10 min, tareas diarias)
 *  - Código de configuración para crear el primer Super Admin
 * Guarda todos los IDs en Script Properties (nunca en el código).
 */
const SetupService = {
  SEED_TIPOS: ['Restaurante', 'Cafetería', 'Panadería', 'Tienda / Minimercado', 'Droguería', 'Peluquería / Barbería',
    'Spa / Estética', 'Ropa y accesorios', 'Calzado', 'Ferretería', 'Taller automotriz / Motos', 'Repuestos',
    'Consultorio médico', 'Odontología', 'Veterinaria', 'Hotel / Hospedaje', 'Papelería', 'Tecnología / Celulares',
    'Construcción', 'Transporte', 'Educación', 'Gimnasio', 'Servicios profesionales', 'Eventos', 'Agropecuario', 'Otro'],
  SEED_MOTIVOS: ['Precio', 'No tiene presupuesto', 'Ya tiene página web', 'No le interesa', 'No contesta', 'Eligió a otro proveedor', 'Negocio cerrado'],
  SEED_PRODUCTOS: [
    ['Página web básica', 'Página informativa con portada, quiénes somos, servicios, ubicación y contacto.', false],
    ['Página con catálogo', 'Página con catálogo de productos y precios.', false],
    ['Agenda de citas', 'Módulo para que los clientes reserven citas.', true],
    ['Base de datos de clientes', 'Registro de clientes del negocio.', true]
  ],

  run() {
    const log = [];
    // 1) Carpetas
    let systemId = Props.get('SYSTEM_FOLDER_ID');
    if (!SetupService._folderOk(systemId)) {
      systemId = DriveApp.createFolder(CONFIG.SYSTEM_FOLDER_NAME).getId();
      Props.set('SYSTEM_FOLDER_ID', systemId);
      log.push('Carpeta creada: ' + CONFIG.SYSTEM_FOLDER_NAME);
    }
    let rootId = Props.get('ROOT_FOLDER_ID');
    if (!SetupService._folderOk(rootId)) {
      rootId = DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME).getId();
      Props.set('ROOT_FOLDER_ID', rootId);
      log.push('Carpeta creada: ' + CONFIG.ROOT_FOLDER_NAME);
    }
    const systemFolder = DriveApp.getFolderById(systemId);

    // 2) Hojas de cálculo
    let dbId = Props.get('DB_SPREADSHEET_ID');
    if (!SetupService._fileOk(dbId)) {
      const ss = SpreadsheetApp.create(CONFIG.DB_SPREADSHEET_NAME);
      DriveApp.getFileById(ss.getId()).moveTo(systemFolder);
      dbId = ss.getId();
      Props.set('DB_SPREADSHEET_ID', dbId);
      log.push('Spreadsheet creado: ' + CONFIG.DB_SPREADSHEET_NAME);
    }
    let auditId = Props.get('AUDIT_SPREADSHEET_ID');
    if (!SetupService._fileOk(auditId)) {
      const ss = SpreadsheetApp.create(CONFIG.AUDIT_SPREADSHEET_NAME);
      DriveApp.getFileById(ss.getId()).moveTo(systemFolder);
      auditId = ss.getId();
      Props.set('AUDIT_SPREADSHEET_ID', auditId);
      log.push('Spreadsheet creado: ' + CONFIG.AUDIT_SPREADSHEET_NAME);
    }
    SheetService.resetCache();
    const db = SpreadsheetApp.openById(dbId);
    const auditDb = SpreadsheetApp.openById(auditId);
    db.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
    auditDb.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

    SchemaUtil.tables().forEach(table => {
      const ss = SCHEMA[table].auditDb ? auditDb : db;
      log.push(SetupService._ensureSheet(ss, table));
    });
    [db, auditDb].forEach(ss => {
      ['Hoja 1', 'Hoja1', 'Sheet1'].forEach(n => {
        const sh = ss.getSheetByName(n);
        if (sh && ss.getSheets().length > 1 && sh.getLastRow() === 0) ss.deleteSheet(sh);
      });
    });
    SheetService.resetCache();

    // 3) Datos iniciales
    Auth._pepper();
    Ctx.set('SISTEMA:setup', 'SISTEMA', 'SISTEMA');
    SetupService._seed(log);

    // 4) Activadores
    SetupService.installTriggers();
    log.push('Activadores instalados.');

    // 5) Primer Super Admin
    if (!Auth.hasActiveSuperAdmin()) {
      const code = Auth.createSetupCode();
      log.push('');
      log.push('=== CÓDIGO DE CONFIGURACIÓN (válido ' + CONFIG.SETUP_CODE_TTL_HOURS + ' h, un solo uso): ' + code + ' ===');
      log.push('Abra la URL de la aplicación con ?page=admin y elija "Crear Super Admin".');
    }
    Audit.log({ accion: 'SETUP', entidad: 'SISTEMA', detalle: 'Instalación/actualización ejecutada' });
    Audit.flush();
    log.push('');
    log.push('DB: ' + db.getUrl());
    log.push('Auditoría: ' + auditDb.getUrl());
    log.push('CLIENTES_WEB: ' + DriveService.folderUrl(rootId));
    return log;
  },

  _folderOk(id) {
    if (!id) return false;
    try { return !DriveApp.getFolderById(id).isTrashed(); } catch (e) { return false; }
  },
  _fileOk(id) {
    if (!id) return false;
    try { return !DriveApp.getFileById(id).isTrashed(); } catch (e) { return false; }
  },

  /** Crea la hoja o agrega columnas faltantes al final (nunca borra ni reordena). */
  _ensureSheet(ss, table) {
    const cols = SchemaUtil.columnNames(table);
    let sh = ss.getSheetByName(table);
    if (!sh) {
      sh = ss.insertSheet(table);
      if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
      sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#0f3d3e').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
      return 'Hoja creada: ' + table;
    }
    const lastCol = sh.getLastColumn();
    const header = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
    const missing = cols.filter(c => header.indexOf(c) < 0);
    if (!missing.length) return 'Hoja OK: ' + table;
    const start = header.length + 1;
    if (sh.getMaxColumns() < start + missing.length - 1) sh.insertColumnsAfter(sh.getMaxColumns(), start + missing.length - 1 - sh.getMaxColumns());
    sh.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#0f3d3e').setFontColor('#ffffff');
    sh.getRange(1, start, sh.getMaxRows(), missing.length).setNumberFormat('@');
    return 'Hoja actualizada: ' + table + ' (+' + missing.join(', ') + ')';
  },

  _seed(log) {
    const existing = SheetService.readAll('CATALOGOS', { includeDeleted: true });
    const has = (cat, label) => existing.some(r => r.CATALOGO === cat && Utils.normalizeText(r.ETIQUETA) === Utils.normalizeText(label));
    const rows = [];
    CONFIG.MUNICIPIOS.forEach((m, i) => {
      if (!has('MUNICIPIOS', m.name)) rows.push({ CATALOGO: 'MUNICIPIOS', CODIGO: m.code, ETIQUETA: m.name, ORDEN: i + 1, ACTIVO: true });
    });
    SetupService.SEED_TIPOS.forEach((t, i) => {
      if (!has('TIPOS_ESTABLECIMIENTO', t)) rows.push({ CATALOGO: 'TIPOS_ESTABLECIMIENTO', CODIGO: Utils.slug(t, 40), ETIQUETA: t, ORDEN: i + 1, ACTIVO: true });
    });
    SetupService.SEED_MOTIVOS.forEach((t, i) => {
      if (!has('MOTIVOS_CIERRE', t)) rows.push({ CATALOGO: 'MOTIVOS_CIERRE', CODIGO: Utils.slug(t, 40), ETIQUETA: t, ORDEN: i + 1, ACTIVO: true });
    });
    if (rows.length) {
      SheetService.insertMany('CATALOGOS', rows);
      log.push('Catálogos iniciales: ' + rows.length + ' ítems');
    }
    if (!SheetService.readAll('PRODUCTOS_WEB', { includeDeleted: true }).length) {
      SheetService.insertMany('PRODUCTOS_WEB', SetupService.SEED_PRODUCTOS.map((p, i) => ({
        NOMBRE: p[0], DESCRIPCION: p[1], PRECIO_SUGERIDO: 0, ES_ADICIONAL: p[2], ACTIVO: true, ORDEN: i + 1
      })));
      log.push('Productos web iniciales creados (precio sugerido vacío, editable).');
    }
    const cfg = SheetService.readAll('CONFIGURACION');
    const cfgRows = Object.keys(EDITABLE_SETTINGS).filter(k => !cfg.some(r => r.CLAVE === k)).map(k => {
      const d = EDITABLE_SETTINGS[k];
      const v = d.type === 'b' ? (d.def ? 'SI' : 'NO') : String(d.def);
      return { CLAVE: k, VALOR: v, DESCRIPCION: d.label, ACTUALIZADO_EN: Utils.nowIso(), ACTUALIZADO_POR: 'SISTEMA:setup' };
    });
    if (cfgRows.length) SheetService.insertMany('CONFIGURACION', cfgRows);
    Settings.invalidate();
  },

  installTriggers() {
    const wanted = { cronProcessReports: true, cronDaily: true };
    ScriptApp.getProjectTriggers().forEach(t => {
      if (wanted[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('cronProcessReports').timeBased().everyMinutes(10).create();
    ScriptApp.newTrigger('cronDaily').timeBased().everyDays(1).atHour(6).create();
  },

  /** Tareas diarias: índice, revisión de permisos, borradores abandonados. */
  daily() {
    SearchIndex.invalidate();
    SearchIndex.build();
    const problems = DriveService.auditSharing();
    if (problems.length) {
      Audit.log({ accion: 'ALERTA_PERMISOS', entidad: 'SISTEMA', resultado: 'ERROR', detalle: 'Compartido públicamente: ' + problems.join(', ') });
    }
    // Borradores sin actividad por más de 60 días pasan a DESCARTADA.
    const limit = Utils.addDays(null, -60);
    SheetService.readColumns('SOLICITUDES', ['SOLICITUD_ID', 'ESTADO', 'ACTUALIZADO_EN']).forEach(s => {
      if (s.ESTADO === 'BORRADOR' && String(s.ACTUALIZADO_EN).slice(0, 10) < limit) {
        SheetService.update('SOLICITUDES', s.SOLICITUD_ID, { ESTADO: 'DESCARTADA', OBSERVACIONES_ADMIN: 'Borrador abandonado (60 días).' });
      }
    });
    return { permisos: problems };
  }
};

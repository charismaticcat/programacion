/**
 * Code.gs — Punto de entrada de la Web App y funciones del propietario.
 *
 * Web App (una sola implementación):
 *   …/exec              → formulario público (Index.html)
 *   …/exec?page=admin   → Super Admin (Admin.html, exige inicio de sesión)
 *
 * Reglas de exposición (Parte H):
 *   - google.script.run solo puede llamar funciones GLOBALES cuyo nombre NO termina en "_".
 *   - Funciones públicas:        pub_*  (PublicApi.gs)
 *   - Funciones administrativas: adm_*  (AdminApi.gs) → exigen token de sesión y permiso.
 *   - Mantenimiento: setup, createSetupCode, installTriggers, importBase, rebuildIndex, cron*
 *     → solo el propietario (editor) o un activador instalado por él.
 *   - Todo lo demás vive dentro de objetos (servicios) o termina en "_" y no es invocable.
 */

function doGet(e) {
  const isAdmin = e && e.parameter && e.parameter.page === 'admin';
  const template = HtmlService.createTemplateFromFile(isAdmin ? 'Admin' : 'Index');
  let title = isAdmin ? CONFIG.ADMIN_APPLICATION_NAME : CONFIG.APPLICATION_NAME;
  if (!isAdmin) {
    try { title = Settings.get('APP_NAME'); } catch (err) { /* antes de setup() */ }
  }
  template.pageTitle = title;
  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Incluye un archivo HTML dentro de una plantilla. Privada (termina en "_"). */
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** Limpia el estado de módulo (en Apps Script cada llamada ya es un proceso nuevo). */
function resetRequestState_() {
  Audit._buffer = [];
  Lock._depth = 0;
  SearchIndex._mem = null;
  Settings._mem = null;
  Ctx.set('ANONIMO', '', 'SISTEMA', null);
}

/** Ejecuta una operación y devuelve {ok, data} o {ok:false, error} serializable. */
function respond_(fn) {
  resetRequestState_();
  try {
    const data = fn();
    Audit.flush();
    return JSON.parse(JSON.stringify({ ok: true, data: data === undefined ? null : data }));
  } catch (err) {
    Audit.flush();
    return errorResponse_(err);
  }
}

function errorResponse_(err) {
  if (err instanceof AppError) {
    return JSON.parse(JSON.stringify({ ok: false, error: { code: err.code, message: err.message, details: err.details } }));
  }
  const ref = Utils.newId('ERR');
  console.error(ref, err && err.stack ? err.stack : err);
  return { ok: false, error: { code: 'INTERNAL', message: 'Ocurrió un error inesperado. Referencia: ' + ref } };
}

/** Ejecuta una tarea del propietario y registra el resultado en el log de ejecución. */
function ownerRun_(e, name, fn) {
  resetRequestState_();
  Auth.requireOwnerContext(e);
  try {
    const out = fn();
    Audit.flush();
    const text = Array.isArray(out) ? out.join('\n') : JSON.stringify(out, null, 2);
    console.log(name + '\n' + text);
    Logger.log(text);
    return out;
  } catch (err) {
    Audit.flush();
    console.error(name, err && err.stack ? err.stack : err);
    throw err;
  }
}

// ======================= Funciones del propietario =======================

/** 1) Ejecutar una vez desde el editor: crea hojas, carpetas, activadores y el código del primer Super Admin. */
function setup() {
  return ownerRun_(null, 'setup', () => SetupService.run());
}

/** Genera un nuevo código para crear o restablecer un Super Admin (recuperación de acceso). */
function createSetupCode() {
  return ownerRun_(null, 'createSetupCode', () => {
    const code = Auth.createSetupCode();
    return ['Código de configuración (válido ' + CONFIG.SETUP_CODE_TTL_HOURS + ' h, un solo uso): ' + code,
      'Abra la aplicación con ?page=admin y elija "Crear o restablecer Super Admin".'];
  });
}

/** Reinstala los activadores (reportes cada 10 min y tareas diarias). */
function installTriggers() {
  return ownerRun_(null, 'installTriggers', () => { SetupService.installTriggers(); return ['Activadores instalados.']; });
}

/**
 * Importa la base existente. Antes: suba CCH@E26-9316.xlsx a Drive, ábralo como
 * Hojas de cálculo de Google y guarde su ID en la propiedad IMPORT_SOURCE_SPREADSHEET_ID.
 */
function importBase() {
  return ownerRun_(null, 'importBase', () => ImportService.start());
}

/** Reconstruye el índice de búsqueda de clientes. */
function rebuildIndex() {
  return ownerRun_(null, 'rebuildIndex', () => { SearchIndex.invalidate(); return { registros: SearchIndex.build().length }; });
}

// ======================= Activadores =======================

function cronProcessReports(e) {
  return ownerRun_(e, 'cronProcessReports', () => {
    // Mantiene el índice de búsqueda en caché para que el formulario no espere a reconstruirlo.
    const indexRebuilt = SearchIndex.warm();
    const r = ReportService.processPending();
    r.indiceReconstruido = indexRebuilt;
    return r;
  });
}

function cronDaily(e) {
  return ownerRun_(e, 'cronDaily', () => SetupService.daily());
}

function cronContinueImport(e) {
  return ownerRun_(e, 'cronContinueImport', () => ImportService.continueImport());
}

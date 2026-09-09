/**
 * Config.gs — Foro Educativo Comunal Neiva 2026
 *
 * Sustituye el patrón de constantes hardcodeadas de FEI 3.1 (Código.js:34-107,
 * ver docs/01-auditoria-fei-3.1.md §1.4/§4) por una hoja `ConfiguracionComunal`
 * (clave/valor) editable sin tocar código. El único dato que sigue siendo
 * "de arranque" es el ID del spreadsheet del proyecto nuevo — pero ni
 * siquiera ese se hardcodea: se autoprovisiona la primera vez que se
 * ejecuta el proyecto (mismo patrón que obtenerSpreadsheetAnalisisFEM_ en
 * 3.1, que crea el spreadsheet satélite perezosamente si no existe) y su
 * ID se guarda en PropertiesService.
 *
 * IMPORTANTE — proyecto independiente: la clave de PropertiesService usada
 * aquí (SPREADSHEET_ID_COMUNAL) es exclusiva de este proyecto y no
 * coincide con ninguna usada por FEI 3.1; no se leen ni escriben hojas ni
 * propiedades de 3.1 en ningún punto de este archivo.
 */

var CLAVE_PROP_SPREADSHEET_COMUNAL = "SPREADSHEET_ID_COMUNAL";

/** Valores por defecto — se usan para sembrar ConfiguracionComunal la primera vez y como fallback. */
var CONFIG_POR_DEFECTO_ = {
  ID_FORO_COMUNAL: "FEC-NEIVA-2026",
  NOMBRE_FORO: "Foro Educativo Comunal Neiva 2026",
  SUBTITULO: "Encuentro de voces que construyen territorio",
  FECHA: "2026-09-24",
  CARPETA_DRIVE_ID: "",
  // URL del deployment publicado (Implementar → Nueva implementación →
  // copiar la URL /exec) — si se deja vacía, construirUrlAcceso_ recurre a
  // ScriptApp.getService().getUrl(), que fuera de una petición web real
  // (p. ej. al ejecutar una función desde el editor) devuelve el /dev del
  // proyecto, inválido para cualquier persona sin permiso de edición.
  URL_WEBAPP: "",
  PLANTILLA_INFORME: "",
  // Logo del Foro y logo de la SEM Neiva — usados en la transición inicial
  // (uno tras otro) y en el encabezado del informe. IDs reales de Drive
  // dados por la SEM; si se necesita cambiarlos, basta con editar
  // ConfiguracionComunal (no hace falta tocar código).
  LOGO_ENCABEZADO_ID: "1mFOOUZ5aFAuwM-JMxNUaDnPPznDlQ2bj",
  LOGO_PIE_ID: "1Cmx7c3ec2gQCjRc8kcNeUbZt5LiURyD5",
  CORREO_REMITENTE: "",
  COPIAS_CORREO: "",
  TIEMPO_SESION: "90",
  TIEMPO_MAXIMO_SOCIALIZACION: "10",
  MAX_SESIONES_SIMULTANEAS_GRUPO: "4",
  ACTIVO: "SI"
};

var CABECERAS_CONFIGURACION_ = ["CLAVE", "VALOR", "DESCRIPCION"];

var _configCacheComunal_ = null;

/**
 * Devuelve la configuración del proyecto (objeto {CLAVE: valor}), leída de
 * la hoja ConfiguracionComunal, con fallback a CONFIG_POR_DEFECTO_ para
 * cualquier clave ausente. Memoizada en memoria durante la ejecución
 * actual (cada invocación de Apps Script es de corta duración, así que no
 * hace falta CacheService entre ejecuciones).
 */
function getConfig() {
  if (_configCacheComunal_) return _configCacheComunal_;

  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty(CLAVE_PROP_SPREADSHEET_COMUNAL);
  var ss;

  if (!spreadsheetId) {
    ss = SpreadsheetApp.create("Foro Educativo Comunal Neiva 2026 — Datos");
    spreadsheetId = ss.getId();
    props.setProperty(CLAVE_PROP_SPREADSHEET_COMUNAL, spreadsheetId);
    inicializarHojasBase_(ss);
  } else {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }

  var config = leerConfiguracionComunal_(ss);
  config.SPREADSHEET_ID = spreadsheetId;
  _configCacheComunal_ = config;
  return config;
}

/** Fuerza a releer la configuración en la próxima llamada (tras escribir cambios). */
function invalidarCacheConfig_() {
  _configCacheComunal_ = null;
}

function leerConfiguracionComunal_(ss) {
  var hoja = ss.getSheetByName("ConfiguracionComunal");
  var config = Object.assign({}, CONFIG_POR_DEFECTO_);
  if (!hoja) return config;
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return config;
  var filas = hoja.getRange(2, 1, ultimaFila - 1, 2).getValues();
  filas.forEach(function (fila) {
    var clave = String(fila[0] || "").trim();
    if (!clave) return;
    var valor = fila[1];
    config[clave] = valor === "" || valor === null || valor === undefined ? config[clave] || "" : valor;
  });
  return config;
}

/** Escribe (UPSERT) una clave de configuración. Usado por Drive.gs al autoprovisionar carpetas, etc. */
function escribirConfig_(clave, valor) {
  var ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  var hoja = ss.getSheetByName("ConfiguracionComunal");
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "CLAVE", clave);
  if (fila === -1) {
    hoja.appendRow([clave, valor, ""]);
  } else {
    hoja.getRange(fila, mapa["VALOR"]).setValue(valor);
  }
  invalidarCacheConfig_();
}

/**
 * Crea, en el spreadsheet nuevo recién autoprovisionado, todas las hojas
 * base del proyecto con sus cabeceras (sin datos). Ver estructura completa
 * en docs/02-arquitectura-nuevo-proyecto.md, Fase 6.
 */
function inicializarHojasBase_(ss) {
  function crear(nombre, cabeceras) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) {
      hoja = ss.insertSheet(nombre);
    }
    if (hoja.getLastRow() === 0) {
      hoja.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]);
      hoja.setFrozenRows(1);
    }
    return hoja;
  }

  // Sheets crea la spreadsheet nueva con una "Hoja 1" por defecto; la
  // reaprovechamos para GruposComunal en vez de dejarla vacía y suelta.
  // Los encabezados de cada hoja se piden a la función cabecerasX_()
  // "dueña" de esa hoja (Instituciones.gs, Access.gs, Data.gs, etc.) en
  // vez de duplicarlos aquí, para que nunca se desincronicen entre sí.
  var hojaPorDefecto = ss.getSheets()[0];
  if (hojaPorDefecto && ss.getSheets().length === 1 && hojaPorDefecto.getLastRow() === 0) {
    var cabecerasGrupos = cabecerasGruposComunal_();
    hojaPorDefecto.setName("GruposComunal");
    hojaPorDefecto.getRange(1, 1, 1, cabecerasGrupos.length).setValues([cabecerasGrupos]);
    hojaPorDefecto.setFrozenRows(1);
  } else {
    crear("GruposComunal", cabecerasGruposComunal_());
  }

  crear("AccesosGrupo", cabecerasAccesosGrupo_());
  crear("ParticipacionComunal", cabecerasParticipacionComunal_());
  crear("Sesion1Comunal", cabecerasSesion1Comunal_());
  crear("ConectaEduca", cabecerasConectaEduca_());
  crear("InformesComunal", cabecerasInformesComunal_());
  crear(HOJA_CARACTERIZACION_IE_, cabecerasCaracterizacionIE_());
  crear(HOJA_RESPONSABLES_COMUNAL_, cabecerasResponsablesComunal_());
  crear(HOJA_VALORACION_COMUNAL_, cabecerasValoracionComunal_());
  crear(HOJA_PARTICIPACION_ESTAMENTO_, cabecerasParticipacionEstamentoIE_());
  crear(HOJA_PREPARACION_IE_, cabecerasPreparacionIE_());
  crear(HOJA_INVITADOS_, cabecerasInvitadosPreparacion_());
  crear("EnviosDiferidosComunal", ["ID_GRUPO", "FECHA_REGISTRO", "REINTENTADO"]);

  var hojaConfig = crear("ConfiguracionComunal", CABECERAS_CONFIGURACION_);
  if (hojaConfig.getLastRow() < 2) {
    var filas = Object.keys(CONFIG_POR_DEFECTO_).map(function (clave) {
      return [clave, CONFIG_POR_DEFECTO_[clave], ""];
    });
    hojaConfig.getRange(2, 1, filas.length, 3).setValues(filas);
  }
}

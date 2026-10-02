/**
 * Drive.gs — Foro Educativo Comunal Neiva 2026
 *
 * Estructura de Google Drive exacta de la spec (sección 16):
 *   FORO EDUCATIVO COMUNAL NEIVA 2026/
 *   ├── 01_DATOS
 *   ├── 02_ASISTENCIA
 *   ├── 03_INFORMES_GRUPALES/GRUPO N/{Informe Grupo N.docx, .pdf}
 *   ├── 04_CONECTAEDUCA
 *   ├── 05_CONSOLIDADO_MUNICIPAL
 *   └── 06_EVIDENCIAS
 * Patrón "obtener o crear" idéntico a crearOFolderHija_ /
 * crearEstructuraCarpetasGrupoFEM_ de FEI 3.1
 * (docs/01-auditoria-fei-3.1.md §1.4/§2.5) — el activo más directamente
 * reutilizable de 3.1 para este proyecto. La carpeta raíz NO se
 * hardcodea: se autoprovisiona la primera vez (igual que el spreadsheet
 * en Config.gs) y su ID se guarda en ConfiguracionComunal.CARPETA_DRIVE_ID.
 *
 * Independencia: nunca se usa DRIVE_CARPETA_FEM_ID ni ningún ID de
 * carpeta de FEI 3.1 — todo lo que este archivo crea es propio del
 * proyecto nuevo.
 */

var SUBCARPETAS_RAIZ_ = ["01_DATOS", "02_ASISTENCIA", "03_INFORMES_GRUPALES", "04_CONECTAEDUCA", "05_CONSOLIDADO_MUNICIPAL", "06_EVIDENCIAS"];

function crearOFolderHija_(padre, nombre) {
  var it = padre.getFoldersByName(nombre);
  if (it.hasNext()) return it.next();
  return padre.createFolder(nombre);
}

/** Intenta compartir un archivo como "cualquiera con el enlace, solo lectura"; nunca lanza si falla. */
function hacerPublicoSiEsPosible_(file) {
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log("No se pudo cambiar el permiso de compartir: " + e.message);
  }
}

/**
 * Ejecuta `fn` una sola vez por instalación, marcada por `bandera` en
 * ConfiguracionComunal — usada por las funciones `asegurarXPublicos_` de
 * abajo, que se llaman en CADA doGet (Code.gs) para no depender de que
 * alguien las ejecute a mano, pero solo deben hacer el trabajo real
 * (llamadas a Drive) la primera vez.
 *
 * BUG REAL corregido aquí (reportado por el usuario como "la página tarda
 * casi un minuto en cargar"): como `getConfig()` solo se memoiza DENTRO de
 * una misma ejecución de doGet, no entre ejecuciones, y en un evento en
 * vivo llegan muchas peticiones simultáneas (varios grupos/dispositivos
 * cargando la página casi al mismo tiempo), TODAS esas ejecuciones
 * concurrentes veían la bandera todavía vacía a la vez y corrían el mismo
 * bucle de llamadas a Drive en paralelo — con `asegurarLogosIEPublicos_`
 * (36 archivos) eso significa cientos de llamadas a la API de Drive
 * disparadas a la vez desde el mismo script, lo que agota la cuota por
 * minuto y hace que Apps Script reintente con back-off exponencial
 * (varios segundos por reintento) en cada llamada — de ahí los ~60
 * segundos. La corrección: un `LockService.getScriptLock().tryLock(0)` NO
 * bloqueante — si otra ejecución ya está haciendo el trabajo, esta
 * simplemente no hace nada y sigue de inmediato (la bandera quedará en
 * "SI" en unos segundos gracias a la que sí obtuvo el lock), en vez de
 * que todas esperen o, peor, que todas corran el bucle en paralelo.
 */
function ejecutarUnaSolaVezConLock_(bandera, fn) {
  var config = getConfig();
  if (String(config[bandera] || "") === "SI") return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;
  try {
    invalidarCacheConfig_();
    if (String(getConfig()[bandera] || "") === "SI") return;
    fn();
    escribirConfig_(bandera, "SI");
  } finally {
    lock.releaseLock();
  }
}

/**
 * Asegura que los logos del splash inicial (Foro y SEM Neiva) sean
 * visibles para cualquier persona con el enlace — si el archivo de Drive
 * solo era visible para quien lo cargó, la miniatura (urlImagenDrive en
 * el cliente) no carga para nadie más ("Logo de Foro no carga"). Se
 * ejecuta una sola vez por instalación (bandera LOGOS_SPLASH_PUBLICOS en
 * ConfiguracionComunal), no en cada carga de página.
 */
function asegurarLogosSplashPublicos_() {
  ejecutarUnaSolaVezConLock_("LOGOS_SPLASH_PUBLICOS", function () {
    var config = getConfig();
    [config.LOGO_ENCABEZADO_ID, config.LOGO_PIE_ID].forEach(function (fileId) {
      if (!fileId) return;
      try {
        hacerPublicoSiEsPosible_(DriveApp.getFileById(fileId));
      } catch (e) {
        Logger.log("No se pudo asegurar el logo público " + fileId + ": " + e.message);
      }
    });
  });
}

/**
 * Igual que asegurarLogosSplashPublicos_, pero para el logo del
 * desarrollador del pie de página (spec del usuario) — bandera propia
 * (LOGO_DESARROLLADOR_PUBLICO) porque LOGOS_SPLASH_PUBLICOS ya pudo
 * haberse marcado "SI" en instalaciones existentes, antes de que este
 * logo se agregara.
 */
function asegurarLogoDesarrolladorPublico_() {
  ejecutarUnaSolaVezConLock_("LOGO_DESARROLLADOR_PUBLICO", function () {
    var fileId = getConfig().LOGO_DESARROLLADOR_ID;
    if (!fileId) return;
    try {
      hacerPublicoSiEsPosible_(DriveApp.getFileById(fileId));
    } catch (e) {
      Logger.log("No se pudo asegurar el logo del desarrollador público: " + e.message);
    }
  });
}

/**
 * Igual que asegurarLogosSplashPublicos_, pero para los dos logos de
 * sección (Encuentro/ConectaEduca, pantallaEleccionSeccion en Index.html)
 * — bandera propia (LOGOS_SECCION_PUBLICOS) porque LOGOS_SPLASH_PUBLICOS
 * ya pudo haberse marcado "SI" en instalaciones existentes, antes de que
 * estos dos logos se agregaran.
 */
function asegurarLogosSeccionPublicos_() {
  ejecutarUnaSolaVezConLock_("LOGOS_SECCION_PUBLICOS", function () {
    var config = getConfig();
    [config.LOGO_ENCUENTRO_ID, config.LOGO_CONECTAEDUCA_ID].forEach(function (fileId) {
      if (!fileId) return;
      try {
        hacerPublicoSiEsPosible_(DriveApp.getFileById(fileId));
      } catch (e) {
        Logger.log("No se pudo asegurar el logo de sección público " + fileId + ": " + e.message);
      }
    });
  });
}

/**
 * Igual que las anteriores, pero para los formatos de asistencia oficiales
 * (Encuentro + Conecta Educa, items 13 y 18 del Documento Orientador
 * FEM2026) — se enlazan directamente desde Index.html y por correo, así
 * que deben ser visibles para cualquiera con el enlace.
 */
function asegurarFormatosAsistenciaPublicos_() {
  ejecutarUnaSolaVezConLock_("FORMATOS_ASISTENCIA_PUBLICOS", function () {
    var config = getConfig();
    [config.FORMATO_ASISTENCIA_ENCUENTRO_ID, config.FORMATO_ASISTENCIA_CONECTAEDUCA_ID].forEach(function (fileId) {
      if (!fileId) return;
      try {
        hacerPublicoSiEsPosible_(DriveApp.getFileById(fileId));
      } catch (e) {
        Logger.log("No se pudo asegurar el formato de asistencia público " + fileId + ": " + e.message);
      }
    });
  });
}

/**
 * Igual que las anteriores, pero para los 36 escudos institucionales
 * (CaracterizacionIE.LOGO_ID) — a diferencia de los logos del Foro/SEM,
 * estos se importaron desde otra fuente (Importacion.gs) y nunca se
 * les ajustó el permiso de compartir, así que `urlImagenDrive` (JS.html)
 * falla en silencio para cualquiera que no sea el propietario del
 * archivo: el usuario lo reporta como "los escudos no cargan". Bandera
 * propia (LOGOS_IE_PUBLICOS) porque estos 36 archivos son independientes
 * de los logos institucionales del Foro.
 */
function asegurarLogosIEPublicos_() {
  ejecutarUnaSolaVezConLock_("LOGOS_IE_PUBLICOS", function () {
    obtenerTodasLasInstitucionesActivas().forEach(function (ie) {
      if (!ie.logoId) return;
      try {
        hacerPublicoSiEsPosible_(DriveApp.getFileById(ie.logoId));
      } catch (e) {
        Logger.log("No se pudo asegurar el escudo público de " + ie.institucion + " (" + ie.logoId + "): " + e.message);
      }
    });
  });
}

/** Carpeta raíz del proyecto: la autoprovisiona si ConfiguracionComunal.CARPETA_DRIVE_ID está vacío. */
function obtenerCarpetaRaiz_() {
  var config = getConfig();
  var id = String(config.CARPETA_DRIVE_ID || "").trim();
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      // el ID guardado ya no es válido — se recrea abajo
    }
  }
  var carpeta = DriveApp.createFolder("FORO EDUCATIVO COMUNAL NEIVA 2026");
  escribirConfig_("CARPETA_DRIVE_ID", carpeta.getId());
  return carpeta;
}

/** Crea (si hace falta) las 6 subcarpetas raíz de la spec. Idempotente. */
function asegurarEstructuraDriveComunal_() {
  var raiz = obtenerCarpetaRaiz_();
  var carpetas = {};
  SUBCARPETAS_RAIZ_.forEach(function (nombre) {
    carpetas[nombre] = crearOFolderHija_(raiz, nombre);
  });
  return carpetas;
}

/** Carpeta "GRUPO N" dentro de 03_INFORMES_GRUPALES, con el informe único del grupo. */
function asegurarCarpetaGrupo_(grupo) {
  var estructura = asegurarEstructuraDriveComunal_();
  return crearOFolderHija_(estructura["03_INFORMES_GRUPALES"], grupo);
}

function obtenerCarpetaConectaEduca_() {
  return asegurarEstructuraDriveComunal_()["04_CONECTAEDUCA"];
}

function obtenerCarpetaEvidencias_() {
  return asegurarEstructuraDriveComunal_()["06_EVIDENCIAS"];
}

function obtenerCarpetaConsolidadoMunicipal_() {
  return asegurarEstructuraDriveComunal_()["05_CONSOLIDADO_MUNICIPAL"];
}

/** Carpeta "GRUPO N" dentro de 02_ASISTENCIA, para el listado físico subido (Asistencia.gs). */
function asegurarCarpetaAsistenciaGrupo_(grupo) {
  var estructura = asegurarEstructuraDriveComunal_();
  return crearOFolderHija_(estructura["02_ASISTENCIA"], grupo);
}

/** Carpeta "GRUPO N" dentro de 06_EVIDENCIAS, para la fotografía del encuentro del grupo. */
function asegurarCarpetaEvidenciasGrupo_(grupo) {
  var estructura = asegurarEstructuraDriveComunal_();
  return crearOFolderHija_(estructura["06_EVIDENCIAS"], grupo);
}

/**
 * Sube un archivo (foto de evidencia, PDF de asistencia manual, etc.) a
 * la carpeta correspondiente. `datosBase64` viene del cliente
 * (google.script.run no admite Blob directo desde HTML sin FileReader).
 */
function subirArchivoAGrupo_(idGrupo, grupo, datosBase64, nombreArchivo, mimeType, subcarpeta) {
  var carpetaGrupo = asegurarCarpetaGrupo_(grupo);
  var destino = subcarpeta ? crearOFolderHija_(carpetaGrupo, subcarpeta) : carpetaGrupo;
  var blob = Utilities.newBlob(Utilities.base64Decode(datosBase64), mimeType, nombreArchivo);
  var file = destino.createFile(blob);
  return file;
}

/** Sube un archivo a una carpeta ya resuelta (usado por Asistencia.gs para 02_ASISTENCIA/06_EVIDENCIAS). */
function subirArchivoACarpeta_(carpeta, datosBase64, nombreArchivo, mimeType) {
  var blob = Utilities.newBlob(Utilities.base64Decode(datosBase64), mimeType, nombreArchivo);
  return carpeta.createFile(blob);
}

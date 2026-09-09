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

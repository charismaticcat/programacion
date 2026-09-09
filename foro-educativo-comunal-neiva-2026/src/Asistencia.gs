/**
 * Asistencia.gs — Foro Educativo Comunal Neiva 2026
 *
 * Métodos de registro de asistencia del grupo, adaptados de FEI 3.1
 * (docs/01-auditoria-fei-3.1.md §1.5/§2.2/§2.5): cada grupo elige UNO de
 * dos métodos —
 *
 *   - QR / enlace: se comparte el enlace público de esta pantalla
 *     (`?asistencia=<idGrupo>`, sin token/código, igual que
 *     paginaAsistenciaQR_ en 3.1) como link o como código QR, para que
 *     cada persona firme desde su propio celular. Usa el mismo registro
 *     individual que Data.gs (registrarParticipante) — es la MISMA tabla
 *     ParticipacionComunal, solo que se llega a ella por una página
 *     pública ligera en vez del formulario completo del aplicativo.
 *   - Listado físico + fotografía: para grupos donde no es práctico que
 *     cada persona firme por celular, se sube una foto o PDF del listado
 *     en papel (a 02_ASISTENCIA/GRUPO N) y una fotografía del encuentro
 *     (a 06_EVIDENCIAS/GRUPO N) — equivalente a subirAsistenciaPDF /
 *     subirEvidenciasFEM de 3.1.
 *
 * A diferencia de 3.1 (que borraba los registros del método anterior al
 * cambiar de método — eliminarAsistenciaQRPorCambioMetodo/PDF), aquí
 * cambiar de método NO borra nada: por seguridad ante un evento en vivo,
 * se prefiere conservar cualquier dato ya capturado antes que arriesgar
 * perderlo por un cambio accidental de método. Queda documentado como
 * decisión deliberada, distinta de 3.1 (ver auditoría §7 sobre el riesgo
 * de pérdida de datos que motivó esta elección).
 */

/** Logo propio del grupo (o "" si no se ha configurado con asignarLogoGrupo). */
function obtenerLogoGrupo_(idGrupo) {
  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", String(idGrupo || "").trim());
  if (fila === -1) return "";
  return String(hoja.getRange(fila, mapa["LOGO_ID"]).getValue() || "").trim();
}

/** URL pública (sin token) para firmar asistencia por QR/enlace, propia del grupo. */
function construirUrlAsistencia_(idGrupo) {
  var config = getConfig();
  var base = config.URL_WEBAPP || ScriptApp.getService().getUrl();
  return base + "?asistencia=" + encodeURIComponent(idGrupo);
}

/** Guarda el método de asistencia elegido para el grupo (QR o LISTADO). No borra datos ya capturados. */
function guardarMetodoAsistencia(idGrupo, tokenSesion, dispositivoId, metodo) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  metodo = String(metodo || "").toUpperCase();
  if (metodo !== "QR" && metodo !== "LISTADO") {
    return { ok: false, mensaje: "Método de asistencia no reconocido." };
  }
  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", idGrupo);
  if (fila === -1) return { ok: false, mensaje: "No existe acceso para este grupo." };
  hoja.getRange(fila, mapa["METODO_ASISTENCIA"]).setValue(metodo);
  return { ok: true };
}

/** Sube el listado físico de asistencia (foto o PDF) a 02_ASISTENCIA/GRUPO N. */
function subirListadoAsistencia(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  var grupoInfo = obtenerGrupoPorId(idGrupo);
  if (!grupoInfo) return { ok: false, mensaje: "Grupo no encontrado." };

  return conLock_(function () {
    var carpeta = asegurarCarpetaAsistenciaGrupo_(grupoInfo.grupo);
    var file = subirArchivoACarpeta_(carpeta, datosBase64, nombreArchivo, mimeType);
    hacerPublicoSiEsPosible_(file);

    var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", idGrupo);
    if (fila !== -1) {
      hoja.getRange(fila, mapa["ID_LISTADO_ASISTENCIA"]).setValue(file.getId());
      hoja.getRange(fila, mapa["METODO_ASISTENCIA"]).setValue("LISTADO");
    }
    return { ok: true, fileId: file.getId(), url: file.getUrl() };
  }, 30000);
}

/** Sube la fotografía de evidencia del encuentro a 06_EVIDENCIAS/GRUPO N. */
function subirFotoEvidencia(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  var grupoInfo = obtenerGrupoPorId(idGrupo);
  if (!grupoInfo) return { ok: false, mensaje: "Grupo no encontrado." };

  return conLock_(function () {
    var carpeta = asegurarCarpetaEvidenciasGrupo_(grupoInfo.grupo);
    var file = subirArchivoACarpeta_(carpeta, datosBase64, nombreArchivo, mimeType);
    hacerPublicoSiEsPosible_(file);

    var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", idGrupo);
    if (fila !== -1) hoja.getRange(fila, mapa["ID_FOTO_EVIDENCIA"]).setValue(file.getId());
    return { ok: true, fileId: file.getId(), url: file.getUrl() };
  }, 30000);
}

/**
 * Sube la fotografía general del grupo (Participación) — independiente
 * del método de asistencia elegido, a diferencia de la fotografía de
 * evidencia (que solo aplica al método "Listado físico"). Se guarda en
 * la misma carpeta 06_EVIDENCIAS/GRUPO N, en su propia columna
 * (FOTO_GRUPO_ID) para no mezclarla con esa evidencia.
 */
function subirFotoGrupo(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  var grupoInfo = obtenerGrupoPorId(idGrupo);
  if (!grupoInfo) return { ok: false, mensaje: "Grupo no encontrado." };

  return conLock_(function () {
    var carpeta = asegurarCarpetaEvidenciasGrupo_(grupoInfo.grupo);
    var file = subirArchivoACarpeta_(carpeta, datosBase64, nombreArchivo, mimeType);
    hacerPublicoSiEsPosible_(file);

    var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", idGrupo);
    if (fila !== -1) hoja.getRange(fila, mapa["FOTO_GRUPO_ID"]).setValue(file.getId());
    return { ok: true, fileId: file.getId(), url: file.getUrl() };
  }, 30000);
}

/**
 * Registro rápido de asistencia desde la página pública de QR/enlace
 * (sin token/código, igual que registrarAsistenciaQR en 3.1 — es un punto
 * de entrada de conveniencia, no un control de seguridad). Reutiliza la
 * misma tabla y el mismo deduplicado que registrarParticipante (Data.gs).
 */
function registrarAsistenciaPublica(idGrupo, idIE, nombre, estamento, correo) {
  idGrupo = String(idGrupo || "").trim();
  var hojaAccesos = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapaAccesos = obtenerMapaCabeceras_(hojaAccesos);
  var fila = buscarFilaPorColumna_(hojaAccesos, mapaAccesos, "ID_GRUPO", idGrupo);
  if (fila === -1) return { ok: false, mensaje: "Enlace de asistencia no válido." };
  var estado = String(hojaAccesos.getRange(fila, mapaAccesos["ESTADO"]).getValue() || "").toUpperCase();
  if (estado === "BLOQUEADO" || estado === "INACTIVO") {
    return { ok: false, mensaje: "Este acceso ya no está disponible." };
  }
  // La página pública QR no pide "rol en el foro" (es para firma general,
  // no para responsables) — queda vacío, distinto del estamento.
  return registrarParticipante(idGrupo, idIE, nombre, estamento, "", correo, "PUBLICO-QR");
}

/**
 * Sirve la página pública de asistencia QR/enlace (doGet ?asistencia=idGrupo,
 * ver Code.gs). Muestra el logo y nombre del grupo, sus IE, y el panel de
 * firmantes — sin pedir token/código, igual que paginaAsistenciaQR_ en 3.1.
 */
function paginaAsistenciaGrupo_(idGrupo) {
  var grupoInfo = obtenerGrupoPorId(idGrupo);
  var template = HtmlService.createTemplateFromFile("AsistenciaPublica");
  template.ID_GRUPO = idGrupo;
  template.NOMBRE_GRUPO = grupoInfo ? grupoInfo.grupo : "";
  template.EXISTE_GRUPO = !!grupoInfo;
  template.LOGO_GRUPO_URL = "";
  var logoId = obtenerLogoGrupo_(idGrupo);
  if (logoId) {
    try {
      template.LOGO_GRUPO_URL = "https://drive.google.com/thumbnail?id=" + logoId;
    } catch (e) {
      template.LOGO_GRUPO_URL = "";
    }
  }
  template.NOMBRE_FORO = getConfig().NOMBRE_FORO;
  template.instituciones = grupoInfo ? obtenerInstitucionesDelGrupo(idGrupo) : [];
  return template
    .evaluate()
    .setTitle("Asistencia — " + (grupoInfo ? grupoInfo.grupo : getConfig().NOMBRE_FORO))
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Correo.gs — Foro Educativo Comunal Neiva 2026
 *
 * Envío de accesos por grupo y del informe final, con cola de reintento
 * por cuota diaria de Gmail agotada — el patrón más valioso a preservar
 * tal cual de FEI 3.1 (enviarInformeFEM, docs/01-auditoria-fei-3.1.md
 * §1.5/§2.4/§5): un correo que falla por cuota NO es un error para quien
 * usa el formulario (el informe ya quedó generado y disponible para
 * descarga), se reintenta después vía EnviosDiferidosComunal.
 */

var HOJA_ENVIOS_DIFERIDOS_ = "EnviosDiferidosComunal";

/** Reconoce el error de cuota diaria de correo agotada, sin importar el idioma del mensaje. */
function esErrorCuotaCorreoAgotada_(error) {
  var msg = String((error && error.message) || error || "").toLowerCase();
  return /demasiadas veces|too many times|quota/.test(msg);
}

function registrarEnvioDiferido_(idGrupo) {
  var hoja = obtenerHoja_(HOJA_ENVIOS_DIFERIDOS_, ["ID_GRUPO", "FECHA_REGISTRO", "REINTENTADO"]);
  var mapa = obtenerMapaCabeceras_(hoja);
  if (buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", idGrupo) !== -1) return; // ya estaba en la cola
  hoja.appendRow([idGrupo, new Date(), "NO"]);
}

/** Remitente configurado, validado contra la cuenta/alias que realmente ejecuta el script. */
function remitenteValido_() {
  var config = getConfig();
  var remitente = String(config.CORREO_REMITENTE || "").trim();
  if (!remitente) return { ok: true, remitente: Session.getEffectiveUser().getEmail() };
  var cuenta = Session.getEffectiveUser().getEmail().toLowerCase();
  var aliases = GmailApp.getAliases().map(function (a) {
    return a.toLowerCase();
  });
  if (cuenta === remitente.toLowerCase() || aliases.indexOf(remitente.toLowerCase()) !== -1) {
    return { ok: true, remitente: remitente };
  }
  return { ok: false, mensaje: "La cuenta de Apps Script no puede enviar como " + remitente + ". Configure esa cuenta o un alias, o deje CORREO_REMITENTE vacío." };
}

/** Envía TOKEN + código de acceso al responsable del grupo. */
function enviarAccesosGrupo(idGrupo) {
  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", String(idGrupo || "").trim());
  if (fila === -1) return { ok: false, mensaje: "No existe acceso para ese grupo. Ejecute generarAccesosGrupo() primero." };

  var acceso = leerFilaComoObjeto_(hoja, fila, mapa);
  var destinatario = String(acceso.EMAIL_RESPONSABLE_GRUPO || "").trim();
  if (!destinatario) return { ok: false, mensaje: "El grupo no tiene EMAIL_RESPONSABLE_GRUPO registrado en AccesosGrupo." };

  var remitente = remitenteValido_();
  if (!remitente.ok) return remitente;

  var config = getConfig();
  var asunto = "Acceso al " + config.NOMBRE_FORO + " — " + acceso.GRUPO;
  var cuerpo =
    "Estimado(a) responsable de " + acceso.GRUPO + ":\n\n" +
    "Le compartimos el acceso al " + config.NOMBRE_FORO + " (\"" + config.SUBTITULO + "\").\n\n" +
    "Enlace de ingreso: " + acceso.URL_ACCESO + "\n" +
    "Código de acceso: " + acceso.CODIGO_ACCESO + "\n\n" +
    "Secretaría de Educación de Neiva — " + config.NOMBRE_FORO;

  try {
    GmailApp.sendEmail(destinatario, asunto, cuerpo, {
      from: remitente.remitente,
      name: "Secretaría de Educación de Neiva",
      cc: String(config.COPIAS_CORREO || "").trim()
    });
  } catch (error) {
    return { ok: false, mensaje: "No fue posible enviar el correo de acceso: " + error.message };
  }
  return { ok: true };
}

/** Envía el informe ya generado del grupo (mismo correo/PDF a todas las IE del grupo, spec sección 18). */
function enviarInformeGrupo(idGrupo) {
  var grupoInfo = obtenerGrupoPorId(idGrupo);
  if (!grupoInfo) return { ok: false, mensaje: "Grupo no encontrado." };

  var informe = obtenerInformeGrupo(idGrupo);
  if (!informe || !informe.PDF_ID) return { ok: false, mensaje: "El grupo todavía no tiene informe generado." };

  var hojaAccesos = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapaAccesos = obtenerMapaCabeceras_(hojaAccesos);
  var filaAcceso = buscarFilaPorColumna_(hojaAccesos, mapaAccesos, "ID_GRUPO", idGrupo);

  // El informe va al responsable de envío + asistentes registrados desde
  // la app (Responsables.gs); si el grupo todavía no registró a nadie,
  // se usa el correo de acceso inicial como respaldo.
  var responsables = listarResponsablesEnvio(idGrupo);
  var destinatario = responsables.principal ? responsables.principal.correo : "";
  var copiasAsistentes = responsables.asistentes.map(function (a) { return a.correo; }).filter(Boolean);
  if (!destinatario) {
    destinatario = filaAcceso === -1 ? "" : String(hojaAccesos.getRange(filaAcceso, mapaAccesos["EMAIL_RESPONSABLE_GRUPO"]).getValue() || "").trim();
  }
  if (!destinatario) return { ok: false, mensaje: "El grupo no tiene un responsable de envío ni EMAIL_RESPONSABLE_GRUPO registrado." };

  var remitente = remitenteValido_();
  if (!remitente.ok) return remitente;

  var config = getConfig();
  var pdfFile;
  try {
    pdfFile = DriveApp.getFileById(informe.PDF_ID);
  } catch (e) {
    return { ok: false, mensaje: "No fue posible ubicar el PDF del informe: " + e.message };
  }

  var asunto = "Informe del " + grupoInfo.grupo + " — " + config.NOMBRE_FORO;
  var cuerpo =
    "Estimado(a) responsable de " + grupoInfo.grupo + ":\n\n" +
    "Adjuntamos el informe consolidado del " + grupoInfo.grupo + " del " + config.NOMBRE_FORO + ".\n" +
    "Este mismo informe está disponible para todas las instituciones educativas del grupo.\n\n" +
    "Enlace de descarga: " + informe.URL + "\n\n" +
    "Secretaría de Educación de Neiva — " + config.NOMBRE_FORO;

  var copias = copiasAsistentes.concat([String(config.COPIAS_CORREO || "").trim()]).filter(Boolean).join(",");

  try {
    GmailApp.sendEmail(destinatario, asunto, cuerpo, {
      from: remitente.remitente,
      name: "Secretaría de Educación de Neiva",
      cc: copias,
      attachments: [pdfFile.getBlob()]
    });
  } catch (error) {
    if (!esErrorCuotaCorreoAgotada_(error)) {
      return { ok: false, mensaje: "No fue posible enviar el informe: " + error.message };
    }
    registrarEnvioDiferido_(idGrupo);
    return {
      ok: true,
      diferido: true,
      mensaje: "Se alcanzó el límite diario de envíos de correo. El informe ya quedó generado y disponible " +
        "para descargar; se enviará automáticamente al correo del grupo cuando la cuota se renueve."
    };
  }

  if (filaAcceso !== -1 && mapaAccesos["FECHA_ENVIO_DEFINITIVO"]) {
    hojaAccesos.getRange(filaAcceso, mapaAccesos["FECHA_ENVIO_DEFINITIVO"]).setValue(new Date());
    hojaAccesos.getRange(filaAcceso, mapaAccesos["ESTADO"]).setValue("ENVIADO");
  }
  return { ok: true, diferido: false, url: informe.URL };
}

/** Reintenta los envíos diferidos por cuota agotada (ejecutar vía trigger diario, p. ej. a medianoche). */
function reintentarEnviosDiferidos() {
  var hoja = obtenerHoja_(HOJA_ENVIOS_DIFERIDOS_, ["ID_GRUPO", "FECHA_REGISTRO", "REINTENTADO"]);
  var mapa = obtenerMapaCabeceras_(hoja);
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { ok: true, reintentados: 0 };

  var filas = hoja.getRange(2, 1, ultimaFila - 1, 3).getValues();
  var reintentados = 0;
  filas.forEach(function (fila, i) {
    if (String(fila[2]).toUpperCase() === "SI") return;
    var idGrupo = String(fila[0] || "").trim();
    if (!idGrupo) return;
    var resultado = enviarInformeGrupo(idGrupo);
    if (resultado.ok && !resultado.diferido) {
      hoja.getRange(i + 2, mapa["REINTENTADO"]).setValue("SI");
      reintentados++;
    }
  });
  return { ok: true, reintentados: reintentados };
}

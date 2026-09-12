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

/**
 * Envía por correo el enlace de invitado (estudiante/egresado(a) o
 * padre/madre/acudiente) del grupo — pantalla de "invitados especiales"
 * que aparece tras la Confirmación de caracterización cuando el grupo
 * declaró estudiantes/egresados o padres/madres/acudientes en la matriz
 * de participación por estamento (spec del usuario). El enlace se envía
 * "por separado" al responsable de envío que cada grupo (estudiantes/
 * egresados o padres/madres/acudientes) haya designado para distribuirlo
 * a su vez entre sus pares — no se envía uno a uno a cada estudiante o
 * acudiente individual.
 */
function enviarEnlaceInvitados(idGrupo, tokenSesion, dispositivoId, tipoInvitado, correos) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  tipoInvitado = String(tipoInvitado || "").trim().toUpperCase();
  if (tipoInvitado !== "ESTUDIANTE" && tipoInvitado !== "ACUDIENTE") {
    return { ok: false, mensaje: "Tipo de invitado no válido." };
  }

  var regexCorreo = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var destinatarios = (Array.isArray(correos) ? correos : String(correos || "").split(/[,;\n]+/))
    .map(function (c) { return String(c || "").trim(); })
    .filter(function (c) { return c && regexCorreo.test(c); });
  if (!destinatarios.length) {
    return { ok: false, mensaje: "Escriba al menos un correo electrónico válido." };
  }

  var enlaces = obtenerEnlacesInvitadosGrupo(idGrupo);
  if (!enlaces.ok) return enlaces;
  var enlace = tipoInvitado === "ESTUDIANTE" ? enlaces.estudiante : enlaces.acudiente;

  var remitente = remitenteValido_();
  if (!remitente.ok) return remitente;

  var grupoInfo = obtenerGrupoPorId(idGrupo);
  var nombreGrupo = (grupoInfo && grupoInfo.grupo) || idGrupo;
  var config = getConfig();
  var textoGrupoInvitado = tipoInvitado === "ESTUDIANTE" ? "estudiantes y egresados(as)" : "padres, madres y acudientes";

  var asunto = "Enlace para " + textoGrupoInvitado + " — " + nombreGrupo + " del " + config.NOMBRE_FORO;
  var cuerpo =
    "Hola:\n\n" +
    "Les compartimos el enlace para que los/las " + textoGrupoInvitado + " del " + nombreGrupo + " registren sus " +
    "aportes al " + config.NOMBRE_FORO + " (\"" + config.SUBTITULO + "\").\n\n" +
    "Enlace: " + enlace + "\n\n" +
    "Instrucciones:\n" +
    "1. Abran el enlace desde un celular o computador.\n" +
    "2. Confirmen si son estudiantes actuales, egresados(as) o adultos responsables, según corresponda.\n" +
    "3. Seleccionen su institución educativa dentro del " + nombreGrupo + ".\n" +
    "4. Completen y envíen sus aportes — no necesitan ningún código de acceso.\n\n" +
    "Secretaría de Educación de Neiva — " + config.NOMBRE_FORO;

  try {
    GmailApp.sendEmail(destinatarios.join(","), asunto, cuerpo, {
      from: remitente.remitente,
      name: "Secretaría de Educación de Neiva"
    });
  } catch (error) {
    return { ok: false, mensaje: "No fue posible enviar el correo: " + error.message };
  }
  return { ok: true, enviados: destinatarios.length };
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

  // El informe también llega a los correos institucionales de todas las
  // IE del grupo (spec del usuario: "llegará a los correos de las IE...
  // que conformaron el grupo"), no solo al responsable de envío.
  var correosIE = obtenerInstitucionesDelGrupo(idGrupo)
    .map(function (ie) { return String(ie.email || "").trim(); })
    .filter(Boolean);

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

  var copias = copiasAsistentes.concat(correosIE, [String(config.COPIAS_CORREO || "").trim()]).filter(Boolean).join(",");

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

/* ------------------------------------------------------------------ *
 * Envío de grupo — recorrido de prueba
 *
 * Envía a personal de la Alcaldía/SEM (nunca a un correo institucional de
 * una IE) el acceso real de un grupo elegido al azar entre los ya
 * cargados en GruposComunal/AccesosGrupo, para que hagan un recorrido de
 * prueba de principio a fin antes de la puesta en producción. El
 * HTML/CSS del cuerpo es el mismo que construirCorreoAccesoIE_ de FEI 3.1
 * (paleta institucional verde #0B6A44 / amarillo #F4B400), adaptado de
 * "IE" a "grupo" — mismo estilo de los correos ya enviados por la SEM.
 * ------------------------------------------------------------------ */

/**
 * URL del deployment publicado ("Foro comunal 1.0"), tomada de `clasp
 * deployments`. Respaldo de ConfiguracionComunal.URL_WEBAPP: la primera
 * vez que se envía un recorrido de prueba, si esa clave sigue vacía se
 * guarda con este valor y se corre actualizarUrlsAcceso() — así el
 * recorrido de prueba (y cualquier acceso enviado después) deja de
 * apuntar al /dev del editor, que un destinatario externo sin permiso de
 * edición sobre el proyecto no puede abrir.
 */
var URL_WEBAPP_PUBLICADA_ = "https://script.google.com/macros/s/AKfycbwpkc8qv90P43dDt-F5NpIBwBkABUc44BGvdFfebtODDZ6b2t1Y_BTAtWB87eJss69G6g/exec";

function _asegurarUrlWebAppConfigurada_() {
  var config = getConfig();
  if (!config.URL_WEBAPP) {
    escribirConfig_("URL_WEBAPP", URL_WEBAPP_PUBLICADA_);
    actualizarUrlsAcceso();
  }
}

/** Direcciones autorizadas para el recorrido de prueba (personal de la Alcaldía, nunca correos de IE). */
var DESTINATARIOS_RECORRIDO_PRUEBA_ = [
  "adriana.cedeno@alcaldianeiva.gov.co",
  "ana.torres@alcaldianeiva.gov.co",
  "carolina.soto@alcaldianeiva.gov.co",
  "edna.rivera@alcaldianeiva.gov.co",
  "nelson.herrera@alcaldianeiva.gov.co",
  "ronald.polania@alcaldianeiva.gov.co",
  "rosa.gonzalez@alcaldianeiva.gov.co",
  "rosario.valenzuela@alcaldianeiva.gov.co"
];

function _construirCorreoRecorridoPrueba_(grupoNombre, codigo, url) {
  var asunto = "🎓 Envío de grupo — recorrido de prueba — Foro Educativo Comunal Neiva 2026 — " + grupoNombre;
  var textoEnlace = "Ingresar como " + grupoNombre + " (recorrido de prueba)";

  var cuerpoTexto =
    "Secretaría de Educación de Neiva\n\n" +
    "Envío de grupo — recorrido de prueba del Foro Educativo Comunal Neiva 2026.\n\n" +
    "Se le asignó al azar el acceso real del " + grupoNombre + " para que recorra la aplicación de " +
    "principio a fin antes de la puesta en producción.\n\n" +
    "Código de acceso: " + codigo + "\n\n" +
    textoEnlace + ":\n" + url + "\n\n" +
    "Este código y este enlace son de uso interno para la prueba: no deben compartirse fuera de la " +
    "Secretaría de Educación.\n\n" +
    "Secretaría de Educación de Neiva\n" +
    "Foro Educativo Comunal Neiva 2026\n" +
    "\"Encuentro de voces que construyen territorio\"";

  var cuerpoHTML =
    "<div style=\"background:#F7F8FA;padding:28px 12px;font-family:Arial,Helvetica,sans-serif;\">" +
    "<div style=\"max-width:520px;margin:0 auto;background:#FFFFFF;border-radius:16px;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,.10);\">" +
    "<div style=\"background:#0B6A44;padding:26px 28px;text-align:center;\">" +
    "<div style=\"color:#FFFFFF;font-size:20px;font-weight:700;\">Foro Educativo Comunal</div>" +
    "<div style=\"color:#CFE8DC;font-size:14px;margin-top:2px;\">Neiva 2026</div>" +
    "</div>" +
    "<div style=\"padding:28px 28px 8px;\">" +
    "<div style=\"display:inline-block;background:#FFF8E1;color:#7A5B00;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;padding:6px 12px;border-radius:20px;margin:0 0 16px;\">🧪 Envío de grupo — recorrido de prueba</div>" +
    "<p style=\"font-size:16px;color:#333333;margin:0 0 14px;\">Estimado(a) colaborador(a) de la Secretaría de Educación de Neiva:</p>" +
    "<p style=\"font-size:15px;color:#4A4A4A;line-height:1.6;margin:0 0 22px;\">" +
    "Le compartimos un acceso de prueba para recorrer el Foro Educativo Comunal Neiva 2026 de principio " +
    "a fin, con los datos reales del <strong>" + grupoNombre + "</strong>, asignado al azar." +
    "</p>" +
    "<div style=\"background:#F7F8FA;border-left:6px solid #F4B400;border-radius:10px;padding:16px 20px;margin:0 0 24px;text-align:center;\">" +
    "<div style=\"font-size:12px;font-weight:700;color:#0B6A44;text-transform:uppercase;letter-spacing:.5px;\">Código de acceso</div>" +
    "<div style=\"font-size:30px;font-weight:700;letter-spacing:6px;color:#0B6A44;margin-top:4px;\">" + codigo + "</div>" +
    "</div>" +
    "<div style=\"text-align:center;margin:0 0 24px;\">" +
    "<a href=\"" + url + "\" target=\"_blank\" style=\"display:inline-block;background:#0B6A44;color:#FFFFFF;text-decoration:none;font-weight:700;font-size:15px;padding:14px 26px;border-radius:10px;\">" + textoEnlace + "</a>" +
    "</div>" +
    "<div style=\"background:#F7F8FA;border:1px dashed #C7CDD1;border-radius:10px;padding:10px 14px;margin:0 0 24px;text-align:center;\">" +
    "<p style=\"font-size:11px;font-weight:700;color:#888888;text-transform:uppercase;letter-spacing:.4px;margin:0 0 4px;\">También puede copiar este enlace</p>" +
    "<p style=\"font-size:12px;color:#0B6A44;word-break:break-all;margin:0;\">" + url + "</p>" +
    "</div>" +
    "<div style=\"background:#FFF8E1;border-left:6px solid #F4B400;border-radius:10px;padding:12px 16px;margin:0 0 20px;\">" +
    "<p style=\"font-size:13px;color:#7A5B00;margin:0;\">🔒 Este código y este enlace son de uso interno para la prueba: no deben compartirse fuera de la Secretaría de Educación.</p>" +
    "</div>" +
    "</div>" +
    "<div style=\"background:#F7F8FA;padding:18px 28px;text-align:center;border-top:1px solid #E5E7EA;\">" +
    "<p style=\"font-size:13px;color:#0B6A44;font-weight:700;margin:0;\">Secretaría de Educación de Neiva</p>" +
    "<p style=\"font-size:12px;color:#888888;margin:4px 0 0;font-style:italic;\">“Encuentro de voces que construyen territorio”</p>" +
    "</div>" +
    "</div>" +
    "</div>";

  return { asunto: asunto, cuerpoTexto: cuerpoTexto, cuerpoHTML: cuerpoHTML };
}

/**
 * Accesos de grupos REALES ya generados (TOKEN+CODIGO), fuente para el
 * sorteo del recorrido de prueba — excluye explícitamente GRUPO-PRUEBA
 * (spec: "grupos al azar reales").
 */
function _accesosGrupoDisponibles_() {
  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  return leerFilasComoObjetos_(hoja).filter(function (f) {
    var idGrupo = String(f.ID_GRUPO || "").trim();
    return idGrupo && idGrupo !== "GRUPO-PRUEBA" && String(f.TOKEN || "").trim() && String(f.CODIGO_ACCESO || "").trim();
  });
}

/**
 * Envía el "recorrido de prueba" a cada una de las 8 direcciones fijas de
 * DESTINATARIOS_RECORRIDO_PRUEBA_, cada una con un grupo real elegido al
 * azar (con reemplazo — hay más destinatarios que grupos). Nunca envía a
 * ningún correo institucional de una IE. Requiere que generarAccesosGrupo()
 * ya se haya ejecutado.
 */
function enviarGrupoRecorridoPrueba() {
  var remitente = remitenteValido_();
  if (!remitente.ok) return remitente;

  _asegurarUrlWebAppConfigurada_();

  var accesos = _accesosGrupoDisponibles_();
  if (!accesos.length) {
    return { ok: false, mensaje: "No hay grupos con acceso generado todavía. Ejecute generarAccesosGrupo() primero." };
  }

  var resultados = DESTINATARIOS_RECORRIDO_PRUEBA_.map(function (destinatario) {
    var acceso = accesos[Math.floor(Math.random() * accesos.length)];
    var grupoNombre = acceso.GRUPO || acceso.ID_GRUPO;
    var correo = _construirCorreoRecorridoPrueba_(grupoNombre, acceso.CODIGO_ACCESO, construirUrlAcceso_(acceso.TOKEN));
    try {
      GmailApp.sendEmail(destinatario, correo.asunto, correo.cuerpoTexto, {
        htmlBody: correo.cuerpoHTML,
        from: remitente.remitente,
        name: "Secretaría de Educación de Neiva"
      });
      return { destinatario: destinatario, grupo: grupoNombre, ok: true };
    } catch (error) {
      return { destinatario: destinatario, grupo: grupoNombre, ok: false, mensaje: error.message };
    }
  });

  Logger.log(JSON.stringify(resultados));
  return { ok: true, resultados: resultados };
}

/**
 * Mismo recorrido de prueba, pero un único envío a la dirección indicada
 * (para una verificación final antes o después del envío masivo).
 */
function enviarGrupoRecorridoPruebaA(destinatario) {
  var remitente = remitenteValido_();
  if (!remitente.ok) return remitente;

  _asegurarUrlWebAppConfigurada_();

  var accesos = _accesosGrupoDisponibles_();
  if (!accesos.length) {
    return { ok: false, mensaje: "No hay grupos con acceso generado todavía. Ejecute generarAccesosGrupo() primero." };
  }

  var acceso = accesos[Math.floor(Math.random() * accesos.length)];
  var grupoNombre = acceso.GRUPO || acceso.ID_GRUPO;
  var correo = _construirCorreoRecorridoPrueba_(grupoNombre, acceso.CODIGO_ACCESO, construirUrlAcceso_(acceso.TOKEN));
  try {
    GmailApp.sendEmail(destinatario, correo.asunto, correo.cuerpoTexto, {
      htmlBody: correo.cuerpoHTML,
      from: remitente.remitente,
      name: "Secretaría de Educación de Neiva"
    });
    return { ok: true, destinatario: destinatario, grupo: grupoNombre };
  } catch (error) {
    return { ok: false, mensaje: error.message };
  }
}

/**
 * Envía el informe por correo, pero solo si el grupo ya valoró el Foro y
 * ya descargó el informe (condiciones explícitas de esta entrega —
 * distinto de generarInformeCompletoGrupo, que solo exige la valoración:
 * aquí además hace falta haberlo descargado, spec de este pedido). Es el
 * punto de entrada del botón "Enviar informe por correo" de la pantalla
 * de informe generado.
 */
function enviarInformeSiCorresponde(idGrupo, tokenSesion, dispositivoId) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  if (!obtenerValoracionGrupo(idGrupo)) {
    return { ok: false, mensaje: "Debe completar la valoración del Foro antes de enviar el informe." };
  }
  var informe = obtenerInformeGrupo(idGrupo);
  if (!informe || String(informe.DESCARGADO || "").toUpperCase() !== "SI") {
    return { ok: false, mensaje: "Debe descargar el informe antes de enviarlo por correo." };
  }
  return enviarInformeGrupo(idGrupo);
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

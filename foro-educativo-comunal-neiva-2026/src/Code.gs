/**
 * Code.gs — Foro Educativo Comunal Neiva 2026
 *
 * Router HTTP único (doGet) + include() para componer Index+CSS+JS en un
 * solo HTML servido — mismo patrón que FEI 3.1 (docs/01-auditoria-fei-3.1.md
 * §1.1), adaptado: el token ahora resuelve un GRUPO, no una IE, y se
 * incluyen 4 archivos de frontend en vez de 2 (JS, Components, Modal,
 * InformeStyles).
 *
 * Este archivo también reúne los "RPC wrappers" que el cliente invoca vía
 * google.script.run — son la única superficie pública del backend. Cada
 * uno está envuelto en ejecutarRpcSeguro_ (Utils.gs) para que cualquier
 * excepción no prevista llegue al cliente como un mensaje legible
 * (r.mensaje) en vez del genérico "Ocurrió un error de comunicación con
 * el servidor" del withFailureHandler.
 */

function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function doGet(e) {
  var params = (e && e.params) || (e && e.parameter) || {};

  // Página pública de asistencia (QR/enlace), sin token/código — mismo
  // patrón que paginaAsistenciaQR_ en FEI 3.1 (auditoría §1.1/§2.2).
  var idGrupoAsistencia = String(params.asistencia || "").trim();
  if (idGrupoAsistencia) {
    return paginaAsistenciaGrupo_(idGrupoAsistencia);
  }

  var token = String(params.t || params.token || params.TOKEN || "").trim();

  var template = HtmlService.createTemplateFromFile("Index");
  template.TOKEN_ACCESO = token;
  template.NOMBRE_GRUPO_ACCESO = "";
  template.ID_GRUPO_ACCESO = "";
  template.NOMBRE_FORO = getConfig().NOMBRE_FORO;
  template.SUBTITULO_FORO = getConfig().SUBTITULO;
  template.LOGO_ENCABEZADO_ID = getConfig().LOGO_ENCABEZADO_ID;
  template.LOGO_PIE_ID = getConfig().LOGO_PIE_ID;
  try {
    asegurarLogosSplashPublicos_();
  } catch (err) {
    Logger.log("doGet: no fue posible asegurar los logos públicos: " + err.message);
  }
  try {
    asegurarLimiteSesionesGrupoRazonable_();
  } catch (err) {
    Logger.log("doGet: no fue posible ajustar el límite de sesiones por grupo: " + err.message);
  }

  if (token) {
    try {
      var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
      var mapa = obtenerMapaCabeceras_(hoja);
      var fila = buscarFilaPorColumna_(hoja, mapa, "TOKEN", token);
      if (fila !== -1) {
        template.NOMBRE_GRUPO_ACCESO = String(hoja.getRange(fila, mapa["GRUPO"]).getValue() || "");
        // Cada enlace de acceso (?t=TOKEN) es exclusivo de un grupo, así que
        // ya se puede acotar el carrusel de bienvenida a las IE de ESE
        // grupo, incluso antes de que el visitante ingrese el código.
        template.ID_GRUPO_ACCESO = String(hoja.getRange(fila, mapa["ID_GRUPO"]).getValue() || "");
      }
    } catch (err) {
      Logger.log("doGet: no fue posible resolver el grupo por token: " + err.message);
    }
  }

  return template
    .evaluate()
    .setTitle(getConfig().NOMBRE_FORO)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------------------------------------------ *
 * RPC — Acceso y sesión
 * ------------------------------------------------------------------ */

function rpcValidarAcceso(token, codigo, dispositivoId, forzar) {
  return ejecutarRpcSeguro_(function () {
    return validarAccesoGrupo(token, codigo, dispositivoId, forzar);
  });
}

function rpcMantenerSesion(idGrupo, dispositivoId, tokenSesion) {
  return ejecutarRpcSeguro_(function () {
    return mantenerSesionGrupo(idGrupo, dispositivoId, tokenSesion);
  });
}

function rpcLiberarSesion(idGrupo, dispositivoId, tokenSesion) {
  return ejecutarRpcSeguro_(function () {
    return liberarSesionGrupo_(idGrupo, dispositivoId, tokenSesion);
  });
}

function rpcTransferirPrincipal(idGrupo, dispositivoId, tokenSesion) {
  return ejecutarRpcSeguro_(function () {
    return transferirResponsablePrincipalGrupo(idGrupo, dispositivoId, tokenSesion);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Participación / firmantes
 * ------------------------------------------------------------------ */

function rpcRegistrarParticipante(idGrupo, idIE, nombre, estamento, rolForo, correo, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return registrarParticipante(idGrupo, idIE, nombre, estamento, rolForo, correo, dispositivoId);
  });
}

function rpcEstadoFirmantes(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return { total: contarParticipantesGrupo(idGrupo), firmantes: listarFirmantesGrupo(idGrupo).slice(0, 50) };
  });
}

function rpcMatrizParticipacion(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerMatrizParticipacionGrupo(idGrupo);
  });
}

/** Conteo manual de participantes por estamento e IE (ParticipacionEstamento.gs — mismo formato que Participación de 3.1). */
function rpcObtenerParticipacionEstamento(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerParticipacionEstamentoGrupo(idGrupo);
  });
}

function rpcGuardarParticipacionEstamentoIE(idGrupo, tokenSesion, dispositivoId, idIE, valores) {
  return ejecutarRpcSeguro_(function () {
    return guardarParticipacionEstamentoIE(idGrupo, tokenSesion, dispositivoId, idIE, valores);
  });
}

function rpcRolesForo() {
  return ejecutarRpcSeguro_(function () {
    return ROLES_FORO_;
  });
}

/** Rector(a) editable en la ficha de Confirmación de caracterización (igual que en FEI 3.1). */
function rpcActualizarRectorIE(idGrupo, tokenSesion, dispositivoId, idIE, nombreRector) {
  return ejecutarRpcSeguro_(function () {
    return actualizarRectorIE(idGrupo, tokenSesion, dispositivoId, idIE, nombreRector);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Sesión de preparación (pre-socialización, Preparacion.gs)
 * ------------------------------------------------------------------ */
function rpcObtenerPreparacionIE(idGrupo, idIE) {
  return ejecutarRpcSeguro_(function () {
    return obtenerPreparacionIE(idGrupo, idIE);
  });
}

function rpcGuardarPreparacionIE(idGrupo, tokenSesion, dispositivoId, idIE, responsable, respuestas) {
  return ejecutarRpcSeguro_(function () {
    return guardarPreparacionIE(idGrupo, tokenSesion, dispositivoId, idIE, responsable, respuestas);
  });
}

function rpcMarcarPreparacionEnviada(idGrupo, tokenSesion, dispositivoId, idIE) {
  return ejecutarRpcSeguro_(function () {
    return marcarPreparacionEnviada(idGrupo, tokenSesion, dispositivoId, idIE);
  });
}

/** Aportes de preparación ya enviados por todas las IE del grupo, para la sección de socialización de Sesión 1. */
function rpcPreparacionesEnviadasGrupo(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerPreparacionesEnviadasGrupo(idGrupo);
  });
}

/** Todos los aportes de invitados del grupo, agrupados por IE — listado agregado en Selección de IE (Invitados.gs). */
function rpcAportesInvitadosGrupo(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerAportesInvitadosGrupo(idGrupo);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Acceso de invitado (estudiante/acudiente, sin código de acceso)
 * ------------------------------------------------------------------ */
function rpcIniciarAccesoInvitado(idIE, tipoInvitado, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return iniciarAccesoInvitado(idIE, tipoInvitado, dispositivoId);
  });
}

function rpcObtenerPreparacionIEInvitado(tokenInvitado, idIE, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return obtenerPreparacionIEInvitado(tokenInvitado, idIE, dispositivoId);
  });
}

function rpcGuardarCaracterizacionInvitado(tokenInvitado, idIE, dispositivoId, datos) {
  return ejecutarRpcSeguro_(function () {
    return guardarCaracterizacionInvitado(tokenInvitado, idIE, dispositivoId, datos);
  });
}

function rpcGuardarPreparacionIEInvitado(tokenInvitado, idIE, tipoInvitado, dispositivoId, respuestas) {
  return ejecutarRpcSeguro_(function () {
    return guardarPreparacionIEInvitado(tokenInvitado, idIE, tipoInvitado, dispositivoId, respuestas);
  });
}

function rpcMarcarPreparacionEnviadaInvitado(tokenInvitado, idIE, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return marcarPreparacionEnviadaInvitado(tokenInvitado, idIE, dispositivoId);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Responsable de envío y asistentes de envío
 * ------------------------------------------------------------------ */

function rpcGuardarResponsable(idGrupo, tokenSesion, dispositivoId, tipo, datos) {
  return ejecutarRpcSeguro_(function () {
    return guardarResponsableEnvio(idGrupo, tokenSesion, dispositivoId, tipo, datos);
  });
}

function rpcListarResponsables(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return listarResponsablesEnvio(idGrupo);
  });
}

function rpcEliminarResponsable(idGrupo, idRegistro, tokenSesion, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return eliminarResponsableEnvio(idGrupo, idRegistro, tokenSesion, dispositivoId);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Consentimiento informado del grupo (sección 3.2)
 * ------------------------------------------------------------------ */

function rpcGuardarConsentimientoGrupo(idGrupo, tokenSesion, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return guardarConsentimientoGrupo(idGrupo, tokenSesion, dispositivoId);
  });
}

function rpcGuardarUltimaPantalla(idGrupo, idPantalla) {
  return ejecutarRpcSeguro_(function () {
    return guardarUltimaPantallaGrupo(idGrupo, idPantalla);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Método de asistencia (QR/enlace o listado físico + foto)
 * ------------------------------------------------------------------ */

function rpcGuardarMetodoAsistencia(idGrupo, tokenSesion, dispositivoId, metodo) {
  return ejecutarRpcSeguro_(function () {
    return guardarMetodoAsistencia(idGrupo, tokenSesion, dispositivoId, metodo);
  });
}

function rpcUrlAsistencia(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return construirUrlAsistencia_(idGrupo);
  });
}

function rpcSubirListadoAsistencia(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType) {
  return ejecutarRpcSeguro_(function () {
    return subirListadoAsistencia(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType);
  });
}

function rpcSubirFotoEvidencia(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType) {
  return ejecutarRpcSeguro_(function () {
    return subirFotoEvidencia(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType);
  });
}

/** Foto general del grupo (Participación) — independiente del método de asistencia elegido. */
function rpcSubirFotoGrupo(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType) {
  return ejecutarRpcSeguro_(function () {
    return subirFotoGrupo(idGrupo, tokenSesion, dispositivoId, datosBase64, nombreArchivo, mimeType);
  });
}

/** Usada desde AsistenciaPublica.html (página pública QR, sin token/código). */
function rpcRegistrarAsistenciaPublica(idGrupo, idIE, nombre, estamento, correo) {
  return ejecutarRpcSeguro_(function () {
    return registrarAsistenciaPublica(idGrupo, idIE, nombre, estamento, correo);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Sesión 1
 * ------------------------------------------------------------------ */

function rpcGuardarSesion1(idGrupo, tokenSesion, dispositivoId, campos) {
  return ejecutarRpcSeguro_(function () {
    return guardarSesion1(idGrupo, tokenSesion, dispositivoId, campos);
  });
}

function rpcObtenerSesion1(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerSesion1(idGrupo);
  });
}

function rpcEnviarSesion1(idGrupo, tokenSesion, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return enviarSesion1Definitiva(idGrupo, tokenSesion, dispositivoId);
  });
}

/** Archivos descargables necesarios antes de iniciar Sesión 1 (Recursos.gs). */
function rpcObtenerRecursosSesion1(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerRecursosSesion1(idGrupo);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — ConectaEduca (Sesión 2)
 * ------------------------------------------------------------------ */

function rpcGuardarActorConectaEduca(idGrupo, tokenSesion, dispositivoId, registro) {
  return ejecutarRpcSeguro_(function () {
    return guardarActorConectaEduca(idGrupo, tokenSesion, dispositivoId, registro);
  });
}

function rpcListarConectaEduca(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return listarConectaEduca(idGrupo);
  });
}

function rpcEliminarActorConectaEduca(idGrupo, idRegistro, tokenSesion, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return eliminarActorConectaEduca(idGrupo, idRegistro, tokenSesion, dispositivoId);
  });
}

function rpcEnviarSesion2(idGrupo, tokenSesion, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return enviarSesion2Definitiva(idGrupo, tokenSesion, dispositivoId);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Informe / cierre
 * ------------------------------------------------------------------ */

/**
 * Generar el informe ya NO envía el correo automáticamente (a diferencia
 * de antes): el envío es ahora una acción explícita y separada
 * (rpcEnviarInformePorCorreo), habilitada solo tras valorar y descargar
 * el informe. generarInformeCompletoGrupo ya exige la valoración antes
 * de generar (Grupos.gs).
 */
function rpcGenerarInforme(idGrupo, tokenSesion, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return generarInformeCompletoGrupo(idGrupo, tokenSesion, dispositivoId);
  });
}

function rpcObtenerInforme(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerInformeGrupo(idGrupo);
  });
}

/** Se llama al hacer clic en el enlace de descarga del informe (habilita el envío por correo). */
function rpcMarcarInformeDescargado(idGrupo, tokenSesion, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return marcarInformeDescargado(idGrupo, tokenSesion, dispositivoId);
  });
}

/** Envía el informe por correo — exige valoración enviada y descarga registrada (Correo.gs). */
function rpcEnviarInformePorCorreo(idGrupo, tokenSesion, dispositivoId) {
  return ejecutarRpcSeguro_(function () {
    return enviarInformeSiCorresponde(idGrupo, tokenSesion, dispositivoId);
  });
}

function rpcEstadoGrupo(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerEstadoGrupo(idGrupo);
  });
}

/* ------------------------------------------------------------------ *
 * RPC — Valoración del Foro (condición para generar el informe)
 * ------------------------------------------------------------------ */

function rpcGuardarValoracion(idGrupo, tokenSesion, dispositivoId, respuestas) {
  return ejecutarRpcSeguro_(function () {
    return guardarValoracionGrupo(idGrupo, tokenSesion, dispositivoId, respuestas);
  });
}

/** Valoración pública, sin código de acceso — ver comentario en Valoracion.gs. */
function rpcGuardarValoracionAsistentePublica(idGrupo, respuestas) {
  return ejecutarRpcSeguro_(function () {
    return guardarValoracionAsistentePublica(idGrupo, respuestas);
  });
}

function rpcObtenerValoracion(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerValoracionGrupo(idGrupo);
  });
}

/**
 * Todas las IE activas de todos los grupos — respaldo del carrusel de
 * bienvenida solo para un enlace de acceso genérico, sin token de grupo.
 */
function rpcTodasLasInstituciones() {
  return ejecutarRpcSeguro_(function () {
    return obtenerTodasLasInstitucionesActivas();
  });
}

/**
 * IE activas de un solo grupo — usada por el carrusel de bienvenida
 * cuando el enlace de acceso (?t=TOKEN) ya identifica un grupo concreto,
 * para no mostrar las IE de los demás grupos.
 */
function rpcInstitucionesDelGrupo(idGrupo) {
  return ejecutarRpcSeguro_(function () {
    return obtenerInstitucionesDelGrupo(idGrupo);
  });
}

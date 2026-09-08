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
 * google.script.run — son la única superficie pública del backend.
 */

function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function doGet(e) {
  var params = (e && e.params) || (e && e.parameter) || {};
  var token = String(params.t || params.token || params.TOKEN || "").trim();

  var template = HtmlService.createTemplateFromFile("Index");
  template.TOKEN_ACCESO = token;
  template.NOMBRE_GRUPO_ACCESO = "";
  template.NOMBRE_FORO = getConfig().NOMBRE_FORO;
  template.SUBTITULO_FORO = getConfig().SUBTITULO;

  if (token) {
    try {
      var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
      var mapa = obtenerMapaCabeceras_(hoja);
      var fila = buscarFilaPorColumna_(hoja, mapa, "TOKEN", token);
      if (fila !== -1) {
        template.NOMBRE_GRUPO_ACCESO = String(hoja.getRange(fila, mapa["GRUPO"]).getValue() || "");
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
  return validarAccesoGrupo(token, codigo, dispositivoId, forzar);
}

function rpcMantenerSesion(idGrupo, dispositivoId, tokenSesion) {
  return mantenerSesionGrupo(idGrupo, dispositivoId, tokenSesion);
}

function rpcLiberarSesion(idGrupo, dispositivoId, tokenSesion) {
  return liberarSesionGrupo_(idGrupo, dispositivoId, tokenSesion);
}

function rpcTransferirPrincipal(idGrupo, dispositivoId, tokenSesion) {
  return transferirResponsablePrincipalGrupo(idGrupo, dispositivoId, tokenSesion);
}

/* ------------------------------------------------------------------ *
 * RPC — Participación / firmantes
 * ------------------------------------------------------------------ */

function rpcRegistrarParticipante(idGrupo, idIE, nombre, rol, correo, dispositivoId) {
  return registrarParticipante(idGrupo, idIE, nombre, rol, correo, dispositivoId);
}

function rpcEstadoFirmantes(idGrupo) {
  return { total: contarParticipantesGrupo(idGrupo), firmantes: listarFirmantesGrupo(idGrupo).slice(0, 50) };
}

/* ------------------------------------------------------------------ *
 * RPC — Sesión 1
 * ------------------------------------------------------------------ */

function rpcGuardarSesion1(idGrupo, tokenSesion, dispositivoId, campos) {
  return guardarSesion1(idGrupo, tokenSesion, dispositivoId, campos);
}

function rpcObtenerSesion1(idGrupo) {
  return obtenerSesion1(idGrupo);
}

function rpcEnviarSesion1(idGrupo, tokenSesion, dispositivoId) {
  return enviarSesion1Definitiva(idGrupo, tokenSesion, dispositivoId);
}

/* ------------------------------------------------------------------ *
 * RPC — ConectaEduca (Sesión 2)
 * ------------------------------------------------------------------ */

function rpcGuardarActorConectaEduca(idGrupo, tokenSesion, dispositivoId, registro) {
  return guardarActorConectaEduca(idGrupo, tokenSesion, dispositivoId, registro);
}

function rpcListarConectaEduca(idGrupo) {
  return listarConectaEduca(idGrupo);
}

function rpcEliminarActorConectaEduca(idGrupo, idRegistro, tokenSesion, dispositivoId) {
  return eliminarActorConectaEduca(idGrupo, idRegistro, tokenSesion, dispositivoId);
}

function rpcEnviarSesion2(idGrupo, tokenSesion, dispositivoId) {
  return enviarSesion2Definitiva(idGrupo, tokenSesion, dispositivoId);
}

/* ------------------------------------------------------------------ *
 * RPC — Informe / cierre
 * ------------------------------------------------------------------ */

function rpcGenerarInforme(idGrupo, tokenSesion, dispositivoId) {
  var resultado = generarInformeCompletoGrupo(idGrupo, tokenSesion, dispositivoId);
  if (resultado.ok) {
    try {
      enviarInformeGrupo(idGrupo);
    } catch (e) {
      Logger.log("El informe se generó, pero el envío por correo falló: " + e.message);
    }
  }
  return resultado;
}

function rpcObtenerInforme(idGrupo) {
  return obtenerInformeGrupo(idGrupo);
}

function rpcEstadoGrupo(idGrupo) {
  return obtenerEstadoGrupo(idGrupo);
}

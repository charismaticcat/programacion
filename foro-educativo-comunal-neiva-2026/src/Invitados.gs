/**
 * Invitados.gs — Foro Educativo Comunal Neiva 2026
 *
 * Acceso de invitado (estudiante o padre/madre de familia — acudiente):
 * NO requiere el código de acceso del grupo (spec: "crear perfil de
 * estudiante y de padre de familia (no necesita código) solo accede
 * dando click en soy invitado"). Elige su institución educativa (de
 * todas las IE reales, no solo las de un grupo) y su tipo de invitado, y
 * entra DIRECTAMENTE a la Sesión de preparación de esa IE para enviar
 * sus aportes; una vez enviados, no tiene más acceso a la aplicación
 * (spec: "envía aportes y no tiene más acceso").
 *
 * Es una sesión de privilegio mínimo y de un solo uso: el token de
 * invitado NUNCA pasa por sesionActivaPorIdGrupo_ (el mecanismo de
 * sesión completa del grupo, ligado al código de acceso) — solo
 * autoriza, mediante sesionInvitadoValida_(), guardar y enviar la
 * preparación de la IE concreta que el invitado eligió al entrar. No da
 * acceso a Participación, Sesión 1, Sesión 2, ni a ninguna otra escritura
 * del grupo.
 */

var HOJA_INVITADOS_ = "InvitadosPreparacion";

function cabecerasInvitadosPreparacion_() {
  return ["TOKEN_INVITADO", "ID_GRUPO", "ID_IE", "TIPO_INVITADO", "DISPOSITIVO_ID", "FECHA_INGRESO", "ENVIADO"];
}

/**
 * Inicia una sesión de invitado para una IE concreta — sin código de
 * acceso. `tipoInvitado` es "ESTUDIANTE" o "ACUDIENTE". El grupo al que
 * pertenece la IE se resuelve automáticamente (obtenerGrupoDeInstitucion_,
 * Instituciones.gs), igual que ya hace el resto de la app.
 */
function iniciarAccesoInvitado(idIE, tipoInvitado, dispositivoId) {
  idIE = String(idIE || "").trim();
  tipoInvitado = String(tipoInvitado || "").trim().toUpperCase();
  if (tipoInvitado !== "ESTUDIANTE" && tipoInvitado !== "ACUDIENTE") {
    return { ok: false, mensaje: "Selecciona si eres estudiante o acudiente." };
  }
  if (!idIE) return { ok: false, mensaje: "Selecciona tu institución educativa." };

  var grupoInfo = obtenerGrupoDeInstitucion_(idIE);
  if (!grupoInfo) return { ok: false, mensaje: "No se encontró esa institución educativa." };

  var instituciones = obtenerInstitucionesDelGrupo(grupoInfo.idGrupo);
  var ie = instituciones.find(function (i) { return i.idIE === idIE; });
  if (!ie) return { ok: false, mensaje: "Esa institución no está activa en ningún grupo todavía." };

  var tokenInvitado = Utilities.getUuid();
  var hoja = obtenerHoja_(HOJA_INVITADOS_, cabecerasInvitadosPreparacion_());
  hoja.appendRow([
    tokenInvitado, grupoInfo.idGrupo, idIE, tipoInvitado, String(dispositivoId || "").trim(), new Date(), "NO"
  ]);

  return {
    ok: true,
    tokenInvitado: tokenInvitado,
    idGrupo: grupoInfo.idGrupo,
    grupo: grupoInfo.grupo,
    idIE: idIE,
    institucion: ie.institucion,
    logoId: ie.logoId || ""
  };
}

/** Verifica que el token de invitado corresponda a esa IE y a ese dispositivo. Devuelve la fila o null. */
function sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId) {
  var hoja = obtenerHoja_(HOJA_INVITADOS_, cabecerasInvitadosPreparacion_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "TOKEN_INVITADO", String(tokenInvitado || "").trim());
  if (fila === -1) return null;
  var obj = leerFilaComoObjeto_(hoja, fila, mapa);
  if (String(obj.ID_IE || "").trim() !== String(idIE || "").trim()) return null;
  if (String(obj.DISPOSITIVO_ID || "").trim() !== String(dispositivoId || "").trim()) return null;
  return obj;
}

/** Estado de preparación de la IE del invitado — misma información que obtenerPreparacionIE, autenticada distinto. */
function obtenerPreparacionIEInvitado(tokenInvitado, idIE, dispositivoId) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return null;
  return obtenerPreparacionIE(sesion.ID_GRUPO, idIE);
}

/** Guarda las respuestas de preparación del invitado — mismo almacenamiento que Preparacion.gs, autenticación de invitado. */
function guardarPreparacionIEInvitado(tokenInvitado, idIE, tipoInvitado, dispositivoId, respuestas) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión de invitado ya no es válida." };
  var etiqueta = String(sesion.TIPO_INVITADO || tipoInvitado || "").toUpperCase() === "ESTUDIANTE"
    ? "Un(a) estudiante invitado(a)"
    : "Un(a) acudiente invitado(a)";
  return _guardarPreparacionIEInterno_(sesion.ID_GRUPO, idIE, etiqueta, respuestas);
}

/**
 * Envío definitivo de un invitado: marca la preparación como enviada y
 * cierra su sesión de invitado (ENVIADO=SI en InvitadosPreparacion) — a
 * partir de aquí sesionInvitadoValida_ lo sigue reconociendo (para poder
 * mostrarle la confirmación), pero cualquier intento de volver a guardar
 * o enviar debe rechazarse: ya cumplió su única tarea.
 */
function marcarPreparacionEnviadaInvitado(tokenInvitado, idIE, dispositivoId) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión de invitado ya no es válida." };
  if (String(sesion.ENVIADO || "") === "SI") {
    return { ok: false, mensaje: "Ya enviaste tus aportes con esta sesión de invitado." };
  }
  var resultado = _marcarPreparacionEnviadaInterno_(sesion.ID_GRUPO, idIE);
  if (resultado.ok) {
    upsertFila_(HOJA_INVITADOS_, cabecerasInvitadosPreparacion_(), "TOKEN_INVITADO", String(tokenInvitado || "").trim(), {
      ENVIADO: "SI"
    });
  }
  return resultado;
}

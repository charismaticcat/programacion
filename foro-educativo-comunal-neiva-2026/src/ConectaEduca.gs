/**
 * ConectaEduca.gs — Foro Educativo Comunal Neiva 2026
 *
 * Sesión 2: ConectaEduca (spec sección 11). A diferencia de Sesión 1 (un
 * único registro fusionado por grupo), aquí cada actor/entidad es una fila
 * independiente — se puede registrar más de un actor por grupo, por eso es
 * *append* + edición/borrado de la fila propia, no UPSERT de un único
 * registro. La información pertenece al GRUPO (spec sección 11: "La
 * información pertenece al GRUPO").
 */

var HOJA_CONECTAEDUCA_ = "ConectaEduca";

function cabecerasConectaEduca_() {
  return [
    "ID_REGISTRO", "ID_GRUPO", "ACTOR", "TIPO_ACTOR", "AREA", "NECESIDADES_ARTICULACION",
    "OPORTUNIDAD", "ALIANZA", "IE_INTERESADAS", "CONEXIONES", "OBSERVACIONES", "FECHA"
  ];
}

/** Registra un nuevo actor/entidad de ConectaEduca para el grupo. */
function guardarActorConectaEduca(idGrupo, tokenSesion, dispositivoId, registro) {
  idGrupo = String(idGrupo || "").trim();
  if (!idGrupo) return { ok: false, mensaje: "Falta el grupo." };
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  registro = registro || {};
  var actor = String(registro.ACTOR || "").trim();
  if (!actor) return { ok: false, mensaje: "El nombre del actor/entidad es obligatorio." };

  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_CONECTAEDUCA_, cabecerasConectaEduca_());
    var idRegistro = Utilities.getUuid();
    hoja.appendRow([
      idRegistro, idGrupo, actor,
      String(registro.TIPO_ACTOR || "").trim(),
      String(registro.AREA || "").trim(),
      String(registro.NECESIDADES_ARTICULACION || "").trim(),
      String(registro.OPORTUNIDAD || "").trim(),
      String(registro.ALIANZA || "").trim(),
      String(registro.IE_INTERESADAS || "").trim(),
      String(registro.CONEXIONES || "").trim(),
      String(registro.OBSERVACIONES || "").trim(),
      new Date()
    ]);
    return { ok: true, idRegistro: idRegistro };
  }, 10000);
}

/** Lista los actores de ConectaEduca registrados por el grupo. */
function listarConectaEduca(idGrupo) {
  var hoja = obtenerHoja_(HOJA_CONECTAEDUCA_, cabecerasConectaEduca_());
  var filas = leerFilasComoObjetos_(hoja);
  var objetivo = String(idGrupo || "").trim();
  return filas.filter(function (f) {
    return String(f.ID_GRUPO || "").trim() === objetivo;
  });
}

/** Elimina un registro propio de ConectaEduca (corrección de un error de captura). */
function eliminarActorConectaEduca(idGrupo, idRegistro, tokenSesion, dispositivoId) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_CONECTAEDUCA_, cabecerasConectaEduca_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_REGISTRO", String(idRegistro || "").trim());
    if (fila === -1) return { ok: false, mensaje: "El registro ya no existe." };
    if (String(hoja.getRange(fila, mapa["ID_GRUPO"]).getValue()).trim() !== idGrupo) {
      return { ok: false, mensaje: "Ese registro no pertenece a este grupo." };
    }
    hoja.deleteRow(fila);
    return { ok: true };
  }, 10000);
}

/**
 * Marca Sesión 2 (ConectaEduca) como enviada definitivamente. Los 4
 * campos de la segunda parte de ConectaEduca (PRIORIDADES_CE,
 * ACUERDOS_CE, PROPUESTAS_CE, RUTA_CE) se guardan, junto con el resto de
 * "consolidados de grupo", en Sesion1Comunal (mismo mecanismo genérico
 * de UPSERT-por-grupo con fusión — ver Sesion1.gs), así que se validan
 * aquí leyendo obtenerSesion1().
 */
function enviarSesion2Definitiva(idGrupo, tokenSesion, dispositivoId) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  // Cualquiera con sesión activa puede enviar (spec del usuario), pero
  // una vez enviada, no se permiten más envíos.
  if (obtenerEstadoGrupo(idGrupo).sesion2Enviada) {
    return { ok: false, codigo: "YA_ENVIADO", mensaje: "Sesión 2 (ConectaEduca) ya fue enviada de forma definitiva. No se permiten más envíos." };
  }

  var datos = obtenerSesion1(idGrupo);
  var vacios = CAMPOS_SESION2_OBLIGATORIOS_.filter(function (c) {
    return !datos || !String(datos[c] || "").trim();
  });
  if (vacios.length) {
    return { ok: false, mensaje: "Faltan campos por completar en ConectaEduca: " + vacios.join(", ") };
  }
  var fueraDeRango = _validarRangoPalabrasSesion1_(datos, CAMPOS_SESION2_OBLIGATORIOS_);
  if (fueraDeRango.length) {
    return {
      ok: false,
      mensaje: "Estos campos deben tener entre " + MIN_PALABRAS_SESION1_ + " y " + MAX_PALABRAS_SESION1_ +
        " palabras: " + fueraDeRango.join(", ")
    };
  }

  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", idGrupo);
    if (fila === -1) return { ok: false, mensaje: "No existe acceso para este grupo." };
    hoja.getRange(fila, mapa["SESION2_ENVIADA"]).setValue("SI");
    hoja.getRange(fila, mapa["FECHA_ENVIO_S2"]).setValue(new Date());
    hoja.getRange(fila, mapa["ESTADO"]).setValue("SESION2_ENVIADA");
    return { ok: true };
  }, 15000);
}

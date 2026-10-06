/**
 * Valoracion.gs — Foro Educativo Comunal Neiva 2026
 *
 * La valoración del Foro DENTRO de la app que gateaba el informe (4
 * preguntas de satisfacción, condición para generar el informe) se
 * eliminó por completo del proyecto en su momento — spec del usuario:
 * "Eliminar valoración". La hoja ValoracionComunal de esa época queda
 * intacta como registro histórico, sin relación con lo de abajo.
 *
 * También queda aquí, sin cambios, la valoración PÚBLICA y ANÓNIMA de la
 * página de asistencia QR (AsistenciaPublica.html).
 *
 * Lo nuevo (tarjeta "Valoración del evento", Participación): digitalización
 * del formato oficial D02.02.F03 ("Evaluación de Encuentro Voces que
 * construyen territorio" / "...Conecta Educa") — ver
 * obtenerValoracionEvento/guardarValoracionEvento más abajo. No gatea nada
 * (igual que el resto de Participación), es independiente de la
 * valoración pública del QR y de la antigua ValoracionComunal.
 */

/**
 * Valoración pública (sin código de acceso) — 4 preguntas de satisfacción
 * diligenciadas por cualquier asistente desde la página pública de
 * asistencia QR (spec del usuario: "La valoración del foro debe aparecer
 * en la asistencia del código QR con las mismas preguntas"). Cada
 * asistente que la diligencia deja su propia fila — no hay upsert ni
 * clave única, es anónima y opcional.
 */
var HOJA_VALORACION_ASISTENTES_PUBLICA_ = "ValoracionAsistentesPublica";

function cabecerasValoracionAsistentesPublica_() {
  return [
    "ID_GRUPO", "P1", "P2", "P3", "P4",
    "P1_MEJORA", "P2_MEJORA", "P3_MEJORA", "P4_MEJORA", "FECHA"
  ];
}

/**
 * Las 4 respuestas de "mejora" (mejoraP1..mejoraP4) son de texto libre y
 * opcionales — en el cliente (AsistenciaPublica.html) solo se despliega
 * el textarea correspondiente cuando la respuesta es Malo o Deficiente
 * (mismo criterio que inicializarValoracion en JS.html/Index.html), pero
 * aquí se guarda lo que llegue sin volver a exigir esa condición: es
 * anónimo y no tiene sentido rechazar un envío por un campo opcional.
 */
function guardarValoracionAsistentePublica(idGrupo, respuestas) {
  idGrupo = String(idGrupo || "").trim();
  if (!idGrupo) return { ok: false, mensaje: "Falta el grupo." };
  respuestas = respuestas || {};
  var p1 = Number(respuestas.p1), p2 = Number(respuestas.p2), p3 = Number(respuestas.p3), p4 = Number(respuestas.p4);
  if (!p1 || !p2 || !p3 || !p4 || p1 < 1 || p1 > 5 || p2 < 1 || p2 > 5 || p3 < 1 || p3 > 5 || p4 < 1 || p4 > 5) {
    return { ok: false, mensaje: "Seleccione una opción de 1 a 5 en las cuatro preguntas antes de enviar." };
  }
  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_VALORACION_ASISTENTES_PUBLICA_, cabecerasValoracionAsistentesPublica_());
    hoja.appendRow([
      idGrupo, p1, p2, p3, p4,
      String(respuestas.mejoraP1 || "").trim(),
      String(respuestas.mejoraP2 || "").trim(),
      String(respuestas.mejoraP3 || "").trim(),
      String(respuestas.mejoraP4 || "").trim(),
      new Date()
    ]);
    return { ok: true };
  }, 10000);
}

/**
 * Valoración del evento (Participación, tarjeta "Valoración del evento") —
 * el responsable de envío digitaliza, una vez por grupo y por sección
 * (ENCUENTRO/CONECTAEDUCA, puede haber una de cada una), la calificación
 * PROMEDIO que resulta de la muestra de encuestas en papel del formato
 * oficial D02.02.F03 (ambos documentos, "...VOCES QUE CONSTRUYEN
 * TERRITORIO" y "...CONECTA EDUCA") que ya escaneó y subió como PDF (ver
 * subirEncuestaSatisfaccion, Asistencia.gs) — spec del usuario: "una sola
 * casilla de calificación promedio de acuerdo a la hoja que se subirá y
 * que manualmente el responsable de envío escribirá". Escala de 0.5 en
 * 0.5, de 1.0 a 5.0 (no el 1-5 entero de las 11 preguntas individuales del
 * formato, que no se digitalizan una por una). Más "Recomendaciones y
 * sugerencias" de texto libre, igual que al final de ambos formatos.
 */
var HOJA_VALORACION_EVENTO_ = "ValoracionEventoComunal";
var OPCIONES_PROMEDIO_VALORACION_EVENTO_ = ["1.0", "1.5", "2.0", "2.5", "3.0", "3.5", "4.0", "4.5", "5.0"];

function cabecerasValoracionEventoComunal_() {
  return ["CLAVE", "ID_GRUPO", "SECCION", "PROMEDIO", "RECOMENDACIONES", "ULTIMA_ACTUALIZACION"];
}

function _claveValoracionEvento_(idGrupo, seccion) {
  return String(idGrupo || "").trim() + "|" + String(seccion || "").trim();
}

/** Valores ya guardados (si los hay) de la valoración del evento de este grupo y sección, listos para pintar el formulario. */
function obtenerValoracionEvento(idGrupo, seccion) {
  var hoja = obtenerHoja_(HOJA_VALORACION_EVENTO_, cabecerasValoracionEventoComunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "CLAVE", _claveValoracionEvento_(idGrupo, seccion));
  return {
    opciones: OPCIONES_PROMEDIO_VALORACION_EVENTO_,
    promedio: fila === -1 ? "" : String(hoja.getRange(fila, mapa["PROMEDIO"]).getValue() || ""),
    recomendaciones: fila === -1 ? "" : String(hoja.getRange(fila, mapa["RECOMENDACIONES"]).getValue() || "")
  };
}

/**
 * Guarda (autoguardado, UPSERT) la valoración del evento — `valores` trae
 * PROMEDIO (opcional, 1.0-5.0 en pasos de 0.5) y/o RECOMENDACIONES
 * (opcional, texto libre); se fusiona con lo ya guardado, igual que el
 * resto de autoguardados de la app.
 */
function guardarValoracionEvento(idGrupo, tokenSesion, dispositivoId, seccion, valores) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  seccion = String(seccion || "").toUpperCase();
  if (seccion !== "ENCUENTRO" && seccion !== "CONECTAEDUCA") {
    return { ok: false, mensaje: "Sección no reconocida." };
  }
  valores = valores || {};
  var datos = { ID_GRUPO: idGrupo, SECCION: seccion, ULTIMA_ACTUALIZACION: new Date() };
  if (Object.prototype.hasOwnProperty.call(valores, "PROMEDIO")) {
    var promedio = Number(valores.PROMEDIO);
    // Pasos de 0.5: el doble debe ser un entero (1.0->2, 1.5->3, ...).
    if (promedio && promedio >= 1 && promedio <= 5 && Math.round(promedio * 2) === promedio * 2) {
      datos.PROMEDIO = promedio;
    }
  }
  if (Object.prototype.hasOwnProperty.call(valores, "RECOMENDACIONES")) {
    datos.RECOMENDACIONES = String(valores.RECOMENDACIONES || "").trim();
  }

  return conLock_(function () {
    upsertFila_(
      HOJA_VALORACION_EVENTO_,
      cabecerasValoracionEventoComunal_(),
      "CLAVE",
      _claveValoracionEvento_(idGrupo, seccion),
      datos
    );
    return { ok: true };
  }, 10000);
}

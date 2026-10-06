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
 * Valoración del evento (Participación, tarjeta "Valoración del evento",
 * lote del formato oficial D02.02.F03) — el funcionario digitaliza, una
 * vez por grupo y por sección (ENCUENTRO/CONECTAEDUCA, puede haber una de
 * cada una), el resultado representativo de la muestra de encuestas en
 * papel que ya escaneó y subió como PDF (ver subirEncuestaSatisfaccion,
 * Asistencia.gs). Mismos 11 aspectos del formato oficial, idénticos en
 * ambos documentos de origen:
 *   - "D02.02_F03_v3 EVALUACION DE VOCES QUE CONSTRUYEN TERRITORIO.docx"
 *   - "D02.02_F03_v3 EVALUACION DE CONECTA EDUCA.docx"
 * (el documento trae una fila duplicada de "Se llegó a conclusiones o
 * acuerdos concretos." — se dejó una sola vez aquí). Escala 1-5
 * (Malo=1 … Excelente=5), más "Recomendaciones y sugerencias" de texto
 * libre, tal como aparece al final de ambos formatos.
 */
var HOJA_VALORACION_EVENTO_ = "ValoracionEventoComunal";
var ASPECTOS_VALORACION_EVENTO_ = [
  { clave: "AGENDA_TIEMPOS", etiqueta: "La agenda se cumplió en los tiempos previstos." },
  { clave: "PARTICIPACION_ACTIVA", etiqueta: "Hubo participación activa con preguntas, aportes y propuestas." },
  { clave: "TEMAS_NECESIDADES_REALES", etiqueta: "Los temas tratados respondían a necesidades reales de la comunidad." },
  { clave: "METODOLOGIA_DIALOGO", etiqueta: "La metodología favoreció el diálogo (mesas de trabajo, preguntas abiertas, plenaria)." },
  { clave: "AMBIENTE_RESPETO", etiqueta: "El ambiente fue de respeto, escucha y convivencia." },
  { clave: "CONCLUSIONES_ACUERDOS", etiqueta: "Se llegó a conclusiones o acuerdos concretos." },
  { clave: "PLAN_SOCIALIZAR", etiqueta: "Hay un plan para socializar los resultados con la comunidad." },
  { clave: "HORARIO_UBICACION", etiqueta: "El horario y la ubicación facilitaron que la comunidad asistiera." },
  { clave: "RESULTADOS_VISIBLES", etiqueta: "Los resultados de la discusión eran visibles y compartidos para todos los asistentes." },
  { clave: "RELACION_PEI_PLAN", etiqueta: "El foro se relacionó con el PEI, el Plan de Desarrollo Municipal u otros proyectos en curso." },
  { clave: "PROPUESTAS_VIABLES", etiqueta: "Hubo propuestas viables que nacieron desde las realidades institucionales." }
];

function cabecerasValoracionEventoComunal_() {
  return ["CLAVE", "ID_GRUPO", "SECCION"]
    .concat(ASPECTOS_VALORACION_EVENTO_.map(function (a) { return a.clave; }))
    .concat(["RECOMENDACIONES", "ULTIMA_ACTUALIZACION"]);
}

function _claveValoracionEvento_(idGrupo, seccion) {
  return String(idGrupo || "").trim() + "|" + String(seccion || "").trim();
}

/** Valores ya guardados (si los hay) de la valoración del evento de este grupo y sección, listos para pintar el formulario. */
function obtenerValoracionEvento(idGrupo, seccion) {
  var hoja = obtenerHoja_(HOJA_VALORACION_EVENTO_, cabecerasValoracionEventoComunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "CLAVE", _claveValoracionEvento_(idGrupo, seccion));
  var valores = {};
  ASPECTOS_VALORACION_EVENTO_.forEach(function (a) {
    valores[a.clave] = fila === -1 ? "" : String(hoja.getRange(fila, mapa[a.clave]).getValue() || "");
  });
  return {
    aspectos: ASPECTOS_VALORACION_EVENTO_,
    valores: valores,
    recomendaciones: fila === -1 ? "" : String(hoja.getRange(fila, mapa["RECOMENDACIONES"]).getValue() || "")
  };
}

/**
 * Guarda (autoguardado, UPSERT) la valoración del evento — `valores` trae
 * cualquier subconjunto de los aspectos (1-5, o vacío para "sin
 * responder todavía") más, opcionalmente, RECOMENDACIONES; se fusiona
 * con lo ya guardado, igual que el resto de autoguardados de la app.
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
  var clavesValidas = {};
  ASPECTOS_VALORACION_EVENTO_.forEach(function (a) { clavesValidas[a.clave] = true; });
  Object.keys(valores).forEach(function (clave) {
    if (clave === "RECOMENDACIONES") {
      datos.RECOMENDACIONES = String(valores.RECOMENDACIONES || "").trim();
      return;
    }
    if (!clavesValidas[clave]) return;
    var valor = Number(valores[clave]);
    if (!valor) return;
    if (valor < 1 || valor > 5) return;
    datos[clave] = Math.round(valor);
  });

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

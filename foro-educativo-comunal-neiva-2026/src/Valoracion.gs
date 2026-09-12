/**
 * Valoracion.gs — Foro Educativo Comunal Neiva 2026
 *
 * Encuesta de cierre "Valoración del Foro", copiada del instrumento real
 * de FEI 3.1 (App.html:7970 inicializarValoracionFEM_ / Index.html
 * "pantallaCierreFEM": 4 preguntas con escala de 1 a 5 corazones + 1
 * pregunta abierta final cuyo texto cambia según el promedio — ver
 * docs/01-auditoria-fei-3.1.md §2.4), adaptada de "Foro Educativo
 * Institucional"/IE a "Foro Educativo Comunal"/grupo. Igual que en 3.1,
 * la responde solo quien diligenció el formulario (cualquier dispositivo
 * con sesión activa), una sola vez por grupo.
 *
 * DIFERENCIA DELIBERADA frente a 3.1: allá la valoración "no bloquea
 * nada" (se ofrece después de generar/enviar el informe, como cierre
 * formal). Aquí, por pedido explícito de esta entrega, la valoración SÍ
 * es condición: no se puede generar el informe del grupo sin haberla
 * enviado primero (ver generarInformeCompletoGrupo en Grupos.gs).
 */

var HOJA_VALORACION_COMUNAL_ = "ValoracionComunal";

function cabecerasValoracionComunal_() {
  return [
    "ID_GRUPO", "P1_DIALOGO_REFLEXION", "P2_PARTICIPACION", "P3_IDEAS_PROPUESTAS",
    "P4_SATISFACCION_INSTRUMENTO", "NOTA_PROMEDIO", "P1_MEJORA", "P2_MEJORA", "P3_MEJORA", "P4_MEJORA",
    "P5_SUGERENCIAS", "FECHA"
  ];
}

/** Valoración ya enviada por el grupo, o null si todavía no la ha enviado. */
function obtenerValoracionGrupo(idGrupo) {
  var hoja = obtenerHoja_(HOJA_VALORACION_COMUNAL_, cabecerasValoracionComunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", String(idGrupo || "").trim());
  if (fila === -1) return null;
  return leerFilaComoObjeto_(hoja, fila, mapa);
}

/**
 * Guarda la valoración del grupo (una sola vez — un segundo envío
 * simplemente actualiza la misma fila, igual que el resto de UPSERT de
 * este proyecto). p1..p4 son obligatorios (1 a 5 corazones); las 4
 * respuestas de "mejora" y p5 son de texto libre y opcionales.
 */
function guardarValoracionGrupo(idGrupo, tokenSesion, dispositivoId, respuestas) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  respuestas = respuestas || {};
  var p1 = Number(respuestas.p1), p2 = Number(respuestas.p2), p3 = Number(respuestas.p3), p4 = Number(respuestas.p4);
  if (!p1 || !p2 || !p3 || !p4 || p1 < 1 || p1 > 5 || p2 < 1 || p2 > 5 || p3 < 1 || p3 > 5 || p4 < 1 || p4 > 5) {
    return { ok: false, mensaje: "Seleccione de 1 a 5 corazones en las cuatro preguntas antes de enviar." };
  }
  var nota = (p1 + p2 + p3 + p4) / 4;

  return conLock_(function () {
    upsertFila_(HOJA_VALORACION_COMUNAL_, cabecerasValoracionComunal_(), "ID_GRUPO", idGrupo, {
      P1_DIALOGO_REFLEXION: p1,
      P2_PARTICIPACION: p2,
      P3_IDEAS_PROPUESTAS: p3,
      P4_SATISFACCION_INSTRUMENTO: p4,
      NOTA_PROMEDIO: nota,
      P1_MEJORA: String(respuestas.mejoraP1 || "").trim(),
      P2_MEJORA: String(respuestas.mejoraP2 || "").trim(),
      P3_MEJORA: String(respuestas.mejoraP3 || "").trim(),
      P4_MEJORA: String(respuestas.mejoraP4 || "").trim(),
      P5_SUGERENCIAS: String(respuestas.p5 || "").trim(),
      FECHA: new Date()
    });
    return { ok: true, notaPromedio: nota };
  }, 15000);
}

/**
 * Valoración pública (sin código de acceso) — mismas 4 preguntas de
 * satisfacción de arriba, pero diligenciada por cualquier asistente desde
 * la página pública de asistencia QR (spec del usuario: "La valoración
 * del foro debe aparecer en la asistencia del código QR con las mismas
 * preguntas"). A diferencia de ValoracionComunal (una sola fila por
 * grupo, la del responsable), aquí CADA asistente que la diligencia deja
 * su propia fila — no hay upsert ni clave única, es anónima y opcional.
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
 * el textarea correspondiente cuando la respuesta es de 1 o 2 corazones
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
    return { ok: false, mensaje: "Seleccione de 1 a 5 corazones en las cuatro preguntas antes de enviar." };
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

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
 *
 * Documento Orientador FEM2026, item 20: valoraciones SEPARADAS para el
 * Encuentro y para Conecta Educa — cada sección tiene su propia fila,
 * identificada por CLAVE = ID_GRUPO + "|" + SECCION (mismo patrón que
 * _claveSocializacion_ en Socializacion.gs), en vez de una sola fila por
 * grupo. El informe sigue siendo uno solo por grupo (spec sección 18): lo
 * que se separa es la valoración, no el documento generado.
 */

var HOJA_VALORACION_COMUNAL_ = "ValoracionComunal";

function cabecerasValoracionComunal_() {
  return [
    "CLAVE", "ID_GRUPO", "SECCION", "P1_DIALOGO_REFLEXION", "P2_PARTICIPACION", "P3_IDEAS_PROPUESTAS",
    "P4_SATISFACCION_INSTRUMENTO", "NOTA_PROMEDIO", "P1_MEJORA", "P2_MEJORA", "P3_MEJORA", "P4_MEJORA",
    "P5_SUGERENCIAS", "FECHA"
  ];
}

function _claveValoracion_(idGrupo, seccion) {
  return String(idGrupo || "").trim() + "|" + String(seccion || "").trim();
}

/**
 * Valoración ya enviada por el grupo para `seccion` ("ENCUENTRO" o
 * "CONECTAEDUCA"), o null si todavía no la ha enviado. Si se omite
 * `seccion`, devuelve la primera valoración que encuentre del grupo (de
 * cualquier sección) — usado solo donde ya no importa cuál sección la dio
 * (enviarInformeSiCorresponde en Correo.gs: el informe ya existe para
 * ese momento, sin importar qué sección lo generó).
 */
function obtenerValoracionGrupo(idGrupo, seccion) {
  idGrupo = String(idGrupo || "").trim();
  var hoja = obtenerHoja_(HOJA_VALORACION_COMUNAL_, cabecerasValoracionComunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  if (seccion) {
    // Misma normalización que guardarValoracionGrupo_ (mayúsculas) — para
    // que un "encuentro" en minúscula desde el cliente no falle en
    // encontrar la fila guardada como "ENCUENTRO".
    var fila = buscarFilaPorColumna_(hoja, mapa, "CLAVE", _claveValoracion_(idGrupo, String(seccion).toUpperCase()));
    return fila === -1 ? null : leerFilaComoObjeto_(hoja, fila, mapa);
  }
  var filas = leerFilasComoObjetos_(hoja);
  return filas.find(function (f) { return String(f.ID_GRUPO || "") === idGrupo; }) || null;
}

/**
 * Guarda la valoración del grupo para `seccion` (una sola vez por
 * sección — un segundo envío simplemente actualiza la misma fila, igual
 * que el resto de UPSERT de este proyecto). p1..p4 son obligatorios (1 a
 * 5, ver escala en Index.html); las 4 respuestas de "mejora" y p5 son de
 * texto libre y opcionales.
 */
function guardarValoracionGrupo(idGrupo, tokenSesion, dispositivoId, respuestas, seccion) {
  idGrupo = String(idGrupo || "").trim();
  seccion = String(seccion || "ENCUENTRO").toUpperCase();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  respuestas = respuestas || {};
  var p1 = Number(respuestas.p1), p2 = Number(respuestas.p2), p3 = Number(respuestas.p3), p4 = Number(respuestas.p4);
  if (!p1 || !p2 || !p3 || !p4 || p1 < 1 || p1 > 5 || p2 < 1 || p2 > 5 || p3 < 1 || p3 > 5 || p4 < 1 || p4 > 5) {
    return { ok: false, mensaje: "Seleccione una opción de 1 a 5 en las cuatro preguntas antes de enviar." };
  }
  var nota = (p1 + p2 + p3 + p4) / 4;

  return conLock_(function () {
    upsertFila_(HOJA_VALORACION_COMUNAL_, cabecerasValoracionComunal_(), "CLAVE", _claveValoracion_(idGrupo, seccion), {
      ID_GRUPO: idGrupo,
      SECCION: seccion,
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
    // Sin flush(), un clic casi inmediato en "Generar informe" (otra
    // ejecución de Apps Script, disparada apenas el cliente marca
    // estado.valoracionCompletada = true) puede leer obtenerValoracionGrupo
    // antes de que esta fila sea visible, y rechazar con "Debe completar la
    // valoración..." aunque sí se guardó — mismo caso que el fix de
    // firmantes en vivo en Data.gs (registrarParticipante).
    SpreadsheetApp.flush();
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

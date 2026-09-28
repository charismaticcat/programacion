/**
 * Valoracion.gs — Foro Educativo Comunal Neiva 2026
 *
 * La valoración del Foro DENTRO de la app (4 preguntas de satisfacción
 * que respondía quien diligenciaba el formulario, una vez por
 * grupo/sección, condición para generar el informe) se eliminó por
 * completo del proyecto — spec del usuario: "Eliminar valoración". Ya no
 * gatea generarInformeCompletoGrupo (Grupos.gs) ni enviarInformeSiCorresponde
 * (Correo.gs); tampoco existe en el HTML (tarjetaValoracion) ni en el
 * cliente (JS.html). La hoja ValoracionComunal con las respuestas ya
 * recogidas antes de este cambio queda intacta como registro histórico,
 * simplemente ya no se escribe ni se lee desde la app.
 *
 * Lo único que queda aquí es la valoración PÚBLICA y ANÓNIMA de la
 * página de asistencia QR (AsistenciaPublica.html) — distinta, no
 * bloquea nada y no se pidió eliminarla.
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

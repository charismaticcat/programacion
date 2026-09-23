/**
 * Socializacion.gs — Foro Educativo Comunal Neiva 2026
 *
 * "Sesión de socialización" (Documento Orientador FEM2026, sección D,
 * item 10) — reemplaza a "Preparación" (item 9, oculta junto con todas
 * sus subpáginas): en vez de que cada IE llene un formulario previo, una
 * sola persona (spec del usuario: "una persona va a ingresar los datos")
 * lleva en vivo, durante el evento, un checklist de qué institución ya
 * socializó sus reflexiones, conclusiones y propuestas del Foro
 * Educativo Institucional, con un temporizador de apoyo por IE (ver
 * JS.html). No hay toma manual de notas aquí (spec del usuario: "solo
 * dejar el temporizador") — las respuestas reales del grupo se muestran
 * en Sesión 1, ver SintesisGrupos.gs.
 *
 * El checklist "¿ya socializó?" ya NO se guarda en una hoja de cálculo
 * (spec del usuario: "eliminar todo el registro del spreadsheets") — es
 * una ayuda puramente visual para quien lleva el registro en vivo durante
 * la sesión (una sola persona, en un solo dispositivo), así que su estado
 * se mantiene solo en el cliente (JS.html, estadoSocializacion.datos) y
 * se reinicia si se recarga la página. obtenerSocializacionGrupo sigue
 * siendo la fuente de las instituciones del grupo, todas con socializo
 * en falso al cargar.
 */

/** Checklist de socialización de todas las IE del grupo, listo para pintar en pantalla. */
function obtenerSocializacionGrupo(idGrupo) {
  idGrupo = String(idGrupo || "").trim();
  var instituciones = obtenerInstitucionesDelGrupo(idGrupo);
  return instituciones.map(function (ie) {
    return { idIE: ie.idIE, institucion: ie.institucion, socializo: false };
  });
}

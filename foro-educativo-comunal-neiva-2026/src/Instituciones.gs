/**
 * Instituciones.gs — Foro Educativo Comunal Neiva 2026
 *
 * Las IE son miembros de un GRUPO (no la unidad raíz). Todo se consulta
 * contra GruposComunal, que es la fuente única de la relación Grupo → IE
 * (spec sección 6: "la información Grupo → IE debe provenir de Google
 * Sheets. NO codificar esta relación directamente en HTML o JavaScript").
 * Reemplaza el catálogo hardcodeado y duplicado de FEI 3.1
 * (GRUPOS_INSTITUCIONES en App.html + GRUPOS_INSTITUCIONES_FEM_ en
 * Código.js — ver docs/01-auditoria-fei-3.1.md §1.6/§7).
 */

var HOJA_GRUPOS_COMUNAL_ = "GruposComunal";

function cabecerasGruposComunal_() {
  return ["ID_GRUPO", "GRUPO", "ID_IE", "INSTITUCION", "COMUNA", "ACTIVO"];
}

/** Todas las IE (activas) que pertenecen a un grupo, tal como las cargó la SEM en GruposComunal. */
function obtenerInstitucionesDelGrupo(idGrupo) {
  var hoja = obtenerHoja_(HOJA_GRUPOS_COMUNAL_, cabecerasGruposComunal_());
  var filas = leerFilasComoObjetos_(hoja);
  var objetivo = String(idGrupo || "").trim();
  return filas
    .filter(function (f) {
      return String(f.ID_GRUPO || "").trim() === objetivo && String(f.ACTIVO || "SI").toUpperCase() !== "NO";
    })
    .map(function (f) {
      return {
        idIE: String(f.ID_IE || "").trim(),
        institucion: String(f.INSTITUCION || "").trim(),
        comuna: String(f.COMUNA || "").trim()
      };
    });
}

/** Busca a qué grupo pertenece una IE (por ID_IE o por nombre normalizado). */
function obtenerGrupoDeInstitucion_(identificadorIE) {
  var hoja = obtenerHoja_(HOJA_GRUPOS_COMUNAL_, cabecerasGruposComunal_());
  var filas = leerFilasComoObjetos_(hoja);
  var objetivo = String(identificadorIE || "").trim();
  var objetivoNormalizado = normalizarTexto_(identificadorIE);
  var encontrada = filas.find(function (f) {
    return (
      String(f.ID_IE || "").trim() === objetivo ||
      normalizarTexto_(f.INSTITUCION) === objetivoNormalizado
    );
  });
  return encontrada ? { idGrupo: String(encontrada.ID_GRUPO).trim(), grupo: String(encontrada.GRUPO).trim() } : null;
}

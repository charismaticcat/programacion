/**
 * Recursos.gs — Foro Educativo Comunal Neiva 2026
 *
 * Enlaces de descarga necesarios antes de iniciar la Sesión 1 (spec:
 * "para realizar la primera parte es necesario descargar los siguientes
 * archivos"): dos recursos fijos (informe de síntesis municipal y FAQs
 * del Foro Educativo Institucional) y dos recursos que dependen del grupo
 * activo (Respuestas Compiladas e Informe de Síntesis de ese grupo).
 *
 * Los IDs de Drive de este archivo fueron verificados uno a uno el
 * 2026-09-09 dentro de la carpeta pública "Informes por grupos" del FEI
 * (https://drive.google.com/drive/folders/1SAjGsKFNudF94ag_xCWmPNuqE0nj9T85)
 * — regla de no invención (spec sección 23): nunca se generan ni adivinan
 * IDs de archivo.
 */

/** Recursos fijos (no dependen del grupo), en el orden en que deben mostrarse al final de la lista. */
var RECURSOS_FIJOS_SESION1_ = [
  {
    emoji: "📘",
    titulo: "Informe de síntesis municipal",
    descripcion: "Consolidado municipal del Foro Educativo Institucional 2026, con los hallazgos comunes de todas las instituciones de Neiva.",
    url: "https://drive.google.com/file/d/1H8ro89vUYRKRE0UxGiXiCQw44_n17klI/view?usp=drive_link"
  },
  {
    emoji: "❓",
    titulo: "FAQs del Foro Educativo Institucional",
    descripcion: "Preguntas frecuentes sobre el desarrollo del Foro Educativo Institucional 2026.",
    url: "https://drive.google.com/file/d/1sI2g4NKUVJ5PyyK3IyotPmIaJ9SfulUj/view?usp=drive_link"
  }
];

/**
 * Documentos propios de cada grupo (Respuestas Compiladas + Informe de
 * Síntesis), ambos ubicados en la carpeta de Drive "Informes por grupos"
 * del FEI. Documentos de Google Docs (no PDF): se enlazan con
 * export?format=pdf para que la descarga sea directa en un clic.
 */
var RECURSOS_POR_GRUPO_ = {
  G1: { respuestasCompiladasId: "1Lixu0vWKJ3lUaYF03O5RZUhdIdpj9CksdiE4vmOd6dI", informeSintesisId: "1QGLrKTKEEzvNU6b4YWaPf5mdHURf-hL-TVIngfuhWLc" },
  G2: { respuestasCompiladasId: "1Lee1F7WfwdMaNPBKBJpfpcd_yXiJMznsz_lqMaNtpFM", informeSintesisId: "1xoHE3E2yHgGOBOFuO2_NSc2AhUujUwfRVlIz3eyQN-Y" },
  G3: { respuestasCompiladasId: "11mOH1RzgKAO7uecBjq7fBF9pgNWO2B7PUou_fL-09r4", informeSintesisId: "14NDJ52CArtGXGG0SrjqZloru-nyA2-VVw-ImB_MgC8A" },
  G4: { respuestasCompiladasId: "1fxgl4fJLMPoixf8aWXbXGO8iqQe_1qJIZvPmtck42Ug", informeSintesisId: "1pSxRhlvQRcdcPA3knSyNC1HmGAMhwHODHMCNO834L_E" },
  G5: { respuestasCompiladasId: "1WafPL1sfqVRabZSCrgibZMUwXG1SvSVNj66fcL_eCmk", informeSintesisId: "1cZ2QTkSR_OWTLTsqSbGRCEqBWw9lbJd3YfCHqcirrS4" },
  G6: { respuestasCompiladasId: "1QXFnuaG7Zqt9fSUEO65_s5Oc6QPCZjxoqtoMoFc7p8w", informeSintesisId: "1y6pYoaZJDtgjXwwP5bVwtsJtJbqF79T1NnZazRBKzHk" }
};

function _urlDescargaDocumentoDrive_(idArchivo) {
  return "https://docs.google.com/document/d/" + idArchivo + "/export?format=pdf";
}

/**
 * Los recursos descargables necesarios antes de iniciar Sesión 1, en
 * orden: primero los 2 dependientes del grupo activo (Respuestas
 * Compiladas, Consolidado del grupo), luego los 2 fijos (síntesis
 * municipal, FAQs). Si el grupo no tiene documentos propios mapeados
 * todavía (por ejemplo GRUPO-PRUEBA, o un grupo nuevo) esos dos ítems se
 * omiten en vez de inventar un enlace — nunca se muestra un enlace roto.
 */
function obtenerRecursosSesion1(idGrupo) {
  var idGrupoStr = String(idGrupo || "").trim();
  var recursos = [];
  var porGrupo = RECURSOS_POR_GRUPO_[idGrupoStr];

  if (porGrupo && porGrupo.respuestasCompiladasId) {
    recursos.push({
      emoji: "📥",
      titulo: "Respuestas por I.E. del Foro Educativo Institucional",
      descripcion: "Respuestas compiladas de cada institución educativa del " + idGrupoStr + " en el Foro Educativo Institucional 2026.",
      url: _urlDescargaDocumentoDrive_(porGrupo.respuestasCompiladasId)
    });
  }
  if (porGrupo && porGrupo.informeSintesisId) {
    recursos.push({
      emoji: "📊",
      titulo: "Consolidado de " + idGrupoStr,
      descripcion: "Informe de síntesis del " + idGrupoStr + ": elementos comunes y particularidades de sus instituciones en el Foro Educativo Institucional 2026.",
      url: _urlDescargaDocumentoDrive_(porGrupo.informeSintesisId)
    });
  }

  RECURSOS_FIJOS_SESION1_.forEach(function (r) {
    recursos.push(r);
  });

  return recursos;
}

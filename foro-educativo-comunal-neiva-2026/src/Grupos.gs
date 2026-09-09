/**
 * Grupos.gs — Foro Educativo Comunal Neiva 2026
 *
 * GRUPO es la unidad raíz del sistema (spec sección 27: "la unidad
 * principal NO ES LA INSTITUCIÓN EDUCATIVA. La unidad principal es EL
 * GRUPO"). Este archivo expone el catálogo de grupos (derivado de
 * GruposComunal, nunca hardcodeado — a diferencia de
 * GRUPOS_INSTITUCIONES_FEM_ en 3.1) y orquesta la generación del informe
 * único por grupo.
 */

/** Lista de grupos distintos presentes en GruposComunal, con el conteo de IE de cada uno. */
function obtenerGrupos() {
  var hoja = obtenerHoja_(HOJA_GRUPOS_COMUNAL_, cabecerasGruposComunal_());
  var filas = leerFilasComoObjetos_(hoja);
  var porGrupo = {};
  var orden = [];
  filas.forEach(function (f) {
    var idGrupo = String(f.ID_GRUPO || "").trim();
    if (!idGrupo) return;
    if (String(f.ACTIVO || "SI").toUpperCase() === "NO") return;
    if (!porGrupo[idGrupo]) {
      porGrupo[idGrupo] = { idGrupo: idGrupo, grupo: String(f.GRUPO || idGrupo).trim(), totalIE: 0 };
      orden.push(idGrupo);
    }
    porGrupo[idGrupo].totalIE++;
  });
  return orden.map(function (id) {
    return porGrupo[id];
  });
}

function obtenerGrupoPorId(idGrupo) {
  return obtenerGrupos().find(function (g) {
    return g.idGrupo === String(idGrupo || "").trim();
  }) || null;
}

/**
 * Estado agregado de un grupo (para el dashboard administrativo, Fase 25
 * de la spec): participación registrada, si Sesión 1/2 están guardadas o
 * enviadas, si ya tiene informe.
 */
function obtenerEstadoGrupo(idGrupo) {
  var idGrupoStr = String(idGrupo || "").trim();

  var hojaAccesos = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapaAccesos = obtenerMapaCabeceras_(hojaAccesos);
  var filaAcceso = buscarFilaPorColumna_(hojaAccesos, mapaAccesos, "ID_GRUPO", idGrupoStr);
  var acceso = filaAcceso === -1 ? null : leerFilaComoObjeto_(hojaAccesos, filaAcceso, mapaAccesos);

  var totalParticipantes = contarParticipantesGrupo(idGrupoStr);

  var hojaInformes = obtenerHoja_("InformesComunal", cabecerasInformesComunal_());
  var mapaInformes = obtenerMapaCabeceras_(hojaInformes);
  var filaInforme = buscarFilaPorColumna_(hojaInformes, mapaInformes, "ID_GRUPO", idGrupoStr);
  var informe = filaInforme === -1 ? null : leerFilaComoObjeto_(hojaInformes, filaInforme, mapaInformes);

  return {
    idGrupo: idGrupoStr,
    estado: acceso ? acceso.ESTADO : "PENDIENTE",
    sesion1Enviada: acceso ? acceso.SESION1_ENVIADA === "SI" : false,
    sesion2Enviada: acceso ? acceso.SESION2_ENVIADA === "SI" : false,
    totalParticipantes: totalParticipantes,
    informeGenerado: !!(informe && informe.DOC_ID),
    informeUrl: informe ? informe.URL : ""
  };
}

/**
 * Orquesta la generación del informe único del grupo: asegura la carpeta
 * de Drive del grupo, genera Doc+PDF (Informes.gs) y deja el registro en
 * InformesComunal. Todas las IE del grupo apuntan al mismo resultado
 * (spec sección 18) — esta es la única función que crea el informe; no
 * existe una función equivalente "por IE".
 */
function generarInformeCompletoGrupo(idGrupo, tokenSesion, dispositivoId) {
  return conLock_(function () {
    if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
      return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
    }
    if (!esPrincipalDeGrupo_(idGrupo, dispositivoId, tokenSesion)) {
      return { ok: false, mensaje: "Solo el responsable principal del grupo puede generar el informe definitivo." };
    }
    // Condición explícita de esta entrega (distinta de FEI 3.1, ver
    // Valoracion.gs): no se genera el informe sin haber enviado antes la
    // valoración del Foro.
    if (!obtenerValoracionGrupo(idGrupo)) {
      return {
        ok: false,
        codigo: "VALORACION_REQUERIDA",
        mensaje: "Debe completar la valoración del Foro antes de generar el informe."
      };
    }
    var resultado = generarInformeGrupo(idGrupo);
    if (!resultado.ok) return resultado;

    marcarEstadoAccesoGrupo_(idGrupo, "INFORME_GENERADO");
    return resultado;
  }, 30000);
}

function marcarEstadoAccesoGrupo_(idGrupo, estado) {
  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", String(idGrupo || "").trim());
  if (fila === -1) return;
  hoja.getRange(fila, mapa["ESTADO"]).setValue(estado);
}

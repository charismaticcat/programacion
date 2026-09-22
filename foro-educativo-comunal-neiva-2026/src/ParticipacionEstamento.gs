/**
 * ParticipacionEstamento.gs — Foro Educativo Comunal Neiva 2026
 *
 * Conteo de participantes por estamento e institución, en la pantalla de
 * Participación — mismo formato que la hoja "Participación" de FEI 3.1
 * (docs/01-auditoria-fei-3.1.md §3.1: Rector(a), Coordinador(a), Docentes,
 * Tutor PTA PFI/3.0, Orientador(a), Estudiantes, Padres/madres/acudientes,
 * Personal administrativo, Egresados, Sector productivo, Otros — con
 * autosuma), pero repetido POR CADA IE del grupo (en 3.1 era un formulario
 * por IE porque la unidad raíz era la IE; aquí la unidad raíz es el
 * GRUPO, así que cada grupo diligencia una tabla de estas por cada una de
 * sus instituciones).
 *
 * Es un conteo manual (números que alguien del grupo escribe), distinto
 * del conteo de firmantes en vivo (ParticipacionComunal/registrarParticipante
 * en Data.gs, usado por el panel de firmantes y por rpcEstadoFirmantes) —
 * cuando el grupo usa el método QR, la pantalla de Participación muestra
 * ambos números lado a lado (firmantes / participantes declarados) para
 * que el grupo compare cuántas de las firmas ya registradas corresponden
 * a lo que declararon aquí.
 */

var HOJA_PARTICIPACION_ESTAMENTO_ = "ParticipacionEstamentoIE";

/**
 * Estamentos — mismo orden y nombres que la hoja Participacion de FEI 3.1,
 * más FUNCIONARIO_SEM (Documento Orientador FEM2026, item 23): igual que
 * SECTOR ("Sector productivo"), no está afiliado a una IE en particular,
 * pero se registra bajo la columna de la IE que corresponda al contexto
 * de su asistencia, mismo criterio ya usado para SECTOR.
 */
var ESTAMENTOS_PARTICIPACION_ = [
  { clave: "RECTOR", etiqueta: "Rector(a)" },
  { clave: "COORDINADOR", etiqueta: "Coordinador(a)" },
  { clave: "DOCENTES", etiqueta: "Docentes" },
  { clave: "TUTOR_PTA", etiqueta: "Tutor PTA PFI/3.0" },
  { clave: "ORIENTADOR", etiqueta: "Orientador(a)" },
  { clave: "ESTUDIANTES", etiqueta: "Estudiantes" },
  { clave: "PADRES", etiqueta: "Padres/madres/acudientes" },
  { clave: "ADMINISTRATIVOS", etiqueta: "Personal administrativo" },
  { clave: "EGRESADOS", etiqueta: "Egresados" },
  { clave: "SECTOR", etiqueta: "Sector productivo" },
  { clave: "FUNCIONARIO_SEM", etiqueta: "Funcionario Secretaría de Educación" },
  { clave: "OTROS", etiqueta: "Otros" }
];

function cabecerasParticipacionEstamentoIE_() {
  return ["CLAVE", "ID_GRUPO", "ID_IE"]
    .concat(ESTAMENTOS_PARTICIPACION_.map(function (e) { return e.clave; }))
    .concat(["ULTIMA_ACTUALIZACION", "PRESENTE"]);
}

function _claveParticipacionEstamento_(idGrupo, idIE) {
  return String(idGrupo || "").trim() + "|" + String(idIE || "").trim();
}

/**
 * Conteo por estamento de todas las IE del grupo, listas para la
 * pantalla de Participación (una IE que todavía no diligenció nada
 * aparece con todos los estamentos en 0, nunca se omite). Cada IE trae
 * además `presente` (spec del usuario: "Responsable de envio debe
 * seleccionar qué instituciones educativas están presentes") — el
 * cliente usa esa marca para filtrar la matriz y el informe a solo las
 * IE presentes; `totalesPorEstamento`/`totalGeneral` YA vienen calculados
 * solo sobre las IE presentes.
 */
function obtenerParticipacionEstamentoGrupo(idGrupo) {
  var idGrupoStr = String(idGrupo || "").trim();
  var hoja = obtenerHoja_(HOJA_PARTICIPACION_ESTAMENTO_, cabecerasParticipacionEstamentoIE_());
  var filas = leerFilasComoObjetos_(hoja);
  var porIE = {};
  filas.forEach(function (f) {
    if (String(f.ID_GRUPO || "").trim() !== idGrupoStr) return;
    porIE[String(f.ID_IE || "").trim()] = f;
  });

  var totalesPorEstamento = {};
  ESTAMENTOS_PARTICIPACION_.forEach(function (e) {
    totalesPorEstamento[e.clave] = 0;
  });

  var totalGeneral = 0;
  var instituciones = obtenerInstitucionesDelGrupo(idGrupoStr).map(function (ie) {
    var fila = porIE[ie.idIE];
    var presente = !!fila && String(fila.PRESENTE || "") === "SI";
    var valores = {};
    var totalIE = 0;
    ESTAMENTOS_PARTICIPACION_.forEach(function (e) {
      var valor = fila ? Number(fila[e.clave] || 0) : 0;
      valores[e.clave] = valor;
      totalIE += valor;
      if (presente) totalesPorEstamento[e.clave] += valor;
    });
    if (presente) totalGeneral += totalIE;
    return { idIE: ie.idIE, institucion: ie.institucion, valores: valores, total: totalIE, presente: presente };
  });

  return {
    estamentos: ESTAMENTOS_PARTICIPACION_,
    instituciones: instituciones,
    totalesPorEstamento: totalesPorEstamento,
    totalGeneral: totalGeneral
  };
}

/**
 * Marca (UPSERT) si una IE del grupo está presente en esta sesión — spec
 * del usuario: casilla de "instituciones educativas presentes" justo
 * después de Responsable de envío. No pisa los conteos por estamento ya
 * guardados de esa IE si la fila ya existía.
 */
function guardarPresenciaIE(idGrupo, tokenSesion, dispositivoId, idIE, presente) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  idIE = String(idIE || "").trim();
  if (!idIE) return { ok: false, mensaje: "Falta la institución." };
  var perteneceAlGrupo = obtenerInstitucionesDelGrupo(idGrupo).some(function (ie) {
    return ie.idIE === idIE;
  });
  if (!perteneceAlGrupo) return { ok: false, mensaje: "Esa institución no pertenece a este grupo." };

  return conLock_(function () {
    upsertFila_(
      HOJA_PARTICIPACION_ESTAMENTO_,
      cabecerasParticipacionEstamentoIE_(),
      "CLAVE",
      _claveParticipacionEstamento_(idGrupo, idIE),
      { ID_GRUPO: idGrupo, ID_IE: idIE, PRESENTE: presente ? "SI" : "NO" }
    );
    return { ok: true };
  }, 10000);
}

/** Guarda (UPSERT) el conteo por estamento de UNA IE del grupo — igual patrón de autoguardado que el resto de la app. */
function guardarParticipacionEstamentoIE(idGrupo, tokenSesion, dispositivoId, idIE, valores) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  idIE = String(idIE || "").trim();
  if (!idIE) return { ok: false, mensaje: "Falta la institución." };
  // Nunca se guarda el conteo de una IE que no pertenezca a este grupo.
  var perteneceAlGrupo = obtenerInstitucionesDelGrupo(idGrupo).some(function (ie) {
    return ie.idIE === idIE;
  });
  if (!perteneceAlGrupo) return { ok: false, mensaje: "Esa institución no pertenece a este grupo." };

  var datos = { ID_GRUPO: idGrupo, ID_IE: idIE, ULTIMA_ACTUALIZACION: new Date() };
  var total = 0;
  ESTAMENTOS_PARTICIPACION_.forEach(function (e) {
    var valor = Math.max(0, Math.round(Number((valores || {})[e.clave]) || 0));
    datos[e.clave] = valor;
    total += valor;
  });

  return conLock_(function () {
    upsertFila_(
      HOJA_PARTICIPACION_ESTAMENTO_,
      cabecerasParticipacionEstamentoIE_(),
      "CLAVE",
      _claveParticipacionEstamento_(idGrupo, idIE),
      datos
    );
    return { ok: true, total: total };
  }, 10000);
}

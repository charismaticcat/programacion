/**
 * Conversatorio.gs — Foro Educativo Comunal Neiva 2026
 *
 * "Conversatorio" (pedido del usuario): la nueva primera parte, autocontenida
 * e independiente, de Conecta Educa — reemplaza el sistema anterior de
 * acceso "por grupo comunal" (código propio de cada uno de los 6 grupos de
 * GruposComunal, y luego el administrador único de la SEM del lote
 * anterior). Ahora cada Institución Educativa se elige a sí misma de un
 * listado — sin código (pedido del usuario: "omite que se generen codigos
 * por IE y mas bien que ellos mismos escojan la IE de un listado") — sin
 * pasar por un administrador ni por los 6 grupos comunales del Encuentro.
 *
 * El resto de Conecta Educa (pantallaConsentimientoConectaEduca,
 * pantallaSesion2, etc., construidas alrededor de GruposComunal) NO se
 * tocó — sigue existiendo en el código tal cual estaba, sin conectarse con
 * el Conversatorio (decisión explícita del usuario).
 *
 * Datos fuente: hoja de cálculo de Google externa (agrupación de las
 * Instituciones Educativas de Neiva en 7 grupos por técnica/articulación
 * SENA — "Grupos ConectaEduca", "Matriz IE-Técnica" — y el detalle de
 * programas aprobados por resolución/decreto por IE — "Resolución Vs IE").
 * importarDatosConversatorio() la lee directamente (misma cuenta de Google
 * que el proyecto) y puebla las hojas propias del proyecto — se ejecuta a
 * mano desde el editor cada vez que la Secretaría actualice ese documento
 * fuente; nunca se llama automáticamente (mismo patrón que
 * generarAccesosGrupo).
 */

var ID_HOJA_FUENTE_CONVERSATORIO_ = "1_ewh04NaoSCfbndj8oaZtzZdrD-sJpeH";

var HOJA_CONVERSATORIO_GRUPOS_ = "ConversatorioGrupos";
var HOJA_CONVERSATORIO_MATRIZ_ = "ConversatorioMatrizIETecnica";
var HOJA_CONVERSATORIO_RESOLUCION_ = "ConversatorioResolucionIE";
var HOJA_CONVERSATORIO_ACCESOS_ = "AccesosIEConversatorio";
// Registro de cambios (pedido del usuario: "Esos cambios deben mencionarse
// explícitamente en el informe... la IE X cambió la respuesta X por Y") —
// se llena solo cuando se SOBRESCRIBE un valor que ya tenía algo (ver
// guardarCampoTecnicaConversatorio), nunca en el llenado inicial.
var HOJA_CONVERSATORIO_CAMBIOS_ = "ConversatorioCambios";
// Único DOC_ID/PDF_ID del informe consolidado (mismo patrón "nunca
// acumular" que InformesComunal).
var HOJA_CONVERSATORIO_INFORME_ = "ConversatorioInforme";

function cabecerasConversatorioGrupos_() {
  return ["ID_GRUPO", "NOMBRE_GRUPO", "NUM_IE", "TECNICAS_INCLUIDAS", "CRITERIO_AGRUPACION"];
}
function cabecerasConversatorioMatriz_() {
  return ["ID_GRUPO", "NOMBRE_GRUPO", "INSTITUCION_EDUCATIVA", "TECNICA"];
}
function cabecerasConversatorioResolucion_() {
  return [
    "ID_FILA", "INSTITUCION_EDUCATIVA", "TECNICA_SENA", "PROGRAMA_APROBADO", "RESOLUCION_DECRETO",
    "PREGUNTA_APERTURA_GRADO10_2027", "PREGUNTA_NUEVA_ARTICULACION_2027", "NUEVA_ARTICULACION_DETALLE",
    "CONFIRMADO_POR_IE", "FECHA_CONFIRMACION", "ACTUALIZADO"
  ];
}
function cabecerasConversatorioAccesos_() {
  return [
    "INSTITUCION_EDUCATIVA", "ESTADO", "GRUPO_ELEGIDO", "ULTIMA_ACTIVIDAD",
    "RESPONSABLE_NOMBRE", "RESPONSABLE_ROL", "ACOMPANANTE"
  ];
}
function cabecerasConversatorioCambios_() {
  return ["INSTITUCION_EDUCATIVA", "TECNICA_SENA", "CAMPO", "VALOR_ANTERIOR", "VALOR_NUEVO", "FECHA"];
}
function cabecerasConversatorioInforme_() {
  return ["DOC_ID", "PDF_ID", "URL", "FECHA"];
}

/** Etiqueta legible por campo — usada tanto en el registro de cambios como en el informe. */
var ETIQUETAS_CAMPOS_TECNICA_CONVERSATORIO_ = {
  TECNICA_SENA: "técnica / programa SENA",
  PROGRAMA_APROBADO: "programa aprobado por resolución",
  RESOLUCION_DECRETO: "resolución o decreto",
  PREGUNTA_APERTURA_GRADO10_2027: "¿Se aperturará esta articulación para grado 10° en el 2027?",
  PREGUNTA_NUEVA_ARTICULACION_2027: "¿Se planea incluir una nueva articulación para el año 2027?",
  NUEVA_ARTICULACION_DETALLE: "detalle de la nueva articulación"
};

/**
 * Nombres de Institución Educativa que aparecen distinto en la hoja
 * "Resolución Vs IE" del documento fuente que en "Matriz IE-Técnica" (para
 * la misma institución) — sin este mapa, la comparación por nombre
 * normalizado (sin prefijo "IE", mayúsculas, sin tildes) no los uniría.
 */
var ALIAS_IE_CONVERSATORIO_ = {
  "NACIONAL SANTA LIBRADA": "SANTA LIBRADA",
  "LICEO DE SANTA LIBRADA": "LICEO SANTA LIBRADA",
  "MARIA AUXILIADORA FORTALECILLAS": "MARIA AUXILIADORA - FORTALECILLAS"
};

function normalizarNombreIEConversatorio_(nombre) {
  var limpio = String(nombre || "")
    .trim()
    .toUpperCase()
    .replace(/^IE\s+/, "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return ALIAS_IE_CONVERSATORIO_[limpio] || limpio;
}

/**
 * Lee las 3 hojas necesarias del documento fuente y puebla las hojas del
 * proyecto. Idempotente: ConversatorioGrupos/ConversatorioMatriz se
 * reescriben por completo (no hay datos de usuario ahí); en
 * ConversatorioResolucionIE solo se agregan filas nuevas (institución +
 * técnica que todavía no exista) — las respuestas y ediciones ya guardadas
 * por una IE nunca se pisan. Las filas de AccesosIEConversatorio (una por
 * institución, sin código) tampoco se duplican para instituciones que ya
 * la tienen.
 */
function importarDatosConversatorio() {
  var ssFuente = SpreadsheetApp.openById(ID_HOJA_FUENTE_CONVERSATORIO_);

  var hojaGruposFuente = ssFuente.getSheetByName("Grupos ConectaEduca");
  var hojaMatrizFuente = ssFuente.getSheetByName("Matriz IE-Técnica");
  var hojaResolucionFuente = ssFuente.getSheetByName("Resolución Vs IE");
  if (!hojaGruposFuente || !hojaMatrizFuente || !hojaResolucionFuente) {
    return {
      ok: false,
      mensaje: "El documento fuente no tiene alguna de las hojas esperadas: " +
        '"Grupos ConectaEduca", "Matriz IE-Técnica", "Resolución Vs IE".'
    };
  }

  // --- Grupos (G01..G07) ---
  var filasGrupos = hojaGruposFuente.getDataRange().getValues();
  var grupos = [];
  for (var i = 1; i < filasGrupos.length; i++) {
    var f = filasGrupos[i];
    if (!String(f[0] || "").trim()) continue;
    grupos.push({
      ID_GRUPO: String(f[0]).trim(),
      NOMBRE_GRUPO: String(f[1] || "").trim(),
      NUM_IE: f[2] || "",
      TECNICAS_INCLUIDAS: String(f[3] || "").trim(),
      CRITERIO_AGRUPACION: String(f[4] || "").trim()
    });
  }
  var hojaGrupos = obtenerHoja_(HOJA_CONVERSATORIO_GRUPOS_, cabecerasConversatorioGrupos_());
  if (hojaGrupos.getLastRow() > 1) {
    hojaGrupos.getRange(2, 1, hojaGrupos.getLastRow() - 1, hojaGrupos.getLastColumn()).clearContent();
  }
  if (grupos.length) {
    hojaGrupos.getRange(2, 1, grupos.length, cabecerasConversatorioGrupos_().length).setValues(
      grupos.map(function (g) {
        return cabecerasConversatorioGrupos_().map(function (c) { return g[c]; });
      })
    );
  }

  // --- Matriz IE-Técnica ---
  var filasMatriz = hojaMatrizFuente.getDataRange().getValues();
  var matriz = [];
  var institucionesCanonicas = {}; // normalizado -> nombre canónico (tal como aparece en Matriz)
  for (var j = 1; j < filasMatriz.length; j++) {
    var fm = filasMatriz[j];
    var institucion = String(fm[2] || "").trim();
    if (!institucion) continue;
    matriz.push({
      ID_GRUPO: String(fm[0] || "").trim(),
      NOMBRE_GRUPO: String(fm[1] || "").trim(),
      INSTITUCION_EDUCATIVA: institucion,
      TECNICA: String(fm[3] || "").trim()
    });
    institucionesCanonicas[normalizarNombreIEConversatorio_(institucion)] = institucion;
  }
  var hojaMatriz = obtenerHoja_(HOJA_CONVERSATORIO_MATRIZ_, cabecerasConversatorioMatriz_());
  if (hojaMatriz.getLastRow() > 1) {
    hojaMatriz.getRange(2, 1, hojaMatriz.getLastRow() - 1, hojaMatriz.getLastColumn()).clearContent();
  }
  if (matriz.length) {
    hojaMatriz.getRange(2, 1, matriz.length, cabecerasConversatorioMatriz_().length).setValues(
      matriz.map(function (m) {
        return cabecerasConversatorioMatriz_().map(function (c) { return m[c]; });
      })
    );
  }

  // --- Resolución Vs IE (solo instituciones que sí pertenecen a algún
  //     grupo de Matriz IE-Técnica; "arrastra hacia abajo" el nombre de
  //     institución en filas de continuación con ese campo vacío) ---
  var filasResolucion = hojaResolucionFuente.getDataRange().getValues();
  var institucionActual = "";
  var hojaResolucion = obtenerHoja_(HOJA_CONVERSATORIO_RESOLUCION_, cabecerasConversatorioResolucion_());
  var existentesResolucion = leerFilasComoObjetos_(hojaResolucion);
  var clavesExistentes = {};
  existentesResolucion.forEach(function (fila) {
    clavesExistentes[
      normalizarNombreIEConversatorio_(fila.INSTITUCION_EDUCATIVA) + "||" + String(fila.TECNICA_SENA || "").trim().toUpperCase()
    ] = true;
  });
  var nuevasFilasResolucion = [];
  for (var k = 1; k < filasResolucion.length; k++) {
    var fr = filasResolucion[k];
    var nombreFila = String(fr[1] || "").trim();
    if (nombreFila) institucionActual = nombreFila;
    if (!institucionActual) continue;
    var canonico = institucionesCanonicas[normalizarNombreIEConversatorio_(institucionActual)];
    if (!canonico) continue; // fuera de los 7 grupos de ConectaEduca — no aplica al Conversatorio
    var tecnica = String(fr[2] || "").trim();
    var clave = normalizarNombreIEConversatorio_(canonico) + "||" + tecnica.toUpperCase();
    if (clavesExistentes[clave]) continue;
    clavesExistentes[clave] = true;
    nuevasFilasResolucion.push([
      Utilities.getUuid(),
      canonico,
      tecnica,
      String(fr[3] || "").trim(),
      String(fr[4] || "").trim(),
      "", "", "", "", "", ""
    ]);
  }
  if (nuevasFilasResolucion.length) {
    hojaResolucion.getRange(hojaResolucion.getLastRow() + 1, 1, nuevasFilasResolucion.length, cabecerasConversatorioResolucion_().length)
      .setValues(nuevasFilasResolucion);
  }

  // --- Una fila de acceso por institución (idempotente, sin código) ---
  var hojaAccesos = obtenerHoja_(HOJA_CONVERSATORIO_ACCESOS_, cabecerasConversatorioAccesos_());
  var existentesAccesos = leerFilasComoObjetos_(hojaAccesos);
  var conAcceso = {};
  existentesAccesos.forEach(function (fila) {
    conAcceso[normalizarNombreIEConversatorio_(fila.INSTITUCION_EDUCATIVA)] = true;
  });
  var nuevosAccesos = [];
  Object.keys(institucionesCanonicas).forEach(function (clave) {
    if (conAcceso[clave]) return;
    nuevosAccesos.push([institucionesCanonicas[clave], "ACTIVO", "", "", "", "", ""]);
  });
  if (nuevosAccesos.length) {
    hojaAccesos.getRange(hojaAccesos.getLastRow() + 1, 1, nuevosAccesos.length, cabecerasConversatorioAccesos_().length)
      .setValues(nuevosAccesos);
  }

  return {
    ok: true,
    grupos: grupos.length,
    institucionesEnMatriz: Object.keys(institucionesCanonicas).length,
    filasResolucionAgregadas: nuevasFilasResolucion.length,
    institucionesNuevas: nuevosAccesos.length
  };
}

/** Nombres de las Instituciones Educativas disponibles en el Conversatorio, para el listado de elección. */
function obtenerInstitucionesConversatorio() {
  var hoja = obtenerHoja_(HOJA_CONVERSATORIO_ACCESOS_, cabecerasConversatorioAccesos_());
  var filas = leerFilasComoObjetos_(hoja);
  var nombres = filas
    .filter(function (fila) { return String(fila.ESTADO || "").toUpperCase() !== "BLOQUEADO"; })
    .map(function (fila) { return String(fila.INSTITUCION_EDUCATIVA || "").trim(); })
    .filter(function (n) { return n; });
  nombres.sort(function (a, b) { return a.localeCompare(b, "es"); });
  return { ok: true, instituciones: nombres };
}

/** Fila (objeto) de AccesosIEConversatorio para una institución, o null si no existe/está bloqueada. */
function buscarAccesoConversatorioPorInstitucion_(institucion) {
  institucion = String(institucion || "").trim();
  if (!institucion) return null;
  var hoja = obtenerHoja_(HOJA_CONVERSATORIO_ACCESOS_, cabecerasConversatorioAccesos_());
  var filas = leerFilasComoObjetos_(hoja);
  var buscado = normalizarNombreIEConversatorio_(institucion);
  for (var i = 0; i < filas.length; i++) {
    if (normalizarNombreIEConversatorio_(filas[i].INSTITUCION_EDUCATIVA) === buscado) return filas[i];
  }
  return null;
}

/** Grupos técnicos (los 7) a los que pertenece una institución, según ConversatorioMatriz. */
function gruposDeInstitucionConversatorio_(institucion) {
  var hoja = obtenerHoja_(HOJA_CONVERSATORIO_MATRIZ_, cabecerasConversatorioMatriz_());
  var filas = leerFilasComoObjetos_(hoja);
  var idsGrupo = {};
  filas.forEach(function (fila) {
    if (normalizarNombreIEConversatorio_(fila.INSTITUCION_EDUCATIVA) === normalizarNombreIEConversatorio_(institucion)) {
      idsGrupo[String(fila.ID_GRUPO).trim()] = true;
    }
  });
  return idsGrupo;
}

/**
 * Catálogo de técnicas únicas (los 7 grupos, sin duplicar) — para el
 * listado que se despliega cuando la IE responde "Sí" a "¿Se planea
 * incluir una nueva articulación para el año 2027?" (spec del usuario).
 */
function obtenerCatalogoTecnicasConversatorio() {
  var hoja = obtenerHoja_(HOJA_CONVERSATORIO_MATRIZ_, cabecerasConversatorioMatriz_());
  var filas = leerFilasComoObjetos_(hoja);
  var vistas = {};
  var tecnicas = [];
  filas.forEach(function (fila) {
    var tecnica = String(fila.TECNICA || "").trim();
    if (!tecnica || vistas[tecnica.toUpperCase()]) return;
    vistas[tecnica.toUpperCase()] = true;
    tecnicas.push(tecnica);
  });
  tecnicas.sort(function (a, b) { return a.localeCompare(b, "es"); });
  return { ok: true, tecnicas: tecnicas };
}

/**
 * Confirma la institución elegida de la lista y devuelve el catálogo
 * completo de los 7 grupos técnicos, marcando a cuáles pertenece esta
 * institución (spec del usuario: mostrar la información de "Grupos
 * ConectaEduca" y, ahí mismo, dejar elegir el grupo al que pertenece).
 */
function seleccionarInstitucionConversatorio(institucion) {
  var acceso = buscarAccesoConversatorioPorInstitucion_(institucion);
  if (!acceso) return { ok: false, mensaje: "Esa institución no está disponible en el Conversatorio." };

  var hojaGrupos = obtenerHoja_(HOJA_CONVERSATORIO_GRUPOS_, cabecerasConversatorioGrupos_());
  var grupos = leerFilasComoObjetos_(hojaGrupos);
  var idsDeLaIE = gruposDeInstitucionConversatorio_(acceso.INSTITUCION_EDUCATIVA);

  return {
    ok: true,
    institucion: acceso.INSTITUCION_EDUCATIVA,
    grupoElegido: acceso.GRUPO_ELEGIDO || "",
    responsableNombre: acceso.RESPONSABLE_NOMBRE || "",
    responsableRol: acceso.RESPONSABLE_ROL || "",
    acompanante: acceso.ACOMPANANTE || "",
    grupos: grupos.map(function (g) {
      return {
        idGrupo: g.ID_GRUPO,
        nombreGrupo: g.NOMBRE_GRUPO,
        numIE: g.NUM_IE,
        tecnicasIncluidas: g.TECNICAS_INCLUIDAS,
        criterioAgrupacion: g.CRITERIO_AGRUPACION,
        perteneceIE: !!idsDeLaIE[g.ID_GRUPO]
      };
    })
  };
}

/**
 * Guarda quién está diligenciando el Conversatorio en nombre de la IE
 * (spec del usuario: "hacer pantalla de responsable de llenar sesion de
 * conecta educa") — nombre, rol y, si aplica, acompañante. Alimenta el
 * banner "Respetado(a), <rol> de la IE <institución> y <acompañante>".
 */
function guardarResponsableConversatorio(institucion, nombreResponsable, rolResponsable, acompanante) {
  var acceso = buscarAccesoConversatorioPorInstitucion_(institucion);
  if (!acceso) return { ok: false, mensaje: "Esa institución no está disponible en el Conversatorio." };
  nombreResponsable = String(nombreResponsable || "").trim();
  rolResponsable = String(rolResponsable || "").trim();
  if (!nombreResponsable) return { ok: false, mensaje: "Ingrese el nombre del responsable." };
  if (!rolResponsable) return { ok: false, mensaje: "Seleccione el rol del responsable." };

  return conLock_(function () {
    upsertFila_(HOJA_CONVERSATORIO_ACCESOS_, cabecerasConversatorioAccesos_(), "INSTITUCION_EDUCATIVA", acceso.INSTITUCION_EDUCATIVA, {
      RESPONSABLE_NOMBRE: nombreResponsable,
      RESPONSABLE_ROL: rolResponsable,
      ACOMPANANTE: String(acompanante || "").trim(),
      ULTIMA_ACTIVIDAD: new Date()
    });
    return { ok: true };
  }, 10000);
}

/** Registra qué grupo técnico eligió trabajar la IE ("direccionar a grupo que pertenezca"). */
function elegirGrupoConversatorio(institucion, idGrupo) {
  var acceso = buscarAccesoConversatorioPorInstitucion_(institucion);
  if (!acceso) return { ok: false, mensaje: "Esa institución no está disponible en el Conversatorio." };
  var idsDeLaIE = gruposDeInstitucionConversatorio_(acceso.INSTITUCION_EDUCATIVA);
  idGrupo = String(idGrupo || "").trim();
  if (!idsDeLaIE[idGrupo]) {
    return { ok: false, mensaje: "Esta institución no pertenece a ese grupo." };
  }
  return conLock_(function () {
    upsertFila_(HOJA_CONVERSATORIO_ACCESOS_, cabecerasConversatorioAccesos_(), "INSTITUCION_EDUCATIVA", acceso.INSTITUCION_EDUCATIVA, {
      GRUPO_ELEGIDO: idGrupo,
      ULTIMA_ACTIVIDAD: new Date()
    });
    return { ok: true };
  }, 10000);
}

/** Filas de ConversatorioResolucionIE (técnicas SENA de la IE, con las 2 preguntas). */
function obtenerTecnicasConversatorio(institucion) {
  var acceso = buscarAccesoConversatorioPorInstitucion_(institucion);
  if (!acceso) return { ok: false, mensaje: "Esa institución no está disponible en el Conversatorio." };

  var hoja = obtenerHoja_(HOJA_CONVERSATORIO_RESOLUCION_, cabecerasConversatorioResolucion_());
  var filas = leerFilasComoObjetos_(hoja);
  var propias = filas.filter(function (fila) {
    return normalizarNombreIEConversatorio_(fila.INSTITUCION_EDUCATIVA) === normalizarNombreIEConversatorio_(acceso.INSTITUCION_EDUCATIVA);
  });

  return {
    ok: true,
    institucion: acceso.INSTITUCION_EDUCATIVA,
    tecnicas: propias.map(function (f) {
      return {
        idFila: f.ID_FILA,
        tecnicaSena: f.TECNICA_SENA || "",
        programaAprobado: f.PROGRAMA_APROBADO || "",
        resolucionDecreto: f.RESOLUCION_DECRETO || "",
        preguntaApertura2027: f.PREGUNTA_APERTURA_GRADO10_2027 || "",
        preguntaNuevaArticulacion2027: f.PREGUNTA_NUEVA_ARTICULACION_2027 || "",
        nuevaArticulacionDetalle: f.NUEVA_ARTICULACION_DETALLE || "",
        confirmado: String(f.CONFIRMADO_POR_IE || "") === "SI"
      };
    })
  };
}

/**
 * Marca una técnica como revisada y confirmada por la IE, sin cambiar
 * ningún dato (spec del usuario: "Boton de es correcto para informacion
 * de tecnica o deseamos editar esta información") — alternativa a abrir
 * el modo edición cuando los datos ya están correctos.
 */
function confirmarTecnicaConversatorio(institucion, idFila) {
  var acceso = buscarAccesoConversatorioPorInstitucion_(institucion);
  if (!acceso) return { ok: false, mensaje: "Esa institución no está disponible en el Conversatorio." };

  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_CONVERSATORIO_RESOLUCION_, cabecerasConversatorioResolucion_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_FILA", idFila);
    if (fila === -1) return { ok: false, mensaje: "Esa técnica ya no existe." };
    var filaInstitucion = String(hoja.getRange(fila, mapa["INSTITUCION_EDUCATIVA"]).getValue() || "").trim();
    if (normalizarNombreIEConversatorio_(filaInstitucion) !== normalizarNombreIEConversatorio_(acceso.INSTITUCION_EDUCATIVA)) {
      return { ok: false, mensaje: "Esa técnica no pertenece a esta institución." };
    }
    hoja.getRange(fila, mapa["CONFIRMADO_POR_IE"]).setValue("SI");
    hoja.getRange(fila, mapa["FECHA_CONFIRMACION"]).setValue(new Date());
    return { ok: true };
  }, 10000);
}

var CAMPOS_EDITABLES_TECNICA_CONVERSATORIO_ = {
  tecnicaSena: "TECNICA_SENA",
  programaAprobado: "PROGRAMA_APROBADO",
  resolucionDecreto: "RESOLUCION_DECRETO",
  preguntaApertura2027: "PREGUNTA_APERTURA_GRADO10_2027",
  preguntaNuevaArticulacion2027: "PREGUNTA_NUEVA_ARTICULACION_2027",
  nuevaArticulacionDetalle: "NUEVA_ARTICULACION_DETALLE"
};

/**
 * Autoguardado de un solo campo (técnica editable o una de las 2
 * preguntas) por fila. Spec del usuario: "poner todo en solo lectura, y
 * si hay necesidad de cambiar algo... Esos cambios deben mencionarse
 * explícitamente en el informe" — si el valor guardado YA tenía contenido
 * y el nuevo es distinto, se registra en ConversatorioCambios (nunca en
 * el llenado inicial de un campo vacío).
 */
function guardarCampoTecnicaConversatorio(institucion, idFila, campo, valor) {
  var acceso = buscarAccesoConversatorioPorInstitucion_(institucion);
  if (!acceso) return { ok: false, mensaje: "Esa institución no está disponible en el Conversatorio." };
  var columna = CAMPOS_EDITABLES_TECNICA_CONVERSATORIO_[campo];
  if (!columna) return { ok: false, mensaje: "Campo no reconocido." };
  valor = String(valor == null ? "" : valor);

  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_CONVERSATORIO_RESOLUCION_, cabecerasConversatorioResolucion_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_FILA", idFila);
    if (fila === -1) return { ok: false, mensaje: "Esa técnica ya no existe." };
    var filaInstitucion = String(hoja.getRange(fila, mapa["INSTITUCION_EDUCATIVA"]).getValue() || "").trim();
    if (normalizarNombreIEConversatorio_(filaInstitucion) !== normalizarNombreIEConversatorio_(acceso.INSTITUCION_EDUCATIVA)) {
      return { ok: false, mensaje: "Esa técnica no pertenece a esta institución." };
    }
    var valorAnterior = String(hoja.getRange(fila, mapa[columna]).getValue() || "").trim();
    if (valorAnterior && valorAnterior !== valor.trim()) {
      var tecnicaFila = String(hoja.getRange(fila, mapa["TECNICA_SENA"]).getValue() || "").trim();
      obtenerHoja_(HOJA_CONVERSATORIO_CAMBIOS_, cabecerasConversatorioCambios_()).appendRow([
        filaInstitucion, tecnicaFila, ETIQUETAS_CAMPOS_TECNICA_CONVERSATORIO_[columna] || columna,
        valorAnterior, valor.trim(), new Date()
      ]);
    }
    hoja.getRange(fila, mapa[columna]).setValue(valor);
    hoja.getRange(fila, mapa["ACTUALIZADO"]).setValue(new Date());
    return { ok: true };
  }, 10000);
}

/** Marca el Conversatorio de esta IE como completado. */
function finalizarConversatorio(institucion) {
  var acceso = buscarAccesoConversatorioPorInstitucion_(institucion);
  if (!acceso) return { ok: false, mensaje: "Esa institución no está disponible en el Conversatorio." };
  return conLock_(function () {
    upsertFila_(HOJA_CONVERSATORIO_ACCESOS_, cabecerasConversatorioAccesos_(), "INSTITUCION_EDUCATIVA", acceso.INSTITUCION_EDUCATIVA, {
      ESTADO: "COMPLETADO",
      ULTIMA_ACTIVIDAD: new Date()
    });
    return { ok: true };
  }, 10000);
}

/* ------------------------------------------------------------------ *
 * Panel de superadministración del Conversatorio (pedido del usuario:
 * "haz un login de superadmin con el codigo que ya me habias dado
 * antes") — un único administrador de la Secretaría de Educación de
 * Neiva entra con un código maestro y puede: ver el estado de todas las
 * instituciones, entrar a cualquiera de ellas (misma pantalla de
 * técnicas que usa la propia IE, con el mismo modo solo-lectura + lápiz),
 * y generar el informe consolidado con el registro de cambios.
 * ------------------------------------------------------------------ */

/** Genera (si hace falta) el código único de administrador — idempotente, nunca regenera uno ya emitido. */
function generarCodigoSuperadminConversatorio() {
  var config = getConfig();
  if (String(config.CODIGO_SUPERADMIN_CONECTAEDUCA || "").trim()) {
    return { ok: true, codigo: config.CODIGO_SUPERADMIN_CONECTAEDUCA, yaExistia: true };
  }
  var caracteres = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";
  var codigo = "SEM-";
  for (var i = 0; i < 6; i++) {
    codigo += caracteres.charAt(Math.floor(Math.random() * caracteres.length));
  }
  escribirConfig_("CODIGO_SUPERADMIN_CONECTAEDUCA", codigo);
  return { ok: true, codigo: codigo, yaExistia: false };
}

/** Valida el código maestro y devuelve el estado de todas las instituciones, para el panel. */
function validarSuperadminConversatorio(codigo) {
  codigo = String(codigo || "").trim();
  if (!codigo) return { ok: false, mensaje: "Ingrese el código de acceso." };
  var codigoConfigurado = String(getConfig().CODIGO_SUPERADMIN_CONECTAEDUCA || "").trim();
  if (!codigoConfigurado) {
    return {
      ok: false,
      mensaje: "Todavía no se ha generado un código de administrador para el Conversatorio. Ejecutar generarCodigoSuperadminConversatorio() desde el editor."
    };
  }
  if (codigo !== codigoConfigurado) return { ok: false, mensaje: "El código es incorrecto." };

  var hoja = obtenerHoja_(HOJA_CONVERSATORIO_ACCESOS_, cabecerasConversatorioAccesos_());
  var filas = leerFilasComoObjetos_(hoja);
  var instituciones = filas
    .map(function (fila) {
      return {
        institucion: String(fila.INSTITUCION_EDUCATIVA || "").trim(),
        estado: String(fila.ESTADO || "").trim() || "ACTIVO",
        grupoElegido: String(fila.GRUPO_ELEGIDO || "").trim()
      };
    })
    .filter(function (i) { return i.institucion; });
  instituciones.sort(function (a, b) { return a.institucion.localeCompare(b.institucion, "es"); });
  return { ok: true, instituciones: instituciones };
}

/**
 * Genera el informe consolidado del Conversatorio: técnicas y respuestas
 * de todas las instituciones, y al final un apartado con el registro de
 * cambios (spec del usuario: "la IE X cambió la respuesta X por Y").
 * Misma técnica que generarInformeGrupo (Informes.gs): la parte lenta
 * (crear el Doc, exportarlo a PDF) va FUERA de conLock_ — el lock de
 * Apps Script es de todo el proyecto, no solo de este informe.
 */
function generarInformeConversatorio(codigoSuperadmin) {
  var validacion = validarSuperadminConversatorio(codigoSuperadmin);
  if (!validacion.ok) return validacion;

  var config = getConfig();
  var hojaResolucion = obtenerHoja_(HOJA_CONVERSATORIO_RESOLUCION_, cabecerasConversatorioResolucion_());
  var tecnicas = leerFilasComoObjetos_(hojaResolucion);
  var porInstitucion = {};
  tecnicas.forEach(function (t) {
    var inst = String(t.INSTITUCION_EDUCATIVA || "").trim();
    if (!inst) return;
    if (!porInstitucion[inst]) porInstitucion[inst] = [];
    porInstitucion[inst].push(t);
  });
  var institucionesOrdenadas = Object.keys(porInstitucion).sort(function (a, b) { return a.localeCompare(b, "es"); });

  var hojaCambios = obtenerHoja_(HOJA_CONVERSATORIO_CAMBIOS_, cabecerasConversatorioCambios_());
  var cambios = leerFilasComoObjetos_(hojaCambios);

  var carpeta = obtenerCarpetaConectaEduca_();
  var nombreBase = "Informe Conversatorio Conecta Educa";

  var doc = DocumentApp.create(nombreBase);
  var docFile = DriveApp.getFileById(doc.getId());
  carpeta.addFile(docFile);
  try {
    DriveApp.getRootFolder().removeFile(docFile);
  } catch (e) {
    Logger.log("No fue posible quitar el Doc del informe del Conversatorio de la raíz de Drive: " + e.message);
  }

  var body = doc.getBody();
  body.clear();
  body.setPageWidth(612).setPageHeight(792).setMarginTop(30).setMarginBottom(30).setMarginLeft(50).setMarginRight(50);

  var pTitulo = body.appendParagraph(nombreBase);
  pTitulo.setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pTitulo.editAsText().setForegroundColor(COLOR_VERDE_INFORME_);
  var pSubtitulo = body.appendParagraph(config.NOMBRE_FORO || "");
  pSubtitulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setItalic(true);
  body.appendParagraph(formatearFechaLargaEs_(new Date(), true)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  if (!institucionesOrdenadas.length) {
    parrafo_(body, "Todavía no hay técnicas registradas.");
  }
  institucionesOrdenadas.forEach(function (institucion) {
    titulo1_(body, institucion);
    porInstitucion[institucion].forEach(function (t) {
      subtitulo_(body, t.TECNICA_SENA || "(sin técnica registrada)");
      parrafo_(body, "Programa aprobado por resolución: " + (t.PROGRAMA_APROBADO || "Sin información registrada."));
      parrafo_(body, "Resolución o decreto: " + (t.RESOLUCION_DECRETO || "Sin información registrada."));
      parrafo_(body, ETIQUETAS_CAMPOS_TECNICA_CONVERSATORIO_.PREGUNTA_APERTURA_GRADO10_2027 + " " +
        (t.PREGUNTA_APERTURA_GRADO10_2027 || "Sin responder."));
      var respuestaQ2 = ETIQUETAS_CAMPOS_TECNICA_CONVERSATORIO_.PREGUNTA_NUEVA_ARTICULACION_2027 + " " +
        (t.PREGUNTA_NUEVA_ARTICULACION_2027 || "Sin responder.");
      if (t.PREGUNTA_NUEVA_ARTICULACION_2027 === "Sí" && t.NUEVA_ARTICULACION_DETALLE) {
        respuestaQ2 += " (" + t.NUEVA_ARTICULACION_DETALLE + ")";
      }
      parrafo_(body, respuestaQ2);
    });
  });

  titulo1_(body, "Cambios presentados por las Instituciones Educativas en las respuestas");
  if (!cambios.length) {
    parrafo_(body, "No se registraron cambios sobre respuestas ya guardadas.");
  } else {
    cambios.forEach(function (c) {
      var tecnicaTexto = c.TECNICA_SENA ? ' (técnica "' + c.TECNICA_SENA + '")' : "";
      parrafo_(
        body,
        'La IE ' + c.INSTITUCION_EDUCATIVA + " cambió la respuesta de " + c.CAMPO + tecnicaTexto +
          ' de "' + c.VALOR_ANTERIOR + '" a "' + c.VALOR_NUEVO + '".'
      );
    });
  }

  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  pdfBlob.setName(nombreBase + ".pdf");
  var pdfFile = carpeta.createFile(pdfBlob);
  hacerPublicoSiEsPosible_(pdfFile);
  var url = pdfFile.getUrl();

  return conLock_(function () {
    var hojaInforme = obtenerHoja_(HOJA_CONVERSATORIO_INFORME_, cabecerasConversatorioInforme_());
    if (hojaInforme.getLastRow() > 1) {
      var anterior = leerFilaComoObjeto_(hojaInforme, 2, obtenerMapaCabeceras_(hojaInforme));
      try { if (anterior.DOC_ID) DriveApp.getFileById(anterior.DOC_ID).setTrashed(true); } catch (e) { Logger.log("No se pudo eliminar el Doc anterior del informe del Conversatorio: " + e.message); }
      try { if (anterior.PDF_ID) DriveApp.getFileById(anterior.PDF_ID).setTrashed(true); } catch (e) { Logger.log("No se pudo eliminar el PDF anterior del informe del Conversatorio: " + e.message); }
      hojaInforme.getRange(2, 1, hojaInforme.getLastRow() - 1, hojaInforme.getLastColumn()).clearContent();
    }
    hojaInforme.appendRow([doc.getId(), pdfFile.getId(), url, new Date()]);
    return { ok: true, url: url };
  }, 15000);
}

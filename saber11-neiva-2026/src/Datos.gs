/**
 * Datos.gs — Portal IE SABER 11 Neiva 2026
 *
 * Lectura/escritura de la hoja de cálculo real, siempre validando el
 * token de IE (o de administrador) antes de tocar nada. El docente NUNCA
 * recibe un enlace a la hoja de cálculo real — todo pasa por aquí.
 */

/** Pantalla de distribución de links: nombre + logo (base64) de cada IE, sin ningún dato sensible. */
function listarInstitucionesParaDistribucion() {
  const logos = listarLogos_();
  return CFG.IES.map(ie => {
    const logo = obtenerLogoBase64IE_(ie, logos);
    return { nombreIE: ie, logoBase64: logo.base64, logoMime: logo.mimeType };
  });
}

function filaVacia_(fila) {
  const C = CFG.COL;
  return !texto_(fila[C.docente - 1]) && !texto_(fila[C.nombre - 1]);
}

function filaARegistro_(numeroFila, fila, niveles) {
  const C = CFG.COL;
  return {
    fila: numeroFila,
    cantidad: fila[C.cantidad - 1],
    docente: texto_(fila[C.docente - 1]),
    nombre: texto_(fila[C.nombre - 1]),
    jornada: texto_(fila[C.jornada - 1]),
    genero: texto_(fila[C.genero - 1]),
    curso: texto_(fila[C.grupo - 1]),
    tipoDoc: texto_(fila[C.tipoDoc - 1]),
    numDoc: texto_(fila[C.numDoc - 1]),
    puntaje: fila[C.puntaje - 1] === '' ? '' : fila[C.puntaje - 1],
    nivel: texto_(fila[C.nivel - 1]),
    intensificacion: texto_(fila[C.intensificacion - 1]) || 'NO',
    sena: texto_(fila[C.sena - 1]) || 'NO',
    academico: texto_(fila[C.academico - 1]) || 'NO'
  };
}

/**
 * Lista de valores permitidos para una columna (lista desplegable que ya
 * tenía la hoja, puesta por el script de CONFIGURAR_TODO o a mano) — null
 * si la columna es de texto libre. Se lee de PRIMERA_FILA, asumiendo que
 * la validación es uniforme en toda la columna (así la aplica
 * aplicarFormato_ del script de la hoja).
 */
function opcionesColumna_(sh, columna) {
  const dv = sh.getRange(CFG.PRIMERA_FILA, columna).getDataValidation();
  if (!dv || dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return null;
  const valores = (dv.getCriteriaValues()[0] || []).map(v => texto_(v).toUpperCase()).filter(Boolean);
  return valores.length ? valores : null;
}

/** Listas desplegables reales de la hoja para las columnas de texto — por si el docente debe escoger de una lista exacta. */
function opcionesColumnasIE_(sh) {
  const C = CFG.COL;
  return {
    jornada: opcionesColumna_(sh, C.jornada),
    genero: opcionesColumna_(sh, C.genero),
    curso: opcionesColumna_(sh, C.grupo),
    tipoDoc: opcionesColumna_(sh, C.tipoDoc),
    numDoc: opcionesColumna_(sh, C.numDoc)
  };
}

/**
 * Filas de la IE: todas las que ya tienen algo escrito, más hasta 15
 * filas vacías a continuación para seguir agregando estudiantes sin
 * tener que traer las 500 filas completas en cada carga.
 */
function obtenerDatosIE(nombreIE, token) {
  const nombreReal = exigirAccesoIEoAdminComoIE_(nombreIE, token);
  const ss = abrirSpreadsheet_();
  const sh = ss.getSheetByName(nombreReal);
  if (!sh) throw new Error('No se encontró la hoja de "' + nombreReal + '". Avise al administrador.');
  const F = CFG.PRIMERA_FILA, L = CFG.ULTIMA_FILA, N = L - F + 1;
  const valores = sh.getRange(F, 1, N, CFG.COL.academico).getValues();
  let ultimaConDatos = -1;
  valores.forEach((fila, i) => { if (!filaVacia_(fila)) ultimaConDatos = i; });
  const limite = Math.min(N - 1, ultimaConDatos + 15);
  const filas = [];
  for (let i = 0; i <= limite; i++) filas.push(filaARegistro_(F + i, valores[i]));
  return {
    nombreIE: nombreReal,
    colorEstado: sh.getTabColor() || '#9E9E9E',
    opciones: opcionesColumnasIE_(sh),
    filas: filas
  };
}

/** Obliga a que, de las 3 columnas de tipo de grupo, quede como máximo un "SI" (mismo criterio que el script de la hoja). */
function normalizarTiposGrupo_(intensificacion, sena, academico) {
  const valores = [norm_(intensificacion), norm_(sena), norm_(academico)];
  const indiceSI = valores.findIndex(v => v === 'SI');
  return [0, 1, 2].map(i => (indiceSI >= 0 && i !== indiceSI) ? 'NO' : (valores[i] === 'SI' ? 'SI' : 'NO'));
}

/**
 * Guarda una fila de estudiante (crea o actualiza). No escribe CANTIDAD
 * (A) ni NIVEL OBTENIDO (J): son fórmulas ya puestas por el script de la
 * hoja (CONFIGURAR_TODO) que se recalculan solas con lo que se escriba
 * aquí; sobrescribirlas rompería el cálculo automático.
 */
function guardarFilaIE(nombreIE, token, fila, datos) {
  const nombreReal = exigirAccesoIEoAdminComoIE_(nombreIE, token);
  fila = Number(fila);
  if (!fila || fila < CFG.PRIMERA_FILA || fila > CFG.ULTIMA_FILA) {
    throw new Error('Fila fuera de rango.');
  }
  const ss = abrirSpreadsheet_();
  const sh = ss.getSheetByName(nombreReal);
  if (!sh) throw new Error('No se encontró la hoja de "' + nombreReal + '".');
  datos = datos || {};
  const C = CFG.COL;
  const mayus = v => texto_(v).toUpperCase();
  const [intensificacion, sena, academico] = normalizarTiposGrupo_(datos.intensificacion, datos.sena, datos.academico);
  const puntaje = datos.puntaje === '' || datos.puntaje == null ? '' : Math.max(0, Math.min(100, Number(datos.puntaje) || 0));

  // Si la columna tiene una lista desplegable real en la hoja, el valor
  // debe estar en esa lista exacta — nunca se escribe texto libre ahí,
  // porque Sheets deja la celda con un valor "inválido" en silencio (la
  // API no respeta la validación al escribir) y el admin solo lo nota
  // después, al abrir la hoja y ver el aviso rojo en esa celda.
  const opciones = opcionesColumnasIE_(sh);
  const jornada = mayus(datos.jornada), genero = mayus(datos.genero), curso = mayus(datos.curso);
  const tipoDoc = mayus(datos.tipoDoc), numDoc = mayus(datos.numDoc);
  [['jornada', jornada], ['genero', genero], ['curso', curso], ['tipoDoc', tipoDoc], ['numDoc', numDoc]].forEach(([campo, valor]) => {
    const lista = opciones[campo];
    if (valor && lista && lista.indexOf(valor) < 0) {
      throw new Error('"' + valor + '" no es un valor válido. Use uno de: ' + lista.join(', ') + '.');
    }
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sh.getRange(fila, C.docente).setValue(mayus(datos.docente));
    sh.getRange(fila, C.nombre).setValue(mayus(datos.nombre));
    sh.getRange(fila, C.jornada).setValue(jornada);
    sh.getRange(fila, C.genero).setValue(genero);
    sh.getRange(fila, C.grupo).setValue(curso);
    sh.getRange(fila, C.tipoDoc).setValue(tipoDoc);
    sh.getRange(fila, C.numDoc).setValue(numDoc);
    sh.getRange(fila, C.puntaje).setValue(puntaje);
    sh.getRange(fila, C.intensificacion).setValue(intensificacion);
    sh.getRange(fila, C.sena).setValue(sena);
    sh.getRange(fila, C.academico).setValue(academico);
    SpreadsheetApp.flush();
    const actualizada = sh.getRange(fila, 1, 1, C.academico).getValues()[0];
    return { ok: true, registro: filaARegistro_(fila, actualizada) };
  } finally {
    lock.releaseLock();
  }
}

/** Lee una hoja de resumen (RESUMEN DE ENVÍOS / REPORTE DIARIO) tal cual, para pintarla de forma genérica en el portal. */
function leerHojaComoTabla_(nombreHoja) {
  const ss = abrirSpreadsheet_();
  const sh = ss.getSheetByName(nombreHoja);
  if (!sh) return { encontrada: false, filas: [] };
  const filas = sh.getLastRow(), columnas = sh.getLastColumn();
  if (filas < 1 || columnas < 1) return { encontrada: true, filas: [] };
  const valores = sh.getRange(1, 1, filas, columnas).getDisplayValues();
  const fondos = sh.getRange(1, 1, filas, columnas).getBackgrounds();
  return { encontrada: true, filas: valores, fondos: fondos };
}

/**
 * Solo administrador — spec del usuario: "no muestres resumen de envío a
 * los usuarios" (esa hoja guarda, en la columna B, el listado de todos
 * los tokens de acceso, así que una IE nunca debe poder leerla).
 */
function obtenerResumenEnvios(token) {
  exigirAccesoAdmin_(token);
  return leerHojaComoTabla_(CFG.HOJA_RESUMEN_ENVIOS);
}
function obtenerReporteDiario(nombreIE, token) {
  exigirAccesoIEoAdmin_(nombreIE, token);
  return leerHojaComoTabla_(CFG.HOJA_REPORTE_DIARIO);
}
function exigirAccesoIEoAdmin_(nombreIE, token) {
  const comoAdmin = validarAccesoAdmin(token);
  if (comoAdmin.ok) return;
  exigirAccesoIE_(nombreIE, token);
}

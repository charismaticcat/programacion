/**
 * Reportes.gs — Portal IE SABER 11 Neiva 2026
 *
 * Mismo cálculo que "REPORTE DIARIO", "REPORTE B1 Y B+" y "Actualizar
 * resumen de envíos" del script atado a la hoja de cálculo (menú
 * REPORTES) — portado aquí porque el administrador pidió poder
 * actualizarlos desde el Portal, sin tener que abrir la hoja real ni
 * pegar el código en el editor de Apps Script de la hoja. Es la MISMA
 * lógica (mismas columnas, mismos textos, mismo formato), no una
 * versión nueva: cualquier ajuste futuro a esos reportes debería
 * hacerse en los dos lugares a la vez (o retirar esa lógica del script
 * atado a la hoja, para no mantener dos copias).
 *
 * Nota importante: al terminar, igual que el script original, se
 * reconstruye "RESUMEN DE ENVÍOS" desde cero (colores de pestaña +
 * tabla) — eso borra el acomodo de columna B con los tokens
 * (adminEscribirTokensEnResumenEnvios, Admin.gs); hay que volver a usar
 * el botón "Escribir todos los tokens…" después de actualizar
 * cualquiera de estos dos reportes.
 */

/** Igual que esHojaIE_ del script de la hoja: ¿la fila de encabezados tiene "PUNTAJE..." en la columna I? */
function esHojaIE_(sh) {
  if (sh.getMaxRows() < CFG.FILA_ENCABEZADO || sh.getMaxColumns() < CFG.COL.academico) return false;
  return norm_(sh.getRange(CFG.FILA_ENCABEZADO, CFG.COL.puntaje).getDisplayValue()).indexOf('PUNTAJE') === 0;
}

/** Una fila es inválida solo si tiene dos o más "SI" entre intensificación/SENA/académico. */
function condicionGrupoInvalida_(fila) {
  const opciones = [
    [fila[9], 'INTENSIFICACIÓN'],
    [fila[10], 'ARTICULACIÓN SENA'],
    [fila[11], 'ACADÉMICO']
  ].filter(([valor]) => norm_(valor || '') === 'SI').map(([, etiqueta]) => etiqueta);
  return opciones.length > 1 ? opciones.join(' / ') : '';
}
function mensajeGrupoInvalido_(condiciones) {
  return 'No se permite seleccionar dos tipos de grupo para un solo curso; actualmente tiene: ' + condiciones.join('; ') + '.';
}

function richTextHipervinculoIE_(ss, nombreIE) {
  const hojaIE = ss.getSheetByName(nombreIE);
  if (!hojaIE) return SpreadsheetApp.newRichTextValue().setText(nombreIE).build();
  const url = ss.getUrl() + '#gid=' + hojaIE.getSheetId() + '&range=A1';
  return SpreadsheetApp.newRichTextValue().setText(nombreIE).setLinkUrl(url).build();
}
function ponerHipervinculoIE_(ss, hojaReporte, fila, columna, nombreIE) {
  hojaReporte.getRange(fila, columna).setRichTextValue(richTextHipervinculoIE_(ss, nombreIE));
}

/** Pasa a mayúsculas el texto escrito a mano en el área dada (nunca fórmulas, nunca el texto guía, nunca NIVEL). */
function mayusculasEn_(area) {
  const vals = area.getValues(), fx = area.getFormulas();
  const c0 = area.getColumn();
  const cambios = [];
  vals.forEach((row, i) => row.forEach((v, j) => {
    if (c0 + j === CFG.COL.nivel || fx[i][j] || typeof v !== 'string' || v === CFG.MARCADOR) return;
    const up = v.toUpperCase();
    if (up !== v) cambios.push([i, j, up]);
  }));
  if (!cambios.length) return;
  if (cambios.length <= 20) {
    cambios.forEach(([i, j, up]) => area.getCell(i + 1, j + 1).setValue(up));
  } else {
    const out = vals.map((row, i) => row.map((v, j) => fx[i][j] || v));
    cambios.forEach(([i, j, up]) => { out[i][j] = up; });
    area.setValues(out);
  }
}
/** Normaliza a mayúsculas lo escrito desde la fila 36 en adelante (jornada..académico) en todas las IE. */
function normalizarMayusculasDesdeD36_(ss) {
  CFG.IES.forEach(nombreIE => {
    const sh = ss.getSheetByName(nombreIE);
    if (!sh || !esHojaIE_(sh)) return;
    const ultimaFila = Math.min(sh.getLastRow(), CFG.ULTIMA_FILA);
    if (ultimaFila < 36) return;
    mayusculasEn_(sh.getRange(36, CFG.COL.jornada, ultimaFila - 35, CFG.COL.academico - CFG.COL.jornada + 1));
  });
}

/**
 * Marca como revisadas solo las IE sin observaciones en REPORTE DIARIO
 * ni en REPORTE B1 Y B+ (colorea pestañas) y reconstruye RESUMEN DE
 * ENVÍOS con ese resultado — idéntico a actualizarMarcasRevisionIE_ del
 * script atado a la hoja.
 */
function actualizarMarcasRevisionIE_(ss) {
  const diario = ss.getSheetByName(CFG.HOJA_REPORTE_DIARIO);
  const b1 = ss.getSheetByName(CFG.HOJA_REPORTE_B1_MAS);
  const zona = Session.getScriptTimeZone();
  const fechaHora = Utilities.formatDate(new Date(), zona, 'dd/MM/yyyy HH:mm:ss');
  const notaRevision = 'Revisado completo (' + fechaHora + ')';
  const observacionesDiario = new Map();
  const filasDiario = new Map();
  const ultimaFilaDiario = diario ? diario.getLastRow() : 0;
  if (diario && ultimaFilaDiario >= 8) {
    const valoresDiario = diario.getRange(8, 1, ultimaFilaDiario - 7, 6).getDisplayValues();
    valoresDiario.forEach((fila, i) => {
      const ie = texto_(fila[0]);
      if (!ie || norm_(ie).indexOf('NO HAY INSTITUCIONES') === 0) return;
      const clave = norm_(ie);
      if (!filasDiario.has(clave)) filasDiario.set(clave, []);
      filasDiario.get(clave).push(8 + i);
      if (!observacionesDiario.has(clave)) observacionesDiario.set(clave, []);
      if (texto_(fila[5])) observacionesDiario.get(clave).push(texto_(fila[5]));
    });
  }
  const observacionesB1 = new Map();
  const filasB1 = new Map();
  const ultimaFilaB1 = b1 ? b1.getLastRow() : 0;
  let filaEncabezadoB1 = 0;
  if (b1 && ultimaFilaB1 >= 6) {
    const columnasAB = b1.getRange(1, 1, ultimaFilaB1, 2).getDisplayValues();
    filaEncabezadoB1 = columnasAB.findIndex(fila => norm_(fila[0]) === 'N° EN IE' || norm_(fila[1]) === 'INSTITUCION EDUCATIVA') + 1;
  }
  if (filaEncabezadoB1 > 6) {
    const resumenB1 = b1.getRange(6, 1, filaEncabezadoB1 - 6, 13).getDisplayValues();
    let ieActual = '';
    resumenB1.forEach((fila, i) => {
      const nombre = texto_(fila[0]);
      if (norm_(nombre) === 'TOTAL GENERAL') { ieActual = ''; return; }
      if (nombre) ieActual = nombre;
      if (!ieActual) return;
      const clave = norm_(ieActual);
      if (!filasB1.has(clave)) filasB1.set(clave, []);
      filasB1.get(clave).push(6 + i);
      if (!observacionesB1.has(clave)) observacionesB1.set(clave, []);
      if (texto_(fila[12])) observacionesB1.get(clave).push(texto_(fila[12]));
    });
  }
  const estadosEnvio = [];
  const tieneObservacionNoGrupo = observaciones => (observaciones || []).some(observacion =>
    texto_(observacion).replace(/No se permite seleccionar dos tipos de grupo para un solo curso; actualmente tiene: [^.]*\./g, ' ').trim() !== ''
  );
  CFG.IES.forEach(nombreIE => {
    const clave = norm_(nombreIE);
    const filasD = filasDiario.get(clave) || [];
    const filasR = filasB1.get(clave) || [];
    const tieneObsDiario = tieneObservacionNoGrupo(observacionesDiario.get(clave));
    const tieneObsB1 = tieneObservacionNoGrupo(observacionesB1.get(clave));
    const tieneAmbosReportes = filasD.length > 0 && filasR.length > 0;
    const hojaIE = ss.getSheetByName(nombreIE);
    let tieneDatosFuente = false;
    const cursosConTiposInvalidos = new Map();
    if (hojaIE && esHojaIE_(hojaIE)) {
      const ultimaFilaFuente = Math.min(CFG.ULTIMA_FILA, hojaIE.getLastRow());
      if (ultimaFilaFuente >= CFG.PRIMERA_FILA) {
        const fuente = hojaIE.getRange(CFG.PRIMERA_FILA, CFG.COL.docente,
          ultimaFilaFuente - CFG.PRIMERA_FILA + 1, CFG.COL.academico - CFG.COL.docente + 1).getDisplayValues();
        fuente.forEach(fila => {
          const tieneRegistro = fila.slice(1, 9).some(valor => texto_(valor) !== '');
          if (tieneRegistro) tieneDatosFuente = true;
          const tipos = [
            [fila[9], 'INTENSIFICACIÓN'],
            [fila[10], 'ARTICULACIÓN SENA'],
            [fila[11], 'ACADÉMICO']
          ].filter(([valor]) => norm_(valor) === 'SI').map(([, tipo]) => tipo);
          if (tipos.length > 1) {
            const combinacion = tipos.join(' / ');
            if (!cursosConTiposInvalidos.has(combinacion)) cursosConTiposInvalidos.set(combinacion, new Set());
            cursosConTiposInvalidos.get(combinacion).add(texto_(fila[4]) || 'SIN CURSO');
          }
        });
      }
    }
    const reportada = tieneDatosFuente || filasD.length > 0 || filasR.length > 0;
    const tieneObsGrupo = cursosConTiposInvalidos.size > 0;
    const revisadaCompleta = tieneAmbosReportes && !tieneObsDiario && !tieneObsB1 && !tieneObsGrupo;
    const estadoEnvio = !reportada ? 'REPORTE NO ENVIADO' :
      (revisadaCompleta ? 'REPORTE COMPLETO' : 'CON OBSERVACIONES O INCOMPLETO');
    const colorIE = revisadaCompleta ? '#D9EAD3' : ((tieneObsDiario || tieneObsB1 || tieneObsGrupo) ? '#FCE8E6' : '#FFFFFF');
    const colorEstado = !reportada ? '#EA4335' : (revisadaCompleta ? '#34A853' : '#F9AB00');
    if (hojaIE) hojaIE.setTabColor(colorEstado);
    const observacionesCrudas = (observacionesDiario.get(clave) || []).concat(observacionesB1.get(clave) || []);
    const partesObservacion = [];
    observacionesCrudas.forEach(observacion => {
      let limpia = texto_(observacion).replace(/No se permite seleccionar dos tipos de grupo para un solo curso; actualmente tiene: [^.]*\./g, ' ');
      limpia = limpia.replace(/(\d)\.\s+/g, '$1\u0001').replace(/\s+/g, ' ').replace(/\s+\./g, '.').trim();
      limpia.split(/(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÑ])/).forEach(parte => {
        const normalizada = texto_(parte).replace(/\u0001/g, '. ');
        if (normalizada) partesObservacion.push(normalizada);
      });
    });
    const observacionesUnicas = Array.from(new Set(partesObservacion));
    const erroresPorCurso = Array.from(cursosConTiposInvalidos.entries()).map(([combinacion, cursos]) =>
      'No se permite seleccionar dos tipos de grupo para un solo curso; actualmente tiene: ' + combinacion + '. Cursos: ' +
      Array.from(cursos).sort((a, b) => a.localeCompare(b, 'es', { numeric: true, sensitivity: 'base' })).join(', ') + '.'
    );
    const observacionesEstado = observacionesUnicas.concat(erroresPorCurso).join(' | ');
    const pendientes = [];
    if (reportada && !filasD.length) pendientes.push('FALTA GENERAR REPORTE DIARIO');
    if (reportada && !filasR.length) pendientes.push('FALTA GENERAR REPORTE B1 Y B+');
    const detalleEstado = observacionesEstado || pendientes.join(' | ') ||
      (reportada ? 'SIN OBSERVACIONES' : 'PENDIENTE DE ENVÍO');
    estadosEnvio.push({
      nombre: nombreIE,
      estado: estadoEnvio,
      detalle: detalleEstado,
      color: colorEstado,
      reportada: reportada,
      completa: revisadaCompleta
    });
    filasD.forEach(fila => {
      if (!diario) return;
      const celdaIE = diario.getRange(fila, 1);
      const celdaObs = diario.getRange(fila, 6);
      celdaIE.setBackground(colorIE);
      if (revisadaCompleta) celdaObs.setNote(notaRevision);
      else celdaObs.clearNote();
    });
    if (filasR.length) {
      if (!b1) return;
      const inicio = filasR[0], fin = filasR[filasR.length - 1];
      const celdaIE = b1.getRange(inicio, 1, fin - inicio + 1, 1);
      celdaIE.setBackground(colorIE);
      filasR.forEach(fila => {
        const celdaObs = b1.getRange(fila, 13);
        if (revisadaCompleta) celdaObs.setNote(notaRevision);
        else celdaObs.clearNote();
      });
    }
  });
  actualizarHojaResumenEnvios_(ss, estadosEnvio, fechaHora);
}

/** Reconstruye por completo "RESUMEN DE ENVÍOS" — idéntico a actualizarHojaResumenEnvios_ del script de la hoja. */
function actualizarHojaResumenEnvios_(ss, estados, fechaHora) {
  const nombreHoja = CFG.HOJA_RESUMEN_ENVIOS;
  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) hoja = ss.insertSheet(nombreHoja, 0);
  hoja.setFrozenRows(0);
  hoja.getRange(1, 1, hoja.getMaxRows(), hoja.getMaxColumns()).breakApart();
  hoja.clear();
  const completas = estados.filter(item => item.completa).length;
  const pendientes = estados.filter(item => !item.reportada).length;
  const incompletas = estados.length - completas - pendientes;
  hoja.getRange('A1:H1').merge().setValue('RESUMEN DE ENVÍOS — PRUEBA SABER 11');
  hoja.getRange('A1:H1').setBackground('#188038').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(16).setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.setRowHeight(1, 40);
  hoja.getRange('A2:H2').merge().setValue('Última actualización: ' + fechaHora);
  hoja.getRange('A2:H2').setFontStyle('italic').setHorizontalAlignment('center');
  hoja.getRange('A4:D4').setValues([['TOTAL IE', 'COMPLETAS', 'CON OBSERVACIONES / INCOMPLETAS', 'FALTAN POR ENVIAR']]);
  hoja.getRange('A5:D5').setValues([[estados.length, completas, incompletas, pendientes]]);
  hoja.getRange('A4:D4').setBackground('#D9EAD3').setFontWeight('bold').setWrap(true)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.getRange('A5:D5').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  hoja.getRange('B5').setBackground('#34A853').setFontColor('#FFFFFF');
  hoja.getRange('C5').setBackground('#F9AB00').setFontColor('#000000');
  hoja.getRange('D5').setBackground('#EA4335').setFontColor('#FFFFFF');
  hoja.getRange('E4:H4').merge().setValue('PROGRESO POR ESTADO');
  hoja.getRange('E4:H4').setBackground('#D9EAD3').setFontWeight('bold').setHorizontalAlignment('center');
  hoja.getRange('E5:H5').merge();
  const total = Math.max(estados.length, 1);
  const sep = separadorFormula_(ss);
  const opcionesBarra = sep === ','
    ? `{"charttype","bar";"max",${total};"color1","#34A853";"color2","#F9AB00";"color3","#EA4335"}`
    : `{"charttype"\\"bar";"max"\\${total};"color1"\\"#34A853";"color2"\\"#F9AB00";"color3"\\"#EA4335"}`;
  const formulaBarra = sep === ','
    ? `=SPARKLINE(B5:D5,${opcionesBarra})`
    : `=SPARKLINE(B5:D5;${opcionesBarra})`;
  hoja.getRange('E5').setFormula(formulaBarra);
  hoja.setRowHeight(5, 34);
  hoja.getRange('A7:H7').merge().setValue('VERDE: REPORTE COMPLETO SIN OBSERVACIONES   |   NARANJA: REPORTE CON OBSERVACIONES O INCOMPLETO   |   ROJO: REPORTE NO ENVIADO');
  hoja.getRange('A7:H7').setWrap(true).setFontWeight('bold').setHorizontalAlignment('center');
  hoja.getRange('A9:C9').setValues([['INSTITUCIÓN EDUCATIVA', 'ESTADO DEL REPORTE', 'OBSERVACIONES / PENDIENTES']]);
  hoja.getRange('A9:C9').setBackground('#188038').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  if (estados.length) {
    hoja.getRange(10, 1, estados.length, 3).setValues(estados.map(item => [item.nombre, item.estado, item.detalle]));
    hoja.getRange(10, 1, estados.length, 1).setRichTextValues(estados.map(item => [richTextHipervinculoIE_(ss, item.nombre)]));
    hoja.getRange(10, 1, estados.length, 1).setBackgrounds(estados.map(item => [item.color]));
    hoja.getRange(10, 2, estados.length, 1).setBackgrounds(estados.map(item => [item.color]));
    hoja.getRange(10, 1, estados.length, 3).setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, true, true, true, true, true);
    hoja.setRowHeights(10, estados.length, 36);
  }
  hoja.setColumnWidth(1, 330);
  hoja.setColumnWidth(2, 300);
  hoja.setColumnWidth(3, 720);
  hoja.setColumnWidth(4, 150);
  for (let columna = 5; columna <= 8; columna++) hoja.setColumnWidth(columna, 65);
  hoja.setTabColor('#188038');
  hoja.setFrozenRows(9);
  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(1);
}

/**
 * Agrupa por docente las filas de UNA IE (grupos reportados, cantidad de
 * estudiantes, qué le falta a cada registro, combinaciones de tipo de
 * grupo inválidas) — aislado de generarReporteDiarioSaber11_ para
 * poder usarlo también en el resumen en vivo de una sola IE
 * (obtenerResumenEnvioIE), sin tener que leer REPORTE DIARIO.
 */
function calcularDocentesIE_(sh, nombreIE) {
  const primeraFila = CFG.PRIMERA_FILA;
  const ultimaFila = Math.min(CFG.ULTIMA_FILA, sh.getLastRow());
  const docentes = {};
  let totalEstudiantesIE = 0;
  const nivelesIE = new Set();
  const esTecnicoIpc = norm_(nombreIE) === norm_('INSTITUTO TECNICO IPC ANDRES ROSA');
  if (ultimaFila < primeraFila) return { docentes, totalEstudiantesIE, nivelesIE, esTecnicoIpc };
  const datos = sh.getRange(primeraFila, CFG.COL.docente, ultimaFila - primeraFila + 1,
    CFG.COL.academico - CFG.COL.docente + 1).getDisplayValues();
  let docenteActual = '';
  datos.forEach(fila => {
    const docenteCelda = String(fila[0] || '').trim();
    const estudiante = String(fila[1] || '').trim();
    const curso = String(fila[4] || '').trim();
    const nivel = norm_(fila[8] || '');
    const esMarcador = norm_(docenteCelda) === norm_(CFG.MARCADOR);
    if (docenteCelda && !esMarcador) docenteActual = docenteCelda;
    if (!estudiante) return;
    const nombreDocente = (!esMarcador && docenteCelda) ? docenteCelda : (docenteActual || 'DOCENTE NO REGISTRADO');
    if (!docentes[nombreDocente]) {
      docentes[nombreDocente] = {
        grupos: {}, estudiantes: 0,
        faltantes: { genero: false, tipoDoc: false, numeroDoc: false, tipoGrupo: false },
        combinacionesGrupoInvalidas: {}
      };
    }
    const registroDocente = docentes[nombreDocente];
    if (esTecnicoIpc && /(CICLO|NOCTUR|SABAT)/i.test(norm_(curso))) return;
    totalEstudiantesIE++;
    if (nivel) nivelesIE.add(nivel);
    docentes[nombreDocente].estudiantes++;
    if (!String(fila[3] || '').trim()) registroDocente.faltantes.genero = true;
    if (!String(fila[5] || '').trim()) registroDocente.faltantes.tipoDoc = true;
    if (!String(fila[6] || '').trim()) registroDocente.faltantes.numeroDoc = true;
    const tipoGrupo = [fila[9], fila[10], fila[11]].map(v => norm_(v || ''));
    if (!tipoGrupo.includes('SI')) registroDocente.faltantes.tipoGrupo = true;
    const condicionInvalida = condicionGrupoInvalida_(fila);
    if (condicionInvalida) registroDocente.combinacionesGrupoInvalidas[condicionInvalida] = true;
    if (curso) docentes[nombreDocente].grupos[curso] = true;
  });
  return { docentes, totalEstudiantesIE, nivelesIE, esTecnicoIpc };
}

/**
 * Por curso de UNA IE: cuántos estudiantes tiene reportados y de qué
 * tipo es (intensificación/SENA/académico) — mismo criterio de
 * "totalesTodosPorCurso" de generarReporteB1BMas_, aislado para
 * reusarlo en el resumen en vivo de una sola IE.
 */
function calcularCursosIE_(sh) {
  const ultimaFila = Math.min(CFG.ULTIMA_FILA, sh.getLastRow());
  const totalesPorCurso = new Map();
  if (ultimaFila < CFG.PRIMERA_FILA) return totalesPorCurso;
  const datos = sh.getRange(CFG.PRIMERA_FILA, CFG.COL.docente, ultimaFila - CFG.PRIMERA_FILA + 1,
    CFG.COL.academico - CFG.COL.docente + 1).getDisplayValues();
  datos.forEach(fila => {
    const tieneRegistroEstudiante = fila.slice(1, 4).concat(fila.slice(5, 9)).some(valor => String(valor || '').trim() !== '');
    if (!tieneRegistroEstudiante) return;
    const nombreCurso = String(fila[4] || '').trim() || 'SIN CURSO';
    const datosCurso = totalesPorCurso.get(nombreCurso) || { totalReportados: 0, intensificacion: false, sena: false, academico: false };
    datosCurso.totalReportados++;
    datosCurso.intensificacion = datosCurso.intensificacion || norm_(fila[9] || '') === 'SI';
    datosCurso.sena = datosCurso.sena || norm_(fila[10] || '') === 'SI';
    datosCurso.academico = datosCurso.academico || norm_(fila[11] || '') === 'SI';
    totalesPorCurso.set(nombreCurso, datosCurso);
  });
  return totalesPorCurso;
}

/**
 * Resumen en vivo de UNA IE (spec del usuario) — reemplaza la antigua
 * pestaña "Reporte diario" (mostraba las 36 IE sin filtrar) por algo
 * propio de la IE: estado con color, cursos reportados (nombre, tipo,
 * cantidad de estudiantes), docentes que han reportado y observaciones.
 * Se calcula directo de la hoja de la IE — no depende de que REPORTE
 * DIARIO/B1+ ya se hayan actualizado.
 */
function obtenerResumenEnvioIE(nombreIE, token) {
  const nombreReal = exigirAccesoIEoAdminComoIE_(nombreIE, token);
  const ss = abrirSpreadsheet_();
  const sh = ss.getSheetByName(nombreReal);
  if (!sh) throw new Error('No se encontró la hoja de "' + nombreReal + '".');

  const { docentes, esTecnicoIpc } = calcularDocentesIE_(sh, nombreReal);
  const totalesPorCurso = calcularCursosIE_(sh);

  const observaciones = [];
  Object.keys(docentes).forEach(nombreDocente => {
    const registro = docentes[nombreDocente];
    const camposFaltantes = [
      ['genero', 'Género'], ['tipoDoc', 'Tipo de documento'],
      ['numeroDoc', 'Número de documento'], ['tipoGrupo', 'Tipo de Grupo']
    ].filter(([clave]) => registro.faltantes[clave]).map(([, etiqueta]) => etiqueta);
    if (camposFaltantes.length) observaciones.push(nombreDocente + ': falta ' + camposFaltantes.join(', ') + '.');
    const gruposInvalidos = Object.keys(registro.combinacionesGrupoInvalidas);
    if (gruposInvalidos.length) observaciones.push(nombreDocente + ': ' + mensajeGrupoInvalido_(gruposInvalidos));
  });
  if (Object.keys(docentes).some(n => norm_(n) === 'DOCENTE NO REGISTRADO')) {
    observaciones.push('Hay estudiantes sin docente asignado (DOCENTE NO REGISTRADO).');
  }
  if (esTecnicoIpc) observaciones.push('No se deben reportar grupos diferentes a grado 11 (se excluyen ciclos nocturnos y sabatinos).');

  const cursos = Array.from(totalesPorCurso.entries())
    .sort((a, b) => a[0].localeCompare(b[0], 'es', { numeric: true, sensitivity: 'base' }))
    .map(([nombreCurso, datosCurso]) => ({
      nombreCurso: nombreCurso,
      tipo: [
        datosCurso.intensificacion ? 'INTENSIFICACIÓN' : '',
        datosCurso.sena ? 'ARTICULACIÓN SENA' : '',
        datosCurso.academico ? 'ACADÉMICO' : ''
      ].filter(Boolean).join(' / ') || 'SIN DATO',
      cantidadEstudiantes: datosCurso.totalReportados
    }));

  return {
    colorEstado: sh.getTabColor() || '#9E9E9E',
    totalCursos: cursos.length,
    cursos: cursos,
    docentes: Object.keys(docentes).filter(n => norm_(n) !== 'DOCENTE NO REGISTRADO'),
    observaciones: observaciones
  };
}

/**
 * Reconstruye REPORTE DIARIO leyendo las 36 hojas de IE — idéntico a
 * generarReporteDiarioSaber11() del script de la hoja, salvo que no usa
 * SpreadsheetApp.getUi() (no existe en una app web): devuelve los
 * totales en vez de mostrar una alerta.
 */
function generarReporteDiarioSaber11_(ss) {
  normalizarMayusculasDesdeD36_(ss);
  const NOMBRE_HOJA_REPORTE = CFG.HOJA_REPORTE_DIARIO;
  let shReporte = ss.getSheetByName(NOMBRE_HOJA_REPORTE);
  if (!shReporte) shReporte = ss.insertSheet(NOMBRE_HOJA_REPORTE);
  shReporte.clear();
  shReporte.getRange('A1:F2').breakApart();
  const filtro = shReporte.getFilter();
  if (filtro) filtro.remove();

  const fechaActualizacion = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  const registros = [];
  let totalInstituciones = 0, totalDocentes = 0, totalEstudiantes = 0;

  CFG.IES.forEach(nombreIE => {
    const sh = ss.getSheetByName(nombreIE);
    if (!sh) return;
    if (!esHojaIE_(sh)) return;
    const { docentes, totalEstudiantesIE, nivelesIE, esTecnicoIpc } = calcularDocentesIE_(sh, nombreIE);
    const nombresDocentes = Object.keys(docentes);
    if (nombresDocentes.length === 0) return;
    const soloNivelesAltos = totalEstudiantesIE > 0 && nivelesIE.size > 0 &&
      Array.from(nivelesIE).every(nivel => nivel === 'B1' || nivel === 'B+');
    totalInstituciones++;
    totalDocentes += nombresDocentes.length;
    nombresDocentes.forEach(nombreDocente => {
      const registro = docentes[nombreDocente];
      const grupos = Object.keys(registro.grupos).sort((a, b) => a.localeCompare(b, 'es', { numeric: true, sensitivity: 'base' }));
      totalEstudiantes += registro.estudiantes;
      const camposFaltantes = [
        ['genero', 'Género'], ['tipoDoc', 'Tipo de documento'],
        ['numeroDoc', 'Número de documento'], ['tipoGrupo', 'Tipo de Grupo']
      ].filter(([clave]) => registro.faltantes[clave]).map(([, etiqueta]) => etiqueta);
      const observaciones = [];
      if (camposFaltantes.length) observaciones.push('Falta: ' + camposFaltantes.map((campo, i) => (i + 1) + '. ' + campo + '.').join(' '));
      if (norm_(nombreDocente) === 'DOCENTE NO REGISTRADO') observaciones.push('DOCENTE NO REGISTRADO.');
      if (camposFaltantes.includes('Tipo de Grupo')) observaciones.push('SIN DATO: no se registró el tipo de grupo.');
      if (soloNivelesAltos) observaciones.push('ATENCIÓN: Se debe reportar a todos los estudiantes.');
      const gruposInvalidos = Object.keys(registro.combinacionesGrupoInvalidas);
      if (gruposInvalidos.length) observaciones.push(mensajeGrupoInvalido_(gruposInvalidos));
      if (esTecnicoIpc) observaciones.push('No se deben reportar grupos diferentes a grado 11. Eliminar los grupos de ciclos nocturnos y sabatinos.');
      const gruposParaReporte = esTecnicoIpc ? grupos.filter(grupo => !/(CICLO|NOCTUR|SABAT)/i.test(norm_(grupo))) : grupos;
      registros.push([nombreIE, nombreDocente, gruposParaReporte.join(' & '), registro.estudiantes, fechaActualizacion, observaciones.join(' ')]);
    });
  });

  registros.sort((a, b) => {
    const ie = a[0].localeCompare(b[0], 'es', { sensitivity: 'base' });
    if (ie !== 0) return ie;
    return a[1].localeCompare(b[1], 'es', { sensitivity: 'base' });
  });

  shReporte.getRange('A1:F1').merge();
  shReporte.getRange('A1').setValue('REPORTE DIARIO DE PRUEBAS SABER 11');
  shReporte.getRange('A1').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center').setVerticalAlignment('middle');
  shReporte.setRowHeight(1, 32);
  shReporte.getRange('A2:F2').merge();
  shReporte.getRange('A2').setValue('Última actualización: ' + fechaActualizacion);
  shReporte.getRange('A2').setFontStyle('italic').setHorizontalAlignment('center');
  shReporte.getRange('A4:E4').setValues([['INSTITUCIONES', 'DOCENTES', 'ESTUDIANTES', 'ESTADO', 'ACTUALIZACIÓN']]);
  shReporte.getRange('A4:E4').setFontWeight('bold').setHorizontalAlignment('center');
  shReporte.getRange('A5:E5').setValues([[totalInstituciones, totalDocentes, totalEstudiantes, 'ACTUALIZADO', fechaActualizacion]]);
  shReporte.getRange('A5:E5').setHorizontalAlignment('center');
  const filaEncabezado = 7;
  shReporte.getRange(filaEncabezado, 1, 1, 6).setValues([[
    'INSTITUCIÓN EDUCATIVA', 'DOCENTE', 'GRUPOS REPORTADOS', 'CANTIDAD DE ESTUDIANTES', 'FECHA / HORA', 'OBSERVACIONES'
  ]]);
  shReporte.getRange(filaEncabezado, 1, 1, 6).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  if (registros.length > 0) {
    shReporte.getRange(filaEncabezado + 1, 1, registros.length, 6).setValues(registros);
    shReporte.getRange(filaEncabezado + 1, 1, registros.length, 1).setRichTextValues(
      registros.map(registro => [richTextHipervinculoIE_(ss, registro[0])]));
  } else {
    shReporte.getRange(filaEncabezado + 1, 1, 1, 6).setValues([['NO HAY INSTITUCIONES CON REPORTES', '', '', 0, fechaActualizacion, '']]);
  }
  const ultimaFilaReporte = Math.max(filaEncabezado + registros.length, filaEncabezado + 1);
  shReporte.getRange(filaEncabezado, 1, ultimaFilaReporte - filaEncabezado + 1, 6).setVerticalAlignment('middle');
  shReporte.getRange(filaEncabezado + 1, 3, ultimaFilaReporte - filaEncabezado, 1).setHorizontalAlignment('center');
  shReporte.getRange(filaEncabezado + 1, 6, ultimaFilaReporte - filaEncabezado, 1).setWrap(true).setVerticalAlignment('top');
  if (registros.length) {
    shReporte.getRange(filaEncabezado + 1, 6, registros.length, 1)
      .setBackgrounds(registros.map(registro => [registro[5] ? '#FCE8E6' : '#FFFFFF']))
      .setFontColors(registros.map(registro => [registro[5] ? '#A61C00' : '#000000']));
    shReporte.getRange(filaEncabezado + 1, 1, registros.length, 1)
      .setBackgrounds(registros.map(registro => [registro[5] ? '#FCE8E6' : '#FFFFFF']));
    shReporte.getRange(filaEncabezado + 1, 2, registros.length, 1)
      .setBackgrounds(registros.map(registro => [norm_(registro[1]) === 'DOCENTE NO REGISTRADO' ? '#FCE8E6' : '#FFFFFF']))
      .setFontColors(registros.map(registro => [norm_(registro[1]) === 'DOCENTE NO REGISTRADO' ? '#A61C00' : '#000000']));
  }
  shReporte.getRange(filaEncabezado + 1, 4, ultimaFilaReporte - filaEncabezado, 1).setHorizontalAlignment('center');
  shReporte.getRange(filaEncabezado + 1, 5, ultimaFilaReporte - filaEncabezado, 1).setHorizontalAlignment('center');
  shReporte.getRange(filaEncabezado, 1, ultimaFilaReporte - filaEncabezado + 1, 6)
    .setBorder(true, true, true, true, true, true);
  shReporte.autoResizeColumns(1, 6);
  shReporte.setColumnWidth(1, 300);
  shReporte.setColumnWidth(2, 260);
  shReporte.setColumnWidth(3, 180);
  shReporte.setColumnWidth(4, 150);
  shReporte.setColumnWidth(5, 150);
  shReporte.setColumnWidth(6, 420);
  shReporte.setFrozenRows(7);
  shReporte.getRange(filaEncabezado, 1, ultimaFilaReporte - filaEncabezado + 1, 6).createFilter();
  shReporte.getRange(5, 1, 1, 3).setNumberFormat('0');
  shReporte.getRange(filaEncabezado + 1, 4, Math.max(registros.length, 1), 1).setNumberFormat('0');

  actualizarMarcasRevisionIE_(ss);
  ss.setActiveSheet(shReporte);

  return { ok: true, totalInstituciones: totalInstituciones, totalDocentes: totalDocentes, totalEstudiantes: totalEstudiantes };
}

/**
 * Reconstruye REPORTE B1 Y B+ leyendo las 36 hojas de IE — idéntico a
 * generarReporteB1BMas() del script de la hoja.
 */
function generarReporteB1BMas_(ss) {
  normalizarMayusculasDesdeD36_(ss);
  const nombreHoja = CFG.HOJA_REPORTE_B1_MAS;
  let reporte = ss.getSheetByName(nombreHoja);
  if (!reporte) reporte = ss.insertSheet(nombreHoja);
  reporte.setFrozenRows(0);
  reporte.getRange(1, 1, reporte.getMaxRows(), reporte.getMaxColumns()).breakApart();
  reporte.clear();
  const filtroAnterior = reporte.getFilter();
  if (filtroAnterior) filtroAnterior.remove();

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  const porIE = new Map();
  CFG.IES.forEach(nombreIE => {
    const sh = ss.getSheetByName(nombreIE);
    if (!sh || !esHojaIE_(sh)) return;
    const ultimaFila = Math.min(CFG.ULTIMA_FILA, sh.getLastRow());
    const cantidadFilas = Math.max(0, ultimaFila - CFG.PRIMERA_FILA + 1);
    let datos = [], valores = [];
    if (cantidadFilas > 0) {
      const rangoDatos = sh.getRange(CFG.PRIMERA_FILA, CFG.COL.docente, cantidadFilas, CFG.COL.academico - CFG.COL.docente + 1);
      datos = rangoDatos.getDisplayValues();
      valores = rangoDatos.getValues();
    }
    const estudiantes = [];
    let totalEvaluados = 0, totalReportados = 0;
    const combinacionesGrupoInvalidas = new Set();
    let docenteActual = '';
    const gruposIntensificacion = new Map();
    const totalesTodosPorCurso = new Map();
    datos.forEach((fila, i) => {
      const docenteCelda = String(fila[0] || '').trim();
      const esMarcador = norm_(docenteCelda) === norm_(CFG.MARCADOR);
      if (docenteCelda && !esMarcador) docenteActual = docenteCelda;
      const nivel = norm_(fila[8] || '');
      const nombreEstudiante = String(fila[1] || '').trim();
      const tieneRegistroEstudiante = fila.slice(1, 4).concat(fila.slice(5, 9)).some(valor => String(valor || '').trim() !== '');
      if (tieneRegistroEstudiante) totalReportados++;
      const puntaje = String(valores[i][CFG.COL.puntaje - CFG.COL.docente] || fila[7] || '').trim().replace(',', '.');
      const tienePuntajeValido = puntaje !== '' && Number.isFinite(Number(puntaje));
      if (tieneRegistroEstudiante && (tienePuntajeValido || nivel !== '')) totalEvaluados++;
      const esIntensificacion = norm_(fila[9] || '') === 'SI';
      const condicionInvalida = tieneRegistroEstudiante ? condicionGrupoInvalida_(fila) : '';
      if (condicionInvalida) combinacionesGrupoInvalidas.add(condicionInvalida);
      if (tieneRegistroEstudiante) {
        const nombreCurso = String(fila[4] || '').trim() || 'SIN CURSO';
        const datosCurso = totalesTodosPorCurso.get(nombreCurso) || { totalReportados: 0, b1BMas: 0, intensificacion: false, sena: false, academico: false };
        datosCurso.totalReportados++;
        if (nivel === 'B1' || nivel === 'B+') datosCurso.b1BMas++;
        datosCurso.intensificacion = datosCurso.intensificacion || norm_(fila[9] || '') === 'SI';
        datosCurso.sena = datosCurso.sena || norm_(fila[10] || '') === 'SI';
        datosCurso.academico = datosCurso.academico || norm_(fila[11] || '') === 'SI';
        totalesTodosPorCurso.set(nombreCurso, datosCurso);
      }
      if (esIntensificacion && tieneRegistroEstudiante) {
        const nombreGrupo = String(fila[4] || '').trim() || 'SIN CURSO';
        const datosGrupo = gruposIntensificacion.get(nombreGrupo) || { total: 0, b1BMas: 0 };
        datosGrupo.total++;
        if (nivel === 'B1' || nivel === 'B+') datosGrupo.b1BMas++;
        gruposIntensificacion.set(nombreGrupo, datosGrupo);
      }
      if (!tieneRegistroEstudiante || (nivel !== 'B1' && nivel !== 'B+')) return;
      const tipoGrupo = [
        [fila[9], 'INTENSIFICACIÓN'], [fila[10], 'ARTICULACIÓN SENA'], [fila[11], 'ACADÉMICO']
      ].filter(([valor]) => norm_(valor || '') === 'SI').map(([, etiqueta]) => etiqueta).join(' / ') || 'SIN DATO';
      estudiantes.push([
        nombreEstudiante, String(fila[5] || '').trim(), String(fila[6] || '').trim(), String(fila[7] || '').trim(),
        nivel, tipoGrupo, (!esMarcador && docenteCelda) ? docenteCelda : (docenteActual || 'DOCENTE NO REGISTRADO')
      ]);
    });
    porIE.set(nombreIE, { estudiantes, evaluados: totalEvaluados, totalReportados, combinacionesGrupoInvalidas: Array.from(combinacionesGrupoInvalidas), gruposIntensificacion, totalesTodosPorCurso });
  });

  let totalB1BMas = 0, totalEvaluados = 0, totalIntensificacion = 0, totalB1BMasIntensificacion = 0;
  let totalReportadoCursos = 0, totalB1BMasCursos = 0;
  let sumaPorcentajesCursosIntensificacion = 0, cantidadCursosIntensificacion = 0;
  porIE.forEach(datosIE => {
    totalB1BMas += datosIE.estudiantes.length;
    totalEvaluados += datosIE.evaluados;
    datosIE.gruposIntensificacion.forEach(datosGrupo => {
      totalIntensificacion += datosGrupo.total;
      totalB1BMasIntensificacion += datosGrupo.b1BMas;
      sumaPorcentajesCursosIntensificacion += datosGrupo.total ? datosGrupo.b1BMas / datosGrupo.total : 0;
      cantidadCursosIntensificacion++;
    });
    datosIE.totalesTodosPorCurso.forEach(totalesCurso => {
      totalReportadoCursos += totalesCurso.totalReportados;
      totalB1BMasCursos += totalesCurso.b1BMas;
    });
  });
  const porcentajeGeneral = totalEvaluados ? totalB1BMas / totalEvaluados : 0;
  const porcentajeCursos = totalReportadoCursos ? totalB1BMasCursos / totalReportadoCursos : 0;
  const promedioCursosIntensificacionGeneral = cantidadCursosIntensificacion ? sumaPorcentajesCursosIntensificacion / cantidadCursosIntensificacion : '';
  const resumen = [];
  CFG.IES.forEach(nombreIE => {
    if (!porIE.has(nombreIE)) return;
    const datosIE = porIE.get(nombreIE);
    if (datosIE.totalReportados === 0) return;
    const tieneB1BMas = datosIE.estudiantes.length > 0;
    const esOliverioLara = norm_(nombreIE) === norm_('OLIVERIO LARA BORRERO');
    const porcentajeIE = datosIE.evaluados ? datosIE.estudiantes.length / datosIE.evaluados : 0;
    const cursosIntensificacionIE = Array.from(datosIE.gruposIntensificacion.values());
    const totalEstudiantesIntensificacionIE = cursosIntensificacionIE.reduce((total, grupo) => total + grupo.total, 0);
    const totalB1BMasIntensificacionIE = cursosIntensificacionIE.reduce((total, grupo) => total + grupo.b1BMas, 0);
    const promedioCursosIntensificacionIE = cursosIntensificacionIE.length
      ? cursosIntensificacionIE.reduce((suma, grupo) => suma + (grupo.total ? grupo.b1BMas / grupo.total : 0), 0) / cursosIntensificacionIE.length
      : '';
    const porcentajeIntensificacionIE = esOliverioLara && cursosIntensificacionIE.length ? 'Reporte no válido' : promedioCursosIntensificacionIE;
    const grupos = Array.from(datosIE.totalesTodosPorCurso.entries())
      .sort((a, b) => Number(!a[1].intensificacion) - Number(!b[1].intensificacion) || a[0].localeCompare(b[0], 'es', { numeric: true, sensitivity: 'base' }));
    const observacionesIE = [];
    if (!tieneB1BMas) observacionesIE.push('Pendiente por reportar estudiantes en B1 & B+');
    if (esOliverioLara) observacionesIE.push('Se deben reportar todos los niveles desde A- hasta B+.');
    if (datosIE.combinacionesGrupoInvalidas.length) observacionesIE.push(mensajeGrupoInvalido_(datosIE.combinacionesGrupoInvalidas));
    if (!grupos.length) {
      resumen.push([
        nombreIE, 'SIN CURSO', '', 0, 0, 0, 'NO', 'NO', 'NO', 'SIN DATO', '', '', '',
        tieneB1BMas ? datosIE.estudiantes.length : '', datosIE.evaluados,
        tieneB1BMas ? porcentajeIE : '', observacionesIE.join(' ')
      ]);
      return;
    }
    grupos.forEach(([nombreGrupo, datosGrupo]) => {
      const datosIntensificacion = datosIE.gruposIntensificacion.get(nombreGrupo);
      const porcentajeGrupo = datosIntensificacion && datosIntensificacion.total ? datosIntensificacion.b1BMas / datosIntensificacion.total : '';
      const tiposCurso = [
        datosGrupo.intensificacion ? 'INTENSIFICACIÓN' : '',
        datosGrupo.sena ? 'ARTICULACIÓN SENA' : '',
        datosGrupo.academico ? 'ACADÉMICO' : ''
      ].filter(Boolean);
      const observacionesGrupo = observacionesIE.slice();
      if (!tiposCurso.length) observacionesGrupo.push('SIN DATO: no se registró el tipo de grupo.');
      if (datosIntensificacion && tieneB1BMas && porcentajeGrupo === 1 && !esOliverioLara) observacionesGrupo.push('Se deben reportar todos los niveles MCER');
      const aviso = observacionesGrupo.join(' ');
      resumen.push([
        nombreIE, nombreGrupo,
        datosGrupo.intensificacion && tieneB1BMas ? porcentajeGrupo : (datosGrupo.intensificacion ? '' : 'No aplica'),
        datosGrupo.totalReportados, datosGrupo.b1BMas,
        datosGrupo.totalReportados ? datosGrupo.b1BMas / datosGrupo.totalReportados : 0,
        datosGrupo.intensificacion ? 'SI' : 'NO', datosGrupo.sena ? 'SI' : 'NO', datosGrupo.academico ? 'SI' : 'NO',
        tiposCurso.length ? tiposCurso.join(' / ') : 'SIN DATO',
        datosGrupo.intensificacion ? totalEstudiantesIntensificacionIE : '',
        datosGrupo.intensificacion ? totalB1BMasIntensificacionIE : '',
        datosGrupo.intensificacion ? porcentajeIntensificacionIE : '',
        tieneB1BMas ? datosIE.estudiantes.length : '', datosIE.evaluados, tieneB1BMas ? porcentajeIE : '', aviso
      ]);
    });
  });
  const detalle = [];
  const bloquesDetalleIE = [];
  CFG.IES.forEach(nombreIE => {
    const datosIE = porIE.get(nombreIE);
    const estudiantes = datosIE ? datosIE.estudiantes : [];
    if (estudiantes.length) bloquesDetalleIE.push({ inicio: detalle.length, cantidad: estudiantes.length, nombreIE });
    estudiantes.forEach((estudiante, i) => detalle.push([i + 1, nombreIE, estudiantes.length, ...estudiante]));
  });

  reporte.getRange('A1:M1').merge();
  reporte.getRange('A1').setValue('REPORTE DE ESTUDIANTES CON NIVEL MCER B1 Y B+');
  reporte.getRange('A1').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center').setVerticalAlignment('middle');
  reporte.setRowHeight(1, 32);
  reporte.getRange('A2:M2').merge();
  reporte.getRange('A2').setValue('Última actualización: ' + fecha).setFontStyle('italic').setHorizontalAlignment('center');

  reporte.getRange('A4:A5').merge();
  reporte.getRange('A4').setValue('1. INSTITUCIÓN EDUCATIVA');
  reporte.getRange('B4:D4').merge().setValue('TOTAL INSTITUCIÓN EDUCATIVA');
  reporte.getRange('E4:L4').merge().setValue('TOTAL POR GRUPOS');
  reporte.getRange('M4:M5').merge().setValue('OBSERVACIONES');
  reporte.getRange('B5:L5').setValues([[
    '2. TOTAL ESTUDIANTES EVALUADOS EN LA IE', '3. TOTAL ESTUDIANTES B1/B+ EN LA IE',
    '4. PORCENTAJE B1 Y B+ EN LA IE', '1. CURSO', '2. TIPO DE GRUPO',
    '3. TOTAL ESTUDIANTES REPORTADOS POR CURSO', '4. ESTUDIANTES B1/B+ REPORTADOS POR CADA CURSO',
    '5. % B1/B+ EN CADA CURSO', '6. TOTAL ESTUDIANTES DE INTENSIFICACIÓN EN LA IE EVALUADOS',
    '7. TOTAL ESTUDIANTES B1/B+ EN INTENSIFICACIÓN', '8. % B1/B+ TOTAL EN INTENSIFICACIÓN'
  ]]);
  reporte.getRange('A4:M5').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  reporte.setRowHeight(4, 32);
  reporte.setRowHeight(5, 80);
  if (resumen.length) {
    const resumenVista = resumen.map(fila => [fila[0], fila[14], fila[13], fila[15], fila[1], fila[9], fila[3], fila[4], fila[5], fila[10], fila[11], fila[12], fila[16]]);
    reporte.getRange(6, 1, resumenVista.length, 13).setValues(resumenVista);
    let inicioIE = 0;
    while (inicioIE < resumenVista.length) {
      let finIE = inicioIE + 1;
      while (finIE < resumenVista.length && norm_(resumenVista[finIE][0]) === norm_(resumenVista[inicioIE][0])) finIE++;
      const primeraFilaIE = 6 + inicioIE;
      const filasIE = finIE - inicioIE;
      if (filasIE > 1) {
        for (const columna of [2, 3, 4]) {
          reporte.getRange(primeraFilaIE + 1, columna, filasIE - 1, 1).clearContent();
          reporte.getRange(primeraFilaIE, columna, filasIE, 1).mergeVertically();
        }
        reporte.getRange(primeraFilaIE + 1, 1, filasIE - 1, 1).clearContent();
        reporte.getRange(primeraFilaIE, 1, filasIE, 1).mergeVertically();
      }
      let finIntensificacion = inicioIE;
      while (finIntensificacion < finIE && String(resumenVista[finIntensificacion][5]).split(' / ').includes('INTENSIFICACIÓN')) finIntensificacion++;
      const filasIntensificacion = finIntensificacion - inicioIE;
      if (filasIntensificacion > 1) {
        for (const columna of [10, 11, 12]) {
          reporte.getRange(6 + inicioIE + 1, columna, filasIntensificacion - 1, 1).clearContent();
          reporte.getRange(6 + inicioIE, columna, filasIntensificacion, 1).mergeVertically();
        }
      }
      const filasSinIntensificacion = finIE - finIntensificacion;
      if (filasSinIntensificacion > 0) {
        const bloqueSinIntensificacion = reporte.getRange(6 + finIntensificacion, 10, filasSinIntensificacion, 3);
        bloqueSinIntensificacion.clearContent();
        bloqueSinIntensificacion.merge();
      }
      ponerHipervinculoIE_(ss, reporte, primeraFilaIE, 1, resumenVista[inicioIE][0]);
      inicioIE = finIE;
    }
    reporte.getRange(6 + resumenVista.length, 1, 1, 13)
      .setValues([['TOTAL GENERAL', totalEvaluados, totalB1BMas, porcentajeGeneral, 'TOTAL POR GRUPOS', '', totalReportadoCursos, totalB1BMasCursos, porcentajeCursos, totalIntensificacion, totalB1BMasIntensificacion, promedioCursosIntensificacionGeneral, '']]).setFontWeight('bold');
    const rangoResumen = reporte.getRange(6, 1, resumenVista.length + 1, 13);
    rangoResumen.setWrap(true).setVerticalAlignment('middle');
    reporte.autoResizeRows(6, resumenVista.length + 1);
    reporte.getRange(6, 4, resumenVista.length + 1, 1).setNumberFormat('0.0%');
    reporte.getRange(6, 2, resumenVista.length + 1, 2).setNumberFormat('0');
    reporte.getRange(6, 7, resumenVista.length + 1, 2).setNumberFormat('0');
    reporte.getRange(6, 9, resumenVista.length + 1, 1).setNumberFormat('0.0%');
    reporte.getRange(6, 10, resumenVista.length + 1, 2).setNumberFormat('0');
    reporte.getRange(6, 12, resumenVista.length + 1, 1).setNumberFormat('0.0%');
    resumenVista.forEach((fila, i) => {
      if (String(fila[5]).split(' / ').includes('INTENSIFICACIÓN')) reporte.getRange(6 + i, 6).setBackground('#D9EAD3');
      if (norm_(fila[5]).split(' / ').includes('SIN DATO')) reporte.getRange(6 + i, 6).setBackground('#FCE8E6').setFontColor('#A61C00');
      if (fila[12]) reporte.getRange(6 + i, 13).setBackground('#FCE8E6').setFontColor('#A61C00');
    });
  } else {
    reporte.getRange('A6:M6').setValues([['SIN HOJAS DE IE', 0, 0, 0, 'No aplica', 'SIN DATO', 0, 0, 0, 0, 0, '', '']]);
    reporte.getRange('A6:M6').setWrap(true).setVerticalAlignment('middle');
    reporte.getRange('B6:C6').setNumberFormat('0');
    reporte.getRange('D6').setNumberFormat('0.0%');
  }
  const filaEncabezado = Math.max(8, 8 + resumen.length);
  reporte.getRange(filaEncabezado, 1, 1, 10).setValues([[
    'N° EN IE', 'INSTITUCIÓN EDUCATIVA', 'TOTAL B1/B+ EN IE', 'NOMBRE DEL ESTUDIANTE',
    'TIPO DE DOCUMENTO', 'DOCUMENTO DE IDENTIDAD', 'PUNTAJE', 'NIVEL MCER', 'TIPO DE GRUPO', 'DOCENTE'
  ]]).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  if (detalle.length) {
    reporte.getRange(filaEncabezado + 1, 1, detalle.length, 10).setValues(detalle);
    bloquesDetalleIE.forEach(bloque => {
      const filaIE = filaEncabezado + 1 + bloque.inicio;
      if (bloque.cantidad > 1) {
        reporte.getRange(filaIE + 1, 2, bloque.cantidad - 1, 1).clearContent();
        reporte.getRange(filaIE, 2, bloque.cantidad, 1).mergeVertically();
        reporte.getRange(filaIE + 1, 3, bloque.cantidad - 1, 1).clearContent();
        reporte.getRange(filaIE, 3, bloque.cantidad, 1).mergeVertically();
      }
      ponerHipervinculoIE_(ss, reporte, filaIE, 2, bloque.nombreIE);
    });
  } else {
    reporte.getRange(filaEncabezado + 1, 1, 1, 10).setValues([['', 'SIN REGISTROS', 0, '', '', '', '', '', '', '']]);
  }
  const ultimaFila = filaEncabezado + Math.max(detalle.length, 1);
  reporte.getRange(filaEncabezado, 1, ultimaFila - filaEncabezado + 1, 10)
    .setVerticalAlignment('middle').setWrap(true).setBorder(true, true, true, true, true, true);
  reporte.getRange(filaEncabezado + 1, 3, Math.max(detalle.length, 1), 1).setNumberFormat('0');
  if (detalle.length) {
    detalle.forEach((fila, i) => {
      const filaDestino = filaEncabezado + 1 + i;
      if (norm_(fila[8]).includes('SIN DATO')) reporte.getRange(filaDestino, 9).setBackground('#FCE8E6').setFontColor('#A61C00');
      if (norm_(fila[9]) === 'DOCENTE NO REGISTRADO') reporte.getRange(filaDestino, 10).setBackground('#FCE8E6').setFontColor('#A61C00');
    });
  }
  reporte.getRange(filaEncabezado + 1, 1, ultimaFila - filaEncabezado, 1).setHorizontalAlignment('center');
  reporte.getRange(filaEncabezado + 1, 3, ultimaFila - filaEncabezado, 1).setHorizontalAlignment('center');
  reporte.getRange(filaEncabezado + 1, 7, ultimaFila - filaEncabezado, 2).setHorizontalAlignment('center');
  reporte.autoResizeColumns(1, 13);
  [300, 300, 170, 260, 190, 230, 210, 210, 170, 250, 250, 230, 700].forEach((ancho, i) => reporte.setColumnWidth(i + 1, ancho));
  reporte.getRange(6, 13, Math.max(resumen.length, 1), 1).setWrap(true).setVerticalAlignment('top');
  reporte.setFrozenRows(5);
  reporte.getRange(filaEncabezado, 1, ultimaFila - filaEncabezado + 1, 10).createFilter();

  actualizarMarcasRevisionIE_(ss);
  ss.setActiveSheet(reporte);

  return { ok: true, instituciones: resumen.length, totalB1BMas: totalB1BMas, totalEvaluados: totalEvaluados, porcentajeGeneral: porcentajeGeneral };
}

/** RPC admin: reconstruye REPORTE DIARIO desde las 36 IE. */
function adminActualizarReporteDiario(token) {
  exigirAccesoAdmin_(token);
  const ss = abrirSpreadsheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(5 * 60 * 1000);
  try {
    return generarReporteDiarioSaber11_(ss);
  } finally {
    lock.releaseLock();
  }
}
/** RPC admin: reconstruye REPORTE B1 Y B+ desde las 36 IE. */
function adminActualizarReporteB1BMas(token) {
  exigirAccesoAdmin_(token);
  const ss = abrirSpreadsheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(5 * 60 * 1000);
  try {
    return generarReporteB1BMas_(ss);
  } finally {
    lock.releaseLock();
  }
}

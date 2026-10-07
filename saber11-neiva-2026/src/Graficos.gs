/**
 * Graficos.gs — Portal IE SABER 11 Neiva 2026
 *
 * 1) Letrero colorido en F3:K3 de cada hoja de IE (spec del usuario),
 *    junto al letrero existente "IE <nombre>" de A3:C3.
 * 2) Generación de los 5 gráficos por IE, a partir de la columna N (spec
 *    del usuario), con una tabla-resumen (datos reales, por si el
 *    administrador abre la hoja) y el gráfico insertado al lado. Se
 *    devuelve también una imagen PNG de cada gráfico para mostrarla en
 *    el portal, junto al cuadro de registro, sin que el docente necesite
 *    abrir la hoja real.
 *
 * Nota de diseño: "nivel alcanzado promedio" (gráficos 3-5) no es un
 * promedio real de un texto (A-, A1, A2, B1, B+) — se calcula el PUNTAJE
 * promedio del grupo y, a partir de ese puntaje, el nivel que le
 * correspondería (mismos rangos ICFES del script de la hoja). No se
 * grafica como número (spec del usuario: "no poner niveles por números"):
 * cada categoría es una sola barra (el puntaje promedio) y el nivel real
 * (B1, B+, …) va como parte de la etiqueta de esa barra, nunca como un
 * segundo eje/serie numérico.
 */
const TEXTO_LETRERO_GRAFICOS_ =
  'Se generarán los gráficos de la IE {IE} de acuerdo a la información ingresada. Para generar el gráfico ' +
  'de cada IE es necesario tener toda la información completa. Una vez finalizada la información dar clic ' +
  'en el botón generar gráficos y revise al lado del cuadro donde ingresó la información la imagen.';

function letreroGraficos_(sh, nombreIE) {
  const rango = sh.getRange(CFG.FILA_IE, 6, 1, 6); // F3:K3
  rango.breakApart();
  rango.merge()
    .setValue('📊 ' + TEXTO_LETRERO_GRAFICOS_.replace('{IE}', nombreSinPrefijoIE_(nombreIE)))
    .setBackground('#1A73E8').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  botonGenerarGraficos_(sh);
}

/**
 * "Botón" en L3:M3 (spec del usuario), junto al letrero — en Sheets no
 * existe un botón real que pueda llamar a un script de OTRO proyecto (el
 * Portal es un proyecto aparte de este que está atado a la hoja), así que
 * se resuelve con un =HYPERLINK: al hacer clic abre el Portal IE en una
 * pestaña nueva, donde la generación real ocurre (valida el token de la
 * IE antes de tocar nada — un botón que generara los gráficos sin pasar
 * por ahí sería una puerta trasera que salta esa validación).
 */
function botonGenerarGraficos_(sh) {
  const rango = sh.getRange(CFG.FILA_IE, 12, 1, 2); // L3:M3
  rango.breakApart();
  rango.merge()
    .setFormula(formulaHyperlink_(sh.getParent(), CFG.URL_PORTAL, '📊 GENERAR GRÁFICOS'))
    .setBackground('#34A853').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

/** Admin: pone el letrero en las 36 IE de una vez. */
function aplicarLetreroGraficosTodasLasIE(token) {
  exigirAccesoAdmin_(token);
  const ss = abrirSpreadsheet_();
  const hechas = [], problemas = [];
  CFG.IES.forEach(nombreIE => {
    const sh = ss.getSheetByName(nombreIE);
    if (!sh) { problemas.push(nombreIE + ': no existe la hoja'); return; }
    try { letreroGraficos_(sh, nombreIE); hechas.push(nombreIE); }
    catch (e) { problemas.push(nombreIE + ': ' + e.message); }
  });
  return { ok: true, hechas: hechas.length, problemas: problemas };
}

/** Nivel (texto) que corresponde a un puntaje, con los mismos rangos ICFES del script de la hoja. */
function nivelDePuntaje_(puntaje) {
  for (let i = 0; i < CFG.NIVELES.length; i++) if (puntaje <= CFG.NIVELES[i][0]) return CFG.NIVELES[i][1];
  return CFG.NIVELES[CFG.NIVELES.length - 1][1];
}

/** Tipo de grupo dominante entre varios estudiantes (para etiquetar un curso) — mismas etiquetas de los encabezados oficiales. */
const ETIQUETA_TIPO_GRUPO_ = { intensificacion: 'INTENSIFICACIÓN', sena: 'ARTICULACIÓN SENA', academico: 'ACADÉMICO' };
function tipoDominante_(registros) {
  const conteos = { intensificacion: 0, sena: 0, academico: 0 };
  registros.forEach(r => {
    if (r.intensificacion) conteos.intensificacion++;
    else if (r.sena) conteos.sena++;
    else if (r.academico) conteos.academico++;
  });
  let mejor = '', max = 0;
  Object.keys(conteos).forEach(clave => { if (conteos[clave] > max) { max = conteos[clave]; mejor = clave; } });
  return mejor ? ETIQUETA_TIPO_GRUPO_[mejor] : '';
}

/** Registros con nombre e información suficiente para graficar, leídos directo de la hoja. */
function leerRegistrosParaGraficos_(sh) {
  const F = CFG.PRIMERA_FILA, L = CFG.ULTIMA_FILA, N = L - F + 1, C = CFG.COL;
  const valores = sh.getRange(F, 1, N, C.academico).getDisplayValues();
  const puntajes = sh.getRange(F, C.puntaje, N, 1).getValues();
  const registros = [];
  valores.forEach((fila, i) => {
    if (!texto_(fila[C.nombre - 1])) return;
    registros.push({
      docente: texto_(fila[C.docente - 1]),
      jornada: texto_(fila[C.jornada - 1]),
      genero: texto_(fila[C.genero - 1]),
      curso: texto_(fila[C.grupo - 1]),
      puntaje: typeof puntajes[i][0] === 'number' ? puntajes[i][0] : null,
      nivel: texto_(fila[C.nivel - 1]),
      intensificacion: norm_(fila[C.intensificacion - 1]) === 'SI',
      sena: norm_(fila[C.sena - 1]) === 'SI',
      academico: norm_(fila[C.academico - 1]) === 'SI'
    });
  });
  return registros;
}

/**
 * {completo, faltantes} — exige jornada/género/curso/puntaje/un solo tipo
 * de grupo en cada fila con nombre. El docente NO se exige: es el dato de
 * quién reporta, no algo que los gráficos usen (spec del usuario).
 */
function verificarInformacionCompleta_(registros) {
  const faltantes = [];
  registros.forEach((r, i) => {
    const falta = [];
    if (!r.jornada) falta.push('jornada');
    if (!r.genero) falta.push('género');
    if (!r.curso) falta.push('curso');
    if (r.puntaje == null) falta.push('puntaje obtenido');
    if (!r.intensificacion && !r.sena && !r.academico) falta.push('tipo de grupo (intensificación/SENA/académico)');
    if (falta.length) faltantes.push('Estudiante ' + (i + 1) + ': falta ' + falta.join(', ') + '.');
  });
  return { completo: faltantes.length === 0, faltantes: faltantes };
}

/**
 * Mismo chequeo que al generar, pero de solo lectura — para mostrar en el
 * portal, antes de que el docente dé clic en "Generar gráficos", si la
 * información ya está completa o todavía falta algo (spec del usuario).
 */
function verificarCompletitudIE(nombreIE, token) {
  const nombreReal = exigirAccesoIEoAdminComoIE_(nombreIE, token);
  const ss = abrirSpreadsheet_();
  const sh = ss.getSheetByName(nombreReal);
  if (!sh) throw new Error('No se encontró la hoja de "' + nombreReal + '".');
  const registros = leerRegistrosParaGraficos_(sh);
  if (!registros.length) return { completo: false, faltantes: ['Todavía no hay ningún estudiante registrado en esta IE.'] };
  return verificarInformacionCompleta_(registros);
}

function totalesDe_(registros) {
  const evaluados = registros.filter(r => r.puntaje != null);
  const b1bMas = evaluados.filter(r => r.nivel === 'B1' || r.nivel === 'B+');
  const total = evaluados.length;
  return {
    totalEvaluados: total,
    totalB1BMas: b1bMas.length,
    porcentajeB1BMas: total ? Math.round((b1bMas.length / total) * 1000) / 10 : 0
  };
}

/**
 * Agrupa por la clave dada (jornada/género/curso) y calcula puntaje
 * promedio + nivel promedio por grupo. `categoria` ya trae el nivel (y,
 * si opciones.conTipo, el tipo de grupo dominante) metido en el mismo
 * texto — así el gráfico de una sola barra por categoría (spec del
 * usuario: "solo deja una columna") muestra el nivel en la propia
 * etiqueta del eje, sin necesitar un segundo eje numérico.
 */
function agruparPromedios_(registros, claveFn, opciones) {
  opciones = opciones || {};
  const grupos = new Map();
  registros.forEach(r => {
    if (r.puntaje == null) return;
    const clave = claveFn(r) || 'SIN DATO';
    if (!grupos.has(clave)) grupos.set(clave, []);
    grupos.get(clave).push(r);
  });
  return Array.from(grupos.entries())
    .sort((a, b) => a[0].localeCompare(b[0], 'es', { numeric: true, sensitivity: 'base' }))
    .map(([clave, regs]) => {
      const puntajes = regs.map(r => r.puntaje);
      const promedio = Math.round((puntajes.reduce((a, b) => a + b, 0) / puntajes.length) * 10) / 10;
      const nivel = nivelDePuntaje_(promedio);
      let etiqueta = clave;
      if (opciones.conTipo) {
        const tipo = tipoDominante_(regs);
        if (tipo) etiqueta += ' (' + tipo + ')';
      }
      return {
        categoria: etiqueta + ' — Nivel ' + nivel,
        totalEvaluados: puntajes.length,
        puntajePromedio: promedio,
        nivelPromedio: nivel
      };
    });
}

/** Escribe una tabla (título + encabezados + filas) y un gráfico de columnas a su derecha; devuelve la fila siguiente libre. */
function bloqueTablaYGrafico_(sh, filaInicio, columnaInicio, titulo, encabezados, filas, rangosGrafico, opcionesExtra) {
  const anchoTabla = encabezados.length;
  sh.getRange(filaInicio, columnaInicio, 1, anchoTabla).merge()
    .setValue(titulo).setBackground('#0B5394').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.getRange(filaInicio + 1, columnaInicio, 1, anchoTabla).setValues([encabezados])
    .setBackground('#CFE2F3').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  if (filas.length) {
    sh.getRange(filaInicio + 2, columnaInicio, filas.length, anchoTabla).setValues(filas);
  }
  const rangos = rangosGrafico.map(([colOffset, numCols]) =>
    sh.getRange(filaInicio + 1, columnaInicio + colOffset, filas.length + 1, numCols));
  let builder = sh.newChart().setChartType(Charts.ChartType.COLUMN);
  rangos.forEach(r => { builder = builder.addRange(r); });
  builder = builder.setOption('title', titulo).setOption('legend', { position: 'top' })
    .setOption('useFirstColumnAsDomain', true)
    .setPosition(filaInicio, columnaInicio + anchoTabla + 1, 0, 0);
  if (opcionesExtra) builder = opcionesExtra(builder);
  const chart = builder.build();
  sh.insertChart(chart);
  return { chart: chart, filaSiguiente: filaInicio + Math.max(filas.length + 3, 18) };
}

/**
 * Quita lo que haya quedado de una generación anterior. Se eliminan TODOS
 * los gráficos incrustados de la hoja (nunca se crean gráficos en otra
 * parte de una hoja de IE más que aquí, así que no hace falta —y es más
 * seguro no— filtrar por posición: la fila/columna de anclaje que
 * devuelve getContainerInfo() no usa la misma base que setPosition, y
 * comparar mal ambas dejaría gráficos viejos sin borrar).
 */
function limpiarZonaGraficos_(sh) {
  sh.getCharts().forEach(chart => sh.removeChart(chart));
  const ultimaFila = Math.max(sh.getMaxRows(), 200);
  const ultimaColumna = Math.max(sh.getMaxColumns(), CFG.COLUMNA_GRAFICOS + 20);
  const zona = sh.getRange(1, CFG.COLUMNA_GRAFICOS, ultimaFila, ultimaColumna - CFG.COLUMNA_GRAFICOS + 1);
  // breakApart() sobre un solo rango grande falla ("Debes seleccionar
  // todas las celdas de un intervalo combinado…") si algún combinado de
  // una generación anterior queda solo PARCIALMENTE dentro de ese
  // rectángulo. Separar cada combinado por su propio rango exacto
  // (getMergedRanges) nunca tiene ese problema.
  zona.getMergedRanges().forEach(rango => rango.breakApart());
  zona.clearContent().clearFormat();
}

/**
 * Genera (o regenera) los 5 gráficos de la IE. Exige información
 * completa primero (spec del usuario); si falta algo, no genera nada y
 * devuelve la lista de lo que falta para que el docente la complete.
 */
function generarGraficosIE(nombreIE, token) {
  const nombreReal = exigirAccesoIEoAdminComoIE_(nombreIE, token);
  const ss = abrirSpreadsheet_();
  const sh = ss.getSheetByName(nombreReal);
  if (!sh) throw new Error('No se encontró la hoja de "' + nombreReal + '".');

  const registros = leerRegistrosParaGraficos_(sh);
  if (!registros.length) {
    return { ok: false, faltantes: ['Todavía no hay ningún estudiante registrado en esta IE.'] };
  }
  const verificacion = verificarInformacionCompleta_(registros);
  if (!verificacion.completo) {
    return { ok: false, faltantes: verificacion.faltantes };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let resultado;
  try {
    limpiarZonaGraficos_(sh);
    let fila = CFG.FILA_ENCABEZADO;
    const col = CFG.COLUMNA_GRAFICOS;
    const etiquetaIE = nombreSinPrefijoIE_(nombreReal) + ', ' + CFG.ANIO;

    const bloques = [];
    // Gráfico 1: totales generales de la IE.
    const t1 = totalesDe_(registros);
    bloques.push(bloqueTablaYGrafico_(sh, fila, col,
      'TOTAL ESTUDIANTES EVALUADOS / B1-B+ / % B1-B+ EN LA IE ' + etiquetaIE + '.',
      ['CONCEPTO', 'VALOR'],
      [['TOTAL ESTUDIANTES EVALUADOS', t1.totalEvaluados], ['TOTAL ESTUDIANTES B1/B+', t1.totalB1BMas], ['% B1 Y B+ EN LA IE', t1.porcentajeB1BMas]],
      [[0, 2]]));
    fila = bloques[0].filaSiguiente;

    // Gráfico 2: igual, solo el grupo de intensificación.
    const t2 = totalesDe_(registros.filter(r => r.intensificacion));
    bloques.push(bloqueTablaYGrafico_(sh, fila, col,
      'TOTAL EVALUADOS / B1-B+ / % B1-B+ EN INTENSIFICACIÓN EN LA IE ' + etiquetaIE,
      ['CONCEPTO', 'VALOR'],
      [['TOTAL EVALUADOS EN INTENSIFICACIÓN', t2.totalEvaluados], ['TOTAL B1/B+ EN INTENSIFICACIÓN', t2.totalB1BMas], ['% B1/B+ EN INTENSIFICACIÓN', t2.porcentajeB1BMas]],
      [[0, 2]]));
    fila = bloques[1].filaSiguiente;

    // Gráfico 3: por género — una sola barra por género (puntaje promedio); el nivel va en la etiqueta, no como número.
    const porGenero = agruparPromedios_(registros, r => r.genero);
    bloques.push(bloqueTablaYGrafico_(sh, fila, col,
      'PUNTAJE Y NIVEL ALCANZADO PROMEDIO POR GÉNERO EN LA IE ' + etiquetaIE,
      ['GÉNERO (con nivel)', 'TOTAL EVALUADOS', 'PUNTAJE PROMEDIO'],
      porGenero.map(g => [g.categoria, g.totalEvaluados, g.puntajePromedio]),
      [[0, 1], [2, 1]]));
    fila = bloques[2].filaSiguiente;

    // Gráfico 4: por jornada — igual, una sola barra.
    const porJornada = agruparPromedios_(registros, r => r.jornada);
    bloques.push(bloqueTablaYGrafico_(sh, fila, col,
      'PUNTAJE Y NIVEL ALCANZADO POR JORNADA EN LA IE ' + etiquetaIE,
      ['JORNADA (con nivel)', 'TOTAL EVALUADOS', 'PUNTAJE PROMEDIO'],
      porJornada.map(g => [g.categoria, g.totalEvaluados, g.puntajePromedio]),
      [[0, 1], [2, 1]]));
    fila = bloques[3].filaSiguiente;

    // Gráfico 5: por curso (todos los reportados) — etiqueta con el tipo de grupo dominante (intensificación/articulación SENA/académico).
    const porCurso = agruparPromedios_(registros, r => r.curso, { conTipo: true });
    bloques.push(bloqueTablaYGrafico_(sh, fila, col,
      'PUNTAJE Y NIVEL ALCANZADO POR CURSO EN LA IE ' + etiquetaIE,
      ['CURSO (tipo y nivel)', 'TOTAL EVALUADOS', 'PUNTAJE PROMEDIO'],
      porCurso.map(g => [g.categoria, g.totalEvaluados, g.puntajePromedio]),
      [[0, 1], [2, 1]]));

    SpreadsheetApp.flush();
    const imagenes = bloques.map((b, i) => ({
      titulo: ['Totales de la IE', 'Totales en intensificación', 'Por género', 'Por jornada', 'Por curso'][i],
      imagenBase64: Utilities.base64Encode(b.chart.getAs('image/png').getBytes())
    }));
    resultado = { ok: true, imagenes: imagenes, generadoEn: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
  return resultado;
}

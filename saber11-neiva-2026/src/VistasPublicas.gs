/**
 * VistasPublicas.gs — Portal IE SABER 11 Neiva 2026
 *
 * Tres vistas públicas (spec del usuario), visibles ANTES de elegir una
 * IE y sin ningún token — solo lectura, nunca exponen datos de
 * identificación de los estudiantes:
 *   1) Resumen de envío de las 36 IE (docentes que han reportado,
 *      estudiantes reportados, barra de progreso general).
 *   2) Reporte estadístico de nivel B1 y B+ por IE (igual contenido que
 *      REPORTE B1 Y B+, pero sin la columna de observaciones, con el
 *      logo de cada IE) — se recalcula en vivo desde las 36 hojas, no
 *      depende de que el administrador haya actualizado esa hoja.
 *   3) Listado de estudiantes en nivel B1/B+ — sin tipo ni número de
 *      documento, y sin el docente que reportó.
 */

/**
 * 1) Docentes que han reportado + total de estudiantes + observaciones
 * personalizadas (spec del usuario: "realiza observaciones
 * personalizadas para cada dependiendo de lo que falte"), en las 36 IE,
 * con barra de progreso.
 */
function obtenerResumenEnvioTodasLasIE() {
  const ss = abrirSpreadsheet_();
  const instituciones = CFG.IES.map(nombreIE => {
    const sh = resolverHojaIE_(ss, nombreIE);
    if (!sh) return { nombreIE: nombreIE, colorEstado: '#9E9E9E', docentes: [], totalEstudiantes: 0, reportada: false, observaciones: ['No se encontró la hoja de esta institución.'] };
    const { docentes, totalEstudiantesIE, esTecnicoIpc } = calcularDocentesIE_(sh, nombreIE);
    const nombresDocentes = Object.keys(docentes).filter(n => norm_(n) !== 'DOCENTE NO REGISTRADO');
    const colorEstado = sh.getTabColor() || '#9E9E9E';
    // Detallada y con el nombre de cada docente (spec del usuario: "las
    // observaciones estaban más detalladas... decías el nombre del
    // docente y las correcciones que le correspondían a cada docente").
    const observaciones = construirObservacionesIE_(docentes, esTecnicoIpc);
    // El color de la pestaña también depende de "Actualizar REPORTE
    // DIARIO/B1 Y B+" (admin), algo que este chequeo por docente no ve;
    // si ese color dice naranja/rojo y aun así no hay nada que corregir
    // por docente, se avisa en vez de dejarlo sin explicación.
    if (!observaciones.length && colorEstado === '#EA4335') observaciones.push('Todavía no ha reportado información.');
    else if (!observaciones.length && colorEstado === '#F9AB00') {
      observaciones.push('Pendiente de confirmar: actualice "REPORTE DIARIO" y "REPORTE B1 Y B+" para reflejar el estado real.');
    }
    return {
      nombreIE: nombreIE,
      colorEstado: colorEstado,
      docentes: nombresDocentes,
      totalEstudiantes: totalEstudiantesIE,
      reportada: totalEstudiantesIE > 0,
      observaciones: observaciones
    };
  });
  const totalReportadas = instituciones.filter(i => i.reportada).length;
  return {
    instituciones: instituciones,
    totalIE: CFG.IES.length,
    totalReportadas: totalReportadas,
    porcentaje: Math.round((totalReportadas / CFG.IES.length) * 1000) / 10
  };
}

/** 2) Reporte estadístico de nivel B1 y B+ por IE (con logo, sin observaciones) — calculado en vivo. */
function obtenerReporteB1BMasPublico_() {
  const ss = abrirSpreadsheet_();
  const logos = listarLogos_();
  const instituciones = [];
  let totalEvaluadosGeneral = 0, totalB1BMasGeneral = 0;
  CFG.IES.forEach(nombreIE => {
    const sh = resolverHojaIE_(ss, nombreIE);
    if (!sh || !esHojaIE_(sh)) return;
    const datos = leerFilasIE_(sh);
    if (!datos.length) return;
    let totalEvaluados = 0, totalB1BMas = 0;
    // Agrupado por curso + jornada (spec del usuario: "discrimina por
    // jornada... no juntes cursos de misma nomenclatura") — dos "1101"
    // de jornadas distintas son grupos distintos, nunca se suman juntos.
    const porCurso = new Map();
    datos.forEach(fila => {
      const nivel = norm_(fila[8] || '');
      const tieneRegistroEstudiante = fila.slice(1, 4).concat(fila.slice(5, 9)).some(v => String(v || '').trim() !== '');
      if (!tieneRegistroEstudiante) return;
      const puntaje = String(fila[7] || '').trim().replace(',', '.');
      const evaluado = (puntaje !== '' && Number.isFinite(Number(puntaje))) || nivel !== '';
      const esB1BMas = nivel === 'B1' || nivel === 'B+';
      if (evaluado) totalEvaluados++;
      if (evaluado && esB1BMas) totalB1BMas++;
      const nombreCurso = String(fila[4] || '').trim() || 'SIN CURSO';
      const jornada = String(fila[2] || '').trim() || 'SIN JORNADA';
      const claveCurso = nombreCurso + '||' + jornada;
      const dc = porCurso.get(claveCurso) ||
        { curso: nombreCurso, jornada: jornada, totalReportados: 0, b1BMas: 0, intensificacion: false, sena: false, academico: false };
      dc.totalReportados++;
      if (evaluado && esB1BMas) dc.b1BMas++;
      dc.intensificacion = dc.intensificacion || norm_(fila[9] || '') === 'SI';
      dc.sena = dc.sena || norm_(fila[10] || '') === 'SI';
      dc.academico = dc.academico || norm_(fila[11] || '') === 'SI';
      porCurso.set(claveCurso, dc);
    });
    if (!porCurso.size) return; // nada reportado todavía en esta IE: no mostrarla en el público
    const logo = obtenerLogoBase64IE_(nombreIE, logos);
    const cursos = Array.from(porCurso.values())
      .sort((a, b) => a.curso.localeCompare(b.curso, 'es', { numeric: true, sensitivity: 'base' }) ||
        a.jornada.localeCompare(b.jornada, 'es', { sensitivity: 'base' }))
      .map(dc => ({
        curso: dc.curso,
        jornada: dc.jornada,
        tipo: [dc.intensificacion ? 'INTENSIFICACIÓN' : '', dc.sena ? 'ARTICULACIÓN SENA' : '', dc.academico ? 'ACADÉMICO' : '']
          .filter(Boolean).join(' / ') || 'SIN DATO',
        totalReportados: dc.totalReportados,
        b1BMas: dc.b1BMas,
        porcentaje: dc.totalReportados ? Math.round((dc.b1BMas / dc.totalReportados) * 1000) / 10 : 0
      }));
    totalEvaluadosGeneral += totalEvaluados;
    totalB1BMasGeneral += totalB1BMas;
    instituciones.push({
      nombreIE: nombreIE,
      logoBase64: logo.base64,
      logoMime: logo.mimeType,
      totalEvaluados: totalEvaluados,
      totalB1BMas: totalB1BMas,
      porcentaje: totalEvaluados ? Math.round((totalB1BMas / totalEvaluados) * 1000) / 10 : 0,
      cursos: cursos
    });
  });
  return {
    instituciones: instituciones,
    totalEvaluados: totalEvaluadosGeneral,
    totalB1BMas: totalB1BMasGeneral,
    porcentajeGeneral: totalEvaluadosGeneral ? Math.round((totalB1BMasGeneral / totalEvaluadosGeneral) * 1000) / 10 : 0
  };
}

/** 3) Listado de estudiantes en nivel B1/B+ — sin tipo/número de documento ni el docente que reportó. */
function obtenerListadoEstudiantesB1BMasPublico_() {
  const ss = abrirSpreadsheet_();
  const estudiantes = [];
  CFG.IES.forEach(nombreIE => {
    const sh = resolverHojaIE_(ss, nombreIE);
    if (!sh || !esHojaIE_(sh)) return;
    const datos = leerFilasIE_(sh);
    if (!datos.length) return;
    datos.forEach(fila => {
      const nivel = norm_(fila[8] || '');
      if (nivel !== 'B1' && nivel !== 'B+') return;
      const nombreEstudiante = String(fila[1] || '').trim();
      if (!nombreEstudiante) return;
      const tipoGrupo = [
        [fila[9], 'INTENSIFICACIÓN'], [fila[10], 'ARTICULACIÓN SENA'], [fila[11], 'ACADÉMICO']
      ].filter(([v]) => norm_(v || '') === 'SI').map(([, t]) => t).join(' / ') || 'SIN DATO';
      estudiantes.push({
        nombreIE: nombreIE,
        nombreEstudiante: nombreEstudiante,
        curso: String(fila[4] || '').trim(),
        tipoGrupo: tipoGrupo,
        nivel: nivel
      });
    });
  });
  estudiantes.sort((a, b) =>
    a.nombreIE.localeCompare(b.nombreIE, 'es', { sensitivity: 'base' }) ||
    a.nombreEstudiante.localeCompare(b.nombreEstudiante, 'es', { sensitivity: 'base' }));
  return { estudiantes: estudiantes, total: estudiantes.length, anio: CFG.ANIO };
}

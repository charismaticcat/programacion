/**
 * LimpiezaDatos.gs — Portal IE SABER 11 Neiva 2026
 *
 * Herramientas de limpieza/calidad de datos para el administrador (spec
 * del usuario): nunca duplicar el nombre del docente en filas seguidas,
 * sugerir el género faltante aprendiendo de los nombres ya registrados,
 * y eliminar los estudiantes de un curso puntual que no se deba
 * reportar (p. ej. ciclos nocturnos/sabatinos).
 */

/** Nombre de docente que "aplica" en cada fila por arrastre — igual que calcularDocentesIE_ (Reportes.gs). */
function limpiarDocentesDuplicadosEnHoja_(sh) {
  const F = CFG.PRIMERA_FILA;
  const ultimaFila = Math.min(CFG.ULTIMA_FILA, sh.getLastRow());
  if (ultimaFila < F) return 0;
  const N = ultimaFila - F + 1;
  const rango = sh.getRange(F, CFG.COL.docente, N, 1);
  const valores = rango.getValues();
  let anterior = '';
  let cambios = 0;
  const salida = valores.map(fila => {
    const actual = texto_(fila[0]);
    if (actual && norm_(actual) === norm_(anterior)) { cambios++; return ['']; }
    if (actual) anterior = actual;
    return [actual];
  });
  if (cambios) rango.setValues(salida);
  return cambios;
}

/**
 * Admin: recorre las 36 IE y deja el nombre del docente solo en la
 * primera fila de cada bloque, borrando cualquier repetición exacta en
 * las filas siguientes — spec del usuario: "nunca se debe duplicar un
 * nombre de docente". No toca el correo (nota de la celda) ni ningún
 * otro dato del estudiante.
 */
function adminLimpiarDocentesDuplicados(token) {
  exigirAccesoAdmin_(token);
  const ss = abrirSpreadsheet_();
  let total = 0;
  const detalle = [];
  CFG.IES.forEach(nombreIE => {
    const sh = ss.getSheetByName(nombreIE);
    if (!sh) return;
    const cambios = limpiarDocentesDuplicadosEnHoja_(sh);
    if (cambios) { total += cambios; detalle.push(nombreIE + ': ' + cambios); }
  });
  return { ok: true, totalCambios: total, detalle: detalle };
}

/**
 * Admin: borra (deja en blanco) nombre, documento, puntaje y demás
 * datos de los estudiantes reportados en un curso puntual, dentro de
 * una IE — p. ej. "elimina a todos los estudiantes de grupo C6" en el
 * Instituto Técnico IPC Andrés Rosa. El docente no se toca, y nunca se
 * sobrescribe NIVEL (columna con fórmula propia de la hoja).
 */
function adminEliminarEstudiantesPorCurso(token, nombreIE, curso) {
  exigirAccesoAdmin_(token);
  const clave = norm_(nombreIE);
  const idx = CFG.IES.map(norm_).indexOf(clave);
  if (idx < 0) throw new Error('Institución no reconocida.');
  const nombreReal = CFG.IES[idx];
  const cursoClave = norm_(curso);
  if (!cursoClave) throw new Error('Escriba el curso a eliminar.');
  const ss = abrirSpreadsheet_();
  const sh = ss.getSheetByName(nombreReal);
  if (!sh) throw new Error('No se encontró la hoja de "' + nombreReal + '".');
  const F = CFG.PRIMERA_FILA, C = CFG.COL;
  const ultimaFila = Math.min(CFG.ULTIMA_FILA, sh.getLastRow());
  if (ultimaFila < F) return { ok: true, eliminados: 0 };
  const N = ultimaFila - F + 1;

  // nombre..puntaje (nunca toca NIVEL, que es una fórmula) y aparte
  // intensificación/SENA/académico — mismo criterio que guardarFilaIE
  // (Datos.gs). El docente (columna aparte) no se toca.
  const rangoDatos = sh.getRange(F, C.nombre, N, C.puntaje - C.nombre + 1);
  const rangoTipos = sh.getRange(F, C.intensificacion, N, C.academico - C.intensificacion + 1);
  const valoresDatos = rangoDatos.getValues();
  const valoresTipos = rangoTipos.getValues();
  const colCurso = C.grupo - C.nombre;
  let eliminados = 0;
  valoresDatos.forEach((fila, i) => {
    if (norm_(fila[colCurso]) !== cursoClave) return;
    eliminados++;
    for (let j = 0; j < fila.length; j++) fila[j] = '';
    valoresTipos[i] = ['NO', 'NO', 'NO'];
  });
  if (eliminados) {
    rangoDatos.setValues(valoresDatos);
    rangoTipos.setValues(valoresTipos);
  }
  return { ok: true, eliminados: eliminados };
}

/**
 * Diccionario de nombres comunes en español/Colombia con género
 * conocido (spec del usuario: "aprende de los nombres que ya existen...
 * haz tu propio machine learning"). En vez de un modelo entrenado (no
 * aplica dentro de Apps Script), se combinan tres señales, en este
 * orden: 1) este diccionario curado; 2) lo que YA está registrado en
 * las 36 IE (si un nombre no está en el diccionario pero ya aparece
 * varias veces con el mismo género en datos reales, se usa ese); 3)
 * como último recurso, la terminación del nombre (-A suele ser
 * femenino, -O suele ser masculino en español). Si nada de esto da una
 * respuesta confiable, se deja sin sugerir — el docente elige a mano.
 */
const NOMBRES_FEMENINOS_ = ['ANA', 'ANDREA', 'ALEJANDRA', 'ALEXANDRA', 'AMPARO', 'ANGELA', 'ANGIE', 'ARACELY',
  'AURORA', 'BEATRIZ', 'BLANCA', 'BRIGITTE', 'CAMILA', 'CARMEN', 'CAROLINA', 'CECILIA', 'CLAUDIA', 'CONSTANZA',
  'CONSUELO', 'DAMARIS', 'DANA', 'DANIELA', 'DANNA', 'DAYANA', 'DEICY', 'DEISY', 'DIANA', 'DOLORES', 'DORIS',
  'EDNA', 'ELENA', 'EMMA', 'ESPERANZA', 'ESTEFANIA', 'ESTHER', 'FERNANDA', 'FLOR', 'GABRIELA', 'GLADYS', 'GLORIA',
  'GUADALUPE', 'HELENA', 'INES', 'IRENE', 'ISABELLA', 'JENNIFER', 'JESSICA', 'JUANA', 'JUDITH', 'JULIANA', 'KAREN',
  'KARLA', 'KATHERINE', 'KELLY', 'LAURA', 'LEIDY', 'LESLY', 'LIDA', 'LIDIA', 'LIGIA', 'LORENA', 'LUCELA', 'LUCIA',
  'LUZ', 'MABEL', 'MARCELA', 'MARGARITA', 'MARIA', 'MARIANA', 'MARTHA', 'MELISSA', 'MILENA', 'MIRIAM', 'MONICA',
  'MYRIAM', 'NANCY', 'NATALIA', 'NICOLE', 'NOEMI', 'NORMA', 'NUBIA', 'OLGA', 'PAOLA', 'PATRICIA', 'PAULA', 'PILAR',
  'RAQUEL', 'REBECA', 'ROSA', 'RUTH', 'SAMANTHA', 'SANDRA', 'SARA', 'SARAH', 'SHIRLEY', 'SILVIA', 'SOCORRO',
  'SOFIA', 'STEFANY', 'STEPHANIE', 'TATIANA', 'TERESA', 'VALENTINA', 'VALERIA', 'VANESSA', 'VERONICA', 'VICTORIA',
  'VIVIANA', 'XIMENA', 'XIOMARA', 'YAMILE', 'YENI', 'YENILCE', 'YENNY', 'YESENIA', 'YOLANDA', 'YULI', 'YULIANA',
  'YURANY', 'YURI', 'ZULMA'];
const NOMBRES_MASCULINOS_ = ['ADRIAN', 'ALBERTO', 'ALEJANDRO', 'ALVARO', 'ANDRES', 'ANTONIO', 'ARMANDO', 'ARTURO',
  'BRAYAN', 'BRYAN', 'CARLOS', 'CESAR', 'CRISTIAN', 'DANIEL', 'DARIO', 'DAVID', 'DEIBY', 'DEIVY', 'DIEGO',
  'DUVAN', 'EDGAR', 'EDUARDO', 'EDWIN', 'ELKIN', 'EMILIO', 'ENRIQUE', 'ERNESTO', 'FABIAN', 'FELIPE', 'FERNANDO',
  'FRANCISCO', 'FREDDY', 'FREDY', 'GERMAN', 'GUILLERMO', 'GUSTAVO', 'HAROLD', 'HECTOR', 'HENRY', 'HUGO', 'IVAN',
  'JAIME', 'JAIR', 'JAIRO', 'JEFFERSON', 'JEISON', 'JHON', 'JOHN', 'JONATHAN', 'JORGE', 'JOSE', 'JUAN', 'JULIAN',
  'KEVIN', 'LEONARDO', 'LUIS', 'MANUEL', 'MARIO', 'MARTIN', 'MATEO', 'MAURICIO', 'MIGUEL', 'NELSON', 'NICOLAS',
  'ORLANDO', 'OSCAR', 'PABLO', 'PEDRO', 'RAFAEL', 'RAUL', 'RICARDO', 'ROBERTO', 'RODOLFO', 'RODRIGO', 'SAMUEL',
  'SANTIAGO', 'SEBASTIAN', 'SIMON', 'STIVEN', 'STEVEN', 'TOMAS', 'VICTOR', 'WILLIAM', 'WILMER', 'WILSON', 'YAIR',
  'YEFERSON', 'YEISON', 'YESID'];

/** Sugerencia de género para un primer nombre, o '' si no hay señal confiable. */
function generoSegunNombre_(primerNombre, mapaAprendido) {
  const clave = norm_(primerNombre).split(' ')[0];
  if (!clave) return '';
  if (NOMBRES_FEMENINOS_.indexOf(clave) >= 0) return 'FEMENINO';
  if (NOMBRES_MASCULINOS_.indexOf(clave) >= 0) return 'MASCULINO';
  const aprendido = mapaAprendido[clave];
  if (aprendido) {
    const total = aprendido.FEMENINO + aprendido.MASCULINO;
    if (total >= 2) {
      if (aprendido.FEMENINO / total >= 0.8) return 'FEMENINO';
      if (aprendido.MASCULINO / total >= 0.8) return 'MASCULINO';
    }
  }
  if (/A$/.test(clave)) return 'FEMENINO';
  if (/O$/.test(clave)) return 'MASCULINO';
  return '';
}

/** Recorre las 36 IE y "aprende" de los géneros ya puestos a mano (primer nombre -> conteo por género). */
function construirMapaGeneroAprendido_(ss) {
  const mapa = {};
  CFG.IES.forEach(nombreIE => {
    const sh = ss.getSheetByName(nombreIE);
    if (!sh) return;
    const F = CFG.PRIMERA_FILA;
    const ultimaFila = Math.min(CFG.ULTIMA_FILA, sh.getLastRow());
    if (ultimaFila < F) return;
    const N = ultimaFila - F + 1;
    const valores = sh.getRange(F, CFG.COL.nombre, N, CFG.COL.genero - CFG.COL.nombre + 1).getDisplayValues();
    valores.forEach(fila => {
      const nombreCompleto = texto_(fila[0]);
      const genero = norm_(fila[CFG.COL.genero - CFG.COL.nombre]);
      if (!nombreCompleto || (genero !== 'FEMENINO' && genero !== 'MASCULINO')) return;
      const clave = norm_(nombreCompleto).split(' ')[0];
      if (!clave) return;
      if (!mapa[clave]) mapa[clave] = { FEMENINO: 0, MASCULINO: 0 };
      mapa[clave][genero]++;
    });
  });
  return mapa;
}

/** Texto exacto que acepta la lista desplegable real de Género en esta hoja, para la sugerencia femenina/masculina. */
function opcionGeneroReal_(sh, esFemenino) {
  const opciones = opcionesColumna_(sh, CFG.COL.genero);
  if (!opciones) return esFemenino ? 'FEMENINO' : 'MASCULINO';
  const candidatas = esFemenino ? ['FEMENINO', 'F', 'MUJER'] : ['MASCULINO', 'M', 'HOMBRE'];
  const encontrada = opciones.find(op => candidatas.indexOf(op) >= 0);
  return encontrada || opciones[0];
}

/**
 * Admin: sugiere y escribe el género de estudiantes sin género, en las
 * 36 IE — nunca sobrescribe uno que ya esté puesto. Spec del usuario:
 * queda marcado con un aviso fijo en el encabezado "Género" del Portal
 * para que siempre se revise lo autoseleccionado.
 */
function adminAutocompletarGenero(token) {
  exigirAccesoAdmin_(token);
  const ss = abrirSpreadsheet_();
  const mapaAprendido = construirMapaGeneroAprendido_(ss);
  let total = 0;
  const detalle = [];
  CFG.IES.forEach(nombreIE => {
    const sh = ss.getSheetByName(nombreIE);
    if (!sh) return;
    const F = CFG.PRIMERA_FILA;
    const ultimaFila = Math.min(CFG.ULTIMA_FILA, sh.getLastRow());
    if (ultimaFila < F) return;
    const N = ultimaFila - F + 1;
    const rango = sh.getRange(F, CFG.COL.nombre, N, CFG.COL.genero - CFG.COL.nombre + 1);
    const valores = rango.getValues();
    const offsetGenero = CFG.COL.genero - CFG.COL.nombre;
    let cambios = 0;
    valores.forEach(fila => {
      const nombreCompleto = texto_(fila[0]);
      const generoActual = texto_(fila[offsetGenero]);
      if (!nombreCompleto || generoActual) return;
      const sugerido = generoSegunNombre_(nombreCompleto, mapaAprendido);
      if (!sugerido) return;
      fila[offsetGenero] = opcionGeneroReal_(sh, sugerido === 'FEMENINO');
      cambios++;
    });
    if (cambios) {
      rango.setValues(valores);
      total += cambios;
      detalle.push(nombreIE + ': ' + cambios);
    }
  });
  return { ok: true, totalCambios: total, detalle: detalle };
}

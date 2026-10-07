/**
 * Config.gs — Portal IE SABER 11 Neiva 2026
 *
 * App web INDEPENDIENTE (proyecto de Apps Script aparte, NO el script atado
 * a la hoja de cálculo) que lee y escribe la MISMA hoja de cálculo
 * "REPORTE DE RESULTADOS PRUEBA SABER 11º.2026" (CFG.SPREADSHEET_ID) vía
 * SpreadsheetApp.openById — nunca toca el script atado a la hoja (menú
 * REPORTES, CONFIGURAR_TODO, etc.), que sigue intacto para el administrador.
 *
 * Por qué un proyecto aparte: la hoja de cálculo original se compartía
 * como "cualquiera con el enlace puede editar" — cualquier persona con el
 * link veía y editaba las 36 pestañas de IE. Compartir por pestaña no es
 * posible en Sheets (la protección de rangos solo bloquea edición, nunca
 * oculta contenido). La única forma real de que cada IE vea y edite SOLO
 * su información es que nunca reciba acceso directo a la hoja: en vez de
 * eso, entra a esta app web (ejecuta como el propietario, acceso
 * "cualquiera"), que filtra todo del lado del servidor según el token que
 * ingresó. Antes de poner esto en producción, el archivo original debe
 * pasar de "cualquiera con el enlace" a "restringido" — si no, el link
 * viejo sigue siendo una puerta trasera sin filtro.
 */
const CFG = {
  SPREADSHEET_ID: '1xYaRpD6VxA6fz628zuPQaB5vJj0JCHuWjZtns4Gpz14',
  ANIO: 2026, // para los títulos de los gráficos ("EN LA IE <nombre>, 2026")

  // Spec del usuario: acceso automático de las IE — con solo elegir su
  // institución, entra directo a sus datos (sin token ni correo).
  // exigirAccesoIEoAdminComoIE_ (Auth.gs) solo valida que la IE exista.
  // El token por IE queda deshabilitado (TOKENS_PORTAL_IE no se borra,
  // así que puede volver a usarse si esto se revierte); el único
  // candado real que sigue activo es el token de administrador.
  ACCESO_IE_AUTOMATICO: true,
  ACCESO_TOKENS_IE_HABILITADO: false,
  CARPETA_LOGOS: '1QVfDyYjhjX5H60U7SyeLtikodQbhGu1B',
  // IE cuyo logo no se detectó bien por nombre — ID del archivo de Drive puesto a mano (spec del usuario).
  LOGOS_MANUALES: {
    'INEM JULIAN MOTTA SALAS': '1AResGxiz_RYP7hcWpQzjnVaC0lQ1ee0Z',
    'JAIRO MOSQUERA MORENO': '1AResGxiz_RYP7hcWpQzjnVaC0lQ1ee0Z',
    'SANTA LIBRADA': '1M2q_Pe0JLkCBdc78rdOsAfHvghcNMvYm'
  },

  HOJA_RESUMEN_ENVIOS: 'RESUMEN DE ENVÍOS',
  HOJA_REPORTE_DIARIO: 'REPORTE DIARIO',
  HOJA_REPORTE_B1_MAS: 'REPORTE B1 Y B+',
  HOJA_TOKENS: 'TOKENS_PORTAL_IE', // hoja de control oculta, creada por este proyecto en la misma hoja de cálculo

  // Texto guía que el script de la hoja escribe en la celda del docente
  // (desaparece al escribir encima) — Reportes.gs lo trata como vacío,
  // igual que el script atado a la hoja.
  MARCADOR: 'Escriba una sola vez su nombre por el grupo que reportará',

  FILA_IE: 3,
  FILA_ENCABEZADO: 4,
  PRIMERA_FILA: 5,
  ULTIMA_FILA: 504,
  COLUMNA_GRAFICOS: 14, // N — donde empiezan las tablas-resumen y los gráficos de cada IE

  COL: {
    cantidad: 1, docente: 2, nombre: 3, jornada: 4, genero: 5, grupo: 6,
    tipoDoc: 7, numDoc: 8, puntaje: 9, nivel: 10, intensificacion: 11, sena: 12, academico: 13
  },
  ENCABEZADOS: ['CANTIDAD', 'ESCRIBA AQUI EL NOMBRE DEL DOCENTE QUE REPORTA', 'NOMBRE COMPLETO DE ESTUDIANTE',
    'JORNADA', 'GÉNERO', 'CURSO', 'TIPO DE DOCUMENTO', 'NÚMERO DE DOCUMENTO', 'PUNTAJE OBTENIDO',
    'NIVEL OBTENIDO', 'GRUPO DE INTENSIFICACIÓN', 'GRUPO DE ARTICULACIÓN SENA', 'GRUPO ACADÉMICO'],

  // [puntaje máximo, nivel, color de fondo, color de letra] — igual que el script de la hoja.
  NIVELES: [
    [47, 'A-', '#D32F2F', '#FFFFFF'],
    [57, 'A1', '#E65100', '#FFFFFF'],
    [67, 'A2', '#FFB74D', '#000000'],
    [78, 'B1', '#A5D6A7', '#000000'],
    [100, 'B+', '#1B5E20', '#FFFFFF']
  ],

  // URL del despliegue como app web — usada para el botón "Generar gráficos"
  // que se pone en la hoja real (L3:M3, ver letreroGraficos_ en Graficos.gs).
  URL_PORTAL: 'https://script.google.com/macros/s/AKfycbwUxLvJ_arl63vRzDci65M6AlUJf79o_BAIpHvJ28L67wgOEre4-AjGTTep2Pq_qSYv/exec',

  // Nombres EXACTOS de las pestañas — "CLARETIANO GUSTAVO TORRES PARRA" sin
  // el prefijo "I.E." (la pestaña real del archivo no lo tiene, aunque el
  // script original de la hoja sí lo escribía así en su propia lista).
  IES: [
    'AGUSTIN CODAZZI', 'AIPECITO', 'ANGEL MARIA PAREDES', 'ATANASIO GIRARDOT', 'CEINAR', 'CHAPINERO',
    'CLARETIANO GUSTAVO TORRES PARRA', 'DEPARTAMENTAL TIERRA DE PROMISIÓN', 'EDUARDO SANTOS', 'EL CAGUAN',
    'EL LIMONAR', 'ENRIQUE OLAYA HERRERA', 'ESCUELA NORMAL SUPERIOR', 'GABRIEL GARCIA MARQUEZ',
    'HUMBERTO TAFUR CHARRY', 'INEM JULIAN MOTTA SALAS', 'JAIRO MORERA LIZCANO', 'JAIRO MOSQUERA MORENO',
    'JOSE EUSTASIO RIVERA', 'JUAN DE CABRERA', 'LICEO DE SANTA LIBRADA', 'LUIS IGNACIO ANDRADE',
    'MARIA AUXILIADORA FORTALECILLAS', 'MISAEL PASTRANA BORRERO', 'OLIVERIO LARA BORRERO', 'PROMOCION SOCIAL',
    'RICARDO BORRERO ALVAREZ', 'ROBERTO DURAN ALVIRA', 'RODRIGO LARA BONILLA', 'SAN ANTONIO DE ANACONIA',
    'SAN LUIS BELTRAN', 'SAN MIGUEL ARCANGEL', 'SANTA LIBRADA', 'SANTA TERESA',
    'INSTITUTO TECNICO IPC ANDRES ROSA', 'TECNICO SUPERIOR'
  ]
};

function abrirSpreadsheet_() {
  return SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
}

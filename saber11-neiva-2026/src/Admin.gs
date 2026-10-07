/**
 * Admin.gs — Portal IE SABER 11 Neiva 2026
 *
 * Pantalla de Administrador (spec del usuario: "solo yo tenga acceso"),
 * gateada por el token de administrador (Auth.gs), nunca por el correo
 * de quien visita — la app corre siempre como el propietario del script
 * (acceso "cualquiera"), así que no hay una identidad de visitante
 * confiable para comparar contra un correo.
 *
 * Lo que NO está aquí: "actualizar resumen de envíos" sigue siendo el
 * menú REPORTES del script atado a la hoja de cálculo (el que ya tenía
 * el administrador) — ese script lee REPORTE DIARIO y REPORTE B1 Y B+
 * para recalcular el estado de cada IE; duplicar esa lógica en un
 * proyecto aparte sería mantener dos copias de lo mismo. El panel de
 * aquí solo gestiona lo que es propio de este portal: tokens de acceso
 * por IE y el letrero de gráficos.
 */

/** Lista de IE con su token y su estado (color de pestaña) — para distribuir los enlaces. */
function adminListarTokens(token) {
  exigirAccesoAdmin_(token);
  const ss = abrirSpreadsheet_();
  const mapa = asegurarTokensIE_(ss);
  return CFG.IES.map(ie => {
    const sh = ss.getSheetByName(ie);
    return {
      nombreIE: ie,
      token: mapa[norm_(ie)] || '',
      colorEstado: sh ? (sh.getTabColor() || '#9E9E9E') : '#9E9E9E',
      existeHoja: !!sh
    };
  });
}

/**
 * Pone el token de acceso de cada IE en la columna B de RESUMEN DE
 * ENVÍOS, justo al frente del nombre de la IE (columna A) — spec del
 * usuario. La tabla original (actualizarHojaResumenEnvios_ del script
 * atado a la hoja) ya usaba esa columna para "ESTADO DEL REPORTE" y la
 * C para "OBSERVACIONES / PENDIENTES"; para no perder esa información
 * se corre un lugar a la derecha: B (estado) -> C, C (observaciones) ->
 * D, dejando B libre para el token. El token de administrador (no es
 * de ninguna IE en particular) se deja en un letrero aparte justo
 * debajo de la tabla.
 *
 * Esa misma hoja se recrea por completo (`hoja.clear()`) cada vez que
 * el administrador usa "Actualizar resumen de envíos" (menú REPORTES)
 * — eso borra este acomodo también, así que hay que volver a usar este
 * botón después (de ahí el botón, no una escritura automática que
 * quedaría desactualizada sin avisar).
 */
function adminEscribirTokensEnResumenEnvios(token) {
  exigirAccesoAdmin_(token);
  const ss = abrirSpreadsheet_();
  const hoja = ss.getSheetByName(CFG.HOJA_RESUMEN_ENVIOS);
  if (!hoja) {
    throw new Error('No existe la hoja "' + CFG.HOJA_RESUMEN_ENVIOS + '" todavía — ejecute primero ' +
      '"Actualizar resumen de envíos" desde el menú REPORTES de la hoja de cálculo.');
  }
  const mapa = asegurarTokensIE_(ss);
  const FILA_ENCABEZADO = 9;
  const FILA_INICIO = 10; // primera fila de IE en la tabla de actualizarHojaResumenEnvios_

  // ¿Cuántas IE tiene realmente el listado (columna A, desde la fila 10)?
  const filasDisponibles = Math.max(hoja.getLastRow() - FILA_INICIO + 1, 1);
  const columnaA = hoja.getRange(FILA_INICIO, 1, filasDisponibles, 1).getDisplayValues();
  let total = 0;
  while (total < columnaA.length && texto_(columnaA[total][0])) total++;
  if (!total) {
    throw new Error('La tabla de instituciones está vacía. Ejecute primero "Actualizar resumen de envíos" desde el menú REPORTES.');
  }
  const nombres = columnaA.slice(0, total).map(f => f[0]);
  const estados = hoja.getRange(FILA_INICIO, 2, total, 1).getValues();
  const coloresEstado = hoja.getRange(FILA_INICIO, 2, total, 1).getBackgrounds();
  const observaciones = hoja.getRange(FILA_INICIO, 3, total, 1).getValues();

  // Corre el contenido existente: observaciones (C) -> D, estado (B) -> C.
  hoja.getRange(FILA_INICIO, 3, total, 1).copyTo(hoja.getRange(FILA_INICIO, 4, total, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  hoja.getRange(FILA_INICIO, 4, total, 1).setValues(observaciones);
  hoja.getRange(FILA_INICIO, 3, total, 1).setValues(estados).setBackgrounds(coloresEstado);

  // Columna B, ahora libre: el token de cada IE, frente a su nombre.
  hoja.getRange(FILA_INICIO, 2, total, 1)
    .setValues(nombres.map(nombre => [mapa[norm_(nombre)] || '']))
    .setBackground('#FFFFFF').setFontColor('#000000').setFontWeight('bold')
    .setFontFamily('Courier New').setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Encabezados: mismo corrimiento.
  hoja.getRange(FILA_ENCABEZADO, 3).copyTo(hoja.getRange(FILA_ENCABEZADO, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  const headerB = hoja.getRange(FILA_ENCABEZADO, 2).getValue();
  const headerC = hoja.getRange(FILA_ENCABEZADO, 3).getValue();
  hoja.getRange(FILA_ENCABEZADO, 4).setValue(headerC);
  hoja.getRange(FILA_ENCABEZADO, 3).setValue(headerB);
  hoja.getRange(FILA_ENCABEZADO, 2).setValue('TOKEN DE ACCESO');

  hoja.setColumnWidth(2, 170);
  hoja.setColumnWidth(3, 300);
  hoja.setColumnWidth(4, 720);

  // Token de administrador: letrero aparte justo debajo de la tabla.
  // Antes de escribirlo, borra cualquier bloque de una versión anterior
  // de este botón (quedaba más abajo, en la fila 46, 47 o 50 según la
  // versión).
  const filaAdmin = FILA_INICIO + total;
  hoja.getRange(filaAdmin, 1, 60, 8).breakApart().clearContent().clearFormat();
  hoja.getRange(filaAdmin, 1, 1, 2).merge()
    .setValue('🔑 TOKEN DE ADMINISTRADOR: ' + obtenerTokenAdmin_())
    .setBackground('#0B5394').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  return { ok: true, total: total, filaAdmin: filaAdmin };
}

/**
 * Redirección automática al abrir la hoja de cálculo real (spec del
 * usuario): nadie debería trabajar directo ahí, todo pasa por el
 * Portal. Como este es un proyecto APARTE del script atado a la hoja
 * (ver nota en Config.gs), no se edita ese script — en vez de eso, este
 * proyecto instala su PROPIO disparador instalable de "al abrir" sobre
 * la hoja real, con el mismo acceso que ya tiene (SpreadsheetApp). Se
 * ejecuta una sola vez (botón de administrador); después, cualquiera
 * que abra el archivo original ve un aviso que lo manda de inmediato al
 * Portal (con un botón por si el navegador bloquea la redirección
 * automática de un iframe hacia la pestaña de arriba).
 */
function alAbrirHojaReal_(e) {
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>body{font-family:Arial,sans-serif;text-align:center;padding:30px 24px;color:#202124;}' +
    'a.boton{display:inline-block;margin-top:18px;background:#1A73E8;color:#fff;padding:14px 28px;' +
    'border-radius:6px;text-decoration:none;font-weight:bold;font-size:1rem;}</style></head><body>' +
    '<p>Este archivo se gestiona desde el <b>Portal IE — SABER 11º Neiva ' + CFG.ANIO + '</b>.<br>No se debe editar directamente aquí.</p>' +
    '<p>Si no es redirigido automáticamente, dé clic en el botón:</p>' +
    '<a class="boton" href="' + CFG.URL_PORTAL + '" target="_top">Ir al Portal IE</a>' +
    '<script>try{top.location.href="' + CFG.URL_PORTAL + '";}catch(err){}</script>' +
    '</body></html>'
  ).setWidth(420).setHeight(230);
  SpreadsheetApp.getUi().showModalDialog(html, 'Portal IE — SABER 11º Neiva ' + CFG.ANIO);
}

/** Admin: instala (o reinstala, de forma idempotente) el disparador de arriba. */
function adminInstalarRedireccionHojaReal(token) {
  exigirAccesoAdmin_(token);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'alAbrirHojaReal_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('alAbrirHojaReal_').forSpreadsheet(CFG.SPREADSHEET_ID).onOpen().create();
  return { ok: true };
}

function adminRegenerarTokenIE(token, nombreIE) {
  exigirAccesoAdmin_(token);
  const clave = norm_(nombreIE);
  if (CFG.IES.map(norm_).indexOf(clave) < 0) throw new Error('Institución no reconocida.');
  const ss = abrirSpreadsheet_();
  const hoja = ss.getSheetByName(CFG.HOJA_TOKENS) || (asegurarTokensIE_(ss), ss.getSheetByName(CFG.HOJA_TOKENS));
  const datos = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 0), 2).getValues();
  const nuevo = generarTokenAleatorio_();
  let fila = datos.findIndex(f => norm_(f[0]) === clave);
  if (fila < 0) {
    hoja.appendRow([CFG.IES[CFG.IES.map(norm_).indexOf(clave)], nuevo]);
  } else {
    hoja.getRange(fila + 2, 2).setValue(nuevo);
  }
  return { ok: true, token: nuevo };
}

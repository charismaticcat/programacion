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
 * Escribe, en la columna B de RESUMEN DE ENVÍOS, el token de
 * administrador y el de las 36 IE (spec del usuario) — justo debajo del
 * listado de las 36 IE que arma actualizarHojaResumenEnvios_ del script
 * atado a la hoja (ese listado termina en la fila 45), con un título en
 * azul bien visible para no tener que buscarlo. Esa misma función hace
 * `hoja.clear()` sobre TODA la hoja cada vez que el administrador usa
 * "Actualizar resumen de envíos" (menú REPORTES) — eso también borra
 * esta lista, así que hay que volver a escribirla después (de ahí el
 * botón, no una escritura automática que se quedaría desactualizada sin
 * avisar).
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
  const FILA_INICIO = 47; // justo debajo del listado de las 36 IE (termina en la fila 45)
  // Limpia cualquier escritura anterior (en esta posición o en la fila 50
  // que usaba una versión previa) antes de volver a escribir.
  hoja.getRange(46, 2, 50, 1).breakApart().clearContent().clearFormat();
  hoja.getRange(FILA_INICIO, 2)
    .setValue('🔑 TOKENS DE ACCESO (administrador + 36 IE)')
    .setBackground('#0B5394').setFontColor('#FFFFFF').setFontWeight('bold');
  const lineas = [['ADMINISTRADOR: ' + obtenerTokenAdmin_()]];
  CFG.IES.forEach(ie => lineas.push([ie + ': ' + (mapa[norm_(ie)] || '')]));
  hoja.getRange(FILA_INICIO + 1, 2, lineas.length, 1).setValues(lineas);
  return { ok: true, fila: FILA_INICIO, total: lineas.length };
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

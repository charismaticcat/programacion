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

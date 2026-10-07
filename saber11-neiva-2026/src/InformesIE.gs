/**
 * InformesIE.gs — Portal IE SABER 11 Neiva 2026
 *
 * Descarga del "informe" de una IE (spec del usuario: "por cada IE debe
 * generar un informe con los gráficos... PDF o Excel"). No se construye
 * un documento nuevo desde cero: se exporta tal cual la propia hoja de
 * la IE (título, estudiantes y la zona de tablas + gráficos que arma
 * Graficos.gs desde la columna N) usando los exportadores nativos de
 * Google Sheets — así el informe es exactamente lo que hay en la hoja
 * real, con sus gráficos ya incrustados.
 *
 * - PDF: exporta esa única pestaña (por gid) directamente.
 * - Excel: Sheets no permite exportar una sola pestaña como .xlsx desde
 *   una URL; se copia la pestaña a una hoja de cálculo temporal (con una
 *   sola pestaña) y se exporta esa, que luego se borra.
 *
 * Solo para el administrador (botón "Descargar informe" en su panel).
 */

/** PDF de una sola pestaña, apaisado y ajustado al ancho — incluye los gráficos incrustados. */
function exportarHojaComoPDF_(sh) {
  const url = 'https://docs.google.com/spreadsheets/d/' + CFG.SPREADSHEET_ID + '/export' +
    '?format=pdf&gid=' + sh.getSheetId() +
    '&portrait=false&size=A4&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false' +
    '&horizontal_alignment=CENTER&vertical_alignment=TOP';
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  return resp.getBlob();
}

/** .xlsx de una sola pestaña: se copia a una hoja de cálculo temporal (con solo esa pestaña) y se exporta esa. */
function exportarHojaComoXLSX_(sh, nombreArchivo) {
  const temp = SpreadsheetApp.create(nombreArchivo);
  const tempId = temp.getId();
  try {
    const copiada = sh.copyTo(temp);
    copiada.setName(sh.getName());
    temp.getSheets().filter(s => s.getSheetId() !== copiada.getSheetId()).forEach(s => temp.deleteSheet(s));
    SpreadsheetApp.flush();
    const url = 'https://docs.google.com/spreadsheets/d/' + tempId + '/export?format=xlsx';
    const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    return resp.getBlob();
  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

/**
 * La propia IE descarga su informe (PDF o Excel) — spec del usuario:
 * "otro [botón] que diga generar y descargar informe" en la pantalla de
 * la IE, sin tener que pasar por el administrador. Mismo contenido que
 * adminDescargarInformeIE, pero exige el token de la IE (o de
 * administrador) en vez del token de administrador exclusivamente.
 */
function descargarInformeIE(nombreIE, token, formato) {
  const nombreReal = exigirAccesoIEoAdminComoIE_(nombreIE, token);
  const ss = abrirSpreadsheet_();
  const sh = resolverHojaIE_(ss, nombreReal);
  if (!sh) throw new Error('No se encontró la hoja de "' + nombreReal + '".');
  if (!sh.getCharts().length) {
    throw new Error('Todavía no hay gráficos generados. Genere los gráficos primero.');
  }
  const nombreBase = 'Informe ' + nombreSinPrefijoIE_(nombreReal) + ' - SABER 11 ' + CFG.ANIO;
  const esExcel = formato === 'xlsx';
  const blob = esExcel ? exportarHojaComoXLSX_(sh, nombreBase) : exportarHojaComoPDF_(sh);
  return {
    ok: true,
    archivoBase64: Utilities.base64Encode(blob.getBytes()),
    mimeType: esExcel ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/pdf',
    nombreArchivo: nombreBase + (esExcel ? '.xlsx' : '.pdf')
  };
}

/** Admin: descarga el informe (PDF o Excel) de una IE — exige que ya tenga los gráficos generados. */
function adminDescargarInformeIE(token, nombreIE, formato) {
  exigirAccesoAdmin_(token);
  const clave = norm_(nombreIE);
  const idx = CFG.IES.map(norm_).indexOf(clave);
  if (idx < 0) throw new Error('Institución no reconocida.');
  const nombreReal = CFG.IES[idx];
  const ss = abrirSpreadsheet_();
  const sh = resolverHojaIE_(ss, nombreReal);
  if (!sh) throw new Error('No se encontró la hoja de "' + nombreReal + '".');
  if (!sh.getCharts().length) {
    throw new Error('"' + nombreReal + '" todavía no tiene gráficos generados. Genere los gráficos primero (botón "Generar gráficos").');
  }
  const nombreBase = 'Informe ' + nombreSinPrefijoIE_(nombreReal) + ' - SABER 11 ' + CFG.ANIO;
  const esExcel = formato === 'xlsx';
  const blob = esExcel ? exportarHojaComoXLSX_(sh, nombreBase) : exportarHojaComoPDF_(sh);
  return {
    ok: true,
    archivoBase64: Utilities.base64Encode(blob.getBytes()),
    mimeType: esExcel ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/pdf',
    nombreArchivo: nombreBase + (esExcel ? '.xlsx' : '.pdf')
  };
}

/** Admin: descarga el REPORTE B1 Y B+ completo (las 36 IE) — exige que ya se haya generado/actualizado. */
function adminDescargarReporteB1BMas(token, formato) {
  exigirAccesoAdmin_(token);
  const ss = abrirSpreadsheet_();
  const sh = ss.getSheetByName(CFG.HOJA_REPORTE_B1_MAS);
  if (!sh) {
    throw new Error('Todavía no existe "' + CFG.HOJA_REPORTE_B1_MAS + '". Use primero "Actualizar REPORTE B1 Y B+".');
  }
  const nombreBase = 'REPORTE B1 Y B+ - SABER 11 ' + CFG.ANIO;
  const esExcel = formato === 'xlsx';
  const blob = esExcel ? exportarHojaComoXLSX_(sh, nombreBase) : exportarHojaComoPDF_(sh);
  return {
    ok: true,
    archivoBase64: Utilities.base64Encode(blob.getBytes()),
    mimeType: esExcel ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/pdf',
    nombreArchivo: nombreBase + (esExcel ? '.xlsx' : '.pdf')
  };
}

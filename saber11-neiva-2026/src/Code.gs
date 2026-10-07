/**
 * Code.gs — Portal IE SABER 11 Neiva 2026
 *
 * Una sola pantalla (Index.html) con 3 vistas en el cliente: distribución
 * de links, vista de IE, y administrador — todas por `google.script.run`,
 * nunca por parámetros de doGet (así no queda ningún dato en la URL). Un
 * solo link para todos: cada IE elige su logo y escribe su token.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal IE — SABER 11 Neiva 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function incluir_(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

/**
 * Envoltorio único para cada función que el cliente llama — nunca deja
 * que un error de Apps Script (mensaje técnico en inglés, stack trace)
 * llegue al navegador tal cual; siempre {ok:false, mensaje} legible.
 */
function ejecutarRpcSeguro_(fn) {
  try {
    const r = fn();
    return (r && typeof r === 'object' && 'ok' in r) ? r : { ok: true, datos: r };
  } catch (e) {
    return { ok: false, mensaje: e.message || 'Ocurrió un error inesperado.' };
  }
}

function rpcListarInstituciones() {
  return ejecutarRpcSeguro_(() => listarInstitucionesParaDistribucion());
}
function rpcValidarAccesoIE(nombreIE, token) {
  return ejecutarRpcSeguro_(() => validarAccesoIE(nombreIE, token));
}
function rpcValidarAccesoAdmin(token) {
  return ejecutarRpcSeguro_(() => validarAccesoAdmin(token));
}
function rpcObtenerDatosIE(nombreIE, token) {
  return ejecutarRpcSeguro_(() => obtenerDatosIE(nombreIE, token));
}
function rpcVerificarCompletitudIE(nombreIE, token) {
  return ejecutarRpcSeguro_(() => verificarCompletitudIE(nombreIE, token));
}
function rpcGuardarFilaIE(nombreIE, token, fila, datos) {
  return ejecutarRpcSeguro_(() => guardarFilaIE(nombreIE, token, fila, datos));
}
function rpcObtenerResumenEnvios(token) {
  return ejecutarRpcSeguro_(() => obtenerResumenEnvios(token));
}
function rpcAdminEscribirTokensEnResumenEnvios(token) {
  return ejecutarRpcSeguro_(() => adminEscribirTokensEnResumenEnvios(token));
}
function rpcObtenerResumenEnvioIE(nombreIE, token) {
  return ejecutarRpcSeguro_(() => obtenerResumenEnvioIE(nombreIE, token));
}
function rpcGenerarGraficosIE(nombreIE, token) {
  return ejecutarRpcSeguro_(() => generarGraficosIE(nombreIE, token));
}
function rpcAdminListarTokens(token) {
  return ejecutarRpcSeguro_(() => adminListarTokens(token));
}
function rpcAdminRegenerarTokenIE(token, nombreIE) {
  return ejecutarRpcSeguro_(() => adminRegenerarTokenIE(token, nombreIE));
}
function rpcAdminAplicarLetreroGraficos(token) {
  return ejecutarRpcSeguro_(() => aplicarLetreroGraficosTodasLasIE(token));
}
function rpcAdminUrlHojaReal(token) {
  return ejecutarRpcSeguro_(() => { exigirAccesoAdmin_(token); return { url: abrirSpreadsheet_().getUrl() }; });
}
function rpcAdminActualizarReporteDiario(token) {
  return ejecutarRpcSeguro_(() => adminActualizarReporteDiario(token));
}
function rpcAdminActualizarReporteB1BMas(token) {
  return ejecutarRpcSeguro_(() => adminActualizarReporteB1BMas(token));
}
function rpcAdminDescargarInformeIE(token, nombreIE, formato) {
  return ejecutarRpcSeguro_(() => adminDescargarInformeIE(token, nombreIE, formato));
}
function rpcAdminDescargarReporteB1BMas(token, formato) {
  return ejecutarRpcSeguro_(() => adminDescargarReporteB1BMas(token, formato));
}
function rpcDescargarInformeIE(nombreIE, token, formato) {
  return ejecutarRpcSeguro_(() => descargarInformeIE(nombreIE, token, formato));
}
function rpcAdminInstalarRedireccionHojaReal(token) {
  return ejecutarRpcSeguro_(() => adminInstalarRedireccionHojaReal(token));
}
function rpcAdminPrepararAvisoPublicacionWeb(token) {
  return ejecutarRpcSeguro_(() => adminPrepararAvisoPublicacionWeb(token));
}

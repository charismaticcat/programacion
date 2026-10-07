/**
 * Auth.gs — Portal IE SABER 11 Neiva 2026
 *
 * Dos credenciales distintas, sin relación con los permisos de Drive:
 *   - TOKEN POR IE: una palabra por institución, guardada en la hoja de
 *     control TOKENS_PORTAL_IE (dentro de la misma hoja de cálculo). Da
 *     acceso de lectura/escritura SOLO a los datos de esa IE.
 *   - TOKEN DE ADMINISTRADOR: una sola palabra secreta, guardada en las
 *     propiedades del script (nunca en una celda), solo para quien use
 *     la pantalla de Administrador.
 * El token de IE es la credencial real — cada función RPC de Datos.gs y
 * Graficos.gs lo vuelve a validar en cada llamada (no hay sesión de
 * servidor que recordar entre llamadas); el cliente solo lo guarda en
 * localStorage para no pedirlo en cada clic.
 */
const PROP_ADMIN_TOKEN_ = 'ADMIN_TOKEN';
const ALFABETO_TOKEN_ = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; // sin 0/O/1/I/L para evitar confusiones al dictar el token

function generarTokenAleatorio_(longitud) {
  longitud = longitud || 8;
  let s = '';
  for (let i = 0; i < longitud; i++) {
    s += ALFABETO_TOKEN_.charAt(Math.floor(Math.random() * ALFABETO_TOKEN_.length));
  }
  return s;
}

function cabecerasTokens_() { return ['IE', 'TOKEN']; }

/** Crea la hoja de tokens si falta y genera el token de las IE que todavía no tengan uno. Devuelve {IE: token}. */
function asegurarTokensIE_(ss) {
  let hoja = ss.getSheetByName(CFG.HOJA_TOKENS);
  if (!hoja) {
    hoja = ss.insertSheet(CFG.HOJA_TOKENS);
    hoja.getRange(1, 1, 1, 2).setValues([cabecerasTokens_()]);
    hoja.setFrozenRows(1);
    hoja.hideSheet();
  }
  const ultimaFila = hoja.getLastRow();
  const datos = ultimaFila >= 2 ? hoja.getRange(2, 1, ultimaFila - 1, 2).getValues() : [];
  const mapa = {};
  datos.forEach(fila => { if (texto_(fila[0])) mapa[norm_(fila[0])] = texto_(fila[1]); });
  const faltantes = [];
  CFG.IES.forEach(ie => {
    if (!mapa[norm_(ie)]) {
      const token = generarTokenAleatorio_();
      mapa[norm_(ie)] = token;
      faltantes.push([ie, token]);
    }
  });
  if (faltantes.length) {
    hoja.getRange(hoja.getLastRow() + 1, 1, faltantes.length, 2).setValues(faltantes);
  }
  return mapa;
}

/** {ok, mensaje} — nunca dice si la IE existe o no cuando el token falla, para no ayudar a adivinar. */
function validarAccesoIE(nombreIE, token) {
  if (!CFG.ACCESO_TOKENS_IE_HABILITADO) {
    return { ok: false, mensaje: 'El acceso de autoservicio está temporalmente pausado. Use el formulario de solicitud de acceso.' };
  }
  const ss = abrirSpreadsheet_();
  const mapa = asegurarTokensIE_(ss);
  const clave = norm_(nombreIE);
  if (CFG.IES.map(norm_).indexOf(clave) < 0) {
    return { ok: false, mensaje: 'Institución no reconocida.' };
  }
  if (!token || mapa[clave] !== texto_(token).toUpperCase()) {
    return { ok: false, mensaje: 'No puede acceder a esta institución: el token no corresponde. Verifique con el administrador.' };
  }
  return { ok: true, nombreIE: CFG.IES[CFG.IES.map(norm_).indexOf(clave)] };
}

/** Lanza un error (y por lo tanto corta la ejecución del RPC) si el token de IE no es válido. */
function exigirAccesoIE_(nombreIE, token) {
  const r = validarAccesoIE(nombreIE, token);
  if (!r.ok) throw new Error(r.mensaje);
  return r.nombreIE;
}

/**
 * Para funciones que dan acceso de lectura/escritura a UNA IE concreta
 * (Datos.gs, Graficos.gs). Spec del usuario: acceso automático — con
 * solo existir la IE, entra; el token de IE queda sin usarse (revisar
 * ACCESO_IE_AUTOMATICO en Config.gs). El token de administrador sigue
 * dando acceso igual, por si ese valor se vuelve a poner en falso.
 * Devuelve el nombre real de la IE (con la misma capitalización que
 * CFG.IES).
 */
function exigirAccesoIEoAdminComoIE_(nombreIE, token) {
  const clave = norm_(nombreIE);
  const idx = CFG.IES.map(norm_).indexOf(clave);
  if (idx < 0) throw new Error('Institución no reconocida.');
  if (CFG.ACCESO_IE_AUTOMATICO) return CFG.IES[idx];
  if (validarAccesoAdmin(token).ok) return CFG.IES[idx];
  return exigirAccesoIE_(nombreIE, token);
}

function obtenerTokenAdmin_() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty(PROP_ADMIN_TOKEN_);
  if (!token) {
    token = generarTokenAleatorio_(10);
    props.setProperty(PROP_ADMIN_TOKEN_, token);
  }
  return token;
}

/**
 * Para que el propietario obtenga/regenere el token de administrador:
 * ejecutar esta función una vez desde el editor de Apps Script (Ejecutar
 * → verTokenAdmin) y leer el resultado en "Ver registros" (Ctrl+Enter).
 * Nunca se expone por la app web.
 */
function verTokenAdmin() {
  Logger.log('TOKEN DE ADMINISTRADOR: ' + obtenerTokenAdmin_());
}
function regenerarTokenAdmin() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_ADMIN_TOKEN_);
  Logger.log('NUEVO TOKEN DE ADMINISTRADOR: ' + obtenerTokenAdmin_());
}

function validarAccesoAdmin(token) {
  if (!token || texto_(token).toUpperCase() !== obtenerTokenAdmin_()) {
    return { ok: false, mensaje: 'Token de administrador incorrecto.' };
  }
  return { ok: true };
}
function exigirAccesoAdmin_(token) {
  const r = validarAccesoAdmin(token);
  if (!r.ok) throw new Error(r.mensaje);
}

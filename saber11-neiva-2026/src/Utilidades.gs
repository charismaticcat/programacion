/**
 * Utilidades.gs — Portal IE SABER 11 Neiva 2026
 *
 * Funciones genéricas de texto/formato, y la búsqueda de logos por
 * similitud de nombre, portadas tal cual del script original atado a la
 * hoja de cálculo (gs_para_excel_CORREGIDO41) para que el portal muestre
 * el mismo logo por IE en la pantalla de distribución de links.
 */
function texto_(v) { return String(v == null ? '' : v).trim(); }

function norm_(v) {
  return texto_(v).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
}

function letra_(n) {
  let s = '';
  while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/**
 * Separador de argumentos de fórmula ("," o ";") según la configuración
 * regional del ARCHIVO (no del script) — mismo truco que el script atado
 * a la hoja: se prueba con una hoja temporal y se borra enseguida.
 * Necesario para cualquier fórmula que escriba este proyecto (p. ej. el
 * botón HYPERLINK de Graficos.gs), porque escribir ";" en un archivo
 * configurado en inglés (",") da #NAME? y viceversa.
 */
let SEPARADOR_FORMULA_ = null;
function separadorFormula_(ss) {
  if (SEPARADOR_FORMULA_) return SEPARADOR_FORMULA_;
  const tmp = ss.insertSheet('_TMP_' + Date.now());
  try {
    const c = tmp.getRange('A1');
    c.setFormula('=SUM(1,2)');
    SpreadsheetApp.flush();
    SEPARADOR_FORMULA_ = c.getDisplayValue() === '3' ? ',' : ';';
  } finally {
    ss.deleteSheet(tmp);
  }
  return SEPARADOR_FORMULA_;
}
/** Arma "=HYPERLINK(url<sep>"texto")" con el separador correcto del archivo. */
function formulaHyperlink_(ss, url, texto) {
  return '=HYPERLINK("' + url + '"' + separadorFormula_(ss) + '"' + texto + '")';
}

/** Nombre de IE sin el prefijo "I.E."/"IE" — para mostrar y para los letreros. */
function nombreSinPrefijoIE_(nombreHoja) {
  return texto_(nombreHoja).replace(/^\s*(I\.?\s*E\.?|IE)\s+/i, '').trim();
}

const VACIAS_LOGO_ = ['IE', 'I', 'E', 'INSTITUCION', 'EDUCATIVA', 'COLEGIO', 'DE', 'DEL', 'LA', 'LAS', 'LOS', 'Y', 'SEDE'];
function tokensLogo_(s) {
  return norm_(String(s).replace(/\.(png|jpe?g|gif|webp)$/i, ''))
    .replace(/[^A-Z0-9 ]/g, ' ').split(/\s+/).filter(t => t && VACIAS_LOGO_.indexOf(t) < 0);
}
function distanciaLogo_(a, b) {
  const d = [];
  for (let i = 0; i <= a.length; i++) d[i] = [i];
  for (let j = 0; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
  }
  return d[a.length][b.length];
}
function similitudLogo_(a, b) {
  const m = Math.max(a.length, b.length);
  return m ? 1 - distanciaLogo_(a, b) / m : 0;
}
function coberturaLogo_(a, b) {
  if (!a.length) return 0;
  return a.filter(t => b.some(u => u === t || (t.length >= 5 && distanciaLogo_(t, u) <= (t.length >= 8 ? 2 : 1)))).length / a.length;
}
function listarLogos_() {
  const out = [];
  const it = DriveApp.getFolderById(CFG.CARPETA_LOGOS).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (String(f.getMimeType()).indexOf('image/') === 0) out.push(f);
  }
  return out;
}
/** Busca el archivo de logo (Drive) cuyo nombre se parece más al de la IE. */
function logoDe_(nombreIE, logos) {
  const manual = CFG.LOGOS_MANUALES[nombreIE];
  if (manual) return DriveApp.getFileById(manual);
  const th = tokensLogo_(nombreIE);
  let mejor = null, puntaje = 0;
  logos.forEach(f => {
    const tf = tokensLogo_(f.getName());
    if (!tf.length) return;
    const p = Math.max(similitudLogo_(th.join(''), tf.join('')), (coberturaLogo_(th, tf) + coberturaLogo_(tf, th)) / 2);
    if (p > puntaje) { puntaje = p; mejor = f; }
  });
  return puntaje >= 0.8 ? mejor : null;
}

/**
 * Logo de una IE como base64, listo para un <img src="data:...">. Sirve
 * los bytes a través de un RPC (nunca un doGet binario: Apps Script no
 * puede devolver una imagen cruda desde doGet, solo HTML o texto) y los
 * guarda en caché por IE hasta 6 horas — sin esto, la pantalla de
 * distribución de links tendría que leer los 36 logos de Drive en cada
 * visita.
 */
function obtenerLogoBase64IE_(nombreIE, logos) {
  const cache = CacheService.getScriptCache();
  const clave = 'LOGO_' + norm_(nombreIE);
  const cacheado = cache.get(clave);
  if (cacheado) return JSON.parse(cacheado);
  const archivo = logoDe_(nombreIE, logos);
  let resultado = { base64: '', mimeType: '' };
  if (archivo) {
    let blob;
    try { blob = archivo.getThumbnail() || archivo.getBlob(); } catch (e) { blob = archivo.getBlob(); }
    resultado = { base64: Utilities.base64Encode(blob.getBytes()), mimeType: blob.getContentType() || 'image/png' };
  }
  try { cache.put(clave, JSON.stringify(resultado), 21600); } catch (e) {} // 100 KB máx. por clave; si no cabe, simplemente no cachea
  return resultado;
}

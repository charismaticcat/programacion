/**
 * Utils.gs — Foro Educativo Comunal Neiva 2026
 * Helpers genéricos sin lógica de negocio. Varios adaptados tal cual de
 * FEI 3.1 (ver docs/01-auditoria-fei-3.1.md §2.7) por ser utilidades
 * triviales de formato/normalización.
 */

/**
 * Normaliza texto para comparaciones tolerantes: quita tildes, pasa a
 * mayúsculas y colapsa espacios. Igual patrón que normalizarNombreIE_ /
 * normalizarNombreParaGrupoFEM_ de FEI 3.1.
 */
function normalizarTexto_(texto) {
  return String(texto == null ? "" : texto)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

/** Number() tolerante a coma decimal española (equivalente a numeroLocalizado_ de 3.1). */
function numeroLocalizado_(texto) {
  if (texto === null || texto === undefined || texto === "") return 0;
  if (typeof texto === "number") return texto;
  var limpio = String(texto).trim().replace(/\./g, "").replace(",", ".");
  var n = Number(limpio);
  return isNaN(n) ? 0 : n;
}

/** Conteo genérico de opciones seleccionadas (ej. selección múltiple con "Otro"). */
function tallyOpciones_(personas, campo) {
  var conteo = {};
  (personas || []).forEach(function (p) {
    var valor = p && p[campo];
    if (!valor) return;
    var opciones = Array.isArray(valor) ? valor : String(valor).split(",");
    opciones.forEach(function (op) {
      var clave = String(op).trim();
      if (!clave) return;
      conteo[clave] = (conteo[clave] || 0) + 1;
    });
  });
  return conteo;
}

/** Top-3 legible de un tally (objeto {opcion: conteo}). */
function top3Texto_(tally) {
  var entradas = Object.keys(tally || {}).map(function (k) {
    return { etiqueta: k, valor: tally[k] };
  });
  entradas.sort(function (a, b) {
    return b.valor - a.valor;
  });
  return entradas
    .slice(0, 3)
    .map(function (e) {
      return e.etiqueta + " (" + e.valor + ")";
    })
    .join(", ");
}

/** Heurística simple anti-texto-basura en respuestas abiertas (equivalente a textoTieneSentidoFEM_). */
function textoTieneSentido_(texto) {
  var limpio = String(texto || "").trim();
  if (limpio.length < 3) return false;
  if (!/[a-zA-ZáéíóúñÁÉÍÓÚÑ]/.test(limpio)) return false;
  // rechaza secuencias de un solo carácter repetido ("aaaaaa", "......")
  if (/^(.)\1{4,}$/.test(limpio.replace(/\s+/g, ""))) return false;
  return true;
}

/** Formatea una fecha en español largo, p.ej. "24 de septiembre de 2026". */
function formatearFechaLargaEs_(fecha, capitalizarMes) {
  if (!fecha) return "";
  var d = fecha instanceof Date ? fecha : new Date(fecha);
  if (isNaN(d.getTime())) return String(fecha);
  var meses = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
  ];
  var mes = meses[d.getMonth()];
  if (capitalizarMes) mes = mes.charAt(0).toUpperCase() + mes.slice(1);
  return d.getDate() + " de " + mes + " de " + d.getFullYear();
}

/** Formatea fecha+hora corta para registros de firma/asistencia. */
function formatearFechaHora_(fecha) {
  var d = fecha instanceof Date ? fecha : new Date(fecha || Date.now());
  return Utilities.formatDate(d, "America/Bogota", "dd/MM/yyyy HH:mm");
}

/** Quita el prefijo institucional ("I.E. ", "IE ", "I.E ") del nombre de una IE. */
function nombreIESinPrefijo_(ie) {
  return String(ie || "").replace(/^\s*I\.?E\.?\s*/i, "").trim();
}

/** Capitaliza cada palabra de un nombre (para presentación en informes/correos). */
function capitalizarNombre_(nombre) {
  return String(nombre || "")
    .toLowerCase()
    .replace(/(^|\s|\/)([a-záéíóúñ])/g, function (m, sep, letra) {
      return sep + letra.toUpperCase();
    });
}

/** Cuenta palabras de un texto libre (mismo criterio que el contador en vivo del cliente: split por espacios). */
function contarPalabras_(texto) {
  var limpio = String(texto || "").trim();
  return limpio ? limpio.split(/\s+/).length : 0;
}

/** Genera un identificador legible corto para filas (ConectaEduca, etc.), no un UUID completo. */
function generarIdCorto_() {
  return Utilities.getUuid().split("-")[0];
}

/** Escapa HTML básico para render en cliente (equivalente a escapeHtml de App.html). */
function escapeHtml_(texto) {
  return String(texto == null ? "" : texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Envuelve toda función RPC (Code.gs) para que una excepción no prevista
 * (p. ej. un límite de cuota, un timeout, un bug real del servidor) se
 * convierta en una respuesta {ok:false, mensaje:...} legible en el
 * cliente, en vez de propagarse como el genérico "Ocurrió un error de
 * comunicación con el servidor" que dispara el withFailureHandler de
 * llamarServidor() (JS.html). El error real queda en los logs de Apps
 * Script (Ejecuciones) para depuración.
 */
function ejecutarRpcSeguro_(fn) {
  try {
    return fn();
  } catch (error) {
    Logger.log("Error no controlado en RPC: " + (error && error.stack ? error.stack : error));
    return {
      ok: false,
      mensaje: "Ocurrió un error inesperado en el servidor. Intente de nuevo en unos segundos."
    };
  }
}

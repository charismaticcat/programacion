/**
 * Data.gs — Foro Educativo Comunal Neiva 2026
 *
 * Utilidades genéricas de acceso a Google Sheets: apertura con reintento,
 * mapa de cabeceras, búsqueda de fila por columna clave, UPSERT genérico
 * (con fusión opcional de campos concurrentes) y lectura de filas como
 * objetos. Generalizan los patrones de FEI 3.1 (abrirSpreadsheet_,
 * buscarFilaPorIdForo_, guardarAvanceForo — ver docs/01-auditoria-fei-3.1.md
 * §1.3 y §5.7) para que cualquier hoja nueva (Sesion1Comunal, ConectaEduca,
 * AccesosGrupo, ...) los reutilice sin duplicar código, a diferencia de 3.1
 * donde cada hoja tenía su propio UPSERT copiado a mano.
 */

/** Abre el spreadsheet del proyecto con reintento (igual patrón que abrirSpreadsheet_ de 3.1). */
function abrirSpreadsheet_() {
  var ultimoError = null;
  for (var i = 0; i < 4; i++) {
    try {
      return SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
    } catch (error) {
      ultimoError = error;
      Utilities.sleep(500 * (i + 1));
    }
  }
  throw new Error(
    "El servicio de Hojas de cálculo no respondió después de varios intentos: " +
      (ultimoError ? ultimoError.message : "error desconocido")
  );
}

/**
 * Obtiene una hoja por nombre; si no existe, la crea con las cabeceras
 * dadas en la fila 1. Si existe pero le faltan columnas de la lista, las
 * agrega al final (nunca reordena ni borra columnas existentes).
 */
function obtenerHoja_(nombreHoja, cabeceras) {
  var ss = abrirSpreadsheet_();
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
    hoja.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]);
    hoja.setFrozenRows(1);
    return hoja;
  }
  var mapa = obtenerMapaCabeceras_(hoja);
  var faltantes = cabeceras.filter(function (c) {
    return !mapa[c];
  });
  if (faltantes.length) {
    var ultimaColumna = hoja.getLastColumn();
    hoja.getRange(1, ultimaColumna + 1, 1, faltantes.length).setValues([faltantes]);
  }
  return hoja;
}

/** Mapa {NOMBRE_CABECERA: numeroDeColumna(1-based)} leído de la fila 1. */
function obtenerMapaCabeceras_(hoja) {
  var ultimaColumna = hoja.getLastColumn();
  if (ultimaColumna === 0) return {};
  var cabeceras = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0];
  var mapa = {};
  cabeceras.forEach(function (nombre, indice) {
    var clave = String(nombre || "").trim();
    if (clave !== "") mapa[clave] = indice + 1;
  });
  return mapa;
}

/**
 * Busca el número de fila (1-based) cuya columna `columnaClave` sea igual
 * a `valorClave` (comparación como texto, trim). Devuelve -1 si no existe.
 * Equivalente genérico de buscarFilaPorIdForo_ de 3.1.
 */
function buscarFilaPorColumna_(hoja, mapaCabeceras, columnaClave, valorClave) {
  var columna = mapaCabeceras[columnaClave];
  if (!columna) {
    throw new Error('La hoja "' + hoja.getName() + '" no tiene la columna ' + columnaClave + ".");
  }
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return -1;
  var valores = hoja.getRange(2, columna, ultimaFila - 1, 1).getValues();
  var objetivo = String(valorClave).trim();
  for (var i = 0; i < valores.length; i++) {
    if (String(valores[i][0] || "").trim() === objetivo) return i + 2;
  }
  return -1;
}

/** Todas las filas que cumplan columna=valor, como número de fila (1-based). */
function buscarFilasPorColumna_(hoja, mapaCabeceras, columnaClave, valorClave) {
  var columna = mapaCabeceras[columnaClave];
  if (!columna) return [];
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  var valores = hoja.getRange(2, columna, ultimaFila - 1, 1).getValues();
  var objetivo = String(valorClave).trim();
  var filas = [];
  for (var i = 0; i < valores.length; i++) {
    if (String(valores[i][0] || "").trim() === objetivo) filas.push(i + 2);
  }
  return filas;
}

/** Lee una fila completa como objeto {CABECERA: valor}. */
function leerFilaComoObjeto_(hoja, fila, mapaCabeceras) {
  var ultimaColumna = hoja.getLastColumn();
  var valores = hoja.getRange(fila, 1, 1, ultimaColumna).getValues()[0];
  var obj = {};
  Object.keys(mapaCabeceras).forEach(function (cabecera) {
    obj[cabecera] = valores[mapaCabeceras[cabecera] - 1];
  });
  return obj;
}

/** Lee todas las filas de datos (desde la fila 2) como arreglo de objetos {CABECERA: valor}. */
function leerFilasComoObjetos_(hoja) {
  var mapaCabeceras = obtenerMapaCabeceras_(hoja);
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  var ultimaColumna = hoja.getLastColumn();
  var valores = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
  return valores.map(function (fila) {
    var obj = {};
    Object.keys(mapaCabeceras).forEach(function (cabecera) {
      obj[cabecera] = fila[mapaCabeceras[cabecera] - 1];
    });
    return obj;
  });
}

/**
 * UPSERT genérico por columna clave. `valores` es un objeto parcial
 * {CABECERA: valor} — solo esas columnas se escriben.
 *
 * opciones.fusionar = true reproduce el patrón anti-"última escritura
 * gana" de guardarAvanceForo en 3.1 (docs/01-auditoria-fei-3.1.md §5.7):
 * si la fila ya existe, un valor entrante SOLO sobreescribe al ya
 * guardado cuando trae contenido real (string no vacío tras trim, o
 * cualquier valor no vacío/no-cero para números/booleanos) — así varios
 * dispositivos conectados al mismo grupo pueden guardar en paralelo sin
 * que el último en llegar borre en silencio lo que otro ya guardó.
 *
 * Devuelve el número de fila escrita (nueva o existente).
 */
function upsertFila_(nombreHoja, cabeceras, columnaClave, valorClave, valores, opciones) {
  opciones = opciones || {};
  var hoja = obtenerHoja_(nombreHoja, cabeceras);
  var mapaCabeceras = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapaCabeceras, columnaClave, valorClave);

  if (fila === -1) {
    var nuevaFila = cabeceras.map(function (c) {
      if (c === columnaClave) return valorClave;
      return c in valores ? valores[c] : "";
    });
    hoja.appendRow(nuevaFila);
    return hoja.getLastRow();
  }

  var columnasAEscribir = Object.keys(valores).filter(function (c) {
    return mapaCabeceras[c];
  });
  if (!columnasAEscribir.length) return fila;

  if (opciones.fusionar) {
    columnasAEscribir = columnasAEscribir.filter(function (c) {
      var entrante = valores[c];
      var tieneContenido =
        entrante !== null && entrante !== undefined && String(entrante).trim() !== "";
      if (tieneContenido) return true;
      // si no trae contenido, solo escribir si la celda actual también está vacía
      var actual = hoja.getRange(fila, mapaCabeceras[c]).getValue();
      return String(actual || "").trim() === "";
    });
  }

  columnasAEscribir.forEach(function (c) {
    hoja.getRange(fila, mapaCabeceras[c]).setValue(valores[c]);
  });
  return fila;
}

/**
 * Participación y asistencia — integradas dentro de PARTICIPACIÓN (spec
 * sección 7), no como módulo aparte. Adaptado de actualizarParticipacion_ /
 * registrarAsistenciaQR de FEI 3.1 (docs/01-auditoria-fei-3.1.md §2.2).
 */
var HOJA_PARTICIPACION_COMUNAL_ = "ParticipacionComunal";

function cabecerasParticipacionComunal_() {
  return [
    "ID_PARTICIPANTE", "ID_GRUPO", "ID_IE", "NOMBRE", "ROL", "CORREO",
    "CONFIRMACION_ASISTENCIA", "FECHA", "ESTADO", "DISPOSITIVO_ID"
  ];
}

/**
 * Registra la asistencia de un participante. Deduplicación por
 * (ID_GRUPO, ID_IE, NOMBRE normalizado) — mismo criterio que la
 * deduplicación por documento de AsistenciaQR en 3.1 (auditoría §5.6): si
 * ya existe, devuelve el registro existente en vez de crear uno nuevo.
 */
function registrarParticipante(idGrupo, idIE, nombre, rol, correo, dispositivoId) {
  return conLock_(function () {
    idGrupo = String(idGrupo || "").trim();
    nombre = String(nombre || "").trim();
    if (!idGrupo) return { ok: false, mensaje: "Falta el grupo." };
    if (!nombre) return { ok: false, mensaje: "El nombre es obligatorio." };

    var hoja = obtenerHoja_(HOJA_PARTICIPACION_COMUNAL_, cabecerasParticipacionComunal_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var ultimaFila = hoja.getLastRow();
    var nombreNormalizado = normalizarTexto_(nombre);

    if (ultimaFila >= 2) {
      var valores = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
      for (var i = 0; i < valores.length; i++) {
        var fila = valores[i];
        var mismoGrupo = String(fila[mapa["ID_GRUPO"] - 1] || "").trim() === idGrupo;
        var mismoNombre = normalizarTexto_(fila[mapa["NOMBRE"] - 1]) === nombreNormalizado;
        if (mismoGrupo && mismoNombre) {
          return { ok: true, yaRegistrado: true, participante: leerFilaComoObjeto_(hoja, i + 2, mapa) };
        }
      }
    }

    var idParticipante = Utilities.getUuid();
    hoja.appendRow([
      idParticipante, idGrupo, String(idIE || "").trim(), nombre, String(rol || "").trim(),
      String(correo || "").trim(), "SI", new Date(), "REGISTRADO", String(dispositivoId || "").trim()
    ]);
    return { ok: true, yaRegistrado: false, idParticipante: idParticipante };
  }, 10000);
}

/** Cuenta los firmantes de un grupo (para el panel permanente, spec sección 8). */
function contarParticipantesGrupo(idGrupo) {
  return listarFirmantesGrupo(idGrupo).length;
}

/** Lista completa de firmantes de un grupo: nombre, institución, rol (más recientes primero). */
function listarFirmantesGrupo(idGrupo) {
  var hoja = obtenerHoja_(HOJA_PARTICIPACION_COMUNAL_, cabecerasParticipacionComunal_());
  var filas = leerFilasComoObjetos_(hoja);
  var objetivo = String(idGrupo || "").trim();
  var deIE = {};
  obtenerInstitucionesDelGrupo(objetivo).forEach(function (ie) {
    deIE[ie.idIE] = ie.institucion;
  });
  return filas
    .filter(function (f) {
      return String(f.ID_GRUPO || "").trim() === objetivo;
    })
    .map(function (f) {
      return {
        nombre: String(f.NOMBRE || "").trim(),
        institucion: deIE[String(f.ID_IE || "").trim()] || "",
        rol: String(f.ROL || "").trim(),
        fecha: f.FECHA
      };
    })
    .reverse();
}

/** Ejecuta `fn` bajo LockService.getScriptLock(), liberando siempre el lock. */
function conLock_(fn, timeoutMs) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(timeoutMs || 10000);
  } catch (error) {
    return { ok: false, codigo: "SISTEMA_OCUPADO", mensaje: "El sistema está ocupado, intenta de nuevo en unos segundos." };
  }
  try {
    return fn();
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // no ocultar el error real de negocio si liberar el lock falla
    }
  }
}

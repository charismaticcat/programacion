/**
 * Access.gs — Foro Educativo Comunal Neiva 2026
 *
 * Acceso por TOKEN + código a nivel de GRUPO (no de IE). Adaptado de
 * validarAccesoIE / generarAccesosIE de FEI 3.1
 * (docs/01-auditoria-fei-3.1.md §1.2, §2.1, §5.1, §5.2). A diferencia de
 * 3.1, hay UN SOLO generador de código de acceso (3.1 tenía dos
 * incompatibles conviviendo — ver auditoría §7 — aquí eso no se
 * reproduce).
 */

var HOJA_ACCESOS_GRUPO_ = "AccesosGrupo";
var COLUMNAS_CODIGO_CONTINGENCIA_ = ["CODIGO_CONTINGENCIA_1", "CODIGO_CONTINGENCIA_2", "CODIGO_CONTINGENCIA_3"];

/**
 * Valida token+código+dispositivo y reclama un cupo de sesión para el
 * grupo. Devuelve el registro del grupo si es válido. El cliente nunca
 * decide ID_GRUPO/ID_FORO_COMUNAL — se resuelven aquí, en el servidor.
 */
function validarAccesoGrupo(token, codigo, dispositivoId, forzar) {
  token = String(token || "").trim();
  codigo = String(codigo || "").trim();
  dispositivoId = String(dispositivoId || "").trim();

  if (token === "") return { ok: false, codigo: "TOKEN_REQUERIDO", mensaje: "No se recibió el token de acceso." };
  if (codigo === "") return { ok: false, codigo: "CODIGO_REQUERIDO", mensaje: "Debe ingresar el código de acceso." };
  if (dispositivoId === "") return { ok: false, codigo: "DISPOSITIVO_REQUERIDO", mensaje: "No fue posible identificar este dispositivo." };

  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var numeroFila = buscarFilaPorColumna_(hoja, mapa, "TOKEN", token);
  if (numeroFila === -1) {
    return { ok: false, codigo: "TOKEN_INVALIDO", mensaje: "El enlace de acceso no es válido." };
  }

  var registro = hoja.getRange(numeroFila, 1, 1, hoja.getLastColumn()).getDisplayValues()[0];
  var valorColumna = function (nombre) {
    return mapa[nombre] ? String(registro[mapa[nombre] - 1] || "").trim() : "";
  };

  var idGrupo = valorColumna("ID_GRUPO");
  var grupo = valorColumna("GRUPO");
  var estado = valorColumna("ESTADO").toUpperCase();

  var codigosValidos = ["CODIGO_ACCESO"].concat(COLUMNAS_CODIGO_CONTINGENCIA_)
    .map(valorColumna)
    .filter(function (v) {
      return v !== "";
    });

  if (codigosValidos.indexOf(codigo) === -1) {
    return {
      ok: false,
      codigo: "CODIGO_INCORRECTO",
      mensaje: "El código es incorrecto. Verifique el código enviado a su grupo."
    };
  }

  if (idGrupo === "") {
    return { ok: false, codigo: "GRUPO_FALTANTE", mensaje: "El acceso no tiene un grupo asignado." };
  }

  if (estado === "BLOQUEADO" || estado === "INACTIVO") {
    return { ok: false, codigo: "ACCESO_BLOQUEADO", mensaje: "Este acceso ya no está disponible." };
  }

  if (mapa["HABILITAR_DESDE"]) {
    var valorHabilitar = hoja.getRange(numeroFila, mapa["HABILITAR_DESDE"]).getValue();
    if (valorHabilitar instanceof Date && !isNaN(valorHabilitar.getTime()) && new Date() < valorHabilitar) {
      var zona = Session.getScriptTimeZone();
      var hora = Utilities.formatDate(valorHabilitar, zona, "h:mm a").replace("AM", "a. m.").replace("PM", "p. m.");
      return {
        ok: false,
        codigo: "BLOQUEADO_POR_HORARIO",
        mensaje: "El Foro Educativo Comunal se habilitará a las " + hora + ". Por favor ingrese nuevamente a partir de esa hora."
      };
    }
  }

  var sesion = reclamarSesionGrupo_(idGrupo, dispositivoId, !!forzar);
  if (!sesion.ok) return sesion;

  hoja.getRange(numeroFila, mapa["ULTIMA_ACTIVIDAD"]).setValue(new Date());

  return {
    ok: true,
    idGrupo: idGrupo,
    grupo: grupo,
    idForoComunal: valorColumna("ID_FORO_COMUNAL"),
    tokenSesion: sesion.tokenSesion,
    esPrincipal: sesion.esPrincipal,
    instituciones: obtenerInstitucionesDelGrupo(idGrupo)
  };
}

/** Único generador de código de acceso (evita caracteres ambiguos: sin I,O,0,1). */
function generarCodigoAcceso_(codigosUsados) {
  var caracteres = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";
  var codigo;
  do {
    codigo = "FEC-";
    for (var i = 0; i < 5; i++) {
      codigo += caracteres.charAt(Math.floor(Math.random() * caracteres.length));
    }
  } while (codigosUsados[codigo]);
  codigosUsados[codigo] = true;
  return codigo;
}

function generarTokenUnico_(tokensUsados) {
  var token;
  do {
    token = Utilities.getUuid().replace(/-/g, "");
  } while (tokensUsados[token]);
  tokensUsados[token] = true;
  return token;
}

function cabecerasAccesosGrupo_() {
  return [
    "ID_ACCESO", "ID_GRUPO", "GRUPO", "ID_FORO_COMUNAL", "TOKEN", "CODIGO_ACCESO",
    "CODIGO_CONTINGENCIA_1", "CODIGO_CONTINGENCIA_2", "CODIGO_CONTINGENCIA_3",
    "URL_ACCESO", "ESTADO", "HABILITAR_DESDE", "EMAIL_RESPONSABLE_GRUPO",
    "FECHA_GENERACION", "ULTIMA_ACTIVIDAD", "SESION1_ENVIADA", "SESION2_ENVIADA",
    "FECHA_ENVIO_S1", "FECHA_ENVIO_S2", "FECHA_ENVIO_DEFINITIVO"
  ];
}

/**
 * Genera (una sola vez) TOKEN + CODIGO_ACCESO + 3 códigos de contingencia
 * para cada grupo presente en GruposComunal que todavía no tenga fila en
 * AccesosGrupo. Idempotente: nunca regenera credenciales de un grupo que
 * ya las tiene (mismo criterio que generarAccesosIE en 3.1 — el servidor
 * jamás regenera un identificador ya emitido).
 */
function generarAccesosGrupo() {
  return conLock_(function () {
    var grupos = obtenerGrupos();
    if (!grupos.length) {
      return { ok: false, mensaje: "GruposComunal no tiene grupos activos todavía. Complete esa hoja primero." };
    }

    var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var existentes = leerFilasComoObjetos_(hoja);
    var idsConAcceso = {};
    existentes.forEach(function (fila) {
      idsConAcceso[String(fila.ID_GRUPO || "").trim()] = true;
    });

    var codigosUsados = {};
    var tokensUsados = {};
    existentes.forEach(function (fila) {
      if (fila.CODIGO_ACCESO) codigosUsados[fila.CODIGO_ACCESO] = true;
      if (fila.TOKEN) tokensUsados[fila.TOKEN] = true;
    });

    var idForoComunal = getConfig().ID_FORO_COMUNAL;
    var creados = [];
    grupos.forEach(function (g) {
      if (idsConAcceso[g.idGrupo]) return;
      var token = generarTokenUnico_(tokensUsados);
      var fila = [
        Utilities.getUuid(),
        g.idGrupo,
        g.grupo,
        idForoComunal,
        token,
        generarCodigoAcceso_(codigosUsados),
        generarCodigoAcceso_(codigosUsados),
        generarCodigoAcceso_(codigosUsados),
        generarCodigoAcceso_(codigosUsados),
        construirUrlAcceso_(token),
        "DISPONIBLE",
        "", "", new Date(), "", "NO", "NO", "", "", ""
      ];
      hoja.appendRow(fila);
      creados.push(g.idGrupo);
    });

    return { ok: true, gruposCreados: creados, totalGrupos: grupos.length };
  }, 30000);
}

/** URL de acceso del grupo. Usa URL_WEBAPP de ConfiguracionComunal si está definida; si no, la del deployment activo. */
function construirUrlAcceso_(token) {
  var config = getConfig();
  var base = config.URL_WEBAPP || ScriptApp.getService().getUrl();
  return base + "?t=" + token;
}

/** Recalcula URL_ACCESO de todos los grupos (usar tras publicar un nuevo deployment). */
function actualizarUrlsAcceso() {
  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { ok: true, actualizados: 0 };
  var datos = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
  datos.forEach(function (fila, i) {
    var token = fila[mapa["TOKEN"] - 1];
    if (!token) return;
    hoja.getRange(i + 2, mapa["URL_ACCESO"]).setValue(construirUrlAcceso_(token));
  });
  return { ok: true, actualizados: datos.length };
}

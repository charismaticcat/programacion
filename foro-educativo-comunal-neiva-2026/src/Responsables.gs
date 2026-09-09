/**
 * Responsables.gs — Foro Educativo Comunal Neiva 2026
 *
 * Responsable de envío (principal) y hasta 3 asistentes de envío por
 * grupo — las personas cuyo correo recibirá el informe final del grupo
 * (Correo.gs). Se capturan DESDE la propia aplicación (a diferencia del
 * correo de acceso inicial en AccesosGrupo.EMAIL_RESPONSABLE_GRUPO, que
 * se introduce manualmente antes de enviar los accesos — no hay forma de
 * evitar ese primer correo "de arranque").
 *
 * ROLES_FORO_ es el catálogo oficial de roles operativos del Foro, tomado
 * textualmente del "Documento Orientador FEM2026" (págs. 9-10, roles
 * definidos para el desarrollo del FEI y reutilizados en el Encuentro
 * Comunal) — no se inventan roles nuevos.
 */

var HOJA_RESPONSABLES_COMUNAL_ = "ResponsablesComunal";
var MAX_ASISTENTES_ENVIO_ = 3;

var ROLES_FORO_ = [
  "Líder (Rector/Rectora)",
  "Dinamizador Pedagógico",
  "Dinamizador de Mesas de Trabajo",
  "Relator(a)",
  "Dinamizador del Tiempo",
  "Dinamizador de la Sistematización",
  "Participante",
  "Otro"
];

function cabecerasResponsablesComunal_() {
  return ["ID_REGISTRO", "ID_GRUPO", "TIPO", "NOMBRE", "ID_IE", "ROL_FORO", "CORREO", "FECHA"];
}

/**
 * Guarda el responsable de envío principal (TIPO="PRINCIPAL", UPSERT — solo
 * puede haber uno, se reemplaza si ya existía) o agrega un asistente de
 * envío (TIPO="ASISTENTE", hasta MAX_ASISTENTES_ENVIO_).
 */
function guardarResponsableEnvio(idGrupo, tokenSesion, dispositivoId, tipo, datos) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  tipo = String(tipo || "").toUpperCase();
  if (tipo !== "PRINCIPAL" && tipo !== "ASISTENTE") {
    return { ok: false, mensaje: "Tipo de responsable no reconocido." };
  }
  datos = datos || {};
  var nombre = String(datos.nombre || "").trim();
  var correo = String(datos.correo || "").trim();
  if (!nombre) return { ok: false, mensaje: "El nombre es obligatorio." };
  if (!correo) return { ok: false, mensaje: "El correo es obligatorio." };

  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_RESPONSABLES_COMUNAL_, cabecerasResponsablesComunal_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var existentes = leerFilasComoObjetos_(hoja).filter(function (f) {
      return String(f.ID_GRUPO || "").trim() === idGrupo;
    });

    if (tipo === "PRINCIPAL") {
      var filaPrincipal = existentes.find(function (f) {
        return String(f.TIPO).toUpperCase() === "PRINCIPAL";
      });
      var valores = {
        NOMBRE: nombre, ID_IE: String(datos.idIE || "").trim(),
        ROL_FORO: String(datos.rolForo || "").trim(), CORREO: correo, FECHA: new Date()
      };
      if (filaPrincipal) {
        var filaNum = buscarFilaPorColumna_(hoja, mapa, "ID_REGISTRO", filaPrincipal.ID_REGISTRO);
        Object.keys(valores).forEach(function (c) {
          hoja.getRange(filaNum, mapa[c]).setValue(valores[c]);
        });
        return { ok: true, idRegistro: filaPrincipal.ID_REGISTRO };
      }
      var idRegistro = Utilities.getUuid();
      hoja.appendRow([idRegistro, idGrupo, "PRINCIPAL", nombre, valores.ID_IE, valores.ROL_FORO, correo, new Date()]);
      return { ok: true, idRegistro: idRegistro };
    }

    // ASISTENTE
    var asistentes = existentes.filter(function (f) {
      return String(f.TIPO).toUpperCase() === "ASISTENTE";
    });
    if (asistentes.length >= MAX_ASISTENTES_ENVIO_) {
      return { ok: false, mensaje: "Ya se registraron los " + MAX_ASISTENTES_ENVIO_ + " asistentes de envío permitidos." };
    }
    var idRegistroAsistente = Utilities.getUuid();
    hoja.appendRow([
      idRegistroAsistente, idGrupo, "ASISTENTE", nombre, String(datos.idIE || "").trim(),
      String(datos.rolForo || "").trim(), correo, new Date()
    ]);
    return { ok: true, idRegistro: idRegistroAsistente };
  }, 15000);
}

/** Responsable principal + asistentes de envío ya registrados para el grupo. */
function listarResponsablesEnvio(idGrupo) {
  var hoja = obtenerHoja_(HOJA_RESPONSABLES_COMUNAL_, cabecerasResponsablesComunal_());
  var filas = leerFilasComoObjetos_(hoja);
  var objetivo = String(idGrupo || "").trim();
  var deIE = {};
  obtenerInstitucionesDelGrupo(objetivo).forEach(function (ie) {
    deIE[ie.idIE] = ie.institucion;
  });
  var deGrupo = filas.filter(function (f) {
    return String(f.ID_GRUPO || "").trim() === objetivo;
  });
  var mapear = function (f) {
    return {
      idRegistro: f.ID_REGISTRO,
      tipo: String(f.TIPO || "").toUpperCase(),
      nombre: f.NOMBRE,
      institucion: deIE[String(f.ID_IE || "").trim()] || "",
      rolForo: f.ROL_FORO,
      correo: f.CORREO
    };
  };
  return {
    principal: (function () {
      var f = deGrupo.find(function (x) { return String(x.TIPO).toUpperCase() === "PRINCIPAL"; });
      return f ? mapear(f) : null;
    })(),
    asistentes: deGrupo
      .filter(function (f) { return String(f.TIPO).toUpperCase() === "ASISTENTE"; })
      .map(mapear)
  };
}

/** Elimina un asistente de envío (el principal no se elimina, se reemplaza guardando otro). */
function eliminarResponsableEnvio(idGrupo, idRegistro, tokenSesion, dispositivoId) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_RESPONSABLES_COMUNAL_, cabecerasResponsablesComunal_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_REGISTRO", String(idRegistro || "").trim());
    if (fila === -1) return { ok: false, mensaje: "El registro ya no existe." };
    if (String(hoja.getRange(fila, mapa["ID_GRUPO"]).getValue()).trim() !== idGrupo) {
      return { ok: false, mensaje: "Ese registro no pertenece a este grupo." };
    }
    if (String(hoja.getRange(fila, mapa["TIPO"]).getValue()).toUpperCase() === "PRINCIPAL") {
      return { ok: false, mensaje: "El responsable principal no se elimina; registra otro para reemplazarlo." };
    }
    hoja.deleteRow(fila);
    return { ok: true };
  }, 10000);
}

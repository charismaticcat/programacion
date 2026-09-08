/**
 * Sesion1.gs — Foro Educativo Comunal Neiva 2026
 *
 * Sesión 1: socialización de resultados del FEI + construcción colectiva
 * (spec secciones 9-10). UN SOLO registro por GRUPO (nunca por IE — spec:
 * "Estas respuestas pertenecen al GRUPO. NO crear un formulario
 * independiente por IE"). UPSERT con fusión de campos concurrentes,
 * reutilizando el patrón anti-"última escritura gana" de guardarAvanceForo
 * en 3.1 (docs/01-auditoria-fei-3.1.md §5.7), generalizado en
 * upsertFila_(..., {fusionar:true}) de Data.gs.
 */

var HOJA_SESION1_COMUNAL_ = "Sesion1Comunal";

function cabecerasSesion1Comunal_() {
  return [
    "ID_GRUPO", "REFLEXIONES", "CONCLUSIONES", "PROPUESTAS_IE", "EXPERIENCIAS", "RETOS",
    "APORTES_TERRITORIALES", "CONVERGENCIAS", "APUESTAS", "DESAFIOS", "IDENTIDAD",
    "PRIORIDADES", "PROPUESTAS_COLECTIVAS", "ACUERDOS", "RUTA", "ULTIMA_ACTUALIZACION"
  ];
}

var CAMPOS_SESION1_ = cabecerasSesion1Comunal_().filter(function (c) {
  return c !== "ID_GRUPO" && c !== "ULTIMA_ACTUALIZACION";
});

/**
 * Guarda (UPSERT con fusión) los campos de Sesión 1 recibidos del cliente.
 * `campos` es un objeto parcial {NOMBRE_CAMPO: texto} — el cliente puede
 * enviar solo los campos que su dispositivo tiene abiertos/editados.
 */
function guardarSesion1(idGrupo, tokenSesion, dispositivoId, campos) {
  idGrupo = String(idGrupo || "").trim();
  if (!idGrupo) return { ok: false, mensaje: "Falta el grupo." };
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }

  return conLock_(function () {
    var valores = {};
    CAMPOS_SESION1_.forEach(function (campo) {
      if (campos && Object.prototype.hasOwnProperty.call(campos, campo)) {
        valores[campo] = String(campos[campo] == null ? "" : campos[campo]);
      }
    });
    valores["ULTIMA_ACTUALIZACION"] = new Date();

    upsertFila_(HOJA_SESION1_COMUNAL_, cabecerasSesion1Comunal_(), "ID_GRUPO", idGrupo, valores, { fusionar: true });
    return { ok: true, guardadoEn: new Date().toISOString() };
  }, 15000);
}

/** Lee el estado actual de Sesión 1 del grupo (para restaurar/mostrar a cualquier dispositivo que entre). */
function obtenerSesion1(idGrupo) {
  var hoja = obtenerHoja_(HOJA_SESION1_COMUNAL_, cabecerasSesion1Comunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", String(idGrupo || "").trim());
  if (fila === -1) return null;
  return leerFilaComoObjeto_(hoja, fila, mapa);
}

/** Marca Sesión 1 como enviada definitivamente (solo el responsable principal, spec sección 22). */
function enviarSesion1Definitiva(idGrupo, tokenSesion, dispositivoId) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  if (!esPrincipalDeGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, mensaje: "Solo el responsable principal del grupo puede enviar Sesión 1 de forma definitiva." };
  }

  var datos = obtenerSesion1(idGrupo);
  var vacios = CAMPOS_SESION1_.filter(function (c) {
    return !datos || !String(datos[c] || "").trim();
  });
  if (vacios.length) {
    return { ok: false, mensaje: "Faltan campos por completar en Sesión 1: " + vacios.join(", ") };
  }

  return conLock_(function () {
    var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
    var mapa = obtenerMapaCabeceras_(hoja);
    var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", idGrupo);
    if (fila === -1) return { ok: false, mensaje: "No existe acceso para este grupo." };
    hoja.getRange(fila, mapa["SESION1_ENVIADA"]).setValue("SI");
    hoja.getRange(fila, mapa["FECHA_ENVIO_S1"]).setValue(new Date());
    hoja.getRange(fila, mapa["ESTADO"]).setValue("SESION1_ENVIADA");
    return { ok: true };
  }, 15000);
}

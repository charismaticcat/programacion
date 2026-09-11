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
    "ID_GRUPO",
    // Consolidado de Socialización (Sesión 1, tarjeta 1) — 4 preguntas, min 50/máx 400 palabras.
    "REFLEXIONES", "DESAFIOS", "APUESTAS", "CONCLUSIONES",
    // Construcción colectiva del grupo (Sesión 1, tarjeta 2).
    "PRIORIDADES", "PROPUESTAS_COLECTIVAS", "ACUERDOS", "RUTA",
    // Preguntas de grupo antes de ConectaEduca (Sesión 2) — se guardan
    // aquí por simplicidad, igual que el resto: no son por actor, son
    // UN solo par de respuestas por grupo.
    "NECESIDADES_ARTICULACION_GRUPO", "OPORTUNIDADES_GRUPO",
    // Segunda parte de ConectaEduca (Sesión 2) — min 50/máx 400 palabras.
    "PRIORIDADES_CE", "ACUERDOS_CE", "PROPUESTAS_CE", "RUTA_CE",
    // Espacio libre y opcional por sesión para hallazgos propios de la
    // comunidad que no encajan en las preguntas orientadoras — mismo
    // espíritu que la Sesión Propia/4 (opcional) de FEI 3.1
    // (docs/01-auditoria-fei-3.1.md §2.6), aquí uno por Sesión 1 y otro
    // por Sesión 2/ConectaEduca (se guardan aquí por simplicidad: ambas
    // sesiones ya usan el mismo UPSERT-por-grupo con fusión de campos).
    "APORTE_PROPIO_S1_TITULO", "APORTE_PROPIO_S1_TEXTO",
    "APORTE_PROPIO_S2_TITULO", "APORTE_PROPIO_S2_TEXTO",
    // Columnas heredadas de versiones anteriores de la app, ya no se
    // muestran en pantalla (spec: "elimina todo y solo deja..."), pero se
    // conservan aquí para no perder ni desalinear datos ya capturados por
    // grupos que las hayan diligenciado antes de este cambio.
    "PROPUESTAS_IE", "EXPERIENCIAS", "RETOS", "APORTES_TERRITORIALES", "CONVERGENCIAS", "IDENTIDAD",
    "ULTIMA_ACTUALIZACION"
  ];
}

/** Campos con mínimo 50 / máximo 400 palabras (Sesión 1 tarjeta 1 + ConectaEduca segunda parte). */
var CAMPOS_SESION1_CON_RANGO_PALABRAS_ = ["REFLEXIONES", "DESAFIOS", "APUESTAS", "CONCLUSIONES", "PRIORIDADES_CE", "ACUERDOS_CE", "PROPUESTAS_CE", "RUTA_CE"];
var MIN_PALABRAS_SESION1_ = 50;
var MAX_PALABRAS_SESION1_ = 400;

/** Campos obligatorios para el envío definitivo de Sesión 1 (la síntesis colectiva). */
var CAMPOS_SESION1_OBLIGATORIOS_ = [
  "REFLEXIONES", "DESAFIOS", "APUESTAS", "CONCLUSIONES",
  "PRIORIDADES", "PROPUESTAS_COLECTIVAS", "ACUERDOS", "RUTA"
];

/** Campos obligatorios para el envío definitivo de Sesión 2 / ConectaEduca (Sesion1.gs los guarda, ConectaEduca.gs los valida). */
var CAMPOS_SESION2_OBLIGATORIOS_ = ["PRIORIDADES_CE", "ACUERDOS_CE", "PROPUESTAS_CE", "RUTA_CE"];

/** Todos los campos de contenido que se pueden guardar (obligatorios + aportes propios, opcionales). */
var CAMPOS_SESION1_ = cabecerasSesion1Comunal_().filter(function (c) {
  return c !== "ID_GRUPO" && c !== "ULTIMA_ACTUALIZACION";
});

/** Valida el rango de palabras (50-400) de los campos que lo exigen; devuelve los que están fuera de rango. */
function _validarRangoPalabrasSesion1_(datos, campos) {
  return campos.filter(function (c) {
    var n = contarPalabras_(datos ? datos[c] : "");
    return n < MIN_PALABRAS_SESION1_ || n > MAX_PALABRAS_SESION1_;
  });
}

/**
 * Guarda (UPSERT con fusión) los campos de Sesión 1 (y/o los aportes
 * propios de Sesión 1/Sesión 2) recibidos del cliente. `campos` es un
 * objeto parcial {NOMBRE_CAMPO: texto} — el cliente puede enviar solo los
 * campos que su dispositivo tiene abiertos/editados.
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

/**
 * Marca Sesión 1 como enviada definitivamente. Cualquier dispositivo con
 * sesión activa en el grupo puede enviarla (spec del usuario: revierte la
 * restricción de "solo el responsable principal" — "cualquiera pueda
 * enviar, pero al final solo debe haber un envío"); una vez enviada, no
 * se permiten más envíos (ver comprobación de SESION1_ENVIADA abajo).
 */
function enviarSesion1Definitiva(idGrupo, tokenSesion, dispositivoId) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  if (obtenerEstadoGrupo(idGrupo).sesion1Enviada) {
    return { ok: false, codigo: "YA_ENVIADO", mensaje: "Sesión 1 ya fue enviada de forma definitiva. No se permiten más envíos." };
  }

  var datos = obtenerSesion1(idGrupo);
  var vacios = CAMPOS_SESION1_OBLIGATORIOS_.filter(function (c) {
    return !datos || !String(datos[c] || "").trim();
  });
  if (vacios.length) {
    return { ok: false, mensaje: "Faltan campos por completar en Sesión 1: " + vacios.join(", ") };
  }
  var fueraDeRango = _validarRangoPalabrasSesion1_(datos, CAMPOS_SESION1_CON_RANGO_PALABRAS_.filter(function (c) {
    return CAMPOS_SESION1_OBLIGATORIOS_.indexOf(c) !== -1;
  }));
  if (fueraDeRango.length) {
    return {
      ok: false,
      mensaje: "Estos campos deben tener entre " + MIN_PALABRAS_SESION1_ + " y " + MAX_PALABRAS_SESION1_ +
        " palabras: " + fueraDeRango.join(", ")
    };
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

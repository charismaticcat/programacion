/**
 * Socializacion.gs — Foro Educativo Comunal Neiva 2026
 *
 * "Sesión de socialización" (Documento Orientador FEM2026, sección D,
 * item 10) — reemplaza a "Preparación" (item 9, oculta junto con todas
 * sus subpáginas): en vez de que cada IE llene un formulario previo, una
 * sola persona (spec del usuario: "una persona va a ingresar los datos")
 * lleva en vivo, durante el evento, un checklist de qué institución ya
 * socializó sus reflexiones, conclusiones y propuestas del Foro
 * Educativo Institucional (con un temporizador de apoyo por IE, ver
 * JS.html) y anota los datos relevantes de lo que cada una compartió.
 *
 * Mismo patrón de clave compuesta ID_GRUPO+ID_IE (columna CLAVE) que
 * PreparacionIE/ParticipacionEstamento (ver Preparacion.gs).
 */

var HOJA_SOCIALIZACION_IE_ = "SocializacionIE";

function cabecerasSocializacionIE_() {
  return ["CLAVE", "ID_GRUPO", "ID_IE", "SOCIALIZO", "DATOS_RELEVANTES", "ULTIMA_ACTUALIZACION"];
}

function _claveSocializacion_(idGrupo, idIE) {
  return String(idGrupo || "").trim() + "|" + String(idIE || "").trim();
}

/** Checklist de socialización de todas las IE del grupo, listo para pintar en pantalla. */
function obtenerSocializacionGrupo(idGrupo) {
  idGrupo = String(idGrupo || "").trim();
  var instituciones = obtenerInstitucionesDelGrupo(idGrupo);
  var hoja = obtenerHoja_(HOJA_SOCIALIZACION_IE_, cabecerasSocializacionIE_());
  var porClave = {};
  leerFilasComoObjetos_(hoja).forEach(function (f) {
    porClave[String(f.CLAVE || "")] = f;
  });

  return instituciones.map(function (ie) {
    var guardada = porClave[_claveSocializacion_(idGrupo, ie.idIE)];
    return {
      idIE: ie.idIE,
      institucion: ie.institucion,
      socializo: guardada ? String(guardada.SOCIALIZO || "") === "SI" : false,
      datosRelevantes: guardada ? String(guardada.DATOS_RELEVANTES || "") : ""
    };
  });
}

/**
 * Marca/desmarca una IE como ya socializada y guarda sus datos
 * relevantes (UPSERT, ambos juntos) — spec del usuario: al dar
 * "Finalizar" en el temporizador se abre la pantalla de datos relevantes
 * de esa IE, y de ahí se guarda todo de una vez.
 */
function guardarSocializacionIE(idGrupo, tokenSesion, dispositivoId, idIE, socializo, datosRelevantes) {
  idGrupo = String(idGrupo || "").trim();
  idIE = String(idIE || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  if (!idIE) return { ok: false, mensaje: "Falta la institución." };

  return conLock_(function () {
    upsertFila_(HOJA_SOCIALIZACION_IE_, cabecerasSocializacionIE_(), "CLAVE", _claveSocializacion_(idGrupo, idIE), {
      ID_GRUPO: idGrupo,
      ID_IE: idIE,
      SOCIALIZO: socializo ? "SI" : "NO",
      DATOS_RELEVANTES: String(datosRelevantes == null ? "" : datosRelevantes),
      ULTIMA_ACTUALIZACION: new Date()
    });
    return { ok: true };
  }, 15000);
}

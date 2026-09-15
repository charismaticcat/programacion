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

/**
 * PDF "Aportes relevantes del <Grupo>" (pedido del usuario): un
 * documento aparte del informe oficial del grupo, con los datos
 * relevantes que cada IE dejó registrados en la Sesión de socialización
 * — para consultarlo desde Sesión 1 mientras se responde el Consolidado.
 * Se regenera cada vez que se pide (se guarda un solo DOC_ID/PDF_ID por
 * grupo — la versión anterior se manda a la papelera antes de crear la
 * nueva, mismo criterio de "nunca acumular" del resto del proyecto).
 */
var HOJA_APORTES_RELEVANTES_SOCIALIZACION_ = "AportesRelevantesSocializacion";

function cabecerasAportesRelevantesSocializacion_() {
  return ["ID_GRUPO", "DOC_ID", "PDF_ID", "URL", "FECHA"];
}

/** PDF ya generado para el grupo, o null si todavía no se ha generado. */
function obtenerAportesRelevantesSocializacionGrupo_(idGrupo) {
  var hoja = obtenerHoja_(HOJA_APORTES_RELEVANTES_SOCIALIZACION_, cabecerasAportesRelevantesSocializacion_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", String(idGrupo || "").trim());
  return fila === -1 ? null : leerFilaComoObjeto_(hoja, fila, mapa);
}

function generarPdfAportesRelevantesSocializacion(idGrupo, tokenSesion, dispositivoId) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  var grupoInfo = obtenerGrupoPorId(idGrupo);
  if (!grupoInfo) return { ok: false, mensaje: "Grupo no encontrado." };

  var conAportes = obtenerSocializacionGrupo(idGrupo).filter(function (ie) {
    return String(ie.datosRelevantes || "").trim();
  });
  if (!conAportes.length) {
    return { ok: false, mensaje: "Todavía no hay datos relevantes registrados en la Sesión de socialización de este grupo." };
  }

  // Crear el Doc y exportarlo a PDF puede tardar varios segundos (llamadas
  // reales a Docs/Drive) — a propósito, esta parte NO va dentro de
  // conLock_(): el lock de Apps Script (LockService.getScriptLock()) es
  // de todo el proyecto, compartido por cualquier grupo que esté
  // guardando algo al mismo tiempo, así que mantenerlo tomado durante
  // varios segundos de trabajo lento podía sentirse como que el resto de
  // la app "se queda pegado" mientras tanto. Solo se usa el lock más
  // abajo, para la escritura final (rápida) en la hoja de seguimiento.
  var config = getConfig();
  var carpetaGrupo = asegurarCarpetaGrupo_(grupoInfo.grupo);
  var nombreBase = "Aportes relevantes del " + grupoInfo.grupo;

  var doc = DocumentApp.create(nombreBase);
  var docFile = DriveApp.getFileById(doc.getId());
  carpetaGrupo.addFile(docFile);
  try {
    DriveApp.getRootFolder().removeFile(docFile);
  } catch (e) {
    Logger.log("No fue posible quitar el Doc de aportes relevantes de la raíz de Drive: " + e.message);
  }

  var body = doc.getBody();
  body.clear();
  body.setPageWidth(612).setPageHeight(792).setMarginTop(30).setMarginBottom(30).setMarginLeft(50).setMarginRight(50);

  var pTitulo = body.appendParagraph(nombreBase);
  pTitulo.setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pTitulo.editAsText().setForegroundColor(COLOR_VERDE_INFORME_);
  var pSubtitulo = body.appendParagraph("Sesión de socialización — " + config.NOMBRE_FORO);
  pSubtitulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setItalic(true);
  body.appendParagraph(formatearFechaLargaEs_(new Date(), true)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  conAportes.forEach(function (ie) {
    titulo1_(body, ie.institucion);
    parrafo_(body, ie.datosRelevantes);
  });

  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  pdfBlob.setName(nombreBase + ".pdf");
  var pdfFile = carpetaGrupo.createFile(pdfBlob);
  hacerPublicoSiEsPosible_(pdfFile);
  var url = pdfFile.getUrl();

  return conLock_(function () {
    // Se regenera cada vez (nunca acumula copias): la versión anterior
    // (si la había, guardada en la hoja de seguimiento) se manda a la
    // papelera después de que la nueva ya quedó lista — así, si algo
    // falla arriba, la anterior sigue disponible en vez de perderse.
    var anterior = obtenerAportesRelevantesSocializacionGrupo_(idGrupo);
    if (anterior) {
      try { if (anterior.DOC_ID) DriveApp.getFileById(anterior.DOC_ID).setTrashed(true); } catch (e) { Logger.log("No se pudo eliminar el Doc anterior de aportes relevantes: " + e.message); }
      try { if (anterior.PDF_ID) DriveApp.getFileById(anterior.PDF_ID).setTrashed(true); } catch (e) { Logger.log("No se pudo eliminar el PDF anterior de aportes relevantes: " + e.message); }
    }
    upsertFila_(HOJA_APORTES_RELEVANTES_SOCIALIZACION_, cabecerasAportesRelevantesSocializacion_(), "ID_GRUPO", idGrupo, {
      DOC_ID: doc.getId(),
      PDF_ID: pdfFile.getId(),
      URL: url,
      FECHA: new Date()
    });
    return { ok: true, url: url };
  }, 15000);
}

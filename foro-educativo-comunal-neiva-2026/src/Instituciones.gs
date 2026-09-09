/**
 * Instituciones.gs — Foro Educativo Comunal Neiva 2026
 *
 * Las IE son miembros de un GRUPO (no la unidad raíz). Todo se consulta
 * contra GruposComunal, que es la fuente única de la relación Grupo → IE
 * (spec sección 6: "la información Grupo → IE debe provenir de Google
 * Sheets. NO codificar esta relación directamente en HTML o JavaScript").
 * Reemplaza el catálogo hardcodeado y duplicado de FEI 3.1
 * (GRUPOS_INSTITUCIONES en App.html + GRUPOS_INSTITUCIONES_FEM_ en
 * Código.js — ver docs/01-auditoria-fei-3.1.md §1.6/§7).
 *
 * CaracterizacionIE guarda los datos oficiales de directorio (equivalente
 * a la hoja `Oficiales` de FEI 3.1, ver auditoría §3.1) de forma
 * independiente: importados por lectura una sola vez
 * (Importacion.gs → importarCaracterizacionRealDesdeFEI31), y desde
 * entonces mantenidos por este proyecto, sin volver a tocar 3.1.
 * obtenerInstitucionesDelGrupo() es el único punto donde se combinan
 * ambas hojas — así, cualquier lugar que ya llama a esta función (acceso,
 * pantalla de Participación, generación del informe) recibe la
 * caracterización completa automáticamente, sin tener que pedirla aparte.
 */

var HOJA_GRUPOS_COMUNAL_ = "GruposComunal";
var HOJA_CARACTERIZACION_IE_ = "CaracterizacionIE";

function cabecerasGruposComunal_() {
  return ["ID_GRUPO", "GRUPO", "ID_IE", "INSTITUCION", "COMUNA", "ACTIVO"];
}

function cabecerasCaracterizacionIE_() {
  return ["ID_IE", "INSTITUCION", "DIRECCION", "SECTOR", "COMUNA", "ZONA", "EMAIL", "RECTOR", "SEDES", "LOGO_ID"];
}

/** Mapa {ID_IE: fila de CaracterizacionIE}, para enriquecer listados sin leer la hoja una vez por IE. */
function mapaCaracterizacionPorIE_() {
  var hoja = obtenerHoja_(HOJA_CARACTERIZACION_IE_, cabecerasCaracterizacionIE_());
  var mapa = {};
  leerFilasComoObjetos_(hoja).forEach(function (f) {
    var idIE = String(f.ID_IE || "").trim();
    if (idIE) mapa[idIE] = f;
  });
  return mapa;
}

/** Caracterización completa de una sola IE (dirección, sector, zona, email, rector, sedes), o null si no está cargada. */
function obtenerCaracterizacionIE(idIE) {
  var fila = mapaCaracterizacionPorIE_()[String(idIE || "").trim()];
  if (!fila) return null;
  return {
    idIE: String(fila.ID_IE || "").trim(),
    institucion: String(fila.INSTITUCION || "").trim(),
    direccion: String(fila.DIRECCION || "").trim(),
    sector: String(fila.SECTOR || "").trim(),
    comuna: String(fila.COMUNA || "").trim(),
    zona: String(fila.ZONA || "").trim(),
    email: String(fila.EMAIL || "").trim(),
    rector: String(fila.RECTOR || "").trim(),
    sedes: String(fila.SEDES || "").trim(),
    logoId: String(fila.LOGO_ID || "").trim()
  };
}

/**
 * Asigna el logo de una IE (mismo patrón que asignarLogoGrupo en Access.gs
 * — ID de archivo de Drive, no la imagen en sí). Uso: ejecutar desde el
 * editor de Apps Script una vez por IE cuando la SEM entregue los logos
 * reales. Si la IE todavía no tiene fila en CaracterizacionIE, la crea.
 */
function asignarLogoIE(idIE, logoFileId) {
  var hoja = obtenerHoja_(HOJA_CARACTERIZACION_IE_, cabecerasCaracterizacionIE_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var idIEStr = String(idIE || "").trim();
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_IE", idIEStr);
  if (fila === -1) {
    hoja.appendRow([idIEStr, "", "", "", "", "", "", "", "", String(logoFileId || "").trim()]);
    return { ok: true, creada: true };
  }
  hoja.getRange(fila, mapa["LOGO_ID"]).setValue(String(logoFileId || "").trim());
  return { ok: true, creada: false };
}

/**
 * Actualiza el nombre del rector(a) de una IE (editable desde la pantalla
 * de Confirmación de caracterización, igual que en FEI 3.1 — campo
 * "Rector(a) ✏️" de la caracterización). Solo se permite si la IE
 * pertenece al mismo grupo de la sesión activa: nunca se puede editar la
 * caracterización de una IE de otro grupo.
 */
function actualizarRectorIE(idGrupo, tokenSesion, dispositivoId, idIE, nombreRector) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  idIE = String(idIE || "").trim();
  var perteneceAlGrupo = obtenerInstitucionesDelGrupo(idGrupo).some(function (ie) {
    return ie.idIE === idIE;
  });
  if (!perteneceAlGrupo) return { ok: false, mensaje: "Esa institución no pertenece a este grupo." };

  upsertFila_(HOJA_CARACTERIZACION_IE_, cabecerasCaracterizacionIE_(), "ID_IE", idIE, {
    RECTOR: String(nombreRector || "").trim()
  });
  return { ok: true };
}

/**
 * Todas las IE (activas) que pertenecen a un grupo, tal como las cargó la
 * SEM en GruposComunal — enriquecidas automáticamente con su
 * caracterización completa desde CaracterizacionIE cuando existe (si una
 * IE todavía no tiene fila en CaracterizacionIE, esos campos quedan
 * vacíos en vez de romper el listado).
 */
function obtenerInstitucionesDelGrupo(idGrupo) {
  var hoja = obtenerHoja_(HOJA_GRUPOS_COMUNAL_, cabecerasGruposComunal_());
  var filas = leerFilasComoObjetos_(hoja);
  var objetivo = String(idGrupo || "").trim();
  var caracterizacion = mapaCaracterizacionPorIE_();

  return filas
    .filter(function (f) {
      return String(f.ID_GRUPO || "").trim() === objetivo && String(f.ACTIVO || "SI").toUpperCase() !== "NO";
    })
    .map(function (f) {
      var idIE = String(f.ID_IE || "").trim();
      var c = caracterizacion[idIE] || {};
      return {
        idIE: idIE,
        institucion: String(f.INSTITUCION || "").trim(),
        comuna: String(f.COMUNA || "").trim(),
        direccion: String(c.DIRECCION || "").trim(),
        sector: String(c.SECTOR || "").trim(),
        zona: String(c.ZONA || "").trim(),
        email: String(c.EMAIL || "").trim(),
        rector: String(c.RECTOR || "").trim(),
        sedes: String(c.SEDES || "").trim(),
        logoId: String(c.LOGO_ID || "").trim()
      };
    });
}

/**
 * Todas las IE activas de todos los grupos, en el orden de GruposComunal —
 * usada por la animación de bienvenida (carrusel de logos y nombres de
 * IE, spec: "aparece uno a uno cada logo de las IE"), que se muestra antes
 * de que el visitante ingrese el código de un grupo en particular, por lo
 * que no puede filtrarse todavía por ID_GRUPO.
 */
function obtenerTodasLasInstitucionesActivas() {
  var hoja = obtenerHoja_(HOJA_GRUPOS_COMUNAL_, cabecerasGruposComunal_());
  var filas = leerFilasComoObjetos_(hoja);
  var caracterizacion = mapaCaracterizacionPorIE_();
  return filas
    .filter(function (f) {
      return String(f.ACTIVO || "SI").toUpperCase() !== "NO";
    })
    .map(function (f) {
      var idIE = String(f.ID_IE || "").trim();
      var c = caracterizacion[idIE] || {};
      return {
        idIE: idIE,
        institucion: String(f.INSTITUCION || "").trim(),
        grupo: String(f.GRUPO || "").trim(),
        logoId: String(c.LOGO_ID || "").trim()
      };
    });
}

/** Busca a qué grupo pertenece una IE (por ID_IE o por nombre normalizado). */
function obtenerGrupoDeInstitucion_(identificadorIE) {
  var hoja = obtenerHoja_(HOJA_GRUPOS_COMUNAL_, cabecerasGruposComunal_());
  var filas = leerFilasComoObjetos_(hoja);
  var objetivo = String(identificadorIE || "").trim();
  var objetivoNormalizado = normalizarTexto_(identificadorIE);
  var encontrada = filas.find(function (f) {
    return (
      String(f.ID_IE || "").trim() === objetivo ||
      normalizarTexto_(f.INSTITUCION) === objetivoNormalizado
    );
  });
  return encontrada ? { idGrupo: String(encontrada.ID_GRUPO).trim(), grupo: String(encontrada.GRUPO).trim() } : null;
}

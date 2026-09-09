/**
 * Informes.gs — Foro Educativo Comunal Neiva 2026
 *
 * Genera el informe ÚNICO por grupo (spec sección 18: "GENERAR UN ÚNICO
 * INFORME POR GRUPO... El sistema debe conservar un único DOC_ID y PDF_ID
 * por grupo"). Motor adaptado de generarInformeFEM /
 * generarInformeSintesisGrupoFEM_ de FEI 3.1
 * (docs/01-auditoria-fei-3.1.md §1.5/§2.3/§2.5): DocumentApp.create desde
 * cero (nunca copia una plantilla externa — 3.1 abandonó ese enfoque por
 * incompatibilidad de formato, ver auditoría §7), gráficos vía Charts
 * insertados como imagen, ajuste de "título huérfano" con el servicio
 * avanzado Docs v1, exportación a PDF.
 *
 * A diferencia de 3.1, los helpers de estilo (titulo_/subtitulo_/parrafo_)
 * se definen UNA sola vez aquí (recomendación explícita de la auditoría
 * §7: en 3.1 estaban redefinidos función a función).
 */

var COLOR_VERDE_INFORME_ = "#0B6A44";
var COLOR_GRIS_TEXTO_INFORME_ = "#555555";
var COLOR_GRIS_BORDE_INFORME_ = "#CCCCCC";

/**
 * Cabeceras de InformesComunal — única fuente de verdad (evita el
 * problema de drift ya sufrido antes en este proyecto por repetir el
 * mismo arreglo de cabeceras en más de un archivo): DESCARGADO se agregó
 * para la condición de "descargado" que habilita enviar el informe por
 * correo (ver marcarInformeDescargado más abajo).
 */
function cabecerasInformesComunal_() {
  return ["ID_GRUPO", "DOC_ID", "PDF_ID", "URL", "FECHA", "ESTADO", "DESCARGADO"];
}

var _titulosKeepWithNext_ = [];

function titulo1_(body, texto) {
  var p = body.appendParagraph(texto.toUpperCase());
  p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  p.setSpacingBefore(16).setSpacingAfter(4);
  p.editAsText().setForegroundColor(COLOR_VERDE_INFORME_).setBold(true).setFontFamily(DocumentApp.FontFamily.ARIAL);
  _titulosKeepWithNext_.push(p.getText());
  return p;
}

function subtitulo_(body, texto) {
  var p = body.appendParagraph(texto.toUpperCase());
  p.setSpacingBefore(10).setSpacingAfter(2);
  p.editAsText().setForegroundColor(COLOR_VERDE_INFORME_).setBold(true).setFontSize(11).setFontFamily(DocumentApp.FontFamily.ARIAL);
  _titulosKeepWithNext_.push(p.getText());
  return p;
}

function parrafo_(body, texto) {
  var p = body.appendParagraph(String(texto || "Sin información registrada."));
  p.setSpacingBefore(0).setSpacingAfter(8);
  p.editAsText().setBold(false).setForegroundColor("#000000").setFontFamily(DocumentApp.FontFamily.ARIAL).setFontSize(11);
  return p;
}

function tablaSimple_(body, filas) {
  var tabla = body.appendTable(filas);
  tabla.setBorderColor(COLOR_GRIS_BORDE_INFORME_);
  for (var i = 0; i < tabla.getNumRows(); i++) {
    var fila = tabla.getRow(i);
    for (var j = 0; j < fila.getNumCells(); j++) {
      var celda = fila.getCell(j);
      celda.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      if (j === 0) celda.editAsText().setBold(true);
    }
  }
  return tabla;
}

function construirGraficoColumnas_(titulo, etiquetas, valores) {
  var dt = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Categoría").addColumn(Charts.ColumnType.NUMBER, "Cantidad");
  etiquetas.forEach(function (e, i) {
    dt.addRow([e, valores[i]]);
  });
  return Charts.newColumnChart()
    .setDataTable(dt.build())
    .setTitle(titulo)
    .setDimensions(560, 300)
    .setColors([COLOR_VERDE_INFORME_])
    .setOption("hAxis.slantedText", true)
    .setOption("hAxis.slantedTextAngle", 30)
    .setOption("chartArea", { left: 60, top: 40, width: "85%", height: "60%" })
    .build()
    .getAs("image/png");
}

/** Ajuste de "título huérfano" vía Docs v1 (requiere el servicio avanzado habilitado en appsscript.json). */
function aplicarKeepWithNextATitulosInforme_(docId, textosTitulos) {
  if (!textosTitulos || !textosTitulos.length) return;
  try {
    var documento = Docs.Documents.get(docId);
    var contenido = (documento.body && documento.body.content) || [];
    var pendientes = {};
    textosTitulos.forEach(function (t) {
      pendientes[t] = true;
    });
    var requests = [];
    contenido.forEach(function (elemento) {
      if (!elemento.paragraph) return;
      var texto = (elemento.paragraph.elements || [])
        .map(function (e) {
          return (e.textRun && e.textRun.content) || "";
        })
        .join("")
        .replace(/\n$/, "");
      if (pendientes[texto]) {
        requests.push({
          updateParagraphStyle: {
            range: { startIndex: elemento.startIndex, endIndex: elemento.endIndex },
            paragraphStyle: { keepWithNext: true },
            fields: "keepWithNext"
          }
        });
        delete pendientes[texto];
      }
    });
    if (requests.length) Docs.Documents.batchUpdate({ requests: requests }, docId);
  } catch (error) {
    Logger.log("No fue posible aplicar 'mantener con el siguiente' (el informe queda generado igual): " + error.message);
  }
}

/** Inserta el logo dado (por ID de Drive) si existe; nunca lanza si falla o el ID está vacío. */
function insertarLogoSiExiste_(parrafo, logoId, ancho, alto) {
  if (!logoId) return;
  try {
    parrafo.appendInlineImage(DriveApp.getFileById(logoId).getBlob()).setWidth(ancho).setHeight(alto);
  } catch (e) {
    Logger.log("No se pudo insertar el logo " + logoId + ": " + e.message);
  }
}

/**
 * Genera el informe único del grupo: Doc (privado, en la carpeta GRUPO N)
 * + PDF (público, misma carpeta) + registro en InformesComunal. Si el
 * grupo YA tiene un informe generado, lo regenera in situ (mismo DOC_ID)
 * en vez de crear uno nuevo — nunca hay más de un DOC_ID/PDF_ID por grupo
 * (spec sección 18).
 */
function generarInformeGrupo(idGrupo) {
  _titulosKeepWithNext_ = [];
  var grupoInfo = obtenerGrupoPorId(idGrupo);
  if (!grupoInfo) return { ok: false, mensaje: "Grupo no encontrado." };

  var config = getConfig();
  var instituciones = obtenerInstitucionesDelGrupo(idGrupo);
  var sesion1 = obtenerSesion1(idGrupo) || {};
  var conectaEduca = listarConectaEduca(idGrupo);
  var firmantes = listarFirmantesGrupo(idGrupo);
  var responsables = listarResponsablesEnvio(idGrupo);
  var matrizParticipacion = obtenerMatrizParticipacionGrupo(idGrupo);

  var carpetaGrupo = asegurarCarpetaGrupo_(grupoInfo.grupo);
  var nombreBase = "Informe " + grupoInfo.grupo;

  var doc = DocumentApp.create(nombreBase);
  var docFile = DriveApp.getFileById(doc.getId());
  carpetaGrupo.addFile(docFile);
  try {
    DriveApp.getRootFolder().removeFile(docFile);
  } catch (e) {
    Logger.log("No fue posible quitar el Doc de la raíz de Drive: " + e.message);
  }

  var body = doc.getBody();
  body.clear();
  body.setPageWidth(612).setPageHeight(792).setMarginTop(30).setMarginBottom(30).setMarginLeft(50).setMarginRight(50);

  // Encabezado con logos institucionales (SEM/Foro), igual patrón que 3.1.
  var header = doc.getHeader() || doc.addHeader();
  header.clear();
  var tablaEncabezado = header.appendTable([["", ""]]);
  tablaEncabezado.setBorderWidth(0);
  insertarLogoSiExiste_(tablaEncabezado.getCell(0, 0).getChild(0).asParagraph(), config.LOGO_PIE_ID, 90, 45);
  var pDerecho = tablaEncabezado.getCell(0, 1).getChild(0).asParagraph();
  pDerecho.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  insertarLogoSiExiste_(pDerecho, config.LOGO_ENCABEZADO_ID, 110, 60);

  var footer = doc.getFooter() || doc.addFooter();
  footer.clear();
  var fp = footer.appendParagraph(
    "Generado el " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy 'a las' HH:mm") +
      " — " + config.NOMBRE_FORO
  );
  fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  fp.editAsText().setForegroundColor(COLOR_GRIS_TEXTO_INFORME_).setFontSize(9);

  // Portada.
  var pTitulo = body.appendParagraph(config.NOMBRE_FORO);
  pTitulo.setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pTitulo.editAsText().setForegroundColor(COLOR_VERDE_INFORME_);
  var pSubtitulo = body.appendParagraph(config.SUBTITULO);
  pSubtitulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setItalic(true);
  var pGrupo = body.appendParagraph(grupoInfo.grupo);
  pGrupo.setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pGrupo.editAsText().setForegroundColor(COLOR_VERDE_INFORME_);
  body.appendParagraph(formatearFechaLargaEs_(config.FECHA, true)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendPageBreak();

  // Instituciones participantes.
  titulo1_(body, "Instituciones educativas del grupo");
  if (instituciones.length) {
    // Dirección/rector llegan automáticamente desde CaracterizacionIE
    // (obtenerInstitucionesDelGrupo() ya las combina, ver Instituciones.gs)
    // cuando esa hoja tiene la fila de la IE; si no, quedan en "—".
    var filasIE = [["Institución educativa", "Comuna", "Dirección", "Rector(a)"]].concat(
      instituciones.map(function (ie) {
        return [ie.institucion, ie.comuna || "—", ie.direccion || "—", ie.rector || "—"];
      })
    );
    tablaSimple_(body, filasIE);
  } else {
    parrafo_(body, "No hay instituciones registradas todavía para este grupo en GruposComunal.");
  }

  // Participación y asistencia.
  titulo1_(body, "Participación y asistencia");
  parrafo_(body, "Total de personas registradas: " + firmantes.length + ".");
  if (firmantes.length) {
    var tallyEstamentos = tallyOpciones_(firmantes, "estamento");
    var etiquetasEstamentos = Object.keys(tallyEstamentos);
    if (etiquetasEstamentos.length) {
      try {
        var imagenEstamentos = construirGraficoColumnas_(
          "Participación por estamento",
          etiquetasEstamentos,
          etiquetasEstamentos.map(function (r) {
            return tallyEstamentos[r];
          })
        );
        body.appendImage(imagenEstamentos);
      } catch (e) {
        Logger.log("No se pudo generar el gráfico de participación: " + e.message);
      }
    }
  }

  // Matriz de participación por estamento e IE (equivalente a la hoja
  // Participacion de 3.1, calculada en vivo — ver Data.gs).
  if (matrizParticipacion.estamentos.length && matrizParticipacion.instituciones.length) {
    subtitulo_(body, "Cantidad de asistentes por estamento e institución");
    var encabezadoMatriz = ["Estamento"].concat(
      matrizParticipacion.instituciones.map(function (ie) { return ie.institucion; }),
      ["Total"]
    );
    var filasMatriz = matrizParticipacion.estamentos.map(function (estamento) {
      var fila = [estamento];
      matrizParticipacion.instituciones.forEach(function (ie) {
        fila.push(String((matrizParticipacion.matriz[estamento][ie.idIE] || 0)));
      });
      fila.push(String(matrizParticipacion.totalesPorEstamento[estamento] || 0));
      return fila;
    });
    tablaSimple_(body, [encabezadoMatriz].concat(filasMatriz));
  }

  // Responsable de envío y asistentes de envío.
  subtitulo_(body, "Responsables de envío del grupo");
  if (responsables.principal) {
    parrafo_(
      body,
      "Principal: " + responsables.principal.nombre +
        (responsables.principal.rolForo ? " — " + responsables.principal.rolForo : "") +
        (responsables.principal.institucion ? " (" + responsables.principal.institucion + ")" : "") +
        " — " + responsables.principal.correo
    );
  } else {
    parrafo_(body, "El grupo todavía no registró un responsable de envío principal.");
  }
  if (responsables.asistentes.length) {
    responsables.asistentes.forEach(function (a) {
      parrafo_(
        body,
        "Asistente: " + a.nombre + (a.rolForo ? " — " + a.rolForo : "") +
          (a.institucion ? " (" + a.institucion + ")" : "") + " — " + a.correo
      );
    });
  }

  // Síntesis Sesión 1 — socialización.
  titulo1_(body, "Socialización de resultados institucionales (Sesión 1)");
  [
    ["Reflexiones", "REFLEXIONES"], ["Conclusiones", "CONCLUSIONES"], ["Propuestas de las IE", "PROPUESTAS_IE"],
    ["Experiencias", "EXPERIENCIAS"], ["Retos", "RETOS"], ["Aportes territoriales", "APORTES_TERRITORIALES"]
  ].forEach(function (par) {
    subtitulo_(body, par[0]);
    parrafo_(body, sesion1[par[1]]);
  });

  // Construcción colectiva del grupo.
  titulo1_(body, "Construcción colectiva del grupo");
  [
    ["Convergencias", "CONVERGENCIAS"], ["Apuestas compartidas", "APUESTAS"], ["Desafíos comunes", "DESAFIOS"],
    ["Identidad territorial", "IDENTIDAD"], ["Prioridades", "PRIORIDADES"],
    ["Propuestas colectivas", "PROPUESTAS_COLECTIVAS"], ["Acuerdos", "ACUERDOS"], ["Ruta de trabajo", "RUTA"]
  ].forEach(function (par) {
    subtitulo_(body, par[0]);
    parrafo_(body, sesion1[par[1]]);
  });

  // Aporte propio de la comunidad (opcional) para Sesión 1.
  if (String(sesion1.APORTE_PROPIO_S1_TITULO || "").trim() || String(sesion1.APORTE_PROPIO_S1_TEXTO || "").trim()) {
    subtitulo_(body, sesion1.APORTE_PROPIO_S1_TITULO || "Aporte propio del grupo");
    parrafo_(body, sesion1.APORTE_PROPIO_S1_TEXTO);
  }

  // ConectaEduca.
  titulo1_(body, "ConectaEduca — oportunidades de articulación");
  if (conectaEduca.length) {
    var filasCE = [["Actor / entidad", "Tipo", "Área", "Oportunidad", "IE interesadas"]].concat(
      conectaEduca.map(function (r) {
        return [r.ACTOR, r.TIPO_ACTOR, r.AREA, r.OPORTUNIDAD, r.IE_INTERESADAS];
      })
    );
    tablaSimple_(body, filasCE);
    conectaEduca.forEach(function (r) {
      if (r.NECESIDADES_ARTICULACION || r.ALIANZA || r.CONEXIONES || r.OBSERVACIONES) {
        subtitulo_(body, r.ACTOR);
        parrafo_(
          body,
          [
            r.NECESIDADES_ARTICULACION ? "Necesidades de articulación: " + r.NECESIDADES_ARTICULACION : "",
            r.ALIANZA ? "Posibles alianzas: " + r.ALIANZA : "",
            r.CONEXIONES ? "Conexiones realizadas: " + r.CONEXIONES : "",
            r.OBSERVACIONES ? "Observaciones: " + r.OBSERVACIONES : ""
          ]
            .filter(Boolean)
            .join("\n")
        );
      }
    });
  } else {
    parrafo_(body, "No se registraron actores de ConectaEduca para este grupo.");
  }

  // Aporte propio de la comunidad (opcional) para Sesión 2 / ConectaEduca.
  if (String(sesion1.APORTE_PROPIO_S2_TITULO || "").trim() || String(sesion1.APORTE_PROPIO_S2_TEXTO || "").trim()) {
    subtitulo_(body, sesion1.APORTE_PROPIO_S2_TITULO || "Aporte propio del grupo");
    parrafo_(body, sesion1.APORTE_PROPIO_S2_TEXTO);
  }

  // Insumos para el FEM 2026.
  titulo1_(body, "Insumos para el Foro Educativo Municipal FEM 2026");
  parrafo_(
    body,
    "Este informe consolida los resultados del " + grupoInfo.grupo + " del Foro Educativo Comunal Neiva 2026 " +
      "como insumo directo para el Foro Educativo Municipal FEM 2026."
  );

  doc.saveAndClose();
  aplicarKeepWithNextATitulosInforme_(doc.getId(), _titulosKeepWithNext_);

  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  pdfBlob.setName(nombreBase + ".pdf");
  var pdfFile = carpetaGrupo.createFile(pdfBlob);
  hacerPublicoSiEsPosible_(pdfFile);
  // El .docx editable NO se comparte públicamente (mismo criterio de
  // seguridad que 3.1 — ver docs/02-arquitectura-nuevo-proyecto.md Fase 7).

  var url = pdfFile.getUrl();
  upsertFila_(
    "InformesComunal",
    cabecerasInformesComunal_(),
    "ID_GRUPO",
    idGrupo,
    { DOC_ID: doc.getId(), PDF_ID: pdfFile.getId(), URL: url, FECHA: new Date(), ESTADO: "GENERADO", DESCARGADO: "NO" }
  );

  return { ok: true, docId: doc.getId(), pdfId: pdfFile.getId(), url: url };
}

/** Informe ya generado de un grupo (para que cualquier IE del grupo lo consulte — mismo informe para todas). */
function obtenerInformeGrupo(idGrupo) {
  var hoja = obtenerHoja_("InformesComunal", cabecerasInformesComunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", String(idGrupo || "").trim());
  if (fila === -1) return null;
  return leerFilaComoObjeto_(hoja, fila, mapa);
}

/**
 * Marca que el informe del grupo ya fue descargado — condición explícita
 * de esta entrega (junto con la valoración) para habilitar el envío del
 * informe por correo. Es un registro de "intención de descarga" (se
 * dispara al hacer clic en el enlace de descarga): Apps Script no puede
 * observar el evento real de descarga del navegador, así que se confía
 * en el mismo clic, igual que el resto de acciones autoreportadas de
 * esta aplicación (consentimientos, envíos de sesión, etc.).
 */
function marcarInformeDescargado(idGrupo, tokenSesion, dispositivoId) {
  idGrupo = String(idGrupo || "").trim();
  if (!sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)) {
    return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión ya no está activa en este dispositivo." };
  }
  var hoja = obtenerHoja_("InformesComunal", cabecerasInformesComunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", idGrupo);
  if (fila === -1) return { ok: false, mensaje: "El grupo todavía no tiene informe generado." };
  hoja.getRange(fila, mapa["DESCARGADO"]).setValue("SI");
  return { ok: true };
}

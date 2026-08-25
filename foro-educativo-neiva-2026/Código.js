/*****************************************************
 * FORO EDUCATIVO INSTITUCIONAL
 * Neiva 2026
 *
 * VERSIÓN INTEGRADA:
 * - Acceso personalizado por TOKEN + código.
 * - Un único ID_FORO oficial por IE.
 * - Carga automática de datos de la IE.
 * - Guardado local automático.
 * - Sincronización con AvancesForo.
 * - UPSERT por ID_FORO.
 * - No generar nuevos ID_FORO.
 * - Pruebas de validación y correo conservadas.
 *****************************************************/


/*****************************************************
 * CONFIGURACIÓN
 *****************************************************/

const SPREADSHEET_ID =
  "1ROYRM7hLY3qVQGifSKUQhgHgW8HK1iwB";

const HOJA_OFICIALES =
  "Oficiales";

const HOJA_AVANCES =
  "AvancesForo";

const HOJA_ACCESOS = "AccesosIE";
const HOJA_PARTICIPACION = "Participacion";
const DRIVE_CARPETA_FEM_ID = "1IqcFgQUSKocvGX3JwvNOu-xJzt0gfKc8";
const TEMPLATE_INFORME_ID = "1Gtsccdbnlcyjl6TcDDjTOA7pAW3JQbHM";
const LOGO_ENCABEZADO_ID = "1mFOOUZ5aFAuwMJMxNUaDnPPznDlQ2bj";
const LOGO_PIE_ID = "1Cmx7c3ec2gQCjRc8kcNeUbZt5LiURyD5";
const REMITENTE_FEM = "educacion@alcaldianeiva.gov.co";
const COPIAS_INFORME_FEM = [
  "adriana.cedeno@alcaldianeiva.gov.co",
  "angelica.rojas@alcaldianeiva.gov.co",
  "ronald.polania@alcaldianeiva.gov.co"
];


/*****************************************************
 * INCLUIR ARCHIVOS HTML
 *****************************************************/

function include(nombre) {

  return HtmlService
    .createHtmlOutputFromFile(nombre)
    .getContent();

}


/*****************************************************
 * OBTENER INSTITUCIONES EDUCATIVAS
 *****************************************************/

function doGet(e) {

  /*
   * =================================================
   * ACCESO POR URL PERSONALIZADA
   * =================================================
   *
   * Ejemplo:
   * .../dev?t=TOKEN
   */

  const parametros =
  (e && e.parameter)
    ? e.parameter
    : {};

const token =
  String(
    parametros.t ||
    parametros.token ||
    parametros.TOKEN ||
    ""
  ).trim();

  const plantilla =
    HtmlService.createTemplateFromFile("Index");

  plantilla.tokenAcceso = token;
  plantilla.nombreIEAcceso = "";

  /*
   * El servidor identifica la IE a partir del TOKEN.
   * El código de acceso se valida posteriormente.
   */
  if (token !== "") {

    try {

      const ss =
        SpreadsheetApp.openById(
          SPREADSHEET_ID
        );

      const hoja =
        ss.getSheetByName("AccesosIE");

      if (hoja) {

        const datos =
          hoja.getDataRange().getDisplayValues();

        if (datos.length > 1) {

          const cabeceras =
            datos[0].map(function(valor){
              return String(valor || "").trim();
            });

          const colIE =
            cabeceras.indexOf("IE");

          const colToken =
            cabeceras.indexOf("TOKEN");

          if (
            colIE !== -1 &&
            colToken !== -1
          ) {

            for (
              let i = 1;
              i < datos.length;
              i++
            ) {

              const tokenFila =
                String(
                  datos[i][colToken] || ""
                ).trim();

              if (tokenFila === token) {

                plantilla.nombreIEAcceso =
                  String(
                    datos[i][colIE] || ""
                  ).trim();

                break;

              }

            }

          }

        }

      }

    } catch (error) {

      console.error(
        "Error obteniendo nombre de IE:",
        error
      );

    }

  }

  return plantilla
    .evaluate()
    .setTitle(
      "Foro Educativo Institucional - Neiva 2026"
    )
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1"
    );

}

/*****************************************************
 * DEVOLVER INSTITUCIONES EN JSON
 *****************************************************/

function normalizarCabeceraInstituciones_(texto) {

  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[¿?]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toUpperCase();

}


function buscarColumnaInstituciones_(
  cabeceras,
  candidatos
) {

  const normalizadas =
    cabeceras.map(
      normalizarCabeceraInstituciones_
    );

  for (
    let i = 0;
    i < candidatos.length;
    i++
  ) {

    const buscada =
      normalizarCabeceraInstituciones_(
        candidatos[i]
      );

    const indice =
      normalizadas.indexOf(buscada);

    if (indice !== -1) {
      return indice;
    }

  }

  /*
   * Último recurso: detectar automáticamente una cabecera
   * que claramente represente el nombre de una institución.
   */
  for (let i = 0; i < normalizadas.length; i++) {

    const h = normalizadas[i];

    if (
      h.includes("DENOMINACION") ||
      (h.includes("INSTITUCION") && (
        h.includes("EDUCATIVA") ||
        h.includes("NOMBRE")
      )) ||
      (h.includes("ESTABLECIMIENTO") && h.includes("EDUCATIVO"))
    ) {
      return i;
    }

  }

  return -1;

}


/*****************************************************
 * OBTENER INSTITUCIONES EDUCATIVAS
 *
 * Lee la hoja Oficiales sin depender de que las
 * columnas estén exactamente en una posición fija.
 *
 * Devuelve:
 * {
 *   "NOMBRE IE": {
 *      dane,
 *      rector,
 *      direccion,
 *      zona,
 *      comuna,
 *      correo
 *   }
 * }
 *****************************************************/
function obtenerInstituciones() {

    const json =
        obtenerInstitucionesJSON();

    return JSON.parse(json);

}


function normalizarNombreIE_(texto) {

  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

}


function buscarInstitucionOficial_(instituciones, nombre) {

  if (
    !instituciones ||
    !nombre
  ) {
    return null;
  }

  if (
    instituciones[nombre]
  ) {
    return {
      nombre: nombre,
      datos: instituciones[nombre]
    };
  }

  const objetivo =
    normalizarNombreIE_(nombre);

  const clave =
    Object.keys(instituciones)
      .find(
        function(nombreIE){

          return (
            normalizarNombreIE_(nombreIE) ===
            objetivo
          );

        }
      );

  if(!clave){
    return null;
  }

  return {
    nombre: clave,
    datos: instituciones[clave]
  };

}


/*****************************************************
 * OBTENER INSTITUCIONES PARA EL FORMULARIO
 *
 * La hoja "Oficiales" tiene los encabezados reales
 * en la FILA 5.
 *****************************************************/

function obtenerInstitucionesJSON(){

    try{

        const ss =
            SpreadsheetApp.openById(
                SPREADSHEET_ID
            );

        const hoja =
            ss.getSheetByName(
                "Oficiales"
            );

        if(!hoja){

            throw new Error(
                'No existe la hoja "Oficiales".'
            );

        }


        /*
         * =============================================
         * LA FILA 5 CONTIENE LOS ENCABEZADOS REALES
         * =============================================
         */

        const FILA_ENCABEZADOS = 5;

        const ultimaFila =
            hoja.getLastRow();

        const ultimaColumna =
            hoja.getLastColumn();


        if(
            ultimaFila <= FILA_ENCABEZADOS ||
            ultimaColumna < 8
        ){

            throw new Error(
                "La hoja Oficiales no contiene datos suficientes."
            );

        }


        /*
         * Leer encabezados de la fila 5.
         */

        const cabeceras =
            hoja
                .getRange(
                    FILA_ENCABEZADOS,
                    1,
                    1,
                    ultimaColumna
                )
                .getDisplayValues()[0]
                .map(function(valor){

                    return String(
                        valor || ""
                    )
                    .trim()
                    .toUpperCase();

                });


        const colIE =
            cabeceras.indexOf(
                "INSTITUCIÓN/SEDE"
            );

        const colDane =
            cabeceras.indexOf(
                "CODIGO DANE"
            );

        const colDireccion =
            cabeceras.indexOf(
                "DIRECCIÓN"
            );

        const colSector =
            cabeceras.indexOf(
                "SECTOR"
            );

        const colComuna =
            cabeceras.indexOf(
                "COMUNA"
            );

        const colZona =
            cabeceras.indexOf(
                "ZONA"
            );

        const colCorreo =
            cabeceras.indexOf(
                "E-MAIL INSTITUCIONAL"
            );

        const colRector =
            cabeceras.indexOf(
                "RECTOR (A)"
            );


        /*
         * Verificar columna indispensable.
         */

        if(colIE === -1){

            throw new Error(
                'No se encontró la columna "Institución/Sede" en la fila 5.'
            );

        }


        /*
         * =============================================
         * LEER TODAS LAS INSTITUCIONES
         * =============================================
         */

        const datos =
            hoja
                .getRange(
                    FILA_ENCABEZADOS + 1,
                    1,
                    ultimaFila - FILA_ENCABEZADOS,
                    ultimaColumna
                )
                .getDisplayValues();


        const instituciones = {};


        datos.forEach(function(fila){

            const nombreIE =
    String(
        fila[colIE] || ""
    ).trim();


/*
 * Ignorar filas vacías.
 */
if(!nombreIE){
    return;
}


/*
 * =================================================
 * MOSTRAR SOLO IE CENTRALES
 * =================================================
 *
 * En la hoja "Oficiales":
 *
 * - Las IE centrales están escritas en MAYÚSCULAS.
 * - Las sedes aparecen con escritura normal.
 *
 * Por tanto, excluimos cualquier registro que
 * no esté completamente escrito en mayúsculas.
 */

const nombreMayusculas =
    nombreIE.toUpperCase();

const nombreMinusculas =
    nombreIE.toLowerCase();


if(
    nombreIE !== nombreMayusculas ||
    nombreIE === nombreMinusculas
){
    return;
}


            /*
             * Evitar duplicados.
             *
             * Si una institución aparece varias veces
             * por tener varias sedes, conservamos la
             * primera ficha encontrada.
             */

            if(
                instituciones[nombreIE]
            ){

                return;

            }


            const dane =
                colDane >= 0
                    ? String(
                        fila[colDane] || ""
                    ).trim()
                    : "";


            const direccion =
                colDireccion >= 0
                    ? String(
                        fila[colDireccion] || ""
                    ).trim()
                    : "";


            const sector =
                colSector >= 0
                    ? String(
                        fila[colSector] || ""
                    ).trim()
                    : "";


            const comuna =
                colComuna >= 0
                    ? String(
                        fila[colComuna] || ""
                    ).trim()
                    : "";


            const zona =
                colZona >= 0
                    ? String(
                        fila[colZona] || ""
                    ).trim()
                    : "";


            const correoIE =
                colCorreo >= 0
                    ? String(
                        fila[colCorreo] || ""
                    ).trim()
                    : "";


            const rector =
                colRector >= 0
                    ? String(
                        fila[colRector] || ""
                    ).trim()
                    : "";


            /*
             * Determinar grupo usando el catálogo
             * GRUPOS_INSTITUCIONES que ya existe
             * en el proyecto.
             */

            let grupo = "";

            try{

                if(
                    typeof obtenerGrupoInstitucion ===
                    "function"
                ){

                    grupo =
                        obtenerGrupoInstitucion(
                            nombreIE
                        );

                }

            }catch(error){

                console.warn(
                    "No se pudo determinar grupo para:",
                    nombreIE,
                    error
                );

            }


            instituciones[nombreIE] = {

                dane:
                    dane,

                rector:
                    rector,

                direccion:
                    direccion,

                sector:
                    sector,

                comuna:
                    comuna,

                zona:
                    zona,

                correo:
    correoIE,

                grupo:
                    grupo

            };

        });


        /*
         * Registrar resultado.
         */

        console.log(
            "Instituciones cargadas:",
            Object.keys(
                instituciones
            ).length
        );


        return JSON.stringify(
            instituciones
        );


    }catch(error){

        console.error(
            "obtenerInstitucionesJSON:",
            error
        );


        /*
         * No ocultar el error.
         *
         * El frontend recibirá el error real.
         */

        throw new Error(
            "No fue posible cargar las instituciones: " +
            error.message
        );

    }

}


/*****************************************************
 * DIAGNÓSTICO DE LA HOJA OFICIALES
 *****************************************************/

function diagnosticarOficiales() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName(HOJA_OFICIALES);

  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_OFICIALES + '".');
  }

  const datos = hoja.getDataRange().getDisplayValues();
  const cabeceras = datos.length ? datos[0] : [];

  Logger.log("========================================");
  Logger.log("CABECERAS DE OFICIALES:");
  cabeceras.forEach(function(cabecera, i) {
    Logger.log(i + " -> [" + String(cabecera || "") + "]");
  });
  Logger.log("========================================");

  return cabeceras;

}


function diagnosticarOficiales() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const hoja =
    ss.getSheetByName(
      HOJA_OFICIALES
    );

  if (!hoja) {
    throw new Error(
      'No existe la hoja "' +
      HOJA_OFICIALES +
      '".'
    );
  }

  const ultimaFila =
    hoja.getLastRow();

  const ultimaColumna =
    hoja.getLastColumn();

  Logger.log(
    "========================================"
  );

  Logger.log(
    "DIAGNÓSTICO REAL DE LA HOJA OFICIALES"
  );

  Logger.log(
    "Nombre hoja: " +
    hoja.getName()
  );

  Logger.log(
    "Última fila: " +
    ultimaFila
  );

  Logger.log(
    "Última columna: " +
    ultimaColumna
  );

  Logger.log(
    "========================================"
  );


  /*
   * Leer las primeras filas y columnas.
   *
   * No asumimos todavía dónde están
   * los encabezados.
   */

  const filasARevisar =
    Math.min(
      ultimaFila,
      25
    );

  const columnasARevisar =
    Math.min(
      ultimaColumna,
      20
    );


  if(
    filasARevisar === 0 ||
    columnasARevisar === 0
  ){

    Logger.log(
      "La hoja no contiene datos."
    );

    return;
  }


  const datos =
    hoja
      .getRange(
        1,
        1,
        filasARevisar,
        columnasARevisar
      )
      .getDisplayValues();


  Logger.log(
    "========================================"
  );

  Logger.log(
    "PRIMERAS " +
    filasARevisar +
    " FILAS:"
  );

  Logger.log(
    "========================================"
  );


  datos.forEach(
    function(fila, numeroFila){

      const valores =
        fila.map(
          function(valor, numeroColumna){

            const texto =
              String(
                valor || ""
              ).trim();

            if(!texto){
              return null;
            }

            return (
              "C" +
              (numeroColumna + 1) +
              "=" +
              texto
            );

          }
        )
        .filter(Boolean);


      Logger.log(
        "FILA " +
        (numeroFila + 1) +
        " -> " +
        (
          valores.length
            ? valores.join(" | ")
            : "[VACÍA]"
        )
      );

    }
  );


  /*
   * Revisar celdas combinadas.
   */

  Logger.log(
    "========================================"
  );

  Logger.log(
    "CELDAS COMBINADAS:"
  );

  const rangosCombinados =
    hoja.getDataRange()
      .getMergedRanges();


  if(
    rangosCombinados.length === 0
  ){

    Logger.log(
      "No hay celdas combinadas."
    );

  }else{

    rangosCombinados.forEach(
      function(rango){

        Logger.log(
          rango.getA1Notation()
        );

      }
    );

  }


  Logger.log(
    "========================================"
  );

  Logger.log(
    "FIN DEL DIAGNÓSTICO"
  );

  return {
    filas: ultimaFila,
    columnas: ultimaColumna,
    datos: datos
  };

}


/*****************************************************
 * PROBAR INSTITUCIONES
 *****************************************************/

function probarInstituciones() {

  const datos =
    obtenerInstituciones();

  const nombres =
    Object.keys(datos);


  Logger.log(
    "TOTAL DE INSTITUCIONES: " +
    nombres.length
  );


  Logger.log(
    "PRIMERAS INSTITUCIONES: " +
    nombres
      .slice(0, 10)
      .join(" | ")
  );


  Logger.log(
    "EL LIMONAR: " +
    JSON.stringify(
      datos["EL LIMONAR"]
    )
  );

}


/*****************************************************
 * CABECERAS DE AVANCESFORO
 *
 * Esta es la estructura de la ETAPA 1.
 *
 * DATOS se conserva temporalmente como respaldo
 * y compatibilidad con la aplicación existente.
 *****************************************************/

function obtenerCabecerasAvancesForo() {

  return [

    "ID_FORO",
    "INSTITUCION",
    "DANE",
    "FECHA_INICIO",
    "ULTIMA_ACTUALIZACION",
    "ESTADO",

    "S1_P1",
    "S1_P2",

    "S2_P1",

    "S2_P2_ACCION_1",
    "S2_P2_ACCION_2",
    "S2_P2_ACCION_3",
    "S2_P2_ACCION_4",
    "S2_P2_ACCION_5",

    "S2_P3",
    "S2_P4",
    "S2_P5",

    "S3_P1",
    "S3_P2_ACCION_1",
    "S3_P2_ACCION_2",
    "S3_P2_ACCION_3",
    "S3_P2_ACCION_4",
    "S3_P2_ACCION_5",
    "S3_P3",
    "S3_P4",

    "S1_ENVIADA",
    "S2_ENVIADA",
    "S3_ENVIADA",
    "FECHA_ENVIO_S1",
    "FECHA_ENVIO_S2",
    "FECHA_ENVIO_S3",
    "FECHA_ENVIO_DEFINITIVO",
    "ID_INFORME",
    "ID_PDF_INFORME",

    "DATOS"

  ];

}


/*****************************************************
 * CREAR / ACTUALIZAR ESTRUCTURA DE AVANCESFORO
 *
 * No borra datos existentes.
 * No elimina columnas.
 *
 * Si la hoja no existe:
 *   la crea con la estructura nueva.
 *
 * Si ya existe:
 *   verifica las cabeceras.
 *   agrega únicamente las que falten al final.
 *****************************************************/

function prepararHojaAvancesForo() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


  let hoja =
    ss.getSheetByName(
      HOJA_AVANCES
    );


  if (!hoja) {

    hoja =
      ss.insertSheet(
        HOJA_AVANCES
      );

  }


  const cabeceras =
    obtenerCabecerasAvancesForo();


  const ultimaColumna =
    hoja.getLastColumn();


  if (ultimaColumna === 0) {

    hoja
      .getRange(
        1,
        1,
        1,
        cabeceras.length
      )
      .setValues([
        cabeceras
      ]);

  } else {

    const existentes =
      hoja
        .getRange(
          1,
          1,
          1,
          ultimaColumna
        )
        .getValues()[0]
        .map(String);


    const faltantes =
      cabeceras.filter(
        function(cabecera) {

          return existentes
            .indexOf(cabecera) === -1;

        }
      );


    if (faltantes.length > 0) {

      hoja
        .getRange(
          1,
          ultimaColumna + 1,
          1,
          faltantes.length
        )
        .setValues([
          faltantes
        ]);

    }

  }


  return {
    ok: true,
    hoja: HOJA_AVANCES,
    cabeceras:
      obtenerMapaCabeceras_(hoja)
  };

}


/*****************************************************
 * MAPA DE CABECERAS
 *****************************************************/

function obtenerMapaCabeceras_(hoja) {

  const ultimaColumna =
    hoja.getLastColumn();


  if (ultimaColumna === 0) {
    return {};
  }


  const cabeceras =
    hoja
      .getRange(
        1,
        1,
        1,
        ultimaColumna
      )
      .getValues()[0];


  const mapa = {};


  cabeceras.forEach(
    function(nombre, indice) {

      const clave =
        String(
          nombre || ""
        ).trim();


      if (clave !== "") {

        mapa[clave] =
          indice + 1;

      }

    }
  );


  return mapa;

}


/*****************************************************
 * OBTENER VALOR DE CAMPO DEL FORMULARIO
 *
 * El autoguardado actual guarda los campos dentro de:
 *
 * datos.campos
 *
 * Se mantiene compatibilidad con el objeto completo.
 *****************************************************/

function obtenerCampoFormulario_(
  campos,
  nombre
) {

  if (
    !campos ||
    typeof campos !== "object"
  ) {

    return "";

  }


  const valor =
    campos[nombre];


  if (
    valor === null ||
    valor === undefined
  ) {

    return "";

  }


  return valor;

}


/*****************************************************
 * NORMALIZAR VALOR PARA SHEETS
 *****************************************************/

function normalizarValorHoja_(valor) {

  if (
    valor === null ||
    valor === undefined
  ) {

    return "";

  }


  if (
    Array.isArray(valor)
  ) {

    return valor.join("\n");

  }


  if (
    typeof valor === "object"
  ) {

    return JSON.stringify(
      valor
    );

  }


  return valor;

}


/*****************************************************
 * EXTRAER RESPUESTAS DE SESIÓN 1 Y SESIÓN 2
 *
 * Esta función NO modifica el objeto original.
 *****************************************************/

function extraerRespuestasSesiones_(datos) {

  const campos =
    datos &&
    datos.campos &&
    typeof datos.campos === "object"
      ? datos.campos
      : {};


  return {

    S1_P1:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion1"
      ),

    S1_P2:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion1Pregunta2"
      ),


    S2_P1:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta1"
      ),


    S2_P2_ACCION_1:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta2Accion1"
      ),

    S2_P2_ACCION_2:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta2Accion2"
      ),

    S2_P2_ACCION_3:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta2Accion3"
      ),

    S2_P2_ACCION_4:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta2Accion4"
      ),

    S2_P2_ACCION_5:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta2Accion5"
      ),


    S2_P3:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta3"
      ),


    S2_P4:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta4"
      ),


    S2_P5:
      obtenerCampoFormulario_(
        campos,
        "respuestaSesion2Pregunta5"
      ),

    S3_P1:
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta1"),
    S3_P2_ACCION_1:
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta2Accion1"),
    S3_P2_ACCION_2:
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta2Accion2"),
    S3_P2_ACCION_3:
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta2Accion3"),
    S3_P2_ACCION_4:
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta2Accion4"),
    S3_P2_ACCION_5:
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta2Accion5"),
    S3_P3:
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta3"),
    S3_P4:
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta4")

  };

}


/*****************************************************
 * BUSCAR FILA POR ID_FORO
 *****************************************************/

function buscarFilaPorIdForo_(
  hoja,
  idForo,
  mapaCabeceras
) {

  const columnaId =
    mapaCabeceras["ID_FORO"];


  if (!columnaId) {

    throw new Error(
      "AvancesForo no tiene la columna ID_FORO."
    );

  }


  const ultimaFila =
    hoja.getLastRow();


  if (ultimaFila < 2) {
    return -1;
  }


  const valores =
    hoja
      .getRange(
        2,
        columnaId,
        ultimaFila - 1,
        1
      )
      .getValues();


  const objetivo =
    String(idForo).trim();


  for (
    let i = 0;
    i < valores.length;
    i++
  ) {

    if (
      String(
        valores[i][0] || ""
      ).trim() === objetivo
    ) {

      return i + 2;

    }

  }


  return -1;

}


/*****************************************************
 * GUARDAR AVANCE DEL FORO
 *
 * ETAPA 1
 *
 * REGLAS:
 *
 * 1. ID_FORO es obligatorio.
 * 2. El servidor NO genera otro ID.
 * 3. Un ID_FORO = una sola fila.
 * 4. Si existe, actualiza.
 * 5. Si no existe, crea.
 * 6. Conserva DATOS como respaldo.
 * 7. Utiliza LockService para evitar duplicados.
 *****************************************************/

function obtenerAccesoPorIdForo_(idForo) {

  const id =
    String(idForo || "").trim();

  if (!id) {
    return null;
  }

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const hoja =
    ss.getSheetByName(
      "AccesosIE"
    );

  if (!hoja) {
    return null;
  }

  const datos =
    hoja.getDataRange().getDisplayValues();

  if (datos.length < 2) {
    return null;
  }

  const cabeceras =
    datos[0];

  const mapa = {};

  cabeceras.forEach(
    function(nombre, indice) {

      const clave =
        String(
          nombre || ""
        ).trim();

      if (clave) {
        mapa[clave] = indice;
      }

    }
  );

  const requeridas = [
    "IE",
    "DANE",
    "ID_FORO"
  ];

  for (
    let i = 0;
    i < requeridas.length;
    i++
  ) {

    if (
      mapa[requeridas[i]] === undefined
    ) {
      return null;
    }

  }

  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    const idFila =
      String(
        datos[i][mapa["ID_FORO"]] || ""
      ).trim();

    if (idFila !== id) {
      continue;
    }

    const estado =
      mapa["ESTADO"] !== undefined
        ? String(
            datos[i][mapa["ESTADO"]] || ""
          ).trim().toUpperCase()
        : "";

    if (
      estado === "ENVIADO" ||
      estado === "BLOQUEADO" ||
      estado === "INACTIVO"
    ) {
      return null;
    }

    return {

      fila: i + 1,

      ie:
        String(
          datos[i][mapa["IE"]] || ""
        ).trim(),

      dane:
        String(
          datos[i][mapa["DANE"]] || ""
        ).trim(),

      estado:
        estado

    };

  }

  return null;

}



/*****************************************************
 * RECUPERAR AVANCE GUARDADO EN SERVIDOR
 *
 * Permite continuar un mismo ID_FORO desde otro
 * dispositivo después de una eventualidad.
 *
 * NO crea ni modifica el avance.
 * Solo devuelve la copia guardada en AvancesForo.
 *****************************************************/
function obtenerAvanceForo(idForo) {

  const lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(10000);

    idForo =
      String(idForo || "").trim();

    if (!idForo) {

      return {
        ok: false,
        codigo: "ID_FORO_REQUERIDO",
        mensaje: "No se recibió el ID_FORO."
      };

    }

    const acceso =
      obtenerAccesoPorIdForo_(idForo);

    if (!acceso) {

      return {
        ok: false,
        codigo: "ID_FORO_NO_AUTORIZADO",
        mensaje: "El ID_FORO no está autorizado."
      };

    }

    const ss =
      SpreadsheetApp.openById(
        SPREADSHEET_ID
      );

    const hoja =
      ss.getSheetByName(
        HOJA_AVANCES
      );

    if (!hoja || hoja.getLastRow() < 2) {

      return {
        ok: true,
        encontrado: false,
        idForo: idForo
      };

    }

    const mapa =
      obtenerMapaCabeceras_(
        hoja
      );

    if (!mapa["ID_FORO"] || !mapa["DATOS"]) {

      return {
        ok: true,
        encontrado: false,
        idForo: idForo
      };

    }

    const ultimaFila =
      hoja.getLastRow();

    const datos =
      hoja
        .getRange(
          2,
          1,
          ultimaFila - 1,
          hoja.getLastColumn()
        )
        .getValues();

    for (
      let i = 0;
      i < datos.length;
      i++
    ) {

      const idFila =
        String(
          datos[i][mapa["ID_FORO"] - 1] || ""
        ).trim();

      if (idFila !== idForo) {
        continue;
      }

      const estado =
        mapa["ESTADO"]
          ? String(
              datos[i][mapa["ESTADO"] - 1] || ""
            ).trim().toUpperCase()
          : "";

      if (
        estado === "ENVIADO" ||
        estado === "BLOQUEADO" ||
        estado === "INACTIVO"
      ) {

        return {
          ok: true,
          encontrado: false,
          bloqueado: true,
          idForo: idForo
        };

      }

      const respaldo =
        String(
          datos[i][mapa["DATOS"] - 1] || ""
        ).trim();

      if (!respaldo) {

        return {
          ok: true,
          encontrado: false,
          idForo: idForo
        };

      }

      let borrador = null;

      try {

        borrador =
          JSON.parse(
            respaldo
          );

      } catch (error) {

        return {
          ok: false,
          codigo: "DATOS_INVALIDOS",
          mensaje:
            "El avance guardado en el servidor no pudo ser leído."
        };

      }

      return {
        ok: true,
        encontrado: true,
        idForo: idForo,
        institucion: acceso.ie,
        dane: acceso.dane,
        datos: borrador
      };

    }

    return {
      ok: true,
      encontrado: false,
      idForo: idForo
    };

  } catch (error) {

    return {
      ok: false,
      codigo: "ERROR_RECUPERANDO_AVANCE",
      mensaje:
        error.message
    };

  } finally {

    try {
      lock.releaseLock();
    } catch (error) {
      // No hacer nada.
    }

  }

}


function guardarAvanceForo(datos) {

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(15000);


    if (
      !datos ||
      typeof datos !== "object"
    ) {

      throw new Error(
        "No se recibieron datos válidos."
      );

    }


    /*
     * ID_FORO YA NO SE GENERA EN EL SERVIDOR.
     */

    const idForo =
      String(
        datos.idForo || ""
      ).trim();


    if (idForo === "") {

      throw new Error(
        "ID_FORO obligatorio. El servidor no puede generar un nuevo ID."
      );

    }


    /*
     * =================================================
     * VALIDAR ID_FORO OFICIAL
     * =================================================
     *
     * El navegador nunca puede inventar un ID.
     * El ID debe existir en AccesosIE.
     */
    const accesoForo =
      obtenerAccesoPorIdForo_(idForo);

    if (!accesoForo) {

      throw new Error(
        "El ID_FORO no está autorizado. " +
        "Solicite un nuevo acceso a la Secretaría de Educación de Neiva " +
        "al correo educacion@alcaldianeiva.gov.co " +
        "o al WhatsApp 318 456 1081."
      );

    }

    /*
     * La institución y el DANE se fuerzan desde
     * AccesosIE para evitar inconsistencias.
     */
    datos.institucion =
      accesoForo.ie;

    datos.dane =
      accesoForo.dane;


    const ss =
      SpreadsheetApp.openById(
        SPREADSHEET_ID
      );


    let hoja =
      ss.getSheetByName(
        HOJA_AVANCES
      );


    /*
     * Garantizar estructura.
     */

    if (!hoja) {

      hoja =
        ss.insertSheet(
          HOJA_AVANCES
        );

    }


    prepararHojaAvancesForo();


    /*
     * Volver a obtener la hoja por seguridad.
     */

    hoja =
      ss.getSheetByName(
        HOJA_AVANCES
      );


    const mapaCabeceras =
      obtenerMapaCabeceras_(
        hoja
      );


    /*
     * Verificar cabeceras fundamentales.
     */

    const cabecerasObligatorias = [

      "ID_FORO",
      "INSTITUCION",
      "DANE",
      "FECHA_INICIO",
      "ULTIMA_ACTUALIZACION",
      "ESTADO",
      "S1_P1",
      "S1_P2",
      "S2_P1",
      "S2_P2_ACCION_1",
      "S2_P2_ACCION_2",
      "S2_P2_ACCION_3",
      "S2_P2_ACCION_4",
      "S2_P2_ACCION_5",
      "S2_P3",
      "S2_P4",
      "S2_P5",
      "DATOS"

    ];


    cabecerasObligatorias.forEach(
      function(cabecera) {

        if (
          !mapaCabeceras[cabecera]
        ) {

          throw new Error(
            "Falta la columna " +
            cabecera +
            " en AvancesForo."
          );

        }

      }
    );


    const ahora =
      new Date();


    const filaExistente =
      buscarFilaPorIdForo_(
        hoja,
        idForo,
        mapaCabeceras
      );


    /*
     * Extraer respuestas estructuradas.
     */

    const respuestas =
      extraerRespuestasSesiones_(
        datos
      );


    /*
     * FECHA_INICIO:
     *
     * En una actualización no se reemplaza
     * la fecha original.
     */

    let fechaInicio =
      datos.fechaInicio ||
      "";


    if (
      filaExistente !== -1
    ) {

      const fechaActual =
        hoja.getRange(
          filaExistente,
          mapaCabeceras["FECHA_INICIO"]
        ).getValue();


      if (
        fechaActual !== "" &&
        fechaActual !== null
      ) {

        fechaInicio =
          fechaActual;

      }

    }


    if (
      fechaInicio === ""
    ) {

      fechaInicio =
        ahora;

    }


    /*
     * DATOS:
     *
     * Se conserva temporalmente para
     * compatibilidad y recuperación.
     */

    const datosRespaldo =
      JSON.stringify(
        datos
      );


    /*
     * Construir únicamente las columnas
     * conocidas en esta etapa.
     */

    const valoresPorCabecera = {

      "ID_FORO":
        idForo,

      "INSTITUCION":
        datos.institucion || "",

      "DANE":
        datos.dane || "",

      "FECHA_INICIO":
        fechaInicio,

      "ULTIMA_ACTUALIZACION":
        ahora,

      "ESTADO":
        "En proceso",

      "S1_P1":
        normalizarValorHoja_(
          respuestas.S1_P1
        ),

      "S1_P2":
        normalizarValorHoja_(
          respuestas.S1_P2
        ),

      "S2_P1":
        normalizarValorHoja_(
          respuestas.S2_P1
        ),

      "S2_P2_ACCION_1":
        normalizarValorHoja_(
          respuestas.S2_P2_ACCION_1
        ),

      "S2_P2_ACCION_2":
        normalizarValorHoja_(
          respuestas.S2_P2_ACCION_2
        ),

      "S2_P2_ACCION_3":
        normalizarValorHoja_(
          respuestas.S2_P2_ACCION_3
        ),

      "S2_P2_ACCION_4":
        normalizarValorHoja_(
          respuestas.S2_P2_ACCION_4
        ),

      "S2_P2_ACCION_5":
        normalizarValorHoja_(
          respuestas.S2_P2_ACCION_5
        ),

      "S2_P3":
        normalizarValorHoja_(
          respuestas.S2_P3
        ),

      "S2_P4":
        normalizarValorHoja_(
          respuestas.S2_P4
        ),

      "S2_P5":
        normalizarValorHoja_(
          respuestas.S2_P5
        ),

      "DATOS":
        datosRespaldo

    };


    /*
     * Si es actualización:
     * solamente se actualiza la fila existente.
     */

    if (
      filaExistente !== -1
    ) {

      Object.keys(
        valoresPorCabecera
      ).forEach(
        function(cabecera) {

          const columna =
            mapaCabeceras[cabecera];


          if (columna) {

            hoja
              .getRange(
                filaExistente,
                columna
              )
              .setValue(
                valoresPorCabecera[
                  cabecera
                ]
              );

          }

        }
      );


    } else {

      /*
       * Crear una única fila nueva.
       */

      const ultimaColumna =
        hoja.getLastColumn();


      const filaNueva =
        new Array(
          ultimaColumna
        ).fill("");


      Object.keys(
        valoresPorCabecera
      ).forEach(
        function(cabecera) {

          const columna =
            mapaCabeceras[cabecera];


          if (columna) {

            filaNueva[
              columna - 1
            ] =
              valoresPorCabecera[
                cabecera
              ];

          }

        }
      );


      hoja.appendRow(
        filaNueva
      );

    }


    SpreadsheetApp.flush();


    return {

      ok: true,

      idForo:
        idForo,

      actualizada:
        filaExistente !== -1,

      creada:
        filaExistente === -1,

      fecha:
        ahora.toISOString()

    };


  } catch (error) {

    return {

      ok: false,

      mensaje:
        error.message

    };


  } finally {

    try {

      lock.releaseLock();

    } catch (e) {
      // No hacer nada.
    }

  }

}


/*****************************************************
 * PRUEBA MANUAL DE AVANCESFORO
 *
 * NO genera un ID.
 * Utiliza el ID recibido.
 *
 * Esta función sirve para comprobar:
 * - creación;
 * - actualización;
 * - una sola fila.
 *****************************************************/

function probarGuardarAvanceForo() {

  const idForo =
    "PRUEBA-1234";


  const datos = {

    idForo:
      idForo,

    institucion:
      "IE PRUEBA 1234",

    dane:
      "111",

    fechaInicio:
      new Date(),

    campos: {

      respuestaSesion1:
        "XXXX",

      respuestaSesion1Pregunta2:
        "XXXX",

      respuestaSesion2Pregunta1:
        "XXXX",

      respuestaSesion2Pregunta2Accion1:
        "XXXX",

      respuestaSesion2Pregunta2Accion2:
        "XXXX",

      respuestaSesion2Pregunta2Accion3:
        "XXXX",

      respuestaSesion2Pregunta2Accion4:
        "XXXX",

      respuestaSesion2Pregunta2Accion5:
        "XXXX",

      respuestaSesion2Pregunta3:
        "XXXX",

      respuestaSesion2Pregunta4:
        "XXXX",

      respuestaSesion2Pregunta5:
        "XXXX"

    }

  };


  const resultado =
    guardarAvanceForo(
      datos
    );


  Logger.log(
    JSON.stringify(
      resultado,
      null,
      2
    )
  );


  return resultado;

}
/*****************************************************
 * CREAR ACCESO DE PRUEBA
 *
 * NO modifica la hoja "Oficiales".
 * NO modifica las 37 IE oficiales.
 *
 * Crea:
 *   IE PRUEBA 1234
 *   Código: 1234
 *   ID_FORO: PRUEBA-1234
 *****************************************************/

function crearAccesoPrueba1234() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const NOMBRE_HOJA = "AccesosIE";

  let hoja = ss.getSheetByName(NOMBRE_HOJA);

  /*
   * Crear la hoja si todavía no existe.
   */
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA);
  }

  /*
   * Estructura maestra de AccesosIE.
   *
   * Los campos de prueba adicionales EMAIL_IE y
   * EMAIL_RESPONSABLE se utilizarán posteriormente
   * para el flujo de correo.
   */
  const cabeceras = [
    "IE",
    "DANE",
    "CODIGO_ACCESO",
    "TOKEN",
    "URL_ACCESO",
    "ID_FORO",
    "ESTADO",
    "TOKEN_SESION",
    "DISPOSITIVO_ID",
    "FECHA_GENERACION",
    "FECHA_PRIMER_ACCESO",
    "EMAIL_IE",
    "EMAIL_RESPONSABLE",
    "TIPO"
  ];

  /*
   * Crear encabezados únicamente si la hoja está vacía.
   */
  if (hoja.getLastRow() === 0) {

    hoja
      .getRange(
        1,
        1,
        1,
        cabeceras.length
      )
      .setValues([cabeceras]);

  }


  /*
   * Verificar que la fila de prueba no exista.
   */
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila >= 2) {

    const datos =
      hoja
        .getRange(
          2,
          1,
          ultimaFila - 1,
          cabeceras.length
        )
        .getValues();

    for (let i = 0; i < datos.length; i++) {

      const ie =
        String(datos[i][0] || "").trim();

      const codigo =
        String(datos[i][2] || "").trim();

      if (
        ie === "IE PRUEBA 1234" ||
        codigo === "1234"
      ) {

        Logger.log(
          "La IE de prueba ya existe."
        );

        return {
          ok: true,
          existente: true,
          fila: i + 2
        };

      }

    }

  }


  /*
   * Token interno.
   *
   * NO es el código que se entrega al usuario.
   */
  const token =
    Utilities.getUuid();


  /*
   * ID_FORO reservado.
   *
   * Este será el mismo durante toda la prueba.
   */
  const idForo =
    "PRUEBA-1234";


  /*
   * URL de prueba.
   *
   * El token se recibe mediante ?t=TOKEN y
   * la pantalla de acceso lo valida en backend.
   */
  const urlAcceso =
    ScriptApp.getService().getUrl() +
    "?t=" +
    encodeURIComponent(token);


  const ahora =
    new Date();


  const fila = [

    "IE PRUEBA 1234",

    "111",

    "1234",

    token,

    urlAcceso,

    idForo,

    "DISPONIBLE",

    "",

    "",

    ahora,

    "",

    "jhonefrainsanchez@gmail.com",

    "hablaconhelprofe@gmail.com",

    "PRUEBA"

  ];


  hoja
    .getRange(
      hoja.getLastRow() + 1,
      1,
      1,
      fila.length
    )
    .setValues([fila]);


  SpreadsheetApp.flush();


  Logger.log(
    "===================================="
  );

  Logger.log(
    "ACCESO DE PRUEBA CREADO"
  );

  Logger.log(
    "IE: IE PRUEBA 1234"
  );

  Logger.log(
    "Código: 1234"
  );

  Logger.log(
    "DANE: 111"
  );

  Logger.log(
    "ID_FORO: " + idForo
  );

  Logger.log(
    "TOKEN: " + token
  );

  Logger.log(
    "URL: " + urlAcceso
  );

  Logger.log(
    "===================================="
  );


  return {

    ok: true,

    existente: false,

    ie: "IE PRUEBA 1234",

    dane: "111",

    codigo:
      "1234",

    token:
      token,

    idForo:
      idForo,

    url:
      urlAcceso

  };

}
/*****************************************************
 * VALIDAR ACCESO DE UNA IE
 *
 * ETAPA:
 * Validación backend.
 *
 * Recibe:
 *   token
 *   codigo
 *
 * Devuelve, si es válido:
 *   IE
 *   DANE
 *   ID_FORO
 *   TOKEN
 *   ESTADO
 *
 * IMPORTANTE:
 * La validación ocurre en el servidor.
 * El navegador nunca decide el ID_FORO.
 *****************************************************/


/*****************************************************
 * SESIÓN EXCLUSIVA POR CÓDIGO
 *
 * NO modifica AccesosIE.
 *
 * LockService se usa como exclusión mutua para que
 * dos dispositivos no puedan reclamar simultáneamente
 * el mismo código.
 *****************************************************/

const SESION_CODIGO_TTL_MS = 2 * 60 * 1000;

function obtenerClaveSesionCodigo_(token, codigo, idForo) {
  const claveBase = String(idForo || "").trim() || (String(token || "").trim()+"|"+String(codigo || "").trim());
  return "FEM_SESION_FORO_" + Utilities.base64EncodeWebSafe(claveBase);
}

function reclamarSesionCodigo_(token, codigo, dispositivoId, idForo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const props = PropertiesService.getScriptProperties();
    const clave = obtenerClaveSesionCodigo_(token,codigo,idForo);
    const ahora = Date.now();
    let actual = null;
    const guardado = props.getProperty(clave);
    if (guardado) { try { actual=JSON.parse(guardado); } catch(e){ actual=null; } }
    if (actual && actual.deviceId && actual.deviceId !== dispositivoId && Number(actual.ultimaActividad||0) > ahora-SESION_CODIGO_TTL_MS) {
      return {ok:false,codigo:"SESION_YA_ABIERTA",mensaje:"Este código está siendo utilizado en otro dispositivo. Si acaba de cambiar de equipo, espere 3 minutos mientras se libera la sesión anterior o solicite otro código de su institución al email de la SEM o al WhatsApp 3184561081."};
    }
    const tokenSesion = actual && actual.deviceId===dispositivoId && actual.tokenSesion ? actual.tokenSesion : Utilities.getUuid();
    props.setProperty(clave,JSON.stringify({deviceId:dispositivoId,tokenSesion:tokenSesion,ultimaActividad:ahora,idForo:idForo}));
    return {ok:true,tokenSesion:tokenSesion};
  } catch(e) { return {ok:false,codigo:"LOCK_SESION_ERROR",mensaje:"No fue posible asegurar la sesión de acceso. Intente nuevamente."}; }
  finally { try{lock.releaseLock();}catch(e){} }
}

function mantenerSesionCodigo(token,codigo,dispositivoId,tokenSesion,idForo){
  return mantenerSesionCodigo_(token,codigo,dispositivoId,tokenSesion,idForo);
}
function mantenerSesionCodigo_(token,codigo,dispositivoId,tokenSesion,idForo){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const props=PropertiesService.getScriptProperties();
    const clave=obtenerClaveSesionCodigo_(token,codigo,idForo);
    const guardado=props.getProperty(clave); if(!guardado)return {ok:false,codigo:"SESION_NO_ENCONTRADA"};
    const actual=JSON.parse(guardado);
    if(actual.deviceId!==dispositivoId || actual.tokenSesion!==tokenSesion)return {ok:false,codigo:"SESION_NO_AUTORIZADA"};
    actual.ultimaActividad=Date.now(); props.setProperty(clave,JSON.stringify(actual)); return {ok:true};
  }catch(e){return {ok:false,codigo:"HEARTBEAT_ERROR"};} finally{try{lock.releaseLock();}catch(e){}}
}

function liberarSesionCodigo(token,codigo,dispositivoId,tokenSesion,idForo){return liberarSesionCodigo_(token,codigo,dispositivoId,tokenSesion,idForo);}
function liberarSesionCodigo_(token,codigo,dispositivoId,tokenSesion,idForo){
  const lock=LockService.getScriptLock();
  try{lock.waitLock(10000);const props=PropertiesService.getScriptProperties();const clave=obtenerClaveSesionCodigo_(token,codigo,idForo);const guardado=props.getProperty(clave);if(!guardado)return {ok:true};const actual=JSON.parse(guardado);if(actual.deviceId===dispositivoId&&actual.tokenSesion===tokenSesion)props.deleteProperty(clave);return {ok:true};}
  catch(e){return {ok:false};} finally{try{lock.releaseLock();}catch(e){}}
}

function validarAccesoIE(token, codigo, dispositivoId) {

  try {

    token =
      String(token || "").trim();

    codigo =
      String(codigo || "").trim();

    dispositivoId =
      String(dispositivoId || "").trim();


    /*
     * Validaciones básicas.
     */

    if (token === "") {

      return {
        ok: false,
        codigo: "TOKEN_REQUERIDO",
        mensaje: "No se recibió el token de acceso."
      };

    }


    if (codigo === "") {

      return {
        ok: false,
        codigo: "CODIGO_REQUERIDO",
        mensaje: "Debe ingresar el código de acceso."
      };

    }


    if (dispositivoId === "") {

      return {
        ok: false,
        codigo: "DISPOSITIVO_REQUERIDO",
        mensaje: "No fue posible identificar este dispositivo."
      };

    }


    /*
     * Abrir Spreadsheet.
     */

    const ss =
      SpreadsheetApp.openById(
        SPREADSHEET_ID
      );


    const hoja =
      ss.getSheetByName(
        "AccesosIE"
      );


    if (!hoja) {

      return {
        ok: false,
        codigo: "HOJA_NO_EXISTE",
        mensaje:
          'No existe la hoja "AccesosIE".'
      };

    }


    /*
     * Obtener encabezados.
     */

    const ultimaColumna =
      hoja.getLastColumn();


    if (ultimaColumna < 1) {

      return {
        ok: false,
        codigo: "ACCESOS_SIN_DATOS",
        mensaje:
          "La hoja AccesosIE no contiene datos."
      };

    }


    const cabeceras =
      hoja
        .getRange(
          1,
          1,
          1,
          ultimaColumna
        )
        .getDisplayValues()[0];


    /*
     * Crear mapa de columnas.
     */

    const mapa = {};


    cabeceras.forEach(
      function(nombre, indice) {

        const clave =
          String(
            nombre || ""
          ).trim();


        if (clave !== "") {

          mapa[clave] =
            indice + 1;

        }

      }
    );


    /*
     * Columnas mínimas necesarias.
     */

    const requeridas = [
      "IE",
      "DANE",
      "CODIGO_ACCESO",
      "TOKEN",
      "ID_FORO",
      "ESTADO",
      "TIPO"
    ];


    for (
      let i = 0;
      i < requeridas.length;
      i++
    ) {

      const campo =
        requeridas[i];


      if (!mapa[campo]) {

        return {
          ok: false,
          codigo: "COLUMNA_FALTANTE",
          mensaje:
            "Falta la columna " +
            campo +
            " en AccesosIE."
        };

      }

    }


    /*
     * Leer registros.
     */

    const ultimaFila =
      hoja.getLastRow();


    if (ultimaFila < 2) {

      return {
        ok: false,
        codigo: "SIN_ACCESOS",
        mensaje:
          "No existen accesos registrados."
      };

    }


    const filas =
      hoja
        .getRange(
          2,
          1,
          ultimaFila - 1,
          ultimaColumna
        )
        .getDisplayValues();


    /*
     * Buscar el TOKEN.
     */

    let registro = null;
    let numeroFila = -1;


    for (
      let i = 0;
      i < filas.length;
      i++
    ) {

      const tokenFila =
        String(
          filas[i][
            mapa["TOKEN"] - 1
          ] || ""
        ).trim();


      if (
        tokenFila === token
      ) {

        registro =
          filas[i];

        numeroFila =
          i + 2;

        break;

      }

    }


    /*
     * TOKEN inexistente.
     */

    if (!registro) {

      return {
        ok: false,
        codigo: "TOKEN_INVALIDO",
        mensaje:
          "El enlace de acceso no es válido."
      };

    }


    /*
     * Obtener datos del registro.
     */

    const ie =
      String(
        registro[
          mapa["IE"] - 1
        ] || ""
      ).trim();


    const dane =
      String(
        registro[
          mapa["DANE"] - 1
        ] || ""
      ).trim();


    /*
     * CÓDIGOS VÁLIDOS:
     * - Código principal
     * - 4 códigos de contingencia
     *
     * Todos mantienen el mismo TOKEN e ID_FORO.
     */

    const codigosValidos = [

      "CODIGO_ACCESO",
      "CODIGO_CONTINGENCIA_1",
      "CODIGO_CONTINGENCIA_2",
      "CODIGO_CONTINGENCIA_3",
      "CODIGO_CONTINGENCIA_4"

    ].map(function(nombreColumna){

      if (!mapa[nombreColumna]) {
        return "";
      }

      return String(
        registro[
          mapa[nombreColumna] - 1
        ] || ""
      ).trim();

    }).filter(function(valor){

      return valor !== "";

    });


    const idForo =
      String(
        registro[
          mapa["ID_FORO"] - 1
        ] || ""
      ).trim();


    const estado =
      String(
        registro[
          mapa["ESTADO"] - 1
        ] || ""
      ).trim()
      .toUpperCase();


    const tipo =
      String(
        registro[
          mapa["TIPO"] - 1
        ] || ""
      ).trim()
      .toUpperCase();


    let datosIE = null;


    /*
     * Comprobar código.
     */

    if (
      codigosValidos.indexOf(codigo) === -1
    ) {

      return {
        ok: false,
        codigo: "CODIGO_INCORRECTO",
        mensaje:
                "El código es incorrecto. Verifique el código que se envió a su I.E. Si el error persiste, comuníquese con la Secretaría de Educación de Neiva al correo educacion@alcaldianeiva.gov.co o al WhatsApp 318 456 1081."
      };

    }


    /*
     * Comprobar ID_FORO.
     */

    if (idForo === "") {

      return {
        ok: false,
        codigo: "ID_FORO_FALTANTE",
        mensaje:
          "El acceso no tiene un ID_FORO asignado."
      };

    }


    /*
     * Comprobar estado.
     *
     * En esta etapa:
     *
     * DISPONIBLE = permitido
     *
     * ENVIADO = bloqueado
     * BLOQUEADO = bloqueado
     * INACTIVO = bloqueado
     */

    if (
      estado === "ENVIADO" ||
      estado === "BLOQUEADO" ||
      estado === "INACTIVO"
    ) {

      return {
        ok: false,
        codigo: "ACCESO_BLOQUEADO",
        mensaje:
          "Este acceso ya no está disponible."
      };

    }


    /*
     * Verificación de IE.
     *
     * Las IE oficiales deben existir en Oficiales.
     *
     * EXCEPCIÓN CONTROLADA:
     * TIPO = PRUEBA.
     *
     * Esto permite utilizar IE PRUEBA 1234
     * sin contaminar la lista oficial de 37 IE.
     */

    if (
      tipo !== "PRUEBA"
    ) {

      const instituciones =
        obtenerInstituciones();

      const institucionOficial =
        buscarInstitucionOficial_(
          instituciones,
          ie
        );


      if(!institucionOficial){

        return {
          ok: false,
          codigo: "IE_NO_VALIDA",
          mensaje:
            "La institución asociada al acceso no es válida."
        };

      }

      datosIE =
        institucionOficial.datos;

    } else {

      /*
       * La prueba debe conservar
       * explícitamente su identidad.
       */

      if (
  ie !== "IE PRUEBA 1234" ||
  dane !== "111" ||
  codigo !== "1234" ||
  idForo !== "PRUEBA-1234"
) {

        return {
          ok: false,
          codigo: "PRUEBA_INVALIDA",
          mensaje:
            "El registro de prueba no tiene la configuración esperada."
        };

      }

      datosIE = {
        dane: dane,
        rector: "",
        direccion: "",
        zona: "",
        comuna: "",
        correo: ""
      };

    }


    /*
     * =================================================
     * LOCK DE SESIÓN POR CÓDIGO + DISPOSITIVO
     * =================================================
     *
     * LockService garantiza que dos dispositivos que
     * intenten reclamar el mismo código al mismo tiempo
     * no puedan hacerlo simultáneamente.
     *
     * La ocupación persistente se guarda en ScriptProperties.
     * Cada código tiene su propio lock lógico.
     *
     * La sesión se libera después de 2 minutos sin heartbeat.
     * El navegador la mantiene viva mientras el formulario siga abierto.
     */ 

    const resultadoSesion =
      reclamarSesionCodigo_(token, codigo, dispositivoId, idForo);

    if (!resultadoSesion.ok) {

      return resultadoSesion;

    }


    /*
     * ACCESO VÁLIDO.
     *
     * El ID_FORO sale exclusivamente
     * de AccesosIE.
     */

    return {

      ok: true,

      codigo: "ACCESO_VALIDO",

      mensaje:
        "Acceso autorizado.",

      fila:
        numeroFila,

      ie:
        ie,

      dane:
        dane,

      codigoAcceso:
        codigo,

      token:
        token,

      idForo:
        idForo,

      estado:
        estado,

      tipo:
        tipo,

      tokenSesion:
        resultadoSesion.tokenSesion,

      // Se devuelve el mismo identificador que fue validado para que
      // el frontend pueda conservar la sesión de forma inequívoca.
      dispositivoId:
        dispositivoId,

      datosIE:
        datosIE || {
          dane: dane,
          rector: "",
          direccion: "",
          zona: "",
          comuna: "",
          correo: ""
        },

      sesiones:
        obtenerEstadoSesiones_(idForo)

    };


  } catch (error) {

    return {

      ok: false,

      codigo: "ERROR_SERVIDOR",

      mensaje:
        error.message

    };

  }

}
/*****************************************************
 * PRUEBA DE VALIDACIÓN DEL ACCESO 1234
 *****************************************************/

function probarValidacion1234() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const hoja =
    ss.getSheetByName("AccesosIE");

  if (!hoja) {
    throw new Error(
      'No existe la hoja "AccesosIE".'
    );
  }

  const datos =
    hoja.getDataRange().getValues();

  const cabeceras =
    datos[0];

  const colIE =
    cabeceras.indexOf("IE");

  const colToken =
    cabeceras.indexOf("TOKEN");

  if (
    colIE === -1 ||
    colToken === -1
  ) {
    throw new Error(
      "No se encontraron las columnas IE o TOKEN."
    );
  }

  let tokenPrueba = "";

  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    if (
      String(
        datos[i][colIE] || ""
      ).trim() ===
      "IE PRUEBA 1234"
    ) {

      tokenPrueba =
        String(
          datos[i][colToken] || ""
        ).trim();

      break;
    }
  }

  if (tokenPrueba === "") {

    throw new Error(
      "No se encontró IE PRUEBA 1234 en AccesosIE."
    );
  }

  const resultado =
    validarAccesoIE(
      tokenPrueba,
      "1234"
    );

  Logger.log(
    JSON.stringify(
      resultado,
      null,
      2
    )
  );

  return resultado;
}
/*****************************************************
 * VERIFICAR CUENTA DE ENVÍO
 *****************************************************/

function verificarCuentaEnvio() {

  const cuenta =
    Session.getEffectiveUser().getEmail();

  const aliases =
    GmailApp.getAliases();

  Logger.log(
    "Cuenta efectiva: " + cuenta
  );

  Logger.log(
    "Aliases disponibles: " +
    JSON.stringify(aliases)
  );

  const puedeUsarEducacion =
    cuenta.toLowerCase() ===
      "educacion@alcaldianeiva.gov.co"
    ||
    aliases
      .map(function(a){
        return a.toLowerCase();
      })
      .includes(
        "educacion@alcaldianeiva.gov.co"
      );

  Logger.log(
    "¿Puede enviar como educacion@alcaldianeiva.gov.co?: " +
    puedeUsarEducacion
  );

  return {
    cuenta: cuenta,
    aliases: aliases,
    puedeUsarEducacion:
      puedeUsarEducacion
  };

}
/*****************************************************
 * ENVIAR CORREO DE PRUEBA
 *
 * Utiliza los datos reales de IE PRUEBA 1234
 * almacenados en AccesosIE.
 *****************************************************/

function enviarCorreoPruebaIE1234() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const hoja =
    ss.getSheetByName("AccesosIE");

  if (!hoja) {
    throw new Error(
      'No existe la hoja "AccesosIE".'
    );
  }


  const datos =
    hoja.getDataRange().getValues();

  const cabeceras =
    datos[0];


  const mapa = {};

  cabeceras.forEach(
    function(nombre, indice) {

      const clave =
        String(nombre || "").trim();

      if (clave !== "") {
        mapa[clave] = indice;
      }

    }
  );


  let registro = null;


  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    const ie =
      String(
        datos[i][mapa["IE"]] || ""
      ).trim();

    const tipo =
      String(
        datos[i][mapa["TIPO"]] || ""
      ).trim()
      .toUpperCase();

    if (
      ie === "IE PRUEBA 1234" &&
      tipo === "PRUEBA"
    ) {

      registro = datos[i];
      break;

    }

  }


  if (!registro) {

    throw new Error(
      "No se encontró IE PRUEBA 1234."
    );

  }


  const ie =
    String(
      registro[mapa["IE"]] || ""
    ).trim();

  const codigo =
    String(
      registro[mapa["CODIGO_ACCESO"]] || ""
    ).trim();

  const url =
    String(
      registro[mapa["URL_ACCESO"]] || ""
    ).trim();

  const idForo =
    String(
      registro[mapa["ID_FORO"]] || ""
    ).trim();

  const correoIE =
    String(
      registro[mapa["EMAIL_IE"]] || ""
    ).trim();

  const correoResponsable =
    String(
      registro[mapa["EMAIL_RESPONSABLE"]] || ""
    ).trim();


  if (!correoIE) {

    throw new Error(
      "La IE de prueba no tiene correo registrado."
    );

  }


  /*
   * Verificar que la cuenta tenga autorización
   * para utilizar educacion@alcaldianeiva.gov.co.
   */

  const cuenta =
    Session.getEffectiveUser()
      .getEmail()
      .toLowerCase();

  const aliases =
    GmailApp
      .getAliases()
      .map(function(a){
        return a.toLowerCase();
      });


  const remitente =
    "educacion@alcaldianeiva.gov.co";


  const puedeEnviar =
    cuenta === remitente ||
    aliases.includes(remitente);


  if (!puedeEnviar) {

    throw new Error(
      "La cuenta que ejecuta el script (" +
      cuenta +
      ") no está autorizada para enviar como " +
      remitente +
      "."
    );

  }


  const asunto =
    "Acceso al Foro Educativo Institucional – IE PRUEBA 1234";


  const cuerpoTexto =
    "Secretaría de Educación de Neiva\n\n" +

    "Se informa que la Institución Educativa " +
    ie +
    " ha sido habilitada para participar en el Foro Educativo Institucional – Neiva 2026.\n\n" +

    "Código de acceso: " +
    codigo +
    "\n\n" +

    "Enlace personalizado:\n" +
    url +
    "\n\n" +

    "ID de prueba: " +
    idForo +
    "\n\n" +

    "Este correo corresponde únicamente a una prueba técnica del sistema.\n\n" +

    "Secretaría de Educación de Neiva";


  const cuerpoHTML =
    "<div style=\"font-family:Arial,sans-serif;line-height:1.6\">" +

    "<h2>Foro Educativo Institucional – Neiva 2026</h2>" +

    "<p>Se informa que la Institución Educativa " +
    "<strong>" + ie + "</strong>" +
    " ha sido habilitada para participar en el Foro Educativo Institucional – Neiva 2026.</p>" +

    "<p><strong>Código de acceso:</strong><br>" +
    "<span style=\"font-size:24px;font-weight:bold;letter-spacing:3px\">" +
    codigo +
    "</span></p>" +

    "<p><strong>Enlace personalizado:</strong><br>" +
    "<a href=\"" + url + "\">Ingresar al Foro Educativo</a></p>" +

    "<p style=\"font-size:12px;color:#666\">" +
    "Este correo corresponde únicamente a una prueba técnica del sistema." +
    "</p>" +

    "<p>Secretaría de Educación de Neiva</p>" +

    "</div>";


  const opciones = {

    htmlBody:
      cuerpoHTML,

    name:
      "Secretaría de Educación de Neiva",

    replyTo:
      remitente

  };


  if (
    correoResponsable &&
    correoResponsable !== correoIE
  ) {

    opciones.cc =
      correoResponsable;

  }


  /*
   * Si educacion@ es un alias autorizado,
   * se fuerza el remitente.
   */

  if (cuenta !== remitente) {

    opciones.from =
      remitente;

  }


  GmailApp.sendEmail(
    correoIE,
    asunto,
    cuerpoTexto,
    opciones
  );


  Logger.log(
    "Correo de prueba enviado."
  );

  Logger.log(
    "Destinatario: " +
    correoIE
  );

  Logger.log(
    "CC: " +
    (correoResponsable || "ninguno")
  );

  Logger.log(
    "Remitente: " +
    remitente
  );


  return {
    ok: true,
    ie: ie,
    destinatario: correoIE,
    cc: correoResponsable,
    remitente: remitente
  };

}
/*****************************************************
 * PROGRAMAR CORREO DE PRUEBA
 *
 * Lo programa aproximadamente 3 minutos después
 * de ejecutar esta función.
 *****************************************************/

function programarCorreoPruebaEn3Minutos() {

  const tiempo =
    new Date(
      Date.now() +
      3 * 60 * 1000
    );


  const trigger =
    ScriptApp
      .newTrigger(
        "enviarCorreoPruebaIE1234"
      )
      .timeBased()
      .at(tiempo)
      .create();


  Logger.log(
    "Correo de prueba programado."
  );

  Logger.log(
    "Hora aproximada: " +
    tiempo
  );

  Logger.log(
    "Trigger ID: " +
    trigger.getUniqueId()
  );


  return {
    ok: true,
    fechaProgramada:
      tiempo.toISOString(),
    triggerId:
      trigger.getUniqueId()
  };

}
/*****************************************************
 * PRUEBA DE CORREO - IE PRUEBA 1234
 *
 * IMPORTANTE:
 * Esta función es solamente para pruebas.
 *
 * Remitente:
 * cuenta que ejecuta Apps Script
 *
 * Destinatario:
 * hablaconhelprofe@gmail.com
 *
 * NO utiliza todavía:
 * educacion@alcaldianeiva.gov.co
 *****************************************************/

function enviarCorreoPruebaActual() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const hoja =
    ss.getSheetByName("AccesosIE");

  if (!hoja) {
    throw new Error(
      'No existe la hoja "AccesosIE".'
    );
  }


  const datos =
    hoja.getDataRange().getValues();

  const cabeceras =
    datos[0];

  const mapa = {};

  cabeceras.forEach(
    function(nombre, indice) {

      const clave =
        String(nombre || "").trim();

      if (clave !== "") {
        mapa[clave] = indice;
      }

    }
  );


  let registro = null;


  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    const ie =
      String(
        datos[i][mapa["IE"]] || ""
      ).trim();

    if (
      ie === "IE PRUEBA 1234"
    ) {

      registro =
        datos[i];

      break;

    }

  }


  if (!registro) {

    throw new Error(
      "No se encontró IE PRUEBA 1234."
    );

  }


  const ie =
    String(
      registro[mapa["IE"]] || ""
    ).trim();

  const codigo =
    String(
      registro[mapa["CODIGO_ACCESO"]] || ""
    ).trim();

  const url =
    String(
      registro[mapa["URL_ACCESO"]] || ""
    ).trim();

  const idForo =
    String(
      registro[mapa["ID_FORO"]] || ""
    ).trim();


  /*
   * DESTINATARIO DE PRUEBA
   */
  const destinatario =
    "hablaconhelprofe@gmail.com";


  const asunto =
    "PRUEBA – Acceso Foro Educativo Institucional – " +
    ie;


  const cuerpoTexto =

    "PRUEBA TÉCNICA\n\n" +

    "Foro Educativo Institucional – Neiva 2026\n\n" +

    "Institución Educativa:\n" +
    ie +
    "\n\n" +

    "Código de acceso:\n" +
    codigo +
    "\n\n" +

    "Enlace personalizado:\n" +
    url +
    "\n\n" +

    "ID_FORO de prueba:\n" +
    idForo +
    "\n\n" +

    "Este mensaje corresponde a una prueba técnica " +
    "del sistema de envío de accesos.\n\n" +

    "Secretaría de Educación de Neiva";


  const cuerpoHTML =

    "<div style=\"" +
    "font-family:Arial,sans-serif;" +
    "line-height:1.6;" +
    "max-width:700px;" +
    "margin:auto;" +
    "\">" +

    "<h2>" +
    "Foro Educativo Institucional – Neiva 2026" +
    "</h2>" +

    "<p>" +
    "<strong>PRUEBA TÉCNICA DEL SISTEMA</strong>" +
    "</p>" +

    "<p>" +
    "Institución Educativa:<br>" +
    "<strong>" +
    ie +
    "</strong>" +
    "</p>" +

    "<p>" +
    "Código de acceso:<br>" +

    "<span style=\"" +
    "font-size:26px;" +
    "font-weight:bold;" +
    "letter-spacing:4px;" +
    "\">" +

    codigo +

    "</span>" +

    "</p>" +

    "<p>" +
    "Enlace personalizado:<br>" +

    "<a href=\"" +
    url +
    "\" target=\"_blank\">" +

    "Ingresar al Foro Educativo" +

    "</a>" +

    "</p>" +

    "<p>" +
    "<strong>ID_FORO:</strong> " +
    idForo +
    "</p>" +

    "<hr>" +

    "<p style=\"color:#666;font-size:13px\">" +
    "Este mensaje corresponde únicamente a " +
    "una prueba técnica del sistema de envío " +
    "de accesos." +
    "</p>" +

    "<p>" +
    "Secretaría de Educación de Neiva" +
    "</p>" +

    "</div>";


  /*
   * IMPORTANTE:
   *
   * No especificamos "from".
   * Gmail utilizará la cuenta que ejecuta
   * actualmente el Apps Script.
   */

  GmailApp.sendEmail(
    destinatario,
    asunto,
    cuerpoTexto,
    {
      htmlBody: cuerpoHTML,
      name:
        "Foro Educativo Institucional – Neiva 2026"
    }
  );


  Logger.log(
    "===================================="
  );

  Logger.log(
    "CORREO DE PRUEBA ENVIADO"
  );

  Logger.log(
    "Remitente: " +
    Session.getEffectiveUser().getEmail()
  );

  Logger.log(
    "Destinatario: " +
    destinatario
  );

  Logger.log(
    "IE: " +
    ie
  );

  Logger.log(
    "Código: " +
    codigo
  );

  Logger.log(
    "ID_FORO: " +
    idForo
  );

  Logger.log(
    "===================================="
  );


  return {
    ok: true,
    remitente:
      Session.getEffectiveUser().getEmail(),
    destinatario:
      destinatario,
    ie:
      ie,
    codigo:
      codigo,
    idForo:
      idForo
  };

}


/*****************************************************
 * ENVÍO DE ACCESOS A TODAS LAS IE
 *
 * IMPORTANTE:
 * - Solo envía registros con ESTADO = DISPONIBLE.
 * - Requiere EMAIL_IE.
 * - No cambia códigos, tokens ni ID_FORO.
 * - No se ejecuta automáticamente al guardar el archivo.
 *****************************************************/

function enviarAccesosTodasIE(){

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const hoja =
    ss.getSheetByName(
      "AccesosIE"
    );

  if(!hoja){

    throw new Error(
      'No existe la hoja "AccesosIE".'
    );

  }

  const datos =
    hoja.getDataRange().getDisplayValues();

  if(datos.length < 2){

    return {
      ok: true,
      enviados: 0,
      omitidos: 0,
      errores: []
    };

  }

  const cabeceras =
    datos[0];

  const mapa = {};

  cabeceras.forEach(
    function(nombre, indice){

      const clave =
        String(
          nombre || ""
        ).trim();

      if(clave){
        mapa[clave] = indice;
      }

    }
  );

  const requeridas = [
    "IE",
    "CODIGO_ACCESO",
    "TOKEN",
    "URL_ACCESO",
    "ID_FORO",
    "ESTADO",
    "EMAIL_IE"
  ];

  requeridas.forEach(
    function(campo){

      if(mapa[campo] === undefined){

        throw new Error(
          "Falta la columna " +
          campo +
          " en AccesosIE."
        );

      }

    }
  );

  const remitente =
    "educacion@alcaldianeiva.gov.co";

  const cuenta =
    Session.getEffectiveUser()
      .getEmail()
      .toLowerCase();

  const aliases =
    GmailApp.getAliases()
      .map(function(alias){
        return alias.toLowerCase();
      });

  if(
    cuenta !== remitente &&
    !aliases.includes(remitente)
  ){

    throw new Error(
      "La cuenta que ejecuta Apps Script no puede enviar como " +
      remitente +
      "."
    );

  }

  let enviados = 0;
  let omitidos = 0;
  const errores = [];

  for(
    let i = 1;
    i < datos.length;
    i++
  ){

    const fila =
      datos[i];

    const estado =
      String(
        fila[mapa["ESTADO"]] || ""
      ).trim().toUpperCase();

    if(estado !== "DISPONIBLE"){

      omitidos++;
      continue;

    }

    const correoIE =
      String(
        fila[mapa["EMAIL_IE"]] || ""
      ).trim();

    if(!correoIE){

      omitidos++;
      continue;

    }

    const ie =
      String(
        fila[mapa["IE"]] || ""
      ).trim();

    const codigo =
      String(
        fila[mapa["CODIGO_ACCESO"]] || ""
      ).trim();

    const url =
      String(
        fila[mapa["URL_ACCESO"]] || ""
      ).trim();

    const idForo =
      String(
        fila[mapa["ID_FORO"]] || ""
      ).trim();

    const correoResponsable =
      mapa["EMAIL_RESPONSABLE"] !== undefined
        ? String(
            fila[mapa["EMAIL_RESPONSABLE"]] || ""
          ).trim()
        : "";

    try{

      const asunto =
        "Acceso al Foro Educativo Institucional – " +
        ie;

      const cuerpoTexto =
        "Secretaría de Educación de Neiva\n\n" +
        "La Institución Educativa " +
        ie +
        " ha sido habilitada para participar en el Foro Educativo Institucional – Neiva 2026.\n\n" +
        "Código de acceso: " +
        codigo +
        "\n\n" +
        "Enlace personalizado:\n" +
        url +
        "\n\n" +
        "Este código y este enlace corresponden exclusivamente a esta institución educativa. " +
        "No comparta el acceso con otra IE.\n\n" +
        "Secretaría de Educación de Neiva";

      const cuerpoHTML =
        "<div style=\"font-family:Arial,sans-serif;line-height:1.6;max-width:700px;margin:auto\">" +
        "<h2>Foro Educativo Institucional – Neiva 2026</h2>" +
        "<p>La Institución Educativa <strong>" +
        ie +
        "</strong> ha sido habilitada para participar en el Foro Educativo Institucional.</p>" +
        "<p><strong>Código de acceso:</strong><br>" +
        "<span style=\"font-size:26px;font-weight:bold;letter-spacing:4px\">" +
        codigo +
        "</span></p>" +
        "<p><strong>Enlace personalizado:</strong><br>" +
        "<a href=\"" +
        url +
        "\" target=\"_blank\">Ingresar al Foro Educativo</a></p>" +
        "<p style=\"color:#666;font-size:13px\">" +
        "Este código y este enlace corresponden exclusivamente a esta institución educativa. " +
        "No comparta el acceso con otra IE." +
        "</p>" +
        "<p>Secretaría de Educación de Neiva</p>" +
        "</div>";

      const opciones = {
        htmlBody: cuerpoHTML,
        name: "Secretaría de Educación de Neiva",
        replyTo: remitente
      };

      if(
        cuenta !== remitente
      ){
        opciones.from = remitente;
      }

      if(
        correoResponsable &&
        correoResponsable !== correoIE
      ){
        opciones.cc =
          correoResponsable;
      }

      GmailApp.sendEmail(
        correoIE,
        asunto,
        cuerpoTexto,
        opciones
      );

      enviados++;

    }catch(error){

      errores.push({
        fila: i + 1,
        ie: ie,
        correo: correoIE,
        mensaje: error.message
      });

    }

  }

  return {
    ok: errores.length === 0,
    enviados: enviados,
    omitidos: omitidos,
    errores: errores
  };

}


/*****************************************************
 * PROGRAMAR ENVÍO DE ACCESOS
 *
 * Fecha solicitada:
 * 28 de agosto de 2026 a las 5:45 a. m.
 *
 * NO se ejecuta al guardar el código.
 * Debe ejecutarse manualmente una sola vez cuando
 * la cuenta de envío ya tenga la autorización.
 *****************************************************/

function programarEnvioAccesos28Agosto(){

  const zona =
    Session.getScriptTimeZone();

  /*
   * El proyecto de Apps Script debe estar configurado
   * con la zona horaria correspondiente a Neiva/Colombia
   * (America/Bogota) para que 5:45 a. m. sea correcto.
   */
  if(
    zona !== "America/Bogota"
  ){

    throw new Error(
      "Antes de programar el envío, configure la zona horaria del proyecto en America/Bogota. " +
      "Zona actual: " +
      zona
    );

  }

  const fecha =
    new Date(
      "2026-08-28T05:45:00-05:00"
    );

  const ahora =
    new Date();

  if(
    fecha <= ahora
  ){

    throw new Error(
      "La fecha programada ya pasó."
    );

  }

  /*
   * Evitar duplicar triggers de este envío.
   */
  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(
    function(trigger){

      if(
        trigger.getHandlerFunction() ===
        "enviarAccesosTodasIE"
      ){

        ScriptApp.deleteTrigger(
          trigger
        );

      }

    }
  );

  const trigger =
    ScriptApp
      .newTrigger(
        "enviarAccesosTodasIE"
      )
      .timeBased()
      .at(fecha)
      .create();

  return {
    ok: true,
    fechaProgramada:
      fecha.toISOString(),
    zonaHoraria:
      zona,
    triggerId:
      trigger.getUniqueId()
  };

}


/*****************************************************
 * PRUEBA DE VALIDACIÓN - CÓDIGO INCORRECTO
 *
 * Esta función NO modifica AccesosIE.
 * Solo comprueba que un código incorrecto
 * sea rechazado por validarAccesoIE().
 *****************************************************/

function probarCodigoIncorrecto1234() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const hoja =
    ss.getSheetByName(
      "AccesosIE"
    );

  if (!hoja) {

    throw new Error(
      'No existe la hoja "AccesosIE".'
    );

  }

  const datos =
    hoja
      .getDataRange()
      .getValues();

  const cabeceras =
    datos[0];

  const colIE =
    cabeceras.indexOf("IE");

  const colToken =
    cabeceras.indexOf("TOKEN");

  if (
    colIE === -1 ||
    colToken === -1
  ) {

    throw new Error(
      "No se encontraron las columnas IE o TOKEN."
    );

  }

  let tokenPrueba = "";

  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    if (
      String(
        datos[i][colIE] || ""
      ).trim() ===
      "IE PRUEBA 1234"
    ) {

      tokenPrueba =
        String(
          datos[i][colToken] || ""
        ).trim();

      break;

    }

  }

  if (
    tokenPrueba === ""
  ) {

    throw new Error(
      "No se encontró IE PRUEBA 1234."
    );

  }

  /*
   * Deliberadamente utilizamos 9999.
   * El código real es 1234.
   */

  const resultado =
    validarAccesoIE(
      tokenPrueba,
      "9999"
    );

  Logger.log(
    JSON.stringify(
      resultado,
      null,
      2
    )
  );

  return resultado;
}
/*****************************************************
 * ACTUALIZAR URL DE ACCESO DE LAS IE
 *
 * NO genera nuevos TOKEN.
 * NO genera nuevos códigos.
 * NO modifica ID_FORO.
 *
 * Únicamente reconstruye URL_ACCESO para los
 * registros que ya existen en AccesosIE.
 *****************************************************/

function actualizarURLsAccesoIE(){

    const ss =
        SpreadsheetApp.openById(
            SPREADSHEET_ID
        );

    const hoja =
        ss.getSheetByName(
            "AccesosIE"
        );

    if(!hoja){

        throw new Error(
            'No existe la hoja "AccesosIE".'
        );

    }


    const datos =
        hoja
            .getDataRange()
            .getValues();


    if(datos.length < 2){

        throw new Error(
            "No existen registros en AccesosIE."
        );

    }


    const cabeceras =
        datos[0].map(function(valor){

            return String(
                valor || ""
            ).trim();

        });


    const colIE =
        cabeceras.indexOf("IE");

    const colToken =
        cabeceras.indexOf("TOKEN");

    const colURL =
        cabeceras.indexOf("URL_ACCESO");


    if(colIE === -1){

        throw new Error(
            "Falta la columna IE en AccesosIE."
        );

    }


    if(colToken === -1){

        throw new Error(
            "Falta la columna TOKEN en AccesosIE."
        );

    }


    if(colURL === -1){

        throw new Error(
            "Falta la columna URL_ACCESO en AccesosIE."
        );

    }


    /*
     * URL base del Web App.
     *
     * Se utiliza la URL de la implementación
     * actual del proyecto.
     */

    const urlBase =
        ScriptApp.getService().getUrl();


    if(!urlBase){

        throw new Error(
            "No fue posible obtener la URL del Web App."
        );

    }


    /*
     * Actualizar cada IE.
     */

    for(
        let i = 1;
        i < datos.length;
        i++
    ){

        const ie =
            String(
                datos[i][colIE] || ""
            ).trim();


        const token =
            String(
                datos[i][colToken] || ""
            ).trim();


        /*
         * Ignorar filas incompletas.
         */

        if(
            ie === "" ||
            token === ""
        ){

            continue;

        }


        /*
         * Crear nombre legible para la URL.
         *
         * Ejemplo:
         *
         * Institución Educativa Liceo de Santa Librada
         *
         * ↓
         *
         * IE-Institución-Educativa-Liceo-de-Santa-Librada
         *
         * Si ya comienza por IE, no duplicarlo.
         */

        let nombreURL =
            ie
                .replace(/^IE[\s-]+/i, "")
                .trim();


        nombreURL =
            "IE-" +
            nombreURL;


        /*
         * Limpiar caracteres problemáticos
         * sin eliminar las palabras del nombre.
         */

        nombreURL =
            nombreURL
                .replace(/\s+/g, "-")
                .replace(/[.,;:()[\]{}¿?¡!]/g, "")
                .replace(/["']/g, "")
                .replace(/&/g, "y")
                .replace(/\/+/g, "-")
                .replace(/-+/g, "-");


        /*
         * Construir URL.
         */

        const urlAcceso =
            urlBase +
            "?ie=" +
            encodeURIComponent(
                nombreURL
            ) +
            "&t=" +
            encodeURIComponent(
                token
            );


        /*
         * Escribir únicamente URL_ACCESO.
         */

        hoja
            .getRange(
                i + 1,
                colURL + 1
            )
            .setValue(
                urlAcceso
            );

    }


    SpreadsheetApp.flush();


    return {
        ok: true,
        mensaje:
            "Las URL_ACCESO fueron actualizadas correctamente."
    };

}
function probarCatalogoIE() {

    const json =
        obtenerInstitucionesJSON();

    const instituciones =
        JSON.parse(json || "{}");

    const nombres =
        Object.keys(instituciones);

    Logger.log(
        "========================================"
    );

    Logger.log(
        "TOTAL DE IE DISPONIBLES: " +
        nombres.length
    );

    Logger.log(
        "========================================"
    );

    nombres.forEach(function(nombre, indice){

        Logger.log(
            (indice + 1) +
            ". " +
            nombre
        );

    });

    Logger.log(
        "========================================"
    );

    return instituciones;
}
function actualizarURLsAccesoIE() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const hoja = ss.getSheetByName("AccesosIE");

  if (!hoja) {
    throw new Error('No existe la hoja "AccesosIE".');
  }

  const datos = hoja.getDataRange().getValues();

  if (datos.length < 2) {
    throw new Error(
      "La hoja AccesosIE no tiene registros."
    );
  }

  const cabeceras = datos[0].map(function(valor) {
    return String(valor || "").trim();
  });

  const colIE = cabeceras.indexOf("IE");
  const colTOKEN = cabeceras.indexOf("TOKEN");

  if (colIE === -1) {
    throw new Error(
      'No existe la columna "IE".'
    );
  }

  if (colTOKEN === -1) {
    throw new Error(
      'No existe la columna "TOKEN".'
    );
  }

  /*
   * URL REAL DEL WEB APP
   */
  const urlBase =
    ScriptApp.getService().getUrl();

  Logger.log(
    "URL BASE DEL WEB APP: " + urlBase
  );

  if (!urlBase) {
    throw new Error(
      "Apps Script no pudo obtener la URL del Web App."
    );
  }

  /*
   * Buscar columnas existentes.
   */
  let colURL =
    cabeceras.indexOf("URL_ACCESO");

  let colLINK =
    cabeceras.indexOf("LINK_ACCESO");

  /*
   * Crear URL_ACCESO si no existe.
   */
  if (colURL === -1) {

    colURL = hoja.getLastColumn();

    hoja
      .getRange(1, colURL + 1)
      .setValue("URL_ACCESO");

    colURL = colURL;

  }

  /*
   * Crear LINK_ACCESO si no existe.
   */
  if (colLINK === -1) {

    colLINK = hoja.getLastColumn();

    hoja
      .getRange(1, colLINK + 1)
      .setValue("LINK_ACCESO");

    colLINK = colLINK;

  }

  /*
   * Volver a obtener las columnas después
   * de haberlas creado.
   */
  const encabezadosFinales =
    hoja
      .getRange(
        1,
        1,
        1,
        hoja.getLastColumn()
      )
      .getValues()[0]
      .map(function(valor) {
        return String(valor || "").trim();
      });

  colURL =
    encabezadosFinales.indexOf(
      "URL_ACCESO"
    );

  colLINK =
    encabezadosFinales.indexOf(
      "LINK_ACCESO"
    );

  Logger.log(
    "Columna IE: " + (colIE + 1)
  );

  Logger.log(
    "Columna TOKEN: " + (colTOKEN + 1)
  );

  Logger.log(
    "Columna URL_ACCESO: " + (colURL + 1)
  );

  Logger.log(
    "Columna LINK_ACCESO: " + (colLINK + 1)
  );

  /*
   * Preparar matrices.
   */
  const urls = [];
  const textos = [];

  let total = 0;

  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    const nombreIE =
      String(
        datos[i][colIE] || ""
      ).trim();

    const token =
      String(
        datos[i][colTOKEN] || ""
      ).trim();

    /*
     * Registro sin IE ni TOKEN.
     */
    if (
      !nombreIE &&
      !token
    ) {

      urls.push([""]);
      textos.push([""]);

      continue;
    }

    /*
     * Registro sin TOKEN.
     */
    if (!token) {

      Logger.log(
        "SIN TOKEN: " + nombreIE
      );

      urls.push([""]);
      textos.push([""]);

      continue;
    }

    /*
     * Construir enlace personalizado.
     */
    const url =
      urlBase +
      "?t=" +
      encodeURIComponent(token);

    const texto =
      "IE - " +
      nombreIE;

    urls.push([url]);
    textos.push([texto]);

    total++;

    /*
     * Mostrar algunos resultados en el registro.
     */
    if (total <= 10) {

      Logger.log(
        total +
        ". " +
        nombreIE +
        " -> " +
        url
      );

    }
  }

  /*
   * Escribir URL_ACCESO.
   */
  hoja
    .getRange(
      2,
      colURL + 1,
      urls.length,
      1
    )
    .setValues(urls);

  /*
   * Escribir LINK_ACCESO como texto.
   *
   * Google Sheets reconocerá automáticamente
   * las URLs como enlaces.
   */
  hoja
    .getRange(
      2,
      colLINK + 1,
      textos.length,
      1
    )
    .setValues(textos);

  /*
   * Convertir LINK_ACCESO en enlaces reales
   * mediante RichText.
   */
  for (
    let i = 0;
    i < urls.length;
    i++
  ) {

    if (!urls[i][0]) {
      continue;
    }

    const celda =
      hoja.getRange(
        i + 2,
        colLINK + 1
      );

    const richText =
      SpreadsheetApp
        .newRichTextValue()
        .setText(
          textos[i][0]
        )
        .setLinkUrl(
          urls[i][0]
        )
        .build();

    celda.setRichTextValue(
      richText
    );
  }

  SpreadsheetApp.flush();

  Logger.log(
    "======================================"
  );

  Logger.log(
    "TOTAL DE ENLACES GENERADOS: " +
    total
  );

  Logger.log(
    "======================================"
  );

  return {
    ok: true,
    total: total,
    urlBase: urlBase,
    mensaje:
      "Se generaron " +
      total +
      " enlaces personalizados."
  };
}
/*****************************************************
 * GENERAR ACCESOS FEM 2026
 *
 * CREA:
 * - 1 registro por IE
 * - 1 ID_FORO permanente por IE
 * - 1 TOKEN permanente por IE
 * - 1 código principal
 * - 4 códigos de contingencia
 * - 1 URL personalizada
 *
 * TOTAL:
 * - 37 IE
 * - 37 ID_FORO
 * - 37 TOKEN
 * - 185 códigos (5 por IE)
 * - 37 URLs
 *
 * IMPORTANTE:
 * - NO regenera accesos existentes.
 * - NO cambia ID_FORO existentes.
 * - NO cambia TOKEN existentes.
 * - NO borra datos.
 * - Si no encuentra exactamente 37 IE,
 *   NO modifica AccesosIE.
 *****************************************************/

function generarAccesosIE() {

    const lock =
        LockService.getScriptLock();

    lock.waitLock(30000);

    try {

        const ss =
            SpreadsheetApp.openById(
                SPREADSHEET_ID
            );

        /*
         * =================================================
         * 1. OBTENER LAS IE DESDE OFICIALES
         * =================================================
         */

        const instituciones =
            JSON.parse(
                obtenerInstitucionesJSON()
            );

        const nombresIE =
            Object.keys(
                instituciones || {}
            );

        Logger.log(
            "IE encontradas en Oficiales: " +
            nombresIE.length
        );

        /*
         * SEGURIDAD:
         * Solo continuar si existen exactamente 37.
         */

        if(
            nombresIE.length !== 37
        ){

            Logger.log(
                "NO SE GENERARON ACCESOS."
            );

            Logger.log(
                "Se esperaban 37 IE y se encontraron: " +
                nombresIE.length
            );

            return {
                ok: false,
                cantidad: nombresIE.length,
                mensaje:
                    "No se generaron accesos. " +
                    "El sistema esperaba exactamente 37 IE " +
                    "en la fuente Oficial y encontró " +
                    nombresIE.length +
                    ". Revise la hoja Oficiales."
            };

        }


        /*
         * =================================================
         * 2. OBTENER / CREAR AccesosIE
         * =================================================
         */

        let hoja =
            ss.getSheetByName(
                "AccesosIE"
            );

        if(!hoja){

            hoja =
                ss.insertSheet(
                    "AccesosIE"
                );

        }


        /*
         * =================================================
         * 3. ENCABEZADOS
         * =================================================
         */

        const encabezadosNecesarios = [

            "ID_ACCESO",
            "IE",
            "DANE",

            "CODIGO_ACCESO",
            "CODIGO_CONTINGENCIA_1",
            "CODIGO_CONTINGENCIA_2",
            "CODIGO_CONTINGENCIA_3",
            "CODIGO_CONTINGENCIA_4",

            "TOKEN",

            "URL_ACCESO",
            "LINK_ACCESO",

            "ID_FORO",

            "ESTADO",

            "TOKEN_SESION",
            "DISPOSITIVO_ID",

            "FECHA_GENERACION",
            "FECHA_PRIMER_ACCESO",
            "ULTIMA_ACTIVIDAD",
            "FECHA_ENVIO"

        ];


        let ultimaColumna =
            Math.max(
                hoja.getLastColumn(),
                1
            );


        let encabezadosActuales =
            hoja
                .getRange(
                    1,
                    1,
                    1,
                    ultimaColumna
                )
                .getValues()[0]
                .map(function(valor){

                    return String(
                        valor || ""
                    ).trim();

                });


        /*
         * Si la hoja está vacía,
         * crear encabezados.
         */

        if(
            hoja.getLastRow() === 0 ||
            (
                hoja.getLastRow() === 1 &&
                encabezadosActuales.every(
                    function(valor){
                        return valor === "";
                    }
                )
            )
        ){

            hoja
                .getRange(
                    1,
                    1,
                    1,
                    encabezadosNecesarios.length
                )
                .setValues([
                    encabezadosNecesarios
                ]);

        }else{

            /*
             * Agregar únicamente columnas
             * que no existan.
             */

            encabezadosNecesarios.forEach(
                function(nombre){

                    if(
                        encabezadosActuales.indexOf(
                            nombre
                        ) === -1
                    ){

                        hoja
                            .getRange(
                                1,
                                hoja.getLastColumn() + 1
                            )
                            .setValue(
                                nombre
                            );

                    }

                }
            );

        }


        /*
         * Volver a leer encabezados.
         */

        encabezadosActuales =
            hoja
                .getRange(
                    1,
                    1,
                    1,
                    hoja.getLastColumn()
                )
                .getValues()[0]
                .map(function(valor){

                    return String(
                        valor || ""
                    ).trim();

                });


        /*
         * Función para localizar columnas.
         */

        function columna(nombre){

            return encabezadosActuales.indexOf(
                nombre
            );

        }


        /*
         * =================================================
         * 4. LEER ACCESOS EXISTENTES
         * =================================================
         */

        const ultimaFila =
            hoja.getLastRow();

        let datosAccesos = [];

        if(
            ultimaFila >= 2
        ){

            datosAccesos =
                hoja
                    .getRange(
                        2,
                        1,
                        ultimaFila - 1,
                        hoja.getLastColumn()
                    )
                    .getValues();

        }


        /*
         * Índice IE -> fila existente.
         */

        const indiceExistentes = {};

        datosAccesos.forEach(
            function(fila, indice){

                const nombre =
                    String(
                        fila[
                            columna("IE")
                        ] || ""
                    ).trim();

                if(nombre){

                    indiceExistentes[
                        normalizarAccesoIE_(nombre)
                    ] = {
                        fila: indice + 2,
                        datos: fila
                    };

                }

            }
        );


        /*
         * =================================================
         * 5. TOKENS Y CÓDIGOS YA UTILIZADOS
         * =================================================
         */

        const tokensUsados = {};
        const codigosUsados = {};
        const idsForoUsados = {};

        datosAccesos.forEach(
            function(fila){

                const token =
                    String(
                        fila[
                            columna("TOKEN")
                        ] || ""
                    ).trim();

                if(token){
                    tokensUsados[token] = true;
                }


                [
                    "CODIGO_ACCESO",
                    "CODIGO_CONTINGENCIA_1",
                    "CODIGO_CONTINGENCIA_2",
                    "CODIGO_CONTINGENCIA_3",
                    "CODIGO_CONTINGENCIA_4"
                ].forEach(
                    function(nombreColumna){

                        const indice =
                            columna(
                                nombreColumna
                            );

                        if(indice === -1){
                            return;
                        }

                        const codigo =
                            String(
                                fila[indice] || ""
                            ).trim();

                        if(codigo){
                            codigosUsados[codigo] = true;
                        }

                    }
                );


                const idForo =
                    String(
                        fila[
                            columna("ID_FORO")
                        ] || ""
                    ).trim();

                if(idForo){
                    idsForoUsados[idForo] = true;
                }

            }
        );


        /*
         * =================================================
         * 6. FUNCIONES DE GENERACIÓN
         * =================================================
         */

        function generarCodigoUnico(){

            const caracteres =
                "ABCDEFGHJKMNPQRSTUVWXYZ23456789";

            let codigo = "";

            do{

                codigo =
                    "FEM-" +
                    caracteres[
                        Math.floor(
                            Math.random() *
                            caracteres.length
                        )
                    ] +
                    caracteres[
                        Math.floor(
                            Math.random() *
                            caracteres.length
                        )
                    ] +
                    caracteres[
                        Math.floor(
                            Math.random() *
                            caracteres.length
                        )
                    ] +
                    caracteres[
                        Math.floor(
                            Math.random() *
                            caracteres.length
                        )
                    ] +
                    caracteres[
                        Math.floor(
                            Math.random() *
                            caracteres.length
                        )
                    ];

            }while(
                codigosUsados[codigo]
            );

            codigosUsados[codigo] = true;

            return codigo;

        }


        function generarTokenUnico(){

            let token = "";

            do{

                token =
                    Utilities.getUuid()
                        .replace(
                            /-/g,
                            ""
                        );

            }while(
                tokensUsados[token]
            );

            tokensUsados[token] = true;

            return token;

        }


        function generarIdForoUnico(){

            let idForo = "";

            do{

                idForo =
                    Utilities.getUuid();

            }while(
                idsForoUsados[idForo]
            );

            idsForoUsados[idForo] = true;

            return idForo;

        }


        /*
         * =================================================
         * 7. URL BASE
         * =================================================
         */

        const urlBase =
            ScriptApp.getService()
                .getUrl();

        if(!urlBase){

            throw new Error(
                "No fue posible obtener la URL del Web App."
            );

        }


        /*
         * =================================================
         * 8. CREAR / CONSERVAR ACCESOS
         * =================================================
         */

        let creados = 0;
        let conservados = 0;


        nombresIE.forEach(
            function(nombreIE){

                const clave =
                    normalizarAccesoIE_(
                        nombreIE
                    );

                const existente =
                    indiceExistentes[
                        clave
                    ];


                /*
                 * Datos oficiales.
                 */

                const datosIE =
                    instituciones[
                        nombreIE
                    ] || {};


                const dane =
                    String(
                        datosIE.dane ||
                        ""
                    ).trim();


                /*
                 * =================================================
                 * SI YA EXISTE
                 * =================================================
                 */

                if(existente){

                    conservados++;

                    Logger.log(
                        "CONSERVADO: " +
                        nombreIE
                    );

                    return;

                }


                /*
                 * =================================================
                 * NUEVO ACCESO
                 * =================================================
                 */

                const codigoPrincipal =
                    generarCodigoUnico();

                const contingencia1 =
                    generarCodigoUnico();

                const contingencia2 =
                    generarCodigoUnico();

                const contingencia3 =
                    generarCodigoUnico();

                const contingencia4 =
                    generarCodigoUnico();


                const token =
                    generarTokenUnico();


                /*
                 * ID_FORO permanente.
                 *
                 * Se genera SOLO aquí,
                 * en la preparación administrativa.
                 *
                 * Nunca en el navegador.
                 */

                const idForo =
                    generarIdForoUnico();


                /*
                 * URL personalizada.
                 *
                 * El TOKEN es el identificador técnico.
                 */

                const url =
                    urlBase +
                    "?t=" +
                    encodeURIComponent(
                        token
                    );


                const fecha =
                    new Date();


                /*
                 * ID interno.
                 */

                const idAcceso =
                    Utilities.getUuid();


                /*
                 * Estado inicial.
                 */

                const estado =
                    "DISPONIBLE";


                /*
                 * Crear fila completa.
                 */

                const nuevaFila =
                    new Array(
                        hoja.getLastColumn()
                    ).fill("");


                nuevaFila[
                    columna("ID_ACCESO")
                ] =
                    idAcceso;


                nuevaFila[
                    columna("IE")
                ] =
                    nombreIE;


                nuevaFila[
                    columna("DANE")
                ] =
                    dane;


                nuevaFila[
                    columna("CODIGO_ACCESO")
                ] =
                    codigoPrincipal;


                nuevaFila[
                    columna(
                        "CODIGO_CONTINGENCIA_1"
                    )
                ] =
                    contingencia1;


                nuevaFila[
                    columna(
                        "CODIGO_CONTINGENCIA_2"
                    )
                ] =
                    contingencia2;


                nuevaFila[
                    columna(
                        "CODIGO_CONTINGENCIA_3"
                    )
                ] =
                    contingencia3;


                nuevaFila[
                    columna(
                        "CODIGO_CONTINGENCIA_4"
                    )
                ] =
                    contingencia4;


                nuevaFila[
                    columna("TOKEN")
                ] =
                    token;


                nuevaFila[
                    columna("URL_ACCESO")
                ] =
                    url;


                nuevaFila[
                    columna("LINK_ACCESO")
                ] =
                    "IE - " +
                    nombreIE;


                nuevaFila[
                    columna("ID_FORO")
                ] =
                    idForo;


                nuevaFila[
                    columna("ESTADO")
                ] =
                    estado;


                nuevaFila[
                    columna("FECHA_GENERACION")
                ] =
                    fecha;


                hoja.appendRow(
                    nuevaFila
                );


                /*
                 * Convertir LINK_ACCESO
                 * en enlace clicable.
                 */

                const filaNueva =
                    hoja.getLastRow();


                const celdaLink =
                    hoja.getRange(
                        filaNueva,
                        columna(
                            "LINK_ACCESO"
                        ) + 1
                    );


                const richText =
                    SpreadsheetApp
                        .newRichTextValue()
                        .setText(
                            "IE - " +
                            nombreIE
                        )
                        .setLinkUrl(
                            url
                        )
                        .build();


                celdaLink
                    .setRichTextValue(
                        richText
                    );


                creados++;

                Logger.log(
                    "CREADO: " +
                    nombreIE
                );

            }
        );


        SpreadsheetApp.flush();


        /*
         * =================================================
         * 9. RESULTADO
         * =================================================
         */

        Logger.log(
            "========================================"
        );

        Logger.log(
            "IE encontradas: " +
            nombresIE.length
        );

        Logger.log(
            "Accesos nuevos: " +
            creados
        );

        Logger.log(
            "Accesos conservados: " +
            conservados
        );

        Logger.log(
            "Códigos disponibles por IE: 5"
        );

        Logger.log(
            "Total de códigos para 37 IE: " +
            (
                nombresIE.length * 5
            )
        );

        Logger.log(
            "========================================"
        );


        return {

            ok: true,

            instituciones:
                nombresIE.length,

            creados:
                creados,

            conservados:
                conservados,

            codigosPorIE:
                5,

            totalCodigos:
                nombresIE.length * 5,

            mensaje:
                "Accesos FEM generados/conservados correctamente."

        };


    }finally{

        lock.releaseLock();

    }

}


/*****************************************************
 * NORMALIZAR NOMBRE DE IE PARA ACCESOS
 *****************************************************/
function normalizarAccesoIE_(texto){

    return String(
        texto || ""
    )
    .normalize("NFD")
    .replace(
        /[\u0300-\u036f]/g,
        ""
    )
    .replace(
        /[.,]/g,
        ""
    )
    .replace(
        /\s+/g,
        " "
    )
    .trim()
    .toLowerCase();

}

function asegurarColumnasAccesosIE_(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  let hoja=ss.getSheetByName(HOJA_ACCESOS);
  if(!hoja) hoja=ss.insertSheet(HOJA_ACCESOS);
  const requeridas=["ID_ACCESO","IE","DANE","CODIGO_ACCESO","TOKEN","URL_ACCESO","ID_FORO","ESTADO","TOKEN_SESION","DISPOSITIVO_ID","FECHA_GENERACION","FECHA_PRIMER_ACCESO","ULTIMA_ACTIVIDAD","FECHA_ENVIO","EMAIL_IE","EMAIL_RESPONSABLE","TIPO","S1_ENVIADA","S2_ENVIADA","S3_ENVIADA","ID_INFORME","ID_PDF_INFORME"];
  const last=hoja.getLastColumn();
  const existentes=last?hoja.getRange(1,1,1,last).getValues()[0].map(String):[];
  if(!last){hoja.getRange(1,1,1,requeridas.length).setValues([requeridas]);}
  else{
    const faltantes=requeridas.filter(h=>existentes.indexOf(h)===-1);
    if(faltantes.length) hoja.getRange(1,last+1,1,faltantes.length).setValues([faltantes]);
  }
  return hoja;
}


function mapaHoja_(hoja){
  const last=hoja.getLastColumn(); const h=last?hoja.getRange(1,1,1,last).getDisplayValues()[0]:[]; const m={}; h.forEach((x,i)=>{if(String(x).trim())m[String(x).trim()]=i+1;}); return m;
}


function filaAccesoPorToken_(token){
  const hoja=asegurarColumnasAccesosIE_(); const m=mapaHoja_(hoja); if(!m.TOKEN)return null;
  const last=hoja.getLastRow(); if(last<2)return null;
  const rows=hoja.getRange(2,1,last-1,hoja.getLastColumn()).getValues();
  for(let i=0;i<rows.length;i++) if(String(rows[i][m.TOKEN-1]||"").trim()===String(token||"").trim()) return {hoja:hoja,mapa:m,fila:i+2,valores:rows[i]};
  return null;
}


function generarCodigoAcceso_(){
  const palabras=["LUNA","NUBE","RUTA","VOZ","AULA","FARO","NORTE","RIO","SOL","PUENTE","VIVA","VALLE"];
  const chars="ABCDEFGHJKMNPQRSTUVWXYZ23456789";
  let n=""; for(let i=0;i<4;i++) n+=chars.charAt(Math.floor(Math.random()*chars.length));
  return "FEM-"+palabras[Math.floor(Math.random()*palabras.length)]+n;
}


function nombreHojaIE_(nombre){
  let s=String(nombre||"IE").replace(/[\\/?*\[\]:]/g," ").replace(/\s+/g," ").trim(); if(s.length>95)s=s.substring(0,95); return s||"IE";
}


function inicializarHojasIE(){
  try{
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    const instituciones=obtenerInstituciones();
    const headers=obtenerCabecerasAvancesForo();
    const created=[];
    Object.keys(instituciones).forEach(ie=>{
      const name=nombreHojaIE_(ie);
      let sh=ss.getSheetByName(name);
      if(!sh){sh=ss.insertSheet(name);created.push(name);}
      if(sh.getLastColumn()<headers.length)sh.getRange(1,1,1,headers.length).setValues([headers]);
    });
    inicializarParticipacion_();
    return {ok:true,creadas:created};
  }catch(e){return {ok:false,mensaje:e.message};}
}


function inicializarParticipacion_(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID); let sh=ss.getSheetByName(HOJA_PARTICIPACION); if(!sh)sh=ss.insertSheet(HOJA_PARTICIPACION);
  const h=["IE","ID_FORO","FECHA","Rector(a)","Coordinador(a)","Docentes","Tutor PTA PFI/3.0","Orientador(a)","Estudiantes","Padres/madres/acudientes","Personal administrativo","Egresados","Sector productivo","Otros","Total"];
  const last=sh.getLastColumn(); if(!last)sh.getRange(1,1,1,h.length).setValues([h]); else{const ex=sh.getRange(1,1,1,last).getValues()[0].map(String);const f=h.filter(x=>ex.indexOf(x)===-1);if(f.length)sh.getRange(1,last+1,1,f.length).setValues([f]);}
  return sh;
}


function actualizarParticipacion_(datos){
  const sh=inicializarParticipacion_(); const m=mapaHoja_(sh); const campos=datos.campos||{}; const vals=["participantesRector","participantesCoordinador","participantesDocentes","participantesTutorPTA","participantesOrientador","participantesEstudiantes","participantesPadres","participantesAdministrativos","participantesEgresados","participantesSector","participantesOtros"].map(id=>Number(campos[id]?.valor||0)); const total=vals.reduce((a,b)=>a+b,0); const row={IE:datos.institucion||"",ID_FORO:datos.idForo||"",FECHA:new Date(),"Rector(a)":vals[0],"Coordinador(a)":vals[1],"Docentes":vals[2],"Tutor PTA PFI/3.0":vals[3],"Orientador(a)":vals[4],"Estudiantes":vals[5],"Padres/madres/acudientes":vals[6],"Personal administrativo":vals[7],"Egresados":vals[8],"Sector productivo":vals[9],"Otros":vals[10],"Total":total};
  const last=sh.getLastRow(); let found=-1;if(last>=2){const ids=sh.getRange(2,m.ID_FORO,last-1,1).getValues();for(let i=0;i<ids.length;i++)if(String(ids[i][0]||"")===String(datos.idForo||"")){found=i+2;break;}}
  const out=new Array(sh.getLastColumn()).fill(""); Object.keys(row).forEach(k=>{if(m[k])out[m[k]-1]=row[k];}); if(found>0)sh.getRange(found,1,1,out.length).setValues([out]); else sh.appendRow(out);
}


function guardarEnHojaIE_(datos){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID); const shName=nombreHojaIE_(datos.institucion); let sh=ss.getSheetByName(shName); if(!sh){sh=ss.insertSheet(shName); const headers=obtenerCabecerasAvancesForo(); sh.getRange(1,1,1,headers.length).setValues([headers]);}
  if(!sh)return; const headers=obtenerCabecerasAvancesForo(); if(sh.getLastColumn()<headers.length)sh.getRange(1,1,1,headers.length).setValues([headers]); const m=mapaHoja_(sh); const respuestas=extraerRespuestasSesiones_(datos);
  const row=Object.assign({ID_FORO:datos.idForo,INSTITUCION:datos.institucion,DANE:datos.dane,FECHA_INICIO:datos.fechaInicio||new Date(),ULTIMA_ACTUALIZACION:new Date(),ESTADO:"En proceso"},respuestas,{DATOS:JSON.stringify(datos)});
  const out=new Array(sh.getLastColumn()).fill(""); Object.keys(row).forEach(k=>{if(m[k])out[m[k]-1]=normalizarValorHoja_(row[k]);}); let found=-1;if(sh.getLastRow()>=2){const ids=sh.getRange(2,m.ID_FORO,sh.getLastRow()-1,1).getValues();for(let i=0;i<ids.length;i++)if(String(ids[i][0]||"")===String(datos.idForo||"")){found=i+2;break;}} if(found>0)sh.getRange(found,1,1,out.length).setValues([out]);else sh.appendRow(out);
}


function obtenerEstadoSesiones_(idForo){
  const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HOJA_AVANCES); if(!sh||sh.getLastRow()<2)return {s1:false,s2:false,s3:false}; const m=mapaHoja_(sh); const row=buscarFilaPorIdForo_(sh,idForo,m); if(row<0)return {s1:false,s2:false,s3:false}; return {s1:String(sh.getRange(row,m.S1_ENVIADA).getValue()).toUpperCase()==="SI",s2:String(sh.getRange(row,m.S2_ENVIADA).getValue()).toUpperCase()==="SI",s3:String(sh.getRange(row,m.S3_ENVIADA).getValue()).toUpperCase()==="SI"};
}


function obtenerDatosGuardadosPorIdForo_(idForo){
  const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HOJA_AVANCES); if(!sh||sh.getLastRow()<2)return null; const m=mapaHoja_(sh); const row=buscarFilaPorIdForo_(sh,idForo,m); if(row<0)return null; const raw=sh.getRange(row,m.DATOS).getValue(); if(!raw)return null; try{return JSON.parse(raw);}catch(e){return null;}
}


function crearCarpetaIE_(ie){
  const root=DriveApp.getFolderById(DRIVE_CARPETA_FEM_ID); const it=root.getFoldersByName(ie); if(it.hasNext())return it.next(); return root.createFolder(ie);
}


function hacerPublicoSiEsPosible_(file){try{file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(e){Logger.log("No se pudo cambiar compartir: "+e.message);}}


function subirEvidenciasFEM(idForo,pdfData,pdfName,pdfMime,fotoData,fotoName,fotoMime,datos){
  const acceso=obtenerAccesoPorIdForo_(idForo); if(!acceso)throw new Error("ID_FORO no autorizado."); const folder=crearCarpetaIE_(datos.institucion||acceso.ie);
  const decode=(data)=>{const s=String(data||"");const comma=s.indexOf(",");return Utilities.base64Decode(comma>=0?s.substring(comma+1):s);}; const pb=decode(pdfData),fb=decode(fotoData); if(pb.length>10*1024*1024||fb.length>10*1024*1024)throw new Error("Cada archivo debe pesar máximo 10 MB."); if(pdfMime!=="application/pdf")throw new Error("La asistencia debe ser PDF."); if(["image/jpeg","image/png"].indexOf(fotoMime)<0)throw new Error("La fotografía debe ser JPG o PNG.");
  const pf=folder.createFile(Utilities.newBlob(pb,pdfMime,pdfName));
  const fotoNombre="Foro 2026 ("+(datos.institucion||acceso.ie)+")."+(fotoMime==="image/png"?"png":"jpg");
  const foto=folder.createFile(Utilities.newBlob(fb,fotoMime,fotoNombre));
  pf.setDescription("Listado de asistencia a foro educativo | I.E. "+(datos.institucion||acceso.ie)+" | Responsable: "+(datos.campos?.nombre?.valor||"")+" | Email: "+(datos.campos?.correo?.valor||"")+" | Cargo: "+(datos.campos?.cargo?.valor||"")+" | Grupo de trabajo FEM 2026: "+(datos.campos?.grupo?.valor||"")+" | Fecha y hora de submisión: "+new Date()); foto.setDescription("Participantes del FEM 2026 | I.E. "+(datos.institucion||acceso.ie)+" | Grupo: "+(datos.campos?.grupo?.valor||"")+" | Fecha y hora de submisión: "+new Date()); hacerPublicoSiEsPosible_(pf);hacerPublicoSiEsPosible_(foto); return {ok:true,asistencia:{id:pf.getId(),url:pf.getUrl()},foto:{id:foto.getId(),url:foto.getUrl()},folderId:folder.getId()};
}


function obtenerGraficoParticipacionBlob_(datos){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID); const sh=ss.getSheetByName(HOJA_PARTICIPACION); if(!sh)return null; const m=mapaHoja_(sh); const row=buscarFilaPorIdForo_(sh,datos.idForo,m); if(row<0)return null; const labels=["Rector(a)","Coordinador(a)","Docentes","Tutor PTA PFI/3.0","Orientador(a)","Estudiantes","Padres/madres/acudientes","Personal administrativo","Egresados","Sector productivo","Otros"]; const nums=labels.map(l=>Number(sh.getRange(row,m[l]).getValue()||0)); const tmp=ss.insertSheet("_grafico_"+String(datos.idForo).slice(0,8)); try{tmp.getRange(1,1,labels.length,2).setValues(labels.map((l,i)=>[l,nums[i]])); const chart=tmp.newChart().setChartType(Charts.ChartType.PIE).addRange(tmp.getRange(1,1,labels.length,2)).setOption("title","Participación por estamento — "+(datos.institucion||"")).setPosition(1,4,0,0).build(); tmp.insertChart(chart); SpreadsheetApp.flush(); const c=tmp.getCharts()[0]; return c.getAs("image/png");}finally{ss.deleteSheet(tmp);}}


function construirParrafoSesion_(titulo,contenido){return titulo+"\n\n"+String(contenido||"");}


function generarInformeFEM(idForo,datosCliente){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const estadoFinal=obtenerEstadoSesiones_(idForo); if(!estadoFinal.s1||!estadoFinal.s2||!estadoFinal.s3)throw new Error("Las tres sesiones deben estar enviadas definitivamente antes de generar el informe.");
    const datos=obtenerDatosGuardadosPorIdForo_(idForo)||datosCliente; if(!datos)throw new Error("No hay datos guardados para generar el informe."); const folder=crearCarpetaIE_(datos.institucion||"Institución Educativa");
    const template=DriveApp.getFileById(TEMPLATE_INFORME_ID); const copy=template.makeCopy("Informe Ejecutivo - "+datos.institucion+" FEM 2026",folder); const doc=DocumentApp.openById(copy.getId()); const body=doc.getBody(); body.clear(); body.setPageWidth(612).setPageHeight(792).setMarginTop(50).setMarginBottom(50).setMarginLeft(48).setMarginRight(48);
    const h=doc.getHeader()||doc.addHeader(); h.clear(); const hi=h.appendParagraph(); hi.setAlignment(DocumentApp.HorizontalAlignment.RIGHT); try{hi.appendInlineImage(DriveApp.getFileById(LOGO_ENCABEZADO_ID).getBlob()).setWidth(90).setHeight(50);}catch(e){};
    const footer=doc.getFooter()||doc.addFooter(); footer.clear(); const fp=footer.appendParagraph(); fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER); try{fp.appendInlineImage(DriveApp.getFileById(LOGO_PIE_ID).getBlob()).setWidth(80).setHeight(40);}catch(e){}; footer.appendParagraph("Generado por SEM el "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy 'a las' HH:mm")+". Enviado por "+(datos.campos?.nombre?.valor||"")+" — "+(datos.campos?.correo?.valor||"")+" — "+(datos.campos?.cargo?.valor||""));
    const title=body.appendParagraph("INFORME EJECUTIVO DE "+String(datos.institucion||"").toUpperCase()+" FEM 2026"); title.setHeading(DocumentApp.ParagraphHeading.TITLE); body.appendParagraph("FEM 2026 “Escuela Viva: Voces que construyen territorio”.").setHeading(DocumentApp.ParagraphHeading.HEADING2); body.appendParagraph("Foro Educativo Institucional — Neiva 2026"); body.appendHorizontalRule();
    body.appendParagraph("Caracterización y participación").setHeading(DocumentApp.ParagraphHeading.HEADING1); const table=body.appendTable(); const c=datos.campos||{}; [["Institución Educativa",datos.institucion],["DANE",datos.dane],["Rector(a)",c.rector?.valor||""],["Grupo de trabajo",c.grupo?.valor||""],["Responsable",c.nombre?.valor||""],["Cargo",c.cargo?.valor||""],["Correo responsable",c.correo?.valor||""],["Correo institucional",c.correoIE?.valor||""]].forEach(x=>{const r=table.appendTableRow();r.appendTableCell(x[0]);r.appendTableCell(String(x[1]||"—"));});
    body.appendParagraph("Participantes: "+totalParticipantesServer_(datos)).setHeading(DocumentApp.ParagraphHeading.HEADING2); const chart=obtenerGraficoParticipacionBlob_(datos); if(chart)body.appendImage(chart).setWidth(430);
    body.appendPageBreak();
    const grupos=[{n:"Sesión 1",items:[["Pregunta orientadora",c.respuestaSesion1?.valor||""],["Pregunta 2",c.respuestaSesion1Pregunta2?.valor||""]]},{n:"Sesión 2",items:[["Pregunta 1",c.respuestaSesion2Pregunta1?.valor||""],["Acciones 1–5",[1,2,3,4,5].map(i=>c["respuestaSesion2Pregunta2Accion"+i]?.valor||"").filter(Boolean).join("\n\n")],["Pregunta 3",c.respuestaSesion2Pregunta3?.valor||""],["Pregunta 4",c.respuestaSesion2Pregunta4?.valor||""],["Pregunta 5",c.respuestaSesion2Pregunta5?.valor||""]]},{n:"Sesión 3",items:[["Pregunta 1",c.respuestaSesion3Pregunta1?.valor||""],["Acciones 1–5",[1,2,3,4,5].map(i=>c["respuestaSesion3Pregunta2Accion"+i]?.valor||"").filter(Boolean).join("\n\n")],["Equipos de trabajo",c.respuestaSesion3Pregunta3?.valor||""],["Mecanismos de seguimiento",c.respuestaSesion3Pregunta4?.valor||""]]}];
    grupos.forEach((g,gi)=>{
      body.appendParagraph(g.n).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      const t=body.appendTable();
      t.setBorderWidth(0.5);
      g.items.forEach(it=>{ const r=t.appendTableRow(); r.appendTableCell(String(it[0]||"")); r.appendTableCell(String(it[1]||"—")); });
      if(gi<grupos.length-1)body.appendPageBreak();
    });
    body.appendPageBreak();
    body.appendParagraph("Evidencias de la jornada").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Los soportes documentales originales se encuentran almacenados en la carpeta institucional de la IE en Google Drive.");
    if(c.evidenciaAsistenciaUrl?.valor){ const p1=body.appendParagraph(); p1.appendText("📄 Descargar listado de asistencia").setLinkUrl(String(c.evidenciaAsistenciaUrl.valor)); }
    if(c.evidenciaFotoUrl?.valor){ const p2=body.appendParagraph(); p2.appendText("📷 Ver fotografía de la plenaria").setLinkUrl(String(c.evidenciaFotoUrl.valor)); }
    doc.saveAndClose(); const pdfBlob=DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF).setName("Informe Ejecutivo - "+datos.institucion+" FEM 2026.pdf"); const pdf=folder.createFile(pdfBlob); hacerPublicoSiEsPosible_(copy);hacerPublicoSiEsPosible_(pdf);
    const ac=filaAccesoPorToken_(obtenerTokenPorIdForo_(idForo)); if(ac){const m=ac.mapa; if(m.ID_INFORME)ac.hoja.getRange(ac.fila,m.ID_INFORME).setValue(copy.getId()); if(m.ID_PDF_INFORME)ac.hoja.getRange(ac.fila,m.ID_PDF_INFORME).setValue(pdf.getId());}
    return {ok:true,docId:copy.getId(),docUrl:copy.getUrl(),pdfId:pdf.getId(),pdfUrl:pdf.getUrl(),folderId:folder.getId()};
  }finally{try{lock.releaseLock();}catch(e){}}
}


function totalParticipantesServer_(datos){const c=datos.campos||{};return ["Rector","Coordinador","Docentes","TutorPTA","Orientador","Estudiantes","Padres","Administrativos","Egresados","Sector","Otros"].reduce((s,k)=>s+Number(c["participantes"+k]?.valor||0),0);}


function enviarInformeFEM(idForo,datos,pdfId){
  const acceso=obtenerAccesoPorIdForoRaw_(idForo); if(!acceso)throw new Error("ID_FORO no autorizado."); const c=datos.campos||{}; const ie=datos.institucion||acceso.ie; const destinatario=String(c.correoIE?.valor||acceso.email||"").trim(); const responsable=String(c.correo?.valor||"").trim(); if(!destinatario)throw new Error("La institución no tiene correo institucional registrado."); const aliases=GmailApp.getAliases().map(x=>x.toLowerCase()); const cuenta=Session.getEffectiveUser().getEmail().toLowerCase(); if(cuenta!==REMITENTE_FEM&&aliases.indexOf(REMITENTE_FEM)===-1)throw new Error("La cuenta de Apps Script no puede enviar como "+REMITENTE_FEM+". Configure esa cuenta o un alias."); const subject="Reporte de Informe IE "+ie; const body="Apreciados(as) integrantes de la comunidad educativa de la Institución Educativa "+ie+":\n\nReciban un cordial saludo de la Secretaría de Educación de Neiva.\n\nAgradecemos a la Institución Educativa por su participación y por el tiempo dedicado al desarrollo del Foro Educativo Institucional – Neiva 2026, así como por los aportes, reflexiones y propuestas construidas colectivamente durante la jornada.\n\nAdjuntamos el Informe Ejecutivo del Foro Educativo Institucional – Neiva 2026, que reúne la caracterización institucional, la participación registrada y las respuestas definitivas construidas durante las tres sesiones de trabajo.\n\nAgradecemos especialmente la disposición de la comunidad educativa para participar en este ejercicio de diálogo, reflexión y construcción colectiva orientado al fortalecimiento de la educación en nuestro municipio.\n\nSecretaría de Educación de Neiva\nForo Educativo Institucional – Neiva 2026\n\“Escuela Viva: Voces que construyen territorio\”"; const to=destinatario; const cc=[responsable].concat(COPIAS_INFORME_FEM).filter(Boolean).join(","); const file=DriveApp.getFileById(pdfId); GmailApp.sendEmail(to,subject,body,{htmlBody:"<p>Apreciados(as) integrantes de la comunidad educativa de la <strong>"+ie+"</strong>:</p><p>Reciban un cordial saludo de la Secretaría de Educación de Neiva.</p><p>Agradecemos a la Institución Educativa por su participación y por el tiempo dedicado al desarrollo del <strong>Foro Educativo Institucional – Neiva 2026</strong>, así como por los aportes, reflexiones y propuestas construidas colectivamente durante la jornada.</p><p>Adjuntamos el <strong>Informe Ejecutivo del Foro Educativo Institucional – Neiva 2026</strong>, que reúne la caracterización institucional, la participación registrada y las respuestas definitivas construidas durante las tres sesiones de trabajo.</p><p>Agradecemos especialmente la disposición de la comunidad educativa para participar en este ejercicio de diálogo, reflexión y construcción colectiva orientado al fortalecimiento de la educación en nuestro municipio.</p><p><strong>Secretaría de Educación de Neiva</strong><br>Foro Educativo Institucional – Neiva 2026<br>“Escuela Viva: Voces que construyen territorio”</p>",cc:cc,from:REMITENTE_FEM,name:"Secretaría de Educación de Neiva",attachments:[file.getBlob()]}); return {ok:true};
}


function finalizarFormularioFEM(idForo,tokenSesion,dispositivoId,pdfId){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const ac=obtenerAccesoPorIdForoRaw_(idForo);
    if(!ac)return {ok:false,mensaje:"No se encontró el acceso."};
    const estado=String(ac.hoja.getRange(ac.fila,ac.mapa.ESTADO).getValue()||"").toUpperCase();
    if(estado!=="ENVIADO")return {ok:false,mensaje:"El Foro aún no tiene un envío definitivo registrado."};
    if(!pdfId)return {ok:false,mensaje:"No se ha registrado el PDF del informe."};
    if(ac.mapa.FECHA_ENVIO)ac.hoja.getRange(ac.fila,ac.mapa.FECHA_ENVIO).setValue(new Date());
    const props=PropertiesService.getScriptProperties(); props.deleteProperty(obtenerClaveSesionCodigo_("","",idForo));
    const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HOJA_AVANCES);
    if(sh){const mm=obtenerMapaCabeceras_(sh);const row=buscarFilaPorIdForo_(sh,idForo,mm);if(row>0&&mm.ESTADO)sh.getRange(row,mm.ESTADO).setValue("ENVIADO");}
    return {ok:true,pdfId:pdfId};
  }finally{try{lock.releaseLock();}catch(e){}}
}

/*****************************************************
 * ENVÍO DEFINITIVO POR SESIÓN (Sesión 1, 2 y 3)
 *****************************************************/
function enviarRespuestasSesion(idForo, tokenSesion, dispositivoId, sesion, datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    idForo = String(idForo || "").trim();
    sesion = Number(sesion);

    if (idForo === "") {
      return { ok: false, mensaje: "ID_FORO obligatorio." };
    }

    if ([1, 2, 3].indexOf(sesion) === -1) {
      return { ok: false, mensaje: "Número de sesión inválido." };
    }

    if (!sesionActivaPorIdForo_(idForo, dispositivoId, tokenSesion)) {
      return {
        ok: false,
        mensaje: "La sesión ya no está activa. Ingrese nuevamente con uno de los códigos de la institución."
      };
    }

    const estadoActual = obtenerEstadoSesiones_(idForo);
    const clave = "s" + sesion;

    if (estadoActual[clave]) {
      return {
        ok: true,
        yaEnviada: true,
        mensaje: "Las respuestas de la Sesión " + sesion + " ya habían sido enviadas."
      };
    }

    if (sesion === 2 && !estadoActual.s1) {
      return { ok: false, mensaje: "Debe enviar primero las respuestas de la Sesión 1." };
    }

    if (sesion === 3 && !estadoActual.s2) {
      return { ok: false, mensaje: "Debe enviar primero las respuestas de la Sesión 2." };
    }

    const guardado = guardarAvanceForo(datos);
    if (!guardado || !guardado.ok) {
      throw new Error((guardado && guardado.mensaje) || "No fue posible guardar las respuestas.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(HOJA_AVANCES);
    const mapa = obtenerMapaCabeceras_(hoja);
    const fila = buscarFilaPorIdForo_(hoja, idForo, mapa);

    if (fila < 0) {
      throw new Error("No fue posible ubicar el registro del formulario para esta institución.");
    }

    const ahora = new Date();
    const colEnviada = mapa["S" + sesion + "_ENVIADA"];
    const colFecha = mapa["FECHA_ENVIO_S" + sesion];

    if (colEnviada) hoja.getRange(fila, colEnviada).setValue("SI");
    if (colFecha) hoja.getRange(fila, colFecha).setValue(ahora);

    SpreadsheetApp.flush();

    return { ok: true, sesion: sesion, fecha: ahora.toISOString() };

  } catch (error) {

    return { ok: false, mensaje: error.message };

  } finally {

    try { lock.releaseLock(); } catch (e) { /* No hacer nada. */ }

  }
}
/*****************************************************
 * ENVÍO DEFINITIVO ÚNICO DEL FORO
 * Los 5 códigos comparten ID_FORO y solo existe un envío.
 *****************************************************/
function obtenerAccesoPorIdForoRaw_(idForo){
  const id=String(idForo||"").trim(); if(!id)return null;
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID); const sh=ss.getSheetByName(HOJA_ACCESOS); if(!sh||sh.getLastRow()<2)return null;
  const vals=sh.getDataRange().getDisplayValues(); const h=vals[0].map(String); const m={}; h.forEach((x,i)=>m[String(x).trim()]=i);
  for(let i=1;i<vals.length;i++) if(String(vals[i][m.ID_FORO]||"").trim()===id) return {hoja:sh,fila:i+1,mapa:Object.fromEntries(Object.keys(m).map(k=>[k,m[k]+1])),ie:String(vals[i][m.IE]||""),dane:String(vals[i][m.DANE]||""),email:m.EMAIL_IE!==undefined?String(vals[i][m.EMAIL_IE]||""):""};
  return null;
}
function sesionActivaPorIdForo_(idForo,dispositivoId,tokenSesion){
  const props=PropertiesService.getScriptProperties(); const clave=obtenerClaveSesionCodigo_("","",idForo); const raw=props.getProperty(clave); if(!raw)return false; let a; try{a=JSON.parse(raw);}catch(e){return false;} if(a.deviceId!==String(dispositivoId||"")||a.tokenSesion!==String(tokenSesion||""))return false; if(Date.now()-Number(a.ultimaActividad||0)>SESION_CODIGO_TTL_MS)return false; return true;
}
function validarEnvioFinal_(datos){
  const c=datos?.campos||{}; const v=id=>String(c[id]?.valor||"").trim(); const words=t=>String(t||"").trim().split(/\s+/).filter(Boolean).length;
  const req=["respuestaSesion1","respuestaSesion1Pregunta2","respuestaSesion2Pregunta1","respuestaSesion2Pregunta2Accion1","respuestaSesion2Pregunta2Accion2","respuestaSesion2Pregunta2Accion3","respuestaSesion2Pregunta3","respuestaSesion2Pregunta4","respuestaSesion2Pregunta5","respuestaSesion3Pregunta1","respuestaSesion3Pregunta2Accion1","respuestaSesion3Pregunta2Accion2","respuestaSesion3Pregunta2Accion3","respuestaSesion3Pregunta3","respuestaSesion3Pregunta4"];
  const falt=req.filter(id=>!v(id)); if(falt.length)return {ok:false,mensaje:"Faltan respuestas obligatorias antes de realizar el envío definitivo."};
  const s3p1=words(v("respuestaSesion3Pregunta1")); if(s3p1<40||s3p1>200)return {ok:false,mensaje:"La pregunta 1 de la Sesión 3 debe tener entre 40 y 200 palabras."};
  for(let i=1;i<=3;i++){const n=words(v("respuestaSesion3Pregunta2Accion"+i));if(n<20||n>100)return {ok:false,mensaje:"Las acciones 1, 2 y 3 de la Sesión 3 deben tener entre 20 y 100 palabras."};}
  return {ok:true};
}
function enviarForoDefinitivo(idForo,tokenSesion,dispositivoId,datos){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const raw=obtenerAccesoPorIdForoRaw_(idForo); if(!raw)return {ok:false,mensaje:"La institución no está autorizada."};
    const estado=String(raw.hoja.getRange(raw.fila,raw.mapa.ESTADO).getValue()||"").toUpperCase();
    if(estado==="ENVIADO")return {ok:true,yaEnviado:true,mensaje:"Las respuestas ya fueron enviadas definitivamente."};
    if(!sesionActivaPorIdForo_(idForo,dispositivoId,tokenSesion))return {ok:false,mensaje:"La sesión ya no está activa. Ingrese nuevamente con uno de los códigos de la institución."};
    const valida=validarEnvioFinal_(datos); if(!valida.ok)return valida;
    datos.idForo=String(idForo); datos.institucion=raw.ie; datos.dane=raw.dane;
    const guardado=guardarAvanceForo(datos); if(!guardado?.ok)throw new Error(guardado?.mensaje||"No fue posible guardar las respuestas.");
    const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HOJA_AVANCES); const m=obtenerMapaCabeceras_(sh); const row=buscarFilaPorIdForo_(sh,idForo,m); const now=new Date();
    ["S1_ENVIADA","S2_ENVIADA","S3_ENVIADA"].forEach(k=>{if(m[k])sh.getRange(row,m[k]).setValue("SI")});
    ["FECHA_ENVIO_S1","FECHA_ENVIO_S2","FECHA_ENVIO_S3","FECHA_ENVIO_DEFINITIVO"].forEach(k=>{if(m[k])sh.getRange(row,m[k]).setValue(now)});
    if(m.ESTADO)sh.getRange(row,m.ESTADO).setValue("ENVIADO");
    guardarEnHojaIE_(datos); actualizarParticipacion_(datos); actualizarGraficoHojaIE_(datos);
    const shIE=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nombreHojaIE_(datos.institucion)); if(shIE&&shIE.getLastRow()>=2){const mi=obtenerMapaCabeceras_(shIE);const rr=buscarFilaPorIdForo_(shIE,idForo,mi);if(rr>0&&mi.ESTADO)shIE.getRange(rr,mi.ESTADO).setValue("ENVIADO");}
    actualizarGraficosParticipacion_();
    const am=raw.mapa; if(am.ESTADO)raw.hoja.getRange(raw.fila,am.ESTADO).setValue("ENVIADO"); if(am.FECHA_ENVIO)raw.hoja.getRange(raw.fila,am.FECHA_ENVIO).setValue(now);
    return {ok:true,fecha:now.toISOString(),idForo:idForo};
  }finally{try{lock.releaseLock();}catch(e){}}
}
function actualizarGraficosParticipacion_(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID); const sh=ss.getSheetByName(HOJA_PARTICIPACION); if(!sh||sh.getLastRow()<2)return;
  sh.getCharts().forEach(c=>sh.removeChart(c));
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String); const idx={};headers.forEach((h,i)=>idx[h]=i+1);
  const labels=["Rector(a)","Coordinador(a)","Docentes","Tutor PTA PFI/3.0","Orientador(a)","Estudiantes","Padres/madres/acudientes","Personal administrativo","Egresados","Sector productivo","Otros"];
  const totals=labels.map(l=>{const col=idx[l];if(!col)return 0;return sh.getRange(2,col,sh.getLastRow()-1,1).getValues().reduce((s,r)=>s+Number(r[0]||0),0)});
  const start=sh.getLastColumn()+2; sh.getRange(1,start,labels.length,2).setValues(labels.map((l,i)=>[l,totals[i]]));
  const range=sh.getRange(1,start,labels.length,2);
  const chart=sh.newChart().setChartType(Charts.ChartType.PIE).addRange(range).setOption("title","Participación consolidada por estamento — FEM 2026").setPosition(2,start+3,0,0).build(); sh.insertChart(chart);
  const chart2=sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(range).setOption("title","Total consolidado de participantes por estamento — FEM 2026").setPosition(22,start+3,0,0).build(); sh.insertChart(chart2);
  const total=totals.reduce((a,b)=>a+b,0); sh.getRange(labels.length+3,start,2,2).setValues([["TOTAL MUNICIPAL",total],["IE PARTICIPANTES",sh.getLastRow()-1]]);
}

function actualizarGraficoHojaIE_(datos){
  try{
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID); const sh=ss.getSheetByName(nombreHojaIE_(datos.institucion)); if(!sh)return;
    sh.getCharts().forEach(c=>sh.removeChart(c));
    const c=datos.campos||{}; const labels=["Rector(a)","Coordinador(a)","Docentes","Tutor PTA PFI/3.0","Orientador(a)","Estudiantes","Padres/madres/acudientes","Personal administrativo","Egresados","Sector productivo","Otros"];
    const ids=["Rector","Coordinador","Docentes","TutorPTA","Orientador","Estudiantes","Padres","Administrativos","Egresados","Sector","Otros"];
    const start=sh.getLastColumn()+2; const vals=ids.map((id,i)=>[labels[i],Number(c["participantes"+id]?.valor||0)]);
    sh.getRange(1,start,vals.length,2).setValues(vals);
    const chart=sh.newChart().setChartType(Charts.ChartType.PIE).addRange(sh.getRange(1,start,vals.length,2)).setOption("title","Participación — "+(datos.institucion||"IE")).setPosition(1,start+3,0,0).build(); sh.insertChart(chart);
  }catch(e){Logger.log("No fue posible crear gráfico IE: "+e.message);}
}

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

/*
 * El ID original (1ROYRM7hLY3qVQGifSKUQhgHgW8HK1iwB) resultó ser un
 * archivo .xlsx de Microsoft Excel subido a Drive, no una Hoja de
 * cálculo de Google nativa — por eso SpreadsheetApp.openById() nunca
 * podía abrirlo ("El servicio de Hojas de cálculo no ha podido
 * acceder al documento"), sin importar los permisos ni los
 * reintentos. Se reemplaza por el ID de la conversión a Hojas de
 * cálculo de Google (Archivo > Guardar como Hojas de cálculo de
 * Google), que conserva todas las pestañas y datos.
 */
const SPREADSHEET_ID =
  "1OiBPO8BEsa0TpmYGRfEu2I2tMpxIMKAJdr9WtTRd14Y";

/*
 * URL pública fija de la implementación (deployment) de producción.
 *
 * IMPORTANTE: ScriptApp.getService().getUrl() NO sirve para esto.
 * Cuando una función se ejecuta manualmente desde el editor (botón
 * Ejecutar), esa llamada devuelve la URL /dev en vez de la URL /exec
 * del deployment publicado — y /dev exige que quien la abra tenga
 * permiso de edición sobre el script, algo que las instituciones no
 * tienen. Por eso los enlaces personalizados se construyen siempre
 * a partir de esta constante, y no de getService().getUrl().
 *
 * Si se crea un deployment nuevo (con un deploymentId distinto),
 * hay que actualizar este valor y volver a ejecutar
 * actualizarURLsAccesoIE() / generarAccesosIE().
 */
const URL_WEBAPP_PRODUCCION =
  "https://script.google.com/macros/s/AKfycbzeXpV-I2kR-jOetOH_DqaXz0K9QoUfA49ouc5gUNE3rWXhy4fI77EXM4-Y8e08APQy/exec";

const HOJA_OFICIALES =
  "Oficiales";

const HOJA_AVANCES =
  "AvancesForo";

const HOJA_ACCESOS = "AccesosIE";
const HOJA_PARTICIPACION = "Participacion";
const HOJA_ASISTENCIA_QR = "AsistenciaQR";
const DRIVE_CARPETA_FEM_ID = "1IqcFgQUSKocvGX3JwvNOu-xJzt0gfKc8";

// Paleta usada al construir documentos con DocumentApp (informe
// ejecutivo, listado de asistencia). Misma paleta del formulario.
const COLOR_VERDE_DOC = "#0B6A44";
const COLOR_GRIS_TEXTO_DOC = "#4A4A4A";
const COLOR_GRIS_FONDO_DOC = "#F7F8FA";
const COLOR_GRIS_BORDE_DOC = "#DADCE0";
const COLOR_AZUL_CLARO_DOC = "#EAF3FB";
const COLOR_AMARILLO_DOC = "#F4B400";
const COLOR_NEGRO_DOC = "#000000";
// Ya no se usa: generarInformeFEM() construye el informe con
// DocumentApp.create() en vez de copiar este archivo, que resultó
// ser un Word (.docx) y no un Google Doc nativo. Se conserva el
// ID por si se quiere revisar el diseño original de la plantilla.
const TEMPLATE_INFORME_ID = "1Gtsccdbnlcyjl6TcDDjTOA7pAW3JQbHM";
const LOGO_ENCABEZADO_ID = "1mFOOUZ5aFAuwM-JMxNUaDnPPznDlQ2bj";
const LOGO_PIE_ID = "1Cmx7c3ec2gQCjRc8kcNeUbZt5LiURyD5";
// Marco/composición decorativa de la pantalla de acceso. Se referencia
// como URL fija directamente en CSS.html (que se incluye como texto
// plano, sin plantilla) — por eso necesita estar públicamente
// compartido (ver hacerPublicosLogosGlobales() en Pruebas.js).
const MARCO_ACCESO_ID = "1qKHFEoq61uBOn1tNusZxXcK8rutIxDAS";
const REMITENTE_FEM = "calidadeducacion@alcaldianeiva.gov.co";
const COPIAS_INFORME_FEM = [
  "adriana.cedeno@alcaldianeiva.gov.co",
  "angelica.rojas@alcaldianeiva.gov.co",
  "ronald.polania@alcaldianeiva.gov.co"
];

/*
 * "El servicio de Hojas de cálculo no ha podido acceder al
 * documento" es un error transitorio conocido del servicio de
 * Google (no de este código): ocurre sobre todo cuando una misma
 * ejecución abre la hoja muchas veces seguidas (enviarForoDefinitivo
 * la abre más de diez veces). Se reintenta unas pocas veces con
 * espera corta antes de fallar de verdad, en vez de abrir la hoja
 * directamente en cada función.
 */
function abrirSpreadsheet_(){
  let ultimoError = null;
  for(let i = 0; i < 4; i++){
    try{
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }catch(error){
      ultimoError = error;
      Utilities.sleep(500 * (i + 1));
    }
  }
  throw new Error(
    "El servicio de Hojas de cálculo no respondió después de varios intentos: " +
    (ultimoError ? ultimoError.message : "error desconocido")
  );
}


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

  /*
   * =================================================
   * FIRMA DE ASISTENCIA POR CÓDIGO QR
   * =================================================
   *
   * .../exec?asistencia=ID_FORO
   *
   * Página pública mínima, independiente del resto del
   * formulario, para que cada asistente firme desde su
   * propio celular al escanear el QR mostrado en la
   * pantalla de Evidencias.
   */
  const idForoAsistencia =
    String(parametros.asistencia || "").trim();

  if (idForoAsistencia !== "") {
    return paginaAsistenciaQR_(idForoAsistencia);
  }

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
  plantilla.logoIEUrl = "";
  plantilla.tituloHeaderIE = "FORO EDUCATIVO INSTITUCIONAL";
  plantilla.urlWebapp = URL_WEBAPP_PRODUCCION;
  // Logos institucionales del encabezado (FEM a la izquierda, SEM a
  // la derecha). Antes se intentaban cargar desde variables globales
  // del cliente (LOGO_FEM/LOGO_SEM) que nunca llegaron a declararse
  // en ningún archivo — por eso aparecían como imagen rota en todas
  // las pantallas. Se resuelven aquí, en el servidor, igual que el
  // logo de la IE.
  plantilla.logoFemUrl = urlPublicaLogoDrive_(LOGO_ENCABEZADO_ID);
  plantilla.logoSemUrl = urlPublicaLogoDrive_(LOGO_PIE_ID);

  /*
   * El servidor identifica la IE a partir del TOKEN.
   * El código de acceso se valida posteriormente.
   */
  if (token !== "") {

    try {

      const ss =
        abrirSpreadsheet_();

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

          const colLogo =
            cabeceras.indexOf("LOGO_ID");

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

                plantilla.tituloHeaderIE =
                  "FORO EDUCATIVO INSTITUCIÓN EDUCATIVA " +
                  nombreIESinPrefijoInstitucional_(plantilla.nombreIEAcceso).toUpperCase();

                if (colLogo !== -1) {

                  const logoId =
                    String(
                      datos[i][colLogo] || ""
                    ).trim();

                  if (logoId) {
                    plantilla.logoIEUrl = urlPublicaLogoDrive_(logoId);
                  }

                }

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
            abrirSpreadsheet_();

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

  const ss =
    abrirSpreadsheet_();

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
    abrirSpreadsheet_();


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
    abrirSpreadsheet_();

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

    /*
     * ENVIADO ya NO impide seguir usando este ID_FORO: una vez
     * enviado el formulario debe poder seguir consultándose, editando
     * y regenerando el informe. Solo un bloqueo manual explícito
     * (BLOQUEADO/INACTIVO, que pone la SEM a mano en AccesosIE)
     * impide el acceso.
     */
    if (
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
      abrirSpreadsheet_();

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

      /*
       * ENVIADO ya NO impide seguir consultando el avance guardado:
       * un formulario ya enviado debe poder seguir restaurándose
       * (por ejemplo, al reingresar desde otro dispositivo) en vez
       * de reportarse como "bloqueado". Solo un bloqueo manual
       * explícito en AccesosIE (BLOQUEADO/INACTIVO) impide seguir.
       */
      if (
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
      abrirSpreadsheet_();


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

/*
 * Ya NO hay temporizador de inactividad: una sesión reclamada por un
 * dispositivo permanece activa indefinidamente (sin importar cuánto
 * tiempo pase sin actividad) hasta que:
 *   a) ese mismo dispositivo la libera (liberarSesionCodigo_), o
 *   b) otro dispositivo la toma explícitamente con "forzar:true",
 *      tras la confirmación del usuario en pantalla.
 * "Solo un envío y una conexión es posible por IE."
 */

function obtenerClaveSesionCodigo_(token, codigo, idForo) {
  const claveBase = String(idForo || "").trim() || (String(token || "").trim()+"|"+String(codigo || "").trim());
  return "FEM_SESION_FORO_" + Utilities.base64EncodeWebSafe(claveBase);
}

function reclamarSesionCodigo_(token, codigo, dispositivoId, idForo, forzar) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const props = PropertiesService.getScriptProperties();
    const clave = obtenerClaveSesionCodigo_(token,codigo,idForo);
    const ahora = Date.now();
    let actual = null;
    const guardado = props.getProperty(clave);
    if (guardado) { try { actual=JSON.parse(guardado); } catch(e){ actual=null; } }
    if (actual && actual.deviceId && actual.deviceId !== dispositivoId && !forzar) {
      return {
        ok:false,
        codigo:"SESION_YA_ABIERTA",
        mensaje:"Ya hay una sesión activa para esta IE. Si desea continuar en este dispositivo, se cerrará la conexión del otro dispositivo y podrá seguir en este dispositivo. Solo un envío y una conexión es posible por IE."
      };
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
    // Si otro dispositivo ya tomó la sesión (takeover), se lo informamos
    // al cliente para que deje de intentar seguir trabajando en silencio.
    if(actual.deviceId!==dispositivoId || actual.tokenSesion!==tokenSesion)return {ok:false,codigo:"SESION_NO_AUTORIZADA",mensaje:"Otro dispositivo tomó el control de esta sesión."};
    actual.ultimaActividad=Date.now(); props.setProperty(clave,JSON.stringify(actual)); return {ok:true};
  }catch(e){return {ok:false,codigo:"HEARTBEAT_ERROR"};} finally{try{lock.releaseLock();}catch(e){}}
}

function liberarSesionCodigo(token,codigo,dispositivoId,tokenSesion,idForo){return liberarSesionCodigo_(token,codigo,dispositivoId,tokenSesion,idForo);}
function liberarSesionCodigo_(token,codigo,dispositivoId,tokenSesion,idForo){
  const lock=LockService.getScriptLock();
  try{lock.waitLock(10000);const props=PropertiesService.getScriptProperties();const clave=obtenerClaveSesionCodigo_(token,codigo,idForo);const guardado=props.getProperty(clave);if(!guardado)return {ok:true};const actual=JSON.parse(guardado);if(actual.deviceId===dispositivoId&&actual.tokenSesion===tokenSesion)props.deleteProperty(clave);return {ok:true};}
  catch(e){return {ok:false};} finally{try{lock.releaseLock();}catch(e){}}
}

function validarAccesoIE(token, codigo, dispositivoId, forzar) {

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
      abrirSpreadsheet_();


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
     * DISPONIBLE / ENVIADO = permitido. Un formulario ya enviado
     * debe poder seguir consultándose y reingresándose (por ejemplo,
     * para revisar respuestas, volver a descargar el informe o
     * responder la valoración) — enviar definitivamente ya NO cierra
     * el acceso al código.
     *
     * BLOQUEADO / INACTIVO = bloqueo manual (lo pone la SEM a mano
     * en AccesosIE) — es el único caso que sigue impidiendo el
     * ingreso.
     */

    if (
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
       * TIPO = PRUEBA: fila de prueba controlada a mano en
       * AccesosIE (no está en Oficiales). Antes esto solo aceptaba
       * una única IE de prueba hardcodeada ("IE PRUEBA 1234", DANE
       * "111", código "1234", ID_FORO "PRUEBA-1234"), así que
       * cualquier otra fila marcada TIPO = PRUEBA quedaba rechazada
       * con "El registro de prueba no tiene la configuración
       * esperada." — impedía tener más de una IE de prueba a la
       * vez. Ya no hace falta esa restricción: el TOKEN y el
       * CODIGO_ACCESO/CODIGO_CONTINGENCIA_x ya se validaron arriba
       * contra esta misma fila, así que cualquier fila marcada
       * TIPO = PRUEBA es igual de confiable que una IE encontrada en
       * Oficiales — es el mismo control (solo alguien con acceso a
       * la hoja puede crearla), solo que en otra hoja.
       */
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
     * Sin temporizador de inactividad: la sesión no se libera sola
     * por tiempo. Si otro dispositivo ya tiene la sesión abierta, se
     * devuelve SESION_YA_ABIERTA para que el cliente muestre la
     * confirmación de takeover; si el usuario confirma, esta misma
     * función se vuelve a llamar con forzar=true.
     */

    const resultadoSesion =
      reclamarSesionCodigo_(token, codigo, dispositivoId, idForo, !!forzar);

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
    abrirSpreadsheet_();

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
    REMITENTE_FEM;

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

    const ieSinPrefijo =
      nombreIESinPrefijoInstitucional_(ie);

    const logoIEUrlCorreo =
      mapa["LOGO_ID"] !== undefined
        ? urlPublicaLogoDrive_(String(fila[mapa["LOGO_ID"]] || "").trim())
        : "";

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
        "🎓 Acceso al Foro Educativo Institucional – " +
        ie;

      const textoEnlace =
        "Ingreso de la IE " + ieSinPrefijo + " al Foro Educativo Institucional";

      const cuerpoTexto =
        "Secretaría de Educación de Neiva\n\n" +
        "Estimada comunidad educativa de la Institución Educativa " + ieSinPrefijo + ":\n\n" +
        "Ya pueden ingresar al Foro Educativo Institucional – Neiva 2026 con el código de acceso exclusivo de su institución.\n\n" +
        "Código de acceso: " + codigo + "\n\n" +
        textoEnlace + ":\n" + url + "\n\n" +
        "Este código y este enlace son exclusivos de su institución: no deben compartirse con otra IE.\n\n" +
        "Ante cualquier inconveniente pueden escribir a este mismo correo o comunicarse al WhatsApp 318 456 1081.\n\n" +
        "Secretaría de Educación de Neiva\n" +
        "Foro Educativo Institucional – Neiva 2026\n" +
        "“Escuela Viva: Voces que construyen territorio”";

      /*
       * HTML del correo con el mismo lenguaje visual del formulario
       * (verde institucional, acento amarillo, tarjeta redondeada con
       * sombra) — con estilos en línea para que se vea igual en la
       * mayoría de los clientes de correo, que no siempre respetan
       * bloques <style>.
       */
      const cuerpoHTML =
        "<div style=\"background:#F7F8FA;padding:28px 12px;font-family:Arial,Helvetica,sans-serif;\">" +
        "<div style=\"max-width:520px;margin:0 auto;background:#FFFFFF;border-radius:16px;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,.10);\">" +
        "<div style=\"background:#0B6A44;padding:26px 28px;text-align:center;\">" +
        (logoIEUrlCorreo ? "<img src=\""+logoIEUrlCorreo+"\" alt=\"Logo de la institución educativa\" style=\"display:block;max-width:56px;max-height:56px;margin:0 auto 10px;border-radius:8px;\">" : "") +
        "<div style=\"color:#FFFFFF;font-size:20px;font-weight:700;\">Foro Educativo Institucional</div>" +
        "<div style=\"color:#CFE8DC;font-size:14px;margin-top:2px;\">Neiva 2026</div>" +
        "</div>" +
        "<div style=\"padding:28px 28px 8px;\">" +
        "<p style=\"font-size:16px;color:#333333;margin:0 0 14px;\">Estimada comunidad educativa de la Institución Educativa <strong>" + ieSinPrefijo + "</strong>:</p>" +
        "<p style=\"font-size:15px;color:#4A4A4A;line-height:1.6;margin:0 0 22px;\">" +
        "Ya pueden ingresar al Foro Educativo Institucional – Neiva 2026 con el código de acceso exclusivo de su institución." +
        "</p>" +
        "<div style=\"background:#F7F8FA;border-left:6px solid #F4B400;border-radius:10px;padding:16px 20px;margin:0 0 24px;text-align:center;\">" +
        "<div style=\"font-size:12px;font-weight:700;color:#0B6A44;text-transform:uppercase;letter-spacing:.5px;\">Código de acceso</div>" +
        "<div style=\"font-size:30px;font-weight:700;letter-spacing:6px;color:#0B6A44;margin-top:4px;\">" + codigo + "</div>" +
        "</div>" +
        "<div style=\"text-align:center;margin:0 0 24px;\">" +
        "<a href=\"" + url + "\" target=\"_blank\" style=\"display:inline-block;background:#0B6A44;color:#FFFFFF;text-decoration:none;font-weight:700;font-size:15px;padding:14px 26px;border-radius:10px;\">" + textoEnlace + "</a>" +
        "</div>" +
        "<div style=\"background:#FFF8E1;border-left:6px solid #F4B400;border-radius:10px;padding:12px 16px;margin:0 0 20px;\">" +
        "<p style=\"font-size:13px;color:#7A5B00;margin:0;\">🔒 Este código y este enlace son exclusivos de su institución: no deben compartirse con otra IE.</p>" +
        "</div>" +
        "<p style=\"font-size:13px;color:#888888;margin:0 0 24px;\">Ante cualquier inconveniente, pueden escribir a este mismo correo o comunicarse al WhatsApp 318 456 1081.</p>" +
        "</div>" +
        "<div style=\"background:#F7F8FA;padding:18px 28px;text-align:center;border-top:1px solid #E5E7EA;\">" +
        "<p style=\"font-size:13px;color:#0B6A44;font-weight:700;margin:0;\">Secretaría de Educación de Neiva</p>" +
        "<p style=\"font-size:12px;color:#888888;margin:4px 0 0;font-style:italic;\">“Escuela Viva: Voces que construyen territorio”</p>" +
        "</div>" +
        "</div>" +
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
 * ACTUALIZAR URL DE ACCESO DE LAS IE
 *
 * NO genera nuevos TOKEN.
 * NO genera nuevos códigos.
 * NO modifica ID_FORO.
 *
 * Reconstruye URL_ACCESO y LINK_ACCESO para los
 * registros que ya existen en AccesosIE.
 *****************************************************/
function actualizarURLsAccesoIE() {

  const ss = abrirSpreadsheet_();

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
   * URL REAL DEL WEB APP (deployment de producción, no /dev).
   */
  const urlBase =
    URL_WEBAPP_PRODUCCION;

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
            abrirSpreadsheet_();

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
         * 7. URL BASE (deployment de producción, no /dev)
         * =================================================
         */

        const urlBase =
            URL_WEBAPP_PRODUCCION;

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
  const ss=abrirSpreadsheet_();
  let hoja=ss.getSheetByName(HOJA_ACCESOS);
  if(!hoja) hoja=ss.insertSheet(HOJA_ACCESOS);
  const requeridas=["ID_ACCESO","IE","DANE","CODIGO_ACCESO","TOKEN","URL_ACCESO","ID_FORO","ESTADO","TOKEN_SESION","DISPOSITIVO_ID","FECHA_GENERACION","FECHA_PRIMER_ACCESO","ULTIMA_ACTIVIDAD","FECHA_ENVIO","EMAIL_IE","EMAIL_RESPONSABLE","TIPO","S1_ENVIADA","S2_ENVIADA","S3_ENVIADA","ID_INFORME","ID_PDF_INFORME","LOGO_ID"];
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
    const ss=abrirSpreadsheet_();
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
  const ss=abrirSpreadsheet_(); let sh=ss.getSheetByName(HOJA_PARTICIPACION); if(!sh)sh=ss.insertSheet(HOJA_PARTICIPACION);
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
  const ss=abrirSpreadsheet_(); const shName=nombreHojaIE_(datos.institucion); let sh=ss.getSheetByName(shName); if(!sh){sh=ss.insertSheet(shName); const headers=obtenerCabecerasAvancesForo(); sh.getRange(1,1,1,headers.length).setValues([headers]);}
  if(!sh)return; const headers=obtenerCabecerasAvancesForo(); if(sh.getLastColumn()<headers.length)sh.getRange(1,1,1,headers.length).setValues([headers]); const m=mapaHoja_(sh); const respuestas=extraerRespuestasSesiones_(datos);
  const row=Object.assign({ID_FORO:datos.idForo,INSTITUCION:datos.institucion,DANE:datos.dane,FECHA_INICIO:datos.fechaInicio||new Date(),ULTIMA_ACTUALIZACION:new Date(),ESTADO:"En proceso"},respuestas,{DATOS:JSON.stringify(datos)});
  const out=new Array(sh.getLastColumn()).fill(""); Object.keys(row).forEach(k=>{if(m[k])out[m[k]-1]=normalizarValorHoja_(row[k]);}); let found=-1;if(sh.getLastRow()>=2){const ids=sh.getRange(2,m.ID_FORO,sh.getLastRow()-1,1).getValues();for(let i=0;i<ids.length;i++)if(String(ids[i][0]||"")===String(datos.idForo||"")){found=i+2;break;}} if(found>0)sh.getRange(found,1,1,out.length).setValues([out]);else sh.appendRow(out);
}


function obtenerEstadoSesiones_(idForo){
  const sh=abrirSpreadsheet_().getSheetByName(HOJA_AVANCES); if(!sh||sh.getLastRow()<2)return {s1:false,s2:false,s3:false}; const m=mapaHoja_(sh); const row=buscarFilaPorIdForo_(sh,idForo,m); if(row<0)return {s1:false,s2:false,s3:false}; return {s1:String(sh.getRange(row,m.S1_ENVIADA).getValue()).toUpperCase()==="SI",s2:String(sh.getRange(row,m.S2_ENVIADA).getValue()).toUpperCase()==="SI",s3:String(sh.getRange(row,m.S3_ENVIADA).getValue()).toUpperCase()==="SI"};
}


function obtenerDatosGuardadosPorIdForo_(idForo){
  const sh=abrirSpreadsheet_().getSheetByName(HOJA_AVANCES); if(!sh||sh.getLastRow()<2)return null; const m=mapaHoja_(sh); const row=buscarFilaPorIdForo_(sh,idForo,m); if(row<0)return null; const raw=sh.getRange(row,m.DATOS).getValue(); if(!raw)return null; try{return JSON.parse(raw);}catch(e){return null;}
}


function crearCarpetaIE_(ie){
  const root=DriveApp.getFolderById(DRIVE_CARPETA_FEM_ID); const it=root.getFoldersByName(ie); if(it.hasNext())return it.next(); return root.createFolder(ie);
}


function hacerPublicoSiEsPosible_(file){try{file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(e){Logger.log("No se pudo cambiar compartir: "+e.message);}}

/*
 * URL de imagen directamente visible (sin necesidad de sesión de
 * Google) para un archivo de Drive ya compartido como "cualquiera
 * con el enlace". Se usa para los logos por IE en el encabezado de
 * la página, el favicon, los correos y el informe.
 */
function urlPublicaLogoDrive_(fileId){
  if(!fileId) return "";
  return "https://lh3.googleusercontent.com/d/"+fileId+"=w300";
}

/*
 * Busca en AccesosIE la columna LOGO_ID de una IE por su nombre
 * exacto (tal como está guardado en esa hoja). Devuelve el ID del
 * archivo de Drive, o "" si esa IE no tiene logo vinculado todavía.
 */
function obtenerLogoIdPorNombreIE_(nombreIE){
  const nombre=String(nombreIE||"").trim(); if(!nombre) return "";
  const hoja=asegurarColumnasAccesosIE_();
  const m=mapaHoja_(hoja);
  if(!m.IE || !m.LOGO_ID) return "";
  const ultimaFila=hoja.getLastRow(); if(ultimaFila<2) return "";
  const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();
  for(let i=0;i<valores.length;i++){
    if(String(valores[i][m.IE-1]||"").trim()===nombre){
      return String(valores[i][m.LOGO_ID-1]||"").trim();
    }
  }
  return "";
}


/*****************************************************
 * ASISTENCIA POR CÓDIGO QR
 *
 * Reemplaza la subida manual de un PDF de asistencia:
 * cada participante firma desde su propio celular al
 * escanear el QR mostrado en la pantalla de Evidencias,
 * completando IE, nombre completo, cargo y documento.
 *****************************************************/

// Mismas categorías de la caracterización, en singular — para el
// desplegable de "Cargo" que llena cada asistente al firmar por QR.
const CARGOS_ASISTENCIA_QR=["Rector(a)","Coordinador(a)","Docente","Tutor(a) PTA/PFI 3.0","Orientador(a)","Estudiante","Padre/madre/acudiente","Personal administrativo","Egresado(a)","Sector productivo","Otro"];
const TIPOS_ASISTENCIA_QR=["Presencial","Virtual","No asistió: con permiso institucional o incapacidad médica.","No asistió: con permiso de comisión o con acto administrativo."];
const ROLES_FORO_QR=["👑 Líder – Rector(a)","🎓 Dinamizador Pedagógico – Tutor(a) PTA / PFI 3.0","👥 Dinamizador(a) de Mesas de Trabajo","📝 Relator(a)","⏱️ Dinamizador(a) del Tiempo","💻 Dinamizador(a) de la Sistematización","🙋 Participante"];

function asegurarHojaAsistenciaQR_(){
  const ss=abrirSpreadsheet_();
  let hoja=ss.getSheetByName(HOJA_ASISTENCIA_QR);
  if(!hoja) hoja=ss.insertSheet(HOJA_ASISTENCIA_QR);
  const requeridas=["ID_FORO","IE","NOMBRE_COMPLETO","TIPO_ASISTENCIA","CARGO","ROL_FORO","NUMERO_DOCUMENTO","CORREO","TELEFONO","FECHA","HORA","DISPOSITIVO_ID"];
  const last=hoja.getLastColumn();
  const existentes=last?hoja.getRange(1,1,1,last).getValues()[0].map(String):[];
  if(!last){ hoja.getRange(1,1,1,requeridas.length).setValues([requeridas]); }
  else{
    const faltantes=requeridas.filter(h=>existentes.indexOf(h)===-1);
    if(faltantes.length) hoja.getRange(1,last+1,1,faltantes.length).setValues([faltantes]);
  }
  return hoja;
}

/*
 * Se llama desde la página pública del QR (paginaAsistenciaQR_).
 * No exige sesión ni dispositivo: los asistentes firman desde su
 * propio celular, distinto al que usa quien diligencia el formulario.
 */
/*
 * Texto de confirmación de firma en letras: "Firmado a las 2:32 p. m.
 * el día 26 del mes de agosto del año 2026."
 */
function formatearFechaHoraFirma_(fecha){
  const zona=Session.getScriptTimeZone();
  const meses=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  const dia=Utilities.formatDate(fecha,zona,"dd");
  const mesIndex=Number(Utilities.formatDate(fecha,zona,"M"))-1;
  const anio=Utilities.formatDate(fecha,zona,"yyyy");
  const horas24=Number(Utilities.formatDate(fecha,zona,"H"));
  const minutos=Utilities.formatDate(fecha,zona,"mm");
  const sufijo=horas24>=12?"p. m.":"a. m.";
  let horas12=horas24%12; if(horas12===0)horas12=12;
  return "Firmado a las "+horas12+":"+minutos+" "+sufijo+" el día "+dia+" del mes de "+(meses[mesIndex]||"")+" del año "+anio+".";
}

function registrarAsistenciaQR(idForo, nombre, tipoAsistencia, cargo, rolForo, documento, correo, telefono, dispositivoId){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);

    idForo=String(idForo||"").trim();
    nombre=String(nombre||"").trim();
    tipoAsistencia=String(tipoAsistencia||"").trim();
    cargo=String(cargo||"").trim();
    rolForo=String(rolForo||"").trim();
    documento=String(documento||"").trim();
    correo=String(correo||"").trim();
    telefono=String(telefono||"").trim();
    dispositivoId=String(dispositivoId||"").trim();

    if(!idForo) return {ok:false, mensaje:"Enlace de asistencia inválido."};
    if(!nombre || !tipoAsistencia || !cargo || !rolForo || !documento || !correo) return {ok:false, mensaje:"Complete nombre, tipo de asistencia, cargo, rol en el Foro, número de documento y correo electrónico."};
    if(TIPOS_ASISTENCIA_QR.indexOf(tipoAsistencia)===-1) return {ok:false, mensaje:"Seleccione un tipo de asistencia válido."};
    if(ROLES_FORO_QR.indexOf(rolForo)===-1) return {ok:false, mensaje:"Seleccione un rol válido en el Foro Educativo Institucional."};
    // Validaciones de formato: el documento solo acepta dígitos, y el
    // correo debe tener una forma mínimamente válida (no solo estar
    // presente). El teléfono es opcional, pero si se escribe algo
    // también debe ser solo dígitos.
    if(!/^[0-9]+$/.test(documento)) return {ok:false, mensaje:"El número de documento debe contener solo números."};
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) return {ok:false, mensaje:"Ingrese un correo electrónico válido."};
    if(telefono && !/^[0-9]+$/.test(telefono)) return {ok:false, mensaje:"El teléfono debe contener solo números."};

    const acceso=obtenerAccesoPorIdForo_(idForo);
    if(!acceso) return {ok:false, mensaje:"Este código de asistencia ya no está disponible."};

    const hoja=asegurarHojaAsistenciaQR_();
    const m=mapaHoja_(hoja);

    const textoFirmaDesdeFila_=function(fila){
      const fechaTexto=String(fila[m.FECHA-1]||"").trim();
      const horaTexto=String(fila[m.HORA-1]||"").trim();
      if(!fechaTexto || !horaTexto) return "";
      const partes=fechaTexto.split("/");
      if(partes.length!==3) return "";
      return formatearFechaHoraFirma_(new Date(Number(partes[2]),Number(partes[1])-1,Number(partes[0]),...horaTexto.split(":").map(Number)));
    };

    // Evitar duplicados: misma persona (documento) o mismo dispositivo
    // firmando dos veces para este mismo foro. Solo una firma por
    // dispositivo, sin importar qué datos escriba la segunda vez.
    const ultimaFila=hoja.getLastRow();
    if(ultimaFila>=2){
      const filas=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getValues();
      for(let i=0;i<filas.length;i++){
        const mismoForo=String(filas[i][m.ID_FORO-1]||"").trim()===idForo;
        if(!mismoForo) continue;
        if(String(filas[i][m.NUMERO_DOCUMENTO-1]||"").trim()===documento){
          return {ok:true, yaRegistrado:true, textoFirma:textoFirmaDesdeFila_(filas[i])};
        }
        if(dispositivoId && m.DISPOSITIVO_ID && String(filas[i][m.DISPOSITIVO_ID-1]||"").trim()===dispositivoId){
          return {ok:false, yaFirmoDispositivo:true, mensaje:"Este dispositivo ya registró una firma de asistencia para este Foro. Solo se permite una firma por dispositivo.", textoFirma:textoFirmaDesdeFila_(filas[i])};
        }
      }
    }

    const ahora=new Date();
    const zona=Session.getScriptTimeZone();
    const fila=new Array(hoja.getLastColumn()).fill("");
    const valores={ID_FORO:idForo, IE:acceso.ie, NOMBRE_COMPLETO:nombre, TIPO_ASISTENCIA:tipoAsistencia, CARGO:cargo, ROL_FORO:rolForo, NUMERO_DOCUMENTO:documento, CORREO:correo, TELEFONO:telefono, DISPOSITIVO_ID:dispositivoId,
      FECHA:Utilities.formatDate(ahora,zona,"dd/MM/yyyy"), HORA:Utilities.formatDate(ahora,zona,"HH:mm:ss")};
    Object.keys(valores).forEach(k=>{ if(m[k]) fila[m[k]-1]=valores[k]; });
    hoja.appendRow(fila);
    return {ok:true, textoFirma:formatearFechaHoraFirma_(ahora)};

  }catch(error){
    return {ok:false, mensaje:error.message};
  }finally{
    try{ lock.releaseLock(); }catch(e){}
  }
}

function obtenerAsistentesQR_(idForo){
  const hoja=asegurarHojaAsistenciaQR_();
  const m=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  if(ultimaFila<2) return [];
  const filas=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();
  return filas
    .filter(fila=>String(fila[m.ID_FORO-1]||"").trim()===String(idForo||"").trim())
    .map(fila=>({
      nombre:String(fila[m.NOMBRE_COMPLETO-1]||""),
      tipoAsistencia:String(fila[m.TIPO_ASISTENCIA-1]||""),
      cargo:String(fila[m.CARGO-1]||""),
      rolForo:String(fila[m.ROL_FORO-1]||""),
      documento:String(fila[m.NUMERO_DOCUMENTO-1]||""),
      correo:String(fila[m.CORREO-1]||""),
      telefono:String(fila[m.TELEFONO-1]||""),
      fecha:String(fila[m.FECHA-1]||""),
      hora:String(fila[m.HORA-1]||"")
    }));
}

/*
 * Se llama desde la pantalla de Evidencias para mostrar cuántas
 * personas ya firmaron mientras el QR sigue disponible. La
 * asistencia se sigue firmando durante toda la jornada, así que
 * esto NO bloquea continuar a la plenaria — solo informa.
 */
function contarAsistentesQR(idForo){
  try{
    const datos=obtenerDatosGuardadosPorIdForo_(idForo);
    const estado=obtenerEstadoSesiones_(idForo);
    return {
      ok:true,
      total:obtenerAsistentesQR_(idForo).length,
      totalCaracterizacion:datos?totalParticipantesServer_(datos):0,
      cerrado:!!(estado&&estado.s3)
    };
  }catch(error){
    return {ok:false, mensaje:error.message};
  }
}

/*
 * Agrega al informe ejecutivo, después de la fotografía, el
 * listado de quienes firmaron asistencia por código QR — ya no se
 * genera como un PDF aparte, va incluido al final del informe.
 */
function agregarListadoAsistenciaAlInforme_(body, idForo, datos){
  const asistentes=obtenerAsistentesQR_(idForo);

  const titulo=body.appendParagraph("Listado de asistencia (firmado por código QR)");
  titulo.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titulo.editAsText().setForegroundColor(COLOR_VERDE_DOC).setBold(true);

  if(!asistentes.length){
    const vacio=body.appendParagraph("No se registraron firmas de asistencia por código QR.");
    vacio.editAsText().setForegroundColor(COLOR_GRIS_TEXTO_DOC);
    return;
  }

  const t=body.appendTable();
  t.setBorderColor(COLOR_GRIS_BORDE_DOC); t.setBorderWidth(1);
  const encabezado=t.appendTableRow();
  ["Nombre completo","Asistencia","Cargo","Rol en el Foro","N.° documento","Correo","Teléfono","Fecha","Hora"].forEach(function(texto){
    const celda=encabezado.appendTableCell(texto);
    celda.setBackgroundColor(COLOR_VERDE_DOC);
    celda.editAsText().setForegroundColor("#FFFFFF").setBold(true).setFontSize(9);
  });
  asistentes.forEach(function(a){
    const r=t.appendTableRow();
    [a.nombre,a.tipoAsistencia,a.cargo,a.rolForo,a.documento,a.correo,a.telefono,a.fecha,a.hora].forEach(function(valor){
      r.appendTableCell(String(valor||"—")).editAsText().setForegroundColor(COLOR_GRIS_TEXTO_DOC).setFontSize(9);
    });
  });

  const totalCaracterizacion=totalParticipantesServer_(datos);
  const resumen=body.appendParagraph(
    "Total de participantes según firmas de asistencia (QR): "+asistentes.length+"\n"+
    "Total de participantes según caracterización institucional: "+totalCaracterizacion
  );
  resumen.editAsText().setForegroundColor(COLOR_GRIS_TEXTO_DOC).setBold(true);
}

/*
 * "COLEGIO CHAPINERO" -> "Colegio Chapinero": nombre con mayúscula
 * inicial en cada palabra, en vez de las mayúsculas sostenidas con
 * que se guarda en la hoja de Oficiales.
 */
function capitalizarNombreIE_(nombre){
  return String(nombre||"").toLowerCase().replace(/(^|\s)([a-záéíóúñ])/g,function(_,sep,letra){ return sep+letra.toUpperCase(); });
}

/*
 * Muchas IE ya guardan su nombre con el prefijo incluido ("INSTITUCIÓN
 * EDUCATIVA JUAN XXIII"). Cuando un texto ya dice "Institución
 * Educativa" antes del nombre, usar el nombre tal cual duplica esa
 * frase ("Institución Educativa INSTITUCIÓN EDUCATIVA JUAN XXIII").
 * Esta función quita ese prefijo redundante solo para esos casos —
 * en el resto de usos (títulos, tablas) el nombre se deja igual.
 */
function nombreIESinPrefijoInstitucional_(ie){
  return String(ie||"").replace(/^\s*(instituci[oó]n\s+educativa|i\.?\s*e\.?)\s+/i,"").trim();
}

/*
 * Página pública mínima que abre el QR: no forma parte del
 * formulario principal (App.html), es autocontenida a propósito
 * para que cargue rápido en el celular de cada asistente.
 */
function paginaAsistenciaQR_(idForo){
  const acceso=obtenerAccesoPorIdForo_(idForo);
  const ie=acceso?acceso.ie:"";
  const ieTitulo=capitalizarNombreIE_(ie);
  const logoUrlIE=urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(ie));

  /*
   * El registro de asistencia por QR queda disponible de forma
   * permanente: no se cierra al enviar la Sesión 3. El informe
   * ejecutivo toma el listado de firmas en el momento en que se
   * genera, pero la página de firma sigue abierta después de eso
   * por si llegan más asistentes.
   */

  const opcionesCargo=CARGOS_ASISTENCIA_QR.map(function(c){
    return '<option value="'+c.replace(/"/g,"&quot;")+'">'+c+'</option>';
  }).join("");

  const opcionesTipoAsistencia=TIPOS_ASISTENCIA_QR.map(function(t){
    return '<option value="'+t.replace(/"/g,"&quot;")+'">'+t+'</option>';
  }).join("");

  const opcionesRolForo=ROLES_FORO_QR.map(function(r){
    return '<option value="'+r.replace(/"/g,"&quot;")+'">'+r+'</option>';
  }).join("");

  const tituloPagina="Firmar asistencia al Foro Educativo Institucional "+ieTitulo;

  const html=
    '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">'+
    '<meta name="viewport" content="width=device-width, initial-scale=1">'+
    '<base target="_top">'+
    '<title>'+tituloPagina.replace(/</g,"&lt;")+' — FEM 2026</title>'+
    '<style>'+
    'body{font-family:Arial,Helvetica,sans-serif;background:#F7F8FA;color:#4A4A4A;margin:0;padding:24px;}'+
    '.tarjeta{max-width:420px;margin:0 auto;background:#fff;border-radius:16px;padding:28px;box-shadow:0 8px 24px rgba(0,0,0,.12);}'+
    'h1{color:#0B6A44;font-size:20px;margin:0 0 6px;line-height:1.35;}'+
    'p{line-height:1.5;}'+
    'label{display:block;font-weight:bold;color:#0B6A44;margin:16px 0 6px;}'+
    'input,select{width:100%;padding:12px;font-size:16px;border:1px solid #DADCE0;border-radius:8px;box-sizing:border-box;font-family:inherit;}'+
    'button{width:100%;margin-top:22px;padding:14px;font-size:17px;background:#0B6A44;color:#fff;border:none;border-radius:10px;cursor:pointer;}'+
    'button:disabled{background:#bdbdbd;}'+
    '#estado{margin-top:14px;font-weight:600;min-height:20px;}'+
    '#textoFirma{margin-top:6px;font-size:12px;font-weight:400;color:#4A4A4A;}'+
    '.logoAsistenciaIE{display:block;max-width:64px;max-height:64px;margin:0 auto 10px;border-radius:8px;}'+
    '.correoInvalido{border-color:#C62828 !important;background:#FFFDE7;}'+
    '.mensajeErrorCorreo{display:none;margin-top:6px;}'+
    '.mensajeErrorCorreo b{background:#FFF3CD;color:#C62828;font-weight:600;padding:3px 8px;border-radius:6px;display:inline-block;}'+
    '.mensajeErrorCorreo i{font-style:normal;font-size:11px;color:#555;margin-left:6px;}'+
    '</style></head><body>'+
    '<div class="tarjeta">'+
    (logoUrlIE?'<img src="'+logoUrlIE+'" alt="Logo de la institución educativa" class="logoAsistenciaIE">':'')+
    '<h1>'+tituloPagina.replace(/</g,"&lt;")+'</h1>'+
    '<p>Foro Educativo Institucional — Neiva 2026</p>'+
    '<div id="formulario">'+
    '<label>Institución Educativa</label>'+
    '<input id="ie" value="'+String(ie).replace(/"/g,"&quot;")+'" readonly>'+
    '<label>Nombre completo</label>'+
    '<input id="nombre" autocomplete="name">'+
    '<label>Su asistencia fue</label>'+
    '<select id="tipoAsistencia"><option value="">Seleccione…</option>'+opcionesTipoAsistencia+'</select>'+
    '<label>Cargo</label>'+
    '<select id="cargo"><option value="">Seleccione…</option>'+opcionesCargo+'</select>'+
    '<label>Rol que desempeñó en el Foro Educativo Institucional '+ieTitulo.replace(/</g,"&lt;")+'</label>'+
    '<select id="rolForo"><option value="">Seleccione…</option>'+opcionesRolForo+'</select>'+
    '<label>Número de documento</label>'+
    '<input id="documento" inputmode="numeric" autocomplete="off">'+
    '<label>Correo electrónico</label>'+
    '<input id="correo" type="email" autocomplete="email">'+
    '<div class="mensajeErrorCorreo" id="mensajeErrorCorreo"><b>Ingrese un correo electrónico válido</b><i>(ej: nombre@dominio.com — sin espacios al inicio, al final o en medio)</i></div>'+
    '<label>Teléfono (opcional)</label>'+
    '<input id="telefono" type="tel" autocomplete="tel">'+
    '<button id="btnFirmar" type="button">Firmar asistencia</button>'+
    '</div>'+
    '<div id="estado"></div>'+
    '<div id="textoFirma"></div>'+
    '</div>'+
    '<script>'+
    /*
     * Un dispositivo solo puede firmar una vez por Foro. Esta página
     * la sirve Apps Script desde un subdominio de
     * script.googleusercontent.com que puede cambiar entre una
     * visita y otra (por ejemplo, al volver a escanear el QR), así
     * que localStorage NO es confiable como único mecanismo — puede
     * quedar "vacío" en cada visita aunque sea el mismo teléfono. Por
     * eso el identificador de dispositivo se calcula a partir de
     * características bastante estables del navegador/equipo
     * (user agent, idioma, resolución de pantalla, zona horaria,
     * etc.), no de un valor aleatorio guardado: así, aunque cambie el
     * origen, dos visitas desde el mismo equipo calculan el MISMO
     * identificador y el servidor puede bloquear el segundo intento.
     * localStorage se sigue usando solo como atajo de UI (si
     * persiste, evita ni siquiera mostrar el formulario de nuevo).
     */
    'function calcularHuellaDispositivo(){'+
    'try{'+
    'var partes=[navigator.userAgent||"",navigator.language||"",(screen.width||"")+"x"+(screen.height||""),screen.colorDepth||"",navigator.hardwareConcurrency||"",navigator.platform||"",(Intl.DateTimeFormat().resolvedOptions().timeZone)||""].join("|");'+
    'var hash=0;'+
    'for(var i=0;i<partes.length;i++){ hash=((hash<<5)-hash+partes.charCodeAt(i))|0; }'+
    'return "fp-"+Math.abs(hash).toString(36);'+
    '}catch(e){ return "fp-desconocida"; }'+
    '}'+
    'function claveYaFirmado(){ return "FEM_ASISTENCIA_FIRMADA_"+'+JSON.stringify(String(idForo))+'; }'+
    'function marcarFirmadoLocal(textoFirma){ try{ localStorage.setItem(claveYaFirmado(), textoFirma||"1"); }catch(e){} }'+
    'function mostrarYaFirmado(textoFirma){'+
    'document.getElementById("formulario").style.display="none";'+
    'document.getElementById("estado").textContent="✓ Este dispositivo ya registró su firma de asistencia. ¡Gracias! Ya puede cerrar esta página y continuar en la plenaria.";'+
    'document.getElementById("textoFirma").textContent=textoFirma||"";'+
    '}'+
    'var dispositivoIdAsistencia=calcularHuellaDispositivo();'+
    '(function(){ try{ var previo=localStorage.getItem(claveYaFirmado()); if(previo){ mostrarYaFirmado(previo==="1"?"":previo); } }catch(e){} })();'+
    /*
     * Validación de correo en vivo: mensaje en rojo sobre fondo
     * amarillo (mismo lenguaje visual que los demás errores del
     * formulario principal) con una aclaración pequeña en gris
     * oscuro entre paréntesis. Se exige formato usuario@dominio.tld
     * sin espacios al inicio, al final ni en medio del valor.
     */
    'function correoEsValido(valor){'+
    'var v=String(valor||"");'+
    'if(v!==v.trim())return false;'+
    'if(/\\s/.test(v))return false;'+
    'return /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(v);'+
    '}'+
    'function validarCorreoUI(){'+
    'var campo=document.getElementById("correo");'+
    'var msg=document.getElementById("mensajeErrorCorreo");'+
    'var v=campo.value;'+
    'if(v===""||correoEsValido(v)){campo.classList.remove("correoInvalido");msg.style.display="none";return v!=="";}'+
    'campo.classList.add("correoInvalido");msg.style.display="block";return false;'+
    '}'+
    'document.getElementById("correo").addEventListener("input",validarCorreoUI);'+
    'document.getElementById("correo").addEventListener("blur",validarCorreoUI);'+
    'document.getElementById("btnFirmar").addEventListener("click",function(){'+
    'var btn=this; var estado=document.getElementById("estado");'+
    'var nombre=document.getElementById("nombre").value.trim();'+
    'var tipoAsistencia=document.getElementById("tipoAsistencia").value.trim();'+
    'var cargo=document.getElementById("cargo").value.trim();'+
    'var rolForo=document.getElementById("rolForo").value.trim();'+
    'var documento=document.getElementById("documento").value.trim();'+
    'var correo=document.getElementById("correo").value.trim();'+
    'var telefono=document.getElementById("telefono").value.trim();'+
    'if(!nombre||!tipoAsistencia||!cargo||!rolForo||!documento||!correo){estado.textContent="Complete nombre, tipo de asistencia, cargo, rol en el Foro, número de documento y correo electrónico.";return;}'+
    'if(!correoEsValido(correo)){validarCorreoUI();estado.textContent="Revise el correo electrónico: no es válido.";return;}'+
    'btn.disabled=true; btn.textContent="Firmando…";'+
    'google.script.run.withSuccessHandler(function(res){'+
    'if(res&&res.ok){ marcarFirmadoLocal(res.textoFirma); mostrarYaFirmado(res.textoFirma); }'+
    'else if(res&&res.yaFirmoDispositivo){ marcarFirmadoLocal(res.textoFirma); mostrarYaFirmado(res.textoFirma); }'+
    'else{ btn.disabled=false; btn.textContent="Firmar asistencia"; estado.textContent=(res&&res.mensaje)||"No fue posible registrar la asistencia."; }'+
    '}).withFailureHandler(function(err){ btn.disabled=false; btn.textContent="Firmar asistencia"; estado.textContent="No fue posible registrar la asistencia: "+(err.message||err); })'+
    '.registrarAsistenciaQR('+JSON.stringify(idForo)+',nombre,tipoAsistencia,cargo,rolForo,documento,correo,telefono,dispositivoIdAsistencia);'+
    '});'+
    '</script>'+
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle(tituloPagina+" — FEM 2026")
    .addMetaTag("viewport","width=device-width, initial-scale=1");
}


function subirEvidenciasFEM(idForo,fotoData,fotoName,fotoMime,datos){
  const acceso=obtenerAccesoPorIdForo_(idForo); if(!acceso)throw new Error("ID_FORO no autorizado."); const folder=crearCarpetaIE_(datos.institucion||acceso.ie);
  const decode=(data)=>{const s=String(data||"");const comma=s.indexOf(",");return Utilities.base64Decode(comma>=0?s.substring(comma+1):s);}; const fb=decode(fotoData); if(fb.length>10*1024*1024)throw new Error("La fotografía debe pesar máximo 10 MB."); if(["image/jpeg","image/png"].indexOf(fotoMime)<0)throw new Error("La fotografía debe ser JPG o PNG.");
  const fotoNombre="Foro 2026 ("+(datos.institucion||acceso.ie)+")."+(fotoMime==="image/png"?"png":"jpg");
  const foto=folder.createFile(Utilities.newBlob(fb,fotoMime,fotoNombre));
  foto.setDescription("Participantes del FEM 2026 | I.E. "+(datos.institucion||acceso.ie)+" | Grupo: "+(datos.campos?.grupo?.valor||"")+" | Fecha y hora de submisión: "+new Date()); hacerPublicoSiEsPosible_(foto);
  // La asistencia ya NO produce un archivo aparte: se sigue firmando
  // por QR durante toda la jornada y se incluye al final del informe
  // ejecutivo (ver agregarListadoAsistenciaAlInforme_ en generarInformeFEM).
  return {ok:true,foto:{id:foto.getId(),url:foto.getUrl()},folderId:folder.getId()};
}


function construirParrafoSesion_(titulo,contenido){return titulo+"\n\n"+String(contenido||"");}


function generarInformeFEM(idForo,datosCliente){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const estadoFinal=obtenerEstadoSesiones_(idForo); if(!estadoFinal.s1||!estadoFinal.s2||!estadoFinal.s3)throw new Error("Las tres sesiones deben estar enviadas definitivamente antes de generar el informe.");
    const datos=obtenerDatosGuardadosPorIdForo_(idForo)||datosCliente; if(!datos)throw new Error("No hay datos guardados para generar el informe."); const folder=crearCarpetaIE_(datos.institucion||"Institución Educativa");

    /*
     * El informe se construye desde cero con DocumentApp.create(),
     * en vez de copiar la plantilla externa (TEMPLATE_INFORME_ID).
     * Esa plantilla resultó ser un archivo de Word (.docx) subido a
     * Drive, no un Google Doc nativo — sus copias tampoco lo eran,
     * así que DocumentApp.openById() nunca podía abrirlas. El error
     * "No se puede acceder al documento" no era un problema de
     * tiempos de propagación de Drive: era un formato incompatible,
     * y ningún número de reintentos lo iba a resolver. Crear el
     * documento directamente con DocumentApp siempre produce un
     * Google Doc nativo, sin depender de ningún archivo externo.
     */
    const nombreArchivo="Informe Ejecutivo - "+datos.institucion+" FEM 2026";
    const doc=DocumentApp.create(nombreArchivo);
    const docFile=DriveApp.getFileById(doc.getId());
    folder.addFile(docFile);
    try{ DriveApp.getRootFolder().removeFile(docFile); }catch(errorMover){ Logger.log("No fue posible quitar el informe de la raíz de Drive: "+errorMover.message); }

    /*
     * Paleta institucional del FEM 2026 (la misma del formulario).
     */
    const VERDE=COLOR_VERDE_DOC, GRIS_TEXTO=COLOR_GRIS_TEXTO_DOC, GRIS_FONDO=COLOR_GRIS_FONDO_DOC, GRIS_BORDE=COLOR_GRIS_BORDE_DOC;
    const AZUL_CLARO=COLOR_AZUL_CLARO_DOC, AMARILLO=COLOR_AMARILLO_DOC, NEGRO=COLOR_NEGRO_DOC;

    const body=doc.getBody(); body.clear(); body.setPageWidth(612).setPageHeight(792).setMarginTop(50).setMarginBottom(50).setMarginLeft(48).setMarginRight(48);
    /*
     * Encabezado con el logo de la Alcaldía/SEM a la derecha (se
     * repite en cada página del documento). El logo de la IE ya NO
     * va aquí — a este tamaño de encabezado se veía diminuto — sino
     * grande y centrado en la portada del informe (ver más abajo).
     */
    const h=doc.getHeader()||doc.addHeader(); h.clear();
    const logoIdIE=obtenerLogoIdPorNombreIE_(datos.institucion||"");
    const pLogoAlcaldia=h.appendParagraph(""); pLogoAlcaldia.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    try{ pLogoAlcaldia.appendInlineImage(DriveApp.getFileById(LOGO_ENCABEZADO_ID).getBlob()).setWidth(90).setHeight(50); }catch(e){};
    const footer=doc.getFooter()||doc.addFooter(); footer.clear(); const fp=footer.appendParagraph(""); fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER); try{fp.appendInlineImage(DriveApp.getFileById(LOGO_PIE_ID).getBlob()).setWidth(80).setHeight(40);}catch(e){};
    const fpTexto=footer.appendParagraph("Generado por SEM el "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy 'a las' HH:mm")+". Enviado por "+(datos.campos?.nombre?.valor||"")+" — "+(datos.campos?.correo?.valor||"")+" — "+(datos.campos?.cargo?.valor||"")+" de la "+(datos.institucion||""));
    fpTexto.setAlignment(DocumentApp.HorizontalAlignment.CENTER); fpTexto.editAsText().setForegroundColor(GRIS_TEXTO).setFontSize(9);

    /*
     * Logo de la IE, grande (100x100 — el doble de lo que tenía antes
     * en el encabezado) y centrado, como primer elemento de la
     * portada del informe. Si la IE todavía no tiene logo vinculado,
     * este bloque simplemente no aparece.
     */
    if(logoIdIE){
      try{
        const pLogoIE=body.appendParagraph("");
        pLogoIE.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        pLogoIE.appendInlineImage(DriveApp.getFileById(logoIdIE).getBlob()).setWidth(100).setHeight(100);
      }catch(e){}
    }

    const title=body.appendParagraph("INFORME EJECUTIVO DE "+String(datos.institucion||"").toUpperCase()+" FEM 2026");
    title.setHeading(DocumentApp.ParagraphHeading.TITLE); title.setAlignment(DocumentApp.HorizontalAlignment.CENTER); title.editAsText().setForegroundColor(VERDE);

    const subt=body.appendParagraph("FEM 2026 “Escuela Viva: Voces que construyen territorio”.");
    subt.setHeading(DocumentApp.ParagraphHeading.HEADING2); subt.setAlignment(DocumentApp.HorizontalAlignment.CENTER); subt.editAsText().setForegroundColor(GRIS_TEXTO).setItalic(true);

    const sub2=body.appendParagraph("Foro Educativo Institucional — Neiva 2026");
    sub2.setAlignment(DocumentApp.HorizontalAlignment.CENTER); sub2.editAsText().setForegroundColor(GRIS_TEXTO);

    body.appendHorizontalRule();

    function encabezadoSeccion_(texto){
      const p=body.appendParagraph(texto);
      p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      p.editAsText().setForegroundColor(VERDE).setBold(true);
      return p;
    }

    /*
     * Preguntas de las Sesiones 1, 2 y 3 en el mismo lenguaje visual
     * que .tarjetaPregunta de la plenaria: una sola columna (la
     * pregunta arriba en verde y negrita, la respuesta debajo, no
     * uno al lado del otro), fondo blanco y una línea amarilla a la
     * izquierda a modo de acento — igual técnica de columna angosta
     * que en tablaCaracterizacion_, ya que Documentos no admite un
     * borde de un solo lado.
     */
    function tablaClaveValor_(filas){
      const t=body.appendTable();
      t.setBorderColor("#FFFFFF"); t.setBorderWidth(6);
      filas.forEach(function(x){
        const r=t.appendTableRow();
        const acento=r.appendTableCell("");
        acento.setBackgroundColor(AMARILLO);
        acento.setWidth(6);
        const contenido=r.appendTableCell("");
        contenido.setBackgroundColor("#FFFFFF");
        const pTitulo=contenido.getChild(0).asParagraph();
        pTitulo.setText(String(x[0]||""));
        pTitulo.editAsText().setBold(true).setForegroundColor(VERDE).setFontSize(10);
        const pValor=contenido.appendParagraph(String(x[1]||"—"));
        // Texto de la respuesta: sin negrita — se fija setBold(false)
        // explícitamente porque una celda de tabla nueva en Google
        // Docs puede heredar la negrita del párrafo anterior si no
        // se indica lo contrario.
        pValor.editAsText().setBold(false).setForegroundColor(GRIS_TEXTO).setFontSize(10);
      });
      return t;
    }

    /*
     * Caracterización en una sola columna, una tarjeta apilada por
     * dato — mismo lenguaje visual que .caracterizacionInforme en la
     * plenaria: título en verde y negrita arriba, valor en negro
     * debajo (no uno al lado del otro), fondo azul claro y un acento
     * amarillo a la izquierda. La API de tablas de Google Docs no
     * admite ni esquinas redondeadas ni bordes de un solo lado, así
     * que el acento se aproxima con una primera columna angosta de
     * fondo amarillo, y la separación entre tarjetas con un borde
     * blanco grueso alrededor de cada celda (simula el espacio entre
     * cajas de la cuadrícula original).
     */
    function tablaCaracterizacion_(filas){
      const t=body.appendTable();
      t.setBorderColor("#FFFFFF"); t.setBorderWidth(6);
      filas.forEach(function(x){
        const r=t.appendTableRow();
        const acento=r.appendTableCell("");
        acento.setBackgroundColor(AMARILLO);
        acento.setWidth(6);
        const contenido=r.appendTableCell("");
        contenido.setBackgroundColor(AZUL_CLARO);
        const pTitulo=contenido.getChild(0).asParagraph();
        pTitulo.setText(String(x[0]||""));
        pTitulo.editAsText().setBold(true).setForegroundColor(VERDE).setFontSize(10);
        const pValor=contenido.appendParagraph(String(x[1]||"—"));
        pValor.editAsText().setForegroundColor(NEGRO).setFontSize(10);
      });
      return t;
    }

    /*
     * Participación por estamento con el mismo lenguaje visual que
     * .barraParticipante en la plenaria: etiqueta a la izquierda,
     * una barra verde con el ancho proporcional al porcentaje sobre
     * fondo gris claro, y el valor a la derecha. Se construye con
     * tablas nativas de Documentos (una tabla anidada por fila, de
     * ancho independiente) en vez de una imagen de gráfico — así se
     * ve igual que en la plenaria y no depende de crear/borrar una
     * hoja temporal con un gráfico de Sheets.
     */
    function insertarBarraParticipacion_(celda, pct){
      const ANCHO_BARRA=190;
      const lleno=Math.max(3, Math.round(ANCHO_BARRA*Math.min(1,pct)));
      const vacio=Math.max(1, ANCHO_BARRA-lleno);
      const tablaBarra=celda.appendTable([["",""]]);
      tablaBarra.setBorderWidth(0);
      tablaBarra.setColumnWidth(0,lleno);
      tablaBarra.setColumnWidth(1,vacio);
      tablaBarra.getCell(0,0).setBackgroundColor(VERDE);
      tablaBarra.getCell(0,1).setBackgroundColor(GRIS_FONDO);
    }

    function tablaParticipacionDoc_(datos){
      const defs=["Rector","Coordinador","Docentes","TutorPTA","Orientador","Estudiantes","Padres","Administrativos","Egresados","Sector","Otros"];
      const etiquetas={Rector:"Rector(a)",Coordinador:"Coordinador(a)",Docentes:"Docentes",TutorPTA:"Tutor(a) PTA/PFI 3.0",Orientador:"Orientador(a)",Estudiantes:"Estudiantes",Padres:"Padres/madres/acudientes",Administrativos:"Personal administrativo",Egresados:"Egresados",Sector:"Sector productivo",Otros:"Otros"};
      const c=datos.campos||{};
      const valores=defs.map(function(d){ return Number(c["participantes"+d]?.valor||0); });
      const total=valores.reduce(function(a,b){ return a+b; },0)||1;
      const t=body.appendTable();
      t.setBorderWidth(0);
      defs.forEach(function(d,i){
        const n=valores[i];
        const pct=n/total;
        const r=t.appendTableRow();
        const cEtq=r.appendTableCell(etiquetas[d]);
        cEtq.setWidth(150);
        cEtq.editAsText().setForegroundColor(GRIS_TEXTO).setFontSize(9);
        const cBarra=r.appendTableCell("");
        cBarra.setWidth(200);
        try{ insertarBarraParticipacion_(cBarra, pct); }catch(errorBarra){}
        const cValor=r.appendTableCell(n+" ("+(pct*100).toFixed(1)+"%)");
        cValor.setWidth(70);
        cValor.editAsText().setForegroundColor(GRIS_TEXTO).setFontSize(9);
      });
      return t;
    }

    encabezadoSeccion_("Caracterización");
    const c=datos.campos||{};
    tablaCaracterizacion_([["Institución Educativa",datos.institucion],["DANE",datos.dane],["Rector(a)",c.rector?.valor||""],["Grupo de trabajo",c.grupo?.valor||""],["Responsable",c.nombre?.valor||""],["Cargo",c.cargo?.valor||""],["Correo responsable",c.correo?.valor||""],["Correo institucional",c.correoIE?.valor||""]]);

    /*
     * Logo, título y caracterización quedan en la primera hoja; la
     * participación y el párrafo introductorio empiezan en una
     * segunda hoja aparte.
     */
    body.appendPageBreak();

    encabezadoSeccion_("Participación");
    const totalParticipantesInforme=totalParticipantesServer_(datos);
    const pPart=body.appendParagraph("Participantes: "+totalParticipantesInforme);
    pPart.setHeading(DocumentApp.ParagraphHeading.HEADING2); pPart.editAsText().setForegroundColor(VERDE).setBold(true);
    tablaParticipacionDoc_(datos);

    // Mismo párrafo introductorio que se muestra en la portada de la
    // sesión de plenaria, en la primera página del informe, junto al
    // gráfico de participación.
    const parrafoIntro=body.appendParagraph(
      "La institución educativa "+String(datos.institucion||"")+" construyó colectivamente las conclusiones que se presentan a continuación con la participación de "+totalParticipantesInforme+" integrantes de su comunidad educativa."
    );
    parrafoIntro.editAsText().setForegroundColor(GRIS_TEXTO);

    body.appendPageBreak();

    /*
     * Cada fila de "Acciones" queda dividida en una fila propia por
     * acción (en vez de un solo bloque de texto con todas juntas).
     * Las acciones 1–3 son obligatorias y siempre se muestran; la 4
     * y la 5 son opcionales y solo aparecen si tienen contenido.
     */
    function filasAcciones_(etiqueta, valores){
      return valores
        .map(function(valor, i){ return [etiqueta+" — Acción "+(i+1), valor]; })
        .filter(function(fila, i){ return i<3 || fila[1]; });
    }

    const grupos=[
      {n:"Sesión 1",items:[
        ["Pregunta orientadora (para todos): ¿Cómo hemos avanzado, desde nuestra institución educativa, en el logro de los retos y propósitos planteados en el FEM2025?",c.respuestaSesion1?.valor||""],
        ["Pregunta 2 (para todos): ¿Cómo hemos avanzado, desde nuestra institución educativa, en la implementación de los nuevos grados del nivel de preescolar (jardín, prejardín)?",c.respuestaSesion1Pregunta2?.valor||""]
      ]},
      {n:"Sesión 2",items:[
        ["Pregunta 1 (para todos): ¿Consideran que los currículos actuales que se desarrollan en las instituciones educativas son pertinentes con sus realidades territoriales (sociales, culturales, productivas)? ¿Por qué?",c.respuestaSesion2Pregunta1?.valor||""],
        ...filasAcciones_("Pregunta 2",[1,2,3,4,5].map(i=>c["respuestaSesion2Pregunta2Accion"+i]?.valor||"")),
        ["Pregunta 3 (para todos): ¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar estas acciones?",c.respuestaSesion2Pregunta3?.valor||""],
        ["Pregunta 4 (para todos): ¿Cómo se están articulando estos equipos de trabajo para lograr currículos más pertinentes territorialmente?",c.respuestaSesion2Pregunta4?.valor||""],
        ["Pregunta 5 (para todos): ¿Qué mecanismos de seguimiento se están implementando para que dichas acciones se cumplan?",c.respuestaSesion2Pregunta5?.valor||""]
      ]},
      {n:"Sesión 3",items:[
        ["Pregunta 1 (para todos): ¿Consideran que la toma de decisiones en las instituciones educativas actualmente es participativa y democrática? ¿Por qué?",c.respuestaSesion3Pregunta1?.valor||""],
        ...filasAcciones_("Pregunta 2",[1,2,3,4,5].map(i=>c["respuestaSesion3Pregunta2Accion"+i]?.valor||"")),
        ["Equipos de trabajo (para todos): ¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar las estrategias y mecanismos de participación escolar?",c.respuestaSesion3Pregunta3?.valor||""],
        ["Mecanismos de seguimiento (para todos): ¿Qué mecanismos de seguimiento se están implementando para garantizar las acciones encaminadas a promover gobiernos educativos democráticos?",c.respuestaSesion3Pregunta4?.valor||""]
      ]}
    ];
    grupos.forEach((g,gi)=>{
      encabezadoSeccion_(g.n);
      tablaClaveValor_(g.items);
      if(gi<grupos.length-1)body.appendPageBreak();
    });

    /*
     * Sin salto de página forzado: si queda espacio en la última
     * hoja de las sesiones, la sección de Evidencias continúa ahí
     * mismo en vez de empezar siempre una hoja nueva.
     */
    encabezadoSeccion_("Evidencias de la jornada");
    const pEv=body.appendParagraph("La fotografía original se encuentra almacenada en la carpeta institucional de la IE en Google Drive. La asistencia se firmó por código QR durante la jornada; el listado completo se incluye a continuación.");
    pEv.editAsText().setForegroundColor(GRIS_TEXTO);
    if(c.evidenciaFotoUrl?.valor){ const p2=body.appendParagraph("📷 Ver fotografía de la plenaria"); const t2=p2.editAsText(); t2.setLinkUrl(String(c.evidenciaFotoUrl.valor)); t2.setForegroundColor(VERDE); }

    agregarListadoAsistenciaAlInforme_(body, idForo, datos);

    doc.saveAndClose();
    const pdfBlob=DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName(nombreArchivo+".pdf");
    const pdf=folder.createFile(pdfBlob);
    hacerPublicoSiEsPosible_(docFile); hacerPublicoSiEsPosible_(pdf);
    const ac=obtenerAccesoPorIdForoRaw_(idForo); if(ac){const m=ac.mapa; if(m.ID_INFORME)ac.hoja.getRange(ac.fila,m.ID_INFORME).setValue(doc.getId()); if(m.ID_PDF_INFORME)ac.hoja.getRange(ac.fila,m.ID_PDF_INFORME).setValue(pdf.getId());}
    return {ok:true,docId:doc.getId(),docUrl:doc.getUrl(),pdfId:pdf.getId(),pdfUrl:pdf.getUrl(),folderId:folder.getId()};
  }finally{try{lock.releaseLock();}catch(e){}}
}


function totalParticipantesServer_(datos){const c=datos.campos||{};return ["Rector","Coordinador","Docentes","TutorPTA","Orientador","Estudiantes","Padres","Administrativos","Egresados","Sector","Otros"].reduce((s,k)=>s+Number(c["participantes"+k]?.valor||0),0);}


function enviarInformeFEM(idForo,datos,pdfId){
  const acceso=obtenerAccesoPorIdForoRaw_(idForo); if(!acceso)throw new Error("ID_FORO no autorizado."); const c=datos.campos||{}; const ie=datos.institucion||acceso.ie; const ieSinPrefijo=nombreIESinPrefijoInstitucional_(ie); const logoIEUrlCorreo=urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(ie)); const logoIEHtmlCorreo=logoIEUrlCorreo?("<div style=\"text-align:center;margin:0 0 18px;\"><img src=\""+logoIEUrlCorreo+"\" alt=\"Logo de la institución educativa\" style=\"max-width:56px;max-height:56px;border-radius:8px;\"></div>"):""; const destinatario=String(c.correoIE?.valor||acceso.email||"").trim(); const responsable=String(c.correo?.valor||"").trim(); if(!destinatario)throw new Error("La institución no tiene correo institucional registrado."); const aliases=GmailApp.getAliases().map(x=>x.toLowerCase()); const cuenta=Session.getEffectiveUser().getEmail().toLowerCase(); if(cuenta!==REMITENTE_FEM&&aliases.indexOf(REMITENTE_FEM)===-1)throw new Error("La cuenta de Apps Script no puede enviar como "+REMITENTE_FEM+". Configure esa cuenta o un alias."); const file=DriveApp.getFileById(pdfId); hacerPublicoSiEsPosible_(file); const linkDescarga=file.getUrl(); const subject="Reporte de Informe IE "+ie; const body="Apreciados(as) integrantes de la comunidad educativa de la Institución Educativa "+ieSinPrefijo+":\n\nReciban un cordial saludo de la Secretaría de Educación de Neiva.\n\nAgradecemos a la Institución Educativa por su participación y por el tiempo dedicado al desarrollo del Foro Educativo Institucional – Neiva 2026, así como por los aportes, reflexiones y propuestas construidas colectivamente durante la jornada.\n\nAdjuntamos el Informe Ejecutivo del Foro Educativo Institucional – Neiva 2026, que reúne la caracterización institucional, la participación registrada y las respuestas definitivas construidas durante las tres sesiones de trabajo.\n\nTambién puede descargarlo desde este enlace:\n"+linkDescarga+"\n\nAgradecemos especialmente la disposición de la comunidad educativa para participar en este ejercicio de diálogo, reflexión y construcción colectiva orientado al fortalecimiento de la educación en nuestro municipio.\n\nSecretaría de Educación de Neiva\nForo Educativo Institucional – Neiva 2026\n\“Escuela Viva: Voces que construyen territorio\”"; const to=destinatario; const cc=[responsable].concat(COPIAS_INFORME_FEM).filter(Boolean).join(","); GmailApp.sendEmail(to,subject,body,{htmlBody:logoIEHtmlCorreo+"<p>Apreciados(as) integrantes de la comunidad educativa de la Institución Educativa <strong>"+ieSinPrefijo+"</strong>:</p><p>Reciban un cordial saludo de la Secretaría de Educación de Neiva.</p><p>Agradecemos a la Institución Educativa por su participación y por el tiempo dedicado al desarrollo del <strong>Foro Educativo Institucional – Neiva 2026</strong>, así como por los aportes, reflexiones y propuestas construidas colectivamente durante la jornada.</p><p>Adjuntamos el <strong>Informe Ejecutivo del Foro Educativo Institucional – Neiva 2026</strong>, que reúne la caracterización institucional, la participación registrada y las respuestas definitivas construidas durante las tres sesiones de trabajo.</p><p>📄 <a href=\""+linkDescarga+"\">Descargar el informe aquí</a></p><p>Agradecemos especialmente la disposición de la comunidad educativa para participar en este ejercicio de diálogo, reflexión y construcción colectiva orientado al fortalecimiento de la educación en nuestro municipio.</p><p><strong>Secretaría de Educación de Neiva</strong><br>Foro Educativo Institucional – Neiva 2026<br>“Escuela Viva: Voces que construyen territorio”</p>",cc:cc,from:REMITENTE_FEM,name:"Secretaría de Educación de Neiva",attachments:[file.getBlob()]});

  /*
   * Correo personalizado y directo a quien diligenció el
   * formulario (no solo en copia del correo institucional),
   * agradeciéndole a nombre propio y con el mismo enlace de
   * descarga. Se omite si no dejó correo o si es el mismo
   * correo institucional (para no duplicar el envío).
   */
  if(responsable && responsable.toLowerCase()!==destinatario.toLowerCase()){
    const nombreResponsable=String(c.nombre?.valor||"").trim();
    const saludoResponsable=nombreResponsable?("Estimado(a) "+nombreResponsable+":"):"Estimado(a):";
    const asuntoResponsable="Gracias por diligenciar el Foro Educativo Institucional – "+ie;
    const cuerpoResponsable=saludoResponsable+"\n\nReciba un cordial saludo de la Secretaría de Educación de Neiva.\n\nLe agradecemos personalmente por haber diligenciado el Foro Educativo Institucional – Neiva 2026 en representación de la Institución Educativa "+ieSinPrefijo+".\n\nAdjuntamos el Informe Ejecutivo ya generado. También puede descargarlo desde este enlace:\n"+linkDescarga+"\n\nSecretaría de Educación de Neiva\nForo Educativo Institucional – Neiva 2026\n\“Escuela Viva: Voces que construyen territorio\”";
    try{
      GmailApp.sendEmail(responsable,asuntoResponsable,cuerpoResponsable,{
        htmlBody:logoIEHtmlCorreo+"<p>"+saludoResponsable+"</p><p>Reciba un cordial saludo de la Secretaría de Educación de Neiva.</p><p>Le agradecemos personalmente por haber diligenciado el <strong>Foro Educativo Institucional – Neiva 2026</strong> en representación de la Institución Educativa <strong>"+ieSinPrefijo+"</strong>.</p><p>Adjuntamos el Informe Ejecutivo ya generado. También puede descargarlo desde aquí:</p><p>📄 <a href=\""+linkDescarga+"\">Descargar el informe</a></p><p><strong>Secretaría de Educación de Neiva</strong><br>Foro Educativo Institucional – Neiva 2026<br>“Escuela Viva: Voces que construyen territorio”</p>",
        from:REMITENTE_FEM,
        name:"Secretaría de Educación de Neiva",
        attachments:[file.getBlob()]
      });
    }catch(errorResponsable){
      Logger.log("No fue posible enviar el correo de agradecimiento al responsable: "+errorResponsable.message);
    }
  }

  return {ok:true,linkDescarga:linkDescarga};
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
    const sh=abrirSpreadsheet_().getSheetByName(HOJA_AVANCES);
    if(sh){const mm=obtenerMapaCabeceras_(sh);const row=buscarFilaPorIdForo_(sh,idForo,mm);if(row>0&&mm.ESTADO)sh.getRange(row,mm.ESTADO).setValue("ENVIADO");}
    return {ok:true,pdfId:pdfId};
  }finally{try{lock.releaseLock();}catch(e){}}
}


/*****************************************************
 * VALORACIÓN DE LA ACTIVIDAD (FEMI2026)
 *
 * Subpantalla que responde únicamente quien diligenció el
 * formulario, al finalizar. 4 preguntas de escala 1–5
 * (corazones) + 1 pregunta abierta. Se guarda en una hoja
 * propia, no bloquea ni exige nada más.
 *****************************************************/
const HOJA_VALORACION_FEM = "Valoración FEMI2026";

function asegurarHojaValoracionFEM_(){
  const ss=abrirSpreadsheet_();
  let hoja=ss.getSheetByName(HOJA_VALORACION_FEM);
  if(!hoja) hoja=ss.insertSheet(HOJA_VALORACION_FEM);
  const requeridas=["ID_FORO","IE","FECHA","P1_DIALOGO_REFLEXION","P2_PARTICIPACION","P3_IDEAS_PROPUESTAS","P4_SATISFACCION_INSTRUMENTO","P5_SUGERENCIAS"];
  const last=hoja.getLastColumn();
  if(!last){ hoja.getRange(1,1,1,requeridas.length).setValues([requeridas]); }
  return hoja;
}

function guardarValoracionFEM(idForo, respuestas){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);

    idForo = String(idForo||"").trim();
    if(!idForo) return {ok:false, mensaje:"No fue posible identificar el registro para guardar la valoración."};

    respuestas = respuestas || {};
    const puntajes = ["p1","p2","p3","p4"].map(k=>Number(respuestas[k]));
    for(let i=0;i<puntajes.length;i++){
      if(!(puntajes[i]>=1 && puntajes[i]<=5)){
        return {ok:false, mensaje:"Las 4 primeras preguntas deben calificarse de 1 a 5 corazones."};
      }
    }

    /*
     * La valoración es una encuesta de satisfacción, no un envío que
     * deba bloquearse por una validación estricta de acceso: si por
     * cualquier motivo no se encuentra la fila en AccesosIE (llegó a
     * pasar justo después de "Finalizar", con "El ID_FORO no está
     * autorizado"), se intenta el nombre de la IE desde AvancesForo
     * — donde también queda guardado — en vez de rechazar la
     * valoración completa por no poder mostrar ese dato.
     */
    const acceso = obtenerAccesoPorIdForo_(idForo);
    let nombreIE = acceso ? acceso.ie : "";
    if(!nombreIE){
      const datosGuardados = obtenerDatosGuardadosPorIdForo_(idForo);
      nombreIE = datosGuardados?.institucion || "";
    }

    const hoja=asegurarHojaValoracionFEM_();
    const m=mapaHoja_(hoja);

    // Una sola valoración por IE: si ya se registró una para este
    // ID_FORO, no se crea una fila duplicada (por ejemplo, si se
    // presiona "Enviar valoración" dos veces, o se reingresa después
    // de haberla enviado).
    const ultimaFilaValoracion=hoja.getLastRow();
    if(ultimaFilaValoracion>=2 && m.ID_FORO){
      const idsExistentes=hoja.getRange(2,m.ID_FORO,ultimaFilaValoracion-1,1).getDisplayValues();
      for(let i=0;i<idsExistentes.length;i++){
        if(String(idsExistentes[i][0]||"").trim()===String(idForo)){
          return {ok:true, yaEnviada:true};
        }
      }
    }

    const fila=new Array(hoja.getLastColumn()).fill("");
    const valores={
      ID_FORO:String(idForo),
      IE:nombreIE,
      FECHA:new Date(),
      P1_DIALOGO_REFLEXION:puntajes[0],
      P2_PARTICIPACION:puntajes[1],
      P3_IDEAS_PROPUESTAS:puntajes[2],
      P4_SATISFACCION_INSTRUMENTO:puntajes[3],
      P5_SUGERENCIAS:String(respuestas.p5||"").trim()
    };
    Object.keys(valores).forEach(k=>{ if(m[k]) fila[m[k]-1]=valores[k]; });
    hoja.appendRow(fila);

    return {ok:true};

  }catch(error){
    return {ok:false, mensaje:error.message};
  }finally{
    try{ lock.releaseLock(); }catch(e){}
  }
}

/*
 * Último correo del flujo: se dispara cuando la IE presiona
 * "Cerrar" en la confirmación final, después de enviar la
 * valoración. Confirma la recepción de la valoración y sirve como
 * comprobante de participación en el Foro, indicando el grupo de
 * trabajo de la IE.
 */
function enviarComprobanteParticipacionFEM(idForo, datos){
  const acceso=obtenerAccesoPorIdForoRaw_(idForo);
  if(!acceso) throw new Error("ID_FORO no autorizado.");

  const c=datos?.campos||{};
  const ie=datos?.institucion||acceso.ie;
  const ieSinPrefijo=nombreIESinPrefijoInstitucional_(ie);
  const logoIEUrlCorreo=urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(ie));
  const destinatario=String(c.correoIE?.valor||acceso.email||"").trim();
  const responsable=String(c.correo?.valor||"").trim();
  const nombreResponsable=String(c.nombre?.valor||"").trim();
  const grupo=String(c.grupo?.valor||"").trim() || "sin grupo asignado";

  if(!destinatario) return {ok:false, mensaje:"La institución no tiene correo institucional registrado."};

  const aliases=GmailApp.getAliases().map(x=>x.toLowerCase());
  const cuenta=Session.getEffectiveUser().getEmail().toLowerCase();
  if(cuenta!==REMITENTE_FEM && aliases.indexOf(REMITENTE_FEM)===-1){
    return {ok:false, mensaje:"La cuenta de Apps Script no puede enviar como "+REMITENTE_FEM+". Configure esa cuenta o un alias."};
  }

  const asunto="✅ Comprobante de participación — Foro Educativo Institucional – Neiva 2026";

  const cuerpoTexto=
    "Secretaría de Educación de Neiva\n\n"+
    "Estimada comunidad educativa de la Institución Educativa "+ieSinPrefijo+":\n\n"+
    "Confirmamos la recepción de la valoración de la actividad enviada por "+(nombreResponsable||"su institución")+".\n\n"+
    "Este correo es el comprobante de participación de la Institución Educativa "+ieSinPrefijo+" dentro del grupo "+grupo+" para el Foro Educativo Comunitario, el 24 de septiembre de 2026.\n\n"+
    "Agradecemos nuevamente su participación y los aportes construidos durante la jornada.\n\n"+
    "Secretaría de Educación de Neiva\n"+
    "Foro Educativo Institucional – Neiva 2026\n"+
    "“Escuela Viva: Voces que construyen territorio”";

  const cuerpoHTML=
    "<div style=\"background:#F7F8FA;padding:28px 12px;font-family:Arial,Helvetica,sans-serif;\">"+
    "<div style=\"max-width:520px;margin:0 auto;background:#FFFFFF;border-radius:16px;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,.10);\">"+
    "<div style=\"background:#0B6A44;padding:26px 28px;text-align:center;\">"+
    (logoIEUrlCorreo ? "<img src=\""+logoIEUrlCorreo+"\" alt=\"Logo de la institución educativa\" style=\"display:block;max-width:48px;max-height:48px;margin:0 auto 8px;border-radius:8px;\">" : "")+
    "<div style=\"font-size:40px;line-height:1;margin-bottom:6px;\">✅</div>"+
    "<div style=\"color:#FFFFFF;font-size:18px;font-weight:700;\">Comprobante de participación</div>"+
    "<div style=\"color:#CFE8DC;font-size:13px;margin-top:2px;\">Foro Educativo Institucional — Neiva 2026</div>"+
    "</div>"+
    "<div style=\"padding:28px;\">"+
    "<p style=\"font-size:16px;color:#333333;margin:0 0 14px;\">Estimada comunidad educativa de la Institución Educativa <strong>"+ieSinPrefijo+"</strong>:</p>"+
    "<p style=\"font-size:15px;color:#4A4A4A;line-height:1.6;margin:0 0 20px;\">"+
    "Confirmamos la recepción de la valoración de la actividad enviada por "+(nombreResponsable?"<strong>"+nombreResponsable+"</strong>":"su institución")+"."+
    "</p>"+
    "<div style=\"background:#F7F8FA;border-left:6px solid #F4B400;border-radius:10px;padding:16px 20px;margin:0 0 22px;\">"+
    "<p style=\"font-size:14px;color:#333333;margin:0;\">Este correo es el <strong>comprobante de participación</strong> de la Institución Educativa <strong>"+ieSinPrefijo+"</strong> dentro del <strong>grupo "+grupo+"</strong> para el <strong>Foro Educativo Comunitario</strong>, el <strong>24 de septiembre de 2026</strong>.</p>"+
    "</div>"+
    "<p style=\"font-size:14px;color:#4A4A4A;line-height:1.6;margin:0;\">Agradecemos nuevamente su participación y los aportes construidos colectivamente durante la jornada.</p>"+
    "</div>"+
    "<div style=\"background:#F7F8FA;padding:18px 28px;text-align:center;border-top:1px solid #E5E7EA;\">"+
    "<p style=\"font-size:13px;color:#0B6A44;font-weight:700;margin:0;\">Secretaría de Educación de Neiva</p>"+
    "<p style=\"font-size:12px;color:#888888;margin:4px 0 0;font-style:italic;\">“Escuela Viva: Voces que construyen territorio”</p>"+
    "</div>"+
    "</div>"+
    "</div>";

  const cc=[responsable].concat(COPIAS_INFORME_FEM).filter(Boolean).join(",");
  GmailApp.sendEmail(destinatario, asunto, cuerpoTexto, {
    htmlBody:cuerpoHTML, cc:cc, from:REMITENTE_FEM, name:"Secretaría de Educación de Neiva"
  });

  return {ok:true};
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

    const ss = abrirSpreadsheet_();
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
  const ss=abrirSpreadsheet_(); const sh=ss.getSheetByName(HOJA_ACCESOS); if(!sh||sh.getLastRow()<2)return null;
  const vals=sh.getDataRange().getDisplayValues(); const h=vals[0].map(String); const m={}; h.forEach((x,i)=>m[String(x).trim()]=i);
  for(let i=1;i<vals.length;i++) if(String(vals[i][m.ID_FORO]||"").trim()===id) return {hoja:sh,fila:i+1,mapa:Object.fromEntries(Object.keys(m).map(k=>[k,m[k]+1])),ie:String(vals[i][m.IE]||""),dane:String(vals[i][m.DANE]||""),email:m.EMAIL_IE!==undefined?String(vals[i][m.EMAIL_IE]||""):""};
  return null;
}
function sesionActivaPorIdForo_(idForo,dispositivoId,tokenSesion){
  // Sin temporizador de inactividad: la sesión es válida mientras
  // pertenezca a este mismo dispositivo/token — sin límite de tiempo —
  // hasta que otro dispositivo la tome explícitamente (takeover).
  const props=PropertiesService.getScriptProperties(); const clave=obtenerClaveSesionCodigo_("","",idForo); const raw=props.getProperty(clave); if(!raw)return false; let a; try{a=JSON.parse(raw);}catch(e){return false;} if(a.deviceId!==String(dispositivoId||"")||a.tokenSesion!==String(tokenSesion||""))return false; return true;
}
function validarEnvioFinal_(datos){
  const c=datos?.campos||{}; const v=id=>String(c[id]?.valor||"").trim();
  /*
   * Los mínimos de palabras y de selección de cada sesión ya se
   * validan y se hacen cumplir en su propio momento (Sesión 1, 2 y
   * 3, al enviarlas una por una desde la plenaria/formulario) — no
   * se vuelven a exigir aquí como criterios nuevos. Este envío
   * definitivo solo comprueba que exista una respuesta guardada en
   * cada campo obligatorio; no bloquea el envío por cantidad de
   * palabras ni por ningún otro criterio de longitud o selección.
   */
  const req=["respuestaSesion1","respuestaSesion1Pregunta2","respuestaSesion2Pregunta1","respuestaSesion2Pregunta2Accion1","respuestaSesion2Pregunta2Accion2","respuestaSesion2Pregunta2Accion3","respuestaSesion2Pregunta3","respuestaSesion2Pregunta4","respuestaSesion2Pregunta5","respuestaSesion3Pregunta1","respuestaSesion3Pregunta2Accion1","respuestaSesion3Pregunta2Accion2","respuestaSesion3Pregunta2Accion3","respuestaSesion3Pregunta3","respuestaSesion3Pregunta4"];
  const falt=req.filter(id=>!v(id)); if(falt.length)return {ok:false,mensaje:"Faltan respuestas obligatorias antes de realizar el envío definitivo."};
  return {ok:true};
}
function enviarForoDefinitivo(idForo,tokenSesion,dispositivoId,datos){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const raw=obtenerAccesoPorIdForoRaw_(idForo); if(!raw)return {ok:false,mensaje:"La institución no está autorizada."};
    const estado=String(raw.hoja.getRange(raw.fila,raw.mapa.ESTADO).getValue()||"").toUpperCase();
    if(estado==="ENVIADO")return {ok:true,yaEnviado:true,mensaje:"Las respuestas ya fueron enviadas definitivamente."};
    if(!sesionActivaPorIdForo_(idForo,dispositivoId,tokenSesion))return {ok:false,mensaje:"Otro dispositivo tomó el control de esta sesión. Ingrese nuevamente con uno de los códigos de la institución si desea continuar aquí."};
    const valida=validarEnvioFinal_(datos); if(!valida.ok)return valida;
    datos.idForo=String(idForo); datos.institucion=raw.ie; datos.dane=raw.dane;
    const guardado=guardarAvanceForo(datos); if(!guardado?.ok)throw new Error(guardado?.mensaje||"No fue posible guardar las respuestas.");
    const sh=abrirSpreadsheet_().getSheetByName(HOJA_AVANCES); const m=obtenerMapaCabeceras_(sh); const row=buscarFilaPorIdForo_(sh,idForo,m); const now=new Date();
    ["S1_ENVIADA","S2_ENVIADA","S3_ENVIADA"].forEach(k=>{if(m[k])sh.getRange(row,m[k]).setValue("SI")});
    ["FECHA_ENVIO_S1","FECHA_ENVIO_S2","FECHA_ENVIO_S3","FECHA_ENVIO_DEFINITIVO"].forEach(k=>{if(m[k])sh.getRange(row,m[k]).setValue(now)});
    if(m.ESTADO)sh.getRange(row,m.ESTADO).setValue("ENVIADO");
    guardarEnHojaIE_(datos); actualizarParticipacion_(datos); actualizarGraficoHojaIE_(datos);
    const shIE=abrirSpreadsheet_().getSheetByName(nombreHojaIE_(datos.institucion)); if(shIE&&shIE.getLastRow()>=2){const mi=obtenerMapaCabeceras_(shIE);const rr=buscarFilaPorIdForo_(shIE,idForo,mi);if(rr>0&&mi.ESTADO)shIE.getRange(rr,mi.ESTADO).setValue("ENVIADO");}
    actualizarGraficosParticipacion_();
    const am=raw.mapa; if(am.ESTADO)raw.hoja.getRange(raw.fila,am.ESTADO).setValue("ENVIADO"); if(am.FECHA_ENVIO)raw.hoja.getRange(raw.fila,am.FECHA_ENVIO).setValue(now);
    return {ok:true,fecha:now.toISOString(),idForo:idForo};
  }finally{try{lock.releaseLock();}catch(e){}}
}
function actualizarGraficosParticipacion_(){
  const ss=abrirSpreadsheet_(); const sh=ss.getSheetByName(HOJA_PARTICIPACION); if(!sh||sh.getLastRow()<2)return;
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
    const ss=abrirSpreadsheet_(); const sh=ss.getSheetByName(nombreHojaIE_(datos.institucion)); if(!sh)return;
    sh.getCharts().forEach(c=>sh.removeChart(c));
    const c=datos.campos||{}; const labels=["Rector(a)","Coordinador(a)","Docentes","Tutor PTA PFI/3.0","Orientador(a)","Estudiantes","Padres/madres/acudientes","Personal administrativo","Egresados","Sector productivo","Otros"];
    const ids=["Rector","Coordinador","Docentes","TutorPTA","Orientador","Estudiantes","Padres","Administrativos","Egresados","Sector","Otros"];
    const start=sh.getLastColumn()+2; const vals=ids.map((id,i)=>[labels[i],Number(c["participantes"+id]?.valor||0)]);
    sh.getRange(1,start,vals.length,2).setValues(vals);
    const chart=sh.newChart().setChartType(Charts.ChartType.PIE).addRange(sh.getRange(1,start,vals.length,2)).setOption("title","Participación — "+(datos.institucion||"IE")).setPosition(1,start+3,0,0).build(); sh.insertChart(chart);
  }catch(e){Logger.log("No fue posible crear gráfico IE: "+e.message);}
}

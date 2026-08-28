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
// Foto del diseñador de la aplicación, mostrada en una insignia fija
// abajo a la izquierda de cada pantalla (ver .marcaDisenador en
// CSS.html). Debe estar compartida como "cualquiera con el enlace"
// (ver hacerPublicosLogosGlobales() en Pruebas.js) para poder
// mostrarse como <img> sin sesión de Google.
const DISENADOR_LOGO_ID = "1BXkKDuSH_XhlLbdPtyYlXpJypbFH9f38";
const REMITENTE_FEM = "calidadeducacion@alcaldianeiva.gov.co";
const COPIAS_INFORME_FEM = [
  "adriana.cedeno@alcaldianeiva.gov.co",
  "angelica.rojas@alcaldianeiva.gov.co",
  "ronald.polania@alcaldianeiva.gov.co"
];

/*
 * Corrección de correos institucionales de las 37 IE oficiales, por
 * código DANE (identificador único y exacto — evita cualquier
 * ambigüedad de coincidencia de nombres con prefijos/mayúsculas
 * distintas). Se detectó que la columna "E-MAIL INSTITUCIONAL" de la
 * hoja Oficiales tenía direcciones que no correspondían a estas IE:
 * los correos SÍ se enviaban (Apps Script no valida la existencia del
 * buzón al momento de encolarlos — por eso la cuota de envíos bajaba
 * igual), pero nunca llegaban a destino porque la dirección no era la
 * real de la institución.
 *
 * Esta tabla es la lista verificada y entregada directamente por la
 * Secretaría de Educación de Neiva. enviarAccesosSoloOficialesFEM() y
 * repararEmailIEOficialesFEM() la usan como fuente de verdad —tiene
 * prioridad sobre lo que haya en AccesosIE/Oficiales— y de paso
 * corrigen la columna EMAIL_IE en la hoja para que quede alineada.
 *
 * Si en el futuro cambia algún correo, actualícelo aquí (o, mejor,
 * corrija la hoja Oficiales y quite la entrada correspondiente de
 * este mapa para que vuelva a leerse de ahí).
 */
const CORRECCION_EMAIL_POR_DANE_ = {
  "141001001763": "ieagustincodazzi@alcaldianeiva.gov.co",   // AGUSTIN CODAZZI
  "241001000711": "ieaipecito@alcaldianeiva.gov.co",         // AIPECITO
  "141001002557": "ieangelmaria@alcaldianeiva.gov.co",       // ANGEL MARIA PAREDES
  "141001005866": "ieatanasiog@alcaldianeiva.gov.co",        // ATANASIO GIRARDOT
  "141001004061": "ieceinar@alcaldianeiva.gov.co",           // CEINAR
  "241001001890": "iechapinero@alcaldianeiva.gov.co",        // CHAPINERO
  "141001060441": "ieclaretiano@alcaldianeiva.gov.co",       // I.E. CLARETIANO GUSTAVO TORRES PARRA
  "141001000058": "iedepartamental@alcaldianeiva.gov.co",    // DEPARTAMENTAL TIERRA DE PROMISIÓN
  "141001004720": "ieeduardosantos@alcaldianeiva.gov.co",    // EDUARDO SANTOS
  "441001002747": "iecaguan@alcaldianeiva.gov.co",           // EL CAGUAN
  "141001004452": "ielimonar@alcaldianeiva.gov.co",          // EL LIMONAR
  "141001005301": "ieenriqueolaya@alcaldianeiva.gov.co",     // ENRIQUE OLAYA HERRERA
  "141001002247": "ienormalsuperior@alcaldianeiva.gov.co",   // ESCUELA NORMAL SUPERIOR
  "341001004559": "iegabrielgarcia@alcaldianeiva.gov.co",    // GABRIEL GARCIA MARQUEZ
  "141001004312": "iehumbertotafur@alcaldianeiva.gov.co",    // HUMBERTO TAFUR CHARRY
  "141001003341": "ieinem@alcaldianeiva.gov.co",             // INEM JULIAM MOTTA SALAS
  "141001005181": "iejairomorera@alcaldianeiva.gov.co",      // JAIRO MORERA LIZCANO
  "241001000486": "ieguacirco@alcaldianeiva.gov.co",         // JAIRO MOSQUERA MORENO
  "141001004398": "iejoseeustasio@alcaldianeiva.gov.co",     // JOSE EUSTASIO RIVERA
  "141001001259": "iejuandecabrera@alcaldianeiva.gov.co",    // JUAN DE CABRERA
  "141001000066": "ieliceosantal@alcaldianeiva.gov.co",      // LICEO DE SANTA LIBRADA
  "141001003171": "ieluisignacio@alcaldianeiva.gov.co",      // LUIS IGNACIO ANDRADE
  "441001004839": "iefortalecillas@alcaldianeiva.gov.co",    // MARIA AUXILIADORA FORTALECILLAS
  "141001001038": "iemariacristina@alcaldianeiva.gov.co",    // MARIA CRISTINA ARANGO DE PASTRANA
  "141001003481": "iemisaelpastrana@alcaldianeiva.gov.co",   // MISAEL PASTRANA BORRERO
  "141001000082": "ieoliveriolara@alcaldianeiva.gov.co",     // OLIVERIO LARA BORRERO
  "141001000040": "iepromocion@alcaldianeiva.gov.co",        // PROMOCION SOCIAL
  "141001001321": "iericardoborrero@alcaldianeiva.gov.co",   // RICARDO BORRERO ALVAREZ
  "241001000664": "ierobertoduran@alcaldianeiva.gov.co",     // ROBERTO DURAN ALVIRA
  "141001060336": "ierodrigolara@alcaldianeiva.gov.co",      // RODRIGO LARA BONILLA
  "241001000435": "iesanantonio@alcaldianeiva.gov.co",       // SAN ANTONIO DE ANACONIA
  "441001003433": "iesanluisbeltran@alcaldianeiva.gov.co",   // SAN LUIS BELTRAN
  "141001001593": "menev.cosma-rec@policia.gov.co",          // SAN MIGUEL ARCANGEL
  "141001000023": "ienacionalsl@alcaldianeiva.gov.co",       // SANTA LIBRADA
  "141001000899": "iesantateresa@alcaldianeiva.gov.co",      // SANTA TERESA
  "141001003855": "ieipc@alcaldianeiva.gov.co",              // INSTITUTO TECNICO IPC ANDRES ROSA
  "141001000031": "ietecnicos@alcaldianeiva.gov.co"          // TECNICO SUPERIOR
};

/*
 * Normaliza un código DANE para comparar (solo dígitos, sin espacios
 * ni ceros a la izquierda perdidos por Sheets al mostrarlo como
 * número) antes de buscarlo en CORRECCION_EMAIL_POR_DANE_.
 */
function obtenerCorreoCorregidoPorDane_(dane){
  const clave = String(dane || "").trim();
  return CORRECCION_EMAIL_POR_DANE_[clave] || "";
}

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
  plantilla.disenadorLogoUrl = urlPublicaLogoDrive_(DISENADOR_LOGO_ID);

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
 * SEDES DE UNA IE (para la pregunta de asistencia por QR)
 *
 * En la hoja "Oficiales", cada IE central está en MAYÚSCULAS
 * y sus sedes aparecen en las filas siguientes con escritura
 * normal, hasta la fila de la siguiente IE central. Se usa
 * exactamente el mismo criterio mayúsculas/minúsculas que
 * obtenerInstitucionesJSON() para decidir qué fila es una IE
 * central y cuál es una sede, para no divergir de esa lista.
 *****************************************************/

function obtenerSedesDeIE_(nombreIE){

    try{

        const ss = abrirSpreadsheet_();
        const hoja = ss.getSheetByName("Oficiales");
        if(!hoja) return ["Central/Administrativa"];

        const FILA_ENCABEZADOS = 5;
        const ultimaFila = hoja.getLastRow();
        const ultimaColumna = hoja.getLastColumn();
        if(ultimaFila <= FILA_ENCABEZADOS) return ["Central/Administrativa"];

        const cabeceras = hoja
            .getRange(FILA_ENCABEZADOS, 1, 1, ultimaColumna)
            .getDisplayValues()[0]
            .map(function(v){ return String(v || "").trim().toUpperCase(); });

        const colIE = cabeceras.indexOf("INSTITUCIÓN/SEDE");
        if(colIE === -1) return ["Central/Administrativa"];

        const datos = hoja
            .getRange(FILA_ENCABEZADOS + 1, 1, ultimaFila - FILA_ENCABEZADOS, ultimaColumna)
            .getDisplayValues();

        const objetivo = String(nombreIE || "").trim().toUpperCase();
        const sedes = ["Central/Administrativa"];
        let dentroDeLaIE = false;

        datos.forEach(function(fila){

            const nombre = String(fila[colIE] || "").trim();
            if(!nombre) return;

            const nombreMayusculas = nombre.toUpperCase();
            const nombreMinusculas = nombre.toLowerCase();

            // Misma prueba que obtenerInstitucionesJSON(): si NO es
            // completamente mayúscula, o si es completamente
            // minúscula, es una sede (no una IE central).
            const esSede = (nombre !== nombreMayusculas) || (nombre === nombreMinusculas);

            if(!esSede){
                dentroDeLaIE = (nombreMayusculas === objetivo);
                return;
            }

            if(dentroDeLaIE){
                sedes.push(nombre);
            }

        });

        return sedes;

    }catch(error){

        Logger.log("obtenerSedesDeIE_: " + error.message);
        return ["Central/Administrativa"];

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

    // Sesión 4 / "Sesión Propia": opcional, línea temática adicional
    // propia de la IE. SESION_PROPIA_LINEAS_JSON guarda el arreglo
    // completo de líneas temáticas con sus preguntas y respuestas.
    "SESION_PROPIA_TITULO",
    "SESION_PROPIA_OBJETIVO",
    "SESION_PROPIA_LINEAS_JSON",

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
      obtenerCampoFormulario_(campos, "respuestaSesion3Pregunta4"),

    // Sesión 4 / "Sesión Propia": opcional, línea temática adicional
    // propia de la IE.
    SESION_PROPIA_TITULO:
      obtenerCampoFormulario_(campos, "tituloSesionPropia"),
    SESION_PROPIA_OBJETIVO:
      obtenerCampoFormulario_(campos, "objetivoSesionPropia"),
    SESION_PROPIA_LINEAS_JSON:
      obtenerCampoFormulario_(campos, "sesionPropiaLineasJSON")

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
        "al correo calidadeducacion@alcaldianeiva.gov.co " +
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
     * FUSIÓN DE CAMPOS — hasta 4 dispositivos pueden estar conectados
     * a la vez con el mismo código (ver MAX_SESIONES_SIMULTANEAS_IE).
     * Cada dispositivo guarda TODO su DOM local (construirDatosBorrador
     * en el cliente barre todos los campos con id, no solo los que esa
     * persona escribió) — sin esta fusión, el último dispositivo en
     * guardar sobrescribiría en silencio los campos que otro
     * dispositivo conectado a la vez ya había guardado y que el
     * primero nunca llegó a visitar (quedarían vacíos en su DOM).
     * Por eso el campo entrante solo reemplaza al guardado cuando trae
     * contenido real, o cuando ese campo todavía no existía.
     */
    if (
      filaExistente !== -1 &&
      mapaCabeceras["DATOS"]
    ) {
      try {
        const datosExistentesRaw =
          hoja.getRange(filaExistente, mapaCabeceras["DATOS"]).getValue();
        if (datosExistentesRaw) {
          const datosExistentes = JSON.parse(datosExistentesRaw);
          const camposExistentes = datosExistentes?.campos || {};
          const camposEntrantes = datos.campos || {};
          const camposFusionados = Object.assign({}, camposExistentes);
          Object.keys(camposEntrantes).forEach(function(id) {
            const entrante = camposEntrantes[id];
            const tieneContenido = entrante && (
              entrante.tipo === "checkbox"
                ? true
                : String(entrante.valor ?? "").trim() !== ""
            );
            if (tieneContenido || !(id in camposFusionados)) {
              camposFusionados[id] = entrante;
            }
          });
          datos.campos = camposFusionados;
        }
      } catch (errorFusion) {
        Logger.log("No fue posible fusionar campos con la versión ya guardada (se usa la del cliente tal cual): " + errorFusion.message);
      }
    }


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

/*
 * Hasta 4 dispositivos pueden conectarse SIMULTÁNEAMENTE con el mismo
 * código de acceso de una IE, para que varias personas responsables
 * del diligenciamiento (ver "responsables adicionales" en
 * Caracterización) puedan trabajar a la vez. El envío definitivo del
 * foro sigue siendo uno solo (enviarForoDefinitivo ya controla eso
 * por separado, marcando ESTADO=ENVIADO).
 */
const MAX_SESIONES_SIMULTANEAS_IE = 4;

// Lee el valor guardado como un arreglo de cupos activos. Compatible
// con el formato anterior (un solo objeto, de cuando solo se permitía
// un dispositivo) para no invalidar sesiones ya abiertas al desplegar
// este cambio.
function leerSesionesActivas_(props, clave){
  const guardado = props.getProperty(clave);
  if(!guardado) return [];
  try{
    const parsed = JSON.parse(guardado);
    if(Array.isArray(parsed)) return parsed;
    if(parsed && parsed.deviceId) return [parsed];
    return [];
  }catch(e){ return []; }
}

function reclamarSesionCodigo_(token, codigo, dispositivoId, idForo, forzar) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const props = PropertiesService.getScriptProperties();
    const clave = obtenerClaveSesionCodigo_(token,codigo,idForo);
    const ahora = Date.now();
    let sesiones = leerSesionesActivas_(props, clave);

    /*
     * Responsable principal ("sistematizador/a"): el PRIMER
     * dispositivo en reclamar un cupo para este idForo queda marcado
     * esPrincipal=true de forma permanente (salvo transferencia
     * explícita, ver transferirResponsablePrincipalFEM_). Es quien
     * puede llegar hasta Plenaria y enviar el informe; los demás
     * (hasta 3 colaboradores) llegan como máximo hasta Sesión 4.
     */
    const yaHabiaPrincipal = sesiones.some(s=>s.esPrincipal);

    // Si este dispositivo ya tenía un cupo (recarga de página,
    // reconexión), se reutiliza en vez de contarlo como uno nuevo.
    const existente = sesiones.find(s=>s.deviceId===dispositivoId);
    if(existente){
      existente.ultimaActividad = ahora;
      // Compatibilidad: sesiones creadas antes de que existiera
      // esPrincipal se tratan como principal si nadie más lo es.
      if(existente.esPrincipal===undefined && !yaHabiaPrincipal) existente.esPrincipal = true;
      props.setProperty(clave, JSON.stringify(sesiones));
      return {ok:true, tokenSesion:existente.tokenSesion, esPrincipal:!!existente.esPrincipal};
    }

    if(sesiones.length >= MAX_SESIONES_SIMULTANEAS_IE){
      if(!forzar){
        return {
          ok:false,
          codigo:"SESION_YA_ABIERTA",
          mensaje:"Ya hay "+MAX_SESIONES_SIMULTANEAS_IE+" dispositivos conectados con este código de acceso, el máximo permitido por institución. Si desea continuar en este dispositivo, se cerrará la conexión del dispositivo con menos actividad reciente."
        };
      }
      // No tiene sentido "tomar el lugar de uno en particular" cuando
      // hay hasta 4 cupos: se libera el de menor actividad reciente.
      // Nunca se desaloja al principal por esta vía (si el único
      // candidato a desalojar fuera el principal, se desaloja el
      // siguiente menos activo en su lugar).
      sesiones.sort((a,b)=>(a.ultimaActividad||0)-(b.ultimaActividad||0));
      const indiceExpulsar = sesiones.findIndex(s=>!s.esPrincipal);
      sesiones.splice(indiceExpulsar===-1?0:indiceExpulsar, 1);
    }

    const tokenSesion = Utilities.getUuid();
    const esPrimeraSesion = !sesiones.some(s=>s.esPrincipal);
    sesiones.push({deviceId:dispositivoId, tokenSesion:tokenSesion, ultimaActividad:ahora, esPrincipal:esPrimeraSesion});
    props.setProperty(clave, JSON.stringify(sesiones));
    return {ok:true, tokenSesion:tokenSesion, esPrincipal:esPrimeraSesion};
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
    const sesiones=leerSesionesActivas_(props, clave);
    const mia=sesiones.find(s=>s.deviceId===dispositivoId && s.tokenSesion===tokenSesion);
    // Si este dispositivo ya no tiene cupo (otro lo tomó por
    // inactividad al llenarse los 4), se le informa para que deje de
    // trabajar en silencio creyendo que sigue conectado.
    if(!mia) return {ok:false,codigo:"SESION_NO_AUTORIZADA",mensaje:"Este dispositivo ya no tiene un cupo activo en esta sesión."};
    mia.ultimaActividad=Date.now();
    props.setProperty(clave, JSON.stringify(sesiones));
    return {ok:true, esPrincipal:!!mia.esPrincipal};
  }catch(e){return {ok:false,codigo:"HEARTBEAT_ERROR"};} finally{try{lock.releaseLock();}catch(e){}}
}

/*
 * El responsable principal transfiere su rol a otro colaborador
 * actualmente conectado (el de actividad más reciente) — usado
 * cuando se detecta pérdida de conexión/inactividad prolongada y el
 * propio principal decide ceder el control en vez de continuar.
 */
function transferirResponsablePrincipalFEM(token,codigo,dispositivoId,tokenSesion,idForo){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const props=PropertiesService.getScriptProperties();
    const clave=obtenerClaveSesionCodigo_(token,codigo,idForo);
    const sesiones=leerSesionesActivas_(props, clave);
    const mia=sesiones.find(s=>s.deviceId===dispositivoId && s.tokenSesion===tokenSesion);
    if(!mia) return {ok:false, mensaje:"Este dispositivo ya no tiene un cupo activo en esta sesión."};
    if(!mia.esPrincipal) return {ok:false, mensaje:"Este dispositivo no es el responsable principal del envío."};
    const otras=sesiones.filter(s=>s.deviceId!==dispositivoId);
    if(!otras.length) return {ok:false, mensaje:"No hay otro colaborador conectado en este momento para transferir el control."};
    otras.sort((a,b)=>(b.ultimaActividad||0)-(a.ultimaActividad||0));
    const nuevoPrincipalId=otras[0].deviceId;
    sesiones.forEach(function(s){ s.esPrincipal = (s.deviceId===nuevoPrincipalId); });
    props.setProperty(clave, JSON.stringify(sesiones));
    return {ok:true, nuevoPrincipalDispositivoId:nuevoPrincipalId};
  }catch(e){ return {ok:false, mensaje:"No fue posible transferir el control. Intente nuevamente."}; }
  finally{ try{lock.releaseLock();}catch(e){} }
}

function liberarSesionCodigo(token,codigo,dispositivoId,tokenSesion,idForo){return liberarSesionCodigo_(token,codigo,dispositivoId,tokenSesion,idForo);}
function liberarSesionCodigo_(token,codigo,dispositivoId,tokenSesion,idForo){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const props=PropertiesService.getScriptProperties();
    const clave=obtenerClaveSesionCodigo_(token,codigo,idForo);
    const sesiones=leerSesionesActivas_(props, clave);
    const restantes=sesiones.filter(s=>!(s.deviceId===dispositivoId && s.tokenSesion===tokenSesion));
    if(restantes.length) props.setProperty(clave, JSON.stringify(restantes));
    else props.deleteProperty(clave);
    return {ok:true};
  }
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
                "El código es incorrecto. Verifique el código que se envió a su I.E. Si el error persiste, comuníquese con la Secretaría de Educación de Neiva al correo calidadeducacion@alcaldianeiva.gov.co o al WhatsApp 318 456 1081."
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
     * Bloqueo por horario (HABILITAR_DESDE).
     *
     * Permite dejar el código ya generado y enviado, pero sin abrir
     * el formulario todavía: si la columna HABILITAR_DESDE tiene una
     * fecha/hora futura para esta fila, se rechaza el ingreso con un
     * código especial que el cliente muestra como una página de
     * bloqueo (no un simple mensaje de error), en vez del formulario.
     * Se lee el valor directamente de la celda (no de "filas", que
     * viene en texto) para no depender del formato de fecha local.
     */

    if (mapa["HABILITAR_DESDE"]) {

      const valorHabilitar =
        hoja.getRange(numeroFila, mapa["HABILITAR_DESDE"]).getValue();

      if (
        valorHabilitar instanceof Date &&
        !isNaN(valorHabilitar.getTime()) &&
        new Date() < valorHabilitar
      ) {

        const zonaHabilitar = Session.getScriptTimeZone();
        const horaHabilitacion =
          Utilities.formatDate(valorHabilitar, zonaHabilitar, "h:mm a")
            .replace("AM", "a. m.")
            .replace("PM", "p. m.");

        return {
          ok: false,
          codigo: "BLOQUEADO_POR_HORARIO",
          mensaje:
            "El Foro Educativo Institucional se habilitará a las " +
            horaHabilitacion +
            ". Por favor ingrese nuevamente a partir de esa hora.",
          horaHabilitacion: horaHabilitacion
        };

      }

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

      // true si este dispositivo es el responsable principal
      // ("sistematizador/a"): el único que puede llegar hasta
      // Plenaria y enviar el informe. Los demás (colaboradores) solo
      // llegan hasta Sesión 4 — ver esPrincipal en reclamarSesionCodigo_.
      esPrincipal:
        !!resultadoSesion.esPrincipal,

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


/*
 * Arma el asunto, el texto plano y el HTML del correo de acceso al
 * Foro — el mismo diseño para toda IE que lo reciba (oficial, de
 * prueba, o una simulación): el lenguaje visual del formulario
 * (verde institucional, acento amarillo, tarjeta redondeada) con
 * estilos en línea para que se vea igual en la mayoría de clientes de
 * correo. Extraído de enviarAccesosTodasIE() para reutilizarse
 * también en los envíos restringidos a IE de prueba/oficiales y en la
 * simulación de una IE puntual (ver Pruebas.js).
 */
function construirCorreoAccesoIE_(ie, ieSinPrefijo, codigo, url, logoIEUrlCorreo){
  const asunto = "🎓 Acceso al Foro Educativo Institucional – " + ie;
  const textoEnlace = "Ingreso de la IE " + ieSinPrefijo + " al Foro Educativo Institucional";

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
    /*
     * Un correo no puede ejecutar JavaScript (Gmail/Outlook eliminan
     * cualquier <script> y atributo onclick), así que un botón de
     * "copiar" real no funcionaría dentro del cuerpo del mensaje. En
     * su lugar, el enlace se repite también como texto plano
     * seleccionable — mantener presionado (celular) o triple clic
     * (computador) para copiarlo manualmente.
     */
    "<div style=\"background:#F7F8FA;border:1px dashed #C7CDD1;border-radius:10px;padding:10px 14px;margin:0 0 24px;text-align:center;\">" +
    "<p style=\"font-size:11px;font-weight:700;color:#888888;text-transform:uppercase;letter-spacing:.4px;margin:0 0 4px;\">También puede copiar este enlace</p>" +
    "<p style=\"font-size:12px;color:#0B6A44;word-break:break-all;margin:0;\">" + url + "</p>" +
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

  return { asunto: asunto, textoEnlace: textoEnlace, cuerpoTexto: cuerpoTexto, cuerpoHTML: cuerpoHTML };
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

      const correoArmado = construirCorreoAccesoIE_(ie, ieSinPrefijo, codigo, url, logoIEUrlCorreo);
      const asunto = correoArmado.asunto;
      const cuerpoTexto = correoArmado.cuerpoTexto;
      const cuerpoHTML = correoArmado.cuerpoHTML;

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

                /*
                 * Correo institucional: faltaba esta asignación por
                 * completo, así que cada acceso nuevo quedaba con
                 * EMAIL_IE vacío — el correo de acceso nunca podía
                 * enviarse (enviarAccesosTodasIE/
                 * enviarAccesosSoloOficialesFEM omiten cualquier fila
                 * sin EMAIL_IE). Se prefiere CORRECCION_EMAIL_POR_DANE_
                 * (la lista verificada por la Secretaría) sobre la
                 * columna "E-MAIL INSTITUCIONAL" de Oficiales, que
                 * traía direcciones que no correspondían a la IE real
                 * para varias instituciones.
                 */
                if(columna("EMAIL_IE") >= 0){
                    nuevaFila[
                        columna("EMAIL_IE")
                    ] =
                        obtenerCorreoCorregidoPorDane_(dane) ||
                        String(datosIE.correo || "").trim();
                }


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
  const requeridas=["ID_ACCESO","IE","DANE","CODIGO_ACCESO","TOKEN","URL_ACCESO","ID_FORO","ESTADO","TOKEN_SESION","DISPOSITIVO_ID","FECHA_GENERACION","FECHA_PRIMER_ACCESO","ULTIMA_ACTIVIDAD","FECHA_ENVIO","EMAIL_IE","EMAIL_RESPONSABLE","TIPO","S1_ENVIADA","S2_ENVIADA","S3_ENVIADA","ID_INFORME","ID_PDF_INFORME","LOGO_ID","HABILITAR_DESDE"];
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
  // Comparación normalizada (sin tildes, mayúsculas ni espacios
  // repetidos) en vez de exacta: el nombre que llega aquí puede venir
  // con una capitalización o acentuación ligeramente distinta a como
  // quedó guardado en AccesosIE (por ejemplo, tal como se escribió en
  // Caracterización), y una comparación exacta hacía que el logo de
  // la IE no apareciera al inicio del informe aunque sí existiera.
  const nombreBuscado=normalizarNombreIE_(nombre);
  for(let i=0;i<valores.length;i++){
    if(normalizarNombreIE_(valores[i][m.IE-1]||"")===nombreBuscado){
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
const TIPOS_ASISTENCIA_QR=["Presencial","No asistió: con permiso institucional o incapacidad médica.","No asistió: con permiso de comisión o con acto administrativo."];
// Género femenino fijo (en vez de "(a)"): evita que el "(a)" quede
// literal al insertar el rol dentro de una oración ("En mi papel
// como Relator(a)..."), y no depende de saber el género de quien
// firma.
const ROLES_FORO_QR=["👑 Líder – Rector(a)","🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0","👥 Dinamizador(a) de Mesas de Trabajo","📝 Relator(a)","⏱️ Dinamizador(a) del Tiempo","💻 Dinamizador(a) de la Sistematización","🙋 Participante"];

// Sexo (para el conteo demográfico del informe: niños/niñas,
// adolescentes hombres/mujeres, adultos hombres/mujeres).
const SEXOS_ASISTENCIA_QR=["Masculino","Femenino","Prefiero no decirlo"];

// Edad por RANGOS (ya no un número exacto): alimentan el análisis
// demográfico/de percepción del informe — ver categoriaEdad_() más
// abajo. "no_responde" se cuenta aparte, sin intentar clasificarla.
const RANGOS_EDAD_QR=["0-12","13-18","18-25","25-35","35-45","45-55","55-65","65+","no_responde"];

// Cargos que NO responden la pregunta de condición (jornada/sede):
// según lo pedido, aplica a todos MENOS estos cuatro.
const CARGOS_SIN_CONDICION_QR=["Padre/madre/acudiente","Personal administrativo","Egresado(a)","Sector productivo"];

const JORNADAS_ASISTENCIA_QR=["Mañana","Tarde","Única"];

// Selección múltiple (máx. 3) de fortalezas y dificultades del Foro,
// para el informe institucional. "Otro" despliega un texto libre.
const FORTALEZAS_ASISTENCIA_QR=["Docentes con experiencias exitosas que pueden ser compartidas","Participación activa de los estudiantes","Participación comprometida de las familias","Liderazgo pedagógico de los directivos","Trabajo colaborativo entre docentes","Proyectos institucionales con resultados demostrables","Estrategias que han mejorado los aprendizajes","Experiencias exitosas de inclusión","Buenas prácticas de convivencia","Capacidad institucional para innovar","Uso pertinente de recursos tecnológicos","Aprovechamiento de recursos disponibles en el territorio","Articulación con otras instituciones o entidades","Reconocimiento de las necesidades del contexto","Capacidad de adaptación ante dificultades","Existencia de espacios de participación democrática","Experiencias que pueden ser escaladas o replicadas","Identidad y sentido de pertenencia institucional","Compromiso de la comunidad educativa","Capacidades y talentos de los estudiantes","Redes de apoyo existentes","Experiencias que generan impacto más allá del aula"];

const DIFICULTADES_ASISTENCIA_QR=["Bajo logro de aprendizajes fundamentales.","Brechas de aprendizaje entre estudiantes, grados o sedes.","Estudiantes en riesgo académico sin atención oportuna.","Prácticas pedagógicas con bajo impacto.","Evaluación centrada en calificar y no en mejorar.","Desarticulación entre PEI, currículo y práctica de aula.","Falta de atención a diferentes ritmos y necesidades de aprendizaje.","Exclusión, discriminación o barreras para la inclusión.","Problemas recurrentes de convivencia y acoso escolar.","Baja participación de estudiantes y familias.","Necesidades de formación y acompañamiento docente.","Escaso trabajo colaborativo entre docentes.","Uso insuficiente de recursos y tecnologías.","Infraestructura o dotación que limita el aprendizaje.","Proyectos institucionales sin resultados demostrables.","Buenas experiencias que no se replican institucionalmente.","Decisiones pedagógicas sin suficiente evidencia.","Desconexión entre educación y contexto territorial.","Acciones que deben mantenerse, modificarse o eliminarse.","Prioridades que requieren intervención inmediata.","Responsables y recursos necesarios para actuar.","Compromisos verificables y con resultados esperados."];

// Texto legal completo del consentimiento informado que se muestra
// en un desplegable (con icono de ojo) antes de firmar asistencia.
// "{{IE}}" se reemplaza por el nombre de la institución al generar
// la página (ver paginaAsistenciaQR_).
const TEXTO_CONSENTIMIENTO_ASISTENCIA_QR=
"CONSENTIMIENTO INFORMADO PARA EL TRATAMIENTO DE DATOS PERSONALES\n\n"+
"AVISO INSTITUCIONAL\n\n"+
"Este aplicativo tiene carácter institucional y oficial: es una iniciativa del equipo de Calidad Educativa de la Secretaría de Educación de Neiva (SEM Neiva), de diseño propio y exclusivo para la SEM Neiva, desarrollado por el equipo de Calidad Educativa de la SEM Neiva sin ningún costo para la entidad — no hace parte de ningún contrato con prestadores de servicio externos. La Secretaría de Educación de Neiva se reserva el derecho de reclamación de los derechos de autor de este aplicativo web. Toda la lógica del código base no podrá ser replicada ni modificada bajo ninguna circunstancia.\n\n"+
"En el marco del Foro Educativo Institucional de Neiva 2026 — FEM 2026, la Secretaría de Educación de Neiva (SEM Neiva) y la Institución Educativa {{IE}}, en el ámbito de sus respectivas competencias y de conformidad con la normativa colombiana aplicable sobre protección de datos personales, podrán realizar el tratamiento de los datos personales suministrados mediante este formulario de asistencia.\n\n"+
"Los datos serán tratados exclusivamente para las finalidades informadas en este formulario y de acuerdo con los principios, derechos y condiciones establecidos en la normativa vigente sobre protección de datos personales.\n\n"+
"¿PARA QUÉ SE UTILIZARÁN MIS DATOS?\n\n"+
"Los datos personales registrados mediante este formulario podrán ser utilizados para:\n\n"+
"• Registrar y verificar la asistencia al Foro Educativo Institucional de Neiva 2026.\n"+
"• Elaborar estadísticas generales relacionadas con la participación en el Foro Educativo Municipal — FEM 2026.\n"+
"• Incorporar la información correspondiente en el listado de asistencia que será entregado a la Secretaría de Educación de Neiva (SEM Neiva), para las finalidades institucionales relacionadas con el evento.\n"+
"• Apoyar la elaboración del informe ejecutivo del Foro Educativo Institucional, que será remitido a la Institución Educativa {{IE}}.\n"+
"• Generar información institucional, estadística y de seguimiento relacionada con el desarrollo y participación en el Foro, dentro de las finalidades propias del evento y de acuerdo con la normativa aplicable.\n\n"+
"Los datos no serán utilizados para fines comerciales, publicitarios o ajenos a las finalidades aquí informadas.\n\n"+
"¿QUIÉNES PODRÁN TRATAR LA INFORMACIÓN?\n\n"+
"La información podrá ser tratada por la Secretaría de Educación de Neiva (SEM Neiva) y por la Institución Educativa {{IE}}, de acuerdo con sus respectivas responsabilidades, competencias y finalidades institucionales.\n\n"+
"El tratamiento y circulación de los datos estarán sujetos a las condiciones y restricciones establecidas por la normativa colombiana de protección de datos personales.\n\n"+
"PARTICIPANTES MENORES DE EDAD\n\n"+
"En el caso de niños, niñas y adolescentes, sus datos personales gozan de especial protección de acuerdo con la legislación colombiana.\n\n"+
"El tratamiento de sus datos deberá realizarse respetando sus derechos fundamentales, su interés superior y las condiciones establecidas por la normativa vigente.\n\n"+
"Cuando la autorización para el tratamiento sea requerida, esta deberá ser otorgada por quien tenga la representación legal o facultad para autorizarla, de conformidad con la normativa aplicable.\n\n"+
"La información de niños, niñas y adolescentes no será utilizada para finalidades distintas de las informadas en este consentimiento.\n\n"+
"PARTICIPANTES EN GENERAL\n\n"+
"Los participantes mayores de edad podrán autorizar, cuando corresponda, el tratamiento de sus datos personales para las finalidades expresamente informadas en este formulario.\n\n"+
"La autorización se limita a las finalidades descritas anteriormente y no implica autorización para usos diferentes o incompatibles con dichas finalidades.\n\n"+
"SERVIDORES Y FUNCIONARIOS PÚBLICOS\n\n"+
"Para los servidores y funcionarios públicos que participen en el Foro, la información de asistencia podrá ser tratada en el marco de las actividades institucionales, administrativas, estadísticas y de seguimiento relacionadas con el evento.\n\n"+
"Cuando el tratamiento de datos requiera autorización del titular, esta se solicitará mediante la presente manifestación.\n\n"+
"El tratamiento que se realice en cumplimiento de funciones legales o competencias de las entidades públicas estará sujeto a la base jurídica correspondiente y a las disposiciones de protección de datos personales aplicables.\n\n"+
"DERECHOS DEL TITULAR DE LOS DATOS\n\n"+
"De acuerdo con la normativa colombiana aplicable, el titular de los datos personales podrá, cuando sea legalmente procedente:\n\n"+
"• Conocer los datos personales que sean objeto de tratamiento.\n"+
"• Solicitar la actualización o rectificación de información cuando sea inexacta, incompleta o desactualizada.\n"+
"• Solicitar información sobre el uso y tratamiento de sus datos personales.\n"+
"• Solicitar prueba de la autorización otorgada, cuando esta sea requerida.\n"+
"• Revocar la autorización otorgada, cuando legalmente sea procedente.\n"+
"• Solicitar la supresión de sus datos personales cuando sea legalmente procedente.\n"+
"• Presentar consultas, solicitudes o reclamos relacionados con el tratamiento de sus datos personales a través de los canales institucionales correspondientes.\n\n"+
"El ejercicio de estos derechos estará sujeto a las condiciones, excepciones y obligaciones de conservación previstas en la normativa vigente.\n\n"+
"SEGURIDAD Y CONFIDENCIALIDAD\n\n"+
"La información será tratada aplicando las medidas de seguridad y confidencialidad correspondientes, de acuerdo con las obligaciones establecidas en la normativa colombiana de protección de datos personales.\n\n"+
"La información no será comercializada ni utilizada para fines diferentes de aquellos informados en este consentimiento, salvo que exista una obligación legal, una competencia institucional o una base jurídica que permita dicho tratamiento.\n\n"+
"AUTORIZACIÓN\n\n"+
"Al seleccionar la opción \"He leído y acepto el tratamiento de mis datos personales\" manifiesto que he sido informado(a) sobre las finalidades para las cuales podrán ser tratados los datos personales suministrados mediante este formulario y, cuando la autorización sea legalmente requerida, autorizo su tratamiento para las finalidades descritas anteriormente, de conformidad con la normativa colombiana vigente.\n\n"+
"En el caso de niños, niñas y adolescentes, esta autorización deberá entenderse de acuerdo con las reglas especiales aplicables a su tratamiento de datos personales y, cuando corresponda, deberá ser otorgada por su padre, madre, representante legal o quien se encuentre facultado para ello.\n\n"+
"FINALIDAD ESPECÍFICA DE LOS INFORMES\n\n"+
"La información de asistencia podrá contribuir a las estadísticas del FEM 2026, formar parte del listado de asistencia que será entregado a la Secretaría de Educación de Neiva (SEM Neiva) y ser utilizada, dentro de las finalidades informadas, para la elaboración del informe ejecutivo que será enviado a la Institución Educativa {{IE}}.\n\n"+
"La inclusión de información en estos instrumentos se realizará de acuerdo con las finalidades del evento y las disposiciones aplicables sobre protección de datos personales.";

function asegurarHojaAsistenciaQR_(){
  const ss=abrirSpreadsheet_();
  let hoja=ss.getSheetByName(HOJA_ASISTENCIA_QR);
  if(!hoja) hoja=ss.insertSheet(HOJA_ASISTENCIA_QR);
  const requeridas=["ID_FORO","IE","NOMBRE_COMPLETO","SEXO","EDAD","TIPO_ASISTENCIA","CARGO","ROL_FORO","JORNADA","SEDE","FORTALEZAS","FORTALEZA_OTRO","DIFICULTADES","DIFICULTAD_OTRO","NUMERO_DOCUMENTO","CORREO","TELEFONO","CONSENTIMIENTO","FECHA","HORA","DISPOSITIVO_ID"];
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

/*
 * "26-Agosto-2026, a las 14:32" — usado en el pie de foto de
 * evidencias (pantalla de Evidencias e informe ejecutivo).
 */
function formatearFechaFotoEvidencia_(fecha){
  const zona=Session.getScriptTimeZone();
  const meses=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  const dia=Utilities.formatDate(fecha,zona,"dd");
  const mesIndex=Number(Utilities.formatDate(fecha,zona,"M"))-1;
  const mes=(meses[mesIndex]||"");
  const mesCapitalizado=mes.charAt(0).toUpperCase()+mes.slice(1);
  const anio=Utilities.formatDate(fecha,zona,"yyyy");
  const hora=Utilities.formatDate(fecha,zona,"HH:mm");
  return dia+"-"+mesCapitalizado+"-"+anio+", a las "+hora;
}

/*
 * "27 de agosto de 2026, a las 12:25" — formato largo usado en el
 * nuevo pie de foto (distinto del corto "27-Agosto-2026, a las 14:32"
 * de formatearFechaFotoEvidencia_, que se conserva por si algo más lo
 * sigue usando).
 */
function formatearFechaSubidaFoto_(fecha){
  const zona=Session.getScriptTimeZone();
  const meses=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  const dia=Utilities.formatDate(fecha,zona,"d");
  const mesIndex=Number(Utilities.formatDate(fecha,zona,"M"))-1;
  const mesCapitalizado=(meses[mesIndex]||"").charAt(0).toUpperCase()+(meses[mesIndex]||"").slice(1);
  const anio=Utilities.formatDate(fecha,zona,"yyyy");
  const hora=Utilities.formatDate(fecha,zona,"HH:mm");
  return dia+" de "+mesCapitalizado+" de "+anio+", a las "+hora;
}

function construirPieFotoEvidencia_(ieSinPrefijo,fecha,cantidadParticipantes){
  const cantidad=Number(cantidadParticipantes||0);
  return cantidad+" de participantes de la IE "+ieSinPrefijo+" en actividad del Foro Educativo Institucional — Neiva 2026. Subida el "+formatearFechaSubidaFoto_(fecha)+".";
}

function registrarAsistenciaQR(idForo, nombre, sexo, edad, tipoAsistencia, cargo, rolForo, jornada, sede, fortalezas, fortalezaOtro, dificultades, dificultadOtro, documento, correo, telefono, consentimiento, dispositivoId){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);

    idForo=String(idForo||"").trim();
    nombre=String(nombre||"").trim();
    sexo=String(sexo||"").trim();
    edad=String(edad||"").trim();
    tipoAsistencia=String(tipoAsistencia||"").trim();
    cargo=String(cargo||"").trim();
    rolForo=String(rolForo||"").trim();
    jornada=String(jornada||"").trim();
    sede=String(sede||"").trim();
    fortalezas=Array.isArray(fortalezas)?fortalezas.filter(Boolean).map(String):[];
    fortalezaOtro=String(fortalezaOtro||"").trim();
    dificultades=Array.isArray(dificultades)?dificultades.filter(Boolean).map(String):[];
    dificultadOtro=String(dificultadOtro||"").trim();
    documento=String(documento||"").trim();
    correo=String(correo||"").trim();
    telefono=String(telefono||"").trim();
    dispositivoId=String(dispositivoId||"").trim();

    if(!idForo) return {ok:false, mensaje:"Enlace de asistencia inválido."};
    if(!nombre || !sexo || !edad || !tipoAsistencia || !cargo || !rolForo || !documento || !correo) return {ok:false, mensaje:"Complete nombre, sexo, edad, tipo de asistencia, cargo, rol en el Foro, número de documento y correo electrónico."};
    if(!consentimiento) return {ok:false, mensaje:"Debe aceptar el tratamiento de sus datos personales para continuar."};
    if(SEXOS_ASISTENCIA_QR.indexOf(sexo)===-1) return {ok:false, mensaje:"Seleccione una opción de sexo válida."};
    if(RANGOS_EDAD_QR.indexOf(edad)===-1) return {ok:false, mensaje:"Seleccione una edad válida."};
    if(TIPOS_ASISTENCIA_QR.indexOf(tipoAsistencia)===-1) return {ok:false, mensaje:"Seleccione un tipo de asistencia válido."};
    if(ROLES_FORO_QR.indexOf(rolForo)===-1) return {ok:false, mensaje:"Seleccione un rol válido en el Foro Educativo Institucional."};

    // La pregunta de condición (jornada/sede) solo aplica a quienes
    // NO estén en CARGOS_SIN_CONDICION_QR.
    const requiereCondicion=CARGOS_SIN_CONDICION_QR.indexOf(cargo)===-1;
    if(requiereCondicion){
      if(!jornada || JORNADAS_ASISTENCIA_QR.indexOf(jornada)===-1) return {ok:false, mensaje:"Seleccione la jornada a la que pertenece."};
      if(!sede) return {ok:false, mensaje:"Seleccione la sede a la que pertenece."};
    }

    if(fortalezas.length>3) return {ok:false, mensaje:"Seleccione máximo 3 fortalezas."};
    if(dificultades.length>3) return {ok:false, mensaje:"Seleccione máximo 3 dificultades u oportunidades de mejora."};
    if(fortalezas.indexOf("Otro")!==-1 && !fortalezaOtro) return {ok:false, mensaje:"Especifique la fortaleza que seleccionó como \"Otro\"."};
    if(dificultades.indexOf("Otro")!==-1 && !dificultadOtro) return {ok:false, mensaje:"Especifique la dificultad que seleccionó como \"Otro\"."};

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

    // Evitar duplicados: la MISMA PERSONA (mismo documento) firmando
    // dos veces para este mismo foro no crea una segunda fila — se le
    // devuelve su firma ya registrada. Ya NO se limita a una sola
    // firma por dispositivo: es habitual que varias personas firmen
    // desde el mismo equipo/celular compartido (un tablet institucional,
    // un solo teléfono que se pasa de mano en mano), así que un mismo
    // dispositivo puede registrar tantas firmas de personas distintas
    // como haga falta. dispositivoId se sigue guardando en la fila
    // solo con fines de auditoría/analítica, sin bloquear nada.
    const ultimaFila=hoja.getLastRow();
    if(ultimaFila>=2){
      const filas=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getValues();
      for(let i=0;i<filas.length;i++){
        const mismoForo=String(filas[i][m.ID_FORO-1]||"").trim()===idForo;
        if(!mismoForo) continue;
        if(String(filas[i][m.NUMERO_DOCUMENTO-1]||"").trim()===documento){
          return {ok:true, yaRegistrado:true, textoFirma:textoFirmaDesdeFila_(filas[i])};
        }
      }
    }

    const ahora=new Date();
    const zona=Session.getScriptTimeZone();
    const fila=new Array(hoja.getLastColumn()).fill("");
    const valores={ID_FORO:idForo, IE:acceso.ie, NOMBRE_COMPLETO:nombre, SEXO:sexo, EDAD:edad, TIPO_ASISTENCIA:tipoAsistencia, CARGO:cargo, ROL_FORO:rolForo,
      JORNADA:requiereCondicion?jornada:"", SEDE:requiereCondicion?sede:"",
      FORTALEZAS:fortalezas.join(" | "), FORTALEZA_OTRO:fortalezaOtro,
      DIFICULTADES:dificultades.join(" | "), DIFICULTAD_OTRO:dificultadOtro,
      NUMERO_DOCUMENTO:documento, CORREO:correo, TELEFONO:telefono, CONSENTIMIENTO:"Sí", DISPOSITIVO_ID:dispositivoId,
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
      sexo:String(m.SEXO?fila[m.SEXO-1]:""||""),
      edad:String(m.EDAD?fila[m.EDAD-1]:""||""),
      tipoAsistencia:String(fila[m.TIPO_ASISTENCIA-1]||""),
      cargo:String(fila[m.CARGO-1]||""),
      rolForo:String(fila[m.ROL_FORO-1]||""),
      jornada:String(m.JORNADA?fila[m.JORNADA-1]:""||""),
      sede:String(m.SEDE?fila[m.SEDE-1]:""||""),
      fortalezas:String(m.FORTALEZAS?fila[m.FORTALEZAS-1]:""||"").split(" | ").filter(Boolean),
      fortalezaOtro:String(m.FORTALEZA_OTRO?fila[m.FORTALEZA_OTRO-1]:""||""),
      dificultades:String(m.DIFICULTADES?fila[m.DIFICULTADES-1]:""||"").split(" | ").filter(Boolean),
      dificultadOtro:String(m.DIFICULTAD_OTRO?fila[m.DIFICULTAD_OTRO-1]:""||""),
      documento:String(fila[m.NUMERO_DOCUMENTO-1]||""),
      correo:String(fila[m.CORREO-1]||""),
      telefono:String(fila[m.TELEFONO-1]||""),
      fecha:String(fila[m.FECHA-1]||""),
      hora:String(fila[m.HORA-1]||"")
    }));
}

/*
 * Se llama al cambiar el método de asistencia de QR a PDF: al dejar
 * de usarse el QR, se eliminan todas las firmas ya registradas para
 * este foro. No tendría sentido conservarlas — el análisis demográfico
 * y de fortalezas/dificultades del informe solo se genera a partir de
 * firmas por QR, así que dejarlas huérfanas podría confundirse con
 * datos vigentes del método ya no elegido.
 */
function eliminarAsistenciaQRPorCambioMetodo(idForo){
  try{
    idForo=String(idForo||"").trim();
    if(!idForo) return {ok:false, mensaje:"ID_FORO inválido."};
    const hoja=asegurarHojaAsistenciaQR_();
    const m=mapaHoja_(hoja);
    const ultimaFila=hoja.getLastRow();
    if(ultimaFila<2) return {ok:true, eliminadas:0};
    const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getValues();
    let eliminadas=0;
    for(let i=valores.length-1;i>=0;i--){
      if(String(valores[i][m.ID_FORO-1]||"").trim()===idForo){
        hoja.deleteRow(i+2);
        eliminadas++;
      }
    }
    return {ok:true, eliminadas:eliminadas};
  }catch(error){
    return {ok:false, mensaje:error.message};
  }
}

/*
 * Se llama al cambiar el método de asistencia de PDF a QR: se envía a
 * la papelera el archivo PDF ya subido a la carpeta de Drive de la
 * IE (si lo había), para no dejar un documento huérfano que ya no
 * aparecerá referenciado en el informe.
 */
function eliminarAsistenciaPDFPorCambioMetodo(pdfFileId){
  try{
    pdfFileId=String(pdfFileId||"").trim();
    if(!pdfFileId) return {ok:true, eliminado:false};
    DriveApp.getFileById(pdfFileId).setTrashed(true);
    return {ok:true, eliminado:true};
  }catch(error){
    return {ok:false, mensaje:error.message};
  }
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
 * Lista en vivo de quiénes ya firmaron asistencia por QR — usada por
 * la pantalla "Firmas de asistencia en vivo" (Evidencias y la
 * pantalla de generar informe). Devuelve solo lo necesario para
 * mostrar en pantalla, no el registro completo (documento, correo,
 * teléfono quedan fuera de esta vista pública en tiempo real).
 */
function obtenerListaAsistentesQR(idForo){
  try{
    const asistentes=obtenerAsistentesQR_(idForo).map(function(a){
      return { nombre:a.nombre, cargo:a.cargo, rolForo:a.rolForo, hora:a.hora };
    });
    return { ok:true, asistentes:asistentes };
  }catch(error){
    return { ok:false, mensaje:error.message };
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

/*****************************************************
 * PERFIL DE LOS PARTICIPANTES Y PERCEPCIÓN DEL FORO
 * (a partir de las firmas de asistencia por QR: sexo, edad,
 * jornada, sede, fortalezas y dificultades/oportunidades)
 *
 * Se construye con párrafos ESTÁNDAR (conectores fijos) y las
 * respuestas literales de las personas — nunca parafraseadas: Apps
 * Script no tiene ninguna función de parafraseo, y hacerlo con una
 * IA externa implicaría costo, latencia y riesgo de inventar texto
 * que la persona no escribió. Ver la respuesta dada al usuario.
 *****************************************************/

/*
 * Edad por RANGOS (ver RANGOS_EDAD_QR): 0-12 = niños/niñas, 13-18 =
 * adolescentes, las 6 franjas de 18 en adelante = adultos. "no_responde"
 * se devuelve tal cual, sin clasificar — el informe solo cuenta
 * cuántas personas eligieron esa opción (ver
 * agregarPerfilYPercepcionAlInforme_). Se conserva compatibilidad con
 * firmas antiguas guardadas como número exacto (0-99), de antes de
 * este cambio.
 */
function categoriaEdad_(edad){
  const e=String(edad||"").trim();
  if(e==="no_responde") return "no_responde";
  if(e==="0-12") return "nino";
  if(e==="13-18") return "adolescente";
  if(["18-25","25-35","35-45","45-55","55-65","65+"].indexOf(e)!==-1) return "adulto";
  const n=Number(e);
  if(isNaN(n)) return "";
  if(n<=12) return "nino";
  if(n<=18) return "adolescente";
  return "adulto";
}

function calcularDemografiaAsistentes_(asistentes){
  const t={ninos:0,ninas:0,adolescentesHombres:0,adolescentesMujeres:0,hombresAdultos:0,mujeresAdultas:0,noResponde:0,otro:0};
  asistentes.forEach(function(p){
    const cat=categoriaEdad_(p.edad);
    const sexo=String(p.sexo||"");
    const esHombre=sexo==="Masculino", esMujer=sexo==="Femenino";
    if(cat==="no_responde"){ t.noResponde++; return; }
    if(cat==="nino"){ if(esHombre)t.ninos++; else if(esMujer)t.ninas++; else t.otro++; }
    else if(cat==="adolescente"){ if(esHombre)t.adolescentesHombres++; else if(esMujer)t.adolescentesMujeres++; else t.otro++; }
    else if(cat==="adulto"){ if(esHombre)t.hombresAdultos++; else if(esMujer)t.mujeresAdultas++; else t.otro++; }
  });
  return t;
}

// Cuenta votos de una pregunta de selección múltiple (fortalezas o
// dificultades), sin contar "Otro" como opción tallada — esas se
// listan aparte, literales.
function tallyOpciones_(personas,campo){
  const conteo={};
  personas.forEach(function(p){
    (p[campo]||[]).forEach(function(o){ if(o&&o!=="Otro") conteo[o]=(conteo[o]||0)+1; });
  });
  return Object.keys(conteo).map(function(k){return {opcion:k,votos:conteo[k]};}).sort(function(a,b){return b.votos-a.votos;});
}

function construirGraficoColumnas_(titulo,etiquetas,valores){
  const dt=Charts.newDataTable().addColumn(Charts.ColumnType.STRING,"Categoría").addColumn(Charts.ColumnType.NUMBER,"Cantidad");
  etiquetas.forEach(function(e,i){ dt.addRow([e,valores[i]]); });
  return Charts.newColumnChart().setDataTable(dt.build()).setTitle(titulo).setDimensions(480,300).setColors(["#0B6A44"]).build().getAs("image/png");
}

function construirGraficoBarrasHorizontal_(titulo,etiquetas,valores){
  const dt=Charts.newDataTable().addColumn(Charts.ColumnType.STRING,"Opción").addColumn(Charts.ColumnType.NUMBER,"Votos");
  etiquetas.forEach(function(e,i){ dt.addRow([e,valores[i]]); });
  return Charts.newBarChart().setDataTable(dt.build()).setTitle(titulo).setDimensions(500,320).setColors(["#0B6A44"]).build().getAs("image/png");
}

function agregarGraficoConOtro_(body,estilos,titulo,tally,otrosTextos){
  const top5=tally.slice(0,5);
  if(top5.length){
    try{
      const blob=construirGraficoBarrasHorizontal_(titulo,top5.map(function(x){return x.opcion;}),top5.map(function(x){return x.votos;}));
      body.appendImage(blob).setWidth(430);
    }catch(errorGrafico){ Logger.log("Gráfico \""+titulo+"\": "+errorGrafico.message); }
  }
  if(otrosTextos&&otrosTextos.length){
    const pOtro=body.appendParagraph("Otro:");
    pOtro.editAsText().setBold(true).setForegroundColor(estilos.VERDE);
    otrosTextos.forEach(function(texto){
      const p=body.appendParagraph("• "+texto);
      p.editAsText().setForegroundColor(estilos.GRIS_TEXTO).setFontSize(10);
    });
  }
}

/*
 * Párrafos estándar de percepción por categoría de edad: mismo
 * conector fijo para todas las instituciones (no se elige al azar
 * entre las variantes propuestas, para que el informe sea
 * consistente), con las 3 fortalezas y 3 dificultades más votadas
 * por esa categoría. top3_ recorta el punto final de cada opción
 * (las de dificultades ya lo traen) para que la lista con "; " se
 * lea bien dentro de la frase.
 */
function top3Texto_(tally){
  const texto=tally.slice(0,3).map(function(x){ return x.opcion.replace(/\.$/,""); }).join("; ");
  return texto||"sin datos suficientes registrados en esta jornada";
}

/*
 * Sugerencias/comentarios (P5, pregunta abierta) de la valoración de
 * la actividad, citados de forma literal. NOTA: en el flujo actual,
 * la valoración se responde DESPUÉS de generar el informe ejecutivo
 * (ver enviarAccesoIndividualIEPrueba_/flujo de cierre en App.html),
 * así que esta función normalmente no encontrará nada la primera
 * vez que se genera el informe — solo si este se vuelve a generar
 * más adelante para la misma IE.
 */
function obtenerSugerenciasValoracion_(idForo){
  try{
    const hoja=abrirSpreadsheet_().getSheetByName(HOJA_VALORACION_FEM);
    if(!hoja) return "";
    const m=mapaHoja_(hoja);
    if(!m.ID_FORO||!m.P5_SUGERENCIAS) return "";
    const ultimaFila=hoja.getLastRow();
    if(ultimaFila<2) return "";
    const filas=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();
    // Si se envió más de una vez, se usa la más reciente (última fila).
    for(let i=filas.length-1;i>=0;i--){
      if(String(filas[i][m.ID_FORO-1]||"").trim()===String(idForo||"").trim()){
        return String(filas[i][m.P5_SUGERENCIAS-1]||"").trim();
      }
    }
    return "";
  }catch(error){
    Logger.log("obtenerSugerenciasValoracion_: "+error.message);
    return "";
  }
}

function agregarPerfilYPercepcionAlInforme_(body, idForo, datos, estilos){
  const asistentes=obtenerAsistentesQR_(idForo);
  // Sin prefijo ("IE"/"Institución Educativa"): en esta función
  // ieTitulo siempre se usa como "...de la IE {ieTitulo}" / "...en la
  // IE {ieTitulo}", así que dejarlo con el prefijo original duplicaba
  // la palabra "IE" para instituciones cuyo nombre ya empieza así
  // (p. ej. "de la IE IE PRUEBA 1234").
  const ieTitulo=nombreIESinPrefijoInstitucional_(capitalizarNombreIE_(datos.institucion||""));
  const c=datos.campos||{};

  const tituloSeccion=body.appendParagraph("Perfil de los participantes y percepción del Foro");
  tituloSeccion.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  tituloSeccion.editAsText().setForegroundColor(estilos.VERDE).setBold(true);

  if(!asistentes.length){
    const vacio=body.appendParagraph("No hay firmas de asistencia por código QR registradas para calcular el perfil demográfico ni la percepción del Foro en esta jornada.");
    vacio.editAsText().setForegroundColor(estilos.GRIS_TEXTO);
    return;
  }

  /*
   * Párrafo de consolidación: quién diligenció el formulario, con
   * cuántos participantes y con cuántas firmas de asistencia.
   */
  const nombreResponsable=String(c.nombre?.valor||"quien diligenció el formulario").trim();
  const fechaHoy=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd 'de' MMMM 'de' yyyy 'a las' HH:mm");
  const totalCaracterizacion=totalParticipantesServer_(datos);
  const pConsolidado=body.appendParagraph(
    "El presente informe fue consolidado por "+nombreResponsable+" el día "+fechaHoy+", siguiendo las indicaciones dadas por la Secretaría de Educación de Neiva (SEM Neiva) y en cooperación con "+totalCaracterizacion+" integrantes de la comunidad académica de la IE "+ieTitulo+", con una asistencia registrada de "+asistentes.length+" personas que firmaron por código QR."
  );
  pConsolidado.editAsText().setForegroundColor(estilos.GRIS_TEXTO);

  /*
   * Desglose por sede (a partir de la respuesta a la pregunta de
   * condición en la firma de asistencia por QR).
   */
  const conteoSedes={};
  asistentes.forEach(function(p){ if(p.sede) conteoSedes[p.sede]=(conteoSedes[p.sede]||0)+1; });
  const sedesOrdenadas=Object.keys(conteoSedes).sort(function(a,b){ return conteoSedes[b]-conteoSedes[a]; });
  if(sedesOrdenadas.length){
    const listaSedes=sedesOrdenadas.map(function(s){ return conteoSedes[s]+" a la sede "+s; }).join(", ");
    const pSedes=body.appendParagraph("De los participantes, "+listaSedes+".");
    pSedes.editAsText().setForegroundColor(estilos.GRIS_TEXTO);
    try{
      const blobSedes=construirGraficoColumnas_("Participantes por sede",sedesOrdenadas,sedesOrdenadas.map(function(s){return conteoSedes[s];}));
      body.appendImage(blobSedes).setWidth(430);
    }catch(errorSedes){ Logger.log("Gráfico de sedes: "+errorSedes.message); }
  }

  /*
   * Desglose por jornada.
   */
  const conteoJornadas={};
  asistentes.forEach(function(p){ if(p.jornada) conteoJornadas[p.jornada]=(conteoJornadas[p.jornada]||0)+1; });
  const jornadasOrdenadas=Object.keys(conteoJornadas);
  if(jornadasOrdenadas.length){
    try{
      const blobJornadas=construirGraficoColumnas_("Participantes por jornada",jornadasOrdenadas,jornadasOrdenadas.map(function(j){return conteoJornadas[j];}));
      body.appendImage(blobJornadas).setWidth(430);
    }catch(errorJornadas){ Logger.log("Gráfico de jornadas: "+errorJornadas.message); }
  }

  /*
   * Desglose por sexo y edad.
   */
  const dem=calcularDemografiaAsistentes_(asistentes);
  const totalHombres=dem.hombresAdultos+dem.adolescentesHombres+dem.ninos;
  const totalMujeres=dem.mujeresAdultas+dem.adolescentesMujeres+dem.ninas;
  const pDemografia=body.appendParagraph(
    "En cuanto a la composición demográfica de los asistentes, se registraron "+dem.hombresAdultos+" hombres adultos y "+dem.mujeresAdultas+" mujeres adultas (mayores de 18 años); "+dem.adolescentesHombres+" adolescentes hombres y "+dem.adolescentesMujeres+" adolescentes mujeres (entre los 13 y los 18 años); y "+dem.ninos+" niños y "+dem.ninas+" niñas (entre los 0 y los 12 años)."+
    (dem.otro?" Adicionalmente, "+dem.otro+" personas seleccionaron la opción \"Prefiero no decirlo\" en la pregunta de sexo.":"")+
    (dem.noResponde?" "+dem.noResponde+" persona"+(dem.noResponde===1?"":"s")+" eligió la opción \"Prefiero no responder\" en la pregunta de edad, por lo que no se cuenta con datos de percepción demográfica de "+(dem.noResponde===1?"esa persona":"esas personas")+".":"")
  );
  pDemografia.editAsText().setForegroundColor(estilos.GRIS_TEXTO);
  try{
    const etiquetasSexo=["Niños","Niñas","Adolescentes hombres","Adolescentes mujeres","Hombres adultos","Mujeres adultas"];
    const valoresSexo=[dem.ninos,dem.ninas,dem.adolescentesHombres,dem.adolescentesMujeres,dem.hombresAdultos,dem.mujeresAdultas];
    const blobSexo=construirGraficoColumnas_("Participantes por sexo y edad ("+(totalHombres+totalMujeres)+" total)",etiquetasSexo,valoresSexo);
    body.appendImage(blobSexo).setWidth(430);
  }catch(errorSexo){ Logger.log("Gráfico de sexo/edad: "+errorSexo.message); }

  /*
   * Percepciones por categoría de edad: niños, adolescentes, adultos,
   * y percepción general (a partir de las respuestas abiertas de la
   * valoración de la actividad, citadas de forma literal). Quienes
   * eligieron "Prefiero no responder" ya quedaron contabilizados
   * arriba, pero no entran en ningún grupo de percepción por edad.
   */
  const ninosYNinas=asistentes.filter(function(p){ return categoriaEdad_(p.edad)==="nino"; });
  const adolescentes=asistentes.filter(function(p){ return categoriaEdad_(p.edad)==="adolescente"; });
  const adultos=asistentes.filter(function(p){ return categoriaEdad_(p.edad)==="adulto"; });

  const tituloPercepcion=body.appendParagraph("Percepción de la comunidad educativa sobre el Foro Educativo Institucional");
  tituloPercepcion.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  tituloPercepcion.editAsText().setForegroundColor(estilos.VERDE).setBold(true);

  if(ninosYNinas.length){
    const tF=tallyOpciones_(ninosYNinas,"fortalezas"), tD=tallyOpciones_(ninosYNinas,"dificultades");
    const pNinos=body.appendParagraph(
      "Percepción de los niños y las niñas de la IE "+ieTitulo+" sobre el Foro Educativo Institucional\n\n"+
      "Desde la perspectiva de los niños y niñas, en la IE "+ieTitulo+" se valora especialmente "+top3Texto_(tF)+
      ". No obstante, también manifiestan dificultades relacionadas con "+top3Texto_(tD)+
      ", las cuales constituyen oportunidades para fortalecer la experiencia educativa y la vida escolar."
    );
    pNinos.editAsText().setForegroundColor(estilos.GRIS_TEXTO);
  }

  if(adolescentes.length){
    const tF=tallyOpciones_(adolescentes,"fortalezas"), tD=tallyOpciones_(adolescentes,"dificultades");
    const pAdolescentes=body.appendParagraph(
      "Percepción de los y las adolescentes de la IE "+ieTitulo+" sobre el Foro Educativo Institucional\n\n"+
      "Los y las adolescentes de la IE "+ieTitulo+" reconocen que la institución promueve "+top3Texto_(tF)+
      ". De igual manera, identifican "+top3Texto_(tD)+
      " como aspectos que representan oportunidades para el mejoramiento institucional."
    );
    pAdolescentes.editAsText().setForegroundColor(estilos.GRIS_TEXTO);
  }

  if(adultos.length){
    const tF=tallyOpciones_(adultos,"fortalezas"), tD=tallyOpciones_(adultos,"dificultades");
    const pAdultos=body.appendParagraph(
      "Percepción de los adultos de la IE "+ieTitulo+" sobre el Foro Educativo Institucional\n\n"+
      "De acuerdo con las respuestas de los adultos participantes, la IE "+ieTitulo+" se destaca por "+top3Texto_(tF)+
      ". A su vez, señalan "+top3Texto_(tD)+
      " como aspectos que requieren atención y pueden orientar acciones de mejoramiento institucional."
    );
    pAdultos.editAsText().setForegroundColor(estilos.GRIS_TEXTO);
  }

  /*
   * Percepción general: se cita LITERALMENTE (sin parafrasear) lo
   * que la persona responsable escribió en la pregunta abierta de
   * la valoración de la actividad (P5: sugerencias/comentarios).
   */
  const sugerenciasValoracion=obtenerSugerenciasValoracion_(idForo);
  const pGeneralTitulo=body.appendParagraph("Percepción general de la comunidad de la IE "+ieTitulo+" sobre el Foro Educativo Institucional");
  pGeneralTitulo.editAsText().setBold(true).setForegroundColor(estilos.GRIS_TEXTO);
  const pGeneral=body.appendParagraph(sugerenciasValoracion||"La comunidad educativa no registró comentarios adicionales en la valoración de la actividad.");
  pGeneral.editAsText().setForegroundColor(estilos.GRIS_TEXTO).setItalic(!!sugerenciasValoracion);

  /*
   * Gráficos de fortalezas y dificultades institucionales (top 5,
   * todas las personas), con las respuestas "Otro" listadas de
   * forma literal debajo de cada gráfico.
   */
  const tituloFortalezas=body.appendParagraph("Fortalezas institucionales identificadas en el Foro");
  tituloFortalezas.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  tituloFortalezas.editAsText().setForegroundColor(estilos.VERDE).setBold(true);
  agregarGraficoConOtro_(body,estilos,"Fortalezas más votadas",tallyOpciones_(asistentes,"fortalezas"),
    asistentes.map(function(p){return p.fortalezaOtro;}).filter(Boolean));

  const tituloDificultades=body.appendParagraph("Oportunidades de mejoramiento institucional identificadas en el Foro");
  tituloDificultades.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  tituloDificultades.editAsText().setForegroundColor(estilos.VERDE).setBold(true);
  agregarGraficoConOtro_(body,estilos,"Aspectos de mejora más votados",tallyOpciones_(asistentes,"dificultades"),
    asistentes.map(function(p){return p.dificultadOtro;}).filter(Boolean));

  body.appendPageBreak();
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
  const sedesIE=obtenerSedesDeIE_(ie);

  /*
   * El registro de asistencia por QR queda disponible de forma
   * permanente: no se cierra al enviar la Sesión 3. El informe
   * ejecutivo toma el listado de firmas en el momento en que se
   * genera, pero la página de firma sigue abierta después de eso
   * por si llegan más asistentes.
   */

  const escAttr_=function(t){ return String(t||"").replace(/"/g,"&quot;"); };
  const escHtml_=function(t){ return String(t||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); };

  const opcionesDe_=function(lista){
    return lista.map(function(v){ return '<option value="'+escAttr_(v)+'">'+escHtml_(v)+'</option>'; }).join("");
  };

  const opcionesCargo=opcionesDe_(CARGOS_ASISTENCIA_QR);
  const opcionesTipoAsistencia=opcionesDe_(TIPOS_ASISTENCIA_QR);
  const opcionesRolForo=opcionesDe_(ROLES_FORO_QR);
  const opcionesSexo=opcionesDe_(SEXOS_ASISTENCIA_QR);
  const opcionesJornada=opcionesDe_(JORNADAS_ASISTENCIA_QR);
  const opcionesSede=opcionesDe_(sedesIE);

  /*
   * Edad por RANGOS (ya no un número exacto 0-99): 0-12 alimenta el
   * análisis de "niños/niñas", 13-18 el de "adolescentes", y las 6
   * franjas de 18 en adelante el de "adultos" en el informe (ver
   * categoriaEdad_/agregarPerfilYPercepcionAlInforme_). "Prefiero no
   * responder" queda registrado aparte: el informe indica cuántas
   * personas eligieron esa opción, sin intentar clasificarlas.
   */
  const opcionesEdad=
    '<option value="0-12">0 a 12 años</option>'+
    '<option value="13-18">13 a 18 años</option>'+
    '<option value="18-25">18 a 25 años</option>'+
    '<option value="25-35">25 a 35 años</option>'+
    '<option value="35-45">35 a 45 años</option>'+
    '<option value="45-55">45 a 55 años</option>'+
    '<option value="55-65">55 a 65 años</option>'+
    '<option value="65+">Más de 65 años</option>'+
    '<option value="no_responde">Prefiero no responder</option>';

  const checklistDe_=function(lista, clase){
    return lista.map(function(v,i){
      return '<label class="opcionChecklist"><input type="checkbox" class="'+clase+'" value="'+escAttr_(v)+'"> '+escHtml_(v)+'</label>';
    }).join("")+
    '<label class="opcionChecklist"><input type="checkbox" class="'+clase+'" value="Otro" data-otro="1"> Otro</label>';
  };

  const checklistFortalezas=checklistDe_(FORTALEZAS_ASISTENCIA_QR,"checkFortaleza");
  const checklistDificultades=checklistDe_(DIFICULTADES_ASISTENCIA_QR,"checkDificultad");

  const cargosSinCondicionJSON=JSON.stringify(CARGOS_SIN_CONDICION_QR);

  // {{IE}} siempre aparece en el texto precedido de "la Institución
  // Educativa {{IE}}": se usa la versión sin prefijo para no duplicar
  // esa frase en instituciones cuyo nombre ya empieza con "IE"/
  // "Institución Educativa".
  const textoConsentimiento=TEXTO_CONSENTIMIENTO_ASISTENCIA_QR.split("{{IE}}").join(nombreIESinPrefijoInstitucional_(ieTitulo));

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
    'input,select{width:100%;padding:12px;font-size:16px;border:1px solid #DADCE0;border-radius:8px;box-sizing:border-box;font-family:inherit;transition:border-color .15s ease, background-color .15s ease;}'+
    'button{width:100%;margin-top:22px;padding:14px;font-size:17px;background:#0B6A44;color:#fff;border:none;border-radius:10px;cursor:pointer;}'+
    'button:disabled{background:#bdbdbd;}'+
    '#estado{margin-top:14px;font-weight:600;min-height:20px;}'+
    '#textoFirma{margin-top:6px;font-size:12px;font-weight:400;color:#4A4A4A;}'+
    '.logoAsistenciaIE{display:block;max-width:64px;max-height:64px;margin:0 auto 10px;border-radius:8px;}'+
    // Un campo ya diligenciado se pone verde (en vez de quedarse
    // gris), para que se note de un vistazo qué falta por llenar.
    '.campoCompletado{border-color:#0B6A44;background:#F7FAF7;}'+
    '.correoInvalido{border-color:#C62828 !important;background:#FFFDE7;}'+
    '.mensajeErrorCorreo{display:none;margin-top:6px;}'+
    '.mensajeErrorCorreo b{background:#FFF3CD;color:#C62828;font-weight:600;padding:3px 8px;border-radius:6px;display:inline-block;}'+
    '.mensajeErrorCorreo i{font-style:normal;font-size:11px;color:#555;margin-left:6px;}'+
    '.consentimientoBox{background:#F7FAF7;border:1px solid #DCE7DD;border-left:4px solid #F4B400;border-radius:10px;padding:14px;margin-top:10px;}'+
    '.botonOjoConsentimiento{width:auto;margin:0;padding:8px 14px;font-size:14px;background:#fff;color:#0B6A44;border:1px solid #0B6A44;border-radius:8px;cursor:pointer;}'+
    '.textoConsentimientoQR{display:none;white-space:pre-wrap;font-size:13px;line-height:1.55;max-height:280px;overflow-y:auto;background:#fff;border:1px solid #DADCE0;border-radius:8px;padding:12px;margin-top:10px;}'+
    '.labelCheckConsentimiento{display:flex;align-items:flex-start;gap:8px;font-weight:600;color:#333;margin-top:12px;font-size:14px;}'+
    '.labelCheckConsentimiento input{width:auto;margin-top:3px;}'+
    '.condicionOculta{display:none;}'+
    '.checklistBox{border:1px solid #DADCE0;border-radius:8px;padding:10px 12px;background:#fff;}'+
    '.opcionChecklist{display:flex;align-items:flex-start;gap:8px;font-weight:400;color:#4A4A4A;margin:0;padding:6px 0;font-size:14px;}'+
    '.opcionChecklist input{width:auto;margin-top:3px;}'+
    '.avisoMaximo{font-size:12px;color:#0B6A44;font-weight:600;margin-top:14px;margin-bottom:4px;}'+
    '.otroTextoOculto{display:none;}'+
    '</style></head><body>'+
    '<div class="tarjeta">'+
    (logoUrlIE?'<img src="'+logoUrlIE+'" alt="Logo de la institución educativa" class="logoAsistenciaIE">':'')+
    '<h1>'+tituloPagina.replace(/</g,"&lt;")+'</h1>'+
    '<p>Foro Educativo Institucional — Neiva 2026</p>'+
    '<div id="formulario">'+

    '<div class="consentimientoBox">'+
    '<button type="button" id="btnVerConsentimiento" class="botonOjoConsentimiento">👁️ Ver el texto completo del consentimiento</button>'+
    '<div id="textoConsentimientoQR" class="textoConsentimientoQR">'+escHtml_(textoConsentimiento)+'</div>'+
    '<label class="labelCheckConsentimiento"><input type="checkbox" id="aceptoConsentimiento"> He leído y acepto el tratamiento de mis datos personales.</label>'+
    '</div>'+

    '<label>Institución Educativa</label>'+
    '<input id="ie" value="'+String(ie).replace(/"/g,"&quot;")+'" readonly>'+
    '<label>Nombre completo</label>'+
    '<input id="nombre" autocomplete="name">'+
    '<label>Sexo</label>'+
    '<select id="sexo"><option value="">Seleccione…</option>'+opcionesSexo+'</select>'+
    '<label>Edad</label>'+
    '<select id="edad"><option value="">Seleccione…</option>'+opcionesEdad+'</select>'+
    '<button type="button" id="btnPorQueEdad" class="botonOjoConsentimiento">❓ ¿Por qué esta pregunta?</button>'+
    '<div id="textoPorQueEdad" class="textoConsentimientoQR">'+
        'Esta pregunta alimenta el análisis de percepción demográfico que se incluye en el informe del Foro Educativo Institucional de la IE '+escHtml_(nombreIESinPrefijoInstitucional_(ieTitulo))+': conocer, de forma general, cuántos niños, niñas, adolescentes y adultos participaron, para relacionar sus respuestas de fortalezas y dificultades con su rango de edad.\n\n'+
        'Si prefiere no compartir este dato, puede elegir la opción "Prefiero no responder" en la lista de arriba — el informe indicará cuántas personas eligieron no responder esta pregunta, sin que eso afecte el resto de su registro de asistencia.'+
    '</div>'+
    '<label>Cargo en la Institución Educativa</label>'+
    '<select id="cargo"><option value="">Seleccione…</option>'+opcionesCargo+'</select>'+

    '<div id="bloqueCondicion" class="condicionOculta">'+
    '<label>Jornada a la que pertenece</label>'+
    '<select id="jornada"><option value="">Seleccione…</option>'+opcionesJornada+'</select>'+
    '<label>Sede a la que pertenece</label>'+
    '<select id="sede"><option value="">Seleccione…</option>'+opcionesSede+'</select>'+
    '</div>'+

    '<label>Rol que desempeñó en el Foro Educativo Institucional '+ieTitulo.replace(/</g,"&lt;")+'</label>'+
    '<select id="rolForo"><option value="">Seleccione…</option>'+opcionesRolForo+'</select>'+

    '<label>Su asistencia fue</label>'+
    '<select id="tipoAsistencia"><option value="">Seleccione…</option>'+opcionesTipoAsistencia+'</select>'+

    '<label>Número de documento</label>'+
    '<input id="documento" inputmode="numeric" autocomplete="off">'+
    '<label>Correo electrónico</label>'+
    '<input id="correo" type="email" autocomplete="email">'+
    '<div class="mensajeErrorCorreo" id="mensajeErrorCorreo"><b>Ingrese un correo electrónico válido</b><i>(ej: nombre@dominio.com — sin espacios al inicio, al final o en medio)</i></div>'+
    '<label>Teléfono (opcional)</label>'+
    '<input id="telefono" type="tel" autocomplete="tel">'+

    '<label id="labelFortalezas">En mi papel como <span id="rolEnLabelFortalezas">(seleccione su rol arriba)</span>, ¿qué fortalezas identifiqué en este Foro Educativo en nuestra IE '+nombreIESinPrefijoInstitucional_(ieTitulo).replace(/</g,"&lt;")+'? Seleccione máximo 3.</label>'+
    '<div class="checklistBox" id="listaFortalezas">'+checklistFortalezas+'</div>'+
    '<input type="text" id="fortalezaOtroTexto" class="otroTextoOculto" placeholder="Especifique la fortaleza...">'+

    '<label>¿Qué aspectos pueden ser oportunidad de mejora institucional en la IE '+nombreIESinPrefijoInstitucional_(ieTitulo).replace(/</g,"&lt;")+'? Seleccione máximo 3.</label>'+
    '<div class="checklistBox" id="listaDificultades">'+checklistDificultades+'</div>'+
    '<input type="text" id="dificultadOtroTexto" class="otroTextoOculto" placeholder="Especifique la dificultad...">'+

    '<button id="btnFirmar" type="button" disabled>Firmar asistencia</button>'+
    '</div>'+
    '<div id="estado"></div>'+
    '<div id="textoFirma"></div>'+
    // Varias personas pueden firmar seguidas desde el mismo
    // dispositivo (un tablet o celular compartido): este botón, oculto
    // hasta la primera firma, limpia el formulario para la siguiente
    // persona sin tener que volver a escanear el QR.
    '<button id="btnFirmarOtra" type="button" style="display:none;background:#fff;color:#0B6A44;border:1px solid #0B6A44;">➕ Firmar otra asistencia</button>'+
    '</div>'+
    '<script>'+
    'var CARGOS_SIN_CONDICION='+cargosSinCondicionJSON+';'+
    'document.getElementById("btnVerConsentimiento").addEventListener("click",function(){'+
    'var d=document.getElementById("textoConsentimientoQR");'+
    'var visible=d.style.display==="block";'+
    'd.style.display=visible?"none":"block";'+
    'this.textContent=visible?"👁️ Ver el texto completo del consentimiento":"🙈 Ocultar el texto del consentimiento";'+
    '});'+
    'document.getElementById("btnPorQueEdad").addEventListener("click",function(){'+
    'var d=document.getElementById("textoPorQueEdad");'+
    'var visible=d.style.display==="block";'+
    'd.style.display=visible?"none":"block";'+
    'this.textContent=visible?"❓ ¿Por qué esta pregunta?":"🙈 Ocultar explicación";'+
    '});'+
    'function actualizarBotonFirmar(){'+
    'var acepto=document.getElementById("aceptoConsentimiento").checked;'+
    'document.getElementById("btnFirmar").disabled=!acepto;'+
    '}'+
    'document.getElementById("aceptoConsentimiento").addEventListener("change",actualizarBotonFirmar);'+
    /*
     * Un campo ya diligenciado se pone verde (border-color/fondo) en
     * vez de quedarse gris — igual idea que el resto del formulario
     * principal. No aplica a checkboxes ni a inputs ocultos, y
     * respeta el estado rojo del correo cuando no es válido (esa
     * regla usa !important, así que gana igual aunque las dos
     * clases estén presentes a la vez).
     */
    'function marcarCampoCompletado_(campo){'+
    'var lleno=!!(campo.value&&String(campo.value).trim());'+
    'campo.classList.toggle("campoCompletado",lleno);'+
    '}'+
    'document.querySelectorAll("#formulario input, #formulario select").forEach(function(campo){'+
    'if(campo.type==="checkbox"||campo.type==="hidden") return;'+
    'marcarCampoCompletado_(campo);'+
    'campo.addEventListener("input",function(){ marcarCampoCompletado_(campo); });'+
    'campo.addEventListener("change",function(){ marcarCampoCompletado_(campo); });'+
    '});'+
    'document.getElementById("cargo").addEventListener("change",function(){'+
    'var requiere=CARGOS_SIN_CONDICION.indexOf(this.value)===-1 && this.value!=="";'+
    'document.getElementById("bloqueCondicion").classList.toggle("condicionOculta",!requiere);'+
    '});'+
    /*
     * El texto de la pregunta de fortalezas ("En mi papel como...")
     * se completa con el rol que la persona seleccionó arriba, ya
     * que esta página se genera antes de que exista esa respuesta.
     */
    'document.getElementById("rolForo").addEventListener("change",function(){'+
    'var span=document.getElementById("rolEnLabelFortalezas");'+
    'if(!span) return;'+
    'var texto=String(this.value||"").replace(/^[^\\p{L}]+/u,"").trim();'+
    'span.textContent=texto||"(seleccione su rol arriba)";'+
    '});'+
    'function activarChecklistMaximo3(claseCheck,idOtroTexto){'+
    'var checks=document.querySelectorAll("."+claseCheck);'+
    'checks.forEach(function(c){'+
    'c.addEventListener("change",function(){'+
    'var marcados=Array.prototype.slice.call(document.querySelectorAll("."+claseCheck+":checked"));'+
    'if(marcados.length>3){ this.checked=false; return; }'+
    'var otro=document.querySelector("."+claseCheck+"[data-otro=\\"1\\"]");'+
    'var campoOtro=document.getElementById(idOtroTexto);'+
    'if(otro&&campoOtro){'+
    'if(otro.checked){ campoOtro.classList.remove("otroTextoOculto"); }'+
    'else{ campoOtro.classList.add("otroTextoOculto"); campoOtro.value=""; }'+
    '}'+
    '});'+
    '});'+
    'var campoOtroInicial=document.getElementById(idOtroTexto);'+
    'if(campoOtroInicial){'+
    'campoOtroInicial.addEventListener("input",function(){'+
    'if(!this.value.trim()){'+
    'var otro=document.querySelector("."+claseCheck+"[data-otro=\\"1\\"]");'+
    'if(otro) otro.checked=false;'+
    'this.classList.add("otroTextoOculto");'+
    '}'+
    '});'+
    '}'+
    '}'+
    'activarChecklistMaximo3("checkFortaleza","fortalezaOtroTexto");'+
    'activarChecklistMaximo3("checkDificultad","dificultadOtroTexto");'+
    'function valoresMarcados(clase){'+
    'return Array.prototype.slice.call(document.querySelectorAll("."+clase+":checked")).map(function(c){return c.value;});'+
    '}'+
    /*
     * dispositivoIdAsistencia se sigue calculando y enviando al
     * servidor, pero solo con fines de auditoría/analítica: YA NO
     * bloquea firmas repetidas desde el mismo equipo. Es habitual que
     * varias personas firmen seguidas desde un mismo dispositivo
     * compartido (un tablet institucional, un solo celular que se
     * pasa de mano en mano), así que el formulario se puede volver a
     * diligenciar tantas veces como haga falta en la misma visita —
     * ver el botón "➕ Firmar otra asistencia" en mostrarYaFirmado().
     */
    'function calcularHuellaDispositivo(){'+
    'try{'+
    'var partes=[navigator.userAgent||"",navigator.language||"",(screen.width||"")+"x"+(screen.height||""),screen.colorDepth||"",navigator.hardwareConcurrency||"",navigator.platform||"",(Intl.DateTimeFormat().resolvedOptions().timeZone)||""].join("|");'+
    'var hash=0;'+
    'for(var i=0;i<partes.length;i++){ hash=((hash<<5)-hash+partes.charCodeAt(i))|0; }'+
    'return "fp-"+Math.abs(hash).toString(36);'+
    '}catch(e){ return "fp-desconocida"; }'+
    '}'+
    'function mostrarYaFirmado(nombrePersona,textoFirma){'+
    'document.getElementById("formulario").style.display="none";'+
    'var saludo=nombrePersona?(nombrePersona+", su asistencia está firmada!"):"Su asistencia está firmada.";'+
    'document.getElementById("estado").textContent="✓ "+saludo+" Ya puede cerrar esta página y continuar en la plenaria, o firmar la asistencia de otra persona desde este mismo dispositivo.";'+
    'document.getElementById("textoFirma").textContent=textoFirma||"";'+
    'document.getElementById("btnFirmarOtra").style.display="block";'+
    'window.scrollTo(0,0);'+
    '}'+
    /*
     * Deja el formulario listo para la siguiente persona: se limpian
     * todos los campos y checkboxes, se vuelve a mostrar el
     * formulario y se oculta este mismo botón hasta la próxima firma.
     */
    'function limpiarFormularioFirma(){'+
    '["nombre","documento","correo","telefono","sexo","edad","cargo","jornada","sede","rolForo","tipoAsistencia"].forEach(function(id){'+
    'var c=document.getElementById(id); if(!c) return; c.value=""; c.classList.remove("campoCompletado");'+
    '});'+
    'document.querySelectorAll(".checkFortaleza,.checkDificultad").forEach(function(c){ c.checked=false; });'+
    '["fortalezaOtroTexto","dificultadOtroTexto"].forEach(function(id){ var c=document.getElementById(id); if(!c) return; c.value=""; c.classList.add("otroTextoOculto"); });'+
    'document.getElementById("aceptoConsentimiento").checked=false;'+
    'document.getElementById("bloqueCondicion").classList.add("condicionOculta");'+
    'document.getElementById("rolEnLabelFortalezas").textContent="(seleccione su rol arriba)";'+
    'document.getElementById("correo").classList.remove("correoInvalido");'+
    'document.getElementById("mensajeErrorCorreo").style.display="none";'+
    'document.getElementById("btnFirmar").disabled=true;'+
    'document.getElementById("btnFirmar").textContent="Firmar asistencia";'+
    'document.getElementById("btnFirmarOtra").style.display="none";'+
    'document.getElementById("estado").textContent="";'+
    'document.getElementById("textoFirma").textContent="";'+
    'document.getElementById("formulario").style.display="";'+
    'window.scrollTo(0,0);'+
    '}'+
    'document.getElementById("btnFirmarOtra").addEventListener("click",limpiarFormularioFirma);'+
    'var dispositivoIdAsistencia=calcularHuellaDispositivo();'+
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
    /*
     * Autocompletado de dominio: @g -> @gmail.com, @h -> @hotmail.com,
     * y "usuario@dominio." -> "usuario@dominio.com". Solo se activa
     * cuando la persona está ESCRIBIENDO hacia adelante (inputType
     * que empieza por "insert"): si se ignora esto, al borrar con
     * backspace se vuelve a completar apenas el texto queda de nuevo
     * en "@g"/"@h", y el campo parece que no se puede borrar.
     */
    'document.getElementById("correo").addEventListener("input",function(e){'+
    'var v=this.value;'+
    'var esBorrado=e&&e.inputType&&e.inputType.indexOf("delete")===0;'+
    'if(!esBorrado){'+
    'if(/@g$/i.test(v)){ this.value=v+"mail.com"; }'+
    'else if(/@h$/i.test(v)){ this.value=v+"otmail.com"; }'+
    'else if(/@[^@\\s.]+\\.$/.test(v)){ this.value=v+"com"; }'+
    '}'+
    'validarCorreoUI();'+
    '});'+
    'document.getElementById("correo").addEventListener("blur",validarCorreoUI);'+
    'document.getElementById("btnFirmar").addEventListener("click",function(){'+
    'var btn=this; var estado=document.getElementById("estado");'+
    'var nombre=document.getElementById("nombre").value.trim();'+
    'var sexo=document.getElementById("sexo").value.trim();'+
    'var edad=document.getElementById("edad").value.trim();'+
    'var tipoAsistencia=document.getElementById("tipoAsistencia").value.trim();'+
    'var cargo=document.getElementById("cargo").value.trim();'+
    'var rolForo=document.getElementById("rolForo").value.trim();'+
    'var requiereCondicion=CARGOS_SIN_CONDICION.indexOf(cargo)===-1;'+
    'var jornada=requiereCondicion?document.getElementById("jornada").value.trim():"";'+
    'var sede=requiereCondicion?document.getElementById("sede").value.trim():"";'+
    'var documento=document.getElementById("documento").value.trim();'+
    'var correo=document.getElementById("correo").value.trim();'+
    'var telefono=document.getElementById("telefono").value.trim();'+
    'var acepto=document.getElementById("aceptoConsentimiento").checked;'+
    'var fortalezas=valoresMarcados("checkFortaleza");'+
    'var fortalezaOtro=document.getElementById("fortalezaOtroTexto").value.trim();'+
    'var dificultades=valoresMarcados("checkDificultad");'+
    'var dificultadOtro=document.getElementById("dificultadOtroTexto").value.trim();'+
    'if(!acepto){estado.textContent="Debe aceptar el tratamiento de sus datos personales para continuar.";return;}'+
    'if(!nombre||!sexo||!edad||!tipoAsistencia||!cargo||!rolForo||!documento||!correo){estado.textContent="Complete nombre, sexo, edad, tipo de asistencia, cargo, rol en el Foro, número de documento y correo electrónico.";return;}'+
    'if(requiereCondicion&&(!jornada||!sede)){estado.textContent="Seleccione la jornada y la sede a la que pertenece.";return;}'+
    'if(fortalezas.length>3){estado.textContent="Seleccione máximo 3 fortalezas.";return;}'+
    'if(dificultades.length>3){estado.textContent="Seleccione máximo 3 aspectos de mejora.";return;}'+
    'if(!correoEsValido(correo)){validarCorreoUI();estado.textContent="Revise el correo electrónico: no es válido.";return;}'+
    'btn.disabled=true; btn.textContent="Firmando…";'+
    'google.script.run.withSuccessHandler(function(res){'+
    'if(res&&res.ok){ mostrarYaFirmado(nombre,res.textoFirma); }'+
    'else{ btn.disabled=false; btn.textContent="Firmar asistencia"; estado.textContent=(res&&res.mensaje)||"No fue posible registrar la asistencia."; }'+
    '}).withFailureHandler(function(err){ btn.disabled=false; btn.textContent="Firmar asistencia"; estado.textContent="No fue posible registrar la asistencia: "+(err.message||err); })'+
    '.registrarAsistenciaQR('+JSON.stringify(idForo)+',nombre,sexo,edad,tipoAsistencia,cargo,rolForo,jornada,sede,fortalezas,fortalezaOtro,dificultades,dificultadOtro,documento,correo,telefono,acepto,dispositivoIdAsistencia);'+
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
  return {ok:true,foto:{id:foto.getId(),url:foto.getUrl()},folderId:folder.getId()};
}

/*
 * Sube la asistencia como PDF escaneado — método alternativo al
 * código QR. Se guarda en la misma carpeta de la IE que la fotografía
 * y el informe. Quien llama debe además guardar, entre los campos
 * del formulario, metodoAsistencia="PDF" y numeroAsistentesPDF, y
 * estos mismos (asistenciaPdfUrl/asistenciaPdfId) — generarInformeFEM
 * los lee de ahí para decidir si incluye el listado/análisis QR o el
 * aviso + enlace de PDF.
 */
function subirAsistenciaPDF(idForo,pdfData,pdfName,datos){
  const acceso=obtenerAccesoPorIdForo_(idForo); if(!acceso)throw new Error("ID_FORO no autorizado.");
  const folder=crearCarpetaIE_(datos.institucion||acceso.ie);
  const decode=(data)=>{const s=String(data||"");const comma=s.indexOf(",");return Utilities.base64Decode(comma>=0?s.substring(comma+1):s);};
  const pb=decode(pdfData);
  if(pb.length>15*1024*1024) throw new Error("El PDF de asistencia debe pesar máximo 15 MB.");
  const nombreArchivo="Asistencia Foro Educativo - "+(datos.institucion||acceso.ie)+" FEM 2026.pdf";
  const pdf=folder.createFile(Utilities.newBlob(pb,"application/pdf",nombreArchivo));
  hacerPublicoSiEsPosible_(pdf);
  return {ok:true, pdf:{id:pdf.getId(), url:pdf.getUrl()}};
}


function construirParrafoSesion_(titulo,contenido){return titulo+"\n\n"+String(contenido||"");}

/*
 * Lee la Sesión Propia (Sesión 4) desde datos.campos: título,
 * objetivo, y el JSON de líneas temáticas (cada una con sus
 * preguntas/respuestas). Nunca lanza error — un JSON inválido o
 * ausente se trata como "sin líneas".
 */
function obtenerSesionPropia_(datos){
  const campos=datos?.campos||{};
  const titulo=String(campos.tituloSesionPropia?.valor||"").trim();
  const objetivo=String(campos.objetivoSesionPropia?.valor||"").trim();
  let lineas=[];
  try{
    const parsed=JSON.parse(campos.sesionPropiaLineasJSON?.valor||"[]");
    if(Array.isArray(parsed)) lineas=parsed;
  }catch(e){ lineas=[]; }
  // Solo cuentan las líneas con título o al menos una pregunta con texto.
  lineas=lineas.filter(function(l){
    return String(l?.titulo||"").trim() || (Array.isArray(l?.preguntas) && l.preguntas.some(function(p){ return String(p?.texto||"").trim(); }));
  });
  const tieneContenido=!!(titulo || objetivo || lineas.length);
  return {titulo:titulo, objetivo:objetivo, lineas:lineas, tieneContenido:tieneContenido};
}

/*
 * Agrega al informe ejecutivo la "Sesión Propia creada por la IE
 * ___" (Sesión 4, opcional) — solo si hay contenido real. Incluye el
 * escudo de la IE junto al título (tabla sin bordes, mismo recurso
 * usado en el encabezado del documento), el subtítulo explicando que
 * es autónoma y no entra en las conclusiones del foro comunitario, y
 * el título/objetivo/líneas temáticas/preguntas tal como los creó la
 * institución.
 */
function agregarSesionPropiaAlInforme_(body, datos, ieSinPrefijo, estilos){
  const sesionPropia=obtenerSesionPropia_(datos);
  if(!sesionPropia.tieneContenido) return;

  const VERDE=estilos.VERDE, GRIS_TEXTO=estilos.GRIS_TEXTO;

  body.appendPageBreak();

  /*
   * Ya NO se repite el logo de la IE aquí encima del título: el
   * encabezado de 3 logos (SEM/FEM/IE) ya se repite en TODAS las
   * páginas del informe (limitación de DocumentApp, ver
   * generarInformeFEM), así que este escudo adicional quedaba
   * redundante justo en medio de la página.
   */
  const pTitulo=body.appendParagraph("Sesión Propia creada por la IE "+ieSinPrefijo);
  pTitulo.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  pTitulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pTitulo.editAsText().setForegroundColor(VERDE).setBold(true);

  const subtitulo=body.appendParagraph(
    "Las respuestas de esta sesión corresponden a las necesidades y realidades de la IE "+ieSinPrefijo+
    (sesionPropia.objetivo ? " y se hicieron con el objetivo de "+sesionPropia.objetivo : "")+
    ", de manera autónoma. Esta sesión es un complemento institucional y no se incluye en las conclusiones finales del foro comunitario."
  );
  subtitulo.editAsText().setForegroundColor(GRIS_TEXTO).setItalic(true);

  if(sesionPropia.titulo){
    const pTituloSesion=body.appendParagraph(sesionPropia.titulo);
    pTituloSesion.editAsText().setBold(true).setFontSize(14).setForegroundColor(GRIS_TEXTO);
  }
  if(sesionPropia.objetivo){
    const pObjetivoLabel=body.appendParagraph("Objetivo de la sesión");
    pObjetivoLabel.editAsText().setBold(true).setForegroundColor(VERDE);
    const pObjetivo=body.appendParagraph(sesionPropia.objetivo);
    pObjetivo.editAsText().setForegroundColor(GRIS_TEXTO);
  }

  sesionPropia.lineas.forEach(function(linea, li){
    const pLinea=body.appendParagraph("Línea temática "+(li+1)+(linea.titulo?": "+linea.titulo:""));
    pLinea.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    pLinea.editAsText().setForegroundColor(VERDE).setBold(true);

    (linea.preguntas||[]).forEach(function(pregunta, pi){
      if(!String(pregunta?.texto||"").trim()) return;
      const pPregunta=body.appendParagraph("Pregunta "+(pi+1)+": "+pregunta.texto);
      pPregunta.editAsText().setBold(true).setForegroundColor(GRIS_TEXTO);
      const pRespuesta=body.appendParagraph("Respuesta: "+(pregunta.respuesta||"—"));
      pRespuesta.editAsText().setForegroundColor(GRIS_TEXTO);
    });
  });
}


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

    const body=doc.getBody(); body.clear(); body.setPageWidth(612).setPageHeight(792).setMarginTop(18).setMarginBottom(24).setMarginLeft(36).setMarginRight(36);
    /*
     * Encabezado con los 3 logos institucionales: SEM a la izquierda,
     * FEM en el centro, IE a la derecha — IE al mismo tamaño grande
     * que el logo de la portada (100x100), FEM y SEM proporcionados
     * (el logo de la SEM es naturalmente ancho/rectangular, no
     * cuadrado: forzarlo a 100x100 lo deformaba y se veía "muy
     * largo"; se usa la misma proporción 2:1 ya usada para este mismo
     * logo en el pie de página, 80x40, escalada a 100x50).
     * Documentos/Apps Script NO permite un encabezado "distinto en la
     * primera página" (no existe esa opción en DocumentApp): este
     * mismo encabezado se repite igual en TODAS las páginas, incluida
     * la primera — no hay forma de quitarlo solo de la portada sin
     * también quitarlo de las páginas 2 en adelante, donde sí se pidió
     * expresamente.
     */
    const h=doc.getHeader()||doc.addHeader(); h.clear();
    const logoIdIE=obtenerLogoIdPorNombreIE_(datos.institucion||"");
    // Tabla sin bordes (3 columnas) para fijar SEM/FEM/IE cada uno en
    // su lugar, sin depender de tabuladores.
    const tablaEncabezado=h.appendTable([["","",""]]);
    tablaEncabezado.setBorderWidth(0);
    const celdaLogoSem=tablaEncabezado.getCell(0,0);
    celdaLogoSem.setWidth(120);
    try{ celdaLogoSem.getChild(0).asParagraph().appendInlineImage(DriveApp.getFileById(LOGO_PIE_ID).getBlob()).setWidth(100).setHeight(50); }catch(e){}
    const celdaLogoFem=tablaEncabezado.getCell(0,1);
    celdaLogoFem.setWidth(300);
    const pLogoFem=celdaLogoFem.getChild(0).asParagraph();
    pLogoFem.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    try{ pLogoFem.appendInlineImage(DriveApp.getFileById(LOGO_ENCABEZADO_ID).getBlob()).setWidth(126).setHeight(70); }catch(e){}
    const celdaLogoIE=tablaEncabezado.getCell(0,2);
    celdaLogoIE.setWidth(120);
    if(logoIdIE){
      const pLogoIEEncabezado=celdaLogoIE.getChild(0).asParagraph();
      pLogoIEEncabezado.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      try{ pLogoIEEncabezado.appendInlineImage(DriveApp.getFileById(logoIdIE).getBlob()).setWidth(100).setHeight(100); }catch(e){}
    }
    const footer=doc.getFooter()||doc.addFooter(); footer.clear(); const fp=footer.appendParagraph(""); fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER); try{fp.appendInlineImage(DriveApp.getFileById(LOGO_PIE_ID).getBlob()).setWidth(80).setHeight(40);}catch(e){};
    const fpTexto=footer.appendParagraph("Generado por SEM el "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy 'a las' HH:mm")+". Enviado por "+(datos.campos?.nombre?.valor||"")+" — "+(datos.campos?.correo?.valor||"")+" — "+(datos.campos?.cargo?.valor||"")+" de la "+(datos.institucion||""));
    fpTexto.setAlignment(DocumentApp.HorizontalAlignment.CENTER); fpTexto.editAsText().setForegroundColor(GRIS_TEXTO).setFontSize(9);

    /*
     * Título subido directamente a continuación del encabezado — ya
     * NO se repite el logo de la IE aquí como primer elemento de la
     * portada (quedaba duplicado con el que ya está en el
     * encabezado); esto libera el espacio vertical que hacía falta
     * para que el párrafo introductorio y la Caracterización cupieran
     * en una sola hoja.
     */

    // Título subido (spacingBefore en 0) y 2pt más pequeño que el
    // tamaño por defecto del estilo "Título" de Documentos (26pt),
    // para que la Caracterización quepa completa en una sola página.
    const title=body.appendParagraph("INFORME EJECUTIVO DE "+String(datos.institucion||"").toUpperCase()+" FEM 2026");
    title.setHeading(DocumentApp.ParagraphHeading.TITLE); title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.setSpacingBefore(0).setSpacingAfter(6);
    title.editAsText().setForegroundColor(VERDE).setFontSize(24);

    const subt=body.appendParagraph("FEM 2026 “Escuela Viva: Voces que construyen territorio”.");
    subt.setHeading(DocumentApp.ParagraphHeading.HEADING2); subt.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subt.setSpacingBefore(2).setSpacingAfter(2); subt.editAsText().setForegroundColor(GRIS_TEXTO).setItalic(true);

    const sub2=body.appendParagraph("Foro Educativo Institucional — Neiva 2026");
    sub2.setSpacingBefore(0).setSpacingAfter(4);
    sub2.setAlignment(DocumentApp.HorizontalAlignment.CENTER); sub2.editAsText().setForegroundColor(GRIS_TEXTO);

    body.appendHorizontalRule();

    function encabezadoSeccion_(texto){
      const p=body.appendParagraph(texto);
      p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      p.setSpacingBefore(6).setSpacingAfter(4);
      p.editAsText().setForegroundColor(VERDE).setBold(true);
      return p;
    }

    /*
     * Recorta el padding por defecto de las celdas de una tabla
     * (Documentos usa ~5pt por lado) para que las tablas de
     * Caracterización y Participación ocupen menos alto y quepan
     * cada una en una sola página.
     */
    function reducirPaddingTabla_(tabla, valor){
      for(let f=0; f<tabla.getNumRows(); f++){
        const fila=tabla.getRow(f);
        for(let col=0; col<fila.getNumCells(); col++){
          const celda=fila.getCell(col);
          celda.setPaddingTop(valor).setPaddingBottom(valor).setPaddingLeft(valor).setPaddingRight(valor);
        }
      }
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
      t.setBorderColor("#FFFFFF"); t.setBorderWidth(4);
      filas.forEach(function(x){
        const r=t.appendTableRow();
        const acento=r.appendTableCell("");
        acento.setBackgroundColor(AMARILLO);
        acento.setWidth(6);
        const contenido=r.appendTableCell("");
        contenido.setBackgroundColor(AZUL_CLARO);
        const pTitulo=contenido.getChild(0).asParagraph();
        pTitulo.setSpacingBefore(0).setSpacingAfter(0);
        pTitulo.setText(String(x[0]||""));
        pTitulo.editAsText().setBold(true).setForegroundColor(VERDE).setFontSize(9);
        const pValor=contenido.appendParagraph(String(x[1]||"—"));
        pValor.setSpacingBefore(0).setSpacingAfter(0);
        pValor.editAsText().setForegroundColor(NEGRO).setFontSize(9);
      });
      reducirPaddingTabla_(t, 3);
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

        // Cargo, con la cantidad de participantes debajo (no al
        // frente del porcentaje, como antes).
        const cEtq=r.appendTableCell("");
        cEtq.setWidth(150);
        const pEtq=cEtq.getChild(0).asParagraph();
        pEtq.setSpacingBefore(0).setSpacingAfter(0);
        pEtq.setText(etiquetas[d]);
        pEtq.editAsText().setForegroundColor(GRIS_TEXTO).setFontSize(8).setBold(true);
        const pCantidad=cEtq.appendParagraph(n+" participante"+(n===1?"":"s"));
        pCantidad.setSpacingBefore(0).setSpacingAfter(0);
        pCantidad.editAsText().setForegroundColor(GRIS_TEXTO).setFontSize(7);

        const cBarra=r.appendTableCell("");
        cBarra.setWidth(190);
        try{ insertarBarraParticipacion_(cBarra, pct); }catch(errorBarra){}

        const cValor=r.appendTableCell((pct*100).toFixed(1)+"%");
        cValor.setWidth(50);
        cValor.editAsText().setForegroundColor(GRIS_TEXTO).setFontSize(8);
      });
      reducirPaddingTabla_(t, 2);
      return t;
    }

    encabezadoSeccion_("Caracterización");
    const c=datos.campos||{};
    tablaCaracterizacion_([["Institución Educativa",datos.institucion],["DANE",datos.dane],["Rector(a)",c.rector?.valor||""],["Grupo de trabajo",c.grupo?.valor||""],["Responsable",c.nombre?.valor||""],["Cargo",c.cargo?.valor||""],["Correo responsable",c.correo?.valor||""],["Correo institucional",c.correoIE?.valor||""]]);

    const totalParticipantesInforme=totalParticipantesServer_(datos);

    // Mismo párrafo introductorio que se muestra en la portada de la
    // sesión de plenaria, justo después de la caracterización (no
    // después de la tabla de participación).
    const parrafoIntro=body.appendParagraph(
      "La institución educativa "+String(datos.institucion||"")+" construyó colectivamente las conclusiones que se presentan a continuación con la participación de "+totalParticipantesInforme+" integrantes de su comunidad educativa."
    );
    parrafoIntro.editAsText().setForegroundColor(GRIS_TEXTO);

    /*
     * Logo, título, caracterización y el párrafo introductorio quedan
     * en la primera hoja; la participación empieza en una segunda
     * hoja aparte.
     */
    body.appendPageBreak();

    encabezadoSeccion_("Participación");
    const pPart=body.appendParagraph("Participantes: "+totalParticipantesInforme);
    pPart.setHeading(DocumentApp.ParagraphHeading.HEADING2); pPart.setSpacingBefore(2).setSpacingAfter(4);
    pPart.editAsText().setForegroundColor(VERDE).setBold(true);
    tablaParticipacionDoc_(datos);

    body.appendPageBreak();

    /*
     * El apartado demográfico y de percepción (fortalezas/dificultades)
     * solo se puede construir a partir de las firmas de asistencia por
     * QR (sexo, edad, jornada, sede, fortalezas, dificultades) — un
     * PDF escaneado no trae esos datos estructurados. Si la IE eligió
     * subir la asistencia en PDF, se avisa expresamente que este
     * apartado no se genera, en vez de mostrarlo vacío o a medias.
     */
    const metodoAsistenciaInforme=String(datos.campos?.metodoAsistencia?.valor||"QR");
    if(metodoAsistenciaInforme==="PDF"){
      const tituloSinAnalisis=body.appendParagraph("Perfil de los participantes y percepción del Foro");
      tituloSinAnalisis.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      tituloSinAnalisis.editAsText().setForegroundColor(VERDE).setBold(true);
      const numeroAsistentesPdf=String(datos.campos?.numeroAsistentesPDF?.valor||"0");
      const notaSinAnalisis=body.appendParagraph(
        "La institución educativa "+nombreIESinPrefijoInstitucional_(datos.institucion||"")+" registró la asistencia mediante un PDF escaneado con "+numeroAsistentesPdf+" de participantes según el listado adjunto que se encuentra al final de este consolidado Institucional."
      );
      notaSinAnalisis.editAsText().setForegroundColor(GRIS_TEXTO);
    }else{
      try{ agregarPerfilYPercepcionAlInforme_(body, idForo, datos, {VERDE,GRIS_TEXTO,GRIS_FONDO,GRIS_BORDE,AZUL_CLARO,AMARILLO,NEGRO}); }
      catch(errorPerfil){ Logger.log("No fue posible agregar el perfil de participantes/percepción: "+errorPerfil.message); }
    }

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

    try{ agregarSesionPropiaAlInforme_(body, datos, nombreIESinPrefijoInstitucional_(datos.institucion||""), {VERDE,GRIS_TEXTO}); }
    catch(errorSesionPropia){ Logger.log("No fue posible agregar la Sesión Propia al informe: "+errorSesionPropia.message); }

    /*
     * Sin salto de página forzado: si queda espacio en la última
     * hoja de las sesiones, la sección de Evidencias continúa ahí
     * mismo en vez de empezar siempre una hoja nueva.
     */
    encabezadoSeccion_("Evidencias de la jornada");
    const textoEvidenciasAsistencia=metodoAsistenciaInforme==="PDF"
      ? "La asistencia se registró mediante un PDF escaneado, que puede verse aquí: 🔗"
      : "La asistencia se firmó de manera digital (código QR/link) durante la jornada; el listado completo se incluye a continuación.";
    const pEv=body.appendParagraph(textoEvidenciasAsistencia);
    pEv.editAsText().setForegroundColor(GRIS_TEXTO);

    // El enlace de descarga del PDF de asistencia va justo debajo del
    // aviso ("que puede verse aquí: 🔗"), no varias secciones después
    // de la fotografía como antes.
    if(metodoAsistenciaInforme==="PDF" && datos.campos?.asistenciaPdfUrl?.valor){
      const pEnlacePdfAsistencia=body.appendParagraph("");
      const rangoEnlacePdfAsistencia=pEnlacePdfAsistencia.appendText(String(datos.campos.asistenciaPdfUrl.valor));
      rangoEnlacePdfAsistencia.setLinkUrl(String(datos.campos.asistenciaPdfUrl.valor));
      rangoEnlacePdfAsistencia.setForegroundColor(VERDE);
    }

    /*
     * La fotografía se inserta como imagen dentro del cuerpo del
     * informe (no solo como enlace) con el mismo pie de página que
     * ya se muestra en la pantalla de Evidencias. La fecha/hora del
     * pie es la fecha real de creación del archivo en Drive (el
     * momento en que se subió), no la fecha de generación del
     * informe, que puede ser mucho después.
     */
    if(c.evidenciaFotoId?.valor){
      try{
        const fotoFile=DriveApp.getFileById(String(c.evidenciaFotoId.valor));
        const imagenInsertada=body.appendImage(fotoFile.getBlob());
        const anchoOriginal=imagenInsertada.getWidth(), altoOriginal=imagenInsertada.getHeight();
        const anchoMax=460;
        if(anchoOriginal>anchoMax){
          const factor=anchoMax/anchoOriginal;
          imagenInsertada.setWidth(anchoMax).setHeight(Math.round(altoOriginal*factor));
        }
        const pieFoto=body.appendParagraph(construirPieFotoEvidencia_(nombreIESinPrefijoInstitucional_(datos.institucion||""), fotoFile.getDateCreated(), totalParticipantesServer_(datos)));
        pieFoto.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        pieFoto.editAsText().setForegroundColor(GRIS_TEXTO).setItalic(true).setFontSize(10);
      }catch(errorFoto){
        Logger.log("No fue posible insertar la fotografía en el informe: "+errorFoto.message);
      }
    }

    if(c.evidenciaFotoUrl?.valor){
      const pDescarga=body.appendParagraph("📷 Descarga de fotografía: ");
      pDescarga.editAsText().setForegroundColor(GRIS_TEXTO);
      const rangoEnlace=pDescarga.appendText(String(c.evidenciaFotoUrl.valor));
      rangoEnlace.setLinkUrl(String(c.evidenciaFotoUrl.valor));
      rangoEnlace.setForegroundColor(VERDE);
    }

    if(metodoAsistenciaInforme!=="PDF"){
      agregarListadoAsistenciaAlInforme_(body, idForo, datos);
    }

    /*
     * Firmas originales al final del informe: el líder (rector/a) de
     * la IE, seguido de cada responsable del envío (el principal y
     * cualquier responsable adicional que se haya agregado en
     * Caracterización) — mismo formato para todos: nombre en negrita
     * y subrayado, cargo en cursiva, y para los responsables del
     * envío también su rol dentro del Foro en cursiva.
     */
    body.appendPageBreak();
    const tituloFirmas=body.appendParagraph("Firmas");
    tituloFirmas.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    tituloFirmas.editAsText().setForegroundColor(VERDE).setBold(true);

    function agregarBloqueFirma_(nombre, cargo, rolTexto){
      if(!String(nombre||"").trim()) return;
      const pNombre=body.appendParagraph(String(nombre).trim());
      pNombre.setSpacingBefore(18).setSpacingAfter(2);
      pNombre.editAsText().setBold(true).setUnderline(true).setForegroundColor(NEGRO);
      if(String(cargo||"").trim()){
        const pCargo=body.appendParagraph(String(cargo).trim());
        pCargo.setSpacingBefore(0).setSpacingAfter(rolTexto?2:10);
        pCargo.editAsText().setItalic(true).setForegroundColor(GRIS_TEXTO);
      }
      if(rolTexto){
        const pRol=body.appendParagraph(rolTexto);
        pRol.setSpacingBefore(0).setSpacingAfter(10);
        pRol.editAsText().setItalic(true).setForegroundColor(GRIS_TEXTO);
      }
    }

    agregarBloqueFirma_(c.rector?.valor, "Rector(a)", null);

    const ieParaFirma=nombreIESinPrefijoInstitucional_(datos.institucion||"");
    const responsablesFirma=[{nombre:c.nombre?.valor, cargo:c.cargo?.valor, rol:c.rolForoResponsable?.valor}];
    for(let i=2;i<=4;i++){
      const nombreResponsableAdicional=c["responsable"+i+"Nombre"]?.valor;
      if(!String(nombreResponsableAdicional||"").trim()) continue;
      responsablesFirma.push({nombre:nombreResponsableAdicional, cargo:c["responsable"+i+"Cargo"]?.valor, rol:c["responsable"+i+"RolForo"]?.valor});
    }
    responsablesFirma.forEach(function(r){
      const rolTexto=String(r.rol||"").trim() + " del Foro Educativo Institucional "+ieParaFirma+" 2026";
      agregarBloqueFirma_(r.nombre, r.cargo, rolTexto);
    });

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
  const acceso=obtenerAccesoPorIdForoRaw_(idForo); if(!acceso)throw new Error("ID_FORO no autorizado."); const c=datos.campos||{}; const ie=datos.institucion||acceso.ie; const ieSinPrefijo=nombreIESinPrefijoInstitucional_(ie); const logoIEUrlCorreo=urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(ie)); const logoIEHtmlCorreo=logoIEUrlCorreo?("<div style=\"text-align:center;margin:0 0 18px;\"><img src=\""+logoIEUrlCorreo+"\" alt=\"Logo de la institución educativa\" style=\"max-width:56px;max-height:56px;border-radius:8px;\"></div>"):""; const destinatario=String(c.correoIE?.valor||acceso.email||"").trim(); const responsable=String(c.correo?.valor||"").trim(); if(!destinatario)throw new Error("La institución no tiene correo institucional registrado."); const aliases=GmailApp.getAliases().map(x=>x.toLowerCase()); const cuenta=Session.getEffectiveUser().getEmail().toLowerCase(); if(cuenta!==REMITENTE_FEM&&aliases.indexOf(REMITENTE_FEM)===-1)throw new Error("La cuenta de Apps Script no puede enviar como "+REMITENTE_FEM+". Configure esa cuenta o un alias."); const file=DriveApp.getFileById(pdfId); hacerPublicoSiEsPosible_(file); const linkDescarga=file.getUrl(); const subject="Reporte de Informe IE "+ie; const body="Apreciados(as) integrantes de la comunidad educativa de la Institución Educativa "+ieSinPrefijo+":\n\nReciban un cordial saludo de la Secretaría de Educación de Neiva.\n\nAgradecemos a la Institución Educativa por su participación y por el tiempo dedicado al desarrollo del Foro Educativo Institucional – Neiva 2026, así como por los aportes, reflexiones y propuestas construidas colectivamente durante la jornada.\n\nAdjuntamos el Informe Ejecutivo del Foro Educativo Institucional – Neiva 2026, que reúne la caracterización institucional, la participación registrada y las respuestas definitivas construidas durante las tres sesiones de trabajo.\n\nTambién puede descargarlo desde este enlace:\n"+linkDescarga+"\n\nAgradecemos especialmente la disposición de la comunidad educativa para participar en este ejercicio de diálogo, reflexión y construcción colectiva orientado al fortalecimiento de la educación en nuestro municipio.\n\nSecretaría de Educación de Neiva\nForo Educativo Institucional – Neiva 2026\n\“Escuela Viva: Voces que construyen territorio\”"; const to=destinatario; const cc=COPIAS_INFORME_FEM.filter(Boolean).join(","); GmailApp.sendEmail(to,subject,body,{htmlBody:logoIEHtmlCorreo+"<p>Apreciados(as) integrantes de la comunidad educativa de la Institución Educativa <strong>"+ieSinPrefijo+"</strong>:</p><p>Reciban un cordial saludo de la Secretaría de Educación de Neiva.</p><p>Agradecemos a la Institución Educativa por su participación y por el tiempo dedicado al desarrollo del <strong>Foro Educativo Institucional – Neiva 2026</strong>, así como por los aportes, reflexiones y propuestas construidas colectivamente durante la jornada.</p><p>Adjuntamos el <strong>Informe Ejecutivo del Foro Educativo Institucional – Neiva 2026</strong>, que reúne la caracterización institucional, la participación registrada y las respuestas definitivas construidas durante las tres sesiones de trabajo.</p><p><a href=\""+linkDescarga+"\">Descargar el informe aquí</a></p><p>Agradecemos especialmente la disposición de la comunidad educativa para participar en este ejercicio de diálogo, reflexión y construcción colectiva orientado al fortalecimiento de la educación en nuestro municipio.</p><p><strong>Secretaría de Educación de Neiva</strong><br>Foro Educativo Institucional – Neiva 2026<br>“Escuela Viva: Voces que construyen territorio”</p>",cc:cc,from:REMITENTE_FEM,name:"Secretaría de Educación de Neiva",attachments:[file.getBlob()]});

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
        htmlBody:logoIEHtmlCorreo+"<p>"+saludoResponsable+"</p><p>Reciba un cordial saludo de la Secretaría de Educación de Neiva.</p><p>Le agradecemos personalmente por haber diligenciado el <strong>Foro Educativo Institucional – Neiva 2026</strong> en representación de la Institución Educativa <strong>"+ieSinPrefijo+"</strong>.</p><p>Adjuntamos el Informe Ejecutivo ya generado. También puede descargarlo desde aquí:</p><p><a href=\""+linkDescarga+"\">Descargar el informe</a></p><p><strong>Secretaría de Educación de Neiva</strong><br>Foro Educativo Institucional – Neiva 2026<br>“Escuela Viva: Voces que construyen territorio”</p>",
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
  const requeridas=["ID_FORO","IE","FECHA","P1_DIALOGO_REFLEXION","P2_PARTICIPACION","P3_IDEAS_PROPUESTAS","P4_SATISFACCION_INSTRUMENTO","NOTA_PROMEDIO","P1_MEJORA","P2_MEJORA","P3_MEJORA","P4_MEJORA","P5_SUGERENCIAS"];
  const last=hoja.getLastColumn();
  const existentes=last?hoja.getRange(1,1,1,last).getValues()[0].map(String):[];
  if(!last){ hoja.getRange(1,1,1,requeridas.length).setValues([requeridas]); }
  else{
    const faltantes=requeridas.filter(h=>existentes.indexOf(h)===-1);
    if(faltantes.length) hoja.getRange(1,last+1,1,faltantes.length).setValues([faltantes]);
  }
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

    const notaPromedio=(puntajes[0]+puntajes[1]+puntajes[2]+puntajes[3])/4;

    const fila=new Array(hoja.getLastColumn()).fill("");
    const valores={
      ID_FORO:String(idForo),
      IE:nombreIE,
      FECHA:new Date(),
      P1_DIALOGO_REFLEXION:puntajes[0],
      P2_PARTICIPACION:puntajes[1],
      P3_IDEAS_PROPUESTAS:puntajes[2],
      P4_SATISFACCION_INSTRUMENTO:puntajes[3],
      NOTA_PROMEDIO:Number(notaPromedio.toFixed(2)),
      P1_MEJORA:String(respuestas.mejoraP1||"").trim(),
      P2_MEJORA:String(respuestas.mejoraP2||"").trim(),
      P3_MEJORA:String(respuestas.mejoraP3||"").trim(),
      P4_MEJORA:String(respuestas.mejoraP4||"").trim(),
      P5_SUGERENCIAS:String(respuestas.p5||"").trim()
    };
    Object.keys(valores).forEach(k=>{ if(m[k]) fila[m[k]-1]=valores[k]; });
    hoja.appendRow(fila);

    try{ actualizarAnalisisFEMIndividual_(idForo); }catch(errorAnalisis){ Logger.log("Análisis FEM (guardarValoracionFEM): "+errorAnalisis.message); }

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
  /*
   * BUG CRÍTICO CORREGIDO: esta función seguía interpretando lo
   * guardado como UN SOLO objeto ({deviceId,tokenSesion,...}), pero
   * desde que se permiten hasta 4 dispositivos simultáneos
   * (reclamarSesionCodigo_) lo que se guarda es SIEMPRE un ARRAY de
   * sesiones. Un array no tiene propiedad .deviceId, así que
   * "a.deviceId !== dispositivoId" era SIEMPRE true — esta función
   * devolvía false para absolutamente cualquier dispositivo, lo que
   * bloqueaba enviarRespuestasSesion() y enviarForoDefinitivo() con
   * "la sesión ya no está activa" incluso para el dispositivo dueño
   * legítimo de su propio cupo. Ahora busca correctamente dentro del
   * array (vía leerSesionesActivas_, que además mantiene
   * compatibilidad con el formato antiguo de un solo objeto).
   */
  const props=PropertiesService.getScriptProperties();
  const clave=obtenerClaveSesionCodigo_("","",idForo);
  const sesiones=leerSesionesActivas_(props, clave);
  return sesiones.some(function(s){
    return s.deviceId===String(dispositivoId||"") && s.tokenSesion===String(tokenSesion||"");
  });
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
    try{ actualizarAnalisisFEMIndividual_(idForo); }catch(errorAnalisis){ Logger.log("Análisis FEM (enviarForoDefinitivo): "+errorAnalisis.message); }
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

/*****************************************************
 * DOCUMENTO DE ANÁLISIS FEM — separado del origen
 *
 * Hoja de cálculo aparte (NO el mismo SPREADSHEET_ID) dedicada a
 * procesar y consultar las respuestas y datos de todas las IE por
 * separado: "Respuestas Totales" (una fila por IE con todo su
 * envío), "Gráficos" (consolidados de participación y valoración) y,
 * de ahí en adelante, una hoja por IE en orden alfabético.
 *
 * Se crea la primera vez que hace falta y su ID queda guardado en
 * ScriptProperties (CLAVE_PROP_SPREADSHEET_ANALISIS) — así queda
 * "vinculado" al proyecto sin depender de un ID fijo escrito en el
 * código, que todavía no se puede conocer de antemano.
 *
 * Se actualiza sola (una IE a la vez, rápido) cada vez que una IE
 * envía definitivamente el foro o guarda su valoración — ver las
 * llamadas a actualizarAnalisisFEMIndividual_() dentro de
 * enviarForoDefinitivo() y guardarValoracionFEM(). Los gráficos y el
 * orden alfabético de las hojas NO se recalculan en cada envío (para
 * no alargar cada envío individual): se refrescan con
 * reconstruirAnalisisFEM() (Pruebas.js), pensada para ejecutarse
 * manualmente cuando se quiera.
 *****************************************************/
const CLAVE_PROP_SPREADSHEET_ANALISIS = "SPREADSHEET_ANALISIS_ID";
const HOJA_ANALISIS_TOTALES = "Respuestas Totales";
const HOJA_ANALISIS_GRAFICOS = "Gráficos";

function obtenerSpreadsheetAnalisisFEM_(){
  const props=PropertiesService.getScriptProperties();
  const idGuardado=props.getProperty(CLAVE_PROP_SPREADSHEET_ANALISIS);
  if(idGuardado){
    try{ return SpreadsheetApp.openById(idGuardado); }
    catch(e){ Logger.log("El documento de análisis guardado ("+idGuardado+") ya no es accesible, se creará uno nuevo: "+e.message); }
  }
  const ss=SpreadsheetApp.create("Análisis FEM 2026 — Foro Educativo Institucional Neiva");
  try{
    const archivo=DriveApp.getFileById(ss.getId());
    DriveApp.getFolderById(DRIVE_CARPETA_FEM_ID).addFile(archivo);
    DriveApp.getRootFolder().removeFile(archivo);
  }catch(e){ Logger.log("No fue posible mover el documento de análisis a la carpeta del FEM: "+e.message); }
  try{ ss.getSheets()[0].setName(HOJA_ANALISIS_TOTALES); }catch(e){}
  props.setProperty(CLAVE_PROP_SPREADSHEET_ANALISIS, ss.getId());
  return ss;
}

const ROLES_PARTICIPACION_ANALISIS_=["Rector","Coordinador","Docentes","TutorPTA","Orientador","Estudiantes","Padres","Administrativos","Egresados","Sector","Otros"];
const COLUMNAS_PARTICIPACION_ANALISIS_=["PART_RECTOR","PART_COORDINADOR","PART_DOCENTES","PART_TUTOR_PTA","PART_ORIENTADOR","PART_ESTUDIANTES","PART_PADRES","PART_ADMINISTRATIVOS","PART_EGRESADOS","PART_SECTOR","PART_OTROS"];
const ETIQUETAS_PARTICIPACION_ANALISIS_=["Rector(a)","Coordinador(a)","Docentes","Tutor PTA PFI/3.0","Orientador(a)","Estudiantes","Padres/madres/acudientes","Personal administrativo","Egresados","Sector productivo","Otros"];

function obtenerCabecerasAnalisisTotales_(){
  return obtenerCabecerasAvancesForo().filter(h=>h!=="DATOS").concat(COLUMNAS_PARTICIPACION_ANALISIS_).concat([
    "TOTAL_PARTICIPANTES","TOTAL_ASISTENTES_QR",
    "VAL_NOTA_PROMEDIO","VAL_P1","VAL_P2","VAL_P3","VAL_P4",
    "VAL_P1_MEJORA","VAL_P2_MEJORA","VAL_P3_MEJORA","VAL_P4_MEJORA","VAL_P5_SUGERENCIAS"
  ]);
}

function asegurarHojaAnalisisTotales_(ss){
  let sh=ss.getSheetByName(HOJA_ANALISIS_TOTALES);
  if(!sh) sh=ss.insertSheet(HOJA_ANALISIS_TOTALES);
  const headers=obtenerCabecerasAnalisisTotales_();
  const last=sh.getLastColumn();
  if(!last){ sh.getRange(1,1,1,headers.length).setValues([headers]); sh.setFrozenRows(1); }
  else{
    const ex=sh.getRange(1,1,1,last).getValues()[0].map(String);
    const faltantes=headers.filter(h=>ex.indexOf(h)===-1);
    if(faltantes.length) sh.getRange(1,last+1,1,faltantes.length).setValues([faltantes]);
  }
  return sh;
}

function obtenerValoracionPorIdForo_(idForo){
  const sh=abrirSpreadsheet_().getSheetByName(HOJA_VALORACION_FEM);
  if(!sh||sh.getLastRow()<2) return null;
  const m=mapaHoja_(sh);
  if(!m.ID_FORO) return null;
  const filas=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getDisplayValues();
  for(let i=0;i<filas.length;i++){
    if(String(filas[i][m.ID_FORO-1]||"").trim()===String(idForo||"").trim()){
      const val=(col)=>m[col]?filas[i][m[col]-1]:"";
      return {
        nota:val("NOTA_PROMEDIO"), p1:val("P1_DIALOGO_REFLEXION"), p2:val("P2_PARTICIPACION"),
        p3:val("P3_IDEAS_PROPUESTAS"), p4:val("P4_SATISFACCION_INSTRUMENTO"),
        p1Mejora:val("P1_MEJORA"), p2Mejora:val("P2_MEJORA"), p3Mejora:val("P3_MEJORA"), p4Mejora:val("P4_MEJORA"),
        p5:val("P5_SUGERENCIAS")
      };
    }
  }
  return null;
}

/*
 * Actualiza (o crea), en el documento de análisis, la fila de UNA
 * sola IE en "Respuestas Totales" y su propia hoja de detalle — se
 * llama automáticamente al enviar el foro definitivo y al guardar la
 * valoración. Nunca debe poder romper esos flujos: quien la llama la
 * envuelve en try/catch, y aquí cada bloque también se protege por
 * separado para que una falla parcial no impida el resto.
 *
 * NO recalcula los gráficos ni el orden alfabético de las hojas (eso
 * alargaría cada envío individual) — para eso está
 * reconstruirAnalisisFEM() en Pruebas.js.
 */
function actualizarAnalisisFEMIndividual_(idForo){
  idForo=String(idForo||"").trim();
  if(!idForo) return;

  const shOrigen=abrirSpreadsheet_().getSheetByName(HOJA_AVANCES);
  if(!shOrigen||shOrigen.getLastRow()<2) return;
  const mOrigen=mapaHoja_(shOrigen);
  const filaOrigen=buscarFilaPorIdForo_(shOrigen,idForo,mOrigen);
  if(filaOrigen<0) return;
  const valoresOrigen=shOrigen.getRange(filaOrigen,1,1,shOrigen.getLastColumn()).getDisplayValues()[0];
  const filaDatos={};
  Object.keys(mOrigen).forEach(k=>{ filaDatos[k]=valoresOrigen[mOrigen[k]-1]; });
  const institucion=String(filaDatos.INSTITUCION||"").trim();
  if(!institucion) return;

  const datosGuardados=obtenerDatosGuardadosPorIdForo_(idForo);
  const campos=(datosGuardados&&datosGuardados.campos)||{};
  const conteoParticipacion=ROLES_PARTICIPACION_ANALISIS_.map(id=>Number(campos["participantes"+id]?.valor||0));
  const totalParticipantes=conteoParticipacion.reduce((a,b)=>a+b,0);
  const totalAsistentesQR=obtenerAsistentesQR_(idForo).length;
  const valoracion=obtenerValoracionPorIdForo_(idForo);

  const ss=obtenerSpreadsheetAnalisisFEM_();

  // --- 1. Fila consolidada en "Respuestas Totales" ---
  try{
    const shTotales=asegurarHojaAnalisisTotales_(ss);
    const mTotales=mapaHoja_(shTotales);
    const fila={};
    Object.keys(mOrigen).forEach(k=>{ if(k!=="DATOS") fila[k]=filaDatos[k]; });
    COLUMNAS_PARTICIPACION_ANALISIS_.forEach((col,i)=>{ fila[col]=conteoParticipacion[i]; });
    fila.TOTAL_PARTICIPANTES=totalParticipantes;
    fila.TOTAL_ASISTENTES_QR=totalAsistentesQR;
    if(valoracion){
      fila.VAL_NOTA_PROMEDIO=valoracion.nota; fila.VAL_P1=valoracion.p1; fila.VAL_P2=valoracion.p2;
      fila.VAL_P3=valoracion.p3; fila.VAL_P4=valoracion.p4;
      fila.VAL_P1_MEJORA=valoracion.p1Mejora; fila.VAL_P2_MEJORA=valoracion.p2Mejora;
      fila.VAL_P3_MEJORA=valoracion.p3Mejora; fila.VAL_P4_MEJORA=valoracion.p4Mejora;
      fila.VAL_P5_SUGERENCIAS=valoracion.p5;
    }
    const out=new Array(shTotales.getLastColumn()).fill("");
    Object.keys(fila).forEach(k=>{ if(mTotales[k]) out[mTotales[k]-1]=fila[k]; });
    let encontrada=-1;
    if(shTotales.getLastRow()>=2 && mTotales.ID_FORO){
      const ids=shTotales.getRange(2,mTotales.ID_FORO,shTotales.getLastRow()-1,1).getDisplayValues();
      for(let i=0;i<ids.length;i++) if(String(ids[i][0]||"").trim()===idForo){ encontrada=i+2; break; }
    }
    if(encontrada>0) shTotales.getRange(encontrada,1,1,out.length).setValues([out]);
    else shTotales.appendRow(out);
  }catch(errorTotales){ Logger.log("Análisis FEM — Respuestas Totales: "+errorTotales.message); }

  // --- 2. Hoja propia de la IE (detalle completo) ---
  try{
    const nombreHoja=nombreHojaIE_(institucion);
    let shIE=ss.getSheetByName(nombreHoja);
    const esHojaNueva=!shIE;
    const headersIE=obtenerCabecerasAvancesForo();
    if(!shIE){ shIE=ss.insertSheet(nombreHoja); shIE.getRange(1,1,1,headersIE.length).setValues([headersIE]); shIE.setFrozenRows(1); }
    if(shIE.getLastColumn()<headersIE.length) shIE.getRange(1,1,1,headersIE.length).setValues([headersIE]);
    const mIE=mapaHoja_(shIE);
    const outIE=new Array(shIE.getLastColumn()).fill("");
    Object.keys(mOrigen).forEach(k=>{ if(mIE[k]) outIE[mIE[k]-1]=filaDatos[k]; });
    let filaIE=-1;
    if(shIE.getLastRow()>=2 && mIE.ID_FORO){
      const ids=shIE.getRange(2,mIE.ID_FORO,shIE.getLastRow()-1,1).getDisplayValues();
      for(let i=0;i<ids.length;i++) if(String(ids[i][0]||"").trim()===idForo){ filaIE=i+2; break; }
    }
    if(filaIE>0) shIE.getRange(filaIE,1,1,outIE.length).setValues([outIE]);
    else shIE.appendRow(outIE);

    // Bloques de detalle debajo de la tabla de caracterización: se
    // reescriben siempre a partir de la misma fila, para no ir
    // acumulando copias en cada actualización.
    const inicioDetalle=Math.max(shIE.getLastRow(),2)+3;
    if(shIE.getMaxRows()>=inicioDetalle) shIE.getRange(inicioDetalle,1,shIE.getMaxRows()-inicioDetalle+1,Math.max(shIE.getMaxColumns(),14)).clearContent();

    shIE.getRange(inicioDetalle,1).setValue("PARTICIPACIÓN — "+institucion);
    const filasParticipacion=ETIQUETAS_PARTICIPACION_ANALISIS_.map((etiqueta,i)=>[etiqueta,conteoParticipacion[i]]);
    shIE.getRange(inicioDetalle+1,1,filasParticipacion.length,2).setValues(filasParticipacion);
    shIE.getRange(inicioDetalle+1+filasParticipacion.length,1,1,2).setValues([["TOTAL",totalParticipantes]]);

    const asistentes=obtenerAsistentesQR_(idForo);
    const inicioAsistencia=inicioDetalle+filasParticipacion.length+4;
    shIE.getRange(inicioAsistencia,1).setValue("ASISTENCIA QR — "+institucion+" ("+asistentes.length+")");
    if(asistentes.length){
      const cabecerasAsistencia=["Nombre","Sexo","Edad","Tipo de asistencia","Cargo","Rol en el Foro","Jornada","Sede","Documento","Correo","Teléfono","Fecha","Hora"];
      shIE.getRange(inicioAsistencia+1,1,1,cabecerasAsistencia.length).setValues([cabecerasAsistencia]);
      const filasAsistencia=asistentes.map(a=>[a.nombre,a.sexo,a.edad,a.tipoAsistencia,a.cargo,a.rolForo,a.jornada,a.sede,a.documento,a.correo,a.telefono,a.fecha,a.hora]);
      shIE.getRange(inicioAsistencia+2,1,filasAsistencia.length,cabecerasAsistencia.length).setValues(filasAsistencia);
    }

    if(valoracion){
      const inicioValoracion=inicioAsistencia+asistentes.length+4;
      shIE.getRange(inicioValoracion,1).setValue("VALORACIÓN — "+institucion);
      shIE.getRange(inicioValoracion+1,1,6,2).setValues([
        ["Nota promedio",valoracion.nota],
        ["P1 diálogo y reflexión",valoracion.p1],
        ["P2 participación",valoracion.p2],
        ["P3 ideas y propuestas",valoracion.p3],
        ["P4 satisfacción del instrumento",valoracion.p4],
        ["P5 sugerencias",valoracion.p5]
      ]);
    }

    // Reordenar alfabéticamente solo cuando aparece una IE nueva —
    // en las actualizaciones normales (misma IE) el orden ya es
    // correcto y no hace falta recorrer todas las hojas.
    if(esHojaNueva) reordenarHojasAnalisisFEM_(ss);
  }catch(errorIE){ Logger.log("Análisis FEM — hoja de "+institucion+": "+errorIE.message); }
}

function reordenarHojasAnalisisFEM_(ss){
  const fijas=[HOJA_ANALISIS_TOTALES,HOJA_ANALISIS_GRAFICOS];
  const hojas=ss.getSheets();
  const deIE=hojas.filter(h=>fijas.indexOf(h.getName())===-1);
  deIE.sort((a,b)=>a.getName().localeCompare(b.getName(),"es"));
  let posicion=1;
  fijas.forEach(nombre=>{ const h=ss.getSheetByName(nombre); if(h){ ss.setActiveSheet(h); ss.moveActiveSheet(posicion); posicion++; } });
  deIE.forEach(h=>{ ss.setActiveSheet(h); ss.moveActiveSheet(posicion); posicion++; });
}

/*
 * Reconstruye "Gráficos" en el documento de análisis a partir de
 * "Respuestas Totales": participación consolidada por estamento (de
 * todas las IE) y nota promedio de valoración por IE. Se llama desde
 * reconstruirAnalisisFEM() (Pruebas.js), no en cada envío individual.
 */
function actualizarGraficosAnalisisFEM_(ss){
  let sh=ss.getSheetByName(HOJA_ANALISIS_GRAFICOS);
  if(!sh) sh=ss.insertSheet(HOJA_ANALISIS_GRAFICOS);
  sh.getCharts().forEach(c=>sh.removeChart(c));
  sh.clear();

  const shTotales=ss.getSheetByName(HOJA_ANALISIS_TOTALES);
  if(!shTotales||shTotales.getLastRow()<2){ sh.getRange(1,1).setValue("Todavía no hay respuestas registradas."); return; }
  const m=mapaHoja_(shTotales);
  const filas=shTotales.getRange(2,1,shTotales.getLastRow()-1,shTotales.getLastColumn()).getValues();

  const totalesRoles=COLUMNAS_PARTICIPACION_ANALISIS_.map(col=>{ if(!m[col]) return 0; return filas.reduce((s,f)=>s+Number(f[m[col]-1]||0),0); });
  sh.getRange(1,1,1,2).setValues([["Estamento","Participantes"]]);
  sh.getRange(2,1,ETIQUETAS_PARTICIPACION_ANALISIS_.length,2).setValues(ETIQUETAS_PARTICIPACION_ANALISIS_.map((e,i)=>[e,totalesRoles[i]]));
  const rangoParticipacion=sh.getRange(1,1,ETIQUETAS_PARTICIPACION_ANALISIS_.length+1,2);
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.PIE).addRange(rangoParticipacion).setOption("title","Participación consolidada por estamento — todas las IE").setPosition(1,4,0,0).build());
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(rangoParticipacion).setOption("title","Total de participantes por estamento — todas las IE").setPosition(20,4,0,0).build());

  if(m.INSTITUCION && m.VAL_NOTA_PROMEDIO){
    const inicioVal=ETIQUETAS_PARTICIPACION_ANALISIS_.length+4;
    const datosValoracion=filas.map(f=>[String(f[m.INSTITUCION-1]||""),Number(f[m.VAL_NOTA_PROMEDIO-1]||0)]).filter(f=>f[1]>0);
    if(datosValoracion.length){
      sh.getRange(inicioVal,1,1,2).setValues([["Institución","Nota promedio"]]);
      sh.getRange(inicioVal+1,1,datosValoracion.length,2).setValues(datosValoracion);
      const rangoValoracion=sh.getRange(inicioVal,1,datosValoracion.length+1,2);
      sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(rangoValoracion).setOption("title","Nota promedio de valoración por IE").setOption("vAxis.viewWindow.max",5).setPosition(inicioVal+datosValoracion.length+3,1,0,0).build());
    }
  }
}

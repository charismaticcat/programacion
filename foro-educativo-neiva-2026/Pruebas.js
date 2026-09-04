/*****************************************************
 * FORO EDUCATIVO INSTITUCIONAL
 * Neiva 2026
 *
 * PRUEBAS Y HERRAMIENTAS DE DESARROLLO
 *
 * Funciones que NO forman parte del flujo real de la
 * aplicación (doGet, validarAccesoIE, guardarAvanceForo,
 * enviarForoDefinitivo, etc.). Se ejecutan manualmente
 * desde el editor de Apps Script para probar el sistema
 * con la institución de prueba "IE PRUEBA 1234" o para
 * revisar datos puntuales.
 *
 * Movidas aquí desde Código.js el 2026-08-25 para separar
 * el código de producción del código de desarrollo.
 *****************************************************/


/*****************************************************
 * LIMPIAR INFORMES ROTOS (.docx) DE LA CARPETA FEM
 *
 * generarInformeFEM() usaba antes una plantilla de Word
 * (.docx) copiada con makeCopy(); cada intento fallido dejó
 * una copia vacía de ese mismo formato en la carpeta de la
 * IE correspondiente, sin que el informe llegara a generarse.
 *
 * Esta función busca, dentro de cada carpeta de institución
 * en DRIVE_CARPETA_FEM_ID, archivos que empiecen por "Informe
 * Ejecutivo" Y cuyo tipo sea Word (.docx) — nunca un informe
 * real, que ahora siempre es un Google Doc nativo — y los
 * envía a la papelera (no los borra permanentemente: quedan
 * recuperables desde la papelera de Drive por 30 días).
 *
 * Ejecutar manualmente una sola vez desde el editor.
 *****************************************************/
function limpiarInformesRotos() {

  const MIME_WORD =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  const raiz = DriveApp.getFolderById(DRIVE_CARPETA_FEM_ID);
  const detalles = [];
  let eliminados = 0;

  const carpetas = raiz.getFolders();

  while (carpetas.hasNext()) {

    const carpeta = carpetas.next();
    const archivos = carpeta.getFiles();

    while (archivos.hasNext()) {

      const archivo = archivos.next();

      if (
        archivo.getName().indexOf("Informe Ejecutivo") === 0 &&
        archivo.getMimeType() === MIME_WORD
      ) {

        detalles.push(carpeta.getName() + " / " + archivo.getName());
        archivo.setTrashed(true);
        eliminados++;

      }

    }

  }

  Logger.log(
    "Informes .docx rotos enviados a la papelera: " + eliminados
  );

  Logger.log(detalles.join("\n"));

  return {
    ok: true,
    eliminados: eliminados,
    detalles: detalles
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

  const ss = abrirSpreadsheet_();

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
    URL_WEBAPP_PRODUCCION +
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
 * PRUEBA DE VALIDACIÓN DEL ACCESO 1234
 *****************************************************/

function probarValidacion1234() {

  const ss =
    abrirSpreadsheet_();

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
    abrirSpreadsheet_();

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
    abrirSpreadsheet_();

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
 * PRUEBA DE VALIDACIÓN - CÓDIGO INCORRECTO
 *
 * Esta función NO modifica AccesosIE.
 * Solo comprueba que un código incorrecto
 * sea rechazado por validarAccesoIE().
 *****************************************************/

function probarCodigoIncorrecto1234() {

  const ss =
    abrirSpreadsheet_();

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


/*****************************************************
 * REINICIAR PRUEBA "IE PRUEBA 1234"
 *
 * Borra todo rastro de un envío de prueba anterior para
 * poder volver a probar el formulario completo desde cero
 * con el mismo código de acceso (1234):
 *
 *   - Fila en AvancesForo.
 *   - Fila(s) en Participacion.
 *   - Fila(s) en AsistenciaQR.
 *   - Fila(s) en "Valoración FEMI2026".
 *   - La pestaña propia "IE PRUEBA 1234" (si existe).
 *   - ESTADO, S1/S2/S3_ENVIADA, fechas de envío, TOKEN_SESION,
 *     DISPOSITIVO_ID, ID_INFORME e ID_PDF_INFORME en la fila
 *     de AccesosIE (se dejan en blanco / "DISPONIBLE").
 *   - El candado de sesión activa (PropertiesService).
 *
 * NO borra los archivos ya generados en Drive (fotos,
 * informes) — eso se hace a mano desde la carpeta si hace
 * falta. NO toca ninguna IE oficial.
 *
 * Ejecutar manualmente desde el editor de Apps Script.
 *
 * A partir de este mismo cambio, para reiniciar una prueba
 * YA NO hace falta ejecutar esta función cada vez: basta con
 * borrar a mano la fila de la IE en AvancesForo — la próxima
 * vez que se ingrese con el código, validarAccesoIE detecta
 * que ya no hay un envío completo ahí y reabre el acceso
 * automáticamente, sin tocar nada más. Esta función sigue
 * sirviendo para limpiar TODO de una sola vez (incluyendo
 * Participacion, AsistenciaQR y Valoración, que no se borran
 * solas).
 *****************************************************/
function reiniciarPrueba1234(){

  const idForo = "PRUEBA-1234";
  const ie = "IE PRUEBA 1234";
  const resumen = [];

  const ss = abrirSpreadsheet_();

  function borrarFilasPorIdForo_(nombreHoja){
    const hoja = ss.getSheetByName(nombreHoja);
    if(!hoja){ resumen.push(nombreHoja + ": la hoja no existe."); return; }
    const mapa = mapaHoja_(hoja);
    if(!mapa["ID_FORO"]){ resumen.push(nombreHoja + ": no tiene columna ID_FORO."); return; }
    const ultimaFila = hoja.getLastRow();
    if(ultimaFila < 2){ resumen.push(nombreHoja + ": sin filas."); return; }
    const valores = hoja.getRange(2, mapa["ID_FORO"], ultimaFila - 1, 1).getDisplayValues();
    let borradas = 0;
    // De abajo hacia arriba para no desordenar los índices al borrar.
    for(let i = valores.length - 1; i >= 0; i--){
      if(String(valores[i][0] || "").trim() === idForo){
        hoja.deleteRow(i + 2);
        borradas++;
      }
    }
    resumen.push(nombreHoja + ": " + borradas + " fila(s) borrada(s).");
  }

  borrarFilasPorIdForo_(HOJA_AVANCES);
  borrarFilasPorIdForo_(HOJA_PARTICIPACION);
  borrarFilasPorIdForo_(HOJA_ASISTENCIA_QR);
  borrarFilasPorIdForo_(HOJA_VALORACION_FEM);

  const nombreHojaIe = nombreHojaIE_(ie);
  const hojaIe = ss.getSheetByName(nombreHojaIe);
  if(hojaIe){
    ss.deleteSheet(hojaIe);
    resumen.push("Pestaña \"" + nombreHojaIe + "\": eliminada.");
  }else{
    resumen.push("Pestaña \"" + nombreHojaIe + "\": no existía.");
  }

  const hojaAccesos = ss.getSheetByName(HOJA_ACCESOS);
  if(hojaAccesos){
    const mapaAccesos = mapaHoja_(hojaAccesos);
    const filaAccesos = buscarFilaPorIdForo_(hojaAccesos, idForo, mapaAccesos);
    if(filaAccesos > 0){
      if(mapaAccesos.ESTADO) hojaAccesos.getRange(filaAccesos, mapaAccesos.ESTADO).setValue("DISPONIBLE");
      [
        "S1_ENVIADA","S2_ENVIADA","S3_ENVIADA",
        "FECHA_ENVIO","FECHA_ENVIO_S1","FECHA_ENVIO_S2","FECHA_ENVIO_S3","FECHA_ENVIO_DEFINITIVO",
        "TOKEN_SESION","DISPOSITIVO_ID","ID_INFORME","ID_PDF_INFORME"
      ].forEach(function(col){
        if(mapaAccesos[col]) hojaAccesos.getRange(filaAccesos, mapaAccesos[col]).setValue("");
      });
      resumen.push("AccesosIE: fila " + filaAccesos + " reiniciada a DISPONIBLE.");
    }else{
      resumen.push("AccesosIE: no se encontró la fila de " + ie + ".");
    }
  }

  try{
    PropertiesService.getScriptProperties().deleteProperty(obtenerClaveSesionCodigo_("", "", idForo));
    resumen.push("Candado de sesión: liberado.");
  }catch(error){
    resumen.push("Candado de sesión: " + error.message);
  }

  Logger.log(resumen.join("\n"));
  return { ok: true, resumen: resumen };
}


/*****************************************************
 * REINICIAR TODOS LOS REGISTROS DEL FEM 2026
 *
 * Generaliza reiniciarPrueba1234() a TODAS las filas de AccesosIE
 * (no solo la IE de pruebas): borra todo rastro de progreso o
 * envíos anteriores para poder recorrer el formulario completo
 * desde cero, con los MISMOS códigos y enlaces de acceso ya
 * generados y enviados por correo — no hace falta reenviar nada.
 *
 * Por cada IE de AccesosIE:
 *   - Borra sus fila(s) en AvancesForo, Participacion, AsistenciaQR
 *     y "Valoración FEMI2026".
 *   - Elimina su pestaña propia (si existe).
 *   - Dentro de AccesosIE, deja ESTADO en "DISPONIBLE" y limpia
 *     S1/S2/S3_ENVIADA, todas las fechas de envío, TOKEN_SESION,
 *     DISPOSITIVO_ID, ID_INFORME e ID_PDF_INFORME.
 * Además libera TODOS los candados de sesión activa
 * (PropertiesService) para que ningún dispositivo quede "con la
 * sesión tomada" de una prueba anterior.
 *
 * NO TOCA — quedan exactamente igual que antes de ejecutarla —:
 *   - IE, DANE, CODIGO_ACCESO, TOKEN, URL_ACCESO, ID_FORO,
 *     EMAIL_IE, EMAIL_RESPONSABLE, TIPO ni LOGO_ID: los códigos y
 *     enlaces que ya se enviaron por correo siguen siendo válidos.
 *   - Los archivos ya generados en Drive (fotos e informes) — eso
 *     se borra a mano desde la carpeta de cada IE si hace falta.
 *
 * IMPORTANTE — esto NO alcanza el localStorage del navegador de
 * cada equipo (es un respaldo aparte, guardado fuera de Google, en
 * el propio dispositivo de quien probó). Si al volver a probar con
 * el MISMO navegador siguen apareciendo respuestas viejas después
 * de ejecutar esto, hay que borrar también los datos del sitio en
 * ese navegador (Ajustes del sitio → Borrar datos, o simplemente
 * probar en una ventana de incógnito) — ninguna función del
 * servidor puede alcanzar esos datos.
 *
 * Ejecutar manualmente desde el editor de Apps Script.
 *****************************************************/
function reiniciarTodosLosRegistrosFEM(){

  const ss = abrirSpreadsheet_();
  const resumen = [];

  const hojaAccesos = ss.getSheetByName(HOJA_ACCESOS);
  if(!hojaAccesos){
    Logger.log("No existe la hoja " + HOJA_ACCESOS + ".");
    return { ok: false, mensaje: "No existe " + HOJA_ACCESOS + "." };
  }

  const mapaAccesos = mapaHoja_(hojaAccesos);
  const ultimaFila = hojaAccesos.getLastRow();
  if(ultimaFila < 2){
    Logger.log("AccesosIE no tiene filas.");
    return { ok: true, resumen: ["AccesosIE no tiene filas."] };
  }
  if(!mapaAccesos.ID_FORO || !mapaAccesos.IE){
    Logger.log("Falta la columna ID_FORO o IE en AccesosIE.");
    return { ok: false, mensaje: "Falta la columna ID_FORO o IE en AccesosIE." };
  }

  const totalFilas = ultimaFila - 1;
  const valores = hojaAccesos.getRange(2, 1, totalFilas, hojaAccesos.getLastColumn()).getDisplayValues();
  const idsForo = {};
  const nombresIE = [];
  valores.forEach(function(fila){
    const id = String(fila[mapaAccesos.ID_FORO - 1] || "").trim();
    const ie = String(fila[mapaAccesos.IE - 1] || "").trim();
    if(id) idsForo[id] = true;
    if(ie) nombresIE.push(ie);
  });

  function borrarFilasPorIdForoEnTodas_(nombreHoja){
    const hoja = ss.getSheetByName(nombreHoja);
    if(!hoja){ resumen.push(nombreHoja + ": la hoja no existe."); return; }
    const mapa = mapaHoja_(hoja);
    if(!mapa["ID_FORO"]){ resumen.push(nombreHoja + ": no tiene columna ID_FORO."); return; }
    const ultima = hoja.getLastRow();
    if(ultima < 2){ resumen.push(nombreHoja + ": sin filas."); return; }
    const idsHoja = hoja.getRange(2, mapa["ID_FORO"], ultima - 1, 1).getDisplayValues();
    let borradas = 0;
    // De abajo hacia arriba para no desordenar los índices al borrar.
    for(let i = idsHoja.length - 1; i >= 0; i--){
      const id = String(idsHoja[i][0] || "").trim();
      if(id && idsForo[id]){ hoja.deleteRow(i + 2); borradas++; }
    }
    resumen.push(nombreHoja + ": " + borradas + " fila(s) borrada(s).");
  }

  borrarFilasPorIdForoEnTodas_(HOJA_AVANCES);
  borrarFilasPorIdForoEnTodas_(HOJA_PARTICIPACION);
  borrarFilasPorIdForoEnTodas_(HOJA_ASISTENCIA_QR);
  borrarFilasPorIdForoEnTodas_(HOJA_VALORACION_FEM);

  let pestañasEliminadas = 0;
  nombresIE.forEach(function(ie){
    const nombreHoja = nombreHojaIE_(ie);
    const hoja = ss.getSheetByName(nombreHoja);
    if(hoja){ ss.deleteSheet(hoja); pestañasEliminadas++; }
  });
  resumen.push("Pestañas propias por IE eliminadas: " + pestañasEliminadas + " de " + nombresIE.length + " IE.");

  /*
   * Reiniciar el estado de TODAS las filas de AccesosIE de una
   * sola vez por columna (un solo setValue por columna, no uno por
   * celda) — mucho más rápido y evita agotar la cuota de llamadas
   * con hojas grandes.
   */
  if(mapaAccesos.ESTADO){
    hojaAccesos.getRange(2, mapaAccesos.ESTADO, totalFilas, 1).setValue("DISPONIBLE");
  }
  [
    "S1_ENVIADA", "S2_ENVIADA", "S3_ENVIADA",
    "FECHA_PRIMER_ACCESO", "ULTIMA_ACTIVIDAD", "FECHA_ENVIO",
    "FECHA_ENVIO_S1", "FECHA_ENVIO_S2", "FECHA_ENVIO_S3", "FECHA_ENVIO_DEFINITIVO",
    "TOKEN_SESION", "DISPOSITIVO_ID", "ID_INFORME", "ID_PDF_INFORME"
  ].forEach(function(col){
    if(mapaAccesos[col]) hojaAccesos.getRange(2, mapaAccesos[col], totalFilas, 1).setValue("");
  });
  resumen.push(
    "AccesosIE: " + totalFilas + " fila(s) reiniciada(s) a DISPONIBLE " +
    "(IE, DANE, CODIGO_ACCESO, TOKEN, URL_ACCESO, ID_FORO, EMAIL_IE, " +
    "EMAIL_RESPONSABLE, TIPO y LOGO_ID quedan intactos)."
  );

  try{
    const props = PropertiesService.getScriptProperties();
    const todas = props.getProperties();
    let liberados = 0;
    Object.keys(todas).forEach(function(clave){
      if(clave.indexOf("FEM_SESION_FORO_") === 0){
        props.deleteProperty(clave);
        liberados++;
      }
    });
    resumen.push("Candados de sesión activa liberados: " + liberados + ".");
  }catch(error){
    resumen.push("Candados de sesión: " + error.message);
  }

  resumen.push("");
  resumen.push(
    "⚠ Esto NO borra archivos en Drive (fotos e informes ya generados) " +
    "ni el localStorage del navegador de cada equipo. Si al volver a " +
    "probar con el MISMO navegador siguen apareciendo respuestas " +
    "viejas, hay que borrar los datos del sitio en ese navegador o " +
    "usar una ventana de incógnito."
  );

  Logger.log(resumen.join("\n"));
  return { ok: true, resumen: resumen };
}


/*****************************************************
 * PROBAR REMITENTE calidadeducacion@alcaldianeiva.gov.co
 *
 * Verifica, con un envío real y mínimo, que la cuenta que
 * ejecuta el script (la que aparece en Session.getEffectiveUser())
 * ya puede enviar como REMITENTE_FEM — es decir, que el alias
 * "Enviar correo como" quedó bien configurado y verificado en
 * Gmail. El correo de prueba se manda a la misma cuenta que
 * ejecuta el script, para no molestar a nadie más.
 *
 * Ejecutar manualmente desde el editor de Apps Script después
 * de configurar el alias.
 *****************************************************/
function probarRemitenteFEM(){

  const cuenta = Session.getEffectiveUser().getEmail();
  const aliases = GmailApp.getAliases();

  Logger.log("Cuenta que ejecuta el script: " + cuenta);
  Logger.log("Alias configurados: " + (aliases.length ? aliases.join(", ") : "(ninguno)"));
  Logger.log("REMITENTE_FEM configurado en el código: " + REMITENTE_FEM);

  const autorizado =
    cuenta.toLowerCase() === REMITENTE_FEM.toLowerCase() ||
    aliases.map(a => a.toLowerCase()).indexOf(REMITENTE_FEM.toLowerCase()) !== -1;

  if(!autorizado){
    Logger.log("❌ TODAVÍA NO: " + REMITENTE_FEM + " no aparece como alias autorizado para " + cuenta + ".");
    return { ok:false, mensaje:"El alias aún no está autorizado.", cuenta:cuenta, aliases:aliases };
  }

  try{
    GmailApp.sendEmail(
      cuenta,
      "Prueba de remitente FEM 2026 — " + REMITENTE_FEM,
      "Este es un correo de prueba para confirmar que ya se puede enviar como " + REMITENTE_FEM + ".",
      { from: REMITENTE_FEM, name: "Secretaría de Educación de Neiva" }
    );
    Logger.log("✅ Envío correcto: se pudo enviar como " + REMITENTE_FEM + ". Revise la bandeja de " + cuenta + ".");
    return { ok:true, mensaje:"Envío correcto." };
  }catch(error){
    Logger.log("❌ Error al enviar como " + REMITENTE_FEM + ": " + error.message);
    return { ok:false, mensaje:error.message };
  }

}


/*****************************************************
 * VINCULAR LOGOS DE CADA IE
 *
 * Recorre la carpeta de Drive con los logos institucionales
 * (una por IE, subidos con un nombre parecido al de la
 * institución) y los relaciona con cada fila de AccesosIE por
 * nombre, guardando el ID del archivo en la columna LOGO_ID.
 *
 * El emparejamiento es por nombre normalizado (sin tildes, sin
 * mayúsculas sostenidas, sin el prefijo "Institución Educativa"/
 * "IE"), así que "CHAPINERO.png" coincide con la IE "CHAPINERO"
 * o "INSTITUCIÓN EDUCATIVA CHAPINERO" indistintamente.
 *
 * NO asigna nada cuando hay ambigüedad (el mismo nombre coincide
 * con varios archivos, o el nombre de un archivo coincide con
 * varias IE): esos casos quedan listados en el registro de
 * ejecución para revisarlos y asignarlos a mano en la columna
 * LOGO_ID de AccesosIE.
 *
 * Ejecutar manualmente desde el editor de Apps Script.
 *****************************************************/
function vincularLogosIE(){

  const CARPETA_LOGOS_ID = "1QVfDyYjhjX5H60U7SyeLtikodQbhGu1B";

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  const ultimaFila = hoja.getLastRow();

  if(ultimaFila < 2){
    Logger.log("AccesosIE no tiene filas.");
    return;
  }

  if(!mapa.IE || !mapa.LOGO_ID){
    Logger.log("Falta la columna IE o LOGO_ID en AccesosIE.");
    return;
  }

  // Indexar los archivos de la carpeta por nombre normalizado
  // (sin extensión). Puede haber varios archivos con el mismo
  // nombre normalizado (duplicados subidos más de una vez).
  const carpeta = DriveApp.getFolderById(CARPETA_LOGOS_ID);
  const archivos = carpeta.getFiles();
  const indice = {};

  while(archivos.hasNext()){
    const archivo = archivos.next();
    const nombreSinExtension = archivo.getName().replace(/\.[^.]+$/, "");
    // Se quita también un posible prefijo "IE"/"Institución Educativa"
    // del NOMBRE DEL ARCHIVO, para que "MARIA CRISTINA ARANGO.png" e
    // "IE MARIA CRISTINA ARANGO.png" caigan bajo la misma clave — así
    // se detectan como duplicados en vez de que uno quede "invisible"
    // por tener una clave distinta a la de la IE.
    const clave = normalizarNombreIE_(nombreIESinPrefijoInstitucional_(nombreSinExtension));
    if(!clave) continue;
    if(!indice[clave]) indice[clave] = [];
    indice[clave].push({ id: archivo.getId(), nombre: archivo.getName() });
  }

  const clavesDisponibles = Object.keys(indice);

  // Correcciones manuales para archivos reales de la carpeta cuyo
  // nombre tiene una errata o una redacción distinta a la del nombre
  // oficial de la IE en AccesosIE (confirmado archivo por archivo:
  // "ATASIO" por "ATANASIO", "LIZACANO" por "LIZCANO", etc.). Sin
  // este mapa esas 7 IE quedaban en SIN COINCIDENCIA aunque su logo
  // sí estaba subido a la carpeta.
  const ALIAS_ARCHIVO_POR_IE = {
    "ATANASIO GIRARDOT": "ATASIO GIRARDOT",
    "JAIRO MORERA LIZCANO": "JAIRO MORERA LIZACANO",
    "JAIRO MOSQUERA MORENO": "JARIO MOSQUERA MORENO-GUACIRCO",
    "LICEO DE SANTA LIBRADA": "LICEO SANTALIBRADA",
    "MARIA AUXILIADORA FORTALECILLAS": "MARIA AUXILIADORA DE FOTALECILLAS",
    "SANTA LIBRADA": "NACIONAL SANTALIBRADA",
    "INSTITUTO TECNICO IPC ANDRES ROSA": "TECNICO IPC ANDRES ROSA"
  };
  const claveAliasPorIE = {};
  Object.keys(ALIAS_ARCHIVO_POR_IE).forEach(function(nombreIE){
    claveAliasPorIE[normalizarNombreIE_(nombreIE)] = normalizarNombreIE_(ALIAS_ARCHIVO_POR_IE[nombreIE]);
  });

  const datos = hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getValues();
  const asignados = [];
  const sinCoincidencia = [];
  const ambiguos = [];

  for(let i=0;i<datos.length;i++){

    const ie = String(datos[i][mapa.IE-1] || "").trim();
    if(!ie) continue;

    // Ya tiene logo asignado: no se vuelve a tocar (para no
    // pisar una corrección manual ya hecha en la hoja).
    const logoActual = mapa.LOGO_ID ? String(datos[i][mapa.LOGO_ID-1] || "").trim() : "";
    if(logoActual) continue;

    const claveDirecta = normalizarNombreIE_(ie);
    const claveSinPrefijo = normalizarNombreIE_(nombreIESinPrefijoInstitucional_(ie));

    let candidatos = indice[claveDirecta] || indice[claveSinPrefijo];

    if(!candidatos){
      // Corrección manual conocida (errata o redacción distinta en
      // el nombre del archivo real subido a la carpeta).
      const claveAlias = claveAliasPorIE[claveDirecta] || claveAliasPorIE[claveSinPrefijo];
      if(claveAlias) candidatos = indice[claveAlias];
    }

    if(!candidatos){
      // Coincidencia parcial: el nombre del archivo está contenido
      // en el nombre de la IE, o al revés (para variantes como
      // "SAN LUIS BELTRAN" vs "SAN LUIS BELTRAN SEDE PRINCIPAL").
      const posibles = clavesDisponibles.filter(function(clave){
        return clave.length > 3 && (claveSinPrefijo.indexOf(clave) !== -1 || clave.indexOf(claveSinPrefijo) !== -1);
      });
      if(posibles.length === 1){
        candidatos = indice[posibles[0]];
      }else if(posibles.length > 1){
        ambiguos.push(ie + " -> varias coincidencias parciales: " + posibles.map(function(k){ return indice[k].map(function(a){return a.nombre;}).join("/"); }).join(", "));
        continue;
      }
    }

    if(!candidatos){
      sinCoincidencia.push(ie);
      continue;
    }

    if(candidatos.length > 1){
      ambiguos.push(ie + " -> " + candidatos.length + " archivos con el mismo nombre: " + candidatos.map(function(a){return a.nombre;}).join(", ") + " (se usó el primero)");
    }

    const elegido = candidatos[0];
    hoja.getRange(i+2, mapa.LOGO_ID).setValue(elegido.id);
    try{ DriveApp.getFileById(elegido.id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(errorCompartir){}
    asignados.push(ie + " -> " + elegido.nombre);

  }

  Logger.log("========================================");
  Logger.log("ASIGNADOS (" + asignados.length + "):");
  Logger.log(asignados.join("\n") || "(ninguno)");
  Logger.log("========================================");
  Logger.log("SIN COINCIDENCIA (" + sinCoincidencia.length + ") — asignar a mano en LOGO_ID:");
  Logger.log(sinCoincidencia.join("\n") || "(ninguna)");
  Logger.log("========================================");
  Logger.log("AMBIGUOS, REVISAR (" + ambiguos.length + "):");
  Logger.log(ambiguos.join("\n") || "(ninguno)");
  Logger.log("========================================");

  return { asignados: asignados.length, sinCoincidencia: sinCoincidencia.length, ambiguos: ambiguos.length };

}


/*****************************************************
 * BORRAR INFORMES, FOTOS Y REGISTRO DE PARTICIPACIÓN
 *
 * Envía a la papelera de Drive (recuperable 30 días, NO es un
 * borrado permanente) TODOS los archivos dentro de cada carpeta de
 * institución en DRIVE_CARPETA_FEM_ID: fotografías de evidencia,
 * documentos de informe (Google Doc) y sus PDF — de cualquier IE,
 * oficial o de prueba.
 *
 * También borra TODAS las filas de datos de la hoja Participacion
 * (deja el encabezado).
 *
 * NO toca AccesosIE, AvancesForo, AsistenciaQR ni "Valoración
 * FEMI2026" — para eso está reiniciarTodosLosRegistrosFEM().
 * NO borra las carpetas por IE, solo su contenido.
 *
 * Ejecutar manualmente desde el editor de Apps Script.
 *****************************************************/
function borrarArchivosDriveYParticipacionFEM(){

  const raiz = DriveApp.getFolderById(DRIVE_CARPETA_FEM_ID);
  const detalles = [];
  let archivosBorrados = 0;
  let carpetasRevisadas = 0;

  const carpetas = raiz.getFolders();
  while(carpetas.hasNext()){
    const carpeta = carpetas.next();
    carpetasRevisadas++;
    const archivos = carpeta.getFiles();
    while(archivos.hasNext()){
      const archivo = archivos.next();
      detalles.push(carpeta.getName() + " / " + archivo.getName());
      archivo.setTrashed(true);
      archivosBorrados++;
    }
  }

  const ss = abrirSpreadsheet_();
  const hojaParticipacion = ss.getSheetByName(HOJA_PARTICIPACION);
  let filasParticipacionBorradas = 0;
  if(hojaParticipacion && hojaParticipacion.getLastRow() >= 2){
    filasParticipacionBorradas = hojaParticipacion.getLastRow() - 1;
    hojaParticipacion.deleteRows(2, filasParticipacionBorradas);
  }

  const resumen = [
    "Carpetas de IE revisadas: " + carpetasRevisadas,
    "Archivos enviados a la papelera (fotos + informes + PDF): " + archivosBorrados,
    "Filas borradas en " + HOJA_PARTICIPACION + ": " + filasParticipacionBorradas,
    "",
    "Los archivos quedan en la papelera de Drive por 30 días, no se borraron de forma permanente."
  ];

  Logger.log(resumen.join("\n"));
  Logger.log("Detalle de archivos borrados:\n" + detalles.join("\n"));

  return { ok:true, archivosBorrados: archivosBorrados, filasParticipacionBorradas: filasParticipacionBorradas, resumen: resumen };

}


/*****************************************************
 * CREAR IE DE PRUEBA ADICIONALES
 *
 * Crea, directamente en AccesosIE (nunca en "Oficiales"), una fila
 * por cada IE de la lista IES_PRUEBA_ADICIONALES — con TIPO =
 * "PRUEBA", igual que "IE PRUEBA 1234", para que validarAccesoIE las
 * acepte sin exigir que existan en la hoja oficial. Cada fila recibe
 * su propio CODIGO_ACCESO, TOKEN, URL_ACCESO e ID_FORO generados al
 * azar (misma mecánica que las IE reales), EMAIL_IE y
 * EMAIL_RESPONSABLE con el correo indicado, y el mismo LOGO_ID que
 * ya tenga vinculado "IE PRUEBA 1234" (si tiene).
 *
 * También precarga su caracterización en AvancesForo con un texto
 * de muestra en todos los campos salvo "Rector(a)", y respuestas de
 * relleno en las tres sesiones — así generarInformeFEM() puede
 * producir un informe real de inmediato, sin depender de que nadie
 * complete el formulario a mano.
 *
 * Requiere el cambio en validarAccesoIE que deja de exigir que la
 * fila TIPO = PRUEBA sea exactamente "IE PRUEBA 1234" — con ese
 * cambio, cualquier fila marcada así funciona igual.
 *
 * Se puede ejecutar varias veces: una IE que ya exista (por nombre,
 * sin importar mayúsculas) se salta, nunca se duplica.
 *
 * Ejecutar manualmente desde el editor de Apps Script.
 *****************************************************/
const IES_PRUEBA_ADICIONALES = [
  { ie: "IE Prueba Ronald",   email: "ronald.polania@alcaldianeiva.gov.co" },
  { ie: "IE Prueba Edna",     email: "articulacionsem@alcaldianeiva.gov.co" },
  { ie: "IE Prueba Adriana",  email: "adriana.cedeno@alcaldianeiva.gov.co" },
  { ie: "IE Prueba Carlos Q", email: "ingenierocarlosq@hotmail.com" },
  { ie: "IE Prueba Nelson",   email: "nelson.herrera@alcaldianeiva.gov.co" },
  { ie: "IE Prueba Rosarito", email: "rosario.vidal@alcaldianeiva.gov.co" },
  { ie: "IE Prueba Angélica", email: "angelica.rojas@alcaldianeiva.gov.co" },
  { ie: "IE Prueba Ana",      email: "ana.torres@alcaldianeiva.gov.co" },
  { ie: "IE Prueba Rosa",     email: "rosa.gonzalez@alcaldianeiva.gov.co" }
];

const TEXTO_MUESTRA_CARACTERIZACION_PRUEBA =
  "Esto es una muestra de la información que se mostrará a las IE automáticamente.";
const TEXTO_MUESTRA_RECTOR_PRUEBA =
  "Aquí aparecerá el nombre del rector y será también editable.";
const TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA =
  "Esta es una respuesta de muestra generada para la IE de prueba, con el fin de " +
  "validar el funcionamiento completo del formulario: el guardado de la " +
  "información, el envío definitivo, la generación del informe ejecutivo y el " +
  "envío de los correos electrónicos correspondientes. El contenido no " +
  "corresponde a una reflexión real de ninguna institución educativa — es " +
  "únicamente texto de relleno, con la extensión suficiente para completar el " +
  "recorrido de las tres sesiones sin quedar bloqueado por el conteo mínimo de " +
  "palabras, y así comprobar que el resto del proceso (guardado en la nube, " +
  "envío definitivo, generación del documento y notificación por correo) " +
  "funciona correctamente de principio a fin, antes de que las instituciones " +
  "educativas oficiales completen el Foro Educativo Institucional Neiva 2026 " +
  "con sus propias respuestas.";
const TEXTO_MUESTRA_ACCION_PRUEBA =
  "Acción de muestra para validar el formulario, sus límites de palabras y la generación del informe.";

function crearIEsPruebaAdicionales(){

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  const ultimaFila = hoja.getLastRow();

  // Mapa nombre (en minúscula) -> número de fila, para saltar las
  // que ya existen SIN dejar de poder repararlas (ver más abajo:
  // si a una que ya existía le falta LINK_ACCESO, se completa).
  const filasExistentes = {};
  if(ultimaFila >= 2 && mapa.IE){
    const ies = hoja.getRange(2, mapa.IE, ultimaFila - 1, 1).getDisplayValues();
    ies.forEach(function(fila, i){
      const nombre = String(fila[0]||"").trim().toLowerCase();
      if(nombre) filasExistentes[nombre] = i + 2;
    });
  }

  const urlBase = URL_WEBAPP_PRODUCCION;
  if(!urlBase){
    Logger.log("No fue posible obtener la URL del Web App (URL_WEBAPP_PRODUCCION).");
    return { ok:false, mensaje:"Falta URL_WEBAPP_PRODUCCION." };
  }

  // Reutiliza el logo ya vinculado a "IE PRUEBA 1234" (si lo tiene),
  // para que las IE de prueba no se vean sin logo en ningún lado.
  const logoIdPruebaOriginal = obtenerLogoIdPorNombreIE_("IE PRUEBA 1234");

  const creadas = [];
  const omitidas = [];
  const reparadas = [];

  IES_PRUEBA_ADICIONALES.forEach(function(item){

    const nombreIE = item.ie;
    const filaExistente = filasExistentes[nombreIE.trim().toLowerCase()];

    if(filaExistente){
      omitidas.push(nombreIE + " (ya existía)");
      // Reparación: si ya existía pero le falta LINK_ACCESO (el
      // texto/enlace que arma la lista visual de links y códigos),
      // se completa con su propia URL_ACCESO — sin tocar nada más
      // de esa fila.
      if(mapa.LINK_ACCESO && mapa.URL_ACCESO){
        const linkActual = String(hoja.getRange(filaExistente, mapa.LINK_ACCESO).getDisplayValue()||"").trim();
        const urlExistente = String(hoja.getRange(filaExistente, mapa.URL_ACCESO).getDisplayValue()||"").trim();
        if(!linkActual && urlExistente){
          const richTextReparado = SpreadsheetApp.newRichTextValue()
            .setText("IE - " + nombreIE)
            .setLinkUrl(urlExistente)
            .build();
          hoja.getRange(filaExistente, mapa.LINK_ACCESO).setRichTextValue(richTextReparado);
          reparadas.push(nombreIE);
        }
      }
      return;
    }

    const token = Utilities.getUuid().replace(/-/g, "");
    const idForo = Utilities.getUuid();
    const codigo = generarCodigoAcceso_();
    const url = urlBase + "?t=" + encodeURIComponent(token);
    const fecha = new Date();

    const nuevaFila = new Array(hoja.getLastColumn()).fill("");
    function set(col, valor){ if(mapa[col]) nuevaFila[mapa[col]-1] = valor; }

    set("ID_ACCESO", Utilities.getUuid());
    set("IE", nombreIE);
    set("DANE", "PRUEBA-" + nombreIE.replace(/[^A-Za-z0-9]+/g, "").toUpperCase());
    set("CODIGO_ACCESO", codigo);
    set("TOKEN", token);
    set("URL_ACCESO", url);
    set("ID_FORO", idForo);
    set("ESTADO", "DISPONIBLE");
    set("EMAIL_IE", item.email);
    set("EMAIL_RESPONSABLE", item.email);
    set("TIPO", "PRUEBA");
    set("FECHA_GENERACION", fecha);
    if(logoIdPruebaOriginal) set("LOGO_ID", logoIdPruebaOriginal);
    // LINK_ACCESO (texto "IE - <nombre>" con el enlace real) es la
    // columna que arma la lista visual de links y códigos de la
    // hoja — sin ella la IE queda creada en AccesosIE, pero no
    // aparece en esa lista.
    if(mapa.LINK_ACCESO) set("LINK_ACCESO", "IE - " + nombreIE);

    hoja.appendRow(nuevaFila);
    SpreadsheetApp.flush();

    if(mapa.LINK_ACCESO){
      const filaNueva = hoja.getLastRow();
      const richText = SpreadsheetApp.newRichTextValue()
        .setText("IE - " + nombreIE)
        .setLinkUrl(url)
        .build();
      hoja.getRange(filaNueva, mapa.LINK_ACCESO).setRichTextValue(richText);
    }

    creadas.push(nombreIE);

    /*
     * Precargar caracterización y respuestas de muestra en
     * AvancesForo, para que el informe se pueda generar de una vez
     * (ver enviarTresCorreosIEsPruebaAdicionales).
     */
    const campos = {};
    function campoTexto(id, valor){ campos[id] = { tipo:"text", valor: valor }; }

    ["institucion","dane","direccion","zona","comuna","grupo","correoIE","nombre","correo","cargo"]
      .forEach(function(id){ campoTexto(id, TEXTO_MUESTRA_CARACTERIZACION_PRUEBA); });
    campoTexto("rector", TEXTO_MUESTRA_RECTOR_PRUEBA);
    // El correo institucional y el del responsable sí deben ser
    // reales (no el texto de muestra): son los que reciben el
    // informe y el comprobante de participación.
    campoTexto("correoIE", item.email);
    campoTexto("correo", item.email);

    campoTexto("respuestaSesion1", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);
    campoTexto("respuestaSesion1Pregunta2", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);
    campoTexto("respuestaSesion2Pregunta1", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);
    [1,2,3,4,5].forEach(function(n){ campoTexto("respuestaSesion2Pregunta2Accion"+n, TEXTO_MUESTRA_ACCION_PRUEBA); });
    campoTexto("respuestaSesion2Pregunta3", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);
    campoTexto("respuestaSesion2Pregunta4", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);
    campoTexto("respuestaSesion2Pregunta5", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);
    campoTexto("respuestaSesion3Pregunta1", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);
    [1,2,3,4,5].forEach(function(n){ campoTexto("respuestaSesion3Pregunta2Accion"+n, TEXTO_MUESTRA_ACCION_PRUEBA); });
    campoTexto("respuestaSesion3Pregunta3", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);
    campoTexto("respuestaSesion3Pregunta4", TEXTO_MUESTRA_RESPUESTA_LARGA_PRUEBA);

    ["Rector","Coordinador","Docentes","TutorPTA","Orientador","Estudiantes","Padres","Administrativos","Egresados","Sector","Otros"]
      .forEach(function(id, i){ campos["participantes"+id] = { tipo:"text", valor: i < 2 ? "1" : "0" }; });

    guardarAvanceForo({ idForo: idForo, institucion: nombreIE, dane: "", campos: campos });

  });

  const resumen = [
    "Creadas (" + creadas.length + "): " + (creadas.join(", ") || "(ninguna)"),
    "Omitidas, ya existían (" + omitidas.length + "): " + (omitidas.join(", ") || "(ninguna)"),
    "LINK_ACCESO reparado en filas que ya existían (" + reparadas.length + "): " + (reparadas.join(", ") || "(ninguna)"),
    logoIdPruebaOriginal
      ? "Logo reutilizado de IE PRUEBA 1234: " + logoIdPruebaOriginal
      : "IE PRUEBA 1234 no tiene logo vinculado todavía — las IE de prueba quedan sin logo."
  ];

  Logger.log(resumen.join("\n"));
  return { ok:true, creadas: creadas, omitidas: omitidas, reparadas: reparadas, resumen: resumen };

}


/*****************************************************
 * ENVIAR LOS 3 CORREOS A LAS IE DE PRUEBA ADICIONALES
 *
 * Ejecutar SOLO después de crearIEsPruebaAdicionales(). Para cada
 * IE de la lista IES_PRUEBA_ADICIONALES, en orden:
 *
 *   1. Correo de acceso (código + enlace) — mismo diseño que
 *      enviarAccesosTodasIE(), pero enviado solo a estas 9 filas
 *      (nunca recorre toda AccesosIE, así que ninguna IE oficial
 *      recibe nada de esto).
 *   2. Envío definitivo del foro con las respuestas de muestra ya
 *      precargadas (enviarForoDefinitivo), generación del informe
 *      real (generarInformeFEM) y su correo (enviarInformeFEM).
 *   3. Registro de una valoración de muestra (guardarValoracionFEM)
 *      y el correo de comprobante de participación
 *      (enviarComprobanteParticipacionFEM).
 *
 * Libera la sesión de prueba al final de cada IE para no dejarla
 * "tomada".
 *
 * Ejecutar manualmente desde el editor de Apps Script.
 *****************************************************/
function enviarAccesoIndividualIEPrueba_(nombreIE){

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  const ultimaFila = hoja.getLastRow();
  if(ultimaFila < 2) return { ok:false, mensaje:"AccesosIE no tiene filas." };

  const valores = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(function(f){ return String(f[mapa.IE-1]||"").trim() === nombreIE; });
  if(!fila) return { ok:false, mensaje:"No se encontró la IE " + nombreIE + " en AccesosIE." };

  const correoIE = String(fila[mapa.EMAIL_IE-1]||"").trim();
  if(!correoIE) return { ok:false, mensaje:"La IE " + nombreIE + " no tiene EMAIL_IE." };

  const ie = nombreIE;
  const ieSinPrefijo = nombreIESinPrefijoInstitucional_(ie);
  const logoIEUrlCorreo = mapa.LOGO_ID ? urlPublicaLogoDrive_(String(fila[mapa.LOGO_ID-1]||"").trim()) : "";
  const codigo = String(fila[mapa.CODIGO_ACCESO-1]||"").trim();
  const url = String(fila[mapa.URL_ACCESO-1]||"").trim();
  const correoResponsable = mapa.EMAIL_RESPONSABLE ? String(fila[mapa.EMAIL_RESPONSABLE-1]||"").trim() : "";

  const cuenta = Session.getEffectiveUser().getEmail().toLowerCase();
  const aliases = GmailApp.getAliases().map(function(a){ return a.toLowerCase(); });
  if(cuenta !== REMITENTE_FEM && aliases.indexOf(REMITENTE_FEM) === -1){
    throw new Error("La cuenta que ejecuta Apps Script no puede enviar como " + REMITENTE_FEM + ".");
  }

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

  const opciones = { htmlBody: cuerpoHTML, name: "Secretaría de Educación de Neiva", replyTo: REMITENTE_FEM };
  if(cuenta !== REMITENTE_FEM) opciones.from = REMITENTE_FEM;
  if(correoResponsable && correoResponsable !== correoIE) opciones.cc = correoResponsable;

  GmailApp.sendEmail(correoIE, asunto, cuerpoTexto, opciones);

  return { ok:true };

}

function enviarTresCorreosIEsPruebaAdicionales(){

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  const resultados = [];

  IES_PRUEBA_ADICIONALES.forEach(function(item){

    const nombreIE = item.ie;
    const paso = { ie: nombreIE, correo: false, informe: false, valoracion: false, errores: [] };

    try{

      const ultimaFila = hoja.getLastRow();
      const valores = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getDisplayValues();
      const fila = valores.find(function(f){ return String(f[mapa.IE-1]||"").trim() === nombreIE; });
      if(!fila) throw new Error("No existe en AccesosIE. Ejecute primero crearIEsPruebaAdicionales().");

      const idForo = String(fila[mapa.ID_FORO-1]||"").trim();
      const dispositivoId = "PRUEBA-DISPOSITIVO-" + idForo.slice(0, 8);

      // --- 1. Correo de acceso (código + enlace) ---
      enviarAccesoIndividualIEPrueba_(nombreIE);
      paso.correo = true;

      // --- 2. Envío definitivo, informe y su correo ---
      const sesion = reclamarSesionCodigo_("", "", dispositivoId, idForo, true);
      if(!sesion.ok) throw new Error("No fue posible reclamar la sesión de prueba: " + sesion.mensaje);

      const datosGuardados = obtenerDatosGuardadosPorIdForo_(idForo);
      if(!datosGuardados) throw new Error("No hay datos guardados para " + nombreIE + ". Ejecute primero crearIEsPruebaAdicionales().");
      datosGuardados.idForo = idForo;

      const envio = enviarForoDefinitivo(idForo, sesion.tokenSesion, dispositivoId, datosGuardados);
      if(!envio || !envio.ok) throw new Error((envio && envio.mensaje) || "No fue posible enviar el foro definitivo.");

      const informe = generarInformeFEM(idForo, datosGuardados);
      if(!informe || !informe.ok) throw new Error((informe && informe.mensaje) || "No fue posible generar el informe.");

      enviarInformeFEM(idForo, datosGuardados, informe.pdfId);
      paso.informe = true;

      // --- 3. Valoración de muestra y su correo de comprobante ---
      const valoracion = guardarValoracionFEM(idForo, { p1:5, p2:5, p3:5, p4:5, p5:"Valoración de muestra generada automáticamente para probar el envío del comprobante." });
      if(!valoracion || !valoracion.ok) throw new Error((valoracion && valoracion.mensaje) || "No fue posible guardar la valoración.");

      enviarComprobanteParticipacionFEM(idForo, datosGuardados);
      paso.valoracion = true;

      liberarSesionCodigo_("", "", dispositivoId, sesion.tokenSesion, idForo);

    }catch(error){
      paso.errores.push(error.message);
    }

    resultados.push(paso);
    Logger.log(nombreIE + " -> correo:" + paso.correo + " informe:" + paso.informe + " valoracion:" + paso.valoracion + (paso.errores.length ? " ERRORES: " + paso.errores.join(" | ") : ""));

  });

  Logger.log("========================================");
  Logger.log("RESUMEN ENVÍO 3 CORREOS — IE DE PRUEBA ADICIONALES");
  resultados.forEach(function(r){
    Logger.log(r.ie + ": " + (r.errores.length ? "⚠ " + r.errores.join(" | ") : "✅ correo + informe + valoración enviados"));
  });
  Logger.log("========================================");

  return { ok:true, resultados: resultados };

}


/*****************************************************
 * DIAGNOSTICAR Y LIBERAR IE DE PRUEBA ADICIONALES
 *
 * Para cada IE de IES_PRUEBA_ADICIONALES, revisa su fila real en
 * AccesosIE (existe/no existe, TOKEN, CODIGO_ACCESO, ESTADO, TIPO,
 * URL_ACCESO) y si tiene una sesión de prueba tomada en
 * PropertiesService (puede quedar así si
 * enviarTresCorreosIEsPruebaAdicionales() se interrumpió a mitad de
 * camino en alguna IE, antes de liberar esa sesión) — de ser así, la
 * libera. Un candado de sesión tomado por otro dispositivo es
 * exactamente lo que produce "no ingresa a la interfaz de sesiones":
 * el código se valida bien, pero reclamarSesionCodigo_ la rechaza
 * porque, según ScriptProperties, ya la tiene otro dispositivo (en
 * este caso, el de la propia prueba automática).
 *
 * No modifica AvancesForo, Participacion ni ninguna otra hoja: solo
 * lee AccesosIE y libera candados de sesión si los encuentra.
 *
 * Ejecutar manualmente desde el editor de Apps Script y revisar el
 * log — si describe algo distinto a "todo en orden", pégamelo.
 *****************************************************/
function diagnosticarYLiberarIEsPruebaAdicionales(){

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  const ultimaFila = hoja.getLastRow();
  const props = PropertiesService.getScriptProperties();
  const resumen = [];

  if(ultimaFila < 2){
    Logger.log("AccesosIE no tiene filas.");
    return { ok:false, mensaje:"AccesosIE no tiene filas." };
  }

  const valores = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getDisplayValues();

  IES_PRUEBA_ADICIONALES.forEach(function(item){

    const nombreIE = item.ie;
    const fila = valores.find(function(f){ return String(f[mapa.IE-1]||"").trim() === nombreIE; });

    if(!fila){
      resumen.push(nombreIE + " -> NO EXISTE en AccesosIE. Ejecute crearIEsPruebaAdicionales().");
      return;
    }

    const token = String(fila[mapa.TOKEN-1]||"").trim();
    const codigo = String(fila[mapa.CODIGO_ACCESO-1]||"").trim();
    const estado = String(fila[mapa.ESTADO-1]||"").trim();
    const tipo = mapa.TIPO ? String(fila[mapa.TIPO-1]||"").trim() : "";
    const url = mapa.URL_ACCESO ? String(fila[mapa.URL_ACCESO-1]||"").trim() : "";
    const idForo = String(fila[mapa.ID_FORO-1]||"").trim();

    let sesion = "sin sesión tomada";
    if(idForo){
      const clave = obtenerClaveSesionCodigo_("", "", idForo);
      const guardado = props.getProperty(clave);
      if(guardado){
        try{
          const actual = JSON.parse(guardado);
          sesion = "TOMADA por dispositivo \"" + actual.deviceId + "\" -> LIBERADA ahora";
        }catch(e){
          sesion = "propiedad de sesión ilegible -> LIBERADA ahora";
        }
        props.deleteProperty(clave);
      }
    }

    const problemas = [];
    if(!token) problemas.push("TOKEN vacío");
    if(!codigo) problemas.push("CODIGO_ACCESO vacío");
    if(tipo !== "PRUEBA") problemas.push("TIPO no es \"PRUEBA\" (es \"" + tipo + "\")");
    if(estado === "BLOQUEADO" || estado === "INACTIVO") problemas.push("ESTADO=" + estado + " (bloquea el ingreso)");
    if(!url) problemas.push("URL_ACCESO vacío");

    resumen.push(
      nombreIE + " -> ESTADO=" + estado + " | TIPO=" + tipo + " | CODIGO=" + codigo +
      " | sesión: " + sesion +
      (problemas.length ? " | ⚠ " + problemas.join("; ") : " | sin problemas detectados") +
      "\n   URL: " + url
    );

  });

  Logger.log(resumen.join("\n\n"));
  return { ok:true, resumen: resumen };

}


/*****************************************************
 * HACER PÚBLICOS LOS LOGOS/MARCO GLOBALES
 *
 * LOGO_ENCABEZADO_ID (FEM) y LOGO_PIE_ID (SEM) ya se usaban dentro
 * del informe generado (DriveApp.getFileById(...).getBlob(), con la
 * identidad del script — no necesita que el archivo sea público).
 * Pero para mostrarlos como <img> en el encabezado de la página web
 * (visible sin sesión de Google, para cualquier visitante) sí hace
 * falta que estén compartidos como "cualquiera con el enlace" — lo
 * mismo para MARCO_ACCESO_ID, la composición de fondo de la pantalla
 * de acceso, referenciada como URL fija en CSS.html.
 *
 * Ejecutar UNA sola vez manualmente desde el editor de Apps Script.
 *****************************************************/
function hacerPublicosLogosGlobales(){

  const ids = {
    "LOGO_ENCABEZADO_ID (FEM)": LOGO_ENCABEZADO_ID,
    "LOGO_PIE_ID (SEM)": LOGO_PIE_ID,
    "MARCO_ACCESO_ID": MARCO_ACCESO_ID,
    "DISENADOR_LOGO_ID": DISENADOR_LOGO_ID
  };

  const resumen = [];

  Object.keys(ids).forEach(function(nombre){
    const id = ids[nombre];
    try{
      const archivo = DriveApp.getFileById(id);
      hacerPublicoSiEsPosible_(archivo);
      resumen.push(nombre + " (" + id + "): compartido como \"cualquiera con el enlace\". Nombre real: " + archivo.getName());
    }catch(error){
      resumen.push(nombre + " (" + id + "): ERROR — " + error.message);
    }
  });

  Logger.log(resumen.join("\n"));
  return { ok:true, resumen: resumen };

}


/*****************************************************
 * RESET TOTAL DE PRODUCCIÓN — FEM 2026
 *
 * Borra TODAS las respuestas y todo registro de todas las IE en el
 * spreadsheet de origen (SPREADSHEET_ID):
 *
 *   - Deja únicamente las hojas elementales: Oficiales, AvancesForo,
 *     Participacion, AsistenciaQR y AccesosIE. Cualquier otra pestaña
 *     (Valoración FEMI2026, hojas propias por IE, etc.) se elimina
 *     por completo.
 *   - Vacía todas las filas de datos (deja solo la cabecera) de
 *     Oficiales, AvancesForo, Participacion, AsistenciaQR y
 *     AccesosIE — incluye el catálogo de instituciones y los
 *     códigos/tokens/enlaces de acceso de TODAS las IE, oficiales y
 *     de prueba.
 *   - En Drive, envía a la papelera todas las carpetas por IE dentro
 *     de DRIVE_CARPETA_FEM_ID (fotos, informes Doc y PDF de todas las
 *     instituciones).
 *   - Libera cualquier candado de sesión activa en ScriptProperties.
 *
 * Nada de esto es destrucción permanente inmediata: las hojas
 * eliminadas quedan en el historial de versiones del archivo, y los
 * archivos de Drive quedan en la papelera 30 días.
 *
 * Después de ejecutar esto hay que volver a cargar Oficiales y a
 * generar los accesos (generarAccesosIE() para las IE oficiales;
 * crearAccesoPrueba1234() y crearIEsPruebaAdicionales() para las 10
 * de prueba) antes de poder usar el formulario de nuevo.
 *
 * NO borra el documento de análisis (Análisis FEM 2026): es el
 * histórico separado, pensado para sobrevivir a un reset del origen.
 *
 * Ejecutar manualmente desde el editor de Apps Script.
 *****************************************************/
function resetTotalProduccionFEM(){

  const HOJAS_ELEMENTALES = [HOJA_OFICIALES, HOJA_AVANCES, HOJA_PARTICIPACION, HOJA_ASISTENCIA_QR, HOJA_ACCESOS];
  const resumen = [];
  const ss = abrirSpreadsheet_();

  let eliminadas = 0;
  ss.getSheets().forEach(function(hoja){
    const nombre = hoja.getName();
    if(HOJAS_ELEMENTALES.indexOf(nombre) === -1){
      ss.deleteSheet(hoja);
      eliminadas++;
      resumen.push("Hoja eliminada: " + nombre);
    }
  });
  resumen.push("Total de hojas adicionales eliminadas: " + eliminadas);

  HOJAS_ELEMENTALES.forEach(function(nombreHoja){
    const hoja = ss.getSheetByName(nombreHoja);
    if(!hoja){ resumen.push(nombreHoja + ": no existe."); return; }
    const ultima = hoja.getLastRow();
    const borradas = Math.max(ultima - 1, 0);
    if(ultima >= 2) hoja.deleteRows(2, ultima - 1);
    resumen.push(nombreHoja + ": vaciada (" + borradas + " fila(s) borrada(s)), cabecera intacta.");
  });

  let carpetasBorradas = 0, archivosBorrados = 0;
  try{
    const raiz = DriveApp.getFolderById(DRIVE_CARPETA_FEM_ID);
    const carpetas = raiz.getFolders();
    while(carpetas.hasNext()){
      const carpeta = carpetas.next();
      let archivosEnCarpeta = 0;
      const archivos = carpeta.getFiles();
      while(archivos.hasNext()){ archivos.next(); archivosEnCarpeta++; }
      carpeta.setTrashed(true);
      carpetasBorradas++;
      archivosBorrados += archivosEnCarpeta;
    }
    let sueltos = 0;
    const archivosSueltos = raiz.getFiles();
    while(archivosSueltos.hasNext()){ archivosSueltos.next().setTrashed(true); sueltos++; }
    resumen.push("Drive: " + carpetasBorradas + " carpeta(s) de IE enviadas a la papelera (" + archivosBorrados + " archivo(s): fotos, informes Doc y PDF).");
    if(sueltos) resumen.push("Drive: " + sueltos + " archivo(s) suelto(s) en la carpeta raíz también enviados a la papelera.");
  }catch(error){
    resumen.push("Drive: " + error.message);
  }

  try{
    const props = PropertiesService.getScriptProperties();
    const todas = props.getProperties();
    let liberados = 0;
    Object.keys(todas).forEach(function(clave){
      if(clave.indexOf("FEM_SESION_FORO_") === 0){ props.deleteProperty(clave); liberados++; }
    });
    resumen.push("Candados de sesión activa liberados: " + liberados + ".");
  }catch(error){
    resumen.push("Candados de sesión: " + error.message);
  }

  resumen.push("");
  resumen.push("⚠ Esto NO borra el documento de análisis (Análisis FEM 2026) ni el localStorage de los navegadores que ya usaron el formulario.");
  resumen.push("⚠ Los códigos de acceso de TODAS las IE (oficiales y de prueba) quedaron vacíos: hay que volver a generarlos.");
  resumen.push("⚠ Oficiales quedó vacía (solo cabecera): hay que volver a cargar el catálogo de instituciones antes de generar accesos.");

  Logger.log(resumen.join("\n"));
  return { ok:true, resumen: resumen };

}


/*****************************************************
 * RESTABLECER SOLO LOS ACCESOS DE LAS 37 IE OFICIALES
 *
 * generarAccesosIE() CONSERVA cualquier fila de AccesosIE cuya IE ya
 * exista (así esté rota o incompleta) — por diseño, para no pisar
 * códigos ya entregados. Eso significa que si las 37 IE oficiales
 * quedaron con filas a medias (por ejemplo, tras un
 * resetTotalProduccionFEM() a medio terminar, o una carga previa que
 * falló), volver a ejecutar generarAccesosIE() no arregla nada: las
 * sigue viendo como "ya existentes" y las salta.
 *
 * Esta función SÍ fuerza el restablecimiento completo: borra de
 * AccesosIE únicamente las filas de las 37 IE oficiales (identificadas
 * por nombre normalizado contra la hoja Oficiales) — sin tocar NINGUNA
 * fila de las IE de prueba (IE PRUEBA 1234, IE Prueba Ronald, etc.) —
 * y luego llama a generarAccesosIE(), que al no encontrarlas ya
 * existentes les genera código, contingencias, token, ID_FORO y
 * URL_ACCESO/LINK_ACCESO nuevos para las 37.
 *
 * Requisito de seguridad (heredado de generarAccesosIE()): la hoja
 * Oficiales debe tener exactamente 37 IE cargadas. Si no las tiene,
 * esta función se detiene sin borrar ni generar nada, y lo dice en el
 * log — hay que recargar Oficiales primero.
 *
 * Ejecutar manualmente:  restablecerAccesosOficialesFEM()
 *****************************************************/
function restablecerAccesosOficialesFEM(){
  const resultado = { pasos: {}, errores: [] };
  try{
    const instituciones = JSON.parse(obtenerInstitucionesJSON());
    const nombresOficiales = Object.keys(instituciones || {});
    Logger.log("IE encontradas en Oficiales: " + nombresOficiales.length);

    if(nombresOficiales.length !== 37){
      const mensaje = "ABORTADO: se esperaban exactamente 37 IE en Oficiales y se encontraron " + nombresOficiales.length + ". Recargue el catálogo de Oficiales antes de reintentar — no se borró ni se generó nada.";
      Logger.log("❌ " + mensaje);
      resultado.errores.push(mensaje);
      return { ok:false, resultado: resultado };
    }

    const clavesOficiales = {};
    nombresOficiales.forEach(function(n){ clavesOficiales[normalizarAccesoIE_(n)] = true; });

    const ss = abrirSpreadsheet_();
    const hoja = ss.getSheetByName(HOJA_ACCESOS);
    let borradas = 0;
    const nombresBorrados = [];

    if(!hoja){
      Logger.log("No existe todavía la hoja " + HOJA_ACCESOS + " — no hay filas que borrar, se continúa directo a generarAccesosIE().");
    }else{
      const m = mapaHoja_(hoja);
      const ultimaFila = hoja.getLastRow();
      if(ultimaFila >= 2 && m.IE){
        const valores = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
        for(let i = valores.length - 1; i >= 0; i--){
          const nombre = String(valores[i][m.IE - 1] || "").trim();
          if(nombre && clavesOficiales[normalizarAccesoIE_(nombre)]){
            hoja.deleteRow(i + 2);
            borradas++;
            nombresBorrados.push(nombre);
          }
        }
      }
    }
    resultado.pasos.filasOficialesBorradas = borradas;
    Logger.log("Filas de IE oficiales borradas de " + HOJA_ACCESOS + ": " + borradas + (borradas ? " (" + nombresBorrados.join(", ") + ")" : " (ninguna — probablemente ya estaban vacías o nunca se crearon)"));

    const generado = generarAccesosIE();
    resultado.pasos.generarAccesosIE = generado;
    Logger.log("Resultado de generarAccesosIE(): " + JSON.stringify(generado));
    if(!generado || !generado.ok){
      resultado.errores.push("generarAccesosIE() no terminó OK: " + (generado && generado.mensaje));
      return { ok:false, resultado: resultado };
    }

    /*
     * Imprime código y link de cada una de las 37, igual que
     * crearTodosLosAccesosDePruebaFEM() hace con las de prueba, para
     * poder copiarlos directo del log sin abrir la hoja.
     */
    const hojaFinal = ss.getSheetByName(HOJA_ACCESOS);
    const mFinal = mapaHoja_(hojaFinal);
    const filasFinal = hojaFinal.getRange(2, 1, hojaFinal.getLastRow() - 1, hojaFinal.getLastColumn()).getValues();
    Logger.log("========================================");
    Logger.log("LINKS DE LAS 37 IE OFICIALES");
    filasFinal.forEach(function(fila){
      const nombre = String(fila[mFinal.IE - 1] || "").trim();
      if(!nombre || !clavesOficiales[normalizarAccesoIE_(nombre)]) return;
      const codigo = mFinal.CODIGO_ACCESO ? String(fila[mFinal.CODIGO_ACCESO - 1] || "") : "";
      const link = (mFinal.URL_ACCESO ? String(fila[mFinal.URL_ACCESO - 1] || "") : "") || (mFinal.LINK_ACCESO ? String(fila[mFinal.LINK_ACCESO - 1] || "") : "");
      Logger.log(nombre + " -> código: " + codigo + " | " + link);
    });
    Logger.log("========================================");
    Logger.log("✅ LISTO: las 37 IE oficiales quedaron con accesos nuevos y funcionales.");

    return { ok:true, resultado: resultado };

  }catch(error){
    resultado.errores.push(error.message);
    Logger.log("❌ ERROR: " + error.message);
    return { ok:false, resultado: resultado };
  }
}


/*****************************************************
 * REPARAR/CORREGIR EMAIL_IE DE LAS 37 IE OFICIALES SIN TOCAR CÓDIGOS
 *
 * Dos problemas distintos, corregidos en el mismo lugar:
 *  1) generarAccesosIE() (Código.js) nunca copiaba el correo
 *     institucional desde Oficiales a la columna EMAIL_IE de
 *     AccesosIE al crear un acceso nuevo — quedaba vacía.
 *  2) Incluso cuando SÍ había algo en EMAIL_IE, la columna "E-MAIL
 *     INSTITUCIONAL" de Oficiales traía direcciones que no
 *     correspondían a la IE real: los correos se enviaban (Apps
 *     Script no valida el buzón al encolar el envío, por eso la
 *     cuota de envíos bajaba igual) pero nunca llegaban a destino.
 *
 * Esta función usa CORRECCION_EMAIL_POR_DANE_ (Código.js — la lista
 * de correos verificada por la Secretaría) como fuente de verdad, y
 * si no hay corrección para una IE, cae de vuelta al correo de
 * Oficiales. Corrige EMAIL_IE tanto si estaba vacío como si tenía un
 * valor distinto al correcto — nunca toca CODIGO_ACCESO, TOKEN,
 * URL_ACCESO ni ID_FORO, así que los códigos/enlaces ya generados (y
 * que ya podrían haberse compartido) siguen siendo válidos.
 *
 * Ejecutar manualmente:  repararEmailIEOficialesFEM()
 *****************************************************/
function repararEmailIEOficialesFEM(){
  const resultado = { pasos: {}, errores: [] };
  try{
    const instituciones = JSON.parse(obtenerInstitucionesJSON());
    const nombresOficiales = Object.keys(instituciones || {});
    Logger.log("IE encontradas en Oficiales: " + nombresOficiales.length);
    if(nombresOficiales.length !== 37){
      const mensaje = "ABORTADO: se esperaban exactamente 37 IE en Oficiales y se encontraron " + nombresOficiales.length + ".";
      Logger.log("❌ " + mensaje);
      resultado.errores.push(mensaje);
      return { ok:false, resultado: resultado };
    }

    const hoja = asegurarColumnasAccesosIE_();
    const m = mapaHoja_(hoja);
    if(hoja.getLastRow() < 2){
      const mensaje = "AccesosIE no tiene filas.";
      Logger.log(mensaje);
      return { ok:false, mensaje: mensaje };
    }
    const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();

    let reparadas = 0, yaEstabanCorrectas = 0, sinCorreoDisponible = 0, noEncontradas = 0;
    const detalle = [];

    nombresOficiales.forEach(function(nombreIE){
      const indiceFila = valores.findIndex(function(f){ return normalizarAccesoIE_(String(f[m.IE - 1] || "")) === normalizarAccesoIE_(nombreIE); });
      if(indiceFila === -1){ noEncontradas++; detalle.push(nombreIE + ": ⚠ no existe en AccesosIE (ejecute restablecerAccesosOficialesFEM() primero)."); return; }

      const fila = valores[indiceFila];
      const correoActual = String(fila[m.EMAIL_IE - 1] || "").trim();
      /*
       * CORRECCION_EMAIL_POR_DANE_ (Código.js) tiene prioridad,
       * incluso si ya había ALGO en EMAIL_IE: la hoja Oficiales traía
       * correos que no correspondían a la IE real, así que "ya tenía
       * un correo" no significaba "tenía el correo correcto".
       */
      const dane = m.DANE ? String(fila[m.DANE - 1] || "").trim() : "";
      const correoCorregido = dane ? obtenerCorreoCorregidoPorDane_(dane) : "";
      const correoObjetivo = correoCorregido || String(instituciones[nombreIE]?.correo || "").trim();

      if(!correoObjetivo){ sinCorreoDisponible++; detalle.push(nombreIE + ": ⚠ no hay correo verificado ni en Oficiales para esta IE."); return; }

      if(correoActual === correoObjetivo){ yaEstabanCorrectas++; return; }

      hoja.getRange(indiceFila + 2, m.EMAIL_IE).setValue(correoObjetivo);
      reparadas++;
      detalle.push(nombreIE + ": ✅ EMAIL_IE " + (correoActual ? "corregido de \"" + correoActual + "\" a" : "completado con") + " \"" + correoObjetivo + "\"" + (correoCorregido ? " (fuente: lista verificada)" : " (fuente: Oficiales)") + ".");
    });

    Logger.log("========================================");
    Logger.log("REPARACIÓN DE EMAIL_IE — 37 IE OFICIALES");
    detalle.forEach(function(d){ Logger.log(d); });
    Logger.log("Reparadas/corregidas: " + reparadas + " | Ya estaban correctas: " + yaEstabanCorrectas + " | Sin correo disponible: " + sinCorreoDisponible + " | No encontradas en AccesosIE: " + noEncontradas);
    Logger.log("========================================");

    return { ok:true, reparadas: reparadas, yaEstabanCorrectas: yaEstabanCorrectas, sinCorreoDisponible: sinCorreoDisponible, noEncontradas: noEncontradas };
  }catch(error){
    resultado.errores.push(error.message);
    Logger.log("❌ ERROR: " + error.message);
    return { ok:false, resultado: resultado };
  }
}


/*****************************************************
 * ENVIAR CORREO DE ACCESO A TODAS LAS IE DE PRUEBA
 *
 * Envía el mismo correo real de "acceso al Foro"
 * (construirCorreoAccesoIE_, Código.js) restringido a las 10 IE de
 * prueba (IE PRUEBA 1234 + IES_PRUEBA_ADICIONALES), a sus EMAIL_IE
 * reales registrados en AccesosIE — útil para verificar cómo se ve y
 * se comporta el correo real sin tocar ninguna IE oficial.
 *
 * Ejecutar manualmente:  enviarAccesosTodasLasPruebasFEM()
 *****************************************************/
function enviarAccesosTodasLasPruebasFEM(){
  const cuenta = Session.getEffectiveUser().getEmail().toLowerCase();
  const aliases = GmailApp.getAliases().map(function(a){ return a.toLowerCase(); });
  if(cuenta !== REMITENTE_FEM && aliases.indexOf(REMITENTE_FEM) === -1){
    const mensaje = "La cuenta que ejecuta Apps Script no puede enviar como " + REMITENTE_FEM + ".";
    Logger.log("❌ " + mensaje);
    return { ok:false, mensaje: mensaje };
  }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2){
    Logger.log("AccesosIE no tiene filas.");
    return { ok:false, mensaje:"AccesosIE no tiene filas." };
  }
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const nombresPrueba = ["IE PRUEBA 1234"].concat(IES_PRUEBA_ADICIONALES.map(function(x){ return x.ie; }));

  let enviados = 0;
  const resultados = [];

  nombresPrueba.forEach(function(nombreIE){
    try{
      const fila = valores.find(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIE; });
      if(!fila) throw new Error("No existe en AccesosIE.");
      const correoIE = String(fila[mapa.EMAIL_IE - 1] || "").trim();
      if(!correoIE) throw new Error("Sin EMAIL_IE.");
      const codigo = String(fila[mapa.CODIGO_ACCESO - 1] || "").trim();
      const url = String(fila[mapa.URL_ACCESO - 1] || "").trim();
      const correoResponsable = mapa.EMAIL_RESPONSABLE ? String(fila[mapa.EMAIL_RESPONSABLE - 1] || "").trim() : "";
      const ieSinPrefijo = nombreIESinPrefijoInstitucional_(nombreIE);
      const logoIEUrlCorreo = mapa.LOGO_ID ? urlPublicaLogoDrive_(String(fila[mapa.LOGO_ID - 1] || "").trim()) : "";

      const correoArmado = construirCorreoAccesoIE_(nombreIE, ieSinPrefijo, codigo, url, logoIEUrlCorreo);
      const opciones = { htmlBody: correoArmado.cuerpoHTML, name: "Secretaría de Educación de Neiva", replyTo: REMITENTE_FEM };
      if(cuenta !== REMITENTE_FEM) opciones.from = REMITENTE_FEM;
      if(correoResponsable && correoResponsable !== correoIE) opciones.cc = correoResponsable;

      GmailApp.sendEmail(correoIE, correoArmado.asunto, correoArmado.cuerpoTexto, opciones);
      enviados++;
      resultados.push(nombreIE + ": ✅ enviado a " + correoIE + (opciones.cc ? " (cc " + opciones.cc + ")" : ""));
    }catch(error){
      resultados.push(nombreIE + ": ⚠ " + error.message);
    }
  });

  Logger.log("========================================");
  Logger.log("ENVÍO DE ACCESOS A LAS " + nombresPrueba.length + " IE DE PRUEBA");
  resultados.forEach(function(r){ Logger.log(r); });
  Logger.log("Enviados: " + enviados + " / " + nombresPrueba.length);
  Logger.log("========================================");

  return { ok:true, enviados: enviados, total: nombresPrueba.length, resultados: resultados };
}


/*****************************************************
 * SIMULAR EL CORREO DE ACCESO DE UNA IE OFICIAL PUNTUAL
 * (por defecto, la que contenga "Limonar" en el nombre) — enviado
 * ÚNICAMENTE al correo del administrador, nunca a los correos reales
 * registrados de esa IE.
 *
 * Sirve para revisar cómo se vería el correo real de una IE oficial
 * concreta (logo, código, enlace) sin arriesgarse a que le llegue a
 * la propia institución ni a su responsable: el EMAIL_IE/
 * EMAIL_RESPONSABLE de la fila real NUNCA se usan como destinatario
 * ni como copia en esta función — solo se leen el nombre, el código
 * y el enlace para armar la vista previa.
 *
 * Ejecutar manualmente:  simularCorreoAccesoIEFEM("Limonar")
 * (sin argumento, usa "Limonar" por defecto)
 *****************************************************/
function simularCorreoAccesoIEFEM(nombreIEBuscado){
  const CORREO_SIMULACION = "jhonefrainsanchez@gmail.com";
  nombreIEBuscado = String(nombreIEBuscado || "Limonar").trim();

  const cuenta = Session.getEffectiveUser().getEmail().toLowerCase();
  const aliases = GmailApp.getAliases().map(function(a){ return a.toLowerCase(); });
  if(cuenta !== REMITENTE_FEM && aliases.indexOf(REMITENTE_FEM) === -1){
    const mensaje = "La cuenta que ejecuta Apps Script no puede enviar como " + REMITENTE_FEM + ".";
    Logger.log("❌ " + mensaje);
    return { ok:false, mensaje: mensaje };
  }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2){
    const mensaje = "AccesosIE no tiene filas.";
    Logger.log(mensaje);
    return { ok:false, mensaje: mensaje };
  }
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();

  // Búsqueda flexible (sin tildes/mayúsculas, por si el nombre exacto
  // en AccesosIE trae un prefijo distinto, p. ej. "I.E. Limonar").
  const claveBuscada = normalizarNombreIE_(nombreIEBuscado);
  const fila = valores.find(function(f){ return normalizarNombreIE_(String(f[mapa.IE - 1] || "")).indexOf(claveBuscada) !== -1; });
  if(!fila){
    const mensaje = 'No se encontró ninguna IE que contenga "' + nombreIEBuscado + '" en AccesosIE.';
    Logger.log("❌ " + mensaje);
    return { ok:false, mensaje: mensaje };
  }

  const nombreIEReal = String(fila[mapa.IE - 1] || "").trim();
  const codigo = String(fila[mapa.CODIGO_ACCESO - 1] || "").trim();
  const url = String(fila[mapa.URL_ACCESO - 1] || "").trim();
  const ieSinPrefijo = nombreIESinPrefijoInstitucional_(nombreIEReal);
  const logoIEUrlCorreo = mapa.LOGO_ID ? urlPublicaLogoDrive_(String(fila[mapa.LOGO_ID - 1] || "").trim()) : "";

  const correoArmado = construirCorreoAccesoIE_(nombreIEReal, ieSinPrefijo, codigo, url, logoIEUrlCorreo);

  const avisoSimulacion =
    "<div style=\"background:#FFF3CD;border-left:6px solid #C62828;border-radius:10px;padding:12px 16px;margin:14px auto 0;max-width:520px;font-family:Arial,Helvetica,sans-serif;\">" +
    "<p style=\"font-size:13px;color:#7A5B00;margin:0;\"><strong>🧪 SIMULACIÓN interna:</strong> este correo es una vista previa de lo que recibiría la IE <strong>" + nombreIEReal + "</strong>. Se envió únicamente a " + CORREO_SIMULACION + " — NO se envió a ningún correo registrado de esa institución.</p>" +
    "</div>";

  const opciones = {
    htmlBody: correoArmado.cuerpoHTML + avisoSimulacion,
    name: "Secretaría de Educación de Neiva",
    replyTo: REMITENTE_FEM
  };
  if(cuenta !== REMITENTE_FEM) opciones.from = REMITENTE_FEM;

  /*
   * Destinatario fijo, sin excepción: nunca correoIE ni
   * correoResponsable de la fila real.
   */
  GmailApp.sendEmail(
    CORREO_SIMULACION,
    "[SIMULACIÓN] " + correoArmado.asunto,
    "[SIMULACIÓN — vista previa del correo real de " + nombreIEReal + ", no enviado a la institución]\n\n" + correoArmado.cuerpoTexto,
    opciones
  );

  Logger.log("========================================");
  Logger.log("✅ Simulación enviada a " + CORREO_SIMULACION + " con el contenido real de: " + nombreIEReal);
  Logger.log("Código: " + codigo + " | Link: " + url);
  Logger.log("========================================");

  return { ok:true, ieSimulada: nombreIEReal, enviadoA: CORREO_SIMULACION };
}


/*****************************************************
 * ENVIAR CORREO DE ACCESO A TODAS LAS IE OFICIALES (SIN PRUEBAS)
 *
 * Envía el correo real de acceso (construirCorreoAccesoIE_) a las 37
 * IE oficiales de la hoja Oficiales — igual que enviarAccesosTodasIE(),
 * pero excluyendo explícitamente cualquier IE de prueba (IE PRUEBA
 * 1234, IE Prueba Ronald, etc.), aunque estas también tengan
 * ESTADO=DISPONIBLE en AccesosIE.
 *
 * Requisito de seguridad, igual que generarAccesosIE()/
 * restablecerAccesosOficialesFEM(): Oficiales debe tener exactamente
 * 37 IE cargadas — si no, se aborta sin enviar nada.
 *
 * Ejecutar manualmente:  enviarAccesosSoloOficialesFEM()
 *****************************************************/
function enviarAccesosSoloOficialesFEM(){
  const cuenta = Session.getEffectiveUser().getEmail().toLowerCase();
  const aliases = GmailApp.getAliases().map(function(a){ return a.toLowerCase(); });
  if(cuenta !== REMITENTE_FEM && aliases.indexOf(REMITENTE_FEM) === -1){
    const mensaje = "La cuenta que ejecuta Apps Script no puede enviar como " + REMITENTE_FEM + ".";
    Logger.log("❌ " + mensaje);
    return { ok:false, mensaje: mensaje };
  }

  const instituciones = JSON.parse(obtenerInstitucionesJSON());
  const nombresOficiales = Object.keys(instituciones || {});
  if(nombresOficiales.length !== 37){
    const mensaje = "ABORTADO: se esperaban exactamente 37 IE en Oficiales y se encontraron " + nombresOficiales.length + ". No se envió ningún correo.";
    Logger.log("❌ " + mensaje);
    return { ok:false, mensaje: mensaje };
  }
  const clavesOficiales = {};
  nombresOficiales.forEach(function(n){ clavesOficiales[normalizarAccesoIE_(n)] = true; });

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2){
    const mensaje = "AccesosIE no tiene filas.";
    Logger.log(mensaje);
    return { ok:false, mensaje: mensaje };
  }
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();

  /*
   * El correo verificado en CORRECCION_EMAIL_POR_DANE_ (Código.js)
   * tiene prioridad sobre lo que haya en la columna EMAIL_IE: se
   * detectó que esa columna traía direcciones que no correspondían a
   * la IE real, así que los correos se enviaban (Apps Script no
   * valida el buzón al encolar el envío) pero nunca llegaban.
   */
  function correoRealDeFila_(fila){
    const dane = mapa.DANE ? String(fila[mapa.DANE - 1] || "").trim() : "";
    const corregido = dane ? obtenerCorreoCorregidoPorDane_(dane) : "";
    return corregido || String(fila[mapa.EMAIL_IE - 1] || "").trim();
  }

  // Filas que realmente se van a enviar (oficial + DISPONIBLE + con
  // correo, ya sea corregido o el de la hoja), calculadas de antemano
  // para poder elegir una al azar y marcarla con copia oculta a modo
  // de verificación de que el envío masivo sí está saliendo de verdad.
  const filasAEnviar = valores.filter(function(fila){
    const nombreIE = String(fila[mapa.IE - 1] || "").trim();
    if(!nombreIE || !clavesOficiales[normalizarAccesoIE_(nombreIE)]) return false;
    if(String(fila[mapa.ESTADO - 1] || "").trim().toUpperCase() !== "DISPONIBLE") return false;
    if(!correoRealDeFila_(fila)) return false;
    return true;
  });

  const CORREO_VERIFICACION = "jhonefrainsanchez@gmail.com";
  const filaConCopiaOculta = filasAEnviar.length
    ? filasAEnviar[Math.floor(Math.random() * filasAEnviar.length)]
    : null;
  const ieConCopiaOculta = filaConCopiaOculta ? String(filaConCopiaOculta[mapa.IE - 1] || "").trim() : "";

  let enviados = 0, omitidos = 0, corregidos = 0;
  const errores = [];

  valores.forEach(function(fila, indice){
    const nombreIE = String(fila[mapa.IE - 1] || "").trim();
    // No es una de las 37 oficiales (incluye a todas las de prueba): se ignora.
    if(!nombreIE || !clavesOficiales[normalizarAccesoIE_(nombreIE)]) return;

    const estado = String(fila[mapa.ESTADO - 1] || "").trim().toUpperCase();
    if(estado !== "DISPONIBLE"){ omitidos++; return; }

    const correoIE = correoRealDeFila_(fila);
    if(!correoIE){ omitidos++; return; }

    // Deja la hoja corregida para la próxima vez, si el correo real
    // no coincidía con el que ya estaba guardado en EMAIL_IE.
    const correoGuardado = String(fila[mapa.EMAIL_IE - 1] || "").trim();
    if(correoGuardado !== correoIE){
      hoja.getRange(indice + 2, mapa.EMAIL_IE).setValue(correoIE);
      corregidos++;
    }

    try{
      const codigo = String(fila[mapa.CODIGO_ACCESO - 1] || "").trim();
      const url = String(fila[mapa.URL_ACCESO - 1] || "").trim();
      const correoResponsable = mapa.EMAIL_RESPONSABLE ? String(fila[mapa.EMAIL_RESPONSABLE - 1] || "").trim() : "";
      const ieSinPrefijo = nombreIESinPrefijoInstitucional_(nombreIE);
      const logoIEUrlCorreo = mapa.LOGO_ID ? urlPublicaLogoDrive_(String(fila[mapa.LOGO_ID - 1] || "").trim()) : "";

      const correoArmado = construirCorreoAccesoIE_(nombreIE, ieSinPrefijo, codigo, url, logoIEUrlCorreo);
      const opciones = { htmlBody: correoArmado.cuerpoHTML, name: "Secretaría de Educación de Neiva", replyTo: REMITENTE_FEM };
      if(cuenta !== REMITENTE_FEM) opciones.from = REMITENTE_FEM;
      if(correoResponsable && correoResponsable !== correoIE) opciones.cc = correoResponsable;
      if(fila === filaConCopiaOculta) opciones.bcc = CORREO_VERIFICACION;

      GmailApp.sendEmail(correoIE, correoArmado.asunto, correoArmado.cuerpoTexto, opciones);
      enviados++;
    }catch(error){
      errores.push({ ie: nombreIE, correo: correoIE, mensaje: error.message });
    }
  });

  Logger.log("========================================");
  Logger.log("ENVÍO DE ACCESOS A LAS 37 IE OFICIALES (sin IE de prueba)");
  Logger.log("Enviados: " + enviados + " | Omitidos (sin ESTADO=DISPONIBLE o sin correo): " + omitidos + " | Corregidos en la hoja (EMAIL_IE no coincidía con el correo verificado): " + corregidos + " | Errores: " + errores.length);
  if(ieConCopiaOculta) Logger.log("Verificación: se envió copia oculta (BCC) a " + CORREO_VERIFICACION + " del correo real de la IE elegida al azar: " + ieConCopiaOculta + ".");
  if(errores.length) Logger.log(JSON.stringify(errores, null, 2));
  Logger.log("========================================");

  return { ok: errores.length === 0, enviados: enviados, omitidos: omitidos, corregidos: corregidos, errores: errores, ieConCopiaOculta: ieConCopiaOculta };
}


/*****************************************************
 * CONSTRUIR / RECONSTRUIR EL DOCUMENTO DE ANÁLISIS — FEM 2026
 *
 * Crea (si no existe) el documento de análisis separado y reconstruye
 * "Respuestas Totales", "Gráficos" y una hoja por cada IE que ya
 * tenga respuestas en AvancesForo — en orden alfabético. Es la misma
 * lógica que se dispara sola con cada envío definitivo o valoración
 * (actualizarAnalisisFEMIndividual_), pero recorriendo TODAS las IE
 * de una sola vez: útil después de un reset, o para confirmar que el
 * documento quedó al día y refrescar los gráficos.
 *
 * Ejecutar manualmente:  reconstruirAnalisisFEM()
 *****************************************************/
function reconstruirAnalisisFEM(){
  const sh = abrirSpreadsheet_().getSheetByName(HOJA_AVANCES);
  if(!sh || sh.getLastRow() < 2){
    const ss = obtenerSpreadsheetAnalisisFEM_();
    const mensajeVacio = "AvancesForo no tiene filas todavía. Documento de análisis: " + ss.getUrl();
    Logger.log(mensajeVacio);
    return { ok:true, mensaje: mensajeVacio, procesadas:0, url: ss.getUrl() };
  }
  const m = mapaHoja_(sh);
  if(!m.ID_FORO){ Logger.log("AvancesForo no tiene columna ID_FORO."); return { ok:false, mensaje:"AvancesForo no tiene columna ID_FORO." }; }

  const ids = sh.getRange(2, m.ID_FORO, sh.getLastRow() - 1, 1).getDisplayValues().map(f => String(f[0] || "").trim()).filter(Boolean);

  let procesadas = 0;
  ids.forEach(function(idForo){
    try{ actualizarAnalisisFEMIndividual_(idForo); procesadas++; }
    catch(error){ Logger.log("Reconstruir análisis — " + idForo + ": " + error.message); }
  });

  const ss = obtenerSpreadsheetAnalisisFEM_();
  try{ actualizarGraficosAnalisisFEM_(ss); }catch(error){ Logger.log("Reconstruir análisis — gráficos: " + error.message); }
  try{ reordenarHojasAnalisisFEM_(ss); }catch(error){ Logger.log("Reconstruir análisis — orden de hojas: " + error.message); }

  const mensaje = "Documento de análisis reconstruido: " + procesadas + " de " + ids.length + " IE procesadas. URL: " + ss.getUrl();
  Logger.log(mensaje);
  return { ok:true, procesadas: procesadas, total: ids.length, url: ss.getUrl() };
}


/*****************************************************
 * PRUEBAS DE RETORNO DE DATOS — errores de autoguardado,
 * guardado local y respuestas no enviadas
 *
 * Ejecutar cada una manualmente desde el editor de Apps Script.
 *****************************************************/

// El caso real más común de "autoguardado que falla en plenaria":
// el ID_FORO ya no existe en AccesosIE (token vencido, o el
// navegador guardó un ID_FORO viejo de otra prueba). Debe responder
// con un mensaje claro, no lanzar un error sin explicación.
function probarAutoguardadoConIdForoInvalido(){
  const resultado = guardarAvanceForo({ idForo: "ID-INEXISTENTE-" + Utilities.getUuid(), campos: { respuestaSesion1: { tipo:"text", valor:"prueba" } } });
  const ok = !resultado.ok && /no est[aá] autorizado/i.test(resultado.mensaje || "");
  Logger.log("Autoguardado con ID_FORO inválido -> " + (ok ? "✅ CORRECTO" : "⚠ INESPERADO") + ": " + JSON.stringify(resultado));
  return { ok: ok, resultado: resultado };
}

// Si el navegador llega a llamar al autoguardado sin datos (por
// ejemplo, un JSON corrupto reconstruido desde localStorage), debe
// rechazarse con un mensaje, no reventar el servidor.
function probarAutoguardadoSinDatos(){
  let resultado;
  try{ resultado = guardarAvanceForo(null); }
  catch(error){ resultado = { ok:false, mensaje:error.message }; }
  const ok = !resultado.ok;
  Logger.log("Autoguardado sin datos -> " + (ok ? "✅ CORRECTO (rechazado)" : "⚠ INESPERADO") + ": " + JSON.stringify(resultado));
  return { ok: ok, resultado: resultado };
}

// El guardado en localStorage del navegador vive en cada equipo y no
// se puede probar desde el servidor. Lo que sí se puede probar desde
// aquí es el escenario que esa falla produce en la práctica: el
// formulario, al no saber si el primer intento quedó guardado,
// reintenta guardarAvanceForo() dos veces seguidas con el mismo
// ID_FORO. Verifica que la segunda llamada ACTUALICE la misma fila
// en vez de crear una fila duplicada.
function probarReintentoPorFallaGuardadoLocal(nombreIEPrueba){
  nombreIEPrueba = nombreIEPrueba || "IE PRUEBA 1234";
  function salir(resultado){ Logger.log(resultado.mensaje); return resultado; }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2) return salir({ ok:false, mensaje:"AccesosIE no tiene filas." });
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIEPrueba);
  if(!fila) return salir({ ok:false, mensaje:"No existe " + nombreIEPrueba + " en AccesosIE. Ejecute primero crearAccesoPrueba1234() o crearIEsPruebaAdicionales()." });
  const idForo = String(fila[mapa.ID_FORO - 1] || "").trim();

  const shAvances = abrirSpreadsheet_().getSheetByName(HOJA_AVANCES);
  const filasAntes = shAvances ? shAvances.getLastRow() : 0;

  const datos = { idForo: idForo, campos: { respuestaSesion1: { tipo:"text", valor:"Primer intento — " + new Date().toISOString() } } };
  guardarAvanceForo(datos);
  datos.campos.respuestaSesion1.valor = "Reintento tras falla de guardado local — " + new Date().toISOString();
  guardarAvanceForo(datos);

  const filasDespues = shAvances.getLastRow();
  const ok = filasDespues <= Math.max(filasAntes, 2);
  Logger.log("Reintento por falla de guardado local -> filas antes: " + filasAntes + ", después: " + filasDespues + " -> " + (ok ? "✅ no se duplicó" : "⚠ POSIBLE DUPLICADO"));
  return { ok: ok, filasAntes: filasAntes, filasDespues: filasDespues };
}

// Una IE que llena sesiones pero nunca presiona "Enviar" al final NO
// debe quedar marcada como ENVIADO en AccesosIE ni contarse como
// respuesta definitiva.
function probarRespuestaNoEnviada(nombreIEPrueba){
  nombreIEPrueba = nombreIEPrueba || "IE Prueba Rosa";
  function salir(resultado){ Logger.log(resultado.mensaje); return resultado; }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2) return salir({ ok:false, mensaje:"AccesosIE no tiene filas." });
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIEPrueba);
  if(!fila) return salir({ ok:false, mensaje:"No existe " + nombreIEPrueba + " en AccesosIE." });
  const idForo = String(fila[mapa.ID_FORO - 1] || "").trim();
  const estadoAcceso = String(fila[mapa.ESTADO - 1] || "").trim().toUpperCase();

  const estadoSesiones = obtenerEstadoSesiones_(idForo);
  const ok = estadoAcceso !== "ENVIADO";
  Logger.log(nombreIEPrueba + " (sin envío definitivo) -> ESTADO en AccesosIE: " + estadoAcceso + ", sesiones enviadas: " + JSON.stringify(estadoSesiones) + " -> " + (ok ? "✅ CORRECTO (no cuenta como enviada)" : "⚠ INESPERADO, ya estaba ENVIADO"));
  return { ok: ok, estadoAcceso: estadoAcceso, estadoSesiones: estadoSesiones };
}


/*****************************************************
 * PRUEBA: PLENARIA -> ENVÍO DEFINITIVO -> DOCUMENTO DE ANÁLISIS
 *
 * Verifica de punta a punta que, una vez una IE de prueba envía sus
 * respuestas como definitivas (como ocurre en plenaria), esas
 * respuestas quedan reflejadas automáticamente en el documento de
 * análisis separado: su fila en "Respuestas Totales" y su propia
 * hoja de detalle.
 *
 * Usa "IE PRUEBA 1234" — no envía ningún correo real adicional.
 *
 * Ejecutar manualmente:  probarFlujoPlenariaHastaDocumentoAnalisis()
 *****************************************************/
function probarFlujoPlenariaHastaDocumentoAnalisis(){
  const nombreIE = "IE PRUEBA 1234";

  // Toda salida (éxito o falla) pasa por aquí — así una ejecución
  // desde el editor SIEMPRE deja algo en el registro, en vez de
  // terminar en silencio ("Se completó la ejecución" sin más detalle)
  // cuando falla en un paso intermedio.
  function salir(resultado){ Logger.log(JSON.stringify(resultado)); return resultado; }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2) return salir({ ok:false, mensaje:"AccesosIE no tiene filas." });
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIE);
  if(!fila) return salir({ ok:false, mensaje:"No existe " + nombreIE + ". Ejecute primero crearAccesoPrueba1234()." });
  const idForo = String(fila[mapa.ID_FORO - 1] || "").trim();
  const dispositivoId = "PRUEBA-ANALISIS-" + idForo.slice(0, 8);

  const datosGuardados = obtenerDatosGuardadosPorIdForo_(idForo);
  if(!datosGuardados) return salir({ ok:false, mensaje:"No hay datos guardados para " + nombreIE + ". Complete o precargue su caracterización primero (ej. crearIEsPruebaAdicionales(), o llene el formulario a mano)." });
  datosGuardados.idForo = idForo;

  const sesion = reclamarSesionCodigo_("", "", dispositivoId, idForo, true);
  if(!sesion.ok) return salir({ ok:false, mensaje:"No fue posible reclamar sesión: " + sesion.mensaje });

  const envio = enviarForoDefinitivo(idForo, sesion.tokenSesion, dispositivoId, datosGuardados);
  liberarSesionCodigo_("", "", dispositivoId, sesion.tokenSesion, idForo);
  if(!envio || (!envio.ok && !envio.yaEnviado)) return salir({ ok:false, mensaje:(envio && envio.mensaje) || "Envío definitivo falló." });

  const ss = obtenerSpreadsheetAnalisisFEM_();
  const shTotales = ss.getSheetByName(HOJA_ANALISIS_TOTALES);
  const mTotales = shTotales ? mapaHoja_(shTotales) : {};
  let encontradaEnTotales = false;
  if(shTotales && shTotales.getLastRow() >= 2 && mTotales.ID_FORO){
    const ids = shTotales.getRange(2, mTotales.ID_FORO, shTotales.getLastRow() - 1, 1).getDisplayValues();
    encontradaEnTotales = ids.some(r => String(r[0] || "").trim() === idForo);
  }
  const shIE = ss.getSheetByName(nombreHojaIE_(nombreIE));
  const encontradaHojaIE = !!(shIE && shIE.getLastRow() >= 2);

  const ok = encontradaEnTotales && encontradaHojaIE;
  const resumen = "Envío definitivo: ok. En '" + HOJA_ANALISIS_TOTALES + "': " + (encontradaEnTotales ? "sí" : "NO") + ". En hoja propia de la IE: " + (encontradaHojaIE ? "sí" : "NO") + " -> " + (ok ? "✅ CORRECTO" : "⚠ FALLÓ LA SINCRONIZACIÓN");
  Logger.log(resumen);
  return { ok: ok, resumen: resumen, idForo: idForo, urlDocumentoAnalisis: ss.getUrl() };
}


/*****************************************************
 * GENERADORES DE DATOS AL AZAR — usados por
 * probarEnvioCompletoAleatorio() y simular50RespuestasFEM()
 *****************************************************/
function generarCamposAleatoriosFEM_(){
  const frasesLargas = [
    "El Foro permitió identificar avances importantes en el trabajo colaborativo entre docentes y directivos, con propuestas concretas para el siguiente año lectivo.",
    "Se evidenció una participación activa de estudiantes y familias, aunque persisten retos en la articulación entre sedes y jornadas.",
    "Las mesas de trabajo generaron acuerdos sobre estrategias pedagógicas y de convivencia que se plasmarán en el plan de mejoramiento institucional.",
    "La comunidad educativa valoró positivamente el espacio de reflexión, señalando la necesidad de darle continuidad durante el año.",
    "Se identificaron fortalezas en el uso de recursos tecnológicos y oportunidades de mejora en la atención a la diversidad de los estudiantes."
  ];
  const frase = function(){ return frasesLargas[Math.floor(Math.random() * frasesLargas.length)] + " (dato de prueba generado automáticamente, " + new Date().toISOString() + ")"; };
  const campos = {};
  const t = function(id, valor){ campos[id] = { tipo:"text", valor: valor }; };
  ["direccion", "zona", "comuna", "grupo"].forEach(function(id){ t(id, "Dato de prueba " + id + " " + Math.floor(Math.random() * 1000)); });
  t("rector", "Rector(a) de prueba " + Math.floor(Math.random() * 1000));
  ["respuestaSesion1", "respuestaSesion1Pregunta2", "respuestaSesion2Pregunta1", "respuestaSesion2Pregunta3", "respuestaSesion2Pregunta4", "respuestaSesion2Pregunta5", "respuestaSesion3Pregunta1", "respuestaSesion3Pregunta3", "respuestaSesion3Pregunta4"]
    .forEach(function(id){ t(id, frase()); });
  [1, 2, 3, 4, 5].forEach(function(n){ t("respuestaSesion2Pregunta2Accion" + n, "Acción de prueba " + n + ": " + frase()); t("respuestaSesion3Pregunta2Accion" + n, "Acción de prueba " + n + ": " + frase()); });
  ["Rector", "Coordinador", "Docentes", "TutorPTA", "Orientador", "Estudiantes", "Padres", "Administrativos", "Egresados", "Sector", "Otros"]
    .forEach(function(id){ t("participantes" + id, String(Math.floor(Math.random() * 20))); });
  return campos;
}

function generarAsistenteAleatorioQR_(){
  const nombres = ["Ana", "Carlos", "María", "Luis", "Sofía", "Andrés", "Valentina", "Jorge", "Camila", "Diego"];
  const apellidos = ["Pérez", "Gómez", "Rodríguez", "Martínez", "López", "García", "Torres", "Ramírez", "Vargas", "Castro"];
  const cargo = CARGOS_ASISTENCIA_QR[Math.floor(Math.random() * CARGOS_ASISTENCIA_QR.length)];
  const requiereCondicion = CARGOS_SIN_CONDICION_QR.indexOf(cargo) === -1;
  const fortalezas = [];
  while(fortalezas.length < 1 + Math.floor(Math.random() * 3)){
    const f = FORTALEZAS_ASISTENCIA_QR[Math.floor(Math.random() * FORTALEZAS_ASISTENCIA_QR.length)];
    if(fortalezas.indexOf(f) === -1) fortalezas.push(f);
  }
  const dificultades = [];
  while(dificultades.length < 1 + Math.floor(Math.random() * 3)){
    const d = DIFICULTADES_ASISTENCIA_QR[Math.floor(Math.random() * DIFICULTADES_ASISTENCIA_QR.length)];
    if(dificultades.indexOf(d) === -1) dificultades.push(d);
  }
  const documento = String(1000000000 + Math.floor(Math.random() * 899999999));
  return {
    nombre: nombres[Math.floor(Math.random() * nombres.length)] + " " + apellidos[Math.floor(Math.random() * apellidos.length)],
    sexo: SEXOS_ASISTENCIA_QR[Math.floor(Math.random() * SEXOS_ASISTENCIA_QR.length)],
    edad: String(15 + Math.floor(Math.random() * 50)),
    tipoAsistencia: "Presencial",
    cargo: cargo,
    rolForo: ROLES_FORO_QR[Math.floor(Math.random() * ROLES_FORO_QR.length)],
    jornada: requiereCondicion ? JORNADAS_ASISTENCIA_QR[Math.floor(Math.random() * JORNADAS_ASISTENCIA_QR.length)] : "",
    sede: requiereCondicion ? "Sede de prueba " + (1 + Math.floor(Math.random() * 3)) : "",
    fortalezas: fortalezas,
    fortalezaOtro: "",
    dificultades: dificultades,
    dificultadOtro: "",
    documento: documento,
    correo: "asistente.prueba" + documento + "@ejemplo.com",
    telefono: String(3000000000 + Math.floor(Math.random() * 99999999))
  };
}

function generarValoracionAleatoriaFEM_(){
  const p = function(){ return 1 + Math.floor(Math.random() * 5); };
  const p1 = p(), p2 = p(), p3 = p(), p4 = p();
  const mejora = function(n){ return n <= 2 ? "Sugerencia de mejora de prueba generada automáticamente." : ""; };
  return {
    p1: p1, p2: p2, p3: p3, p4: p4,
    mejoraP1: mejora(p1), mejoraP2: mejora(p2), mejoraP3: mejora(p3), mejoraP4: mejora(p4),
    p5: "Sugerencia final de prueba generada automáticamente para fortalecer el FEM 2027."
  };
}


/*****************************************************
 * PRUEBA DE ENVÍO COMPLETO CON DATOS AL AZAR
 *
 * Simula, de principio a fin, una sola IE de prueba completando el
 * Foro: caracterización y 3 sesiones (texto y números al azar),
 * varias firmas de asistencia por QR (datos al azar), envío
 * definitivo, generación y envío del informe, y valoración final
 * (corazones y comentarios al azar) — sin necesidad de abrir el
 * formulario en el navegador. Usa una IE de prueba ya existente (por
 * defecto "IE PRUEBA 1234"), así que no envía correos a nadie fuera
 * de las 10 IE de prueba ya configuradas.
 *
 * Ejecutar manualmente:  probarEnvioCompletoAleatorio("IE PRUEBA 1234")
 *****************************************************/
function probarEnvioCompletoAleatorio(nombreIE){
  nombreIE = nombreIE || "IE PRUEBA 1234";
  const resultado = { ie: nombreIE, pasos: {}, errores: [] };

  try{
    const hoja = asegurarColumnasAccesosIE_();
    const mapa = mapaHoja_(hoja);
    if(hoja.getLastRow() < 2) throw new Error("AccesosIE no tiene filas.");
    const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
    const fila = valores.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIE);
    if(!fila) throw new Error("No existe " + nombreIE + " en AccesosIE.");
    const idForo = String(fila[mapa.ID_FORO - 1] || "").trim();
    const dispositivoId = "PRUEBA-ALEATORIA-" + idForo.slice(0, 8);

    const datos = { idForo: idForo, campos: generarCamposAleatoriosFEM_() };
    guardarAvanceForo(datos);
    resultado.pasos.caracterizacionYSesiones = true;

    const cantidadAsistentes = 3 + Math.floor(Math.random() * 4);
    let asistentesFirmados = 0;
    for(let i = 0; i < cantidadAsistentes; i++){
      const asistente = generarAsistenteAleatorioQR_();
      const registro = registrarAsistenciaQR(idForo, asistente.nombre, asistente.sexo, asistente.edad, asistente.tipoAsistencia, asistente.cargo, asistente.rolForo, asistente.jornada, asistente.sede, asistente.fortalezas, asistente.fortalezaOtro, asistente.dificultades, asistente.dificultadOtro, asistente.documento, asistente.correo, asistente.telefono, true, dispositivoId + "-QR-" + i);
      if(registro.ok) asistentesFirmados++;
    }
    resultado.pasos.asistenciaQR = asistentesFirmados + " de " + cantidadAsistentes;

    const sesion = reclamarSesionCodigo_("", "", dispositivoId, idForo, true);
    if(!sesion.ok) throw new Error("No fue posible reclamar la sesión: " + sesion.mensaje);
    const datosGuardados = obtenerDatosGuardadosPorIdForo_(idForo);
    datosGuardados.idForo = idForo;
    const envio = enviarForoDefinitivo(idForo, sesion.tokenSesion, dispositivoId, datosGuardados);
    if(!envio || (!envio.ok && !envio.yaEnviado)){ liberarSesionCodigo_("", "", dispositivoId, sesion.tokenSesion, idForo); throw new Error((envio && envio.mensaje) || "Envío definitivo falló."); }
    resultado.pasos.envioDefinitivo = true;

    const informe = generarInformeFEM(idForo, datosGuardados);
    if(!informe || !informe.ok) throw new Error((informe && informe.mensaje) || "No fue posible generar el informe.");
    resultado.pasos.informe = true;
    enviarInformeFEM(idForo, datosGuardados, informe.pdfId);
    resultado.pasos.correoInforme = true;

    const valoracionAleatoria = generarValoracionAleatoriaFEM_();
    const valoracion = guardarValoracionFEM(idForo, valoracionAleatoria);
    if(!valoracion || !valoracion.ok) throw new Error((valoracion && valoracion.mensaje) || "No fue posible guardar la valoración.");
    resultado.pasos.valoracion = valoracionAleatoria;
    enviarComprobanteParticipacionFEM(idForo, datosGuardados);
    resultado.pasos.correoComprobante = true;

    liberarSesionCodigo_("", "", dispositivoId, sesion.tokenSesion, idForo);

    const ss = obtenerSpreadsheetAnalisisFEM_();
    const shIE = ss.getSheetByName(nombreHojaIE_(nombreIE));
    resultado.pasos.reflejadoEnAnalisis = !!(shIE && shIE.getLastRow() >= 2);
    resultado.urlDocumentoAnalisis = ss.getUrl();

  }catch(error){
    resultado.errores.push(error.message);
  }

  Logger.log("RESULTADO PRUEBA ALEATORIA — " + nombreIE + ":\n" + JSON.stringify(resultado, null, 2));
  return resultado;
}


/*****************************************************
 * VERIFICAR PERMISOS Y CUOTA DE ENVÍO DE CORREO — FEM 2026
 *
 * Confirma que la cuenta que ejecuta el script puede enviar como
 * REMITENTE_FEM (calidadeducacion@alcaldianeiva.gov.co) y cuántos
 * correos quedan disponibles hoy en la cuota diaria — importante
 * antes de cualquier envío masivo (accesos, avisos, informes).
 *
 * Ejecutar manualmente desde el editor de Apps Script y revisar el
 * log.
 *****************************************************/
function verificarPermisosEnvioCorreoFEM(){
  const cuenta = Session.getEffectiveUser().getEmail();
  const aliases = GmailApp.getAliases().map(function(a){ return a.toLowerCase(); });
  const puedeEnviarComoFEM = cuenta.toLowerCase() === REMITENTE_FEM || aliases.indexOf(REMITENTE_FEM) !== -1;
  const cuotaRestante = MailApp.getRemainingDailyQuota();

  const resumen = [
    "Cuenta que ejecuta el script: " + cuenta,
    "Aliases de envío disponibles: " + (aliases.join(", ") || "(ninguno)"),
    "¿Puede enviar como " + REMITENTE_FEM + "?: " + (puedeEnviarComoFEM ? "SÍ" : "NO — configure el alias en Gmail antes de enviar."),
    "Cuota de correos restante hoy: " + cuotaRestante
  ];
  Logger.log(resumen.join("\n"));
  return { ok:true, cuenta: cuenta, aliases: aliases, puedeEnviarComoFEM: puedeEnviarComoFEM, cuotaRestante: cuotaRestante, resumen: resumen };
}


/*****************************************************
 * PROGRAMAR AVISO AL EQUIPO DE CALIDAD EDUCATIVA — 6:30 A.M.
 *
 * Crea un disparador de una sola vez para las 6:30 a.m. (hora de
 * Bogotá) de hoy, o de mañana si ya pasaron las 6:30 a.m., que
 * ejecuta enviarAvisoEquipoCalidadFEM_(): envía, a cada una de las 10
 * IE de prueba (correos del equipo de calidad educativa), un correo
 * con el mismo diseño de código + enlace ya usado en
 * enviarAccesoIndividualIEPrueba_(), agregando arriba un aviso breve
 * explicando que el correo llegó programado a las 6:30 a.m. desde la
 * aplicación FEM 2026 y pidiendo verificar cómo se comporta la página
 * y reportar cualquier novedad durante el día.
 *
 * El disparador se autodestruye la primera vez que se ejecuta.
 *
 * Ejecutar manualmente:  programarAvisoEquipoCalidadFEM()
 *****************************************************/
function programarAvisoEquipoCalidadFEM(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === "enviarAvisoEquipoCalidadFEM_") ScriptApp.deleteTrigger(t);
  });

  const zona = "America/Bogota";
  const ahora = new Date();
  let fechaEnvio = new Date(Utilities.formatDate(ahora, zona, "yyyy-MM-dd") + "T06:30:00");
  const horaActual = Number(Utilities.formatDate(ahora, zona, "HH"));
  const minutoActual = Number(Utilities.formatDate(ahora, zona, "mm"));
  if(horaActual > 6 || (horaActual === 6 && minutoActual >= 30)){
    fechaEnvio = new Date(fechaEnvio.getTime() + 24 * 60 * 60 * 1000);
  }

  ScriptApp.newTrigger("enviarAvisoEquipoCalidadFEM_").timeBased().at(fechaEnvio).create();

  const mensaje = "Aviso programado para las 6:30 a.m. (hora de Bogotá) del " + Utilities.formatDate(fechaEnvio, zona, "dd/MM/yyyy") + ".";
  Logger.log(mensaje);
  return { ok:true, fechaEnvio: fechaEnvio.toISOString(), mensaje: mensaje };
}

function enviarAvisoEquipoCalidadFEM_(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === "enviarAvisoEquipoCalidadFEM_") ScriptApp.deleteTrigger(t);
  });

  const cuenta = Session.getEffectiveUser().getEmail().toLowerCase();
  const aliases = GmailApp.getAliases().map(function(a){ return a.toLowerCase(); });
  if(cuenta !== REMITENTE_FEM && aliases.indexOf(REMITENTE_FEM) === -1){
    Logger.log("No se pudo enviar el aviso: la cuenta no puede enviar como " + REMITENTE_FEM + ".");
    return;
  }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2){ Logger.log("AccesosIE no tiene filas."); return; }
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();

  const nombresDestino = ["IE PRUEBA 1234"].concat(IES_PRUEBA_ADICIONALES.map(function(x){ return x.ie; }));
  const resultados = [];

  nombresDestino.forEach(function(nombreIE){
    try{
      const fila = valores.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIE);
      if(!fila) throw new Error("No existe en AccesosIE.");
      const correoIE = String(fila[mapa.EMAIL_IE - 1] || "").trim();
      const codigo = String(fila[mapa.CODIGO_ACCESO - 1] || "").trim();
      const url = String(fila[mapa.URL_ACCESO - 1] || "").trim();
      if(!correoIE) throw new Error("Sin EMAIL_IE.");

      const ieSinPrefijo = nombreIESinPrefijoInstitucional_(nombreIE);
      const asunto = "🧪 Aviso de prueba programada 6:30 a.m. — Foro Educativo Institucional FEM 2026";
      const textoEnlace = "Ingreso de prueba al Foro Educativo Institucional";

      const cuerpoTexto =
        "Secretaría de Educación de Neiva\n\n" +
        "Este correo fue programado para llegar hoy a las 6:30 a.m. desde la aplicación FEM 2026.\n\n" +
        "Por favor verifiquen cómo se comporta la página con el código y el enlace de abajo, y reporten cualquier novedad durante el día.\n\n" +
        "Institución de prueba: " + ieSinPrefijo + "\n\n" +
        "Código de acceso: " + codigo + "\n\n" +
        textoEnlace + ":\n" + url + "\n\n" +
        "Secretaría de Educación de Neiva\n" +
        "Foro Educativo Institucional – Neiva 2026\n" +
        "“Escuela Viva: Voces que construyen territorio”";

      const cuerpoHTML =
        "<div style=\"background:#F7F8FA;padding:28px 12px;font-family:Arial,Helvetica,sans-serif;\">" +
        "<div style=\"max-width:520px;margin:0 auto;background:#FFFFFF;border-radius:16px;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,.10);\">" +
        "<div style=\"background:#0B6A44;padding:26px 28px;text-align:center;\">" +
        "<div style=\"color:#FFFFFF;font-size:20px;font-weight:700;\">Foro Educativo Institucional</div>" +
        "<div style=\"color:#CFE8DC;font-size:14px;margin-top:2px;\">Neiva 2026 — Prueba</div>" +
        "</div>" +
        "<div style=\"padding:28px 28px 8px;\">" +
        "<div style=\"background:#FFF8E1;border-left:6px solid #F4B400;border-radius:10px;padding:16px 20px;margin:0 0 22px;\">" +
        "<p style=\"font-size:14px;color:#7A5B00;margin:0;\"><strong>🧪 Aviso de prueba:</strong> este correo fue programado para llegar hoy a las 6:30 a.m. desde la aplicación FEM 2026. Por favor verifiquen cómo se comporta la página con el código y el enlace de abajo, y reporten cualquier novedad durante el día.</p>" +
        "</div>" +
        "<p style=\"font-size:15px;color:#4A4A4A;line-height:1.6;margin:0 0 22px;\">Institución de prueba: <strong>" + ieSinPrefijo + "</strong></p>" +
        "<div style=\"background:#F7F8FA;border-left:6px solid #F4B400;border-radius:10px;padding:16px 20px;margin:0 0 24px;text-align:center;\">" +
        "<div style=\"font-size:12px;font-weight:700;color:#0B6A44;text-transform:uppercase;letter-spacing:.5px;\">Código de acceso</div>" +
        "<div style=\"font-size:30px;font-weight:700;letter-spacing:6px;color:#0B6A44;margin-top:4px;\">" + codigo + "</div>" +
        "</div>" +
        "<div style=\"text-align:center;margin:0 0 24px;\">" +
        "<a href=\"" + url + "\" target=\"_blank\" style=\"display:inline-block;background:#0B6A44;color:#FFFFFF;text-decoration:none;font-weight:700;font-size:15px;padding:14px 26px;border-radius:10px;\">" + textoEnlace + "</a>" +
        "</div>" +
        "</div>" +
        "<div style=\"background:#F7F8FA;padding:18px 28px;text-align:center;border-top:1px solid #E5E7EA;\">" +
        "<p style=\"font-size:13px;color:#0B6A44;font-weight:700;margin:0;\">Secretaría de Educación de Neiva</p>" +
        "<p style=\"font-size:12px;color:#888888;margin:4px 0 0;font-style:italic;\">“Escuela Viva: Voces que construyen territorio”</p>" +
        "</div>" +
        "</div>" +
        "</div>";

      const opciones = { htmlBody: cuerpoHTML, name: "Secretaría de Educación de Neiva", replyTo: REMITENTE_FEM };
      if(cuenta !== REMITENTE_FEM) opciones.from = REMITENTE_FEM;

      GmailApp.sendEmail(correoIE, asunto, cuerpoTexto, opciones);
      resultados.push(nombreIE + ": ✅ enviado a " + correoIE);
    }catch(error){
      resultados.push(nombreIE + ": ⚠ " + error.message);
    }
  });

  Logger.log("AVISO 6:30 A.M. — RESULTADO:\n" + resultados.join("\n"));
}


/*****************************************************
 * DEBUG COMPLETO — FEM 2026
 *
 * Corre en cadena varias pruebas ya existentes más las nuevas de este
 * archivo y arma un solo reporte pasa/falla en el log. Pensado para
 * ejecutarse manualmente después de un reset o antes de un evento
 * real, sin tener que ejecutar función por función.
 *
 * Ejecutar manualmente:  ejecutarDebugCompletoFEM()
 *****************************************************/
function ejecutarDebugCompletoFEM(){
  const reporte = [];
  function correr(nombre, fn){
    try{
      const r = fn();
      const ok = !r || r.ok === true || r.ok === undefined;
      reporte.push((ok ? "✅ " : "⚠ ") + nombre + (r && r.mensaje ? " — " + r.mensaje : ""));
    }catch(error){
      reporte.push("❌ " + nombre + " — ERROR: " + error.message);
    }
  }

  correr("Catálogo de instituciones (probarCatalogoIE)", probarCatalogoIE);
  correr("Permisos y cuota de correo (verificarPermisosEnvioCorreoFEM)", verificarPermisosEnvioCorreoFEM);
  correr("Validación de código correcto (probarValidacion1234)", probarValidacion1234);
  correr("Validación de código incorrecto (probarCodigoIncorrecto1234)", probarCodigoIncorrecto1234);
  correr("Guardar avance de foro (probarGuardarAvanceForo)", probarGuardarAvanceForo);
  correr("Autoguardado con ID_FORO inválido", probarAutoguardadoConIdForoInvalido);
  correr("Autoguardado sin datos", probarAutoguardadoSinDatos);
  correr("Reintento por falla de guardado local", function(){ return probarReintentoPorFallaGuardadoLocal("IE PRUEBA 1234"); });
  correr("Flujo plenaria -> documento de análisis", probarFlujoPlenariaHastaDocumentoAnalisis);

  Logger.log("========================================");
  Logger.log("DEBUG COMPLETO FEM 2026");
  Logger.log(reporte.join("\n"));
  Logger.log("========================================");
  return { ok:true, reporte: reporte };
}


/*****************************************************
 * SIMULAR 50 RESPUESTAS COMPLETAS — FEM 2026
 *
 * Crea 50 IE de prueba ("IE Simulación 01".."50", TIPO=PRUEBA, con
 * correo en el dominio reservado .invalid — nunca entregable, para no
 * arriesgar ningún envío real) y para cada una: guarda caracterización
 * y 3 sesiones con texto y números al azar, registra entre 3 y 8
 * firmas de asistencia QR al azar, hace el envío definitivo y guarda
 * una valoración al azar — todo reflejado automáticamente en el
 * documento de análisis.
 *
 * A propósito NO genera el informe (Doc + PDF) ni envía ningún correo
 * para las 50: generar 50 informes reales superaría los 6 minutos que
 * permite una ejecución manual desde el editor, y emitiría 50 correos
 * innecesarios. Para ver un informe real completo, usar
 * probarEnvioCompletoAleatorio() sobre una sola IE de prueba de
 * verdad.
 *
 * Devuelve (y deja en el log) el resultado de cada una de las 50 IE.
 *
 * Ejecutar manualmente:  simular50RespuestasFEM()
 *****************************************************/
function simular50RespuestasFEM(){
  const TOTAL = 50;
  const resultados = [];

  for(let i = 1; i <= TOTAL; i++){
    const numero = String(i).padStart(2, "0");
    const nombreIE = "IE Simulación " + numero;
    const fila = { ie: nombreIE, ok:false, asistentes:0, valoracionNota:0, errores:[] };

    try{
      const hoja = asegurarColumnasAccesosIE_();
      const mapa = mapaHoja_(hoja);
      const valoresActuales = hoja.getLastRow() >= 2 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues() : [];
      const filaExistente = valoresActuales.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIE);
      let idForo;
      if(filaExistente){
        idForo = String(filaExistente[mapa.ID_FORO - 1] || "").trim();
      }else{
        idForo = Utilities.getUuid();
        const nuevaFila = new Array(hoja.getLastColumn()).fill("");
        const set = function(col, valor){ if(mapa[col]) nuevaFila[mapa[col] - 1] = valor; };
        set("ID_ACCESO", Utilities.getUuid());
        set("IE", nombreIE);
        set("DANE", "SIMULACION-" + numero);
        set("CODIGO_ACCESO", generarCodigoAcceso_());
        set("TOKEN", Utilities.getUuid().replace(/-/g, ""));
        set("URL_ACCESO", URL_WEBAPP_PRODUCCION + "?t=SIMULACION" + numero);
        set("ID_FORO", idForo);
        set("ESTADO", "DISPONIBLE");
        set("EMAIL_IE", "simulacion" + numero + "@fem2026.invalid");
        set("EMAIL_RESPONSABLE", "simulacion" + numero + "@fem2026.invalid");
        set("TIPO", "PRUEBA");
        set("FECHA_GENERACION", new Date());
        hoja.appendRow(nuevaFila);
      }

      guardarAvanceForo({ idForo: idForo, campos: generarCamposAleatoriosFEM_() });

      const dispositivoId = "SIMULACION-" + numero;
      const cantidadAsistentes = 3 + Math.floor(Math.random() * 6);
      let asistentesFirmados = 0;
      for(let a = 0; a < cantidadAsistentes; a++){
        const asistente = generarAsistenteAleatorioQR_();
        const registro = registrarAsistenciaQR(idForo, asistente.nombre, asistente.sexo, asistente.edad, asistente.tipoAsistencia, asistente.cargo, asistente.rolForo, asistente.jornada, asistente.sede, asistente.fortalezas, asistente.fortalezaOtro, asistente.dificultades, asistente.dificultadOtro, asistente.documento, asistente.correo, asistente.telefono, true, dispositivoId + "-QR-" + a);
        if(registro.ok) asistentesFirmados++;
      }
      fila.asistentes = asistentesFirmados;

      const sesion = reclamarSesionCodigo_("", "", dispositivoId, idForo, true);
      if(!sesion.ok) throw new Error("Sesión: " + sesion.mensaje);
      const datosGuardados = obtenerDatosGuardadosPorIdForo_(idForo);
      datosGuardados.idForo = idForo; datosGuardados.institucion = nombreIE;
      const envio = enviarForoDefinitivo(idForo, sesion.tokenSesion, dispositivoId, datosGuardados);
      liberarSesionCodigo_("", "", dispositivoId, sesion.tokenSesion, idForo);
      if(!envio || (!envio.ok && !envio.yaEnviado)) throw new Error((envio && envio.mensaje) || "Envío definitivo falló.");

      const valoracionAleatoria = generarValoracionAleatoriaFEM_();
      const valoracion = guardarValoracionFEM(idForo, valoracionAleatoria);
      if(!valoracion || !valoracion.ok) throw new Error((valoracion && valoracion.mensaje) || "Valoración falló.");
      fila.valoracionNota = ((valoracionAleatoria.p1 + valoracionAleatoria.p2 + valoracionAleatoria.p3 + valoracionAleatoria.p4) / 4).toFixed(1);

      fila.ok = true;

    }catch(error){
      fila.errores.push(error.message);
    }

    resultados.push(fila);
  }

  const exitosas = resultados.filter(function(r){ return r.ok; }).length;
  Logger.log("========================================");
  Logger.log("SIMULACIÓN DE 50 RESPUESTAS — RESUMEN");
  Logger.log("Exitosas: " + exitosas + " de " + TOTAL);
  resultados.forEach(function(r){
    Logger.log(r.ie + " -> " + (r.ok ? ("✅ asistentes:" + r.asistentes + " nota:" + r.valoracionNota) : ("❌ " + r.errores.join(" | "))));
  });
  Logger.log("========================================");

  return { ok:true, exitosas: exitosas, total: TOTAL, resultados: resultados };
}


/*****************************************************
 * BLOQUEO POR HORARIO — página de bloqueo antes de una hora
 *
 * Pone en AccesosIE, en la columna HABILITAR_DESDE de la fila de la
 * IE indicada, una fecha/hora (hoy, hora de Bogotá). Mientras esa
 * hora no llegue, validarAccesoIE() rechaza el ingreso con el código
 * BLOQUEADO_POR_HORARIO, que el cliente muestra como una página de
 * bloqueo completa (pantallaBloqueoHorarioFEM en Index.html) en vez
 * del formulario — el código de acceso sigue siendo el mismo, no hay
 * que reenviar nada ni cambiar nada más.
 *
 * Al llegar la hora programada, el siguiente intento de ingreso ya
 * entra normalmente: no hace falta ninguna acción manual para
 * "abrir" el acceso.
 *
 * Ejecutar manualmente:
 *   programarBloqueoHorarioIE("IE Prueba Ana", 9, 0)
 *   quitarBloqueoHorarioIE("IE Prueba Ana")            (para levantarlo antes de tiempo)
 *   bloquearPruebaAnaYNelsonHasta9am()                  (las dos de una vez)
 *****************************************************/
function programarBloqueoHorarioIE(nombreIE, hora, minuto){
  function salir(resultado){ Logger.log(resultado.mensaje); return resultado; }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(!mapa.HABILITAR_DESDE) return salir({ ok:false, mensaje:"No fue posible crear la columna HABILITAR_DESDE." });
  if(hoja.getLastRow() < 2) return salir({ ok:false, mensaje:"AccesosIE no tiene filas." });

  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const indice = valores.findIndex(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIE; });
  if(indice === -1) return salir({ ok:false, mensaje:"No existe " + nombreIE + " en AccesosIE." });

  const zona = "America/Bogota";
  const ahora = new Date();
  const fechaHabilitacion = new Date(
    Utilities.formatDate(ahora, zona, "yyyy-MM-dd") + "T" +
    String(hora).padStart(2, "0") + ":" + String(minuto).padStart(2, "0") + ":00"
  );

  hoja.getRange(indice + 2, mapa.HABILITAR_DESDE).setValue(fechaHabilitacion);

  const mensaje = nombreIE + ": bloqueada hasta las " + Utilities.formatDate(fechaHabilitacion, zona, "h:mm a") + " del " + Utilities.formatDate(fechaHabilitacion, zona, "dd/MM/yyyy") + ".";
  Logger.log(mensaje);
  return { ok:true, mensaje: mensaje, fechaHabilitacion: fechaHabilitacion.toISOString() };
}

function quitarBloqueoHorarioIE(nombreIE){
  function salir(resultado){ Logger.log(resultado.mensaje); return resultado; }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(!mapa.HABILITAR_DESDE || hoja.getLastRow() < 2) return salir({ ok:true, mensaje:"Nada que quitar." });
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const indice = valores.findIndex(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIE; });
  if(indice === -1) return salir({ ok:false, mensaje:"No existe " + nombreIE + " en AccesosIE." });
  hoja.getRange(indice + 2, mapa.HABILITAR_DESDE).setValue("");
  const mensaje = nombreIE + ": bloqueo por horario retirado.";
  Logger.log(mensaje);
  return { ok:true, mensaje: mensaje };
}

function bloquearPruebaAnaYNelsonHasta9am(){
  const resultadoAna = programarBloqueoHorarioIE("IE Prueba Ana", 9, 0);
  const resultadoNelson = programarBloqueoHorarioIE("IE Prueba Nelson", 9, 0);
  Logger.log([resultadoAna.mensaje, resultadoNelson.mensaje].join("\n"));
  return { ok:true, ana: resultadoAna, nelson: resultadoNelson };
}


/*****************************************************
 * LINK DEL DOCUMENTO DE ANÁLISIS
 *
 * Devuelve (y deja en el log) la URL del documento de análisis
 * separado (Análisis FEM 2026). Si todavía no existe (nadie lo ha
 * necesitado antes), lo crea vacío en ese mismo momento — no hace
 * falta esperar a una reconstrucción completa solo para obtener el
 * link.
 *
 * Ejecutar manualmente:  obtenerLinkDocumentoAnalisisFEM()
 *****************************************************/
function obtenerLinkDocumentoAnalisisFEM(){
  const ss = obtenerSpreadsheetAnalisisFEM_();
  Logger.log(ss.getUrl());
  return { ok:true, url: ss.getUrl() };
}


/*****************************************************
 * DIAGNOSTICAR UN CÓDIGO/IE DE PRUEBA
 *
 * Muestra, tal como está HOY en AccesosIE, el estado completo de la
 * fila de una IE (por defecto "IE PRUEBA 1234"): si existe o no, su
 * CODIGO_ACCESO, TOKEN, URL_ACCESO, ESTADO, TIPO y HABILITAR_DESDE.
 * Pensado para responder rápido a "no me sirve el código de prueba"
 * sin adivinar — dice exactamente qué hay (o no hay) en la hoja.
 *
 * Ejecutar manualmente:  diagnosticarAccesoPruebaFEM("IE PRUEBA 1234")
 *****************************************************/
function diagnosticarAccesoPruebaFEM(nombreIE){
  nombreIE = nombreIE || "IE PRUEBA 1234";
  const hoja = abrirSpreadsheet_().getSheetByName(HOJA_ACCESOS);
  if(!hoja) { Logger.log("No existe la hoja " + HOJA_ACCESOS + "."); return { ok:false, mensaje:"No existe " + HOJA_ACCESOS + "." }; }
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2){
    const mensaje = HOJA_ACCESOS + " no tiene filas — probablemente por un reset. Hay que volver a ejecutar crearAccesoPrueba1234() / crearIEsPruebaAdicionales() / generarAccesosIE().";
    Logger.log(mensaje);
    return { ok:false, existe:false, mensaje: mensaje };
  }

  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const indice = valores.findIndex(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIE; });

  if(indice === -1){
    const mensaje = "No existe la fila \"" + nombreIE + "\" en " + HOJA_ACCESOS + " — probablemente por un reset. Hay que volver a crearla (crearAccesoPrueba1234() para IE PRUEBA 1234, o crearIEsPruebaAdicionales() para las otras 9).";
    Logger.log(mensaje);
    return { ok:false, existe:false, mensaje: mensaje };
  }

  const fila = valores[indice];
  const val = function(col){ return mapa[col] ? String(fila[mapa[col]-1] || "") : "(sin columna)"; };
  const habilitarDesdeCelda = mapa.HABILITAR_DESDE ? hoja.getRange(indice + 2, mapa.HABILITAR_DESDE).getValue() : "";
  const bloqueadaPorHorario = habilitarDesdeCelda instanceof Date && !isNaN(habilitarDesdeCelda.getTime()) && new Date() < habilitarDesdeCelda;

  const resumen = [
    "IE: " + val("IE"),
    "CODIGO_ACCESO: " + val("CODIGO_ACCESO"),
    "TOKEN: " + val("TOKEN"),
    "URL_ACCESO: " + val("URL_ACCESO"),
    "ID_FORO: " + val("ID_FORO"),
    "ESTADO: " + val("ESTADO"),
    "TIPO: " + val("TIPO"),
    "HABILITAR_DESDE: " + (habilitarDesdeCelda ? Utilities.formatDate(habilitarDesdeCelda, "America/Bogota", "dd/MM/yyyy HH:mm") : "(vacío)") + (bloqueadaPorHorario ? " — TODAVÍA BLOQUEADA por horario" : "")
  ];
  Logger.log(resumen.join("\n"));
  return { ok:true, existe:true, resumen: resumen, bloqueadaPorHorario: bloqueadaPorHorario };
}


/*****************************************************
 * CREAR TODOS LOS LINKS DE PRUEBA DE UNA SOLA VEZ
 *
 * Ejecuta, en orden, crearAccesoPrueba1234() y
 * crearIEsPruebaAdicionales() (ambas se saltan solas la IE que ya
 * exista, nunca duplican), y al final imprime en el log la lista
 * completa de las 10 IE de prueba con su código y su enlace —
 * pensada para usarse justo después de un reset, cuando hace falta
 * volver a tener todos los accesos de prueba listos para copiar.
 *
 * Ejecutar manualmente:  crearTodosLosAccesosDePruebaFEM()
 *****************************************************/
function crearTodosLosAccesosDePruebaFEM(){
  crearAccesoPrueba1234();
  crearIEsPruebaAdicionales();

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  const valores = hoja.getLastRow() >= 2 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues() : [];

  const nombresPrueba = ["IE PRUEBA 1234"].concat(IES_PRUEBA_ADICIONALES.map(function(x){ return x.ie; }));
  const lista = nombresPrueba.map(function(nombreIE){
    const fila = valores.find(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIE; });
    if(!fila) return nombreIE + ": ⚠ no se pudo crear/encontrar.";
    return nombreIE + " -> código: " + String(fila[mapa.CODIGO_ACCESO - 1] || "") + " | " + String(fila[mapa.URL_ACCESO - 1] || "");
  });

  Logger.log("========================================");
  Logger.log("LINKS DE LAS 10 IE DE PRUEBA");
  Logger.log(lista.join("\n"));
  Logger.log("========================================");

  return { ok:true, lista: lista };
}


/*****************************************************
 * LIBERAR EL CANDADO DE SESIÓN DE UNA IE
 *
 * El código de acceso puede estar perfectamente bien en AccesosIE
 * (como confirmó diagnosticarAccesoPruebaFEM) y aun así no dejar
 * entrar: reclamarSesionCodigo_() solo permite UN dispositivo
 * conectado por IE (guarda el candado en ScriptProperties, con la
 * clave FEM_SESION_FORO_<idForo codificado>). Si alguna prueba
 * automática anterior (probarEnvioCompletoAleatorio,
 * probarFlujoPlenariaHastaDocumentoAnalisis,
 * ejecutarDebugCompletoFEM, etc.) reclamó esa sesión con un
 * "dispositivo" simulado y se interrumpió antes de liberarla, el
 * candado queda tomado por ese dispositivo simulado — y el
 * navegador real, al entrar con el código, recibe
 * SESION_YA_ABIERTA en vez de pasar. Esto libera ese candado sin
 * tocar ningún dato de AccesosIE, AvancesForo ni nada más.
 *
 * Ejecutar manualmente:  liberarCandadoSesionIE("IE PRUEBA 1234")
 *****************************************************/
function liberarCandadoSesionIE(nombreIE){
  function salir(resultado){ Logger.log(resultado.mensaje); return resultado; }

  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2) return salir({ ok:false, mensaje:"AccesosIE no tiene filas." });
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIE; });
  if(!fila) return salir({ ok:false, mensaje:"No existe " + nombreIE + " en AccesosIE." });
  const idForo = String(fila[mapa.ID_FORO - 1] || "").trim();
  if(!idForo) return salir({ ok:false, mensaje:nombreIE + " no tiene ID_FORO." });

  const clave = obtenerClaveSesionCodigo_("", "", idForo);
  const props = PropertiesService.getScriptProperties();
  const habia = props.getProperty(clave);
  props.deleteProperty(clave);

  const mensaje = nombreIE + " (ID_FORO " + idForo + "): " + (habia ? "candado de sesión liberado (estaba tomado por: " + habia + ")." : "no tenía ningún candado activo — el código ya estaba libre para entrar.");
  Logger.log(mensaje);
  return { ok:true, habiaCandado: !!habia, mensaje: mensaje };
}


/*****************************************************
 * PRUEBA: HASTA 4 DISPOSITIVOS SIMULTÁNEOS + FUSIÓN DE CAMPOS
 *
 * Simula 4 "dispositivos" reclamando sesión con el mismo código a la
 * vez (el máximo permitido — ver MAX_SESIONES_SIMULTANEAS_IE),
 * confirma que un 5° es rechazado sin forzar y aceptado forzando
 * (desalojando al de menor actividad reciente), y comprueba que dos
 * dispositivos guardando avances con SOLO su propio campo lleno (el
 * resto vacío, como pasaría con el DOM real de cada navegador) NO se
 * borran el trabajo entre sí — la fusión de campos en
 * guardarAvanceForo() debe conservar ambos.
 *
 * Dentro de una sola ejecución de Apps Script las llamadas son
 * secuenciales (no hay paralelismo real de hilos), pero eso es
 * exactamente lo que importa aquí: valida la LÓGICA de cupos y
 * fusión, que es la misma que se ejecutaría si las peticiones
 * llegaran de IPs y dispositivos distintos al mismo tiempo.
 *
 * Usa "IE PRUEBA 1234" — no envía ningún correo ni modifica ESTADO.
 * Libera todas las sesiones de prueba al final, incluso si algo falla
 * a mitad de camino.
 *
 * Ejecutar manualmente:  probarSesionesSimultaneasYFusionDatos()
 *****************************************************/
function probarSesionesSimultaneasYFusionDatos(nombreIEPrueba){
  nombreIEPrueba = nombreIEPrueba || "IE PRUEBA 1234";
  const resultado = { pasos: {}, errores: [] };
  const dispositivos = ["DISP-A", "DISP-B", "DISP-C", "DISP-D", "DISP-E"];
  const tokensPorDispositivo = {};
  let idForo = "";

  function log(mensaje){ Logger.log(mensaje); }

  try{
    const hoja = asegurarColumnasAccesosIE_();
    const mapa = mapaHoja_(hoja);
    if(hoja.getLastRow() < 2) throw new Error("AccesosIE no tiene filas.");
    const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
    const fila = valores.find(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIEPrueba; });
    if(!fila) throw new Error("No existe " + nombreIEPrueba + " en AccesosIE.");
    idForo = String(fila[mapa.ID_FORO - 1] || "").trim();
    if(!idForo) throw new Error(nombreIEPrueba + " no tiene ID_FORO.");

    // Empezar en limpio: liberar cualquier candado que hubiera quedado.
    PropertiesService.getScriptProperties().deleteProperty(obtenerClaveSesionCodigo_("", "", idForo));

    // --- 1. Reclamar 4 dispositivos (el máximo) ---
    for(let i = 0; i < 4; i++){
      const disp = dispositivos[i];
      const sesion = reclamarSesionCodigo_("", "", disp, idForo, false);
      if(!sesion.ok) throw new Error("Dispositivo " + disp + " no pudo conectarse (debería haber cupo): " + sesion.mensaje);
      tokensPorDispositivo[disp] = sesion.tokenSesion;
    }
    resultado.pasos.cuatroDispositivosConectados = true;
    log("✅ 4 dispositivos conectados simultáneamente sin problema.");

    // --- 2. Un 5° sin forzar debe rechazarse ---
    const quintoSinForzar = reclamarSesionCodigo_("", "", "DISP-E", idForo, false);
    if(quintoSinForzar.ok) throw new Error("Un 5° dispositivo pudo conectarse sin forzar — el límite de 4 no se está respetando.");
    resultado.pasos.quintoRechazadoSinForzar = (quintoSinForzar.codigo === "SESION_YA_ABIERTA");
    log("✅ 5° dispositivo rechazado correctamente sin forzar: " + quintoSinForzar.mensaje);

    // --- 3. El 5° SÍ debe poder entrar forzando (desaloja al más inactivo: DISP-A, el primero en conectarse) ---
    const quintoForzando = reclamarSesionCodigo_("", "", "DISP-E", idForo, true);
    if(!quintoForzando.ok) throw new Error("El 5° dispositivo debería poder conectarse forzando: " + quintoForzando.mensaje);
    tokensPorDispositivo["DISP-E"] = quintoForzando.tokenSesion;
    const aTodaviaActivo = sesionActivaPorIdForo_(idForo, "DISP-A", tokensPorDispositivo["DISP-A"]);
    resultado.pasos.dispositivoDesalojadoCorrectamente = !aTodaviaActivo;
    log((aTodaviaActivo ? "❌" : "✅") + " DISP-A " + (aTodaviaActivo ? "sigue activo (no debería)" : "quedó desalojado como se esperaba") + " tras forzar el 5° cupo.");

    // --- 4. Fusión de campos: dos "dispositivos" guardan cada uno SOLO su propio campo, el resto vacío ---
    guardarAvanceForo({
      idForo: idForo,
      campos: {
        respuestaSesion1: { tipo:"text", valor:"PRUEBA FUSIÓN — dispositivo B, Sesión 1 (" + new Date().toISOString() + ")" },
        respuestaSesion2Pregunta1: { tipo:"text", valor:"" }
      }
    });
    guardarAvanceForo({
      idForo: idForo,
      campos: {
        respuestaSesion1: { tipo:"text", valor:"" },
        respuestaSesion2Pregunta1: { tipo:"text", valor:"PRUEBA FUSIÓN — dispositivo C, Sesión 2 (" + new Date().toISOString() + ")" }
      }
    });
    const datosFusionados = obtenerDatosGuardadosPorIdForo_(idForo);
    const s1Conservada = String(datosFusionados?.campos?.respuestaSesion1?.valor || "").indexOf("dispositivo B") !== -1;
    const s2Conservada = String(datosFusionados?.campos?.respuestaSesion2Pregunta1?.valor || "").indexOf("dispositivo C") !== -1;
    resultado.pasos.fusionConservaAmbosCampos = s1Conservada && s2Conservada;
    log((s1Conservada && s2Conservada ? "✅" : "❌") + " Fusión de campos — Sesión 1 (de B) conservada: " + s1Conservada + ", Sesión 2 (de C) conservada: " + s2Conservada + ".");

  }catch(error){
    resultado.errores.push(error.message);
    log("❌ ERROR: " + error.message);
  }finally{
    // Liberar todos los cupos de prueba, pase lo que pase.
    if(idForo){
      Object.keys(tokensPorDispositivo).forEach(function(disp){
        try{ liberarSesionCodigo_("", "", disp, tokensPorDispositivo[disp], idForo); }catch(e){}
      });
      // Por si algún cupo quedó en un estado raro, se limpia del todo.
      try{ PropertiesService.getScriptProperties().deleteProperty(obtenerClaveSesionCodigo_("", "", idForo)); }catch(e){}
    }
  }

  const todoBien = resultado.errores.length === 0 &&
    resultado.pasos.cuatroDispositivosConectados &&
    resultado.pasos.quintoRechazadoSinForzar &&
    resultado.pasos.dispositivoDesalojadoCorrectamente &&
    resultado.pasos.fusionConservaAmbosCampos;

  Logger.log("========================================");
  Logger.log("RESULTADO — SESIONES SIMULTÁNEAS Y FUSIÓN DE DATOS: " + (todoBien ? "✅ TODO CORRECTO" : "⚠ REVISAR"));
  Logger.log(JSON.stringify(resultado, null, 2));
  Logger.log("========================================");

  return { ok: todoBien, resultado: resultado };
}

/*****************************************************
 * PRUEBA: TRANSFERENCIA DE RESPONSABLE PRINCIPAL (4 DISPOSITIVOS)
 *
 * "Secuencia de error 1": responsable principal (RESP-1) ya no puede
 * continuar (se le da de baja como si perdiera la conexión) y
 * transfiere su rol a un colaborador ya conectado (COLAB-2, forzado
 * como el de actividad más reciente entre los 3 colaboradores, que es
 * a quién transferirResponsablePrincipalFEM() siempre elige). Verifica
 * que, tras la transferencia, COLAB-2 queda como principal, RESP-1
 * deja de serlo, y que sesionActivaPorIdForo_ — la función que
 * enviarRespuestasSesion()/enviarForoDefinitivo() usan para decidir si
 * se puede enviar — sigue devolviendo true para COLAB-2 (es decir,
 * "colaborador 2 ahora puede enviar el archivo").
 *
 * Usa "IE PRUEBA 1234" — no envía ningún correo ni modifica ESTADO.
 * Libera todas las sesiones de prueba al final, incluso si algo falla.
 *
 * Ejecutar manualmente:  probarTransferenciaResponsablePrincipalFEM()
 *****************************************************/
function probarTransferenciaResponsablePrincipalFEM(nombreIEPrueba){
  nombreIEPrueba = nombreIEPrueba || "IE PRUEBA 1234";
  const resultado = { pasos: {}, errores: [] };
  const RESP1 = "RESP-1", COLAB2 = "COLAB-2", COLAB3 = "COLAB-3", COLAB4 = "COLAB-4";
  const tokensPorDispositivo = {};
  let idForo = "";
  let clave = "";

  function log(mensaje){ Logger.log(mensaje); }

  try{
    const hoja = asegurarColumnasAccesosIE_();
    const mapa = mapaHoja_(hoja);
    if(hoja.getLastRow() < 2) throw new Error("AccesosIE no tiene filas.");
    const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
    const fila = valores.find(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIEPrueba; });
    if(!fila) throw new Error("No existe " + nombreIEPrueba + " en AccesosIE.");
    idForo = String(fila[mapa.ID_FORO - 1] || "").trim();
    if(!idForo) throw new Error(nombreIEPrueba + " no tiene ID_FORO.");
    clave = obtenerClaveSesionCodigo_("", "", idForo);

    // Empezar en limpio.
    PropertiesService.getScriptProperties().deleteProperty(clave);

    // --- 1. RESP-1 entra primero: debe quedar como responsable principal ---
    const sResp1 = reclamarSesionCodigo_("", "", RESP1, idForo, false);
    if(!sResp1.ok) throw new Error("RESP-1 no pudo conectarse: " + sResp1.mensaje);
    tokensPorDispositivo[RESP1] = sResp1.tokenSesion;
    resultado.pasos.resp1EsPrincipalAlEntrar = sResp1.esPrincipal === true;
    log((sResp1.esPrincipal ? "✅" : "❌") + " RESP-1 quedó como responsable principal al ser el primero en conectarse.");

    // --- 2. Se conectan los otros 3 colaboradores (hasta el máximo de 4) ---
    [COLAB2, COLAB3, COLAB4].forEach(function(disp){
      const s = reclamarSesionCodigo_("", "", disp, idForo, false);
      if(!s.ok) throw new Error(disp + " no pudo conectarse: " + s.mensaje);
      tokensPorDispositivo[disp] = s.tokenSesion;
      if(s.esPrincipal) throw new Error(disp + " no debería quedar como principal (ya había uno: RESP-1).");
    });
    resultado.pasos.tresColaboradoresConectados = true;
    log("✅ 4 dispositivos conectados con el mismo código: 1 responsable principal (RESP-1) + 3 colaboradores.");

    // --- 3. Antes de transferir, los 4 pasan sesionActivaPorIdForo_ ---
    const activosAntes = [RESP1, COLAB2, COLAB3, COLAB4].every(function(disp){
      return sesionActivaPorIdForo_(idForo, disp, tokensPorDispositivo[disp]);
    });
    resultado.pasos.todosActivosAntesDeTransferir = activosAntes;
    log((activosAntes ? "✅" : "❌") + " Los 4 dispositivos pasan sesionActivaPorIdForo_ antes de transferir.");

    // --- 4. SECUENCIA DE ERROR 1: RESP-1 "ya no puede continuar" y transfiere el control ---
    // transferirResponsablePrincipalFEM siempre elige al colaborador de
    // actividad más reciente — se fuerza a COLAB-2 a ser ese, para
    // probar puntualmente "se transfiere el rol a colaborador 2".
    (function forzarColab2ComoMasReciente(){
      const props = PropertiesService.getScriptProperties();
      const sesiones = leerSesionesActivas_(props, clave);
      const ahora = Date.now();
      sesiones.forEach(function(s){
        if(s.deviceId === COLAB2) s.ultimaActividad = ahora;
        else if(s.deviceId !== RESP1) s.ultimaActividad = ahora - 60000;
      });
      props.setProperty(clave, JSON.stringify(sesiones));
    })();

    const transferencia = transferirResponsablePrincipalFEM("", "", RESP1, tokensPorDispositivo[RESP1], idForo);
    resultado.pasos.transferenciaOk = !!transferencia.ok;
    resultado.pasos.transferidoAColab2 = transferencia.ok && transferencia.nuevoPrincipalDispositivoId === COLAB2;
    log((resultado.pasos.transferidoAColab2 ? "✅" : "❌") + " Transferencia de RESP-1 a COLAB-2: " + JSON.stringify(transferencia));

    // --- 5. Tras la transferencia: COLAB-2 es principal, RESP-1 ya no ---
    const sesionesDespues = leerSesionesActivas_(PropertiesService.getScriptProperties(), clave);
    const colab2 = sesionesDespues.find(function(s){ return s.deviceId === COLAB2; });
    const resp1 = sesionesDespues.find(function(s){ return s.deviceId === RESP1; });
    resultado.pasos.colab2EsPrincipalDespues = !!(colab2 && colab2.esPrincipal);
    resultado.pasos.resp1YaNoEsPrincipal = !!(resp1 && !resp1.esPrincipal);
    log((resultado.pasos.colab2EsPrincipalDespues && resultado.pasos.resp1YaNoEsPrincipal ? "✅" : "❌") +
        " Tras transferir: COLAB-2.esPrincipal=" + (colab2 && colab2.esPrincipal) + ", RESP-1.esPrincipal=" + (resp1 && resp1.esPrincipal) + ".");

    // --- 6. COLAB-2 (nuevo principal) sigue pasando sesionActivaPorIdForo_:
    //         es justo la condición que enviarRespuestasSesion() y
    //         enviarForoDefinitivo() exigen para permitir el envío. ---
    const colab2PuedeEnviar = sesionActivaPorIdForo_(idForo, COLAB2, tokensPorDispositivo[COLAB2]);
    resultado.pasos.colab2PuedeEnviar = colab2PuedeEnviar;
    log((colab2PuedeEnviar ? "✅" : "❌") + " COLAB-2 (nuevo responsable principal) puede enviar — sesionActivaPorIdForo_ = " + colab2PuedeEnviar + ".");

    // --- 7. RESP-1 ya no puede volver a transferir (ya no es principal) ---
    const segundaTransferencia = transferirResponsablePrincipalFEM("", "", RESP1, tokensPorDispositivo[RESP1], idForo);
    resultado.pasos.segundaTransferenciaRechazada = segundaTransferencia.ok === false;
    log((resultado.pasos.segundaTransferenciaRechazada ? "✅" : "❌") + " RESP-1 ya no puede transferir de nuevo (no es principal): " + segundaTransferencia.mensaje);

  }catch(error){
    resultado.errores.push(error.message);
    log("❌ ERROR: " + error.message);
  }finally{
    if(idForo){
      Object.keys(tokensPorDispositivo).forEach(function(disp){
        try{ liberarSesionCodigo_("", "", disp, tokensPorDispositivo[disp], idForo); }catch(e){}
      });
      try{ PropertiesService.getScriptProperties().deleteProperty(obtenerClaveSesionCodigo_("", "", idForo)); }catch(e){}
    }
  }

  const todoBien = resultado.errores.length === 0 &&
    resultado.pasos.resp1EsPrincipalAlEntrar &&
    resultado.pasos.tresColaboradoresConectados &&
    resultado.pasos.todosActivosAntesDeTransferir &&
    resultado.pasos.transferenciaOk &&
    resultado.pasos.transferidoAColab2 &&
    resultado.pasos.colab2EsPrincipalDespues &&
    resultado.pasos.resp1YaNoEsPrincipal &&
    resultado.pasos.colab2PuedeEnviar &&
    resultado.pasos.segundaTransferenciaRechazada;

  Logger.log("========================================");
  Logger.log("RESULTADO — TRANSFERENCIA DE RESPONSABLE PRINCIPAL: " + (todoBien ? "✅ TODO CORRECTO" : "⚠ REVISAR"));
  Logger.log(JSON.stringify(resultado, null, 2));
  Logger.log("========================================");

  return { ok: todoBien, resultado: resultado };
}

/*****************************************************
 * PRUEBA: RESPONSABLE PRINCIPAL + 3 COLABORADORES EN UNA SOLA SESIÓN
 *
 * "Secuencia de error 2": los 4 dispositivos permitidos (1 responsable
 * principal + 3 colaboradores) entran a la vez con el mismo código de
 * acceso, al mismo ID_FORO. Confirma que sesionActivaPorIdForo_
 * reconoce a los 4 simultáneamente (sin falsos negativos, que era
 * justo el bug crítico ya corregido — ver el comentario en la propia
 * función en Código.js), que solo uno queda marcado esPrincipal, que
 * un token equivocado se rechaza, que los latidos (heartbeat) de los 3
 * colaboradores no se pisan entre sí ni desalojan a nadie, y que un 5°
 * dispositivo se rechaza sin afectar a los 4 ya conectados.
 *
 * Usa "IE PRUEBA 1234" — no envía ningún correo ni modifica ESTADO.
 * Libera todas las sesiones de prueba al final, incluso si algo falla.
 *
 * Ejecutar manualmente:  probarTodosLosColaboradoresEnUnaSolaSesionFEM()
 *****************************************************/
function probarTodosLosColaboradoresEnUnaSolaSesionFEM(nombreIEPrueba){
  nombreIEPrueba = nombreIEPrueba || "IE PRUEBA 1234";
  const resultado = { pasos: {}, errores: [] };
  const RESP1 = "RESP-1", COLAB2 = "COLAB-2", COLAB3 = "COLAB-3", COLAB4 = "COLAB-4";
  const dispositivos = [RESP1, COLAB2, COLAB3, COLAB4];
  const tokensPorDispositivo = {};
  let idForo = "";
  let clave = "";

  function log(mensaje){ Logger.log(mensaje); }

  try{
    const hoja = asegurarColumnasAccesosIE_();
    const mapa = mapaHoja_(hoja);
    if(hoja.getLastRow() < 2) throw new Error("AccesosIE no tiene filas.");
    const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
    const fila = valores.find(function(f){ return String(f[mapa.IE - 1] || "").trim() === nombreIEPrueba; });
    if(!fila) throw new Error("No existe " + nombreIEPrueba + " en AccesosIE.");
    idForo = String(fila[mapa.ID_FORO - 1] || "").trim();
    if(!idForo) throw new Error(nombreIEPrueba + " no tiene ID_FORO.");
    clave = obtenerClaveSesionCodigo_("", "", idForo);

    PropertiesService.getScriptProperties().deleteProperty(clave);

    // --- 1. Los 4 entran "a la vez", con el mismo código, al mismo ID_FORO ---
    dispositivos.forEach(function(disp){
      const s = reclamarSesionCodigo_("", "", disp, idForo, false);
      if(!s.ok) throw new Error(disp + " no pudo conectarse: " + s.mensaje);
      tokensPorDispositivo[disp] = s.tokenSesion;
    });
    resultado.pasos.cuatroEnUnaSolaSesion = true;
    log("✅ RESP-1 + 3 colaboradores conectados a la vez, mismo código y mismo ID_FORO.");

    // --- 2. Cada uno pasa sesionActivaPorIdForo_ de forma independiente ---
    const estadoActivos = {};
    dispositivos.forEach(function(disp){ estadoActivos[disp] = sesionActivaPorIdForo_(idForo, disp, tokensPorDispositivo[disp]); });
    resultado.pasos.todosActivosSimultaneamente = dispositivos.every(function(disp){ return estadoActivos[disp]; });
    log((resultado.pasos.todosActivosSimultaneamente ? "✅" : "❌") + " Estado activo por dispositivo: " + JSON.stringify(estadoActivos));

    // --- 3. Solo RESP-1 queda marcado esPrincipal; los otros 3 son colaboradores ---
    const sesiones = leerSesionesActivas_(PropertiesService.getScriptProperties(), clave);
    const principales = sesiones.filter(function(s){ return s.esPrincipal; });
    resultado.pasos.unSoloPrincipal = principales.length === 1 && principales[0].deviceId === RESP1;
    log((resultado.pasos.unSoloPrincipal ? "✅" : "❌") + " Un solo responsable principal (RESP-1): " + JSON.stringify(principales.map(function(s){ return s.deviceId; })));

    // --- 4. Un token que no corresponde a nadie NO debe pasar como activo ---
    const tokenEquivocadoActivo = sesionActivaPorIdForo_(idForo, COLAB2, "token-que-no-es-de-nadie");
    resultado.pasos.tokenEquivocadoRechazado = tokenEquivocadoActivo === false;
    log((resultado.pasos.tokenEquivocadoRechazado ? "✅" : "❌") + " Un token que no corresponde a COLAB-2 se rechaza correctamente.");

    // --- 5. Los 3 colaboradores mandan su latido en la misma sesión, sin pisarse entre sí ---
    [COLAB2, COLAB3, COLAB4].forEach(function(disp){
      const latido = mantenerSesionCodigo_("", "", disp, tokensPorDispositivo[disp], idForo);
      if(!latido.ok) throw new Error("Latido de " + disp + " falló: " + JSON.stringify(latido));
    });
    const todosSiguenActivosTrasLatidos = dispositivos.every(function(disp){
      return sesionActivaPorIdForo_(idForo, disp, tokensPorDispositivo[disp]);
    });
    resultado.pasos.latidosNoConflictan = todosSiguenActivosTrasLatidos;
    log((todosSiguenActivosTrasLatidos ? "✅" : "❌") + " Tras los latidos de los 3 colaboradores, los 4 dispositivos siguen activos (ninguno se desalojó).");

    // --- 6. Un 5° dispositivo debe rechazarse: ya hay 4 (el máximo permitido) ---
    const quinto = reclamarSesionCodigo_("", "", "COLAB-5", idForo, false);
    resultado.pasos.quintoRechazado = quinto.ok === false && quinto.codigo === "SESION_YA_ABIERTA";
    log((resultado.pasos.quintoRechazado ? "✅" : "❌") + " Un 5° dispositivo (COLAB-5) es rechazado sin forzar: " + quinto.mensaje);

    // --- 7. Tras el intento rechazado, los 4 originales siguen intactos ---
    const siguenActivosTrasRechazo = dispositivos.every(function(disp){
      return sesionActivaPorIdForo_(idForo, disp, tokensPorDispositivo[disp]);
    });
    resultado.pasos.integridadTrasRechazo = siguenActivosTrasRechazo;
    log((siguenActivosTrasRechazo ? "✅" : "❌") + " Los 4 dispositivos originales siguen intactos tras el intento rechazado del 5°.");

  }catch(error){
    resultado.errores.push(error.message);
    log("❌ ERROR: " + error.message);
  }finally{
    if(idForo){
      dispositivos.forEach(function(disp){
        try{ liberarSesionCodigo_("", "", disp, tokensPorDispositivo[disp], idForo); }catch(e){}
      });
      try{ liberarSesionCodigo_("", "", "COLAB-5", "", idForo); }catch(e){}
      try{ PropertiesService.getScriptProperties().deleteProperty(obtenerClaveSesionCodigo_("", "", idForo)); }catch(e){}
    }
  }

  const todoBien = resultado.errores.length === 0 &&
    resultado.pasos.cuatroEnUnaSolaSesion &&
    resultado.pasos.todosActivosSimultaneamente &&
    resultado.pasos.unSoloPrincipal &&
    resultado.pasos.tokenEquivocadoRechazado &&
    resultado.pasos.latidosNoConflictan &&
    resultado.pasos.quintoRechazado &&
    resultado.pasos.integridadTrasRechazo;

  Logger.log("========================================");
  Logger.log("RESULTADO — TODOS LOS COLABORADORES EN UNA SOLA SESIÓN: " + (todoBien ? "✅ TODO CORRECTO" : "⚠ REVISAR"));
  Logger.log(JSON.stringify(resultado, null, 2));
  Logger.log("========================================");

  return { ok: todoBien, resultado: resultado };
}

/*****************************************************
 * REORDENAR (UNA VEZ) LAS HOJAS YA EXISTENTES POR IE
 *
 * reordenarHojasPorIE_() (Código.js) ya ordena alfabéticamente,
 * justo después de "AvancesForo", cualquier hoja de IE NUEVA que se
 * cree de aquí en adelante — pero las hojas de IE que ya existían
 * ANTES de ese cambio se quedan donde estaban (guardarEnHojaIE_
 * reutiliza la hoja existente y nunca la reordena por sí sola). Esta
 * función corre ese mismo reordenamiento una sola vez, de una vez
 * para todas las hojas de IE que ya existen en el spreadsheet.
 *
 * No borra ni modifica ningún dato — solo cambia la POSICIÓN de las
 * pestañas. Ejecutar manualmente: reordenarHojasIEExistentesFEM()
 *****************************************************/
function reordenarHojasIEExistentesFEM(){
  const ss=abrirSpreadsheet_();
  const antes=ss.getSheets().map(function(h){ return h.getName(); });
  reordenarHojasPorIE_(ss);
  const despues=ss.getSheets().map(function(h){ return h.getName(); });
  Logger.log("Orden ANTES: "+JSON.stringify(antes));
  Logger.log("Orden DESPUÉS: "+JSON.stringify(despues));
  return {ok:true, antes:antes, despues:despues};
}

/*****************************************************
 * REINTENTAR ENVÍOS DE INFORME DIFERIDOS POR CUOTA DE CORREO
 *
 * Cuando enviarInformeFEM() (Código.js) se encuentra con la cuota
 * diaria de GmailApp agotada (bug crítico en vivo, 2026-08-28: ~100
 * IE enviando el mismo día comparten UNA sola cuota de la cuenta que
 * ejecuta el script), ya no falla: guarda el ID_FORO en la hoja
 * "EnviosInformeDiferidos" y le informa a la persona que su informe
 * quedó generado y se enviará por correo más adelante.
 *
 * Esta función reintenta el envío real para cada ID_FORO todavía
 * pendiente (REINTENTADO=NO) — pensada para ejecutarse manualmente al
 * día siguiente (o cuando se sepa que la cuota ya se renovó; Google
 * la renueva a medianoche, hora del script). No reenvía nada que ya
 * se haya marcado como reintentado con éxito.
 *
 * Ejecutar manualmente: reintentarEnviosInformeDiferidosFEM()
 *****************************************************/
function reintentarEnviosInformeDiferidosFEM(){
  const hoja=asegurarHojaEnviosDiferidosFEM_();
  const last=hoja.getLastRow();
  const resultado={reenviados:[],fallidos:[],sinPendientes:false};

  if(last<2){
    resultado.sinPendientes=true;
    Logger.log("No hay envíos de informe diferidos pendientes.");
    return resultado;
  }

  const filas=hoja.getRange(2,1,last-1,3).getValues();
  for(let i=0;i<filas.length;i++){
    const idForo=String(filas[i][0]||"").trim();
    const yaReintentado=String(filas[i][2]||"").toUpperCase()==="SI";
    if(!idForo || yaReintentado) continue;

    try{
      const acceso=obtenerAccesoPorIdForoRaw_(idForo);
      if(!acceso){ resultado.fallidos.push({idForo:idForo, motivo:"ID_FORO no encontrado en AccesosIE."}); continue; }

      const pdfId=acceso.mapa.ID_PDF_INFORME
        ? String(acceso.hoja.getRange(acceso.fila, acceso.mapa.ID_PDF_INFORME).getValue()||"").trim()
        : "";
      if(!pdfId){ resultado.fallidos.push({idForo:idForo, motivo:"Esa IE no tiene ID_PDF_INFORME registrado (el informe no llegó a generarse)."}); continue; }

      const datosGuardados=obtenerDatosGuardadosPorIdForo_(idForo);
      if(!datosGuardados){ resultado.fallidos.push({idForo:idForo, motivo:"No se encontró el avance guardado en AvancesForo."}); continue; }
      datosGuardados.idForo=idForo;
      datosGuardados.institucion=datosGuardados.institucion||acceso.ie;

      const r=enviarInformeFEM(idForo, datosGuardados, pdfId);
      if(r && r.ok && !r.diferido){
        hoja.getRange(i+2,3).setValue("SI");
        resultado.reenviados.push(idForo+" ("+(datosGuardados.institucion||acceso.ie)+")");
      }else{
        resultado.fallidos.push({idForo:idForo, motivo:"Sigue sin cuota disponible — se reintentará en la próxima ejecución."});
      }
    }catch(error){
      resultado.fallidos.push({idForo:idForo, motivo:error.message});
    }
  }

  Logger.log("========================================");
  Logger.log("REINTENTO DE ENVÍOS DE INFORME DIFERIDOS");
  Logger.log("Reenviados con éxito ("+resultado.reenviados.length+"): "+JSON.stringify(resultado.reenviados));
  Logger.log("Pendientes/fallidos ("+resultado.fallidos.length+"): "+JSON.stringify(resultado.fallidos));
  Logger.log("========================================");

  return resultado;
}


/*
 * DIAGNÓSTICO REAL de "¿ya a todas las IE les llegó el correo del
 * informe?" — reintentarEnviosInformeDiferidosFEM() en 0/0 SOLO dice
 * que la cola de diferidos por cuota agotada está vacía; eso NO es lo
 * mismo que "todas recibieron el correo", porque:
 *   1) las IE cuyo correo salió bien a la primera nunca pasan por esa
 *      cola (no hay nada que reintentar), y
 *   2) un error de envío que NO sea por cuota agotada se relanza
 *      (throw) y no queda registrado en ninguna parte persistente.
 * Esta función sí compara, IE por IE, quién tiene informe generado
 * (ID_PDF_INFORME) contra quién tiene registrada la fecha real de
 * envío exitoso del correo (columna FECHA_ENVIO_CORREO_INFORME,
 * agregada para esto — no reutiliza FECHA_ENVIO, que ya significa
 * "fecha de envío definitivo del formulario").
 * Nota: informes generados ANTES de agregar esta columna no van a
 * tener FECHA_ENVIO_CORREO_INFORME aunque su correo sí se haya
 * enviado — para esos casos, esta función solo puede decir "sin
 * registro", no "sin enviar"; conviene revisar el buzón de la IE o
 * volver a llamar a enviarInformeFEM antes de darlos por no enviados.
 */
function diagnosticarCorreosInformeFEMPendientes(){
  const hoja=asegurarColumnasAccesosIE_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  const resultado={conInformeYCorreoConfirmado:[],conInformeSinRegistroDeCorreo:[],sinInformeAun:[]};

  if(ultimaFila<2){
    Logger.log("AccesosIE no tiene filas.");
    return resultado;
  }

  const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();

  for(let i=0;i<valores.length;i++){
    const fila=valores[i];
    const nombreIE=String(fila[mapa.IE-1]||"").trim();
    if(!nombreIE) continue;

    const pdfId=mapa.ID_PDF_INFORME ? String(fila[mapa.ID_PDF_INFORME-1]||"").trim() : "";
    if(!pdfId){ resultado.sinInformeAun.push(nombreIE); continue; }

    const fechaEnvioCorreo=mapa.FECHA_ENVIO_CORREO_INFORME ? String(fila[mapa.FECHA_ENVIO_CORREO_INFORME-1]||"").trim() : "";
    if(fechaEnvioCorreo){
      resultado.conInformeYCorreoConfirmado.push(nombreIE+" ("+fechaEnvioCorreo+")");
    }else{
      resultado.conInformeSinRegistroDeCorreo.push(nombreIE);
    }
  }

  Logger.log("========================================");
  Logger.log("DIAGNÓSTICO — CORREOS DE INFORME FEM");
  Logger.log("Con informe y correo CONFIRMADO ("+resultado.conInformeYCorreoConfirmado.length+"): "+JSON.stringify(resultado.conInformeYCorreoConfirmado));
  Logger.log("Con informe pero SIN registro de correo enviado ("+resultado.conInformeSinRegistroDeCorreo.length+"): "+JSON.stringify(resultado.conInformeSinRegistroDeCorreo));
  Logger.log("Todavía SIN informe generado ("+resultado.sinInformeAun.length+"): "+JSON.stringify(resultado.sinInformeAun));
  Logger.log("========================================");
  if(resultado.conInformeSinRegistroDeCorreo.length){
    Logger.log("Para las IE 'SIN registro de correo enviado' listadas arriba: si su informe se generó ANTES de este diagnóstico, puede que el correo sí se haya enviado pero sin quedar registrado (columna nueva). Revise el buzón de esa IE, o si tiene dudas, vuelva a llamar a enviarInformeFEM para esa IE.");
  }

  return resultado;
}


/*
 * RELLENO RETROACTIVO de FECHA_ENVIO_CORREO_INFORME para los informes
 * que se enviaron ANTES de que esa columna existiera (por eso salen
 * "sin registro" en diagnosticarCorreosInformeFEMPendientes aunque su
 * correo sí haya salido). En vez de adivinar, esta función busca en
 * la bandeja de ENVIADOS de la propia cuenta remitente (fuente de la
 * verdad real, no una suposición) un correo cuyo asunto sea EXACTO al
 * que arma construirCorreoInformeFEM_ ("Reporte de Informe IE "+ie) y,
 * si lo encuentra, registra la fecha real en que se envió.
 *
 * Solo toca IE con ID_PDF_INFORME ya generado y
 * FECHA_ENVIO_CORREO_INFORME todavía vacía — nunca sobrescribe un
 * registro que ya exista, y nunca envía ningún correo nuevo.
 */
function backfillFechaEnvioCorreoInformeDesdeGmailFEM(){
  const hoja=asegurarColumnasAccesosIE_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  const resultado={confirmadosPorGmail:[],sinRastroEnGmail:[],omitidosSinInforme:[],omitidosYaRegistrados:[]};

  if(ultimaFila<2){
    Logger.log("AccesosIE no tiene filas.");
    return resultado;
  }
  if(!mapa.FECHA_ENVIO_CORREO_INFORME){
    throw new Error("La columna FECHA_ENVIO_CORREO_INFORME no existe todavía en AccesosIE.");
  }

  const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();

  for(let i=0;i<valores.length;i++){
    const fila=valores[i];
    const filaHoja=i+2;
    const nombreIE=String(fila[mapa.IE-1]||"").trim();
    if(!nombreIE) continue;

    const pdfId=mapa.ID_PDF_INFORME ? String(fila[mapa.ID_PDF_INFORME-1]||"").trim() : "";
    if(!pdfId){ resultado.omitidosSinInforme.push(nombreIE); continue; }

    const yaRegistrado=String(fila[mapa.FECHA_ENVIO_CORREO_INFORME-1]||"").trim();
    if(yaRegistrado){ resultado.omitidosYaRegistrados.push(nombreIE); continue; }

    const asuntoEsperado="Reporte de Informe IE "+nombreIE;
    try{
      const hilos=GmailApp.search('in:sent subject:"'+asuntoEsperado.replace(/"/g,'\\"')+'"', 0, 5);
      let fechaMasReciente=null;
      hilos.forEach(function(hilo){
        hilo.getMessages().forEach(function(msj){
          if(msj.getSubject()!==asuntoEsperado) return;
          const fecha=msj.getDate();
          if(!fechaMasReciente || fecha.getTime()>fechaMasReciente.getTime()) fechaMasReciente=fecha;
        });
      });
      if(fechaMasReciente){
        hoja.getRange(filaHoja, mapa.FECHA_ENVIO_CORREO_INFORME).setValue(fechaMasReciente);
        resultado.confirmadosPorGmail.push(nombreIE+" ("+fechaMasReciente+")");
      }else{
        resultado.sinRastroEnGmail.push(nombreIE);
      }
    }catch(errorBusqueda){
      resultado.sinRastroEnGmail.push(nombreIE+" (error de búsqueda: "+errorBusqueda.message+")");
    }
  }

  Logger.log("========================================");
  Logger.log("RELLENO RETROACTIVO — FECHA_ENVIO_CORREO_INFORME DESDE GMAIL");
  Logger.log("Confirmados en Enviados, registrados ahora ("+resultado.confirmadosPorGmail.length+"): "+JSON.stringify(resultado.confirmadosPorGmail));
  Logger.log("SIN rastro en Enviados — revisar manualmente ("+resultado.sinRastroEnGmail.length+"): "+JSON.stringify(resultado.sinRastroEnGmail));
  Logger.log("Omitidos, ya tenían registro ("+resultado.omitidosYaRegistrados.length+"): "+JSON.stringify(resultado.omitidosYaRegistrados));
  Logger.log("Omitidos, sin informe generado ("+resultado.omitidosSinInforme.length+"): "+JSON.stringify(resultado.omitidosSinInforme));
  Logger.log("========================================");
  if(resultado.sinRastroEnGmail.length){
    Logger.log("Las IE de 'SIN rastro en Enviados' son las que de verdad hay que revisar con cuidado: no aparece ningún correo con ese asunto exacto en Enviados de "+REMITENTE_FEM+". Puede ser que el asunto haya cambiado con el tiempo, o que el correo realmente nunca haya salido — para estas conviene volver a llamar a enviarInformeFEM.");
  }

  return resultado;
}


/*
 * Regenera el Doc/PDF del informe ejecutivo de UNA IE ya en
 * producción, a partir de sus datos YA GUARDADOS (Caracterización,
 * Sesiones 1-3, asistencia QR) — sin volver a llamar
 * enviarForoDefinitivo ni tocar ninguna respuesta guardada. Sirve
 * para aplicar correcciones de formato del informe (fecha, pie de
 * página, gráficos, títulos...) a una IE cuyo informe ya se había
 * generado con una versión anterior del código. generarInformeFEM ya
 * actualiza por su cuenta ID_INFORME/ID_PDF_INFORME en AccesosIE con
 * el Doc/PDF nuevo. El nombre de la IE debe coincidir con el valor
 * guardado en AccesosIE (columna IE) — mayúsculas/minúsculas no
 * importan.
 *
 * generarInformeFEM() SIEMPRE crea un Doc/PDF nuevo (no sobrescribe
 * el anterior), así que — una vez el nuevo informe se generó bien —
 * el Doc y el PDF VIEJOS (con la fecha en inglés, sin "IE" en el pie,
 * etc.) se envían a la papelera de Drive (setTrashed, recuperable
 * desde la papelera si hiciera falta) para que en la carpeta de la
 * IE quede solo el informe correcto, sin dos versiones distintas
 * conviviendo.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función,
 * cambiar el valor por defecto de nombreIE si hace falta y presionar
 * "Ejecutar". El resultado (con el docUrl/pdfUrl nuevos) queda en
 * "Ver registros de ejecución".
 */
/*
 * Ubica el ID_FORO de una IE por su nombre exacto (mayúsculas/
 * minúsculas no importan) en AccesosIE. Compartido por
 * regenerarInformeFEMPorIE y enviarCorreoPruebaInformeFEM para no
 * duplicar la misma búsqueda.
 */
function buscarIdForoPorNombreIE_(nombreIE){
  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  const ultimaFila = hoja.getLastRow();
  if(ultimaFila < 2) throw new Error("AccesosIE no tiene filas.");
  const valores = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(function(f){
    return String(f[mapa.IE-1]||"").trim().toUpperCase() === String(nombreIE).trim().toUpperCase();
  });
  if(!fila) throw new Error("No existe \""+nombreIE+"\" en AccesosIE.");
  return String(fila[mapa.ID_FORO-1]||"").trim();
}

function regenerarInformeFEMPorIE(nombreIE){
  nombreIE = nombreIE || "EL LIMONAR";
  const idForo = buscarIdForoPorNombreIE_(nombreIE);
  const acceso = obtenerAccesoPorIdForoRaw_(idForo);
  if(!acceso) throw new Error("No se encontró el acceso de "+nombreIE+" (ID_FORO "+idForo+").");

  const datosGuardados = obtenerDatosGuardadosPorIdForo_(idForo);
  if(!datosGuardados) throw new Error("No hay datos guardados para "+nombreIE+" (ID_FORO "+idForo+").");
  datosGuardados.idForo = idForo;

  // IDs del Doc/PDF viejos, ANTES de regenerar, para poder mandarlos
  // a la papelera después (solo si la regeneración sale bien).
  const idDocViejo = acceso.mapa.ID_INFORME ? String(acceso.hoja.getRange(acceso.fila,acceso.mapa.ID_INFORME).getValue()||"").trim() : "";
  const idPdfViejo = acceso.mapa.ID_PDF_INFORME ? String(acceso.hoja.getRange(acceso.fila,acceso.mapa.ID_PDF_INFORME).getValue()||"").trim() : "";

  const informe = generarInformeFEM(idForo, datosGuardados);

  if(informe && informe.ok){
    [idDocViejo, idPdfViejo].forEach(function(idViejo){
      if(!idViejo) return;
      try{ DriveApp.getFileById(idViejo).setTrashed(true); }
      catch(errorPapelera){ Logger.log("No fue posible enviar a la papelera el archivo viejo "+idViejo+": "+errorPapelera.message); }
    });
  }

  Logger.log("Informe regenerado para "+nombreIE+": "+JSON.stringify(informe));
  return informe;
}

/*
 * PASO ÚNICO Y MANUAL: ejecutar esta función UNA SOLA VEZ desde el
 * editor de Apps Script (seleccionarla en el desplegable de
 * funciones y presionar "Ejecutar"), no desde la app web.
 *
 * Se agregó el servicio avanzado "Docs" (API de Documentos de
 * Google) al proyecto para poder corregir los títulos huérfanos del
 * informe (ver aplicarKeepWithNextATitulosInforme_ en Código.js).
 * La PRIMERA vez que un proyecto de Apps Script usa un servicio
 * avanzado nuevo, hace falta autorizar ese permiso nuevo una vez
 * — y esa autorización solo se puede conceder desde el editor
 * (aparece un cuadro de diálogo "Se requiere autorización"), nunca
 * desde una llamada de la app web ya desplegada.
 *
 * Esta función crea un Google Doc de prueba, lo lee con la API de
 * Docs (lo que dispara el cuadro de autorización si hace falta) y
 * lo borra (papelera) al terminar. Si "Ver registros de ejecución"
 * muestra "Autorización de Docs API: OK", quedó todo listo — los
 * próximos informes generados ya podrán corregir sus títulos
 * huérfanos automáticamente.
 */
function autorizarServicioAvanzadoDocs(){
  const doc=DocumentApp.create("PRUEBA — autorizar Docs API (se puede borrar)");
  try{
    const info=Docs.Documents.get(doc.getId());
    Logger.log("Autorización de Docs API: OK. Título leído vía Docs API: \""+info.title+"\".");
  }finally{
    try{ DriveApp.getFileById(doc.getId()).setTrashed(true); }catch(e){}
  }
}

/*
 * Envía el correo del informe (con los 3 agregados nuevos: enlace a
 * la ÚLTIMA VERSIÓN del informe, enlace a la CARPETA de la IE en
 * Drive, y el enlace de cierre de la VALORACIÓN DEL FORO) como
 * PRUEBA a jhonefrainsanchez@gmail.com — nunca a la institución real.
 * Reutiliza construirCorreoInformeFEM_ (Código.js), el mismo armado
 * que usa enviarInformeFEM tanto para el envío inmediato como para
 * el reintento diferido, así que lo que se ve acá es EXACTAMENTE lo
 * que le llegaría a una IE real.
 *
 * Usa los datos YA GUARDADOS de una IE que ya tenga su informe
 * generado (por defecto, "EL LIMONAR"); no envía nada a su correo
 * institucional ni al responsable — el asunto queda marcado
 * "[PRUEBA]" y el cuerpo lleva un aviso adicional en amarillo.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar".
 */
function enviarCorreoPruebaInformeFEM(nombreIE){
  nombreIE = nombreIE || "EL LIMONAR";
  const idForo = buscarIdForoPorNombreIE_(nombreIE);
  const acceso = obtenerAccesoPorIdForoRaw_(idForo);
  if(!acceso) throw new Error("No se encontró el acceso de "+nombreIE+" (ID_FORO "+idForo+").");

  const pdfId = acceso.mapa.ID_PDF_INFORME ? String(acceso.hoja.getRange(acceso.fila,acceso.mapa.ID_PDF_INFORME).getValue()||"").trim() : "";
  if(!pdfId) throw new Error(nombreIE+" todavía no tiene un informe generado (ID_PDF_INFORME vacío). Genere o regenere su informe primero (ver regenerarInformeFEMPorIE).");

  const datosGuardados = obtenerDatosGuardadosPorIdForo_(idForo);
  if(!datosGuardados) throw new Error("No hay datos guardados para "+nombreIE+" (ID_FORO "+idForo+").");
  datosGuardados.idForo = idForo;

  const ie = datosGuardados.institucion || acceso.ie;
  const ieSinPrefijo = nombreIESinPrefijoInstitucional_(ie);
  const logoIEUrlCorreo = urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(ie));

  const file = DriveApp.getFileById(pdfId);
  const linkDescarga = file.getUrl();
  const folderIE = crearCarpetaIE_(ie);
  const linkCarpeta = folderIE.getUrl();
  const linkValoracion = acceso.mapa.URL_ACCESO ? String(acceso.hoja.getRange(acceso.fila,acceso.mapa.URL_ACCESO).getValue()||"").trim() : "";
  const codigoAcceso = acceso.mapa.CODIGO_ACCESO ? String(acceso.hoja.getRange(acceso.fila,acceso.mapa.CODIGO_ACCESO).getValue()||"").trim() : "";
  const valoracionYaCompletada = !!obtenerValoracionPorIdForo_(idForo);

  const correo = construirCorreoInformeFEM_({
    ie:ie, ieSinPrefijo:ieSinPrefijo, logoIEUrlCorreo:logoIEUrlCorreo,
    linkDescarga:linkDescarga, linkCarpeta:linkCarpeta, linkValoracion:linkValoracion,
    valoracionYaCompletada:valoracionYaCompletada, codigoAcceso:codigoAcceso
  });

  const avisoPrueba="<p style=\"background:#FFF3CD;color:#664D03;padding:8px 12px;border-radius:6px;\"><strong>⚠ Correo de PRUEBA</strong> — contenido real de "+ie+", enviado únicamente a jhonefrainsanchez@gmail.com para revisión. No se envió a la institución.</p>";

  GmailApp.sendEmail("jhonefrainsanchez@gmail.com", "[PRUEBA] "+correo.subject, correo.body, {
    htmlBody: avisoPrueba+correo.htmlBody,
    from: REMITENTE_FEM,
    name: "Secretaría de Educación de Neiva (PRUEBA)",
    attachments: [file.getBlob()]
  });

  Logger.log("Correo de prueba del informe (con enlace a la última versión, a la carpeta de Drive y a la Valoración del Foro) enviado a jhonefrainsanchez@gmail.com, con los datos de "+ie+".");
}

/*
 * Envía el recordatorio de la Valoración del Foro (ver
 * enviarRecordatorioValoracionFEM_ en Código.js) a TODAS las IE que
 * ya tengan su informe generado y NO hayan diligenciado todavía la
 * Valoración — de aquí en adelante ese recordatorio ya sale solo,
 * automáticamente, cada vez que se envía el informe (ver
 * enviarInformeFEM), pero las IE que ya habían recibido su informe
 * ANTES de esa mejora no llegaron a recibirlo. Esta función es para
 * ponerse al día una sola vez con esas IE ya enviadas.
 *
 * VALIDACIÓN EXPLÍCITA (a pedido expreso): antes de enviar, se
 * confirma con obtenerValoracionPorIdForo_ que esa IE en particular
 * TODAVÍA NO tiene una valoración registrada — a quien ya la
 * diligenció no le vuelve a llegar el recordatorio.
 *
 * Se detiene (sin marcar como fallidas las que faltan) si se agota
 * la cuota diaria de correo — puede volver a ejecutarse al día
 * siguiente para las que quedaron pendientes, sin duplicar los
 * correos ya enviados.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El resultado (enviados/omitidos/fallidos)
 * queda en "Ver registros de ejecución".
 */
function enviarRecordatoriosValoracionPendientesFEM(){
  const hoja=asegurarColumnasAccesosIE_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  const resultado={enviados:[],omitidosYaValorados:[],omitidosSinInforme:[],omitidosSinCorreo:[],fallidos:[],cuotaAgotada:false};

  if(ultimaFila<2){
    Logger.log("AccesosIE no tiene filas.");
    return resultado;
  }

  const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();

  for(let i=0;i<valores.length;i++){
    if(resultado.cuotaAgotada) break;

    const fila=valores[i];
    const nombreIE=String(fila[mapa.IE-1]||"").trim();
    const idForo=String(fila[mapa.ID_FORO-1]||"").trim();
    if(!nombreIE || !idForo) continue;

    const pdfId=mapa.ID_PDF_INFORME ? String(fila[mapa.ID_PDF_INFORME-1]||"").trim() : "";
    if(!pdfId){ resultado.omitidosSinInforme.push(nombreIE); continue; }

    // Validación explícita: nunca a quien ya diligenció la valoración.
    if(obtenerValoracionPorIdForo_(idForo)){ resultado.omitidosYaValorados.push(nombreIE); continue; }

    try{
      const datosGuardados=obtenerDatosGuardadosPorIdForo_(idForo);
      const c=datosGuardados?.campos||{};
      const ie=datosGuardados?.institucion||nombreIE;
      const ieSinPrefijo=nombreIESinPrefijoInstitucional_(ie);
      const destinatario=String(c.correoIE?.valor||(mapa.EMAIL_IE?fila[mapa.EMAIL_IE-1]:"")||"").trim();
      const responsable=String(c.correo?.valor||"").trim();
      const linkValoracion=mapa.URL_ACCESO ? String(fila[mapa.URL_ACCESO-1]||"").trim() : "";
      const codigoAcceso=mapa.CODIGO_ACCESO ? String(fila[mapa.CODIGO_ACCESO-1]||"").trim() : "";

      if(!destinatario){ resultado.omitidosSinCorreo.push(nombreIE); continue; }
      if(!linkValoracion){ resultado.fallidos.push({ie:nombreIE, motivo:"Sin URL_ACCESO registrada."}); continue; }

      const logoIEUrlCorreo=urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(ie));
      enviarRecordatorioValoracionFEM_(idForo, ie, ieSinPrefijo, logoIEUrlCorreo, destinatario, responsable, linkValoracion, codigoAcceso);
      resultado.enviados.push(nombreIE);
    }catch(error){
      if(esErrorCuotaCorreoAgotada_(error)){
        resultado.cuotaAgotada=true;
        Logger.log("Cuota de correo agotada — se detiene el envío masivo de recordatorios. Vuelva a ejecutar esta función mañana para las IE que quedaron pendientes.");
        break;
      }
      resultado.fallidos.push({ie:nombreIE, motivo:error.message});
    }
  }

  Logger.log("========================================");
  Logger.log("RECORDATORIOS DE VALORACIÓN — RESULTADO");
  Logger.log("Enviados ("+resultado.enviados.length+"): "+JSON.stringify(resultado.enviados));
  Logger.log("Omitidos, ya valoraron ("+resultado.omitidosYaValorados.length+"): "+JSON.stringify(resultado.omitidosYaValorados));
  Logger.log("Omitidos, sin informe generado ("+resultado.omitidosSinInforme.length+"): "+JSON.stringify(resultado.omitidosSinInforme));
  Logger.log("Omitidos, sin correo institucional ("+resultado.omitidosSinCorreo.length+"): "+JSON.stringify(resultado.omitidosSinCorreo));
  Logger.log("Fallidos ("+resultado.fallidos.length+"): "+JSON.stringify(resultado.fallidos));
  if(resultado.cuotaAgotada) Logger.log("Se detuvo por cuota de correo agotada — vuelva a ejecutar mañana.");
  Logger.log("========================================");

  return resultado;
}

/*
 * Busca automáticamente una IE con informe ya generado pero que
 * TODAVÍA NO tenga Valoración registrada, y le corre
 * enviarCorreoPruebaInformeFEM() con esa IE — para ver, en un caso
 * real, cómo se ve la caja "falta su Valoración" (con el botón para
 * diligenciarla) en vez de la de "ya diligenciaron la Valoración".
 * El correo sigue yendo únicamente a jhonefrainsanchez@gmail.com,
 * nunca a la institución real.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". Qué IE se usó queda en "Ver registros de
 * ejecución".
 */
function enviarCorreoPruebaInformeFEMSinValoracion(){
  const hoja=asegurarColumnasAccesosIE_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  if(ultimaFila<2) throw new Error("AccesosIE no tiene filas.");

  const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();

  for(let i=0;i<valores.length;i++){
    const fila=valores[i];
    const nombreIE=String(fila[mapa.IE-1]||"").trim();
    const idForo=String(fila[mapa.ID_FORO-1]||"").trim();
    if(!nombreIE||!idForo) continue;

    const pdfId=mapa.ID_PDF_INFORME ? String(fila[mapa.ID_PDF_INFORME-1]||"").trim() : "";
    if(!pdfId) continue; // sin informe generado todavía

    if(obtenerValoracionPorIdForo_(idForo)) continue; // ya valoró, seguir buscando

    Logger.log("IE usada para la prueba (con informe, sin Valoración todavía): "+nombreIE);
    return enviarCorreoPruebaInformeFEM(nombreIE);
  }

  throw new Error("No se encontró ninguna IE con informe generado y sin Valoración pendiente — todas las que ya tienen informe también ya valoraron.");
}

/*
 * Envía el correo de RECORDATORIO de la Valoración del Foro (el que
 * llega a la IE y al responsable con el enlace personalizado y el
 * código de ingreso — ver construirCorreoRecordatorioValoracionFEM_/
 * enviarRecordatorioValoracionFEM_ en Código.js) como PRUEBA a
 * jhonefrainsanchez@gmail.com, usando los datos reales de una IE que
 * ya tenga informe generado pero SIN Valoración todavía. Si no se
 * indica una IE, busca sola la primera que cumpla esa condición
 * (igual que enviarCorreoPruebaInformeFEMSinValoracion). Nunca envía
 * nada a la institución real.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar".
 */
function enviarCorreoPruebaRecordatorioValoracionFEM(nombreIE){
  const hoja=asegurarColumnasAccesosIE_();
  const mapa=mapaHoja_(hoja);

  if(!nombreIE){
    const ultimaFila=hoja.getLastRow();
    if(ultimaFila<2) throw new Error("AccesosIE no tiene filas.");
    const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();
    for(let i=0;i<valores.length;i++){
      const fila=valores[i];
      const nombreCandidato=String(fila[mapa.IE-1]||"").trim();
      const idForoCandidato=String(fila[mapa.ID_FORO-1]||"").trim();
      if(!nombreCandidato||!idForoCandidato) continue;
      const pdfIdCandidato=mapa.ID_PDF_INFORME ? String(fila[mapa.ID_PDF_INFORME-1]||"").trim() : "";
      if(!pdfIdCandidato) continue;
      if(obtenerValoracionPorIdForo_(idForoCandidato)) continue;
      nombreIE=nombreCandidato;
      break;
    }
    if(!nombreIE) throw new Error("No se encontró ninguna IE con informe generado y sin Valoración pendiente.");
  }

  const idForo=buscarIdForoPorNombreIE_(nombreIE);
  const acceso=obtenerAccesoPorIdForoRaw_(idForo);
  if(!acceso) throw new Error("No se encontró el acceso de "+nombreIE+" (ID_FORO "+idForo+").");

  const linkValoracion=acceso.mapa.URL_ACCESO ? String(acceso.hoja.getRange(acceso.fila,acceso.mapa.URL_ACCESO).getValue()||"").trim() : "";
  if(!linkValoracion) throw new Error(nombreIE+" no tiene URL_ACCESO registrada.");
  const codigoAcceso=acceso.mapa.CODIGO_ACCESO ? String(acceso.hoja.getRange(acceso.fila,acceso.mapa.CODIGO_ACCESO).getValue()||"").trim() : "";

  const ie=acceso.ie||nombreIE;
  const ieSinPrefijo=nombreIESinPrefijoInstitucional_(ie);
  const logoIEUrlCorreo=urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(ie));

  const correo=construirCorreoRecordatorioValoracionFEM_(ieSinPrefijo, logoIEUrlCorreo, linkValoracion, codigoAcceso);
  const avisoPrueba="<p style=\"background:#FFF3CD;color:#664D03;padding:8px 12px;border-radius:6px;\"><strong>Correo de PRUEBA</strong> — contenido real de "+ie+", enviado únicamente a jhonefrainsanchez@gmail.com para revisión. No se envió a la institución.</p>";

  GmailApp.sendEmail("jhonefrainsanchez@gmail.com", "[PRUEBA] "+correo.asunto, correo.cuerpoTexto, {
    htmlBody: avisoPrueba+correo.cuerpoHTML,
    from: REMITENTE_FEM,
    name: "Secretaría de Educación de Neiva (PRUEBA)"
  });

  Logger.log("Correo de prueba del recordatorio de Valoración (con código de ingreso "+(codigoAcceso||"—")+") enviado a jhonefrainsanchez@gmail.com, con los datos de "+ie+".");
}

/*
 * Corrige el nombre del rector(a) guardado para una IE (en
 * Caracterización, campos.rector.valor) y regenera de una vez su
 * informe con el valor corregido. El nombre del rector aparece en
 * DOS lugares del informe — la tabla de Caracterización y la firma
 * "Rector(a)" al final — y ambos toman ese mismo dato guardado, así
 * que corregirlo ahí y regenerar es lo único que hace falta para que
 * el cambio se refleje en todo el documento.
 *
 * El nombre de la IE debe ser EXACTO (mayúsculas/minúsculas no
 * importan) al que aparece en la columna IE de AccesosIE — si no
 * coincide, falla con un error claro en vez de tocar la IE
 * equivocada.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función,
 * ajustar nombreIE/nombreRectorCorrecto si hace falta y presionar
 * "Ejecutar".
 */
function corregirRectorYRegenerarInformeFEM(nombreIE, nombreRectorCorrecto){
  nombreIE = nombreIE || "INSTITUTO TÉCNICO IPC ANDRÉS ROSA";
  nombreRectorCorrecto = nombreRectorCorrecto || "Jorge Luis Polania Vargas";
  const idForo=buscarIdForoPorNombreIE_(nombreIE);

  const hoja=abrirSpreadsheet_().getSheetByName(HOJA_AVANCES);
  if(!hoja) throw new Error("No se encontró la hoja "+HOJA_AVANCES+".");
  const mapa=mapaHoja_(hoja);
  const fila=buscarFilaPorIdForo_(hoja, idForo, mapa);
  if(fila<0) throw new Error("No hay datos guardados en "+HOJA_AVANCES+" para "+nombreIE+" (ID_FORO "+idForo+").");

  const raw=hoja.getRange(fila, mapa.DATOS).getValue();
  if(!raw) throw new Error("La fila de "+nombreIE+" en "+HOJA_AVANCES+" no tiene datos guardados (columna DATOS vacía).");

  const datos=JSON.parse(raw);
  datos.campos=datos.campos||{};
  const nombreAnterior=datos.campos.rector?.valor||"(vacío)";
  datos.campos.rector=Object.assign({}, datos.campos.rector, {valor:String(nombreRectorCorrecto||"").trim()});
  hoja.getRange(fila, mapa.DATOS).setValue(JSON.stringify(datos));

  Logger.log("Rector(a) de "+nombreIE+" corregido: \""+nombreAnterior+"\" -> \""+datos.campos.rector.valor+"\". Regenerando informe...");

  return regenerarInformeFEMPorIE(nombreIE);
}

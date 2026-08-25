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

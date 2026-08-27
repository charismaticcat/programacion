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
    "MARCO_ACCESO_ID": MARCO_ACCESO_ID
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
 *     AvancesForo, Participacion, AsistenciaQR y AccesosIE — incluye
 *     los códigos/tokens/enlaces de acceso de TODAS las IE, oficiales
 *     y de prueba. Oficiales NO se toca: es el catálogo oficial de
 *     instituciones, no una respuesta.
 *   - En Drive, envía a la papelera todas las carpetas por IE dentro
 *     de DRIVE_CARPETA_FEM_ID (fotos, informes Doc y PDF de todas las
 *     instituciones).
 *   - Libera cualquier candado de sesión activa en ScriptProperties.
 *
 * Nada de esto es destrucción permanente inmediata: las hojas
 * eliminadas quedan en el historial de versiones del archivo, y los
 * archivos de Drive quedan en la papelera 30 días.
 *
 * Después de ejecutar esto hay que volver a generar los accesos
 * (generarAccesosIE() para las IE oficiales; crearAccesoPrueba1234()
 * y crearIEsPruebaAdicionales() para las 10 de prueba) antes de poder
 * usar el formulario de nuevo.
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

  [HOJA_AVANCES, HOJA_PARTICIPACION, HOJA_ASISTENCIA_QR, HOJA_ACCESOS].forEach(function(nombreHoja){
    const hoja = ss.getSheetByName(nombreHoja);
    if(!hoja){ resumen.push(nombreHoja + ": no existe."); return; }
    const ultima = hoja.getLastRow();
    const borradas = Math.max(ultima - 1, 0);
    if(ultima >= 2) hoja.deleteRows(2, ultima - 1);
    resumen.push(nombreHoja + ": vaciada (" + borradas + " fila(s) borrada(s)), cabecera intacta.");
  });
  resumen.push(HOJA_OFICIALES + ": NO se tocó (catálogo oficial de instituciones, no es una respuesta).");

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

  Logger.log(resumen.join("\n"));
  return { ok:true, resumen: resumen };

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
    return { ok:true, mensaje:"AvancesForo no tiene filas todavía. Documento de análisis: " + ss.getUrl(), procesadas:0 };
  }
  const m = mapaHoja_(sh);
  if(!m.ID_FORO) return { ok:false, mensaje:"AvancesForo no tiene columna ID_FORO." };

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
  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2) return { ok:false, mensaje:"AccesosIE no tiene filas." };
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIEPrueba);
  if(!fila) return { ok:false, mensaje:"No existe " + nombreIEPrueba + " en AccesosIE. Ejecute primero crearAccesoPrueba1234() o crearIEsPruebaAdicionales()." };
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
  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2) return { ok:false, mensaje:"AccesosIE no tiene filas." };
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIEPrueba);
  if(!fila) return { ok:false, mensaje:"No existe " + nombreIEPrueba + " en AccesosIE." };
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
  const hoja = asegurarColumnasAccesosIE_();
  const mapa = mapaHoja_(hoja);
  if(hoja.getLastRow() < 2) return { ok:false, mensaje:"AccesosIE no tiene filas." };
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const fila = valores.find(f => String(f[mapa.IE - 1] || "").trim() === nombreIE);
  if(!fila) return { ok:false, mensaje:"No existe " + nombreIE + ". Ejecute primero crearAccesoPrueba1234()." };
  const idForo = String(fila[mapa.ID_FORO - 1] || "").trim();
  const dispositivoId = "PRUEBA-ANALISIS-" + idForo.slice(0, 8);

  const datosGuardados = obtenerDatosGuardadosPorIdForo_(idForo);
  if(!datosGuardados) return { ok:false, mensaje:"No hay datos guardados para " + nombreIE + ". Complete o precargue su caracterización primero." };
  datosGuardados.idForo = idForo;

  const sesion = reclamarSesionCodigo_("", "", dispositivoId, idForo, true);
  if(!sesion.ok) return { ok:false, mensaje:"No fue posible reclamar sesión: " + sesion.mensaje };

  const envio = enviarForoDefinitivo(idForo, sesion.tokenSesion, dispositivoId, datosGuardados);
  liberarSesionCodigo_("", "", dispositivoId, sesion.tokenSesion, idForo);
  if(!envio || (!envio.ok && !envio.yaEnviado)) return { ok:false, mensaje:(envio && envio.mensaje) || "Envío definitivo falló." };

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

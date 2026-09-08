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
 * Investigación puntual de UNA IE que quedó "sin rastro en Enviados"
 * en backfillFechaEnvioCorreoInformeDesdeGmailFEM — antes de reenviar
 * a ciegas, busca en Enviados por el NOMBRE de la IE (sin exigir el
 * asunto exacto) para distinguir dos casos muy distintos:
 *   a) si aparece algo, es que el asunto exacto cambió con el tiempo
 *      (por ejemplo el nombre de la IE se corrigió después) y el
 *      correo sí salió — no hay que reenviar nada;
 *   b) si de verdad no aparece nada, el correo nunca salió y sí hay
 *      que volver a llamar a enviarInformeFEM para esa IE.
 * Solo lee — nunca envía correos ni modifica AccesosIE.
 */
function investigarEnvioInformeIE(nombreIE){
  nombreIE = nombreIE || "AIPECITO";
  const idForo = buscarIdForoPorNombreIE_(nombreIE);
  const acceso = obtenerAccesoPorIdForoRaw_(idForo);
  if(!acceso) throw new Error("No se encontró el acceso de "+nombreIE+" (ID_FORO "+idForo+").");

  const pdfId = acceso.mapa.ID_PDF_INFORME ? String(acceso.hoja.getRange(acceso.fila, acceso.mapa.ID_PDF_INFORME).getValue()||"").trim() : "";
  const emailIE = acceso.mapa.EMAIL_IE ? String(acceso.hoja.getRange(acceso.fila, acceso.mapa.EMAIL_IE).getValue()||"").trim() : "";
  const fechaRegistrada = acceso.mapa.FECHA_ENVIO_CORREO_INFORME ? String(acceso.hoja.getRange(acceso.fila, acceso.mapa.FECHA_ENVIO_CORREO_INFORME).getValue()||"").trim() : "";

  Logger.log("========================================");
  Logger.log("INVESTIGACIÓN DE ENVÍO DE INFORME — "+nombreIE);
  Logger.log("ID_FORO: "+idForo+" | ID_PDF_INFORME: "+(pdfId||"(vacío)")+" | EMAIL_IE: "+(emailIE||"(vacío)")+" | FECHA_ENVIO_CORREO_INFORME registrada: "+(fechaRegistrada||"(vacío)"));

  const hilos = GmailApp.search('in:sent "'+nombreIE.replace(/"/g,'\\"')+'"', 0, 10);
  const coincidencias = [];
  hilos.forEach(function(hilo){
    hilo.getMessages().forEach(function(msj){
      coincidencias.push({asunto: msj.getSubject(), fecha: msj.getDate(), para: msj.getTo()});
    });
  });
  coincidencias.sort(function(a,b){ return b.fecha.getTime()-a.fecha.getTime(); });

  Logger.log("Correos en Enviados que mencionan \""+nombreIE+"\" (sin exigir asunto exacto), encontrados ("+coincidencias.length+"):");
  coincidencias.forEach(function(c){ Logger.log("  - \""+c.asunto+"\" -> "+c.para+" ("+c.fecha+")"); });
  if(!coincidencias.length){
    Logger.log("No se encontró NINGÚN correo en Enviados que mencione a esta IE — todo indica que el correo del informe realmente nunca salió. Se recomienda volver a llamar a enviarInformeFEM para esta IE.");
  }
  Logger.log("========================================");

  return {idForo:idForo, pdfId:pdfId, emailIE:emailIE, fechaRegistrada:fechaRegistrada, coincidencias:coincidencias};
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
 * Recordatorio de VENCIMIENTO DE PLAZO (ver
 * construirCorreoRecordatorioVencimientoFEM_/enviarRecordatorioVencimientoFEM_
 * en Código.js) para las IE REALES que todavía NO han culminado el
 * Foro (sin ID_PDF_INFORME) — se excluyen las filas de prueba
 * (TIPO="PRUEBA": IE PRUEBA 1234, IE Prueba *, IE Simulación *) para
 * no mandarle este aviso a nadie que no sea una institución real.
 * Con copia fija a ronald.polania@alcaldianeiva.gov.co en cada envío,
 * a pedido expreso.
 *
 * Se detiene (sin marcar como fallidas las que faltan) si se agota la
 * cuota diaria de correo — puede volver a ejecutarse más tarde para
 * las que quedaron pendientes, sin duplicar los correos ya enviados.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El resultado queda en "Ver registros de
 * ejecución".
 */
function enviarRecordatoriosVencimientoFEMPendientes(){
  const hoja=asegurarColumnasAccesosIE_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  const resultado={enviados:[],omitidosDePrueba:[],omitidosYaCulminaron:[],omitidosSinDatos:[],fallidos:[],cuotaAgotada:false};

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

    const tipo=mapa.TIPO ? String(fila[mapa.TIPO-1]||"").trim().toUpperCase() : "";
    if(tipo==="PRUEBA"){ resultado.omitidosDePrueba.push(nombreIE); continue; }

    const pdfId=mapa.ID_PDF_INFORME ? String(fila[mapa.ID_PDF_INFORME-1]||"").trim() : "";
    if(pdfId){ resultado.omitidosYaCulminaron.push(nombreIE); continue; }

    const destinatario=mapa.EMAIL_IE ? String(fila[mapa.EMAIL_IE-1]||"").trim() : "";
    const responsable=mapa.EMAIL_RESPONSABLE ? String(fila[mapa.EMAIL_RESPONSABLE-1]||"").trim() : "";
    const urlAcceso=mapa.URL_ACCESO ? String(fila[mapa.URL_ACCESO-1]||"").trim() : "";
    const codigoAcceso=mapa.CODIGO_ACCESO ? String(fila[mapa.CODIGO_ACCESO-1]||"").trim() : "";

    if(!destinatario || !urlAcceso){ resultado.omitidosSinDatos.push(nombreIE); continue; }

    try{
      const ieSinPrefijo=nombreIESinPrefijoInstitucional_(nombreIE);
      const logoIEUrlCorreo=urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(nombreIE));
      enviarRecordatorioVencimientoFEM_(destinatario, responsable, ieSinPrefijo, logoIEUrlCorreo, urlAcceso, codigoAcceso);
      resultado.enviados.push(nombreIE);
    }catch(error){
      if(esErrorCuotaCorreoAgotada_(error)){
        resultado.cuotaAgotada=true;
        Logger.log("Cuota de correo agotada — se detiene el envío masivo de recordatorios de vencimiento. Vuelva a ejecutar esta función más tarde para las IE que quedaron pendientes.");
        break;
      }
      resultado.fallidos.push({ie:nombreIE, motivo:error.message});
    }
  }

  Logger.log("========================================");
  Logger.log("RECORDATORIOS DE VENCIMIENTO DE PLAZO — RESULTADO");
  Logger.log("Enviados ("+resultado.enviados.length+"): "+JSON.stringify(resultado.enviados));
  Logger.log("Omitidos, filas de prueba ("+resultado.omitidosDePrueba.length+"): "+JSON.stringify(resultado.omitidosDePrueba));
  Logger.log("Omitidos, ya culminaron ("+resultado.omitidosYaCulminaron.length+"): "+JSON.stringify(resultado.omitidosYaCulminaron));
  Logger.log("Omitidos, sin correo o sin URL_ACCESO ("+resultado.omitidosSinDatos.length+"): "+JSON.stringify(resultado.omitidosSinDatos));
  Logger.log("Fallidos ("+resultado.fallidos.length+"): "+JSON.stringify(resultado.fallidos));
  if(resultado.cuotaAgotada) Logger.log("Se detuvo por cuota de correo agotada — vuelva a ejecutar más tarde.");
  Logger.log("========================================");

  return resultado;
}

/*
 * Correo de PRUEBA del recordatorio de vencimiento — únicamente a
 * jhonefrainsanchez@gmail.com, con el contenido real de una IE que
 * todavía no ha culminado (o la indicada por parámetro). Nunca lleva
 * la copia a ronald.polania@alcaldianeiva.gov.co (esa copia es solo
 * para el envío real, ver enviarRecordatorioVencimientoFEM_).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". Qué IE se usó queda en "Ver registros de
 * ejecución".
 */
function enviarCorreoPruebaRecordatorioVencimientoFEM(nombreIE){
  const hoja=asegurarColumnasAccesosIE_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  if(ultimaFila<2) throw new Error("AccesosIE no tiene filas.");
  const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();

  let fila;
  if(nombreIE){
    fila=valores.find(function(f){ return String(f[mapa.IE-1]||"").trim().toUpperCase()===String(nombreIE).trim().toUpperCase(); });
    if(!fila) throw new Error("No existe \""+nombreIE+"\" en AccesosIE.");
  }else{
    fila=valores.find(function(f){
      const tipo=mapa.TIPO ? String(f[mapa.TIPO-1]||"").trim().toUpperCase() : "";
      const pdfId=mapa.ID_PDF_INFORME ? String(f[mapa.ID_PDF_INFORME-1]||"").trim() : "";
      return tipo!=="PRUEBA" && !pdfId;
    });
    if(!fila) throw new Error("No se encontró ninguna IE real pendiente de culminar. Indique el nombre de la IE por parámetro.");
    nombreIE=String(fila[mapa.IE-1]||"").trim();
  }

  const urlAcceso=mapa.URL_ACCESO ? String(fila[mapa.URL_ACCESO-1]||"").trim() : "";
  const codigoAcceso=mapa.CODIGO_ACCESO ? String(fila[mapa.CODIGO_ACCESO-1]||"").trim() : "";
  const ieSinPrefijo=nombreIESinPrefijoInstitucional_(nombreIE);
  const logoIEUrlCorreo=urlPublicaLogoDrive_(obtenerLogoIdPorNombreIE_(nombreIE));

  const correo=construirCorreoRecordatorioVencimientoFEM_(ieSinPrefijo, logoIEUrlCorreo, urlAcceso, codigoAcceso);
  const avisoPrueba="<p style=\"background:#FFF3CD;color:#664D03;padding:8px 12px;border-radius:6px;\"><strong>⚠ Correo de PRUEBA</strong> — contenido real de "+nombreIE+", enviado únicamente a jhonefrainsanchez@gmail.com para revisión. No se envió a la institución ni lleva copia a ronald.polania@alcaldianeiva.gov.co.</p>";

  GmailApp.sendEmail("jhonefrainsanchez@gmail.com", "[PRUEBA] "+correo.asunto, correo.cuerpoTexto, {
    htmlBody: avisoPrueba+correo.cuerpoHTML,
    from: REMITENTE_FEM,
    name: "Secretaría de Educación de Neiva (PRUEBA)"
  });

  Logger.log("Correo de prueba del recordatorio de vencimiento enviado a jhonefrainsanchez@gmail.com, con los datos de "+nombreIE+".");
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

/*****************************************************
 * ORGANIZACIÓN Y REPORTES POR GRUPO DE TRABAJO (G1-G6)
 *
 * Tres funciones independientes, pensadas para ejecutarse manualmente
 * desde el editor de Apps Script (en cualquier orden, y las veces que
 * hagan falta — todas son idempotentes, no duplican carpetas, copias
 * ni hojas si ya existen de una ejecución anterior):
 *
 *   1) organizarInformesPorGrupoFEM()   — carpetas de Drive por grupo
 *      con copias de los PDF e informes editables ya generados.
 *   2) compilarRespuestasPorGrupoFEM()  — un Doc editable por grupo
 *      con las respuestas de todas sus IE.
 *   3) generarSpreadsheetGrupoFEM()     — una sola hoja de cálculo con
 *      participación, valoración y percepción, todas por grupo.
 *
 * Las tres reutilizan mapaGruposFEM_() como única fuente de "qué IE
 * pertenece a qué grupo", para que las tres cuenten siempre la misma
 * historia.
 *****************************************************/

/*
 * Agrupa todas las IE REALES (excluye TIPO="PRUEBA") de AccesosIE
 * según su grupo de trabajo (G1-G6), usando obtenerGrupoRealDeIEFEM_
 * (Código.js) — que prioriza la respuesta de caracterización ya
 * guardada por la propia IE y solo recurre al catálogo fijo si esa
 * respuesta todavía no existe. Las IE que no calzan en ninguna de las
 * dos fuentes (ni catálogo ni respuesta guardada) quedan aparte, en
 * "SIN_GRUPO", en vez de perderse en silencio.
 */
function mapaGruposFEM_(){
  const hoja=asegurarColumnasAccesosIE_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  const grupos={G1:[],G2:[],G3:[],G4:[],G5:[],G6:[],SIN_GRUPO:[]};
  if(ultimaFila<2) return grupos;

  const valores=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();
  valores.forEach(function(fila){
    const nombreIE=String(fila[mapa.IE-1]||"").trim();
    const idForo=String(fila[mapa.ID_FORO-1]||"").trim();
    if(!nombreIE || !idForo) return;

    const tipo=mapa.TIPO ? String(fila[mapa.TIPO-1]||"").trim().toUpperCase() : "";
    if(tipo==="PRUEBA") return;

    const pdfId=mapa.ID_PDF_INFORME ? String(fila[mapa.ID_PDF_INFORME-1]||"").trim() : "";
    const grupo=obtenerGrupoRealDeIEFEM_(nombreIE, idForo) || "SIN_GRUPO";
    const entrada={nombreIE:nombreIE, idForo:idForo, pdfId:pdfId};
    if(grupos[grupo]) grupos[grupo].push(entrada); else grupos.SIN_GRUPO.push(entrada);
  });

  return grupos;
}

const GRUPOS_FEM_ORDEN_=["G1","G2","G3","G4","G5","G6"];

/*
 * 1) Crea (si no existen) la carpeta de cada grupo y sus dos
 * subcarpetas, y copia ahí — NUNCA MUEVE — el PDF más reciente y el
 * Doc editable de cada IE del grupo que ya tenga informe generado.
 * Se copia, no se mueve, porque el PDF original en la carpeta propia
 * de la IE (DRIVE_CARPETA_FEM_ID/{IE}) sigue siendo el que usan
 * enviarInformeFEM/reintentarEnviosInformeDiferidosFEM para reenviar
 * o adjuntar por correo — moverlo de ahí rompería esos flujos en
 * vivo. Idempotente: si una copia con el mismo nombre ya existe en la
 * carpeta del grupo, no se duplica.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El resultado (qué se copió, qué faltó) queda
 * en "Ver registros de ejecución".
 */
function organizarInformesPorGrupoFEM(){
  const grupos=mapaGruposFEM_();
  const resultado={};

  GRUPOS_FEM_ORDEN_.forEach(function(g){
    const miembros=(grupos[g]||[]).slice().sort(function(a,b){ return a.nombreIE.localeCompare(b.nombreIE,"es"); });
    const carpetas=crearEstructuraCarpetasGrupoFEM_(g);
    const pdfCopiados=[], editablesCopiados=[], sinInforme=[], fallidos=[];

    miembros.forEach(function(m){
      if(!m.pdfId){ sinInforme.push(m.nombreIE); return; }

      try{
        const nombreArchivo="Informe Ejecutivo - "+m.nombreIE+" FEM 2026";
        const pdfOriginal=obtenerPdfInformeMasRecienteFEM_(crearCarpetaIE_(m.nombreIE), nombreArchivo, m.pdfId);
        if(!carpetas.sentFolder.getFilesByName(pdfOriginal.getName()).hasNext()){
          pdfOriginal.makeCopy(pdfOriginal.getName(), carpetas.sentFolder);
        }
        pdfCopiados.push(m.nombreIE);
      }catch(errorPdf){
        fallidos.push({ie:m.nombreIE, motivo:"PDF: "+errorPdf.message});
      }

      try{
        const nombreDoc="Informe Ejecutivo - "+m.nombreIE+" FEM 2026";
        const it=DriveApp.getFolderById(DRIVE_CARPETA_EDITABLES_FEM_ID).getFilesByName(nombreDoc);
        if(it.hasNext()){
          const docOriginal=it.next();
          if(!carpetas.editableFolder.getFilesByName(nombreDoc).hasNext()){
            docOriginal.makeCopy(nombreDoc, carpetas.editableFolder);
          }
          editablesCopiados.push(m.nombreIE);
        }
      }catch(errorDoc){
        fallidos.push({ie:m.nombreIE, motivo:"Editable: "+errorDoc.message});
      }
    });

    resultado[g]={
      carpetaGrupo:carpetas.grupoFolder.getUrl(),
      pdfCopiados:pdfCopiados,
      editablesCopiados:editablesCopiados,
      sinInformeTodavia:sinInforme,
      fallidos:fallidos
    };
  });

  Logger.log("========================================");
  Logger.log("ORGANIZACIÓN DE INFORMES POR GRUPO — RESULTADO");
  Logger.log(JSON.stringify(resultado, null, 2));
  if((grupos.SIN_GRUPO||[]).length){
    Logger.log("IE SIN GRUPO asignado (ni respuesta guardada ni catálogo) — revisar manualmente: "+JSON.stringify(grupos.SIN_GRUPO.map(function(m){return m.nombreIE;})));
  }
  Logger.log("========================================");

  return resultado;
}

/*
 * 2) Un Google Doc editable POR GRUPO que compila, con un título por
 * IE (orden alfabético, con su logo al lado cuando existe) y debajo
 * sus respuestas, las Sesiones 1/2/3 de todas las IE del grupo. Se
 * deja en la raíz de la carpeta del grupo (no dentro de ninguna de
 * las dos subcarpetas: es un documento aparte que sintetiza el grupo
 * completo, distinto de los informes individuales que ya se copian a
 * "Informes editables de grupo GN" con organizarInformesPorGrupoFEM).
 *
 * NUNCA crea un documento nuevo si ya existe uno con el mismo nombre
 * en la carpeta del grupo: lo REESCRIBE completo en el mismo archivo
 * (mismo ID/enlace) — así el enlace que ya se compartió no cambia de
 * una ejecución a la siguiente. Solo la primera vez, si no existe
 * ninguno todavía, se crea.
 *
 * INCIDENTE 2026-09-05: la hoja AvancesForo perdió las filas de la
 * mayoría de las IE que ya habían enviado su Foro (por eso antes solo
 * aparecían 1-2 IE por grupo en el compilado) y, aparte, se confirmó
 * que la Sesión 3 (S3_P1/P3/P4) NUNCA quedó bien reflejada en
 * "Análisis FEM 2026" para NINGUNA IE — está en blanco ahí aunque el
 * Doc real de cada IE sí tenga las respuestas completas. Por eso,
 * para las Sesiones 1/2/3 esta función usa como fuente PRINCIPAL el
 * propio Doc editable ya generado de cada IE (ver
 * extraerSesionesDesdeDocEditableFEM_ en Código.js — es exactamente
 * lo que se envió, la fuente más confiable que existe), y solo si esa
 * IE no tiene Doc todavía recurre a los datos en vivo de AvancesForo
 * y, en último caso, al respaldo de "Análisis FEM 2026" (que sigue
 * sirviendo para Sesión 1 y 2). La Sesión Propia/4 (opcional) se toma
 * de los datos en vivo cuando existen — si una IE la llenó pero su
 * fila de AvancesForo no se pudo recuperar, no hay forma de traerla
 * de vuelta y simplemente no aparece para esa IE, igual que para
 * quienes nunca la llenaron.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar".
 */
/*
 * Resuelve las Sesiones 1/2/3 (y la Sesión Propia/4, si la llenó) de
 * UNA IE, probando en orden: el Doc editable ya generado (fuente más
 * confiable — ver extraerSesionesDesdeDocEditableFEM_), los datos en
 * vivo de AvancesForo, y como último recurso el respaldo de "Análisis
 * FEM 2026" (bueno solo para Sesión 1/2). Compartida entre
 * compilarRespuestasPorGrupoFEM y generarDocumentosAnalisisPorGrupoFEM
 * para no resolver esto de dos formas distintas.
 * Devuelve {sesiones, sesionPropia, fuente} — sesiones es null si
 * ninguna fuente tuvo nada.
 */
function resolverSesionesIEFEM_(m){
  let sesiones=null, fuente="";
  try{
    sesiones=extraerSesionesDesdeDocEditableFEM_(m.nombreIE);
    if(sesiones) fuente="doc";
  }catch(error){ Logger.log("Sesiones desde el Doc editable de "+m.nombreIE+": "+error.message); }

  let datosVivos=null;
  try{ datosVivos=obtenerDatosGuardadosPorIdForo_(m.idForo); }
  catch(error){ Logger.log("Datos guardados de "+m.nombreIE+": "+error.message); }

  if(!sesiones && datosVivos){ sesiones=obtenerRespuestasSesionesParaCompilado_(datosVivos); fuente="vivo"; }

  if(!sesiones){
    try{
      sesiones=obtenerRespuestasSesionesDesdeAnalisisFEM_(m.idForo);
      if(sesiones) fuente="respaldo";
    }catch(error){ Logger.log("Sesiones de respaldo de "+m.nombreIE+": "+error.message); }
  }

  let sesionPropia=null;
  if(datosVivos){
    try{
      const sp=obtenerSesionPropia_(datosVivos);
      if(sp.tieneContenido) sesionPropia=sp;
    }catch(error){ Logger.log("Sesión Propia de "+m.nombreIE+": "+error.message); }
  }

  return {sesiones:sesiones, sesionPropia:sesionPropia, fuente:fuente};
}

function compilarRespuestasPorGrupoFEM(){
  const grupos=mapaGruposFEM_();
  const resultado={};

  GRUPOS_FEM_ORDEN_.forEach(function(g){
    const miembros=(grupos[g]||[]).slice().sort(function(a,b){ return a.nombreIE.localeCompare(b.nombreIE,"es"); });
    const listaConSesiones=[], sinDatos=[], recuperadosDeRespaldo=[];

    miembros.forEach(function(m){
      const resuelto=resolverSesionesIEFEM_(m);
      const sesiones=resuelto.sesiones, sesionPropia=resuelto.sesionPropia, fuente=resuelto.fuente;

      if(!sesiones){ sinDatos.push(m.nombreIE); return; }

      let logoBlob=null;
      try{
        const logoId=obtenerLogoIdPorNombreIE_(m.nombreIE);
        if(logoId) logoBlob=DriveApp.getFileById(logoId).getBlob();
      }catch(error){ Logger.log("Logo de "+m.nombreIE+" para el compilado: "+error.message); }

      listaConSesiones.push({nombreIE:m.nombreIE, sesiones:sesiones, sesionPropia:sesionPropia, logoBlob:logoBlob});
      if(fuente==="respaldo") recuperadosDeRespaldo.push(m.nombreIE);
    });

    if(!listaConSesiones.length){
      resultado[g]={omitido:"Ninguna IE del grupo tiene datos guardados todavía (ni en vivo ni en el respaldo).", sinDatos:sinDatos};
      return;
    }

    const carpetas=crearEstructuraCarpetasGrupoFEM_(g);
    const nombreDoc="Respuestas Compiladas - Grupo "+g+" FEM 2026";
    const existentesIt=carpetas.grupoFolder.getFilesByName(nombreDoc);
    const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
    // Por si quedó más de una copia de una ejecución vieja (de antes
    // de que esta función actualizara el mismo archivo en vez de
    // crear uno nuevo cada vez): se deja solo la más reciente.
    while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

    const archivoDoc=generarDocumentoCompiladoGrupoFEM_(g, listaConSesiones, archivoExistente?archivoExistente.getId():null);
    if(!archivoExistente){
      carpetas.grupoFolder.addFile(archivoDoc);
      try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
    }

    resultado[g]={
      documento:archivoDoc.getUrl(),
      ieIncluidas:listaConSesiones.map(function(x){return x.nombreIE;}),
      recuperadosDeRespaldoDeAnalisisFEM:recuperadosDeRespaldo,
      sinDatos:sinDatos
    };
  });

  Logger.log("========================================");
  Logger.log("DOCUMENTOS COMPILADOS POR GRUPO — RESULTADO");
  Logger.log(JSON.stringify(resultado, null, 2));
  Logger.log("========================================");

  return resultado;
}

/*
 * 3) Hoja de cálculo dedicada, aparte del documento de análisis
 * general (Análisis FEM 2026), con tres hojas: "Participación",
 * "Valoración" y "Percepción por Grupo". Su ID queda guardado en
 * ScriptProperties para que volver a ejecutar esta función actualice
 * la MISMA hoja de cálculo en vez de crear una nueva cada vez.
 */
const CLAVE_PROP_SPREADSHEET_GRUPOS_FEM_="SPREADSHEET_GRUPOS_FEM_ID";

function obtenerSpreadsheetGruposFEM_(){
  const props=PropertiesService.getScriptProperties();
  const idGuardado=props.getProperty(CLAVE_PROP_SPREADSHEET_GRUPOS_FEM_);
  if(idGuardado){
    try{ return SpreadsheetApp.openById(idGuardado); }
    catch(e){ Logger.log("La hoja de grupos guardada ("+idGuardado+") ya no es accesible, se creará una nueva: "+e.message); }
  }
  const ss=SpreadsheetApp.create("Participación, Valoración y Percepción por Grupo — FEM 2026");
  try{
    const archivo=DriveApp.getFileById(ss.getId());
    DriveApp.getFolderById(DRIVE_CARPETA_FEM_ID).addFile(archivo);
    DriveApp.getRootFolder().removeFile(archivo);
  }catch(e){ Logger.log("No fue posible mover la hoja de grupos a la carpeta del FEM: "+e.message); }
  props.setProperty(CLAVE_PROP_SPREADSHEET_GRUPOS_FEM_, ss.getId());
  return ss;
}

function hojaLimpiaGrupoFEM_(ss, nombre){
  let sh=ss.getSheetByName(nombre);
  if(!sh){ sh=ss.insertSheet(nombre); }
  else{ sh.getCharts().forEach(function(c){ sh.removeChart(c); }); sh.clear(); }
  return sh;
}

// Etiqueta legible de un grupo, para encabezados y títulos de gráficos.
function etiquetaGrupoFEM_(g){ return g==="SIN_GRUPO" ? "Sin grupo asignado" : "Grupo "+g; }

function construirHojaParticipacionGrupoFEM_(ss, grupos){
  const sh=hojaLimpiaGrupoFEM_(ss, "Participación");
  const headers=["IE","Grupo"].concat(ETIQUETAS_PARTICIPACION_ANALISIS_).concat(["Total"]);
  const filasIE=[];
  const gruposConSinAsignar=GRUPOS_FEM_ORDEN_.concat((grupos.SIN_GRUPO||[]).length?["SIN_GRUPO"]:[]);

  // sumasPorGrupo[grupo][índice de rol] = total de ese rol en ese grupo.
  const sumasPorGrupo={};
  gruposConSinAsignar.forEach(function(g){ sumasPorGrupo[g]=ROLES_PARTICIPACION_ANALISIS_.map(function(){return 0;}); });

  gruposConSinAsignar.forEach(function(g){
    (grupos[g]||[]).slice().sort(function(a,b){return a.nombreIE.localeCompare(b.nombreIE,"es");}).forEach(function(m){
      let conteo=ROLES_PARTICIPACION_ANALISIS_.map(function(){return 0;});
      try{
        const datos=obtenerDatosGuardadosPorIdForo_(m.idForo);
        const c=(datos&&datos.campos)||{};
        conteo=ROLES_PARTICIPACION_ANALISIS_.map(function(id){ return Number(c["participantes"+id]?.valor||0); });
      }catch(error){ Logger.log("Participación de "+m.nombreIE+": "+error.message); }
      const total=conteo.reduce(function(a,b){return a+b;},0);
      filasIE.push([m.nombreIE, etiquetaGrupoFEM_(g)].concat(conteo).concat([total]));
      conteo.forEach(function(v,i){ sumasPorGrupo[g][i]+=v; });
    });
  });

  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  if(filasIE.length) sh.getRange(2,1,filasIE.length,headers.length).setValues(filasIE);

  // Un gráfico por CADA TIPO DE PARTICIPANTE que sí participó (total
  // municipal > 0), comparando el total de ese rol entre los grupos.
  const colInicioAyuda=headers.length+2;
  let filaAyuda=1;
  ROLES_PARTICIPACION_ANALISIS_.forEach(function(idRol, i){
    const totalRol=gruposConSinAsignar.reduce(function(s,g){return s+sumasPorGrupo[g][i];},0);
    if(totalRol<=0) return;
    const etiqueta=ETIQUETAS_PARTICIPACION_ANALISIS_[i];
    const filaInicio=filaAyuda;
    sh.getRange(filaInicio,colInicioAyuda,1,2).setValues([["Grupo",etiqueta]]);
    const datosGrupo=gruposConSinAsignar.map(function(g){ return [etiquetaGrupoFEM_(g), sumasPorGrupo[g][i]]; });
    sh.getRange(filaInicio+1,colInicioAyuda,datosGrupo.length,2).setValues(datosGrupo);
    const rango=sh.getRange(filaInicio,colInicioAyuda,datosGrupo.length+1,2);
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(rango)
      .setOption("title","Participación de "+etiqueta+" por grupo")
      .setOption("width",420).setOption("height",260)
      .setPosition(filaInicio,colInicioAyuda+3,0,0).build());
    filaAyuda+=datosGrupo.length+3;
  });

  sh.autoResizeColumns(1,headers.length);
  return sh;
}

function construirHojaValoracionGrupoFEM_(ss, grupos){
  const sh=hojaLimpiaGrupoFEM_(ss, "Valoración");
  const headers=["IE","Grupo","Estado","Nota promedio","P1 diálogo y reflexión","P2 participación","P3 ideas y propuestas","P4 satisfacción del instrumento","P1 mejora","P2 mejora","P3 mejora","P4 mejora","P5 sugerencias"];
  const filas=[];
  const gruposConSinAsignar=GRUPOS_FEM_ORDEN_.concat((grupos.SIN_GRUPO||[]).length?["SIN_GRUPO"]:[]);

  gruposConSinAsignar.forEach(function(g){
    (grupos[g]||[]).slice().sort(function(a,b){return a.nombreIE.localeCompare(b.nombreIE,"es");}).forEach(function(m){
      const val=obtenerValoracionPorIdForo_(m.idForo);
      if(val){
        // numeroLocalizado_ (no Number() directo): la nota puede venir
        // con coma decimal ("4,75") por el formato de hoja en español,
        // y Number("4,75") da NaN — ya se vio este mismo bug en el
        // informe individual (ver numeroLocalizado_ en Código.js).
        filas.push([m.nombreIE, etiquetaGrupoFEM_(g), "Diligenciada", numeroLocalizado_(val.nota), val.p1, val.p2, val.p3, val.p4, val.p1Mejora, val.p2Mejora, val.p3Mejora, val.p4Mejora, val.p5]);
      }else{
        filas.push([m.nombreIE, etiquetaGrupoFEM_(g), "Pendiente", "", "", "", "", "", "", "", "", "", ""]);
      }
    });
  });

  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  if(filas.length) sh.getRange(2,1,filas.length,headers.length).setValues(filas);

  // Gráfico 1: nota promedio por IE (solo las que ya valoraron).
  const colAyuda=headers.length+2;
  const conNota=filas.filter(function(f){ return f[2]==="Diligenciada" && Number(f[3])>0; }).map(function(f){ return [f[0], Number(f[3])]; });
  let filaSiguiente=1;
  if(conNota.length){
    sh.getRange(1,colAyuda,1,2).setValues([["IE","Nota promedio"]]);
    sh.getRange(2,colAyuda,conNota.length,2).setValues(conNota);
    const rango=sh.getRange(1,colAyuda,conNota.length+1,2);
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(rango)
      .setOption("title","Nota promedio de valoración por IE").setOption("vAxis.viewWindow.max",5)
      .setOption("width",480).setOption("height",280)
      .setPosition(1,colAyuda+3,0,0).build());
    filaSiguiente=conNota.length+4;
  }

  // Gráfico 2: nota promedio (del promedio de sus IE) por grupo.
  const promediosPorGrupo=gruposConSinAsignar.map(function(g){
    const notas=filas.filter(function(f){ return f[1]===etiquetaGrupoFEM_(g) && f[2]==="Diligenciada" && Number(f[3])>0; }).map(function(f){ return Number(f[3]); });
    const promedio=notas.length ? notas.reduce(function(a,b){return a+b;},0)/notas.length : 0;
    return [etiquetaGrupoFEM_(g), Number(promedio.toFixed(2))];
  }).filter(function(f){ return f[1]>0; });
  if(promediosPorGrupo.length){
    sh.getRange(filaSiguiente,colAyuda,1,2).setValues([["Grupo","Nota promedio"]]);
    sh.getRange(filaSiguiente+1,colAyuda,promediosPorGrupo.length,2).setValues(promediosPorGrupo);
    const rango2=sh.getRange(filaSiguiente,colAyuda,promediosPorGrupo.length+1,2);
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(rango2)
      .setOption("title","Nota promedio de valoración por grupo").setOption("vAxis.viewWindow.max",5)
      .setOption("width",480).setOption("height",280)
      .setPosition(filaSiguiente,colAyuda+3,0,0).build());
  }

  sh.autoResizeColumns(1,headers.length);
  return sh;
}

/*
 * Escribe, a partir de filaActual, el bloque de percepción (texto +
 * tablas con Votos Y Porcentaje + gráficos) de UNA lista de
 * asistentes ya combinada — se usa tanto para el bloque "TODAS LAS
 * IE" (municipal) como para cada uno de los G1-G6. Devuelve la fila
 * siguiente libre para el próximo bloque.
 */
function escribirBloquePercepcionFEM_(sh, filaActual, tituloBloque, asistentes){
  sh.getRange(filaActual,1).setValue(tituloBloque);
  sh.getRange(filaActual,1).setFontWeight("bold").setFontSize(11);
  filaActual+=1;

  if(!asistentes.length){
    sh.getRange(filaActual,1).setValue("Todavía no hay asistencia registrada por código QR.");
    return filaActual+3;
  }

  const tF=tallyOpciones_(asistentes,"fortalezas"), tD=tallyOpciones_(asistentes,"dificultades");
  const total=asistentes.length;
  sh.getRange(filaActual,1).setValue(
    "Los "+total+" participantes que firmaron asistencia por código QR expresaron que las principales fortalezas institucionales identificadas en el Foro fueron "+top3Texto_(tF)+
    ", mientras que las principales oportunidades de mejoramiento institucional identificadas fueron "+top3Texto_(tD)+"."
  );
  sh.getRange(filaActual,1,1,1).setWrap(true);
  filaActual+=2;

  const filaTablas=filaActual;
  function filasConPorcentaje_(tally){
    return tally.map(function(x){ return [x.opcion, x.votos, ((x.votos/total)*100).toFixed(1)+"%"]; });
  }
  sh.getRange(filaTablas,1,1,3).setValues([["Fortaleza","Votos","% de asistentes"]]);
  if(tF.length) sh.getRange(filaTablas+1,1,tF.length,3).setValues(filasConPorcentaje_(tF));

  sh.getRange(filaTablas,5,1,3).setValues([["Oportunidad de mejoramiento","Votos","% de asistentes"]]);
  if(tD.length) sh.getRange(filaTablas+1,5,tD.length,3).setValues(filasConPorcentaje_(tD));

  // Un gráfico de Votos (número) y uno de Porcentaje, para cada lista
  // — así queda visible tanto el número como el porcentaje pedido.
  if(tF.length){
    const rangoFVotos=sh.getRange(filaTablas,1,tF.length+1,2);
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.BAR).addRange(rangoFVotos)
      .setOption("title","Fortalezas más votadas (número) — "+tituloBloque)
      .setOption("width",380).setOption("height",Math.max(220,40+tF.length*22))
      .setPosition(filaTablas,9,0,0).build());
    // Rango no contiguo (columna 1 + columna 3, la de porcentaje):
    // dos addRange() en el mismo gráfico se combinan en una sola
    // tabla de datos, sin necesidad de una columna auxiliar.
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(filaTablas,1,tF.length+1,1))
      .addRange(sh.getRange(filaTablas,3,tF.length+1,1))
      .setOption("title","Fortalezas más votadas (% de asistentes) — "+tituloBloque)
      .setOption("width",380).setOption("height",Math.max(220,40+tF.length*22))
      .setPosition(filaTablas,14,0,0).build());
  }
  if(tD.length){
    const rangoDVotos=sh.getRange(filaTablas,5,tD.length+1,2);
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.BAR).addRange(rangoDVotos)
      .setOption("title","Oportunidades de mejoramiento más votadas (número) — "+tituloBloque)
      .setOption("width",380).setOption("height",Math.max(220,40+tD.length*22))
      .setPosition(filaTablas,19,0,0).build());
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(filaTablas,5,tD.length+1,1))
      .addRange(sh.getRange(filaTablas,7,tD.length+1,1))
      .setOption("title","Oportunidades de mejoramiento más votadas (% de asistentes) — "+tituloBloque)
      .setOption("width",380).setOption("height",Math.max(220,40+tD.length*22))
      .setPosition(filaTablas,24,0,0).build());
  }

  const filasTablaMax=Math.max(tF.length,tD.length)+1;
  const filasChartMax=Math.ceil(Math.max(220,40+Math.max(tF.length,tD.length)*22)/21);
  return filaTablas+Math.max(filasTablaMax,filasChartMax)+3;
}

function construirHojaPercepcionGrupoFEM_(ss, grupos){
  const sh=hojaLimpiaGrupoFEM_(ss, "Percepción por Grupo");
  sh.getRange(1,1).setValue("PERCEPCIÓN DE FORTALEZAS Y OPORTUNIDADES DE MEJORAMIENTO — FEM 2026");
  sh.getRange(1,1).setFontWeight("bold").setFontSize(13);

  let filaActual=3;

  // Bloque municipal: TODAS las IE reales combinadas (a partir de
  // AsistenciaQR completa), a pedido expreso — número Y porcentaje.
  let asistentesMunicipio=[];
  GRUPOS_FEM_ORDEN_.concat((grupos.SIN_GRUPO||[]).length?["SIN_GRUPO"]:[]).forEach(function(g){
    (grupos[g]||[]).forEach(function(m){
      try{ asistentesMunicipio=asistentesMunicipio.concat(obtenerAsistentesQR_(m.idForo)); }
      catch(error){ Logger.log("Asistentes QR de "+m.nombreIE+" (municipal): "+error.message); }
    });
  });
  filaActual=escribirBloquePercepcionFEM_(sh, filaActual, "TODAS LAS IE (percepción general municipal, "+asistentesMunicipio.length+" asistentes)", asistentesMunicipio);
  filaActual+=1;

  GRUPOS_FEM_ORDEN_.forEach(function(g){
    const miembros=grupos[g]||[];
    let asistentes=[];
    miembros.forEach(function(m){
      try{ asistentes=asistentes.concat(obtenerAsistentesQR_(m.idForo)); }
      catch(error){ Logger.log("Asistentes QR de "+m.nombreIE+" ("+g+"): "+error.message); }
    });
    filaActual=escribirBloquePercepcionFEM_(sh, filaActual, etiquetaGrupoFEM_(g)+" — Percepción combinada ("+asistentes.length+" asistentes de "+miembros.length+" IE)", asistentes);
  });

  sh.autoResizeColumns(1,8);
  return sh;
}

/*
 * Orquesta las tres hojas de arriba en una sola hoja de cálculo.
 * Vuelve a ejecutarse las veces que haga falta: cada vez recalcula
 * las tres hojas desde cero con los datos más recientes (no va
 * acumulando versiones viejas).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". La URL de la hoja de cálculo queda en "Ver
 * registros de ejecución".
 */
function generarSpreadsheetGrupoFEM(){
  const grupos=mapaGruposFEM_();
  const ss=obtenerSpreadsheetGruposFEM_();

  construirHojaParticipacionGrupoFEM_(ss, grupos);
  construirHojaValoracionGrupoFEM_(ss, grupos);
  construirHojaPercepcionGrupoFEM_(ss, grupos);

  // La primera hoja que trae toda hoja de cálculo nueva de Apps
  // Script ("Hoja 1"/"Sheet1") ya no hace falta una vez existen las
  // tres de arriba.
  const hojaPorDefecto=ss.getSheetByName("Hoja 1")||ss.getSheetByName("Sheet1");
  if(hojaPorDefecto && ss.getSheets().length>1) ss.deleteSheet(hojaPorDefecto);

  Logger.log("========================================");
  Logger.log("HOJA DE PARTICIPACIÓN, VALORACIÓN Y PERCEPCIÓN POR GRUPO");
  Logger.log(ss.getUrl());
  if((grupos.SIN_GRUPO||[]).length){
    Logger.log("IE SIN GRUPO asignado incluidas aparte — revisar manualmente: "+JSON.stringify(grupos.SIN_GRUPO.map(function(m){return m.nombreIE;})));
  }
  Logger.log("========================================");

  return {url:ss.getUrl(), sinGrupo:(grupos.SIN_GRUPO||[]).map(function(m){return m.nombreIE;})};
}


/*****************************************************
 * RECUPERACIÓN DE ASISTENCIA QR — INCIDENTE 2026-09-05
 *
 * La hoja "AsistenciaQR" del spreadsheet principal apareció
 * completamente vacía (solo el encabezado, cero filas) pese a que 4
 * instituciones sí habían registrado asistencia real por código QR:
 * EL LIMONAR (101), LICEO DE SANTA LIBRADA (68), ATANASIO GIRARDOT
 * (64) y RODRIGO LARA BONILLA (52) — total 285 filas. Se confirmó
 * comparando contra la copia de seguridad que queda automáticamente
 * en el documento "Análisis FEM 2026" (hoja propia de cada IE,
 * bloque "ASISTENCIA QR — {IE}", escrito por
 * actualizarAnalisisFEMIndividual_ en Código.js cada vez que esa IE
 * envía su Foro o su Valoración) — las otras IE reales muestran 0
 * también en esa copia de seguridad, así que para ellas no hay nada
 * que restaurar (nunca tuvieron asistencia QR real registrada, no es
 * una pérdida de datos).
 *
 * LIMITACIÓN CONOCIDA: esa copia de seguridad NO incluye las columnas
 * FORTALEZAS ni DIFICULTADES (el bloque "ASISTENCIA QR" de Análisis
 * FEM solo guarda nombre/sexo/edad/cargo/rol/jornada/sede/documento/
 * correo/teléfono/fecha/hora) — esas dos columnas quedan en blanco en
 * las filas restauradas porque no existe ningún respaldo con esa
 * información al que se tenga acceso. CONSENTIMIENTO se restaura como
 * "Sí" porque el formulario de asistencia QR nunca permite enviar sin
 * aceptarlo — no es una suposición, es un invariante del formulario
 * (ver TEXTO_CONSENTIMIENTO_ASISTENCIA_QR y registrarAsistenciaQR en
 * Código.js, que siempre escribe CONSENTIMIENTO:"Sí").
 *
 * Idempotente: si una fila con el mismo ID_FORO+NUMERO_DOCUMENTO ya
 * existe en la hoja (por ejemplo porque esta función ya se ejecutó
 * antes), no se duplica.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El resultado (cuántas filas se restauraron
 * por IE) queda en "Ver registros de ejecución".
 *****************************************************/
const DATOS_RECUPERADOS_ASISTENCIA_QR_ = JSON.parse('[{"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Juan Sebastian Rojas Guerrero", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1075294195", "correo": "rojassebastian285@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:11:46"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Elvis Andrés Mantilla", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "7716053", "correo": "elanma31@yahoo.com.co", "telefono": "3214773081", "fecha": "28/08/2026", "hora": "11:11:50"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Oscar Fernando Penagos Rojas", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1075269821", "correo": "penagos.oscar@ielimonar.edu.co", "telefono": "3123231459", "fecha": "28/08/2026", "hora": "11:11:56"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Yesid covaleda Gómez", "sexo": "Masculino", "edad": "18-25", "tipoAsistencia": "Presencial", "cargo": "Otro", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "1077226379", "correo": "yesidcovaledagomez@gmail.com", "telefono": "3053474507", "fecha": "28/08/2026", "hora": "11:12:00"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Rudy Leidy Montenegro Sanchez", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1117513715", "correo": "rudylight-0607@hotmail.com", "telefono": "3132752068", "fecha": "28/08/2026", "hora": "11:12:24"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Salomé Gutiérrez Quintero", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1076909452", "correo": "salomequintero2315@gmail.com", "telefono": "3168457200", "fecha": "28/08/2026", "hora": "11:12:28"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Wilder Duban Herrera Torres", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075260612", "correo": "dutoher0318@gmail.commail.com", "telefono": "3138728399", "fecha": "28/08/2026", "hora": "11:12:42"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "MARIA FERNANDA MEDINA PASAJE", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "1112458531", "correo": "Mafecilla116@hotmail.com", "telefono": "3166270607", "fecha": "28/08/2026", "hora": "11:13:05"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Diana Carolina sastoque", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Garabaticos", "documento": "52915045", "correo": "evecarodiana@hotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:13:46"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Raily Juliette Bolivar Castañeda", "sexo": "Femenino", "edad": "no_responde", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Lomalinda", "documento": "1075250433", "correo": "raily.9@hotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:13:50"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "OSCAR ANDRÉS PERDOMO ROJAS", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Tutor(a) PTA/PFI 3.0", "rolForo": "🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0", "jornada": "Única", "sede": "Central/Administrativa", "documento": "7719324", "correo": "osanpero@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:13:55"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Mateo Trujillo hernandez", "sexo": "Masculino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1077233338", "correo": "mateotrujillohernandez375@gmail.com", "telefono": "3204149008", "fecha": "28/08/2026", "hora": "11:13:56"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "José Danluguer Casallas Calderón", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "83228566", "correo": "danluguercasallas@gmail.com", "telefono": "3202290602", "fecha": "28/08/2026", "hora": "11:13:58"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Karen Lizeth Camacho Lara", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1029882281", "correo": "karenlizethcamacholara@gmail.com", "telefono": "3142787560", "fecha": "28/08/2026", "hora": "11:14:07"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Zulma Yineth Baquero Torres", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "36306772", "correo": "docentezulmabaquero2021@gmail.com", "telefono": "3142930454", "fecha": "28/08/2026", "hora": "11:14:10"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Lina Marcela Castaño Cuellar", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "40612569", "correo": "linishmarcecc@gmail.com", "telefono": "3156611644", "fecha": "28/08/2026", "hora": "11:14:12"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "SUSANA SANDOVAL MORENO", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Lomalinda", "documento": "36307077", "correo": "susan54072@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:14:20"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Sandra Lorena Polanco Perdomo", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36312117", "correo": "lorenapolanco_522@hotmail.com", "telefono": "3107938412", "fecha": "28/08/2026", "hora": "11:14:21"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Jorge Andres Herrera Dussan", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Lomalinda", "documento": "1075211950", "correo": "jorreradu@hotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:14:23"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "María Fernández Cuenca Barrero", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36313128", "correo": "mafecuenca@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:14:31"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Olga Lyda Rodríguez Narváez", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "55166549", "correo": "olgalynarvaez877@gmail.com", "telefono": "3162677936", "fecha": "28/08/2026", "hora": "11:14:40"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Derly Constanza Zamudio Trujillo", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Tarde", "sede": "Lomalinda", "documento": "1075224151", "correo": "derly.czt@gmail.commail.com", "telefono": "3125868587", "fecha": "28/08/2026", "hora": "11:14:42"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "DIEGO ARMANDO MORALES MOSQUERA", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "7731104", "correo": "dam_856@hotmail.com", "telefono": "3185201784", "fecha": "28/08/2026", "hora": "11:14:58"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Martha Yineth Bastidas Cangrejo", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Garabaticos", "documento": "55170345", "correo": "mayito.b.c@gmail.com", "telefono": "3166259505", "fecha": "28/08/2026", "hora": "11:14:59"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Kevin Santiago González Peña", "sexo": "Masculino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075252811", "correo": "santiagogonzalez121317@gmail.com", "telefono": "3204971882", "fecha": "28/08/2026", "hora": "11:15:02"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "MARIA EUGENIA", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "55164946", "correo": "maepan2012@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:15:11"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Judy Andrea Hernández Lucuara", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Lomalinda", "documento": "36309418", "correo": "judyandy27@hotmail.com", "telefono": "3144665303", "fecha": "28/08/2026", "hora": "11:15:14"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Hernanado Cardona Salazar", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "7694000", "correo": "hernando.cardona.salazar@gmail.com", "telefono": "3223353913", "fecha": "28/08/2026", "hora": "11:15:15"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Stefani Albarracin Gallego", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "1075216864", "correo": "stefani.albarracin86@gmail.com", "telefono": "3168643991", "fecha": "28/08/2026", "hora": "11:15:20"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "LILIANA BAHAMON ALVAREZ", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "55164073", "correo": "lilianabahamon@hotmail.comotmail.com", "telefono": "3152415161", "fecha": "28/08/2026", "hora": "11:15:26"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Oliver Guillermo Andrade Pérez", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1075258266", "correo": "oliverandradeperez@gmail.com", "telefono": "3108621303", "fecha": "28/08/2026", "hora": "11:15:28"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Arthur Tovar Cruz", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075264883", "correo": "arthur.tovar@usco.edu.co", "telefono": "3165204782", "fecha": "28/08/2026", "hora": "11:15:31"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Pablo Emilio Castillo Quiroga", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Lomalinda", "documento": "80020239", "correo": "pecqdocencia@gmail.com", "telefono": "3137780054", "fecha": "28/08/2026", "hora": "11:15:44"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Veronica Alfonso Gamboa", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "26427372", "correo": "vealgam84@gmail.com", "telefono": "3152855909", "fecha": "28/08/2026", "hora": "11:15:46"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "ANDREA PAOLA CASTRO MOSSOS", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "33751009", "correo": "andrepao1727@hotmail.com", "telefono": "3133265982", "fecha": "28/08/2026", "hora": "11:15:49"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "gabriela triviño ceron", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36295812", "correo": "gbisbri@hotmail.com", "telefono": "3187902204", "fecha": "28/08/2026", "hora": "11:15:50"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Nicolás Baquero Torres", "sexo": "Masculino", "edad": "18-25", "tipoAsistencia": "Presencial", "cargo": "Otro", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "1003866024", "correo": "nicolasbatorres4@gmail.com", "telefono": "3206334197", "fecha": "28/08/2026", "hora": "11:15:54"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Clara Eugenia Pastrana Cuenca", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Lomalinda", "documento": "36178169", "correo": "clarita4851@gmail.com", "telefono": "3123709036", "fecha": "28/08/2026", "hora": "11:15:56"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "JESUS HERNANDO MAZORRA MUÑOZ", "sexo": "Masculino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "4935776", "correo": "hernandomazorra@gmail.com", "telefono": "3138170337", "fecha": "28/08/2026", "hora": "11:16:10"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Ana Violeth Quintero Polanco", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36342599", "correo": "anaquinpol85@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:16:13"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Derly Niyireth Salas Arce", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "55158478", "correo": "sderly76@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:16:25"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "María Esperanza Torres suaza", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075236608", "correo": "maria-e.torres@hotmail.com", "telefono": "3164212971", "fecha": "28/08/2026", "hora": "11:16:39"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Alba Ramírez Vega", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36170148", "correo": "albaramirezv71@gmail.com", "telefono": "3103303929", "fecha": "28/08/2026", "hora": "11:16:40"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Yolima García Monje", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Lomalinda", "documento": "55168317", "correo": "yolimagarciamonje@gmail.com", "telefono": "3163033961", "fecha": "28/08/2026", "hora": "11:16:46"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Christopher", "sexo": "Masculino", "edad": "0-12", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1073919304", "correo": "salomequintero2315@gmail.com", "telefono": "3168457200", "fecha": "28/08/2026", "hora": "11:16:48"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "HECTOR EDMUNDO PANTOJA MUÑOZ", "sexo": "Masculino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "87025143", "correo": "pantojaedmundo@gmail.com", "telefono": "3112629058", "fecha": "28/08/2026", "hora": "11:17:02"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Maria Arlet Rojas Puentes", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Lomalinda", "documento": "36087409", "correo": "arletrojas66@gmail.com", "telefono": "3124488182", "fecha": "28/08/2026", "hora": "11:17:13"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Lucía Moreno M", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36182720", "correo": "morenomedinalucia@gmail.com", "telefono": "3167475181", "fecha": "28/08/2026", "hora": "11:17:20"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Consuelo Fernández Tovar", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Garabaticos", "documento": "36173294", "correo": "connyfer2011@hotmail.comotmail.com", "telefono": "3187747062", "fecha": "28/08/2026", "hora": "11:17:23"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Rina Polania Ninco", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Tarde", "sede": "Garabaticos", "documento": "1075209074", "correo": "rinadianeth0985@gmail.com", "telefono": "3142850896", "fecha": "28/08/2026", "hora": "11:17:31"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Tránsito Polonia Benítez", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "36179085", "correo": "transitopolaniabenitez@gmail.com", "telefono": "315229476", "fecha": "28/08/2026", "hora": "11:17:39"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "claudia Inés Quiroga Perdomo", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "26425628", "correo": "clauquity@gmail.com", "telefono": "3164739542", "fecha": "28/08/2026", "hora": "11:17:50"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Stella Gonzalez Moreno", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36175711", "correo": "stellagonza2026@gmail.com", "telefono": "3212526872", "fecha": "28/08/2026", "hora": "11:17:55"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Deymar Santiago Caicedo", "sexo": "Masculino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1109677221", "correo": "deymarcaicedo78@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:17:57"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Juan Andrés Bahamon Martínez", "sexo": "Masculino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075798209", "correo": "jabm2009grr@gmail.commail.com", "telefono": "3213223203", "fecha": "28/08/2026", "hora": "11:17:59"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "María Eugenia Bello Lara", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "65740895", "correo": "bellolaramariaeugenia3@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:18:02"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Isabella Cortes Henao", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "📝 Relator(a)", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1029887550", "correo": "isabellacorteshenao5@gmail.com", "telefono": "3227154979", "fecha": "28/08/2026", "hora": "11:18:05"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Nicolás Herrera Cabrera", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "7711285", "correo": "herreracabreranicolas@gmail.com", "telefono": "3214716381", "fecha": "28/08/2026", "hora": "11:18:09"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "DORIS GUEVARA PERDOMO", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "36158518", "correo": "gdoris3@gmail.com", "telefono": "3205700289", "fecha": "28/08/2026", "hora": "11:18:11"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Constanza Losada Quintero", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Garabaticos", "documento": "55145249", "correo": "conylosadaquintero@gmail.commail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:18:13"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Luz Dary Rojas Guzman", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36169422", "correo": "luzdaryrojasguzman2219@gmail.com", "telefono": "3176720579", "fecha": "28/08/2026", "hora": "11:18:14"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Marta Milena Castillo Calderón", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Padre/madre/acudiente", "rolForo": "🙋 Participante", "jornada": "", "sede": "", "documento": "26433762", "correo": "martacastillo792@gmail.com", "telefono": "3115733666", "fecha": "28/08/2026", "hora": "11:18:16"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Mayra Fernanda Delgado Quimbaya", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Lomalinda", "documento": "1075259200", "correo": "mayrafernandadelgado@gmail.commail.com", "telefono": "32092617", "fecha": "28/08/2026", "hora": "11:18:18"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Verselio Lozano", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "7691139", "correo": "versel72@hotmail.comotmIl.com", "telefono": "3115850382", "fecha": "28/08/2026", "hora": "11:18:46"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Lina María Coronado", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1080295572", "correo": "linacoronadosilva95@gmail.com", "telefono": "3233228472", "fecha": "28/08/2026", "hora": "11:19:16"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Sandra Milena Cruz Collazos", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Lomalinda", "documento": "36066562", "correo": "smcc1995@hotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:19:41"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "María del Rosario Trujillo Medina", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "36182644", "correo": "rosario-trujillo@hotmail.com", "telefono": "3166234938", "fecha": "28/08/2026", "hora": "11:20:06"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Diego Andrés García Barreiro", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Padre/madre/acudiente", "rolForo": "🙋 Participante", "jornada": "", "sede": "", "documento": "7722154", "correo": "diegoagb354@gmail.com", "telefono": "3132465678", "fecha": "28/08/2026", "hora": "11:20:09"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Liliana Constanza Yustre Cabrera", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "55155915", "correo": "naniliyus@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:20:12"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Aydee Mora Bernal", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "26458599", "correo": "aymober123@gmail.commail.com", "telefono": "3142984435", "fecha": "28/08/2026", "hora": "11:20:48"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Arcelia vargas quintero", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Garabaticos", "documento": "36165961", "correo": "vargasa988@gmail.com.com", "telefono": "3173697520", "fecha": "28/08/2026", "hora": "11:21:06"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Lorena Urriaga Hoyos", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Garabaticos", "documento": "36309153", "correo": "profelorena16@hotmail.com", "telefono": "3167206209", "fecha": "28/08/2026", "hora": "11:21:08"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Maria Francia Ordonez Claros", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "36279591", "correo": "mafranor68@hotmail.com", "telefono": "3204662163", "fecha": "28/08/2026", "hora": "11:21:22"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Martha Patricia Sanchez Marin", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "55152593", "correo": "patricia300034@hotmail.com", "telefono": "3132839139", "fecha": "28/08/2026", "hora": "11:21:41"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Lucero Méndez Murcia", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "55174434", "correo": "lucepeju@hotmail.com", "telefono": "3174948808", "fecha": "28/08/2026", "hora": "11:21:55"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "JACQUELINE BAHAMON MONTENEGRO", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36301772", "correo": "jacque.bahamon@gmail.com", "telefono": "3184732980", "fecha": "28/08/2026", "hora": "11:21:58"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Alicia Claros Rojas", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Lomalinda", "documento": "26452593", "correo": "aliciaclarosrojas@gmail.commal.com", "telefono": "3118311077", "fecha": "28/08/2026", "hora": "11:22:27"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Erika Alexandra Vargas Gutiérrez", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Orientador(a)", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075224698", "correo": "aner1205@gmail.com", "telefono": "3153941622", "fecha": "28/08/2026", "hora": "11:22:48"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Martha Elena Azuero Paredes", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Buenos Aires", "documento": "36164137", "correo": "marthae150@hotmail.comotmsil.com", "telefono": "3158232833", "fecha": "28/08/2026", "hora": "11:23:16"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Jorge Enrique Mendoza Soto", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "93353703", "correo": "jorgemendozan50@hotmail.com", "telefono": "3212077880", "fecha": "28/08/2026", "hora": "11:23:20"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Diana Consuelo Sanmiguel ortiz", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "55163743", "correo": "dianasan145@hotmail.com", "telefono": "3202350437", "fecha": "28/08/2026", "hora": "11:24:06"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Luz Marina Palomino Suarez", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Lomalinda", "documento": "26500701", "correo": "luzmarinapalomino1@gmail.commail.com", "telefono": "3166542857", "fecha": "28/08/2026", "hora": "11:24:11"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Gladys Elena Dussán Díaz", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "55168327", "correo": "gladydus@gmail.com", "telefono": "3167509600", "fecha": "28/08/2026", "hora": "11:24:23"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Martha Lucía Pérez Manrique", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36178818", "correo": "malupema01@gmail.com", "telefono": "3163972587", "fecha": "28/08/2026", "hora": "11:24:47"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Jorge Bravo González", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "7715424", "correo": "jbravo723@hotmail.com", "telefono": "3115336600", "fecha": "28/08/2026", "hora": "11:25:19"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Nancy Sánchez González", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Padre/madre/acudiente", "rolForo": "🙋 Participante", "jornada": "", "sede": "", "documento": "36178703", "correo": "sanchezgonzaleznancy0@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:26:33"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "María Rita Dussan Quiza", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "36162828", "correo": "mariadussan08@gmail.com", "telefono": "3188712554", "fecha": "28/08/2026", "hora": "11:26:34"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Yolanda Arce Quibano", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36167153", "correo": "yolanda.emmita.@hotmail.com", "telefono": "3105293597", "fecha": "28/08/2026", "hora": "11:26:48"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Mercedes Moyano Otalora", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "361625213", "correo": "mercedes.moyano486@gmail.com", "telefono": "3114750713", "fecha": "28/08/2026", "hora": "11:27:18"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Consuelo Cordón Herrera", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Garabaticos", "documento": "36165681", "correo": "consuelo-CordónHerrera@hotmail.com", "telefono": "3125035202", "fecha": "28/08/2026", "hora": "11:27:22"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Ana Yaneth Aldana", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Lomalinda", "documento": "26427967", "correo": "anita47aldana@gmail.com", "telefono": "3138100025", "fecha": "28/08/2026", "hora": "11:27:45"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "María Isabel Farfan Hernández", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Padre/madre/acudiente", "rolForo": "🙋 Participante", "jornada": "", "sede": "", "documento": "36068749", "correo": "mariaisabelfarfanhernandez@gmail.com", "telefono": "3158632109", "fecha": "28/08/2026", "hora": "11:28:35"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Alvaro Alveiro Romero Almario", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "79401850", "correo": "alve3338@hotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:28:54"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Diana Maria Reyes Reyes", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "26430291", "correo": "dsj3104@hotmail.com", "telefono": "3167864975", "fecha": "28/08/2026", "hora": "11:29:14"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Ligia Lievano Hernandez", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Buenos Aires", "documento": "26501242", "correo": "ligialievano.74@gmail.com", "telefono": "33154606341", "fecha": "28/08/2026", "hora": "11:29:36"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Carlos ramon Canencio Ramirez", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "4935811", "correo": "caracara61@gmail.comotmoil.com", "telefono": "3015184858", "fecha": "28/08/2026", "hora": "11:30:37"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Diana Marcela Guzmán Quimbayo", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Padre/madre/acudiente", "rolForo": "🙋 Participante", "jornada": "", "sede": "", "documento": "1081153467", "correo": "aner723@hotmail.com", "telefono": "3202421910", "fecha": "28/08/2026", "hora": "11:32:11"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Henry William Uni Yugcha", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Rector(a)", "rolForo": "👑 Líder – Rector(a)", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "7698786", "correo": "ielimonar@alcaldianeiva.gov.co", "telefono": "3123797087", "fecha": "28/08/2026", "hora": "11:34:00"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Leidy Alexandra Narvaez", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Garabaticos", "documento": "26428223", "correo": "leidyalexandranarvaez@gmail.com", "telefono": "3176592036", "fecha": "28/08/2026", "hora": "11:35:40"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Eder Valderrama", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "7719139", "correo": "edervb81@hotmail.com", "telefono": "3204703111", "fecha": "28/08/2026", "hora": "11:44:59"}, {"idForo": "0ef507cd-0d33-459c-b9f8-e9a03a5241bb", "ie": "EL LIMONAR", "nombre": "Nury Andrade Quimbaya", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Garabaticos", "documento": "55153849", "correo": "nuanqui@gmail.com", "telefono": "3202680875", "fecha": "28/08/2026", "hora": "11:46:01"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Rosa Maria Suarez", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36176709", "correo": "rositasuarezliceo@gmail.com", "telefono": "3136461331", "fecha": "28/08/2026", "hora": "10:44:20"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Jhon Efraín Sanchez Bolaños", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "No asistió: con permiso de comisión o con acto administrativo.", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1083873534", "correo": "jhonefrainsanchez@gmail.com", "telefono": "3184561081", "fecha": "28/08/2026", "hora": "10:47:07"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Amanda Saavedra Perdomo", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36177855", "correo": "amsape63@gmail.com", "telefono": "3172691148", "fecha": "28/08/2026", "hora": "10:49:16"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Carlos Fernando Cardoso", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "17647475", "correo": "carfenando1971@gmail.com", "telefono": "3102856863", "fecha": "28/08/2026", "hora": "10:50:55"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Ana Milé Sabogal Garzón", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1030549073", "correo": "milesabogalgarzon@gmail.com", "telefono": "3002345866", "fecha": "28/08/2026", "hora": "10:52:08"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Luz Teresa Calderón Pacheco", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "55174860", "correo": "luztere75@gmail.com", "telefono": "3183701639", "fecha": "28/08/2026", "hora": "10:54:12"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "kathleen Martinez Ramirez", "sexo": "Prefiero no decirlo", "edad": "no_responde", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "52787840", "correo": "Kathleenmartinez@liceodesantalibrada.edu.co", "telefono": "3182691363", "fecha": "28/08/2026", "hora": "10:54:17"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "FRANCIA GARCIA MONJE", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36184082", "correo": "franciagarcia@liceodesantalibrada.comedu.co", "telefono": "", "fecha": "28/08/2026", "hora": "10:54:33"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Lizeth Caterine Trujillo Fierro", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1075257201", "correo": "lizethtrujillo@liceodesantalibrada.edu.co", "telefono": "3123702407", "fecha": "28/08/2026", "hora": "10:55:11"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "NORMA CONSTANZA BASTO SALAS", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "55180042", "correo": "normacons77@gmail.com", "telefono": "3123254994", "fecha": "28/08/2026", "hora": "10:56:20"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Bibianey Lozano Losada", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36177779", "correo": "bibiloza@hotmail.com", "telefono": "3204570798", "fecha": "28/08/2026", "hora": "10:56:50"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Xiomara Natalia García Cedeño", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1075291003", "correo": "natawi85@gmail.com", "telefono": "3172815400", "fecha": "28/08/2026", "hora": "10:56:53"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Marly Hellen Silva Rojas", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "55178158", "correo": "marlyhellen@gmail.com", "telefono": "3118646153", "fecha": "28/08/2026", "hora": "10:59:05"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Maritza Cuadrado Peraza", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "26431203", "correo": "maritzacuadrado@gmail.com", "telefono": "31640402857", "fecha": "28/08/2026", "hora": "10:59:53"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Jorge Mauricio Durán Basto", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "7693211", "correo": "jorgeduran@liceodesantalibrada.comedu.co", "telefono": "", "fecha": "28/08/2026", "hora": "11:00:57"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Gabriela Isabel González Fernández", "sexo": "Femenino", "edad": "18-25", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1029881949", "correo": "gabrielagf123456@gmail.com", "telefono": "3223522307", "fecha": "28/08/2026", "hora": "11:01:14"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Carolina silva Murcia", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1081157404", "correo": "silvamurciacarolina26@gmail.com", "telefono": "3107937088", "fecha": "28/08/2026", "hora": "11:01:51"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Zammara Tovar Medina", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1077231701", "correo": "zammaratovarm@liceodesantalibrada.edu.co", "telefono": "3223900621", "fecha": "28/08/2026", "hora": "11:03:19"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Danna Valentina Rodríguez Villarreal", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1077726394", "correo": "dannavr211@gmail.com", "telefono": "3106974583", "fecha": "28/08/2026", "hora": "11:03:35"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Luisa Fernanda Pedraza Chavarro", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1076908492", "correo": "luisapedraza.ch@gmail.com", "telefono": "3209709628", "fecha": "28/08/2026", "hora": "11:03:53"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Adriana María Victoria", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36310324", "correo": "adrianamvictoria@gmail.com", "telefono": "3134963710", "fecha": "28/08/2026", "hora": "11:04:02"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Sandra Patricia caballero", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Padre/madre/acudiente", "rolForo": "🙋 Participante", "jornada": "", "sede": "", "documento": "65718089", "correo": "patico.verde@hotmail.com", "telefono": "3214311006", "fecha": "28/08/2026", "hora": "11:05:48"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Adriana Gerena Rivas", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1082214281", "correo": "adrianagerena@liceodesantalibrada.edu.co", "telefono": "", "fecha": "28/08/2026", "hora": "11:05:54"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "María Goretti Azuero Bernal", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36172526", "correo": "mariaazuero@liceodesantalibrada.comcom", "telefono": "3107806142", "fecha": "28/08/2026", "hora": "11:06:31"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "SARA MAHALIA MOLINA TAPIERO", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36310571", "correo": "samamota@gmail.com", "telefono": "3118401318", "fecha": "28/08/2026", "hora": "11:06:55"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Javier Salas Ramírez", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "El Triangulo", "documento": "7728553", "correo": "javiersalas1608@gmail.com", "telefono": "3134687258", "fecha": "28/08/2026", "hora": "11:07:18"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "JULIO EFRAÍN DELGADO BRAVO", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "5209756", "correo": "julefra0811@gmail.com", "telefono": "3105853504", "fecha": "28/08/2026", "hora": "11:07:49"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "MYRIAM CUELLAR", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "El Triangulo", "documento": "36166771", "correo": "cumy618@hotmail.com", "telefono": "3204948868", "fecha": "28/08/2026", "hora": "11:08:18"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Julio Cesar Ayala Plazas", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1075217786", "correo": "julces86@gmail.com", "telefono": "3156264455", "fecha": "28/08/2026", "hora": "11:08:39"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Mayuri suarez gonzalez", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Mañana", "sede": "El Triangulo", "documento": "36305672", "correo": "suarezgonzalezmayuri84@gmail.com", "telefono": "3125655550", "fecha": "28/08/2026", "hora": "11:08:40"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Luz Mary Plazas", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "El Triangulo", "documento": "55163292", "correo": "luzmarypf@gmail.commail.com", "telefono": "3208088907", "fecha": "28/08/2026", "hora": "11:09:32"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Sandra Milena Pinto González", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075254568", "correo": "samypinto1991@gmail.com", "telefono": "3137080356", "fecha": "28/08/2026", "hora": "11:09:34"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Mayuri Yucumá Vanegas", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36307000", "correo": "mayuriyv@gmail.commail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:10:08"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Antonella Gómez Guzmán", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1076911297", "correo": "2210antonellagomez@gmail.com", "telefono": "3143741686", "fecha": "28/08/2026", "hora": "11:10:14"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Ever Oviedo Murcia", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075219694", "correo": "eoviedom17@gmail.com", "telefono": "3028613331", "fecha": "28/08/2026", "hora": "11:10:23"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Luis yeris celis toledo", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "7719197", "correo": "lyeris@hotmail.com", "telefono": "3118330499", "fecha": "28/08/2026", "hora": "11:10:39"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Gina Katherine Otálora Moreno", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "1049604262", "correo": "ginita33@gmail.com", "telefono": "3212816121", "fecha": "28/08/2026", "hora": "11:11:20"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Saturia Escobar Quimbaya", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36167810", "correo": "saturiaescobar@liceodesantalibrafa.comedu.co", "telefono": "3124365158", "fecha": "28/08/2026", "hora": "11:11:57"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Edwin Duarte Vidal", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "7723832", "correo": "eduvi83@hotmail.comotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:11:58"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "jorge mendezmurcia", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "12128318", "correo": "jorgemendezmurcia@gmail.com", "telefono": "3214533881", "fecha": "28/08/2026", "hora": "11:12:04"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Doris Yaneth Vargas Torrejano", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "El Triangulo", "documento": "36178579", "correo": "yaneca36@hotmail.com", "telefono": "3153666872", "fecha": "28/08/2026", "hora": "11:12:32"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Karen Johanna Marlés Quintero", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36312096", "correo": "karen.marles83@gmail.com", "telefono": "3212949103", "fecha": "28/08/2026", "hora": "11:13:41"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Giovanna Salazar", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "55163049", "correo": "giovannasalazar@liceodesantalibrada.edu.co", "telefono": "3178863552", "fecha": "28/08/2026", "hora": "11:14:29"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Maria Enith Bonilla Ramìrez-", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36179867", "correo": "maenbor65@gmail.com", "telefono": "3168677149", "fecha": "28/08/2026", "hora": "11:14:43"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Oscar Ivan no", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "7730433", "correo": "oskkar01@hotmail.com", "telefono": "3187462460", "fecha": "28/08/2026", "hora": "11:15:06"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Carlos Uni Yugcha", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "7686529", "correo": "unicarl@gmail.com", "telefono": "3108018780", "fecha": "28/08/2026", "hora": "11:15:23"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Marlio Saavedra Perdomo", "sexo": "Masculino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "12108519", "correo": "marliosaavedraperdomo@gmail.com", "telefono": "3158353792", "fecha": "28/08/2026", "hora": "11:15:59"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Diana Rocío Vargas Bermeo", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "55165412", "correo": "dianarvargasb@gmail.com", "telefono": "3208396229", "fecha": "28/08/2026", "hora": "11:17:06"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Diana patricia Córdoba salazar", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1077460169", "correo": "dianapatriciacordobasalazar@gmail.com", "telefono": "3123289483", "fecha": "28/08/2026", "hora": "11:17:15"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Yanit Vergel Villarreal", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36184076", "correo": "yanitvergel@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:17:34"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "CLARA INES AHUMADA BUSTOS", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36177400", "correo": "clahum12@gmail.com", "telefono": "3157902279", "fecha": "28/08/2026", "hora": "11:17:41"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Luz Myriam Artunduaga Yunda", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36175777", "correo": "luzmyriam113@gmail.com", "telefono": "3174308794", "fecha": "28/08/2026", "hora": "11:18:07"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Eulalia Parra Torres", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36177885", "correo": "eulalia.parra@hotmail.com", "telefono": "3118557439", "fecha": "28/08/2026", "hora": "11:18:43"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "ema isabel beltran guijo", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1104946764", "correo": "emaisabel.beltran18@gmail.com", "telefono": "3134740527", "fecha": "28/08/2026", "hora": "11:18:54"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Néstor Hernando Rocha Suarez", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "93381184", "correo": "blackmanice2204@gmail.com", "telefono": "3166668537", "fecha": "28/08/2026", "hora": "11:18:57"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Leidy Mariana Sierra Ardila", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1077233519", "correo": "leidymarianasierra39@gmail.com", "telefono": "3214251359", "fecha": "28/08/2026", "hora": "11:19:10"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Beatriz Hurtado Cabrera", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36167705", "correo": "beatrizhurtadoloceodesantalibrada@edu.cogmail.com", "telefono": "3202560831", "fecha": "28/08/2026", "hora": "11:19:53"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "ALBA LIGIA CANO PENAGOS", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "26597636", "correo": "albacano@liceodesantalibrada.edu.co", "telefono": "3176748194", "fecha": "28/08/2026", "hora": "11:20:03"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Yo", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "36171624", "correo": "socorrojaimesabella@hotmail.comoymail.com", "telefono": "3164317870", "fecha": "28/08/2026", "hora": "11:20:38"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Mireya Cabrera Adames", "sexo": "Femenino", "edad": "no_responde", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "55151375", "correo": "murallita_04@hotmail.com", "telefono": "3183624217", "fecha": "28/08/2026", "hora": "11:21:32"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Yesid Bermeo Ochoa", "sexo": "Masculino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "4881095", "correo": "yebero008@gmail.com", "telefono": "3123642942", "fecha": "28/08/2026", "hora": "11:22:41"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Estefanía Pérez Urriago", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1077859142", "correo": "estefyurriago2009@gmail.com", "telefono": "3227086871", "fecha": "28/08/2026", "hora": "11:23:18"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "leidy milena pastrana zambrano", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Tarde", "sede": "Central/Administrativa", "documento": "26431108", "correo": "leidy.pastrana@gmail.com", "telefono": "3157849052", "fecha": "28/08/2026", "hora": "11:25:58"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Natalia Díaz Giraldo", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1076907770", "correo": "1994nata2020@gmail.com", "telefono": "30434863848", "fecha": "28/08/2026", "hora": "11:27:39"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Ana Sofia Fajardo", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "3125603029", "correo": "anafajardo@liceodesantalibrada.edu.co", "telefono": "3125603029", "fecha": "28/08/2026", "hora": "11:28:59"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Hernando Cano Castro", "sexo": "Masculino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "12114660", "correo": "hercaca8@hotmail.com", "telefono": "3153231313", "fecha": "28/08/2026", "hora": "11:39:21"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Martha Lucia Palma Huergo", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "65743441", "correo": "marthapal13@gmail.com", "telefono": "3002197579", "fecha": "28/08/2026", "hora": "19:21:57"}, {"idForo": "36832fb9-13dc-4b9c-9a5c-a9306c6422be", "ie": "LICEO DE SANTA LIBRADA", "nombre": "Paul Alejandro Serpa Quintero", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Tutor(a) PTA/PFI 3.0", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1075222198", "correo": "alsequi26@gmail.com", "telefono": "", "fecha": "31/08/2026", "hora": "16:18:17"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Mario Morales", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "7719007", "correo": "amorales9007@gmail.com", "telefono": "3165304865", "fecha": "28/08/2026", "hora": "11:07:39"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Denis Guihomar Peña Santofimio", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "55155153", "correo": "deguipesatecno1067@hotmail.com", "telefono": "3143630448", "fecha": "28/08/2026", "hora": "11:07:40"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Martha González", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Loma De La Cruz", "documento": "36274474", "correo": "martha.gonzalez19581999@gmail.com", "telefono": "3184409245", "fecha": "28/08/2026", "hora": "11:08:24"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Martha Patricia Rivera Castro", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36174075", "correo": "marthicap63@gmail.com", "telefono": "3124460814", "fecha": "28/08/2026", "hora": "11:08:42"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Martha Cecilia Monje Alvarez", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Liceo Batallon Tenerife", "documento": "36168929", "correo": "martha.monje@hotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:08:50"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Alexander Paredes Martínez", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "1075237484", "correo": "alexanderp@ieatanasiog.com.edu.co", "telefono": "3154872236", "fecha": "28/08/2026", "hora": "11:08:59"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Pedro Francisco Morales García", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "7711484", "correo": "peframo@hotmail.com", "telefono": "3186473749", "fecha": "28/08/2026", "hora": "11:09:24"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Sandra Liliana Perdomo Naranjo", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Tarde", "sede": "Guillermo Montenegro", "documento": "36068374", "correo": "sanli_311@hotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:09:50"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Carolina Jiménez Lasso", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "55113532", "correo": "cajila22@hotmail.com", "telefono": "3163458391", "fecha": "28/08/2026", "hora": "11:09:59"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Eddy Trujillo", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Liceo Batallon Tenerife", "documento": "36158461", "correo": "eddytrujillo56@hotmail.com", "telefono": "3123005028", "fecha": "28/08/2026", "hora": "11:10:11"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Juan sebastian Trujillo Ospitia", "sexo": "Masculino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1083893697", "correo": "sebastrujilloyo@gmail.com", "telefono": "3212068265", "fecha": "28/08/2026", "hora": "11:10:20"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Arbey Burbano Vargas", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "12167583", "correo": "arbeymafis@hotmail.com", "telefono": "3115935075", "fecha": "28/08/2026", "hora": "11:10:26"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Sharik lorena Ordoñez Rios", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1029882394", "correo": "lr880959@gmail.com", "telefono": "3209300251", "fecha": "28/08/2026", "hora": "11:10:35"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Eurustides Montenegro", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "17672511", "correo": "teachemonte@hotmail.com", "telefono": "3232094554", "fecha": "28/08/2026", "hora": "11:10:44"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Carlos Alberto Yaime Ardila", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075246500", "correo": "yaime.poe@gmail.com", "telefono": "573143261761", "fecha": "28/08/2026", "hora": "11:10:48"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "NANCY ALDANA GUTIERREZ", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "26444275", "correo": "minancyprofe@gmail.com", "telefono": "3004829081", "fecha": "28/08/2026", "hora": "11:11:08"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Daniela Andrea Ruiz Suárez", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Liceo Batallon Tenerife", "documento": "1083913900", "correo": "dajucama@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:11:35"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Patricia Sanchez", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Orientador(a)", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36169669", "correo": "patriciasanchez48@hotmail.com", "telefono": "3153249852", "fecha": "28/08/2026", "hora": "11:11:40"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Rafael Hernandez", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "13354329", "correo": "rafaele@ieatanasiog.edu.co", "telefono": "3213860331", "fecha": "28/08/2026", "hora": "11:11:48"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Maria Isabel Gonzalez Ramos", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Personal administrativo", "rolForo": "📝 Relator(a)", "jornada": "", "sede": "", "documento": "36303760", "correo": "migonzalez@ieatanasiog.edu.co", "telefono": "3102232125", "fecha": "28/08/2026", "hora": "11:11:54"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Nancy Elvira  Castañeda Bermeo", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36173123", "correo": "nanelriviera163@gmail.com", "telefono": "3112444274", "fecha": "28/08/2026", "hora": "11:12:02"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Ingrid yulieth Garcia segura", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Tarde", "sede": "Guillermo Montenegro", "documento": "3105302932", "correo": "inyu.garse@gmail.com", "telefono": "3155357761", "fecha": "28/08/2026", "hora": "11:12:17"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Luis Eduardo Reyes Perdomo", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "1075209286", "correo": "serieprimo@hotmail.com", "telefono": "3182253548", "fecha": "28/08/2026", "hora": "11:12:27"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Constanza Castillo Aviles", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "55167696", "correo": "castillo.constanza3@gmail.com", "telefono": "3167457920", "fecha": "28/08/2026", "hora": "11:12:31"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Alerso Rojas Muñoz", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075249387", "correo": "alercresp@gmail.com", "telefono": "3223654604", "fecha": "28/08/2026", "hora": "11:12:33"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Gloria Edith Moreno Cabrera", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36376447", "correo": "gloriaem@atanasiog.edu.co", "telefono": "3118320184", "fecha": "28/08/2026", "hora": "11:12:57"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Elsa Cecilia Navarrete Tamayo", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Liceo Batallon Tenerife", "documento": "36165643", "correo": "cenata@hotmail.com", "telefono": "3112339587", "fecha": "28/08/2026", "hora": "11:13:02"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Eluana Perdomo Córdoba", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1029883831", "correo": "eluanaperdomo4@gmail.com", "telefono": "3150643870", "fecha": "28/08/2026", "hora": "11:13:07"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Laura Santana", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Tarde", "sede": "Guillermo Montenegro", "documento": "1075252989", "correo": "laurasan91@hotmail.com", "telefono": "3", "fecha": "28/08/2026", "hora": "11:13:43"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Ana María Bonilla", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "1075234071", "correo": "ana.maria.bonilla41@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:13:48"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Ángela Viviana Rodríguez Montero", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36069669", "correo": "anviromo282315@gmail.com", "telefono": "3138076238", "fecha": "28/08/2026", "hora": "11:13:52"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Raul Reyes Bahamon", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "12134911", "correo": "reyesbahamonr@gmail.com", "telefono": "3107872734", "fecha": "28/08/2026", "hora": "11:14:03"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Mariana valentina castillo vergara", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1019766050", "correo": "marianavalentina.vergara007@gmail.com", "telefono": "3213182007", "fecha": "28/08/2026", "hora": "11:14:05"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "MARIA STELLA SEGURA CORDÓN", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Liceo Batallon Tenerife", "documento": "55151339", "correo": "stella1616@hotmail.com", "telefono": "3105829558", "fecha": "28/08/2026", "hora": "11:14:06"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Diego Fernando Díaz Espinosa", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "7709198", "correo": "diegofd@ieatanasiog.edu.co", "telefono": "3177248908", "fecha": "28/08/2026", "hora": "11:14:17"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Ramírez María Virginia", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "26553518", "correo": "ramirezmariavirginiao@gmail.com", "telefono": "3203420846", "fecha": "28/08/2026", "hora": "11:14:18"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Paola Andrea Tamayo Vargas", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Personal administrativo", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "", "sede": "", "documento": "1080292809", "correo": "paolatamayov1@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:14:37"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "ALAYAM ABIV VALDERRAMA RODRÍGUEZ", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Rector(a)", "rolForo": "👑 Líder – Rector(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "12128120", "correo": "ieatanasiog@alcaldianeiva.gov.co", "telefono": "3181772523", "fecha": "28/08/2026", "hora": "11:14:56"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Yuderly Medina Tello", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Mañana", "sede": "Loma De La Cruz", "documento": "26428244", "correo": "yuderlymedina@gmail.com", "telefono": "3123636900", "fecha": "28/08/2026", "hora": "11:15:18"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "María Goretti Díaz Rodríguez", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075258575", "correo": "magore1220@gmail.com", "telefono": "3132497130", "fecha": "28/08/2026", "hora": "11:15:29"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Eduardo polania Trujillo", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "79312974", "correo": "eduardopolania08@gmail.com", "telefono": "3177982920", "fecha": "28/08/2026", "hora": "11:15:36"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Amina Surez", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "26592693", "correo": "suarez55.amina@gmail.com", "telefono": "3174300493", "fecha": "28/08/2026", "hora": "11:16:00"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Herminda Garcia Ipuz", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Loma De La Cruz", "documento": "36173616", "correo": "ermindagp@gmail.com", "telefono": "3118759133", "fecha": "28/08/2026", "hora": "11:16:19"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Jorge osorio", "sexo": "Prefiero no decirlo", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "1075218950", "correo": "josorio1987@gmail.commail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:16:33"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Diana Milet Silvestre Cruz", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Personal administrativo", "rolForo": "🙋 Participante", "jornada": "", "sede": "", "documento": "1075213968", "correo": "lautarosilvestrec.13@gmail.com", "telefono": "3107243767", "fecha": "28/08/2026", "hora": "11:16:36"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Flor Delly Trujillo Zuleta", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Guillermo Montenegro", "documento": "26514683", "correo": "flordelly20@gmail.com", "telefono": "3003672249", "fecha": "28/08/2026", "hora": "11:16:54"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Clara Edima Rojas", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Tarde", "sede": "Guillermo Montenegro", "documento": "36169868", "correo": "claraer@ieatanasiog.edu.co", "telefono": "3224707288", "fecha": "28/08/2026", "hora": "11:16:56"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Gladys Perez Ramirez", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "36162594", "correo": "glapera13@hotmail.comotmail.com", "telefono": "3202008318", "fecha": "28/08/2026", "hora": "11:16:59"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Sandra Patricia reyes polanco", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Personal administrativo", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "", "sede": "", "documento": "55173453", "correo": "sandrareyesp72@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:17:08"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Gilberto Murillo Díaz", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "12124850", "correo": "gilmur_diaz@hotmail.es", "telefono": "", "fecha": "28/08/2026", "hora": "11:17:35"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Sara Sofía Popayan Medina", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1029885186", "correo": "sofismedina05@gmail.com", "telefono": "3229870470", "fecha": "28/08/2026", "hora": "11:17:37"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Fernando perez zuleta", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "12129061", "correo": "jfpf2020@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:17:46"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Marilyn campo Montealegre", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "55156124", "correo": "maly6810@hotmail.com", "telefono": "3177904362", "fecha": "28/08/2026", "hora": "11:18:40"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Yeni patricia perdomo castañeda", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Mañana", "sede": "Loma De La Cruz", "documento": "55162421", "correo": "yeniop@ieatanasiog.comedu.co", "telefono": "3172804237", "fecha": "28/08/2026", "hora": "11:19:02"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Diana Margareth Ruiz Sierra", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "26424855", "correo": "dianamr@ieatanasiog.edu.co", "telefono": "3188322299", "fecha": "28/08/2026", "hora": "11:19:18"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Breily Lizeth Gutierrez Ramirez", "sexo": "Femenino", "edad": "13-18", "tipoAsistencia": "Presencial", "cargo": "Estudiante", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1029886348", "correo": "breilyguti@icloud.com", "telefono": "3209843965", "fecha": "28/08/2026", "hora": "11:19:48"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "ISABEL CHARRY DE PERDOMO", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Mañana", "sede": "Guillermo Montenegro", "documento": "36161998", "correo": "isabel.charry10@gmail.commil.com", "telefono": "3188628386", "fecha": "28/08/2026", "hora": "11:22:03"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Ligia ortiz alvarez", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Liceo Batallon Tenerife", "documento": "36164630", "correo": "ligisita@hotmail..com", "telefono": "3187729227", "fecha": "28/08/2026", "hora": "11:22:46"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Faiber Navarro Murcia", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1117504248", "correo": "faiber1188.jdjd@gmail.com", "telefono": "3103764020", "fecha": "28/08/2026", "hora": "11:24:18"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Oscar Rodrigo quimbaya", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Personal administrativo", "rolForo": "🙋 Participante", "jornada": "", "sede": "", "documento": "7705939", "correo": "oscarquimbaya77@gmail.com", "telefono": "3183631214", "fecha": "28/08/2026", "hora": "11:24:32"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Fanny candela", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "36377550", "correo": "fannyc@ieatanasiog.edu.co", "telefono": "", "fecha": "28/08/2026", "hora": "11:47:48"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Andrea Lorena Rodriguez Grijalba", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "26422205", "correo": "andreadeltoboso@gmail.com", "telefono": "3178765444", "fecha": "28/08/2026", "hora": "12:41:42"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Cesar Andrés Vanegas Guevara", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1127070980", "correo": "andresvgs1@gmail.com", "telefono": "3104277927", "fecha": "28/08/2026", "hora": "12:42:02"}, {"idForo": "5abbdb2b-7b66-474b-9004-62c6abed9587", "ie": "ATANASIO GIRARDOT", "nombre": "Leidy Johanna Mora Cardozo", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Tutor(a) PTA/PFI 3.0", "rolForo": "🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "26429912", "correo": "pta.johanna24@gmail.com", "telefono": "3112785158", "fecha": "28/08/2026", "hora": "13:21:00"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Jorge Armando Espinosa Pérez", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Tutor(a) PTA/PFI 3.0", "rolForo": "🎓 Dinamizador(a) Pedagógico(a) – Tutor(a) PTA / PFI 3.0", "jornada": "Única", "sede": "Central/Administrativa", "documento": "7691849", "correo": "espinoxxa1@gmail.com", "telefono": "3132613637", "fecha": "28/08/2026", "hora": "11:30:22"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Freddy Augusto Otalora Portillo", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "12129601", "correo": "dahefrey66@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:34:40"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Nicolás Fabián Tovar Arteaga", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "1075314794", "correo": "scoul1996@gmail.com", "telefono": "3014302630", "fecha": "28/08/2026", "hora": "11:37:00"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Alexandra Mosquera Tovar", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075212567", "correo": "almita_2805@hotmail.comotmail.com", "telefono": "3163923303", "fecha": "28/08/2026", "hora": "11:37:02"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Gabriela Naranjo", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1007328003", "correo": "naranjogabriela06@gmail.com", "telefono": "3132814868", "fecha": "28/08/2026", "hora": "11:37:08"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Paula Andrea Bustamante Celis", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1121956401", "correo": "paulabustamante2015@gmail.com", "telefono": "3125040204", "fecha": "28/08/2026", "hora": "11:37:42"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Saira Valentina Roa", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1233506144", "correo": "sroa@ierodrigolarabonillaneiva.edu.co", "telefono": "3015830520", "fecha": "28/08/2026", "hora": "11:37:48"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Nicolás Jiménez Trujillo", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "1075321354", "correo": "nicolasjimeneztrujillo@gmail.commail.com", "telefono": "3144722406", "fecha": "28/08/2026", "hora": "11:38:24"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "felipe perdomo", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "7733033", "correo": "larsulrich98@hotmail.com", "telefono": "3124226222", "fecha": "28/08/2026", "hora": "11:38:26"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Rosana Bastidas Trujillo", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "26471617", "correo": "rbastidas@ierodrigolarabonillaneiva.comedu.co", "telefono": "", "fecha": "28/08/2026", "hora": "11:38:30"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Astrid Julieth Ramos Gonzalez", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "1075296837", "correo": "asjurago@gmail.com", "telefono": "3185202002", "fecha": "28/08/2026", "hora": "11:39:02"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Diego Leandro Tovar Salazar", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075213662", "correo": "dtovar@ierodrigolarabonillaneiva.edu.co", "telefono": "3215966579", "fecha": "28/08/2026", "hora": "11:39:12"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Yesica Alejandra Olaya Gaspar", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "1075262998", "correo": "yekita2402@hotmail.com", "telefono": "3214825995", "fecha": "28/08/2026", "hora": "11:39:23"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Arnold González Quesada", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075245930", "correo": "argoz7230@gmail.com", "telefono": "3203094499", "fecha": "28/08/2026", "hora": "11:39:28"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "yeny yubely Peña cruz", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075209591", "correo": "yeny182281@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:39:34"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Juan Gabriel Buyucué Perdomo", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "VI centenario", "documento": "16943125", "correo": "juanga09buyucue@gmail.com", "telefono": "3138343483", "fecha": "28/08/2026", "hora": "11:39:40"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Martha Sofía vargas santana", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36303476", "correo": "sofiasantanat19@gmail.com", "telefono": "3105340175", "fecha": "28/08/2026", "hora": "11:39:49"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Oswaldo Martínez", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Coordinador(a)", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "7711520", "correo": "oswaldomartinezmur@gmail.com", "telefono": "3234962933", "fecha": "28/08/2026", "hora": "11:40:10"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Beatriz Elena Villanueva Montenegro", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "41955930", "correo": "bvillanueva@ierodrigolarabonillaneiva.edu.co", "telefono": "", "fecha": "28/08/2026", "hora": "11:40:15"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Richar Fredy Oyola Guzman", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "93343623", "correo": "ritana1965@hotmail.com", "telefono": "3158058050", "fecha": "28/08/2026", "hora": "11:40:17"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Carlos Diaz", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075271029", "correo": "ccarlosdiazzp@gmail.com", "telefono": "3112971847", "fecha": "28/08/2026", "hora": "11:40:27"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Magda Vasquez Ramirez", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075226095", "correo": "vasquezmagddis3@gmail.com", "telefono": "3015460620", "fecha": "28/08/2026", "hora": "11:40:28"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Nataly Sarria Córdoba", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1083908572", "correo": "natalycordoba@gmail.com", "telefono": "3112707312", "fecha": "28/08/2026", "hora": "11:40:35"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Katherine Cordero G", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "1075257977", "correo": "katheriincordero@gmail.com", "telefono": "3168968211", "fecha": "28/08/2026", "hora": "11:40:51"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Alvaro Trujillo Cuenca", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "12116996", "correo": "atrujillo@ierodrigolarabonillaneiva.edu.co", "telefono": "3108843875", "fecha": "28/08/2026", "hora": "11:41:26"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Norma Constanza Barragán Cardona", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "55178166", "correo": "nbarragan94@hotmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:41:36"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Alis Mercedes Vega Cardoso", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "1075223461", "correo": "Avega@ierodrigolarabonillaneiva.edu.co", "telefono": "3202699610", "fecha": "28/08/2026", "hora": "11:41:59"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "EMPERATRIZ PERDOMO CRUZ", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Única", "sede": "VI centenario", "documento": "55153540", "correo": "empera1199@gmail.com", "telefono": "", "fecha": "28/08/2026", "hora": "11:42:35"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "NORMA GRACIELA CASTAÑO CASTRO", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "VI centenario", "documento": "1079605189", "correo": "normacastano1026@gmail.com", "telefono": "3134476639", "fecha": "28/08/2026", "hora": "11:42:37"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "RIOS  GUTIERREZ   OSCAR", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "VI centenario", "documento": "94416819", "correo": "serviciospanacea1a@gmail.com", "telefono": "3122236604", "fecha": "28/08/2026", "hora": "11:42:47"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Cristiam Elias Ramirez Rodriguez", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1117539496", "correo": "cristiamelmago@gmail.com", "telefono": "3045311299", "fecha": "28/08/2026", "hora": "11:42:53"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Jose Adel Barrera Cardozo", "sexo": "Masculino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "83166377", "correo": "joalbaca07@gmail.commail.com", "telefono": "3175860630", "fecha": "28/08/2026", "hora": "11:42:55"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Eliana Carolina Narváez Silva", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36067057", "correo": "enarvaez@ierodrigolarabonillaneiva.edu.co", "telefono": "3105568984", "fecha": "28/08/2026", "hora": "11:43:19"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Edith Johanna Vargas", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "30508514", "correo": "johannavargas0882@hotmail.com", "telefono": "3045470871", "fecha": "28/08/2026", "hora": "11:43:57"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "JHONATAN CAVIEDES", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075227567", "correo": "jhonatan_caviedes@hotmail.com", "telefono": "3112270688", "fecha": "28/08/2026", "hora": "11:44:04"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Luz Nery Barbosa Álvarez", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "26431949", "correo": "neryenglishteacher@gmail.com", "telefono": "3118396590", "fecha": "28/08/2026", "hora": "11:44:09"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Nadya Aranza Vega", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "VI centenario", "documento": "1075253970", "correo": "aranzita18@gmail.com", "telefono": "3115117780", "fecha": "28/08/2026", "hora": "11:44:12"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Laura Cristina Suarez Sanmiguel", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "💻 Dinamizador(a) de la Sistematización", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1032441775", "correo": "laura.suarez1775@gmail.com", "telefono": "3142961187", "fecha": "28/08/2026", "hora": "11:44:39"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Patricia Ramirez Pascuas", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "55151700", "correo": "patyrp12@gmail.com", "telefono": "3123788358", "fecha": "28/08/2026", "hora": "11:44:43"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Angela Constanza Sánchez Polania", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075235925", "correo": "ancosanpo456@hotmail.com", "telefono": "3188641947", "fecha": "28/08/2026", "hora": "11:44:46"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "ANGELA ROSA CASTILLO AGUDELO", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36180928", "correo": "angelarosacastillo@gmail.com", "telefono": "3053455444", "fecha": "28/08/2026", "hora": "11:45:05"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Henry Nelson Ortiz Buitago", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "7716242", "correo": "kepper27@hotmail.com", "telefono": "3156223551", "fecha": "28/08/2026", "hora": "11:45:25"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Luz Andrea Vargas Adames", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "36067028", "correo": "luzandreavargas10@yahoo.es", "telefono": "3202961519", "fecha": "28/08/2026", "hora": "11:46:18"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Marcela Romero", "sexo": "Femenino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Mañana", "sede": "Central/Administrativa", "documento": "1061721223", "correo": "marcelita1606.mr@gmail.com", "telefono": "3104912294", "fecha": "28/08/2026", "hora": "11:47:33"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Juan Carlos Poveda Hernández", "sexo": "Masculino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "7684965", "correo": "volei70@gmail.com", "telefono": "3175310569", "fecha": "28/08/2026", "hora": "11:50:38"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Ángeles Medina Martínez", "sexo": "Femenino", "edad": "55-65", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "36377275", "correo": "anyimedi77@gmail.com", "telefono": "3214868911", "fecha": "28/08/2026", "hora": "11:51:27"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "JUAN CAMILO RAMÍREZ SÁNCHEZ", "sexo": "Masculino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "👥 Dinamizador(a) de Mesas de Trabajo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075256322", "correo": "jcramirez@ierodrigolarabonillaneiva.edu.co", "telefono": "3118997749", "fecha": "28/08/2026", "hora": "12:04:58"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "jhonny leandro guzman perdomo", "sexo": "Masculino", "edad": "35-45", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "8027131", "correo": "jhonnyguzman75@gmail.com", "telefono": "3107906429", "fecha": "28/08/2026", "hora": "12:09:37"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "ANGELA SILVANA AZUERO CUELTAN", "sexo": "Femenino", "edad": "25-35", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "📝 Relator(a)", "jornada": "Única", "sede": "Central/Administrativa", "documento": "1075270129", "correo": "aazuero@ierodrigolarabonillaneiva.edu.co", "telefono": "", "fecha": "28/08/2026", "hora": "12:14:26"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Rosalba Montilla Charry", "sexo": "Femenino", "edad": "65+", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "⏱️ Dinamizador(a) del Tiempo", "jornada": "Única", "sede": "Central/Administrativa", "documento": "26537145", "correo": "lascantoras@hotmail.com", "telefono": "3212016077", "fecha": "28/08/2026", "hora": "12:16:55"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "María Rocío Guarnizo Medina", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "Central/Administrativa", "documento": "55165136", "correo": "guarnizo_chio@hotmail.com", "telefono": "3118112286", "fecha": "28/08/2026", "hora": "12:48:20"}, {"idForo": "e00a0307-2da1-45f1-b2c0-dac5d82510a0", "ie": "RODRIGO LARA BONILLA", "nombre": "Lidia Edith Calderon Guzman", "sexo": "Femenino", "edad": "45-55", "tipoAsistencia": "Presencial", "cargo": "Docente", "rolForo": "🙋 Participante", "jornada": "Única", "sede": "VI centenario", "documento": "55178418", "correo": "e.isilla@hotmail.com", "telefono": "3133769548", "fecha": "29/08/2026", "hora": "07:11:26"}]');
function restaurarAsistenciaQRDesdeRespaldo(){
  const hoja=asegurarHojaAsistenciaQR_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();

  const existentes=new Set();
  if(ultimaFila>=2){
    const filas=hoja.getRange(2,1,ultimaFila-1,hoja.getLastColumn()).getDisplayValues();
    filas.forEach(function(f){
      existentes.add(String(f[mapa.ID_FORO-1]||"").trim()+"|"+String(f[mapa.NUMERO_DOCUMENTO-1]||"").trim());
    });
  }

  const numCols=hoja.getLastColumn();
  const nuevasFilas=[];
  const resumenPorIE={};
  const omitidasPorIE={};

  DATOS_RECUPERADOS_ASISTENCIA_QR_.forEach(function(r){
    const clave=r.idForo+"|"+r.documento;
    if(existentes.has(clave)){
      omitidasPorIE[r.ie]=(omitidasPorIE[r.ie]||0)+1;
      return;
    }
    const fila=new Array(numCols).fill("");
    const valores={
      ID_FORO:r.idForo, IE:r.ie, NOMBRE_COMPLETO:r.nombre, SEXO:r.sexo, EDAD:r.edad,
      TIPO_ASISTENCIA:r.tipoAsistencia, CARGO:r.cargo, ROL_FORO:r.rolForo,
      JORNADA:r.jornada, SEDE:r.sede,
      FORTALEZAS:"", FORTALEZA_OTRO:"", DIFICULTADES:"", DIFICULTAD_OTRO:"",
      NUMERO_DOCUMENTO:r.documento, CORREO:r.correo, TELEFONO:r.telefono,
      CONSENTIMIENTO:"Sí", DISPOSITIVO_ID:"",
      FECHA:r.fecha, HORA:r.hora
    };
    Object.keys(valores).forEach(function(k){ if(mapa[k]) fila[mapa[k]-1]=valores[k]; });
    nuevasFilas.push(fila);
    existentes.add(clave);
    resumenPorIE[r.ie]=(resumenPorIE[r.ie]||0)+1;
  });

  if(nuevasFilas.length){
    hoja.getRange(hoja.getLastRow()+1,1,nuevasFilas.length,numCols).setValues(nuevasFilas);
  }

  Logger.log("========================================");
  Logger.log("RECUPERACIÓN DE ASISTENCIA QR — RESULTADO");
  Logger.log("Filas restauradas ("+nuevasFilas.length+"): "+JSON.stringify(resumenPorIE));
  if(Object.keys(omitidasPorIE).length) Logger.log("Ya existían, omitidas: "+JSON.stringify(omitidasPorIE));
  Logger.log("Nota: FORTALEZAS y DIFICULTADES quedaron en blanco en las filas restauradas -- esa informacion no estaba en ningun respaldo disponible.");
  Logger.log("========================================");

  return {restauradas:resumenPorIE, omitidas:omitidasPorIE};
}

/*****************************************************
 * RESTAURACIÓN DE AvancesForo — INCIDENTE 2026-09-05 (continuación)
 *
 * Además de AsistenciaQR, se confirmó que la hoja AvancesForo del
 * spreadsheet principal perdió las filas de TODAS las IE reales que
 * ya habían enviado su Foro definitivamente antes de cierto momento
 * (quedaron solo las que ya tenían fila por haber enviado después).
 * Sin esa fila, obtenerDatosGuardadosPorIdForo_ no encuentra nada
 * para esas IE — lo cual rompe regenerarInformeFEMPorIE,
 * compilarRespuestasPorGrupoFEM (ya tiene su propio respaldo desde
 * v130) y cualquier otra función que dependa de sus respuestas.
 *
 * Esta función restaura, para cada IE real con ID_PDF_INFORME
 * generado que NO tenga ya una fila en AvancesForo, una fila
 * reconstruida a partir de la copia que sí sigue intacta en
 * "Análisis FEM 2026" (hoja "Respuestas Totales" — sus columnas son
 * las MISMAS de AvancesForo, excepto DATOS, así que se copian
 * directo) más lo que se puede recuperar de otras hojas que nunca se
 * tocaron:
 *   - Sesiones 1/2/3 (S1_P1...S3_P4): copiadas tal cual (ya vienen en
 *     el mismo formato {"valor":...,"tipo":...} que usa DATOS).
 *   - Participación (PART_RECTOR...PART_OTROS): reconstruida como
 *     campos.participantesX.
 *   - Grupo de trabajo: recalculado con obtenerGrupoRealDeIEFEM_.
 *   - Correo institucional y del responsable: de AccesosIE
 *     (EMAIL_IE/EMAIL_RESPONSABLE), que nunca se vio afectado.
 *   - Rector(a): del directorio oficial ("Oficiales"), como mejor
 *     aproximación disponible — OJO: si a alguna IE se le corrigió el
 *     nombre del rector manualmente (como se hizo antes con
 *     Instituto Técnico IPC Andrés Rosa) y esa corrección nunca se
 *     reflejó en "Oficiales", esta restauración trae el nombre
 *     ANTERIOR, no el corregido; conviene revisarlo después.
 *
 * NO se pueden recuperar (quedan simplemente ausentes en campos, no
 * inventados): cargo y nombre de quien envió el formulario, la
 * fotografía de evidencia, la URL del PDF de asistencia (si aplicaba)
 * y el método de asistencia elegido. Regenerar el informe de una IE
 * restaurada por esta función va a mostrar esos datos en blanco.
 *
 * NUNCA sobrescribe una fila que ya exista en AvancesForo — solo
 * agrega las que de verdad faltan. Idempotente: si se vuelve a
 * ejecutar, las ya restauradas se omiten (ya tienen fila).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El resultado (qué se restauró, qué quedó sin
 * poder recuperarse) queda en "Ver registros de ejecución".
 *****************************************************/
function restaurarAvancesForoDesdeAnalisisFEM(){
  const hojaDestino=abrirSpreadsheet_().getSheetByName(HOJA_AVANCES);
  if(!hojaDestino) throw new Error("No se encontró la hoja "+HOJA_AVANCES+".");
  const mapaDestino=mapaHoja_(hojaDestino);

  const ss=obtenerSpreadsheetAnalisisFEM_();
  const hojaOrigen=ss.getSheetByName(HOJA_ANALISIS_TOTALES);
  if(!hojaOrigen||hojaOrigen.getLastRow()<2) throw new Error("No hay datos en \""+HOJA_ANALISIS_TOTALES+"\" para restaurar.");
  const mapaOrigen=mapaHoja_(hojaOrigen);
  const filasOrigen=hojaOrigen.getRange(2,1,hojaOrigen.getLastRow()-1,hojaOrigen.getLastColumn()).getDisplayValues();

  // IE que YA tienen fila en AvancesForo -- nunca se tocan.
  const yaPresentes=new Set();
  const ultimaFilaDestino=hojaDestino.getLastRow();
  if(ultimaFilaDestino>=2 && mapaDestino.ID_FORO){
    hojaDestino.getRange(2,mapaDestino.ID_FORO,ultimaFilaDestino-1,1).getDisplayValues()
      .forEach(function(f){ if(f[0]) yaPresentes.add(String(f[0]).trim()); });
  }

  // AccesosIE (correo institucional y del responsable) y el
  // directorio oficial (rector) -- ninguna de las dos se vio afectada
  // por el incidente.
  const hojaAccesos=asegurarColumnasAccesosIE_();
  const mapaAccesos=mapaHoja_(hojaAccesos);
  const filasAccesos=hojaAccesos.getLastRow()>=2
    ? hojaAccesos.getRange(2,1,hojaAccesos.getLastRow()-1,hojaAccesos.getLastColumn()).getDisplayValues()
    : [];
  function accesoDe_(ie){
    for(let i=0;i<filasAccesos.length;i++){
      if(String(filasAccesos[i][mapaAccesos.IE-1]||"").trim().toUpperCase()===String(ie).trim().toUpperCase()) return filasAccesos[i];
    }
    return null;
  }
  const institucionesOficiales=obtenerInstituciones();

  const NOMBRES_CAMPO_SESION_={
    S1_P1:"respuestaSesion1", S1_P2:"respuestaSesion1Pregunta2",
    S2_P1:"respuestaSesion2Pregunta1",
    S2_P2_ACCION_1:"respuestaSesion2Pregunta2Accion1", S2_P2_ACCION_2:"respuestaSesion2Pregunta2Accion2",
    S2_P2_ACCION_3:"respuestaSesion2Pregunta2Accion3", S2_P2_ACCION_4:"respuestaSesion2Pregunta2Accion4",
    S2_P2_ACCION_5:"respuestaSesion2Pregunta2Accion5",
    S2_P3:"respuestaSesion2Pregunta3", S2_P4:"respuestaSesion2Pregunta4", S2_P5:"respuestaSesion2Pregunta5",
    S3_P1:"respuestaSesion3Pregunta1",
    S3_P2_ACCION_1:"respuestaSesion3Pregunta2Accion1", S3_P2_ACCION_2:"respuestaSesion3Pregunta2Accion2",
    S3_P2_ACCION_3:"respuestaSesion3Pregunta2Accion3", S3_P2_ACCION_4:"respuestaSesion3Pregunta2Accion4",
    S3_P2_ACCION_5:"respuestaSesion3Pregunta2Accion5",
    S3_P3:"respuestaSesion3Pregunta3", S3_P4:"respuestaSesion3Pregunta4"
  };

  const resultado={restauradas:[], omitidasYaPresentes:[]};
  const filasNuevas=[];

  filasOrigen.forEach(function(fila){
    const idForo=mapaOrigen.ID_FORO ? String(fila[mapaOrigen.ID_FORO-1]||"").trim() : "";
    const institucion=mapaOrigen.INSTITUCION ? String(fila[mapaOrigen.INSTITUCION-1]||"").trim() : "";
    if(!idForo || !institucion) return;
    if(yaPresentes.has(idForo)){ resultado.omitidasYaPresentes.push(institucion); return; }

    function val_(col){ return mapaOrigen[col] ? fila[mapaOrigen[col]-1] : ""; }

    const campos={};
    Object.keys(NOMBRES_CAMPO_SESION_).forEach(function(col){
      const crudo=val_(col);
      if(!crudo) return;
      try{ campos[NOMBRES_CAMPO_SESION_[col]]=JSON.parse(crudo); }
      catch(errorParseo){ campos[NOMBRES_CAMPO_SESION_[col]]={valor:String(crudo),tipo:"textarea"}; }
    });
    ROLES_PARTICIPACION_ANALISIS_.forEach(function(id,i){
      const n=Number(val_(COLUMNAS_PARTICIPACION_ANALISIS_[i])||0);
      if(n) campos["participantes"+id]={valor:n,tipo:"number"};
    });
    const grupo=obtenerGrupoRealDeIEFEM_(institucion, idForo);
    if(grupo) campos.grupo={valor:grupo,tipo:"hidden"};

    const acceso=accesoDe_(institucion);
    if(acceso){
      const correoIE=mapaAccesos.EMAIL_IE ? String(acceso[mapaAccesos.EMAIL_IE-1]||"").trim() : "";
      const correoResp=mapaAccesos.EMAIL_RESPONSABLE ? String(acceso[mapaAccesos.EMAIL_RESPONSABLE-1]||"").trim() : "";
      if(correoIE) campos.correoIE={valor:correoIE,tipo:"email"};
      if(correoResp) campos.correo={valor:correoResp,tipo:"email"};
    }
    const oficial=buscarInstitucionOficial_(institucionesOficiales, institucion);
    const rectorOficial=oficial?.datos?.rector ? String(oficial.datos.rector).trim() : "";
    if(rectorOficial) campos.rector={valor:rectorOficial,tipo:"text"};

    const dane=val_("DANE");
    const datosReconstruidos={idForo:idForo, institucion:institucion, dane:dane, campos:campos};

    const filaSalida={};
    Object.keys(mapaOrigen).forEach(function(k){
      if(k==="DATOS") return;
      if(mapaDestino[k]) filaSalida[k]=fila[mapaOrigen[k]-1];
    });
    filaSalida.DATOS=JSON.stringify(datosReconstruidos);

    const filaCompleta=new Array(hojaDestino.getLastColumn()).fill("");
    Object.keys(filaSalida).forEach(function(k){ if(mapaDestino[k]) filaCompleta[mapaDestino[k]-1]=filaSalida[k]; });
    filasNuevas.push(filaCompleta);
    resultado.restauradas.push(institucion);
  });

  if(filasNuevas.length){
    hojaDestino.getRange(hojaDestino.getLastRow()+1,1,filasNuevas.length,hojaDestino.getLastColumn()).setValues(filasNuevas);
  }

  Logger.log("========================================");
  Logger.log("RESTAURACIÓN DE AVANCESFORO DESDE ANÁLISIS FEM — RESULTADO");
  Logger.log("Filas restauradas ("+resultado.restauradas.length+"): "+JSON.stringify(resultado.restauradas));
  if(resultado.omitidasYaPresentes.length) Logger.log("Ya tenían fila, omitidas: "+JSON.stringify(resultado.omitidasYaPresentes));
  Logger.log("Campos que NO se pudieron recuperar en las filas restauradas (quedan ausentes, no inventados): cargo, nombre de quien envió el formulario, evidencia fotográfica, URL del PDF de asistencia, método de asistencia.");
  Logger.log("El rector(a) restaurado viene del directorio oficial \"Oficiales\" -- si a alguna de estas IE se le había corregido el nombre del rector manualmente, esa corrección NO quedó reflejada aquí y conviene revisarla.");
  Logger.log("========================================");

  return resultado;
}

/*
 * CORRECCIÓN 2026-09-05 (AsistenciaQR): al revisar la hoja real con
 * más cuidado se confirmó que los datos SÍ estaban ahí — el problema
 * no era una pérdida de datos, sino que la hoja tenía varias decenas
 * de filas completamente en blanco justo después del encabezado, y
 * las respuestas reales empezaban mucho más abajo. Con el encabezado
 * seguido de tantas filas vacías, la hoja se ve exactamente igual a
 * "no hay nadie registrado" tanto a simple vista como con lecturas
 * automatizadas — de ahí la confusión (mía y del reporte inicial).
 * NO hace falta correr restaurarAsistenciaQRDesdeRespaldo() para
 * este caso: los registros nunca se perdieron.
 *
 * Esta función deja la hoja como se espera que se vea:
 *   1) Corrige ID_FORO en blanco cuando la IE de esa fila sí se puede
 *      identificar sin ambigüedad (una sola celda editada por error a
 *      mano puede dejar una firma "huérfana" que ya no se reconoce
 *      como la misma persona si vuelve a firmar más adelante).
 *   2) Elimina las filas completamente en blanco, para que las
 *      respuestas reales queden justo debajo del encabezado.
 * Nunca toca una fila que tenga cualquier dato, ni siquiera para
 * completar columnas vacías sueltas — solo actúa sobre ID_FORO en
 * blanco (punto 1) y filas 100% vacías (punto 2).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar".
 */
function limpiarAsistenciaQR(){
  const hoja=asegurarHojaAsistenciaQR_();
  const mapa=mapaHoja_(hoja);
  const ultimaFila=hoja.getLastRow();
  const resultado={filasEnBlancoEliminadas:0, idForoCorregidos:[]};
  if(ultimaFila<2){ Logger.log("AsistenciaQR no tiene filas."); return resultado; }

  const numCols=hoja.getLastColumn();
  const valores=hoja.getRange(2,1,ultimaFila-1,numCols).getDisplayValues();

  if(mapa.ID_FORO && mapa.IE){
    for(let i=0;i<valores.length;i++){
      if(String(valores[i][mapa.ID_FORO-1]||"").trim()) continue;
      const ie=String(valores[i][mapa.IE-1]||"").trim();
      if(!ie) continue;
      try{
        const idForoCorrecto=buscarIdForoPorNombreIE_(ie);
        hoja.getRange(i+2, mapa.ID_FORO).setValue(idForoCorrecto);
        resultado.idForoCorregidos.push(ie+" (fila "+(i+2)+")");
      }catch(error){
        Logger.log("No se pudo corregir el ID_FORO en blanco de \""+ie+"\" (fila "+(i+2)+"): "+error.message);
      }
    }
  }

  for(let i=valores.length-1;i>=0;i--){
    const filaVacia=valores[i].every(function(celda){ return String(celda||"").trim()===""; });
    if(filaVacia){ hoja.deleteRow(i+2); resultado.filasEnBlancoEliminadas++; }
  }

  Logger.log("========================================");
  Logger.log("LIMPIEZA DE ASISTENCIAQR — RESULTADO");
  Logger.log("Filas en blanco eliminadas: "+resultado.filasEnBlancoEliminadas);
  Logger.log("ID_FORO corregidos: "+JSON.stringify(resultado.idForoCorregidos));
  Logger.log("========================================");

  return resultado;
}

/*
 * 4) Un Google Doc de ANÁLISIS por grupo (G1-G6). Tiene dos modos,
 * elegidos automáticamente según si el grupo tiene asistencia por
 * código QR registrada en al menos una de sus IE:
 *
 *   - CON asistencia QR: las 9 secciones pedidas (IE del grupo,
 *     participación, edad, fortalezas/debilidades general y por
 *     grupo de edad, comparativo) — dejando explícito, antes de la
 *     sección 3, cuáles IE del grupo son las que sí registraron
 *     asistencia por QR (las demás del grupo simplemente no aportan
 *     datos de percepción, pero sí de participación/caracterización).
 *   - SIN asistencia QR suficiente (ninguna IE del grupo firmó por
 *     QR): en vez de las secciones 3-9 (que quedarían vacías), un
 *     resumen de las respuestas de las Sesiones 1-4 de cada IE y de
 *     su Valoración del Foro, con gráfico y descripción en texto de
 *     las calificaciones (preguntas de selección múltiple 1-5) de
 *     cada IE que ya valoró.
 *
 * Ver generarDocumentoAnalisisGrupoFEM_ en Código.js. Se deja en la
 * raíz de la carpeta del grupo, junto al documento de "Respuestas
 * Compiladas". Igual que ese, NUNCA crea un documento nuevo si ya
 * existe uno con el mismo nombre: lo reescribe en el mismo archivo.
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar".
 */
function generarDocumentosAnalisisPorGrupoFEM(){
  const grupos=mapaGruposFEM_();
  const resultado={};

  GRUPOS_FEM_ORDEN_.forEach(function(g){
    const miembros=(grupos[g]||[]).slice().sort(function(a,b){ return a.nombreIE.localeCompare(b.nombreIE,"es"); });

    const conteoPorRol=ROLES_PARTICIPACION_ANALISIS_.map(function(){ return 0; });
    let asistentesGrupo=[];
    const ieConQR=[];
    miembros.forEach(function(m){
      try{
        const datos=obtenerDatosGuardadosPorIdForo_(m.idForo);
        const c=(datos&&datos.campos)||{};
        ROLES_PARTICIPACION_ANALISIS_.forEach(function(id,i){ conteoPorRol[i]+=Number(c["participantes"+id]?.valor||0); });
      }catch(error){ Logger.log("Participación de "+m.nombreIE+" para análisis de grupo "+g+": "+error.message); }
      try{
        const asistentesIE=obtenerAsistentesQR_(m.idForo);
        if(asistentesIE.length){ ieConQR.push(m.nombreIE); asistentesGrupo=asistentesGrupo.concat(asistentesIE); }
      }catch(error){ Logger.log("Asistentes QR de "+m.nombreIE+" para análisis de grupo "+g+": "+error.message); }
    });

    const totalParticipantes=conteoPorRol.reduce(function(a,b){return a+b;},0);
    const porRol=ETIQUETAS_PARTICIPACION_ANALISIS_.map(function(etiqueta,i){ return {etiqueta:etiqueta, total:conteoPorRol[i]}; });

    // Sin asistencia QR en ninguna IE del grupo: se arma en su lugar
    // el resumen de respuestas (Sesiones 1-4) y valoración por IE.
    let resumenIEs=null;
    if(!asistentesGrupo.length){
      resumenIEs=miembros.map(function(m){
        const resuelto=resolverSesionesIEFEM_(m);
        let valoracion=null;
        try{ valoracion=obtenerValoracionPorIdForo_(m.idForo); }
        catch(error){ Logger.log("Valoración de "+m.nombreIE+" para análisis de grupo "+g+": "+error.message); }
        return {nombreIE:m.nombreIE, sesiones:resuelto.sesiones, sesionPropia:resuelto.sesionPropia, valoracion:valoracion};
      });
    }

    const analisis={
      miembros:miembros.map(function(m){return m.nombreIE;}),
      porRol:porRol,
      totalParticipantes:totalParticipantes,
      asistentesGrupo:asistentesGrupo,
      ieConQR:ieConQR,
      resumenIEs:resumenIEs
    };

    const carpetas=crearEstructuraCarpetasGrupoFEM_(g);
    const nombreDoc="Análisis de Grupo "+g+" FEM 2026";
    const existentesIt=carpetas.grupoFolder.getFilesByName(nombreDoc);
    const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
    while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

    const archivoDoc=generarDocumentoAnalisisGrupoFEM_(g, analisis, archivoExistente?archivoExistente.getId():null);
    if(!archivoExistente){
      carpetas.grupoFolder.addFile(archivoDoc);
      try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
    }

    resultado[g]={
      documento:archivoDoc.getUrl(),
      ieIncluidas:analisis.miembros,
      totalParticipantes:totalParticipantes,
      totalAsistentesQR:asistentesGrupo.length,
      ieConQR:ieConQR,
      modo:asistentesGrupo.length?"percepcion_qr":"resumen_respuestas_valoracion"
    };
  });

  Logger.log("========================================");
  Logger.log("DOCUMENTOS DE ANÁLISIS POR GRUPO — RESULTADO");
  Logger.log(JSON.stringify(resultado, null, 2));
  Logger.log("========================================");

  return resultado;
}

/*
 * INFORME DE SÍNTESIS GRUPAL — GRUPO 1 (muestra completa a revisar
 * antes de replicar a G2-G6).
 *
 * A diferencia de todo lo anterior en este archivo, esta prosa NO se
 * calcula a partir de datos en vivo: se redactó a mano, leyendo
 * directamente el Doc editable ya generado y enviado de cada una de
 * las 6 IE del Grupo 1 (AIPECITO, CHAPINERO, I.E. CLARETIANO GUSTAVO
 * TORRES PARRA, INEM JULIAM MOTTA SALAS, LICEO DE SANTA LIBRADA,
 * PROMOCION SOCIAL), tal como se pidió expresamente: encontrar los
 * elementos en común entre las 6 IE para cada pregunta y señalar,
 * aparte, las particularidades de una sola IE. Las tallas de las 4
 * preguntas de selección múltiple (Sesión 2 P3/P5, Sesión 3 "Equipos
 * de trabajo"/"Mecanismos de seguimiento") se contaron a mano, opción
 * por opción, sobre el texto real seleccionado por cada una de las 6
 * IE.
 *
 * PENDIENTE A COMPLETAR MANUALMENTE por la SEM antes de enviar/firmar
 * el informe (no hay fuente de datos para esto en el sistema): el
 * nombre de quien consolida el informe, su cargo, y el nombre del
 * funcionario o funcionaria de la SEM Neiva a cargo — quedan como
 * texto entre corchetes, editable directamente en el Doc generado.
 *
 * A la fecha de redacción (7 de septiembre de 2026), ninguna de las 6
 * IE del Grupo 1 había registrado su Valoración formal del Foro
 * (todas sus fichas ejecutivas dicen literalmente "La institución
 * educativa aún no ha registrado la valoración de esta jornada"), por
 * lo que las conclusiones NO inventan una valoración: lo dicen
 * explícitamente y usan, como única referencia real disponible, la
 * percepción de fortalezas/oportunidades registrada por firma QR en
 * el Liceo de Santa Librada (única IE del grupo con asistencia QR).
 */
const DATOS_SINTESIS_GRUPO_1_FEM_ = {
  grupo: "G1",
  tituloInforme: "Informe de Síntesis Grupal — Elementos Comunes y Particularidades Institucionales",
  instituciones: [
    "AIPECITO",
    "CHAPINERO",
    "I.E. CLARETIANO GUSTAVO TORRES PARRA",
    "INEM JULIAM MOTTA SALAS",
    "LICEO DE SANTA LIBRADA",
    "PROMOCION SOCIAL"
  ],
  responsableInforme: "[Nombre de quien consolida el informe de síntesis del Grupo 1 — SEM Neiva]",
  fechaPresentacion: "7 de septiembre de 2026",
  secciones: [
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta orientadora",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en el logro de los retos y propósitos planteados en el FEM2025?",
      tipo: "cualitativo",
      comun: "Las seis instituciones educativas del Grupo 1 (AIPECITO, Chapinero, I.E. Claretiano Gustavo Torres Parra, INEM Juliam Motta Salas, Liceo de Santa Librada y Promoción Social) coinciden en reconocer avances sustantivos hacia una educación más pertinente y contextualizada, materializados principalmente en el fortalecimiento de proyectos pedagógicos transversales, experiencias significativas y estrategias de articulación con el territorio (huertas escolares, PRAE, alianzas con el SENA, proyectos productivos agrícolas y culturales). Un segundo eje común es el avance en el reconocimiento del contexto socioeconómico y territorial de los estudiantes como insumo para flexibilizar las prácticas pedagógicas, sea mediante caracterizaciones diagnósticas (Claretiano), la contextualización curricular hacia dinámicas productivas regionales (AIPECITO, INEM) o el fortalecimiento de los hilos conductores institucionales (Promoción Social). No obstante, el análisis conjunto revela que estos avances se ven sistemáticamente limitados por un tercer elemento compartido: la persistencia de dificultades estructurales de infraestructura, dotación y conectividad —deterioro de instalaciones (AIPECITO), brechas tecnológicas (INEM)— que, según señalan las mesas de trabajo, condicionan la capacidad real de sostener y profundizar los avances pedagógicos alcanzados. De igual manera, se identifica de forma recurrente la necesidad de fortalecer la participación efectiva de las familias y la comunidad educativa como condición para consolidar los logros obtenidos y dar continuidad a los procesos iniciados desde el FEM2025.",
      particularidades: "La IE Chapinero, única institución de carácter rural dentro del grupo, señaló como obstáculo específico la alta rotación del personal docente y la ausencia de acompañamiento presencial de la Secretaría de Educación al sector rural, factores que —a diferencia de lo expresado por las demás instituciones del grupo— condicionan la continuidad institucional de las experiencias significativas, al quedar ligadas al docente que las lidera y no a una política de sistematización institucional. Por su parte, el Liceo de Santa Librada planteó una particularidad de tipo administrativo poco común en el grupo: no contar con una población estudiantil propia, lo que exige gestionar ante el ente territorial garantías de sostenibilidad y transporte escolar para estudiantes provenientes de otras comunas."
    },
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta 2",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en la implementación de los nuevos grados del nivel de preescolar (jardín, prejardín)?",
      tipo: "cualitativo",
      comun: "En las seis instituciones se observa un patrón común de implementación parcial y asimétrica entre los grados de jardín y prejardín: el grado jardín registra avances más consolidados —incorporación de nuevos grupos (Claretiano), vinculación de sedes adicionales (AIPECITO), aulas específicas con acompañamiento docente (INEM, Chapinero)—, mientras que la apertura de prejardín se mantiene, en la generalidad del grupo, como un reto pendiente. Dos factores explicativos se repiten transversalmente: de un lado, una demanda de matrícula insuficiente para habilitar la apertura de nuevos grupos (Liceo de Santa Librada, INEM); de otro, limitaciones de infraestructura y disponibilidad de aulas adecuadas para la primera infancia (Liceo de Santa Librada, Chapinero). En el plano pedagógico coincide, igualmente, la apuesta por currículos y prácticas centradas en el desarrollo integral, el juego, el arte y la exploración del entorno (AIPECITO, Claretiano, Promoción Social), con un tránsito declarado desde modelos organizados por dimensiones del desarrollo hacia enfoques por propósitos y valoraciones cualitativas del aprendizaje infantil.",
      particularidades: "La IE Chapinero constituye una particularidad clara dentro del grupo al señalar que las condiciones viales de sus sedes rurales dificultan la movilización de niños y niñas de temprana edad hacia la escuela, un obstáculo de tipo geográfico-territorial no mencionado por ninguna de las otras cinco instituciones, de carácter urbano o periurbano. El INEM Juliam Motta Salas, por su parte, identificó una causa específica de baja demanda de prejardín no reportada por las demás IE: la competencia de los hogares comunitarios (CDI), que retienen a las familias por los beneficios complementarios que allí reciben, dificultando la transición oportuna de esa población hacia la oferta oficial."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que los currículos actuales que se desarrollan en las instituciones educativas son pertinentes con sus realidades territoriales (sociales, culturales, productivas)? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Existe consenso entre las seis instituciones del grupo en calificar la pertinencia curricular como parcial: los currículos institucionales se reconocen alineados con los lineamientos del Ministerio de Educación Nacional, pero insuficientemente contextualizados frente a las realidades sociales, culturales y productivas de cada territorio. Es recurrente la idea de que la escuela \"forma en el territorio, pero no siempre desde el territorio\" (expresión explícita del Liceo de Santa Librada, compartida en sustancia por INEM, Claretiano y Promoción Social), señalando como vacíos comunes la escasa incorporación de saberes ancestrales y culturales, la limitada articulación con las dinámicas productivas locales y la persistencia de un enfoque curricular más estandarizado que situado. Como condición asociada a esta pertinencia parcial, varias instituciones (Claretiano, AIPECITO) vinculan explícitamente la calidad de los aprendizajes con las condiciones materiales e institucionales —infraestructura, conectividad, servicios básicos— que inciden directamente en la posibilidad de desarrollar un currículo realmente contextualizado.",
      particularidades: "La IE Chapinero presenta una particularidad relevante para el grupo al circunscribir la pertinencia territorial casi exclusivamente a la dinámica del cultivo del café, señalando que la ausencia de flexibilización curricular durante la época de cosecha incrementa el ausentismo escolar, una tensión entre calendario productivo agrícola y calendario escolar no reportada por ninguna otra institución del grupo. AIPECITO, por su parte, se distingue por describir una experiencia curricular concreta y ya consolidada de articulación entre matemáticas y proyectos productivos agrícolas en alianza con el SENA, mientras las demás instituciones describen la pertinencia productiva de manera más general o aspiracional."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se han implementado para lograr currículos más pertinentes al territorio?",
      tipo: "cualitativo",
      comun: "Al observar de manera conjunta las acciones reportadas por las seis instituciones, se configuran tres líneas de trabajo comunes al grupo. La primera es la actualización y ajuste permanente de mallas curriculares y planes de área con enfoque de pertinencia territorial, presente explícitamente en INEM (actualización curricular por competencias), Chapinero (ajuste de mallas y plan de estudios) y Claretiano (espacios de reflexión y revisión curricular por niveles). La segunda es la incorporación de metodologías activas y experiencias pedagógicas contextualizadas —salidas pedagógicas, proyectos de aula, uso de recursos tecnológicos y espacios institucionales como escenarios de aprendizaje—, reportada de forma prácticamente unánime (AIPECITO, INEM, Claretiano, Promoción Social). La tercera es el fortalecimiento de proyectos transversales que vinculan saberes ambientales, culturales y productivos del territorio: educación ambiental y manejo de residuos (AIPECITO), integración de saberes ancestrales y agricultura escolar (INEM), identidad huilense y patrimonio cultural (Promoción Social), intensificación de inglés y ferias de ciencia (Claretiano). En conjunto, estas acciones evidencian una apuesta institucional compartida por transformar el currículo desde la vía de los proyectos transversales y las metodologías activas, más que desde una reforma estructural del plan de estudios.",
      particularidades: "El INEM Juliam Motta Salas es la única institución del grupo que reporta una articulación formal y diversificada con el SENA para el desarrollo de modalidades técnicas (turismo, software, agropecuaria, emprendimiento) como estrategia curricular estructurada; aunque AIPECITO también menciona al SENA, lo hace en referencia a un proyecto puntual y no a una articulación por modalidades técnicas. La IE Chapinero, por su parte, es la única en asociar explícitamente el fortalecimiento curricular con la infraestructura tecnológica institucional en el marco de la jornada única, una acción no reportada por las demás instituciones del grupo."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 3",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar estas acciones?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Consejo Académico", count: 6},
        {opcion: "Consejo Directivo", count: 6},
        {opcion: "Consejo de Padres de Familia", count: 6},
        {opcion: "Comités de área", count: 5},
        {opcion: "Consejo Estudiantil", count: 4},
        {opcion: "Personero(a) Estudiantil", count: 4},
        {opcion: "Comisión de Evaluación y Promoción", count: 4},
        {opcion: "Equipo de Autoevaluación Institucional", count: 4}
      ],
      comun: "El análisis cuantitativo de las seis instituciones del grupo muestra que el Consejo Académico, el Consejo Directivo y el Consejo de Padres de Familia son los equipos de trabajo institucional que el 100% de las IE del Grupo 1 identifican como responsables de liderar las acciones hacia currículos más pertinentes. Con un nivel de adopción también mayoritario (5 de 6 instituciones, 83,3%) aparecen los comités de área, y con presencia en 4 de las 6 instituciones (66,7%) se ubican el Consejo Estudiantil, el/la Personero(a) Estudiantil, la Comisión de Evaluación y Promoción y el Equipo de Autoevaluación Institucional. Este patrón evidencia que el grupo recurre de manera prioritaria a las instancias de gobierno escolar de carácter reglamentario y a la participación de las familias como base estructural para liderar la pertinencia curricular, mientras que instancias más especializadas de evaluación institucional, aunque relevantes, no alcanzan aún una adopción unánime en el grupo.",
      particularidades: "AIPECITO es la única institución del grupo en no reportar la conformación del Consejo Estudiantil ni de la figura de Personero(a) Estudiantil como equipos vinculados a esta pregunta, apoyándose en cambio en un Comité de Convivencia Escolar (opción \"Otro\" seleccionada únicamente por esta institución). El Liceo de Santa Librada, por su parte, es la única IE del grupo en reportar equipos de trabajo de carácter cultural y deportivo (grupos de rajaleñas, banda marcial, grupos deportivos, comité de riesgos y desastres escolares) como parte de su estructura de equipos, una particularidad que no se repite en ninguna otra institución del grupo."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 4",
      enunciado: "¿Cómo se están articulando estos equipos de trabajo para lograr currículos más pertinentes territorialmente?",
      tipo: "cualitativo",
      comun: "Las seis instituciones coinciden en que la articulación entre equipos de trabajo institucional se apoya principalmente en espacios ya existentes —reuniones de área, Consejo Académico, jornadas pedagógicas y proyectos transversales— más que en mecanismos formales de coordinación creados para este fin específico. De manera igualmente compartida, se reconoce que dicha articulación es todavía incipiente o insuficiente: el Liceo de Santa Librada señala que los espacios de diálogo son \"limitados\" y que persiste \"falta de compromiso\" de algunos actores; el INEM concluye que \"no existe una articulación efectiva\" y que el trabajo se realiza \"de manera aislada por asignaturas\"; Claretiano plantea la necesidad de \"fortalecer la coordinación institucional\"; y Promoción Social describe una articulación que, aunque existente a través de reuniones y proyectos, requiere superar el trabajo fragmentado. AIPECITO matiza este panorama al destacar el papel del liderazgo directivo cercano y la participación incluyente de familias y estudiantes como catalizadores de una mejor articulación.",
      particularidades: "La IE Chapinero constituye la particularidad más marcada del grupo en esta pregunta: a diferencia de las demás instituciones, que describen limitaciones puntuales dentro de una articulación que de todas formas ocurre, Chapinero señala la falta de trabajo en equipo y de transferencia del conocimiento entre docentes como la debilidad institucional central identificada por su propia nueva dirección, un diagnóstico autocrítico de mayor calado que no tiene equivalente explícito en las otras cinco instituciones del grupo."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 5",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para que dichas acciones se cumplan?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Autoevaluación institucional anual", count: 5},
        {opcion: "Consejo Académico", count: 4},
        {opcion: "Análisis de resultados de Pruebas Saber e ICFES", count: 4},
        {opcion: "Comités de área o núcleos de formación", count: 4},
        {opcion: "Actas de reunión de área", count: 4},
        {opcion: "Plataformas de gestión académica (JIGRA)", count: 4},
        {opcion: "Comparación histórica de resultados internos y externos", count: 3},
        {opcion: "Articulación con el PEI", count: 3},
        {opcion: "Tasa de deserción/reprobación por asignatura", count: 3}
      ],
      comun: "En cuanto a los mecanismos de seguimiento a las acciones curriculares, el elemento más ampliamente compartido por el grupo es la autoevaluación institucional anual, seleccionada por 5 de las 6 instituciones (83,3%). Con una adopción del 66,7% (4 de 6 IE) se ubican el Consejo Académico, el análisis de resultados de Pruebas Saber e ICFES, los comités de área o núcleos de formación, las actas de reunión de área y las plataformas de gestión académica (JIGRA); y con una adopción del 50% (3 de 6 IE), la comparación histórica de resultados internos y externos, la articulación con el PEI y la tasa de deserción o reprobación por asignatura. Este patrón revela que el seguimiento al cumplimiento de las acciones curriculares del grupo se apoya mayoritariamente en instrumentos internos de evaluación institucional y en el análisis de resultados académicos externos, más que en mecanismos de seguimiento comunitario o de participación de las familias en dicho seguimiento.",
      particularidades: "AIPECITO constituye la particularidad más marcada de esta pregunta en todo el grupo: es la única institución que no seleccionó ninguno de los mecanismos de seguimiento del catálogo institucional, reportando en su lugar, exclusivamente, dos mecanismos externos y comunitarios poco convencionales para esta pregunta (el propio Foro Educativo Institucional y la Junta de Acción Comunal), lo que sugiere un mecanismo de seguimiento de tipo comunitario-territorial distinto al del resto del grupo. La IE Chapinero, por su parte, reportó un único mecanismo (autoevaluación institucional anual), el nivel de selección más bajo del grupo, lo que puede sugerir que sus mecanismos de seguimiento curricular están aún en una etapa incipiente de formalización."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que la toma de decisiones en las instituciones educativas actualmente es participativa y democrática? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Existe una coincidencia sustancial entre las seis instituciones del grupo: todas reconocen la existencia de estructuras formales de gobierno escolar (Consejo Directivo, Consejo Académico, Consejo Estudiantil, Personero, Contralor) que en principio garantizan cauces democráticos, pero todas —sin excepción— matizan esa existencia formal señalando límites concretos a la participación real. Se repite la idea de una democracia \"más formal que vivida\" (expresión textual del INEM, compartida en sustancia por Claretiano, Promoción Social y Chapinero): las reuniones tienden a ser informativas más que deliberativas (Liceo de Santa Librada), las decisiones continúan centralizadas en el Consejo Académico o en los directivos (INEM, Promoción Social), y persisten bajos niveles de incidencia efectiva de estudiantes y familias más allá de la elección de representantes. La escasa participación de las familias, explicada en términos de timidez, desinterés, falta de tiempo o desconocimiento de funciones (AIPECITO, INEM, Promoción Social), constituye igualmente un obstáculo transversal identificado por el grupo para consolidar una democracia escolar plenamente incidente.",
      particularidades: "AIPECITO se distingue del resto del grupo por reportar un resultado concreto y tangible de participación comunitaria efectiva —proyectos de señalización y demarcación vial en la sede, surgidos de iniciativas directas de la comunidad—, mientras que las demás instituciones describen la participación fundamentalmente en términos de estructuras y limitaciones, sin un logro equivalente atribuible directamente a la incidencia comunitaria en la toma de decisiones."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se están implementando para canalizar y fortalecer la participación de la comunidad educativa?",
      tipo: "cualitativo",
      comun: "El análisis conjunto de las acciones reportadas por las seis instituciones evidencia tres estrategias comunes de fortalecimiento de la participación escolar. La primera es el fortalecimiento de los procesos electorales y de representación del Gobierno Escolar (elección de representantes por salón y jornada, personeros, contralores), presente en AIPECITO, Liceo de Santa Librada, INEM, Claretiano y Promoción Social. La segunda es la diversificación de canales de comunicación institucional hacia la comunidad —circulares informativas (Liceo de Santa Librada), medios digitales e institucionales (INEM, AIPECITO)— orientada a visibilizar procesos y necesidades y a estrechar la relación entre la institución y sus públicos. La tercera es la apertura de espacios formales de diálogo y acompañamiento a las familias, como las escuelas de padres o de familia (Claretiano, Promoción Social) y las reuniones de rendición de cuentas o seguimiento (Liceo de Santa Librada, INEM). En conjunto, estas acciones muestran una apuesta compartida por ampliar los canales de representación y comunicación, más que por transformar de fondo los mecanismos de deliberación institucional.",
      particularidades: "El INEM Juliam Motta Salas es la única institución del grupo en reportar un canal de comunicación audiovisual propio (\"Inemitas TV\") como estrategia de divulgación institucional y acercamiento a la comunidad, particularidad no replicada por ninguna otra IE del grupo. La IE Chapinero, por su parte, reportó de manera única un \"plan padrino\" liderado por docentes de sedes veredales, orientado a fortalecer la permanencia y motivación de los estudiantes mediante acompañamiento personalizado, estrategia sin equivalente explícito en las demás instituciones."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Equipos de trabajo",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar las estrategias y mecanismos de participación escolar?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Consejo Directivo", count: 6},
        {opcion: "Consejo de Padres de Familia", count: 6},
        {opcion: "Comité Escolar de Convivencia", count: 6},
        {opcion: "Consejo Académico", count: 5},
        {opcion: "Gobierno Escolar", count: 5},
        {opcion: "Consejo Estudiantil", count: 5},
        {opcion: "Personero Estudiantil", count: 5},
        {opcion: "Contralor Estudiantil", count: 5},
        {opcion: "Comité de Calidad / Equipo de Gestión Institucional", count: 4},
        {opcion: "Comisión de Evaluación y Promoción", count: 3}
      ],
      comun: "Para el liderazgo de las estrategias y mecanismos de participación escolar, el Consejo Directivo, el Consejo de Padres de Familia y el Comité Escolar de Convivencia son los equipos que el 100% de las instituciones del grupo reportan haber conformado. Con una adopción del 83,3% (5 de 6 IE) se ubican el Consejo Académico, el Gobierno Escolar en su conjunto, el Consejo Estudiantil, el/la Personero(a) Estudiantil y el/la Contralor(a) Estudiantil; y con una adopción del 66,7% (4 de 6 IE), el Comité de Calidad o Equipo de Gestión Institucional. Este patrón evidencia una estructura de gobierno escolar ampliamente consolidada y homogénea en el grupo, en la que las instancias de representación estudiantil alcanzan un nivel de adopción casi unánime, superior incluso al observado en los equipos de trabajo curricular de la Sesión 2.",
      particularidades: "AIPECITO es, nuevamente, la institución que marca la particularidad más notable del grupo: es la única que no reporta el Gobierno Escolar como tal, ni el Consejo Estudiantil, la Personería ni la Contraloría Estudiantil entre sus equipos de participación, apoyándose en cambio en comités alternativos y específicos —un Comité de Veeduría del PAE y comités de Gestión del Riesgo y Medio Ambiente, ambos exclusivos de esta institución dentro del grupo—, lo que sugiere una estructura de participación escolar organizada de manera distinta a la del resto de instituciones del Grupo 1, con énfasis en la veeduría de programas específicos más que en la representación estudiantil clásica."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Mecanismos de seguimiento",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para garantizar las acciones encaminadas a promover gobiernos educativos democráticos?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Autoevaluación institucional anual (Guía 34 – MEN)", count: 5},
        {opcion: "Plan de Mejoramiento Institucional (PMI)", count: 5},
        {opcion: "Elecciones estudiantiles", count: 5},
        {opcion: "Elección de docentes para consejo directivo", count: 5},
        {opcion: "Elección de docentes para consejo académico", count: 4},
        {opcion: "Reuniones periódicas con entes de gobierno escolar", count: 4},
        {opcion: "Informes de gestión del Consejo Directivo", count: 4},
        {opcion: "Actualización y ajuste permanente del PEI", count: 3},
        {opcion: "Seguimiento periódico al Plan de Mejoramiento", count: 3}
      ],
      comun: "En relación con los mecanismos de seguimiento a las acciones que promueven gobiernos educativos democráticos, la autoevaluación institucional anual (Guía 34 del MEN), el Plan de Mejoramiento Institucional, las elecciones estudiantiles y la elección de docentes para el Consejo Directivo son los mecanismos más ampliamente compartidos por el grupo, cada uno reportado por 5 de las 6 instituciones (83,3%). Con una adopción del 66,7% (4 de 6 IE) se encuentran la elección de docentes para el Consejo Académico, las reuniones periódicas con los distintos entes de gobierno escolar y los informes de gestión del Consejo Directivo. Este patrón muestra que el seguimiento a la democracia escolar en el grupo se sustenta, sobre todo, en los procesos electorales reglamentarios y en los instrumentos institucionales de autoevaluación y mejoramiento continuo dispuestos por la normativa del Ministerio de Educación Nacional, más que en mecanismos de participación comunitaria de base creados específicamente para este fin.",
      particularidades: "AIPECITO vuelve a constituir la particularidad más marcada del grupo en esta pregunta: es la única institución que no seleccionó ninguno de los mecanismos reglamentarios del catálogo, reportando en su lugar, de manera exclusiva, seis mecanismos de seguimiento propios de corte comunitario-territorial (sistematicidad y registro de acuerdos, diversificación de canales informativos en el territorio, mecanismos de consulta y diagnóstico, seguimiento a la transparencia y rendición de cuentas, estrategias de asistencia e involucramiento constante, y reafirmación del Gobierno Escolar), consolidando un patrón consistente a lo largo de las cuatro preguntas mixtas: mientras las demás cinco instituciones del grupo anclan su seguimiento en instrumentos normativos e institucionales estandarizados, AIPECITO construye respuestas propias fuera del catálogo sugerido, lo que amerita un acompañamiento específico de la Secretaría de Educación para verificar si ello refleja una práctica real distinta o una dificultad para reconocerse en las categorías estándar propuestas."
    }
  ],
  conclusiones: "El análisis conjunto de las tres sesiones de trabajo del Grupo 1 permite concluir que las seis instituciones educativas comparten una trayectoria similar: avances reales y verificables en la contextualización curricular, el fortalecimiento de la educación inicial y la consolidación de estructuras de gobierno escolar, que conviven con retos comunes en materia de infraestructura, participación efectiva de las familias y articulación real —más allá de lo formal— entre los distintos equipos de trabajo institucional. Un hallazgo transversal relevante es el patrón particular de la IE AIPECITO, que en las cuatro preguntas de selección múltiple analizadas respondió de manera consistente por fuera del catálogo estándar de opciones, apoyándose en mecanismos y equipos de corte comunitario-territorial; esta recurrencia amerita un seguimiento específico por parte de la Secretaría de Educación Municipal para determinar si corresponde a una práctica institucional genuinamente distinta o a una dificultad de esa institución para reconocerse en las categorías propuestas por el instrumento. En cuanto a la valoración general del Foro, a la fecha de consolidación de este informe ninguna de las seis instituciones del grupo ha registrado formalmente su valoración de la jornada (preguntas de valoración y comentarios/sugerencias), por lo cual no es posible incorporar aquí un balance cuantitativo de percepción ni comentarios o sugerencias institucionales sobre el desarrollo del Foro; se recomienda a la Secretaría de Educación gestionar el diligenciamiento pendiente de este instrumento con las seis instituciones del grupo para completar el ciclo de retroalimentación previsto por el Foro Educativo Institucional. Como única referencia disponible de percepción sobre el Foro, procedente de la asistencia registrada por código QR, la comunidad educativa del Liceo de Santa Librada valoró como principales fortalezas el trabajo colaborativo entre docentes, la participación activa de los estudiantes y la existencia de experiencias exitosas replicables, señalando como oportunidades de mejoramiento la baja participación de estudiantes y familias, las brechas de aprendizaje entre grados y sedes, y el bajo logro de aprendizajes fundamentales — percepción que resulta coherente con los hallazgos cualitativos identificados en las tres sesiones del presente informe para el conjunto del grupo.",
  proyectoNombre: "[Nombre de quien consolida el Informe Consolidado]",
  proyectoCargo: "[Cargo]",
  fechaRealizacion: "7 de septiembre de 2026"
};

/*
 * Genera (o reescribe, si ya existe) el Informe de Síntesis Grupal del
 * Grupo 1, dejándolo en la raíz de la carpeta "Grupo G1" (mismo
 * criterio de idempotencia que compilarRespuestasPorGrupoFEM y
 * generarDocumentosAnalisisPorGrupoFEM: nunca duplica el archivo).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El enlace del documento queda en "Ver
 * registros de ejecución".
 */
function generarInformeSintesisGrupo1FEM(){
  const carpetas=crearEstructuraCarpetasGrupoFEM_("G1");
  const nombreDoc="Informe de Síntesis - Grupo G1 FEM 2026";
  const existentesIt=carpetas.grupoFolder.getFilesByName(nombreDoc);
  const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
  while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

  const archivoDoc=generarInformeSintesisGrupoFEM_(DATOS_SINTESIS_GRUPO_1_FEM_, archivoExistente?archivoExistente.getId():null);
  if(!archivoExistente){
    carpetas.grupoFolder.addFile(archivoDoc);
    try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
  }

  Logger.log("========================================");
  Logger.log("INFORME DE SÍNTESIS — GRUPO 1 — RESULTADO");
  Logger.log("Documento: "+archivoDoc.getUrl());
  Logger.log("========================================");

  return archivoDoc.getUrl();
}

/*
 * INFORME DE SÍNTESIS GRUPAL — GRUPO 2.
 *
 * Mismo criterio que el Grupo 1: prosa redactada a mano leyendo
 * directamente el Doc editable ya generado de cada IE del grupo, y
 * tallas de las 4 preguntas mixtas contadas a mano sobre las opciones
 * realmente seleccionadas.
 *
 * PARTICULARIDAD DE ESTE GRUPO: de las 6 IE del catálogo (MARIA
 * CRISTINA ARANGO DE PASTRANA., LUIS IGNACIO ANDRADE, GABRIEL GARCIA
 * MARQUEZ, EDUARDO SANTOS, MARIA AUXILIADORA FORTALECILLAS, JAIRO
 * MOSQUERA MORENO), MARIA CRISTINA ARANGO DE PASTRANA. NO HABÍA
 * ENVIADO SU INFORME del Foro a la fecha de redacción (7 de septiembre
 * de 2026) — solo existe su logo en Drive, ningún "Informe Ejecutivo".
 * Por eso el análisis de las 11 preguntas y las conclusiones se basan
 * en las 5 IE restantes (denominador 5, no 6), y la portada lista las
 * 6 IE del grupo pero las conclusiones señalan expresamente la
 * ausencia de esta institución.
 *
 * De esas 5 IE, solo Fortalecillas había registrado su Valoración del
 * Foro (5.0/5 en los 4 criterios) a la fecha de redacción.
 */
const DATOS_SINTESIS_GRUPO_2_FEM_ = {
  grupo: "G2",
  tituloInforme: "Informe de Síntesis Grupal — Elementos Comunes y Particularidades Institucionales",
  instituciones: [
    "MARIA CRISTINA ARANGO DE PASTRANA.",
    "LUIS IGNACIO ANDRADE",
    "GABRIEL GARCIA MARQUEZ",
    "EDUARDO SANTOS",
    "MARIA AUXILIADORA FORTALECILLAS",
    "JAIRO MOSQUERA MORENO"
  ],
  responsableInforme: "[Nombre de quien consolida el informe de síntesis del Grupo 2 — SEM Neiva]",
  fechaPresentacion: "7 de septiembre de 2026",
  secciones: [
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta orientadora",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en el logro de los retos y propósitos planteados en el FEM2025?",
      tipo: "cualitativo",
      comun: "Las cinco instituciones del Grupo 2 que a la fecha de este informe han remitido su participación en el Foro (Luis Ignacio Andrade, Gabriel García Márquez, Eduardo Santos, María Auxiliadora Fortalecillas y Jairo Mosquera Moreno) coinciden en reportar avances hacia una educación más pertinente y centrada en la formación integral, con énfasis particular en el fortalecimiento de la atención a la primera infancia y el nivel preescolar: la formalización del ciclo II en Fortalecillas, la apertura y ampliación de jardín y prejardín en Jairo Mosquera Moreno y Eduardo Santos, y el compromiso declarado —aunque con avances aún limitados— de Gabriel García Márquez con la estrategia municipal \"Primera Infancia Feliz y Protegida\". Un segundo eje compartido es el fortalecimiento de proyectos pedagógicos, centros de interés y estrategias de flexibilización curricular orientadas a responder a las características de los estudiantes: los Centros de Interés de Fortalecillas, la recontextualización curricular de Luis Ignacio Andrade y las experiencias basadas en el juego y la exploración del territorio de Jairo Mosquera Moreno. No obstante, el análisis conjunto revela un tercer elemento recurrente: las limitaciones de infraestructura, dotación, recursos didácticos y tecnológicos que condicionan la capacidad institucional de materializar plenamente estos avances, señaladas con particular firmeza por Gabriel García Márquez y Eduardo Santos. Finalmente, se identifica de forma transversal la necesidad de fortalecer la participación y la corresponsabilidad de las familias como condición para consolidar los logros alcanzados.",
      particularidades: "La IE Eduardo Santos se distingue de las demás por expresar de manera explícita su inconformismo frente al que denomina \"abandono\" y la falta de respuesta oportuna de la Secretaría de Educación y la Alcaldía frente a las exigencias institucionales, un señalamiento directo a la administración municipal no formulado en estos términos por ninguna otra institución del grupo. La IE Gabriel García Márquez, por su parte, es la única en vincular explícitamente la seguridad estudiantil en salidas pedagógicas como un requisito institucional que debería formalizarse, y la única en proponer la creación de un comité de alerta temprana para prevenir la deserción escolar."
    },
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta 2",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en la implementación de los nuevos grados del nivel de preescolar (jardín, prejardín)?",
      tipo: "cualitativo",
      comun: "Cuatro de las cinco instituciones del grupo (Luis Ignacio Andrade, Fortalecillas, Jairo Mosquera Moreno y, de manera progresiva, Eduardo Santos) reportan la apertura ya materializada de jardín y, en la mayoría de los casos, también de prejardín, mientras que Gabriel García Márquez es la única institución del grupo en la que esta apertura oficial continúa postergada. Pese a esta diferencia central, existe una coincidencia notable: incluso donde el nivel ya opera, se reconoce la persistencia de limitaciones de infraestructura, dotación y recursos didácticos y tecnológicos —Fortalecillas señala la necesidad de salones propios y mobiliario especializado a mediano plazo, Jairo Mosquera Moreno reclama mayor acompañamiento y dotación de los entes gubernamentales, y Eduardo Santos describe su avance como logrado \"en medio de importantes dificultades\"—, lo que sitúa estas limitaciones materiales como el obstáculo estructural común a todo el grupo, más allá de si el nivel ya se implementó o no. En el plano pedagógico coincide igualmente la apuesta por una educación inicial centrada en el juego, la exploración y la valoración cualitativa del desarrollo infantil, así como la búsqueda de alianzas externas de apoyo, como la articulación de Luis Ignacio Andrade con la Universidad Minuto de Dios para prácticas de primera infancia.",
      particularidades: "Gabriel García Márquez constituye la particularidad más marcada del grupo en esta pregunta: a diferencia de las otras cuatro instituciones, que ya operan el nivel de jardín, esta institución reconoce de manera transparente que la apertura oficial de prejardín y jardín continúa postergada por limitaciones críticas de infraestructura, y plantea la necesidad de un plan inmediato y progresivo con la Secretaría de Educación. Luis Ignacio Andrade, por su parte, es la única institución en reportar una alianza formal con una institución de educación superior para el desarrollo de prácticas pedagógicas en primera infancia."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que los currículos actuales que se desarrollan en las instituciones educativas son pertinentes con sus realidades territoriales (sociales, culturales, productivas)? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Existe consenso entre las cinco instituciones del grupo en calificar la pertinencia curricular como parcial e insuficientemente conectada con las realidades territoriales. Un hallazgo particularmente consistente es la identificación de la presión por los resultados en pruebas estandarizadas —Pruebas Saber, ICFES, DBA— como un factor que limita la flexibilidad curricular necesaria para incorporar el contexto local: Gabriel García Márquez señala que esta presión \"termina uniformando los contenidos\" y relegando la riqueza social y productiva de la región; Jairo Mosquera Moreno indica que la necesidad de responder a lineamientos, estándares y DBA \"valorados mediante pruebas estandarizadas... puede limitar la flexibilidad curricular\"; y Eduardo Santos expresa preocupación porque \"las pruebas nacionales priorizan conocimientos y contenidos y no siempre consideran los procesos ni las particularidades territoriales\". Como segundo elemento común, las cinco instituciones manifiestan la intención de incorporar saberes y dinámicas propias del territorio —la economía local y el patrimonio cultural en Fortalecillas (las achiras, el río, la memoria local), la formación productiva en Eduardo Santos (retomando el proyecto de elaboración de zapatos) y la contextualización social y ambiental en Luis Ignacio Andrade—, aunque reconociendo que estos esfuerzos siguen siendo, en su mayoría, iniciativas puntuales más que una transformación curricular sistemática.",
      particularidades: "Jairo Mosquera Moreno introduce una particularidad relevante para el grupo al señalar que la brecha de pertinencia territorial es más evidente en instituciones con sedes rurales, población dispersa y aulas multigrado, una condición geográfica específica no mencionada por las demás instituciones del grupo. Gabriel García Márquez, por su parte, es la única institución en vincular explícitamente la baja pertinencia curricular con la deserción escolar, al señalar que los estudiantes enfrentan \"la difícil disyuntiva de abandonar la educación formal para ingresar de forma prematura al mercado laboral\" ante la desconexión entre el currículo y las realidades económicas inmediatas."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se han implementado para lograr currículos más pertinentes al territorio?",
      tipo: "cualitativo",
      comun: "El análisis conjunto de las acciones reportadas por las cinco instituciones configura dos líneas de trabajo ampliamente compartidas. La primera es la actualización y contextualización de los planes de estudio y contenidos curriculares, presente en las cinco instituciones bajo distintas modalidades: la depuración y priorización de aprendizajes esenciales en Luis Ignacio Andrade, la integración de lecturas y proyectos de vida ligados al entorno en Gabriel García Márquez, la caracterización permanente de la comunidad educativa en Eduardo Santos, la actualización de planes de estudio liderada por el Consejo Académico en Fortalecillas, y la contextualización de aprendizajes a la vida cotidiana en Jairo Mosquera Moreno. La segunda línea común es el fortalecimiento de proyectos pedagógicos transversales e interdisciplinarios —Plan Lector y Cátedra de Paz en Luis Ignacio Andrade, Aprendizaje Basado en Proyectos en Gabriel García Márquez, proyectos ambientales y olimpiadas en Eduardo Santos, proyectos transversales e investigación escolar en Fortalecillas, y centros de interés como semillas de vida y el club de astronomía y arqueología en Jairo Mosquera Moreno—, como vía principal para conectar los aprendizajes con las realidades del territorio sin renunciar a los estándares nacionales.",
      particularidades: "Fortalecillas y Jairo Mosquera Moreno son las únicas instituciones del grupo en plantear explícitamente el diálogo de saberes con la comunidad como una acción diferenciada —reconociendo a las familias y líderes locales como portadores de conocimiento legítimo (achiras, río y memoria local en Fortalecillas; tradiciones, música y festividades en Jairo Mosquera Moreno)—, mientras que las demás instituciones del grupo no formulan una acción equivalente. Eduardo Santos, por su parte, es la única institución en vincular explícitamente sus acciones curriculares con la prevención del consumo de sustancias psicoactivas y la delincuencia juvenil a través de las olimpiadas deportivas, culturales y académicas."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 3",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar estas acciones?",
      tipo: "mixto",
      totalIE: 5,
      tally: [
        {opcion: "Consejo Académico", count: 5},
        {opcion: "Comités de área", count: 4},
        {opcion: "Consejo Directivo", count: 4},
        {opcion: "Comisión de Evaluación y Promoción", count: 4},
        {opcion: "Consejo Estudiantil", count: 3},
        {opcion: "Personero(a) Estudiantil", count: 3},
        {opcion: "Equipo de Gestión Académica", count: 3}
      ],
      comun: "El Consejo Académico es el único equipo de trabajo que las cinco instituciones del grupo, sin excepción, identifican como responsable de liderar las acciones hacia currículos más pertinentes. Con una adopción del 80% (4 de 5 IE) se ubican los comités de área, el Consejo Directivo y la Comisión de Evaluación y Promoción; y con una adopción del 60% (3 de 5 IE), el Consejo Estudiantil, el/la Personero(a) Estudiantil y el Equipo de Gestión Académica. Este patrón muestra que, salvo por el Consejo Académico —prácticamente universal—, el grupo presenta una estructura de equipos de trabajo curricular más heterogénea que la observada en otras preguntas mixtas del Foro, con una parte relevante de las instituciones apoyándose en estructuras propias o de carácter más específico en lugar de las instancias reglamentarias estándar.",
      particularidades: "Gabriel García Márquez es, con diferencia, la institución con menor selección de equipos del catálogo estándar en el grupo: únicamente marcó Consejo Académico y comités de área, complementando con una opción \"Otro\" (Comité de Convivencia Escolar y Áreas de Gestión), sin reportar Consejo Directivo ni ninguna instancia de representación estudiantil para esta pregunta. En el extremo opuesto, Eduardo Santos es la institución que más se aparta del catálogo estándar por la vía contraria: además de seis equipos reglamentarios, reportó diez opciones \"Otro\" de carácter completamente institucional y específico (Mesa Técnica de Inclusión, COPAS, Olimpiadas Escolares, Proyectos de cuidado de Medioambiente, Simulacros de las Pruebas Nacionales —Proyecto PEMSAR—, Caracterización de la Comunidad, entre otros), una proliferación de equipos propios muy superior a la de cualquier otra institución del grupo."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 4",
      enunciado: "¿Cómo se están articulando estos equipos de trabajo para lograr currículos más pertinentes territorialmente?",
      tipo: "cualitativo",
      comun: "Las cinco instituciones del grupo coinciden en que la articulación entre equipos de trabajo se apoya principalmente en el Consejo Académico, los comités de área y los espacios de reunión y planeación institucional, canalizando allí las propuestas curriculares para su análisis y aprobación. De manera igualmente compartida —con la única excepción parcial de Jairo Mosquera Moreno, que describe el proceso en términos más consolidados—, las instituciones reconocen que esta articulación resulta todavía insuficiente: Luis Ignacio Andrade señala la necesidad de \"fortalecer el trabajo colaborativo y construir una visión curricular común\"; Gabriel García Márquez advierte que la planeación \"se enfoca prioritariamente en el cumplimiento de los DBA y en la preparación para las pruebas ICFES, dejando de lado la integración efectiva del enfoque territorial\"; Eduardo Santos plantea como \"fundamental fortalecer el trabajo colectivo para evitar acciones aisladas\"; y Fortalecillas identifica como reto \"optimizar los tiempos de planeación docentes\" y \"garantizar decisiones participativas\". Este patrón sugiere que, en el Grupo 2, la brecha entre la existencia formal de espacios de articulación y su funcionamiento efectivo es un desafío ampliamente compartido.",
      particularidades: "Jairo Mosquera Moreno se distingue del resto del grupo por describir una articulación curricular apoyada en un número particularmente amplio de proyectos transversales institucionalizados (PRAE, PESCC, Derechos Humanos, Paz y Democracia, Prevención del Riesgo, Estilos de Vida Saludable y Mejoramiento de Logros), una estructura más diversificada que la reportada por las demás instituciones del grupo. Fortalecillas, por su parte, es la única institución en mencionar el diálogo de saberes locales (cultura, achiras, río) como parte explícita de su mecanismo de articulación curricular."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 5",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para que dichas acciones se cumplan?",
      tipo: "mixto",
      totalIE: 5,
      tally: [
        {opcion: "Análisis de resultados de Pruebas Saber e ICFES", count: 3},
        {opcion: "Consejo Académico", count: 3},
        {opcion: "Comités de área o núcleos de formación", count: 3},
        {opcion: "Actas de reunión de área", count: 3},
        {opcion: "Autoevaluación institucional anual", count: 3},
        {opcion: "Articulación con el PEI", count: 2},
        {opcion: "Plataformas de gestión académica (JIGRA)", count: 2}
      ],
      comun: "A diferencia de otras preguntas del grupo, en los mecanismos de seguimiento a las acciones curriculares no se observa un elemento adoptado por la totalidad de las instituciones: el análisis de resultados de Pruebas Saber e ICFES, el Consejo Académico, los comités de área o núcleos de formación, las actas de reunión de área y la autoevaluación institucional anual son, cada uno, los mecanismos más compartidos, presentes en 3 de las 5 instituciones (60%). Con una adopción del 40% (2 de 5 IE) se ubican la articulación con el PEI y las plataformas de gestión académica (JIGRA). Este patrón de dispersión —más marcado que en el resto de preguntas mixtas del grupo— refleja que el seguimiento a las acciones curriculares en el Grupo 2 depende todavía, en buena medida, de mecanismos propios de cada institución más que de un conjunto de instrumentos comunes y estandarizados.",
      particularidades: "Gabriel García Márquez vuelve a ser la institución más alejada del catálogo estándar: solo seleccionó autoevaluación institucional anual y complementó con una opción \"Otro\" describiendo el análisis de la eficiencia académica al finalizar cada período escolar como su único mecanismo adicional. Eduardo Santos, en cambio, seleccionó únicamente actas de reunión de área del catálogo, apoyándose de manera extensa en una respuesta \"Otro\" que describe el seguimiento mediante la revisión de planes de aula y evaluaciones institucionales anuales, y que propone explícitamente extender este seguimiento más allá de los docentes del Decreto 1278 — una observación normativa específica no planteada por ninguna otra institución del grupo."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que la toma de decisiones en las instituciones educativas actualmente es participativa y democrática? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Existe una coincidencia sustancial entre las cinco instituciones del grupo: todas reconocen la existencia de estructuras formales de gobierno escolar (Consejo Directivo, Consejo Académico, Consejo Estudiantil, Personero, Contralor) amparadas en la Ley 115, que en principio garantizan cauces democráticos, pero todas matizan esa existencia formal señalando que la participación real continúa siendo limitada. Gabriel García Márquez resume esta tensión al señalar que estos espacios operan \"bajo un enfoque instrumental o de cumplimiento meramente normativo\"; Fortalecillas advierte que las decisiones estratégicas \"se concentran en la directiva, relegando la participación a un rol informativo o consultivo\"; y Jairo Mosquera Moreno concluye que la participación \"aún no es plenamente incluyente e incidente\". La baja participación e interés de las familias es señalada de manera recurrente como un obstáculo específico —Luis Ignacio Andrade menciona el \"bajo interés de algunos padres de familia\" y Eduardo Santos observa que \"algunos padres muestran poco interés... y suelen participar los mismos representantes año tras año\"—, mientras que la participación estudiantil, aunque valorada positivamente en cuanto a espacios disponibles, se describe como cohibida en la práctica o limitada a la elección de representantes sin mayor incidencia posterior.",
      particularidades: "Fortalecillas se distingue del resto del grupo por proponer una ruta metodológica propia y estructurada para avanzar hacia una democracia escolar real —\"Participar para decidir\", de cinco pasos (escuchar, priorizar, tramitar, decidir mediante acuerdos vinculantes y retroalimentar)—, una herramienta de gestión participativa sin equivalente en ninguna otra institución del grupo. Eduardo Santos, por su parte, es la única institución en incorporar explícitamente al sector productivo y empresarial como actor con representación formal en el Consejo Directivo."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se están implementando para canalizar y fortalecer la participación de la comunidad educativa?",
      tipo: "cualitativo",
      comun: "El análisis conjunto de las acciones reportadas por las cinco instituciones evidencia dos estrategias ampliamente compartidas. La primera es la apertura de espacios de diálogo, consulta y encuentro con la comunidad educativa —asambleas comunitarias y escuelas de padres en Luis Ignacio Andrade, mesas territoriales de codiseño y escuelas de familias en Gabriel García Márquez, asambleas y consultas comunitarias mediante encuestas y buzones de sugerencias en Fortalecillas, y espacios de diálogo y escucha en Jairo Mosquera Moreno—, orientada a recoger inquietudes y propuestas más allá de las reuniones estrictamente informativas. La segunda es el fortalecimiento del liderazgo y la representación estudiantil, presente de manera explícita en Gabriel García Márquez (su Semillero de Veedores Estudiantiles), Eduardo Santos (fortalecimiento del liderazgo estudiantil y comunitario) y Jairo Mosquera Moreno (fortalecimiento de la representación a través de Personero, Contralor y Consejo Estudiantil). El fortalecimiento de canales de comunicación institucional —redes sociales y aplicaciones de mensajería— aparece también como una estrategia compartida por Luis Ignacio Andrade y Eduardo Santos.",
      particularidades: "Gabriel García Márquez presenta la particularidad más singular del grupo en esta pregunta: su Semillero de Veedores Estudiantiles articula explícitamente a la institución con la Personería Municipal para capacitar a sus representantes en veeduría sobre la ejecución del PAE y las compras escolares, una alianza interinstitucional de control social sin equivalente en ninguna otra institución del grupo. Eduardo Santos, por su parte, es la única institución en reportar un sistema de motivación mediante puntos y recompensas dentro del aula como estrategia de participación, y la única en articular sus acciones con un proyecto macro institucional nombrado (\"Soy Constructor de mi Vida\")."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Equipos de trabajo",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar las estrategias y mecanismos de participación escolar?",
      tipo: "mixto",
      totalIE: 5,
      tally: [
        {opcion: "Gobierno Escolar", count: 4},
        {opcion: "Consejo Directivo", count: 4},
        {opcion: "Consejo Académico", count: 4},
        {opcion: "Consejo Estudiantil", count: 4},
        {opcion: "Personero Estudiantil", count: 4},
        {opcion: "Contralor Estudiantil", count: 3},
        {opcion: "Consejo de Padres de Familia", count: 4},
        {opcion: "Comité Escolar de Convivencia", count: 3},
        {opcion: "Comisión de Evaluación y Promoción", count: 3}
      ],
      comun: "Cuatro de las cinco instituciones del grupo (Luis Ignacio Andrade, Gabriel García Márquez, Fortalecillas y Jairo Mosquera Moreno) coinciden en identificar el Gobierno Escolar en su conjunto, el Consejo Directivo, el Consejo Académico, el Consejo Estudiantil, el/la Personero(a) Estudiantil y el Consejo de Padres de Familia como los equipos que lideran las estrategias de participación escolar (80% del grupo en cada caso). Con una adopción del 60% (3 de 5 IE) se ubican el/la Contralor(a) Estudiantil, el Comité Escolar de Convivencia y la Comisión de Evaluación y Promoción. Este patrón replica, con leves variaciones, la estructura de gobierno escolar reglamentaria observada en la pregunta análoga de la Sesión 2, confirmando que las instancias formales de representación estudiantil y de gobierno escolar constituyen el eje común de participación en la mayoría de instituciones del grupo.",
      particularidades: "Eduardo Santos constituye la particularidad más marcada de esta pregunta en todo el grupo: en lugar de seleccionar opciones del catálogo estándar, describió una estructura propia y extensa en texto libre, detallando responsables específicos por nombre para cada frente (un Comité de Ciencias Sociales encargado del seguimiento al Gobierno Escolar, formación técnica a través del SENA a cargo de una docente nombrada, un equipo de gestión para prácticas pedagógicas con la universidad, y asesorías de orientación vocacional con responsables igualmente identificados), lo que impide comparar directamente su respuesta con el catálogo compartido por el resto del grupo. Gabriel García Márquez, por su parte, es la única institución que no reporta ni Contralor(a) Estudiantil ni Comité Escolar de Convivencia entre sus equipos de participación."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Mecanismos de seguimiento",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para garantizar las acciones encaminadas a promover gobiernos educativos democráticos?",
      tipo: "mixto",
      totalIE: 5,
      tally: [
        {opcion: "Autoevaluación institucional anual (Guía 34 – MEN)", count: 5},
        {opcion: "Plan de Mejoramiento Institucional (PMI)", count: 4},
        {opcion: "Seguimiento periódico al Plan de Mejoramiento", count: 4},
        {opcion: "Reuniones periódicas con entes de gobierno escolar", count: 3},
        {opcion: "Elecciones estudiantiles", count: 3},
        {opcion: "Elección de docentes para consejo directivo", count: 3},
        {opcion: "Elección de docentes para consejo académico", count: 3},
        {opcion: "Actualización y ajuste permanente del PEI", count: 3}
      ],
      comun: "La autoevaluación institucional anual conforme a la Guía 34 del MEN es el único mecanismo de seguimiento que las cinco instituciones del grupo, sin excepción, reportan haber implementado para garantizar las acciones hacia gobiernos educativos democráticos. Con una adopción del 80% (4 de 5 IE) se ubican el Plan de Mejoramiento Institucional y su seguimiento periódico; y con una adopción del 60% (3 de 5 IE), las reuniones periódicas con los entes de gobierno escolar, las elecciones estudiantiles, la elección de docentes para los consejos directivo y académico, y la actualización permanente del PEI. Este patrón muestra que, más allá de la autoevaluación institucional —prácticamente universal en el grupo—, existe una diferenciación real en la solidez de los mecanismos de seguimiento entre las instituciones.",
      particularidades: "Gabriel García Márquez constituye el hallazgo más relevante de esta pregunta para todo el grupo: además de ser la única institución en seleccionar un único mecanismo del catálogo (autoevaluación institucional anual), reconoce explícitamente en su respuesta que \"no se están implementando acciones de seguimiento para garantizar las acciones encaminadas a promover gobiernos educativos democráticos\", una admisión directa de vacío institucional que amerita atención específica de la Secretaría de Educación. Fortalecillas, en el extremo opuesto, es la institución con el mecanismo de seguimiento más robusto y diversificado del grupo, siendo la única en reportar una verificación externa del funcionamiento del Gobierno Escolar por parte de la propia Secretaría de Educación Municipal."
    }
  ],
  conclusiones: "El análisis conjunto de las tres sesiones de trabajo del Grupo 2 permite concluir que, entre las cinco instituciones que a la fecha de este informe han remitido su participación en el Foro, existe una trayectoria compartida de avances reales en la atención a la primera infancia, el fortalecimiento de proyectos pedagógicos transversales y la consolidación formal de estructuras de gobierno escolar, junto con retos comunes en materia de infraestructura, pertinencia curricular frente a la presión de las pruebas estandarizadas, y participación real —más allá de lo formal— de las familias y los estudiantes. Dos hallazgos particulares merecen seguimiento específico de la Secretaría de Educación Municipal: la IE Gabriel García Márquez, que en la pregunta sobre mecanismos de seguimiento de la Sesión 3 reconoció explícitamente no tener implementadas acciones de seguimiento para garantizar gobiernos educativos democráticos, y cuya apertura de prejardín y jardín sigue postergada por limitaciones de infraestructura; y la IE Eduardo Santos, cuya respuesta a la pregunta sobre equipos de trabajo de la Sesión 3 se apartó completamente del catálogo estándar mediante una descripción propia y detallada de responsables institucionales. En cuanto a la valoración general del Foro, de las cinco instituciones del grupo con informe remitido, únicamente Fortalecillas registró formalmente su valoración de la jornada, con una calificación de 5,0 sobre 5,0 en los cuatro criterios evaluados (diálogo y reflexión, participación, ideas y propuestas, y satisfacción con el instrumento), sin comentarios o sugerencias adicionales registrados; las cuatro instituciones restantes (Luis Ignacio Andrade, Gabriel García Márquez, Eduardo Santos y Jairo Mosquera Moreno) aún no han registrado su valoración, y la IE María Cristina Arango de Pastrana, sexta institución del grupo, no había remitido su informe del Foro a la fecha de consolidación de este documento. Se recomienda a la Secretaría de Educación gestionar tanto el diligenciamiento pendiente de la valoración con las cuatro instituciones señaladas, como el seguimiento a la IE María Cristina Arango de Pastrana para completar la participación del grupo.",
  proyectoNombre: "[Nombre de quien consolida el Informe Consolidado]",
  proyectoCargo: "[Cargo]",
  fechaRealizacion: "7 de septiembre de 2026"
};

/*
 * Genera (o reescribe, si ya existe) el Informe de Síntesis Grupal del
 * Grupo 2, dejándolo en la raíz de la carpeta "Grupo G2" (mismo
 * criterio de idempotencia que el resto de documentos de grupo).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El enlace del documento queda en "Ver
 * registros de ejecución".
 */
function generarInformeSintesisGrupo2FEM(){
  const carpetas=crearEstructuraCarpetasGrupoFEM_("G2");
  const nombreDoc="Informe de Síntesis - Grupo G2 FEM 2026";
  const existentesIt=carpetas.grupoFolder.getFilesByName(nombreDoc);
  const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
  while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

  const archivoDoc=generarInformeSintesisGrupoFEM_(DATOS_SINTESIS_GRUPO_2_FEM_, archivoExistente?archivoExistente.getId():null);
  if(!archivoExistente){
    carpetas.grupoFolder.addFile(archivoDoc);
    try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
  }

  Logger.log("========================================");
  Logger.log("INFORME DE SÍNTESIS — GRUPO 2 — RESULTADO");
  Logger.log("Documento: "+archivoDoc.getUrl());
  Logger.log("========================================");

  return archivoDoc.getUrl();
}

/*
 * INFORME DE SÍNTESIS GRUPAL — GRUPO 3.
 *
 * Mismo criterio que los Grupos 1 y 2: prosa redactada a mano leyendo
 * directamente el Doc editable ya generado de cada IE del grupo, y
 * tallas de las 4 preguntas mixtas contadas a mano sobre las opciones
 * realmente seleccionadas.
 *
 * A diferencia del Grupo 2, las 6 IE del catálogo (Técnico Superior,
 * Departamental Tierra de Promisión, Santa Librada, Ricardo Borrero
 * Álvarez, Ángel María Paredes, Ceinar) SÍ habían remitido su Informe
 * Ejecutivo a la fecha de redacción (7 de septiembre de 2026), por lo
 * que el denominador de las 4 preguntas mixtas es 6 (no 5).
 *
 * De esas 6 IE, solo Departamental Tierra de Promisión había registrado
 * su Valoración del Foro (5.0/5 en los 4 criterios) a la fecha de
 * redacción.
 */
const DATOS_SINTESIS_GRUPO_3_FEM_ = {
  grupo: "G3",
  tituloInforme: "Informe de Síntesis Grupal — Elementos Comunes y Particularidades Institucionales",
  instituciones: [
    "TECNICO SUPERIOR",
    "DEPARTAMENTAL TIERRA DE PROMISION",
    "SANTA LIBRADA",
    "RICARDO BORRERO ALVAREZ",
    "ANGEL MARIA PAREDES",
    "CEINAR"
  ],
  responsableInforme: "[Nombre de quien consolida el informe de síntesis del Grupo 3 — SEM Neiva]",
  fechaPresentacion: "7 de septiembre de 2026",
  secciones: [
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta orientadora",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en el logro de los retos y propósitos planteados en el FEM2025?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del Grupo 3 (Técnico Superior, Departamental Tierra de Promisión, Santa Librada, Ricardo Borrero Álvarez, Ángel María Paredes y Ceinar) coinciden en reportar avances significativos en la contextualización curricular y el fortalecimiento de proyectos pedagógicos e institucionales propios: la articulación con el SENA y universidades en Técnico Superior y Tierra de Promisión, los centros de interés y proyectos tecnológicos y ambientales de Santa Librada (CI-LEO, MakeCode, robótica, ECOTURO, PRAE), la flexibilización curricular mediante DUA y PIAR en Ricardo Borrero Álvarez, la consolidación de la educación inicial y el énfasis en emprendimiento en Ángel María Paredes, y el hilo conductor institucional articulado a CEINARTE en Ceinar. De manera igualmente compartida, las seis instituciones identifican las limitaciones de infraestructura, dotación, conectividad y recursos tecnológicos como el obstáculo estructural más persistente para consolidar estos avances —desde el bloque pendiente de demolición en Técnico Superior hasta la ausencia de computadores y conectividad en Tierra de Promisión—, y coinciden en señalar la necesidad de fortalecer la corresponsabilidad y la participación permanente de las familias en los procesos formativos.",
      particularidades: "Ricardo Borrero Álvarez presenta la particularidad más vívida del grupo en esta pregunta: describe problemáticas de salubridad concretas —una infestación de palomas que deteriora cielos rasos y bebederos en las sedes Oriente y Central— y un descontento comunitario explícito con la calidad del Programa de Alimentación Escolar y la cobertura del transporte escolar, junto con preocupaciones de seguridad por riñas y microtráfico, un nivel de detalle sobre condiciones materiales adversas no replicado por ninguna otra institución del grupo. Ceinar, por su parte, es la única institución en reportar una estrategia cultural-comunitaria propia y nombrada —la \"Cumbre para el liderazgo por la paz, el arte y la cultura\"— que vincula a la organización comunal y a la secretaría de gobierno municipal en la discusión sobre el papel de la educación."
    },
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta 2",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en la implementación de los nuevos grados del nivel de preescolar (jardín, prejardín)?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del grupo coinciden en reportar la implementación ya en marcha del grado jardín, con distintos niveles de consolidación del prejardín: Tierra de Promisión y Ricardo Borrero Álvarez cuentan con oferta integrada de varios grupos y sedes, mientras que Santa Librada, Ángel María Paredes y Técnico Superior reportan avances más recientes y Ceinar enfrenta la cobertura más baja del grupo. Pese a esta implementación generalizada, existe una coincidencia notable en que la infraestructura y la dotación continúan siendo insuficientes incluso donde el nivel ya opera: Técnico Superior señala la necesidad de adecuar baterías sanitarias y mobiliario, Santa Librada reconoce que las limitaciones de infraestructura condicionan la ampliación de la oferta, y Ricardo Borrero Álvarez identifica la urgencia de renovar una red eléctrica obsoleta. El hallazgo cuantitativo más contundente del grupo en esta pregunta proviene de Ángel María Paredes, donde apenas el 27,1% de los consultados en su ejercicio interno de consulta considera que los recursos disponibles son suficientes y oportunos, frente a un 51% que los califica de insuficientes o parcialmente suficientes.",
      particularidades: "Ceinar constituye la particularidad más marcada del grupo en esta pregunta: a diferencia de las demás instituciones, que describen una demanda de matrícula creciente o estable, Ceinar identifica una desventaja competitiva estructural frente a los Centros de Desarrollo Infantil (CDI) del sector, que ofrecen alimentación y permanencia más prolongada con apoyo de un equipo interdisciplinar del que la institución carece, lo que explica la ausencia de solicitudes de cupo para prejardín. Ricardo Borrero Álvarez, por su parte, es la única institución en señalar explícitamente el tránsito normativo de la evaluación por Derechos Básicos de Aprendizaje (DBA) hacia propósitos de desarrollo como un cambio que obligó a una reestructuración curricular profunda del nivel preescolar durante 2026."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que los currículos actuales que se desarrollan en las instituciones educativas son pertinentes con sus realidades territoriales (sociales, culturales, productivas)? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del grupo coinciden en calificar la pertinencia curricular como un proceso en construcción, con avances reales pero todavía dependientes, en varios casos, de iniciativas puntuales más que de una transformación institucional sistemática: Santa Librada reconoce que muchas experiencias \"dependen principalmente de la iniciativa de determinados docentes\"; Ricardo Borrero Álvarez advierte sobre una \"estructura curricular fragmentada y rígida por asignaturas disciplinares tradicionales\"; y Ángel María Paredes señala como reto pendiente \"la articulación entre áreas\" pese a que el 77,1% de sus consultados considera pertinente el currículo actual. Como segundo elemento compartido, las seis instituciones recurren a alianzas externas —con el SENA (Técnico Superior, Ricardo Borrero Álvarez), con universidades regionales (Universidad de Navarra y Universidad Surcolombiana) o con actores privados como la Fundación Terpel (Ceinar, a través del programa Escuelas que Aprenden)— como estrategia central para fortalecer la pertinencia curricular frente a las dinámicas sociales y productivas del territorio.",
      particularidades: "Ricardo Borrero Álvarez introduce una particularidad relevante para el grupo al vincular explícitamente su reflexión curricular con la atención a una \"comunidad estudiantil flotante\", una condición de movilidad poblacional no mencionada por ninguna otra institución del grupo. Técnico Superior, por su parte, es la única institución en fundamentar su pertinencia curricular en una oferta de nueve especialidades técnicas articuladas con el sector productivo, y la única en mencionar explícitamente la necesidad de incorporar Inteligencia Artificial y tecnologías 4.0 a sus contenidos disciplinares."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se han implementado para lograr currículos más pertinentes al territorio?",
      tipo: "cualitativo",
      comun: "Al observar de manera conjunta las acciones reportadas por las seis instituciones, la línea de trabajo más ampliamente compartida es el fortalecimiento de proyectos pedagógicos transversales que vinculan los aprendizajes con el territorio: los proyectos de aula anclados en el contexto socioproductivo de Técnico Superior, el Proyecto Ambiental Escolar de Tierra de Promisión, los centros de interés de Santa Librada, los proyectos interdisciplinares de Ricardo Borrero Álvarez, la Feria de Emprendimiento y el PRAE de Ángel María Paredes, y el hilo conductor institucional articulado a CEINARTE en Ceinar. Una segunda línea compartida por cuatro de las seis instituciones es el establecimiento de alianzas externas concretas como mecanismo de contextualización curricular: convenios con el SENA y cuatro universidades en Técnico Superior, la alianza con la Universidad Surcolombiana en Tierra de Promisión, la vinculación de entidades externas especializadas para simulacros de pruebas en Ricardo Borrero Álvarez, y la articulación con la Fundación Terpel a través del programa Escuelas que Aprenden en Ceinar.",
      particularidades: "Ricardo Borrero Álvarez es la única institución del grupo en reportar una estrategia de asignación de roles a los propios estudiantes —como líderes y guardianes ambientales, de convivencia, de TIC y de rendimiento académico— combinada con espacios de autocapacitación interna del personal docente, una acción de apropiación estudiantil del currículo sin equivalente en las demás instituciones. Técnico Superior, por su parte, es la única institución en vincular explícitamente sus acciones curriculares con un programa de intensificación en inglés como estrategia de flexibilización."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 3",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar estas acciones?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Consejo Académico", count: 6},
        {opcion: "Comités de área", count: 6},
        {opcion: "Consejo Directivo", count: 4},
        {opcion: "Equipo de Educación Inclusiva / NEE", count: 4},
        {opcion: "Equipo de Preescolar", count: 3},
        {opcion: "Consejo Estudiantil", count: 3},
        {opcion: "Consejo de Padres de Familia", count: 3}
      ],
      comun: "El Consejo Académico y los comités de área son los dos equipos de trabajo que las seis instituciones del grupo, sin excepción, identifican como responsables de liderar las acciones hacia currículos más pertinentes territorialmente. Con una adopción del 66,7% (4 de 6 IE) se ubican el Consejo Directivo y el Equipo de Educación Inclusiva o NEE; y con una adopción del 50% (3 de 6 IE), el Equipo de Preescolar, el Consejo Estudiantil y el Consejo de Padres de Familia. Este patrón evidencia una base común muy sólida en torno a las dos instancias más reglamentarias del gobierno curricular, sobre la cual cada institución añade equipos adicionales de acuerdo con su propia estructura organizativa.",
      particularidades: "Santa Librada es la institución con la estructura de equipos más granular y diferenciada por nivel educativo del grupo, al reportar equipos separados de preescolar, básica primaria y básica secundaria y media, además de un Equipo de Sistematización de la Información sin equivalente en ninguna otra institución. Técnico Superior, en el extremo opuesto, es la única institución que no reporta ninguna instancia de representación estudiantil (Consejo Estudiantil, Personería) entre los equipos de esta pregunta, concentrando su respuesta en tres equipos reglamentarios y un conjunto combinado de \"Equipos Líderes\" de carácter curricular y técnico."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 4",
      enunciado: "¿Cómo se están articulando estos equipos de trabajo para lograr currículos más pertinentes territorialmente?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del grupo coinciden en que la articulación entre equipos de trabajo se canaliza principalmente a través de reuniones periódicas, comités de área y jornadas pedagógicas de planeación institucional, con el Consejo Académico como instancia orientadora central en la mayoría de los casos. Tres de las seis instituciones —Santa Librada, Ricardo Borrero Álvarez y Ángel María Paredes— matizan esta descripción reconociendo que la articulación efectiva sigue siendo un reto: Santa Librada advierte que \"todavía es un reto lograr que esta articulación sea permanente y que no dependa solamente de iniciativas individuales\"; Ricardo Borrero Álvarez señala \"falta de tiempo operativo\" para optimizar el seguimiento; y Ángel María Paredes propone como oportunidad de mejora \"fortalecer la coordinación entre equipos\" y \"consolidar mecanismos de seguimiento con evidencias, indicadores y periodicidad\". Técnico Superior, Tierra de Promisión y Ceinar, en cambio, describen sus mecanismos de articulación en términos más procedimentales, sin formular una autocrítica equivalente.",
      particularidades: "Ricardo Borrero Álvarez es la única institución del grupo en describir una distribución explícita de roles entre sus equipos de articulación curricular —el Consejo Académico lidera la transformación curricular, los comités de área diseñan las propuestas de aula y las familias aportan insumos territoriales— y en nombrar canales de comunicación digital específicos (WhatsApp) junto a los formales (circulares de rectoría). Técnico Superior, por su parte, es la única institución en mencionar explícitamente reuniones de seguimiento entre la coordinación de convivencia y los representantes estudiantiles como parte de su mecanismo de articulación."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 5",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para que dichas acciones se cumplan?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Articulación con el PEI", count: 5},
        {opcion: "Consejo Académico", count: 4},
        {opcion: "Comités de área o núcleos de formación", count: 4},
        {opcion: "Actas de reunión de área", count: 4},
        {opcion: "Autoevaluación institucional anual", count: 4},
        {opcion: "Plataformas de gestión académica (JIGRA)", count: 3},
        {opcion: "Comparación histórica de resultados internos y externos", count: 3}
      ],
      comun: "La articulación con el PEI es, en esta pregunta, el mecanismo de seguimiento más ampliamente compartido por el grupo, presente en cinco de las seis instituciones (83,3%). Con una adopción del 66,7% (4 de 6 IE) se ubican el Consejo Académico, los comités de área, las actas de reunión de área y la autoevaluación institucional anual; y con una adopción del 50% (3 de 6 IE), las plataformas de gestión académica JIGRA y la comparación histórica de resultados internos y externos. Este patrón confirma al PEI como el instrumento articulador central del seguimiento curricular en el Grupo 3, complementado de manera desigual por instancias colegiadas y herramientas de análisis de resultados.",
      particularidades: "Tierra de Promisión y Ceinar son las dos instituciones del grupo que más se apartan del catálogo estándar de mecanismos, respondiendo en su mayoría con opciones \"Otro\" de redacción propia: Tierra de Promisión describe un fortalecimiento académico centrado en pruebas Saber, pensamiento crítico y permanencia escolar, mientras que Ceinar reporta la elaboración de informes de gestión como su mecanismo adicional. Ricardo Borrero Álvarez, por su parte, es la única institución en incluir portafolios docentes y la tasa de deserción o reprobación por asignatura como mecanismos explícitos de seguimiento curricular."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que la toma de decisiones en las instituciones educativas actualmente es participativa y democrática? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Existe una coincidencia sustancial entre las seis instituciones del grupo: todas afirman que la toma de decisiones cuenta con una estructura formal participativa y democrática, sustentada en el funcionamiento del Consejo Directivo, el Consejo Académico, el Consejo Estudiantil y el Consejo de Padres, y todas matizan esa afirmación señalando límites concretos a la incidencia real. La baja participación de las familias es, con diferencia, el obstáculo más recurrente del grupo, mencionado explícitamente por cinco de las seis instituciones: Tierra de Promisión señala que la participación de los padres \"se limita a la elección de representantes, quienes presentan poca asistencia\"; Santa Librada reconoce \"baja asistencia de padres de familia\"; Ricardo Borrero Álvarez atribuye esto a \"apatía y falta de sentido de pertenencia de algunos acudientes\"; Ángel María Paredes reconoce \"dificultades laborales y familiares que pueden limitar su asistencia\"; y Ceinar admite \"falencias vinculadas a la participación de los padres de familia, específicamente el consejo de padres\". Un segundo elemento compartido, explícito en Ricardo Borrero Álvarez y Ángel María Paredes, es la advertencia de que la participación no debe limitarse a la consulta formal, sino traducirse en una incidencia real sobre las decisiones institucionales.",
      particularidades: "Ricardo Borrero Álvarez presenta la particularidad más crítica del grupo en esta pregunta, al señalar que las decisiones académicas o administrativas clave \"se centralizan en el Equipo de Gestión sin comunicarse oportunamente a las bases\", una denuncia de opacidad en la comunicación de decisiones no formulada en estos términos por ninguna otra institución. Ceinar, por su parte, es la única institución en mencionar explícitamente la vinculación de veedores externos —en su caso del sector productivo y comunitario a través de la Cumbre— a sus espacios de participación democrática."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se están implementando para canalizar y fortalecer la participación de la comunidad educativa?",
      tipo: "cualitativo",
      comun: "El análisis conjunto de las acciones reportadas por las seis instituciones evidencia dos líneas de trabajo compartidas. La primera es el fortalecimiento de los órganos del Gobierno Escolar como vehículo de participación: Técnico Superior lo denomina explícitamente \"Gobierno Escolar Activo\", Ricardo Borrero Álvarez desarrolla elecciones democráticas de Personero, Contralor y representantes, Ángel María Paredes fortalece sus órganos de representación, y Ceinar acompaña sus propios órganos de participación estudiantil. La segunda línea, presente con particular fuerza en Tierra de Promisión, Técnico Superior y Ceinar, es el establecimiento de alianzas interinstitucionales con entidades gubernamentales, de protección y de educación superior —ICBF, Policía de Infancia, Comisarías de Familia, Defensoría del Pueblo, universidades regionales— orientadas a fortalecer la convivencia, la orientación vocacional y la proyección comunitaria del Foro más allá de los espacios escolares.",
      particularidades: "Ceinar es la única institución del grupo en reportar un proyecto de proyección territorial con incidencia política explícita —\"La Minga\", que vincula a la institución con el Plan de Ordenamiento Territorial (POT) y con los consejos municipal y consultivo—, una articulación entre participación escolar y agenda pública municipal sin equivalente en las demás instituciones. Tierra de Promisión, por su parte, es la institución que reporta el mayor número de alianzas interinstitucionales de protección y convivencia del grupo (incluyendo Justicia Especial de Paz, Defensoría del Pueblo y Procuraduría), un nivel de articulación interinstitucional notablemente más amplio que el de las demás instituciones."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Equipos de trabajo",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar las estrategias y mecanismos de participación escolar?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Comité Escolar de Convivencia", count: 6},
        {opcion: "Comité de Calidad / Equipo de Gestión Institucional", count: 6},
        {opcion: "Gobierno Escolar", count: 5},
        {opcion: "Consejo Directivo", count: 5},
        {opcion: "Consejo Académico", count: 5},
        {opcion: "Consejo Estudiantil", count: 5},
        {opcion: "Consejo de Padres de Familia", count: 5},
        {opcion: "Comisión de Evaluación y Promoción", count: 5},
        {opcion: "Personero(a) Estudiantil", count: 4},
        {opcion: "Contralor(a) Estudiantil", count: 4}
      ],
      comun: "El Comité Escolar de Convivencia y el Comité de Calidad o Equipo de Gestión Institucional son los dos equipos que las seis instituciones del grupo, sin excepción, reportan haber conformado para liderar las estrategias de participación escolar. Con una adopción del 83,3% (5 de 6 IE) se ubican el Gobierno Escolar en su conjunto, el Consejo Directivo, el Consejo Académico, el Consejo Estudiantil, el Consejo de Padres de Familia y la Comisión de Evaluación y Promoción; y con una adopción del 66,7% (4 de 6 IE), el/la Personero(a) y el/la Contralor(a) Estudiantil. Este patrón evidencia la estructura de gobierno escolar más homogénea y consolidada observada hasta ahora entre los grupos analizados, con una adopción casi universal de la práctica totalidad del catálogo reglamentario.",
      particularidades: "Ceinar constituye, con diferencia, la particularidad más marcada del grupo en esta pregunta: es la única institución que no reporta el Consejo Directivo, el Consejo Académico, el Consejo Estudiantil ni el Consejo de Padres de Familia entre sus equipos de participación, apoyándose en cambio en una estructura propia de cuatro equipos vinculados a sus proyectos institucionales (entorno y sociedad, proyectos transversales, Escuelas que Aprenden con la Fundación Terpel, y reestructuración del PEI). Tierra de Promisión, por su parte, es la única institución que no reporta Gobierno Escolar como tal ni las figuras de Personero(a) o Contralor(a) Estudiantil, complementando su selección con un Comité de Prevención y Atención del Riesgo exclusivo de esta institución."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Mecanismos de seguimiento",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para garantizar las acciones encaminadas a promover gobiernos educativos democráticos?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Reuniones periódicas con entes de gobierno escolar", count: 6},
        {opcion: "Autoevaluación institucional anual (Guía 34 – MEN)", count: 5},
        {opcion: "Elecciones estudiantiles", count: 4},
        {opcion: "Elección de docentes para consejo directivo", count: 4},
        {opcion: "Elección de docentes para consejo académico", count: 4},
        {opcion: "Plan de Mejoramiento Institucional (PMI)", count: 4},
        {opcion: "Actualización y ajuste permanente del PEI", count: 4}
      ],
      comun: "Las reuniones periódicas con los distintos entes de gobierno escolar constituyen, en esta pregunta, el único mecanismo de seguimiento que las seis instituciones del grupo, sin excepción, reportan haber implementado. Con una adopción del 83,3% (5 de 6 IE) se ubica la autoevaluación institucional anual; y con una adopción del 66,7% (4 de 6 IE), las elecciones estudiantiles, la elección de docentes para los consejos directivo y académico, el Plan de Mejoramiento Institucional y la actualización permanente del PEI. Este patrón confirma que, al igual que en la pregunta análoga de equipos de trabajo, el Grupo 3 presenta el nivel más alto y homogéneo de adopción de mecanismos reglamentarios observado hasta ahora, con las reuniones de gobierno escolar como práctica verdaderamente universal.",
      particularidades: "Santa Librada y Tierra de Promisión son las dos instituciones que más se apartan del catálogo estándar: Santa Librada no reporta ninguno de los procesos electorales del catálogo, apoyándose en cambio en la revisión de actas, acuerdos y espacios de diálogo como mecanismo \"Otro\"; y Tierra de Promisión complementa su selección con una descripción propia del ciclo completo de gestión de proyectos institucionales (planeación, programación, ejecución, seguimiento y evaluación) en lugar de los mecanismos normativos de autoevaluación y PMI seleccionados por la mayoría del grupo."
    }
  ],
  conclusiones: "El análisis conjunto de las tres sesiones de trabajo del Grupo 3 permite concluir que las seis instituciones educativas (Técnico Superior, Departamental Tierra de Promisión, Santa Librada, Ricardo Borrero Álvarez, Ángel María Paredes y Ceinar) comparten una trayectoria de avances reales en la contextualización curricular, el fortalecimiento de proyectos pedagógicos propios y una estructura de gobierno escolar particularmente consolidada y homogénea —la más alta adopción de equipos y mecanismos reglamentarios observada hasta ahora entre los grupos analizados—, junto con retos comunes en materia de infraestructura y dotación, y con la baja participación de las familias como el obstáculo más recurrente y explícito de todo el grupo, mencionado por cinco de las seis instituciones. Dos hallazgos particulares merecen destacarse: la denuncia de Ricardo Borrero Álvarez sobre la centralización de decisiones clave en el Equipo de Gestión \"sin comunicarse oportunamente a las bases\", que amerita atención específica; y el dato cuantitativo aportado por Ángel María Paredes, según el cual apenas el 27,1% de su comunidad consultada considera suficientes y oportunos los recursos para la educación inicial, frente a un 51% que los califica de insuficientes o parciales. En cuanto a la valoración general del Foro, de las seis instituciones del grupo, únicamente Tierra de Promisión registró formalmente su valoración de la jornada, con una calificación de 5,0 sobre 5,0 en los cuatro criterios evaluados (diálogo y reflexión, participación, ideas y propuestas, y satisfacción con el instrumento); las cinco instituciones restantes (Técnico Superior, Santa Librada, Ricardo Borrero Álvarez, Ángel María Paredes y Ceinar) aún no habían registrado su valoración ni comentarios o sugerencias a la fecha de consolidación de este informe. Se recomienda a la Secretaría de Educación gestionar el diligenciamiento pendiente de este instrumento con las cinco instituciones señaladas.",
  proyectoNombre: "[Nombre de quien consolida el Informe Consolidado]",
  proyectoCargo: "[Cargo]",
  fechaRealizacion: "7 de septiembre de 2026"
};

/*
 * Genera (o reescribe, si ya existe) el Informe de Síntesis Grupal del
 * Grupo 3, dejándolo en la raíz de la carpeta "Grupo G3" (mismo
 * criterio de idempotencia que el resto de documentos de grupo).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El enlace del documento queda en "Ver
 * registros de ejecución".
 */
function generarInformeSintesisGrupo3FEM(){
  const carpetas=crearEstructuraCarpetasGrupoFEM_("G3");
  const nombreDoc="Informe de Síntesis - Grupo G3 FEM 2026";
  const existentesIt=carpetas.grupoFolder.getFilesByName(nombreDoc);
  const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
  while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

  const archivoDoc=generarInformeSintesisGrupoFEM_(DATOS_SINTESIS_GRUPO_3_FEM_, archivoExistente?archivoExistente.getId():null);
  if(!archivoExistente){
    carpetas.grupoFolder.addFile(archivoDoc);
    try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
  }

  Logger.log("========================================");
  Logger.log("INFORME DE SÍNTESIS — GRUPO 3 — RESULTADO");
  Logger.log("Documento: "+archivoDoc.getUrl());
  Logger.log("========================================");

  return archivoDoc.getUrl();
}

/*
 * INFORME DE SÍNTESIS GRUPAL — GRUPO 4.
 *
 * Mismo criterio que los Grupos 1, 2 y 3: prosa redactada a mano leyendo
 * directamente el Doc editable ya generado de cada IE del grupo, y
 * tallas de las 4 preguntas mixtas contadas a mano sobre las opciones
 * realmente seleccionadas.
 *
 * PARTICULARIDAD DE ESTE GRUPO: de las 6 IE del catálogo (José Eustasio
 * Rivera, Atanasio Girardot, Misael Pastrana Borrero, Humberto Tafur
 * Charry, Enrique Olaya Herrera, Roberto Durán Alvira), HUMBERTO TAFUR
 * CHARRY NO HABÍA ENVIADO SU INFORME EJECUTIVO a la fecha de redacción
 * (7 de septiembre de 2026) — solo existen en Drive el PDF de asistencia
 * y la foto del evento, ningún "Informe Ejecutivo". Por eso el análisis
 * de las 11 preguntas y las conclusiones se basan en las 5 IE restantes
 * (denominador 5, no 6), y la portada lista las 6 IE del grupo pero las
 * conclusiones señalan expresamente la ausencia de esta institución.
 *
 * De esas 5 IE, solo José Eustasio Rivera había registrado su Valoración
 * del Foro (4.5/5) a la fecha de redacción.
 */
const DATOS_SINTESIS_GRUPO_4_FEM_ = {
  grupo: "G4",
  tituloInforme: "Informe de Síntesis Grupal — Elementos Comunes y Particularidades Institucionales",
  instituciones: [
    "JOSE EUSTASIO RIVERA",
    "ATANASIO GIRARDOT",
    "MISAEL PASTRANA BORRERO",
    "HUMBERTO TAFUR CHARRY",
    "ENRIQUE OLAYA HERRERA",
    "ROBERTO DURAN ALVIRA"
  ],
  responsableInforme: "[Nombre de quien consolida el informe de síntesis del Grupo 4 — SEM Neiva]",
  fechaPresentacion: "7 de septiembre de 2026",
  secciones: [
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta orientadora",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en el logro de los retos y propósitos planteados en el FEM2025?",
      tipo: "cualitativo",
      comun: "Cuatro de las cinco instituciones del Grupo 4 que remitieron su informe (José Eustasio Rivera, Atanasio Girardot, Misael Pastrana Borrero y Enrique Olaya Herrera) coinciden en describir su avance frente al FEM2025 como parcial y desigual, con logros reales pero todavía insuficientemente sistematizados: José Eustasio Rivera señala como reto pendiente \"consolidar un trabajo metodológico articulado y sistemático\"; Enrique Olaya Herrera reconoce \"debilidades en los procesos de sistematización\" de la información y las experiencias pedagógicas; y Misael Pastrana Borrero describe un avance \"paulatino y desigual entre instituciones, en la medida en que cada una lo determina según las condiciones y los recursos con que cuenta\". Dentro de este patrón compartido, la educación incluyente aparece como el eje de avance más consistente del grupo: José Eustasio Rivera destaca la formalización de los planes PIAR y el DUA, Atanasio Girardot reporta el aumento de población diversa y una actitud docente positiva hacia la inclusión, y Misael Pastrana Borrero resalta el incremento de estudiantes con PIAR sistematizados a tiempo. Como contraste notable dentro del propio grupo, Roberto Durán Alvira reporta el nivel de avance más alto y consolidado, estimando en cerca del 80% el cumplimiento de las metas del FEM2025 gracias a su macroproyecto institucional \"Cultivando saberes y conexiones rurales\".",
      particularidades: "Roberto Durán Alvira presenta la particularidad más marcada del grupo en esta pregunta al ser la única institución en cuantificar explícitamente su nivel de cumplimiento (80% de las metas), y la única en describir un macroproyecto institucional propio y nombrado que articula la totalidad de sus estrategias con la vocación agropecuaria del corregimiento de Vegalarga. Atanasio Girardot, por su parte, es la única institución en señalar de manera explícita problemáticas de consumo de sustancias psicoactivas y presencia de pandillas entre sus estudiantes como factor asociado al bajo rendimiento y la desmotivación."
    },
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta 2",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en la implementación de los nuevos grados del nivel de preescolar (jardín, prejardín)?",
      tipo: "cualitativo",
      comun: "Cuatro de las cinco instituciones del grupo (José Eustasio Rivera, Atanasio Girardot, Enrique Olaya Herrera y Roberto Durán Alvira) ya tienen en funcionamiento el grado jardín, mientras que Misael Pastrana Borrero es la única en reconocer que, pese a fortalecer el grado de transición mediante recursos para material didáctico y la reducción de su cobertura a 25 estudiantes por aula, no logró avances concretos en la creación de nuevos grados de jardín o prejardín por falta de infraestructura y de personal docente suficiente. Entre las cuatro instituciones que sí ofrecen jardín, existe una coincidencia notable en que el prejardín específicamente enfrenta una barrera de demanda estructural: José Eustasio Rivera, Atanasio Girardot y Enrique Olaya Herrera señalan, en términos casi idénticos, que las familias prefieren los hogares comunitarios o centros de cuidado del ICBF y del sector privado, que ofrecen alimentación y jornada completa que la oferta escolar no puede igualar, lo que explica la baja matrícula o la inactividad de este nivel. Finalmente, incluso donde el nivel ya opera, persisten limitaciones de infraestructura: baterías sanitarias pendientes en José Eustasio Rivera, obras físicas incompletas en Atanasio Girardot, un aula por adecuar en Enrique Olaya Herrera, y hacinamiento y ausencia de zonas de recreación en Roberto Durán Alvira.",
      particularidades: "Roberto Durán Alvira presenta la particularidad más contextual del grupo en esta pregunta, al señalar que los vehículos de transporte escolar disponibles \"no cumplen con las condiciones de seguridad para el transporte de los estudiantes de primera infancia\" en su zona rural, una barrera de movilidad específica no reportada por ninguna otra institución. Misael Pastrana Borrero, por su parte, es la única institución del grupo en admitir con total transparencia que, en la práctica, no logró implementar ningún grado nuevo de preescolar durante el periodo, limitándose a fortalecer el nivel ya existente de transición."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que los currículos actuales que se desarrollan en las instituciones educativas son pertinentes con sus realidades territoriales (sociales, culturales, productivas)? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Tres de las cinco instituciones del grupo (José Eustasio Rivera, Misael Pastrana Borrero y Roberto Durán Alvira) consideran que sus currículos sí son pertinentes con las realidades territoriales, mientras que Atanasio Girardot y Enrique Olaya Herrera matizan esa pertinencia como parcial, señalando retos más profundos de articulación. Pese a esta diferencia de énfasis, existe una coincidencia sustancial en dos estrategias: la articulación con el SENA como vía de pertinencia técnico-laboral, presente en José Eustasio Rivera (asistencia administrativa e inglés), Atanasio Girardot (Enfermería, Sistemas, Contabilidad, AIPI) y Roberto Durán Alvira (competencias agropecuarias); y el fortalecimiento de cátedras y proyectos transversales —Alegría de Leer y Cátedra de Paz en José Eustasio Rivera, las cátedras de paz, afrocolombianidad y socioemocionalidad en Misael Pastrana Borrero, y el macroproyecto \"Cultivando saberes y conexiones rurales\" en Roberto Durán Alvira— como mecanismo central de contextualización curricular.",
      particularidades: "Enrique Olaya Herrera se distingue del resto del grupo por abordar la pregunta desde una reflexión conceptual sobre el territorio —entendido explícitamente como \"un escenario donde confluyen la historia, la cultura, las prácticas sociales, los conocimientos ancestrales, las dinámicas productivas, las necesidades y las expectativas de las comunidades\"— sin describir avances curriculares concretos ya materializados, a diferencia de las demás instituciones. Atanasio Girardot, por su parte, es la única institución en señalar explícitamente falencias en los planes de evacuación y gestión del riesgo ante emergencias reales como un vacío de pertinencia curricular."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se han implementado para lograr currículos más pertinentes al territorio?",
      tipo: "cualitativo",
      comun: "Las cinco instituciones del grupo coinciden en reportar el fortalecimiento de centros de interés y proyectos pedagógicos transversales como la acción más ampliamente compartida para lograr currículos más pertinentes: el PRAE y la estrategia Alegría de Leer en José Eustasio Rivera, los Centros de Interés y la Peña Cultural en Atanasio Girardot, los proyectos de aula articulados con las cátedras transversales en Misael Pastrana Borrero, los centros de interés propios y nombrados —\"La magia de la palabra hablada\", \"Ser, Soñar y Lograr\"— en Enrique Olaya Herrera, y los proyectos pedagógicos productivos en huertas y emprendimientos agrícolas de Roberto Durán Alvira. Como segunda línea compartida por tres de las cinco instituciones, la articulación con el SENA aparece como estrategia concreta de pertinencia técnico-laboral en José Eustasio Rivera, Enrique Olaya Herrera y Roberto Durán Alvira, mientras que la revisión y actualización periódica del PEI es reportada explícitamente por José Eustasio Rivera, Misael Pastrana Borrero y Enrique Olaya Herrera como mecanismo formal que acompaña estas acciones.",
      particularidades: "Atanasio Girardot presenta la particularidad más diversificada del grupo en esta pregunta, al vincular sus acciones curriculares con un conjunto amplio de eventos culturales y deportivos propios —la Peña Cultural, el Festival de la Fruta y la Verdura, la celebración de la Huilensidad y los Juegos SUPÉRATE— y con una alianza explícita con la Policía Antinarcóticos para la prevención, un componente de seguridad no mencionado por ninguna otra institución del grupo. Roberto Durán Alvira, por su parte, es la única institución en fundamentar sus acciones curriculares en un diagnóstico ambiental de campo mediante salidas pedagógicas a fincas de la región, articulando explícitamente café, ganadería y negocios verdes como ejes productivos del territorio."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 3",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar estas acciones?",
      tipo: "mixto",
      totalIE: 5,
      tally: [
        {opcion: "Comités de área", count: 5},
        {opcion: "Consejo Académico", count: 4},
        {opcion: "Comisión de Evaluación y Promoción", count: 4},
        {opcion: "Equipo de Gestión Académica", count: 4},
        {opcion: "Equipo de Educación Inclusiva / NEE", count: 3}
      ],
      comun: "Los comités de área son el único equipo de trabajo que las cinco instituciones del grupo, sin excepción, reportan haber conformado para esta pregunta. Con una adopción del 80% (4 de 5 IE) se ubican el Consejo Académico, la Comisión de Evaluación y Promoción y el Equipo de Gestión Académica; y con una adopción del 60% (3 de 5 IE), el Equipo de Educación Inclusiva o NEE. Este patrón muestra una base compartida relativamente sólida en torno a las instancias curriculares y de evaluación más reglamentarias, sobre la cual cada institución construye una estructura adicional propia y de complejidad muy variable.",
      particularidades: "José Eustasio Rivera y Roberto Durán Alvira son, con diferencia, las instituciones con la estructura de equipos más extensa y diferenciada del grupo —trece equipos reglamentarios cada una, incluyendo divisiones por nivel educativo (preescolar, primaria, secundaria y media) y equipos especializados como Sistematización de la Información y Seguimiento a Egresados—, mientras que Atanasio Girardot presenta, en el extremo opuesto, la selección más reducida de todo el grupo, con apenas tres equipos reportados y sin ninguna instancia de gobierno escolar reglamentario (Consejo Directivo, Consejo Estudiantil) para esta pregunta."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 4",
      enunciado: "¿Cómo se están articulando estos equipos de trabajo para lograr currículos más pertinentes territorialmente?",
      tipo: "cualitativo",
      comun: "Tres de las cinco instituciones del grupo (José Eustasio Rivera, Misael Pastrana Borrero y Enrique Olaya Herrera) describen mecanismos de articulación curricular ya en funcionamiento, apoyados en reuniones periódicas de comités de área y equipos de gestión que canalizan las propuestas hacia su aprobación institucional. Las otras dos instituciones, en cambio, reconocen abiertamente que esta articulación es todavía insuficiente: Roberto Durán Alvira admite la \"necesidad de superar el funcionamiento aislado mediante la concertación de un cronograma unificado\", y Atanasio Girardot llega más lejos al señalar que \"durante el año 2026 no se ha convocado el Consejo Académico ni otras instancias del gobierno escolar\", dejando la reflexión curricular en manos exclusivas de comités de área y equipos específicos. Esta polarización —entre instituciones con mecanismos consolidados y otras con vacíos institucionales explícitos en la convocatoria de sus propias instancias reglamentarias— es más marcada en el Grupo 4 que en los grupos analizados previamente.",
      particularidades: "Atanasio Girardot constituye el hallazgo más crítico de esta pregunta en todo el grupo, al admitir la ausencia de convocatoria del Consejo Académico durante todo el año escolar, un vacío de gobierno curricular no reportado por ninguna otra institución en ningún grupo analizado hasta ahora. Enrique Olaya Herrera, por su parte, es la única institución en vincular explícitamente su articulación curricular con proyectos culturales nombrados (Batuta, la Revista Huellas) como parte de su estructura de equipos."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 5",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para que dichas acciones se cumplan?",
      tipo: "mixto",
      totalIE: 5,
      tally: [
        {opcion: "Comités de área o núcleos de formación", count: 5},
        {opcion: "Tasa de deserción/reprobación por asignatura", count: 5},
        {opcion: "Autoevaluación institucional anual", count: 5},
        {opcion: "Reportes al MEN / Planes de Mejoramiento Institucional", count: 5},
        {opcion: "Análisis de resultados de Pruebas Saber e ICFES", count: 4},
        {opcion: "Consejo Académico", count: 4},
        {opcion: "Articulación con el PEI", count: 4},
        {opcion: "Actas de reunión de área", count: 4},
        {opcion: "Plataformas de gestión académica (JIGRA)", count: 4},
        {opcion: "Visitas de las Secretarías de Educación", count: 4}
      ],
      comun: "Los mecanismos de seguimiento curricular alcanzan, en el Grupo 4, el nivel de consenso más alto observado hasta ahora entre los grupos analizados: los comités de área, la tasa de deserción o reprobación por asignatura, la autoevaluación institucional anual y los reportes al MEN mediante el Plan de Mejoramiento Institucional son adoptados por las cinco instituciones, sin excepción. Con una adopción del 80% (4 de 5 IE) se ubican además el análisis de resultados de Pruebas Saber e ICFES, el Consejo Académico, la articulación con el PEI, las actas de reunión de área, las plataformas de gestión académica JIGRA y las visitas de las Secretarías de Educación. Este patrón evidencia que, más allá de las diferencias en la profundidad curricular reportadas en preguntas anteriores, el grupo comparte un aparato de seguimiento y rendición de cuentas particularmente robusto y estandarizado.",
      particularidades: "Roberto Durán Alvira presenta, con dieciséis mecanismos reportados, el catálogo de seguimiento más extenso y diversificado de todo el grupo, siendo la única institución en incluir la evaluación por matriz DOFA como herramienta formal de análisis institucional. Atanasio Girardot, en el extremo opuesto, es la institución con el catálogo más reducido (seis mecanismos), sin reportar Consejo Académico, articulación con el PEI, actas de reunión de área ni plataformas JIGRA, una discrepancia que resulta coherente con su selección igualmente reducida de equipos de trabajo en la pregunta anterior."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que la toma de decisiones en las instituciones educativas actualmente es participativa y democrática? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Las cinco instituciones del grupo coinciden en reconocer la existencia de estructuras formales de gobierno escolar amparadas en la Ley 115, y las cinco matizan esa existencia formal señalando límites concretos a la participación real. Este patrón es particularmente marcado en tres instituciones —Atanasio Girardot, Misael Pastrana Borrero y Enrique Olaya Herrera—, que describen la participación como predominantemente pasiva, autocrática o meramente administrativa: Atanasio Girardot reporta que el Consejo Directivo y el Consejo Académico solo se han convocado una vez en todo el año; Misael Pastrana Borrero señala que los órganos del gobierno escolar \"suelen cumplir un papel más pasivo que activo\"; y Enrique Olaya Herrera observa que los espacios formales \"a menudo se limitan a cumplir requisitos administrativos\", con estudiantes que no participan en la construcción curricular \"por falta de interés\" y una ausencia total de padres de familia el día del propio Foro. La baja participación de las familias es, además, un obstáculo mencionado explícitamente por cuatro de las cinco instituciones (José Eustasio Rivera, Misael Pastrana Borrero, Enrique Olaya Herrera y Roberto Durán Alvira).",
      particularidades: "Enrique Olaya Herrera presenta la constatación más directa y singular del grupo en esta pregunta, al registrar en su propio informe que \"el día de hoy no contamos con la participación de los padres de familia en el foro\", una ausencia total documentada explícitamente que ninguna otra institución del grupo reporta en estos términos. Atanasio Girardot, por su parte, es la única institución en señalar que los estudiantes de educación inicial no participan en absoluto en los procesos de elección o democracia escolar, una precisión etaria no formulada por las demás instituciones."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se están implementando para canalizar y fortalecer la participación de la comunidad educativa?",
      tipo: "cualitativo",
      comun: "Cuatro de las cinco instituciones del grupo (José Eustasio Rivera, Atanasio Girardot, Misael Pastrana Borrero y Roberto Durán Alvira) coinciden en reportar el fortalecimiento de los órganos del Gobierno Escolar —consejos, personería, contraloría— como acción central para canalizar la participación, mientras que Enrique Olaya Herrera concentra sus acciones, en cambio, en el proceso participativo de actualización del PEI con asesoría externa de la Fundación Terpel. Una segunda línea, igualmente compartida por cuatro de las cinco instituciones (Atanasio Girardot, Misael Pastrana Borrero, Enrique Olaya Herrera y Roberto Durán Alvira), es el fortalecimiento de canales de comunicación —páginas web, redes sociales, WhatsApp y medios digitales— como estrategia explícita para superar barreras de participación, ya sea de desinformación (Atanasio Girardot) o geográficas en zonas rurales apartadas (Roberto Durán Alvira).",
      particularidades: "José Eustasio Rivera es la única institución del grupo en reportar una veeduría comunitaria formalizada y nombrada —el comité de acompañamiento de padres de familia, liderado por el Comité de Ciencias Sociales, que hace seguimiento directo a la gestión, licitación y construcción de la infraestructura de la sede Ciudad Jardín—, una figura de control social con un objeto de vigilancia concreto sin equivalente en las demás instituciones. Atanasio Girardot, por su parte, es la única institución en vincular su fortalecimiento de la participación con alianzas externas nombradas de carácter social y universitario (Jóvenes Pacíficos de Ecopetrol y capacitaciones en resolución de conflictos con la Universidad Surcolombiana)."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Equipos de trabajo",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar las estrategias y mecanismos de participación escolar?",
      tipo: "mixto",
      totalIE: 5,
      tally: [
        {opcion: "Gobierno Escolar", count: 5},
        {opcion: "Consejo Estudiantil", count: 5},
        {opcion: "Comisión de Evaluación y Promoción", count: 5},
        {opcion: "Comité de Calidad / Equipo de Gestión Institucional", count: 5},
        {opcion: "Consejo Directivo", count: 4},
        {opcion: "Consejo Académico", count: 4},
        {opcion: "Personero(a) Estudiantil", count: 4},
        {opcion: "Contralor(a) Estudiantil", count: 4},
        {opcion: "Consejo de Padres de Familia", count: 4},
        {opcion: "Comité Escolar de Convivencia", count: 4}
      ],
      comun: "El Gobierno Escolar en su conjunto, el Consejo Estudiantil, la Comisión de Evaluación y Promoción y el Comité de Calidad o Equipo de Gestión Institucional son los cuatro equipos que las cinco instituciones del grupo, sin excepción, reportan haber conformado para las estrategias de participación escolar. Con una adopción del 80% (4 de 5 IE) se ubican el Consejo Directivo, el Consejo Académico, el/la Personero(a) y Contralor(a) Estudiantil, el Consejo de Padres de Familia y el Comité Escolar de Convivencia. Este patrón revela una estructura de gobierno escolar altamente homogénea entre cuatro de las cinco instituciones, con Atanasio Girardot como la única excepción sistemática en casi todo el catálogo reglamentario.",
      particularidades: "Atanasio Girardot constituye, con notable diferencia, la particularidad más marcada del grupo en esta pregunta: es la única institución que no reporta Consejo Directivo, Consejo Académico, Personero(a) ni Contralor(a) Estudiantil, Consejo de Padres de Familia ni Comité Escolar de Convivencia, seleccionando apenas cuatro equipos del catálogo y complementando con una opción \"Otro\" (comités de área) sin mayor desarrollo — una estructura de participación considerablemente más débil que la de las demás instituciones del grupo. Roberto Durán Alvira, por su parte, es la única institución en incluir el Comité de Alimentación Escolar (CAE) como equipo adicional de participación, vinculando la veeduría del PAE con la estructura de gobierno escolar."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Mecanismos de seguimiento",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para garantizar las acciones encaminadas a promover gobiernos educativos democráticos?",
      tipo: "mixto",
      totalIE: 5,
      tally: [
        {opcion: "Autoevaluación institucional anual (Guía 34 – MEN)", count: 5},
        {opcion: "Plan de Mejoramiento Institucional (PMI)", count: 5},
        {opcion: "Reuniones periódicas con entes de gobierno escolar", count: 4},
        {opcion: "Elecciones estudiantiles", count: 4},
        {opcion: "Elección de docentes para consejo directivo", count: 4},
        {opcion: "Elección de docentes para consejo académico", count: 4},
        {opcion: "Seguimiento periódico al Plan de Mejoramiento", count: 4},
        {opcion: "Actualización y ajuste permanente del PEI", count: 4},
        {opcion: "Informes de gestión del Consejo Directivo", count: 4},
        {opcion: "Verificación del funcionamiento del Gobierno Escolar por la Secretaría de Educación", count: 4}
      ],
      comun: "La autoevaluación institucional anual y el Plan de Mejoramiento Institucional son los dos mecanismos de seguimiento que las cinco instituciones del grupo, sin excepción, reportan haber implementado. Con una adopción del 80% (4 de 5 IE) se ubica prácticamente la totalidad del resto del catálogo reglamentario: las reuniones periódicas con entes de gobierno escolar, las elecciones estudiantiles, la elección de docentes para los consejos directivo y académico, el seguimiento periódico al PMI, la actualización del PEI, los informes de gestión del Consejo Directivo y la verificación externa por parte de la Secretaría de Educación. Este patrón confirma, en paralelo con la pregunta análoga de equipos de trabajo, que Atanasio Girardot es sistemáticamente la única institución que se aparta del comportamiento homogéneo del resto del grupo.",
      particularidades: "Atanasio Girardot constituye, de nuevo, el hallazgo más crítico de esta pregunta en todo el grupo: además de seleccionar solo dos mecanismos del catálogo estándar, admite explícitamente en su respuesta \"Otro\" que \"no existen mecanismos formales, instrumentos consolidados o un seguimiento riguroso tras la etapa inicial de votación y conformación del gobierno escolar\", una confesión de vacío institucional en el seguimiento democrático que resulta coherente con su también escasa convocatoria del Consejo Académico reportada en la Sesión 2. Roberto Durán Alvira, por su parte, es la única institución en complementar el catálogo estándar con mecanismos propios de rendición pública de cuentas y de encuesta y buzón de sugerencias."
    }
  ],
  conclusiones: "El análisis conjunto de las tres sesiones de trabajo del Grupo 4 permite concluir que, entre las cinco instituciones que remitieron su Informe Ejecutivo (José Eustasio Rivera, Atanasio Girardot, Misael Pastrana Borrero, Enrique Olaya Herrera y Roberto Durán Alvira), existe una trayectoria compartida de avances parciales y desiguales hacia los propósitos del FEM2025, con la educación incluyente y el fortalecimiento de proyectos pedagógicos transversales como los ejes de progreso más consistentes, y con la baja participación real de las familias y los estudiantes —más allá de la existencia formal de estructuras de gobierno escolar— como el obstáculo más recurrente del grupo. Dos hallazgos particulares ameritan atención específica de la Secretaría de Educación Municipal: la IE Atanasio Girardot, que reconoció explícitamente no haber convocado su Consejo Académico durante todo el año escolar 2026 y admitió no contar con mecanismos formales de seguimiento a la democracia escolar más allá de la votación inicial, presentando de manera sistemática la estructura de participación más débil de todo el grupo; y la IE Roberto Durán Alvira, cuyo macroproyecto institucional \"Cultivando saberes y conexiones rurales\" reporta el nivel de avance más alto y consolidado del grupo frente a los retos del FEM2025, con cerca del 80% de cumplimiento estimado. En cuanto a la valoración general del Foro, de las cinco instituciones con informe remitido, únicamente José Eustasio Rivera registró formalmente su valoración de la jornada, con un promedio de 4,5 sobre 5,0 (diálogo y reflexión 5,0, participación 4,0, ideas y propuestas 5,0, satisfacción con el instrumento 4,0), acompañado de sugerencias sobre el fortalecimiento de la planeación logística y los recursos para el FEM 2027; las cuatro instituciones restantes (Atanasio Girardot, Misael Pastrana Borrero, Enrique Olaya Herrera y Roberto Durán Alvira) aún no han registrado su valoración. Finalmente, la IE Humberto Tafur Charry, sexta institución del grupo, únicamente registró la asistencia y la evidencia fotográfica de la jornada, sin haber generado ni enviado su Informe Ejecutivo a la fecha de consolidación de este documento; se recomienda a la Secretaría de Educación gestionar de manera prioritaria el diligenciamiento pendiente de este informe con dicha institución.",
  proyectoNombre: "[Nombre de quien consolida el Informe Consolidado]",
  proyectoCargo: "[Cargo]",
  fechaRealizacion: "7 de septiembre de 2026"
};

/*
 * Genera (o reescribe, si ya existe) el Informe de Síntesis Grupal del
 * Grupo 4, dejándolo en la raíz de la carpeta "Grupo G4" (mismo
 * criterio de idempotencia que el resto de documentos de grupo).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El enlace del documento queda en "Ver
 * registros de ejecución".
 */
function generarInformeSintesisGrupo4FEM(){
  const carpetas=crearEstructuraCarpetasGrupoFEM_("G4");
  const nombreDoc="Informe de Síntesis - Grupo G4 FEM 2026";
  const existentesIt=carpetas.grupoFolder.getFilesByName(nombreDoc);
  const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
  while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

  const archivoDoc=generarInformeSintesisGrupoFEM_(DATOS_SINTESIS_GRUPO_4_FEM_, archivoExistente?archivoExistente.getId():null);
  if(!archivoExistente){
    carpetas.grupoFolder.addFile(archivoDoc);
    try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
  }

  Logger.log("========================================");
  Logger.log("INFORME DE SÍNTESIS — GRUPO 4 — RESULTADO");
  Logger.log("Documento: "+archivoDoc.getUrl());
  Logger.log("========================================");

  return archivoDoc.getUrl();
}

/*
 * INFORME DE SÍNTESIS GRUPAL — GRUPO 5.
 *
 * Mismo criterio que los grupos anteriores: prosa redactada a mano
 * leyendo directamente el Doc editable ya generado de cada IE del
 * grupo, y tallas de las 4 preguntas mixtas contadas a mano sobre las
 * opciones realmente seleccionadas.
 *
 * Las 6 IE del catálogo (Santa Teresa, Escuela Normal Superior,
 * Instituto Técnico IPC Andrés Rosa, Juan de Cabrera, Jairo Morera
 * Lizcano, San Antonio de Anaconia) SÍ habían remitido su Informe
 * Ejecutivo a la fecha de redacción (7 de septiembre de 2026), por lo
 * que el denominador de las 4 preguntas mixtas es 6.
 *
 * De esas 6 IE, la Escuela Normal Superior (5.0/5) y el Instituto
 * Técnico IPC Andrés Rosa (4.8/5) habían registrado su Valoración del
 * Foro a la fecha de redacción.
 */
const DATOS_SINTESIS_GRUPO_5_FEM_ = {
  grupo: "G5",
  tituloInforme: "Informe de Síntesis Grupal — Elementos Comunes y Particularidades Institucionales",
  instituciones: [
    "SANTA TERESA",
    "ESCUELA NORMAL SUPERIOR",
    "INSTITUTO TECNICO IPC ANDRES ROSA",
    "JUAN DE CABRERA",
    "JAIRO MORERA LIZCANO",
    "SAN ANTONIO DE ANACONIA"
  ],
  responsableInforme: "[Nombre de quien consolida el informe de síntesis del Grupo 5 — SEM Neiva]",
  fechaPresentacion: "7 de septiembre de 2026",
  secciones: [
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta orientadora",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en el logro de los retos y propósitos planteados en el FEM2025?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del Grupo 5 (Santa Teresa, Escuela Normal Superior, Instituto Técnico IPC Andrés Rosa, Juan de Cabrera, Jairo Morera Lizcano y San Antonio de Anaconia) coinciden en reportar avances sostenidos en la contextualización curricular mediante proyectos pedagógicos propios y transversales: el macroproyecto \"Amémonos y Creceremos\" y las estrategias Líder en Mí y Mentes PRO en Santa Teresa, los Proyectos Pedagógicos de Aula (PPA) en la Escuela Normal Superior, las áreas integradas y los centros de interés en el Instituto Técnico IPC Andrés Rosa, la reestructuración curricular hacia el modelo constructivista en Juan de Cabrera, el PRAE y la feria de emprendimiento en Jairo Morera Lizcano, y la implementación progresiva de la educación inicial en San Antonio de Anaconia. De manera igualmente compartida, cuatro de las seis instituciones (Escuela Normal Superior, Instituto Técnico IPC Andrés Rosa, Juan de Cabrera y Jairo Morera Lizcano) identifican limitaciones de infraestructura, conectividad y personal de apoyo especializado como el obstáculo estructural más persistente para consolidar estos avances, particularmente agudo en las sedes rurales de la Escuela Normal Superior y en la sede Guillermo Liévano de Jairo Morera Lizcano.",
      particularidades: "Jairo Morera Lizcano presenta la particularidad más crítica del grupo en esta pregunta, al describir un deterioro tecnológico explícito —laboratorios inexistentes, tablets obsoletas y conectividad nula en una de sus sedes— que frena directamente la innovación pedagógica pese al esfuerzo docente, y al reconocer que el Programa de Alimentación Escolar opera de forma parcial por carecer de ración para la jornada única y de transporte escolar. San Antonio de Anaconia, por su parte, es la única institución en sustentar sus avances con testimonios directos y nombrados de la comunidad educativa —como el de la madre comunitaria Neyla Olaya—, evidenciando la transformación de la atención a la primera infancia desde la perspectiva de sus propias beneficiarias."
    },
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta 2",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en la implementación de los nuevos grados del nivel de preescolar (jardín, prejardín)?",
      tipo: "cualitativo",
      comun: "Cinco de las seis instituciones del grupo (Santa Teresa, Escuela Normal Superior, Instituto Técnico IPC Andrés Rosa, Jairo Morera Lizcano y San Antonio de Anaconia) ya tienen en funcionamiento el grado jardín, con distintos niveles de consolidación del prejardín, mientras que Juan de Cabrera es la única institución del grupo que reconoce no haber logrado la implementación efectiva de ninguno de los dos grados pese a haber realizado la oferta correspondiente, atribuyendo esta situación a que las familias prefieren mantener a los niños en hogares infantiles o programas de bienestar familiar, lo que impide alcanzar el cupo mínimo requerido. Entre las instituciones que sí operan el nivel, existe una coincidencia notable en que la infraestructura, la dotación y el personal formado siguen siendo insuficientes: Santa Teresa señala la necesidad de fortalecer infraestructura, dotación y formación docente; la Escuela Normal Superior reconoce condiciones diferenciadas entre sedes, con aulas multigrado y personal inestable en las sedes rurales; y Jairo Morera Lizcano denuncia el incumplimiento del gobierno en el envío de materiales pedagógicos y riesgos estructurales críticos, como accesos exclusivamente por escaleras que comprometen la seguridad de los niños más pequeños.",
      particularidades: "El Instituto Técnico IPC Andrés Rosa aporta el hallazgo demográfico más singular del grupo en esta pregunta, al identificar una disminución en la tasa de natalidad del sector como factor que ha reducido la matrícula y la conformación de grupos de preescolar, una dinámica poblacional —y no de infraestructura o de política institucional— que ninguna otra institución del grupo reporta como causa de estancamiento. Jairo Morera Lizcano, por su parte, es la única institución en documentar un riesgo de seguridad física concreto y crítico —accesos limitados exclusivamente por escaleras en el aula de primera infancia— derivado de la falta de adecuación de espacios."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que los currículos actuales que se desarrollan en las instituciones educativas son pertinentes con sus realidades territoriales (sociales, culturales, productivas)? ¿Por qué?",
      tipo: "cualitativo",
      comun: "El Grupo 5 presenta la valoración más dividida observada hasta ahora sobre la pertinencia curricular: Santa Teresa, el Instituto Técnico IPC Andrés Rosa y Juan de Cabrera afirman explícitamente que sus currículos sí son pertinentes con las realidades territoriales, la Escuela Normal Superior y San Antonio de Anaconia describen avances importantes que aún requieren consolidación, y Jairo Morera Lizcano llega a la conclusión opuesta, afirmando sin matices que \"el currículo no es pertinente con la realidad territorial\". Pese a esta divergencia de conclusión, las seis instituciones coinciden en identificar la misma tensión de fondo: la presión de los referentes curriculares y las pruebas estandarizadas nacionales frente a las particularidades sociales, culturales y productivas locales. San Antonio de Anaconia advierte sobre \"una concepción curricular homogénea, en la que se espera que todos los estudiantes desarrollen contenidos y aprendizajes similares, independientemente de su contexto\"; Jairo Morera Lizcano denuncia que \"el sistema tiende a medir a los estudiantes como cifras e indicadores, ignorando sus contextos\"; y el Instituto Técnico IPC Andrés Rosa reconoce la necesidad de \"superar progresivamente una planeación rígida\".",
      particularidades: "Jairo Morera Lizcano constituye el hallazgo más contundente del grupo en esta pregunta al ser la única institución, entre todos los grupos analizados hasta ahora, en concluir de manera explícita y sin matices que su currículo no es pertinente con la realidad territorial, atribuyendo la sostenibilidad de la enseñanza contextualizada casi enteramente a la vocación individual del cuerpo docente más que a una política institucional o estatal. La Escuela Normal Superior, por su parte, es la única institución del grupo en fundamentar su pertinencia curricular en la atención simultánea a una diversidad poblacional excepcionalmente amplia —sedes rurales, contextos urbanos vulnerables, comunidad sorda y oyente, y el Programa de Formación Complementaria— dentro de una misma institución."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se han implementado para lograr currículos más pertinentes al territorio?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del grupo coinciden en reportar el fortalecimiento de proyectos pedagógicos contextualizados como la acción central para lograr currículos más pertinentes, bajo denominaciones propias en cada caso: los macroproyectos y estrategias de Santa Teresa (Amémonos y Creceremos, Líder en Mí, Guardianes de la Naturaleza), los Proyectos Pedagógicos de Aula de la Escuela Normal Superior, los centros de interés del Instituto Técnico IPC Andrés Rosa, la transversalización de proyectos (PRAE, educación vial, emprendimiento) en Juan de Cabrera, el PRAE y los proyectos culturales de Jairo Morera Lizcano, y la formación integral mediante centros de interés de San Antonio de Anaconia. Como segunda línea compartida por cuatro de las seis instituciones, el diagnóstico o la caracterización sistemática del contexto y de la población estudiantil —explícitos en la Escuela Normal Superior, el Instituto Técnico IPC Andrés Rosa, Juan de Cabrera y Jairo Morera Lizcano— aparece como el punto de partida metodológico que orienta los ajustes curriculares posteriores.",
      particularidades: "La Escuela Normal Superior es la única institución del grupo en reportar un espacio de socialización interinstitucional propio y nombrado —el \"Foro Escuela de Vida\"— que le permite comparar y compartir experiencias curriculares entre sus distintas sedes, una estrategia de intercambio horizontal sin equivalente en las demás instituciones. San Antonio de Anaconia, por su parte, es la única institución en enfatizar de manera explícita la interdisciplinariedad como principio articulador transversal de todas sus acciones curriculares."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 3",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar estas acciones?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Consejo Académico", count: 6},
        {opcion: "Comités de área", count: 6},
        {opcion: "Consejo Directivo", count: 5},
        {opcion: "Equipo de Preescolar", count: 5},
        {opcion: "Consejo Estudiantil", count: 4},
        {opcion: "Personero(a) Estudiantil", count: 4}
      ],
      comun: "El Consejo Académico y los comités de área son los dos equipos de trabajo que las seis instituciones del grupo, sin excepción, reportan haber conformado para esta pregunta. Con una adopción del 83,3% (5 de 6 IE) se ubican el Consejo Directivo y el Equipo de Preescolar; y con una adopción del 66,7% (4 de 6 IE), el Consejo Estudiantil y el/la Personero(a) Estudiantil. Este patrón muestra una base reglamentaria sólida y ampliamente compartida, sobre la cual algunas instituciones del grupo —notablemente el Instituto Técnico IPC Andrés Rosa— construyen estructuras adicionales considerablemente más extensas y especializadas.",
      particularidades: "El Instituto Técnico IPC Andrés Rosa presenta, con amplio margen, la estructura de equipos más extensa y diversificada de todo el grupo, reportando cerca de veinte equipos distintos —incluyendo una Mesa de Diálogo Currículo-Comunidad, un Comité Curricular Institucional, un Equipo de Sistematización de la Información y un Equipo del Proyecto Ambientes de Aprendizaje Colaborativos con la Universidad Surcolombiana—, una complejidad organizativa muy superior a la de cualquier otra institución analizada hasta ahora en todos los grupos. Jairo Morera Lizcano, en el extremo opuesto, es la institución con la selección más reducida del grupo, con apenas dos equipos del catálogo estándar (Consejo Académico y comités de área) complementados por referencias a la articulación con el SENA y a un comité de seguimiento a la inclusión."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 4",
      enunciado: "¿Cómo se están articulando estos equipos de trabajo para lograr currículos más pertinentes territorialmente?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del grupo coinciden en que la articulación entre equipos de trabajo se canaliza mediante reuniones periódicas, comités de área y jornadas de planeación institucional, en las que se comparten diagnósticos y se construyen propuestas curriculares de manera colaborativa. Tres de las seis instituciones —Santa Teresa, la Escuela Normal Superior y Jairo Morera Lizcano— matizan esta descripción reconociendo limitaciones concretas: Santa Teresa admite que \"algunas articulaciones son ocasionales\"; la Escuela Normal Superior señala que \"persisten dificultades de tiempo, espacios y coordinación\"; y Jairo Morera Lizcano advierte que \"faltan espacios, tiempos e itinerarios institucionales permanentes para la comunicación, socialización e integración efectiva entre estamentos\". El Instituto Técnico IPC Andrés Rosa, Juan de Cabrera y San Antonio de Anaconia, en cambio, describen sus mecanismos de articulación en términos más consolidados, sin formular una autocrítica equivalente.",
      particularidades: "Santa Teresa es la única institución del grupo en vincular explícitamente la articulación curricular con la participación estudiantil en espacios externos de liderazgo y voluntariado como aporte de \"una mirada territorial y comunitaria\" al proceso, una fuente de insumos curriculares desde fuera del cuerpo docente no reportada por las demás instituciones. Jairo Morera Lizcano, por su parte, es la única institución en nombrar explícitamente sus proyectos transversales de articulación (PRAE, Eduderechos) como vehículos concretos alineados con el Plan de Mejoramiento Institucional."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 5",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para que dichas acciones se cumplan?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Análisis de resultados de Pruebas Saber e ICFES", count: 6},
        {opcion: "Autoevaluación institucional anual", count: 6},
        {opcion: "Actas de reunión de área", count: 6},
        {opcion: "Consejo Académico", count: 5},
        {opcion: "Comités de área o núcleos de formación", count: 4},
        {opcion: "Reportes al MEN / Planes de Mejoramiento Institucional", count: 4},
        {opcion: "Comparación histórica de resultados internos y externos", count: 4},
        {opcion: "Plataformas de gestión académica (JIGRA)", count: 4}
      ],
      comun: "El análisis de resultados de Pruebas Saber e ICFES, la autoevaluación institucional anual y las actas de reunión de área son los tres mecanismos de seguimiento que las seis instituciones del grupo, sin excepción, reportan haber implementado. Con una adopción del 83,3% (5 de 6 IE) se ubica el Consejo Académico; y con una adopción del 66,7% (4 de 6 IE), los comités de área, los reportes al MEN mediante el PMI, la comparación histórica de resultados y las plataformas de gestión académica JIGRA. Este patrón confirma a las pruebas estandarizadas y a la autoevaluación anual como el núcleo verdaderamente universal del seguimiento curricular en el Grupo 5, sobre el cual cada institución agrega herramientas propias de complejidad muy desigual.",
      particularidades: "El Instituto Técnico IPC Andrés Rosa presenta, con diecinueve mecanismos reportados, el catálogo de seguimiento más extenso de todos los grupos analizados hasta ahora, siendo la única institución en incluir de manera simultánea una auditoría curricular inicial frente a los DBA, una evaluación por matriz DOFA y mesas de calidad educativa territoriales. Jairo Morera Lizcano, en el extremo opuesto, presenta el catálogo más reducido del grupo con apenas cuatro mecanismos, sin reportar Consejo Académico ni comités de área pese a haberlos incluido como equipos de trabajo en la pregunta anterior, una inconsistencia entre la estructura declarada y su seguimiento efectivo que no se observa en las demás instituciones."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que la toma de decisiones en las instituciones educativas actualmente es participativa y democrática? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del grupo coinciden en reconocer la existencia de una estructura formal de gobierno escolar democrático —Consejo Directivo, Consejo Académico, Consejo Estudiantil, Consejo de Padres— y las seis matizan esa existencia formal señalando límites a su incidencia real. Cuatro de las seis instituciones (Santa Teresa, Instituto Técnico IPC Andrés Rosa, Juan de Cabrera y Jairo Morera Lizcano) coinciden, además, en identificar una percepción de centralización de las decisiones en los equipos directivos o de gestión: Santa Teresa reconoce que las decisiones administrativas \"pueden concentrarse en los equipos directivos\"; Juan de Cabrera señala que \"algunas decisiones se toman de manera impositiva por parte de directivos docentes o del equipo de gestión\"; y Jairo Morera Lizcano concluye que \"predomina la autoridad directiva o de un grupo reducido, relegando la incidencia real del resto de la comunidad\". San Antonio de Anaconia, por su parte, plantea el reto en términos ligeramente distintos, centrando su preocupación no en quién decide sino en que las decisiones democráticas se vuelvan visibles y tangibles para la comunidad, más allá del papel y las reuniones a puerta cerrada.",
      particularidades: "La Escuela Normal Superior introduce la reflexión conceptual más elaborada del grupo sobre la naturaleza de la democracia escolar, entendiéndola explícitamente no solo como \"votación o delegación\" sino como \"posibilidad de deliberar, expresar perspectivas, construir acuerdos y participar en asuntos institucionales\", y es la única institución en señalar de manera específica el débil acompañamiento institucional a las figuras de personero y contralor estudiantil. San Antonio de Anaconia, por su parte, es la única institución del grupo en formular el reto de la participación democrática en términos de visibilidad pública de sus resultados, más que de ampliación de los espacios formales existentes."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se están implementando para canalizar y fortalecer la participación de la comunidad educativa?",
      tipo: "cualitativo",
      comun: "Cinco de las seis instituciones del grupo (Santa Teresa, Escuela Normal Superior, Instituto Técnico IPC Andrés Rosa, Juan de Cabrera y Jairo Morera Lizcano) coinciden en reportar el fortalecimiento del Gobierno Escolar y de la Escuela de Padres como acción central para canalizar la participación de la comunidad educativa, mientras que San Antonio de Anaconia concentra sus acciones en herramientas de transparencia y co-construcción más innovadoras. Dos instituciones del grupo —Santa Teresa y Jairo Morera Lizcano— se distinguen además por vincular su fortalecimiento de la participación con alianzas interinstitucionales externas de carácter gubernamental y social: Santa Teresa articula a sus estudiantes con el Concejo Municipal, la Asamblea Departamental, la Secretaría de Movilidad y la organización juvenil COSUR; y Jairo Morera Lizcano establece alianzas con Ciudad Limpia, el Ejército, el ICBF, la Personería, el Inder Neiva y la Contraloría.",
      particularidades: "San Antonio de Anaconia presenta, con diferencia, las acciones más innovadoras y específicamente democráticas de todo el grupo: mesas comunitarias de co-creación del PEI, jornadas de rendición pública de cuentas pedagógicas, murales y boletines informativos comunitarios, y presupuestos participativos escolares, un conjunto de herramientas de transparencia y participación directa sin equivalente en ninguna otra institución analizada hasta ahora en ningún grupo. La Escuela Normal Superior, por su parte, es la única institución en fundamentar su estrategia de participación en los Proyectos Pedagógicos de Aula como mecanismo que vincula a estudiantes y familias como actores que \"investigan, aportan, validan y transforman\", en lugar de instancias tradicionales de representación."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Equipos de trabajo",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar las estrategias y mecanismos de participación escolar?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Gobierno Escolar", count: 6},
        {opcion: "Consejo Directivo", count: 6},
        {opcion: "Consejo Académico", count: 6},
        {opcion: "Consejo Estudiantil", count: 6},
        {opcion: "Comité Escolar de Convivencia", count: 6},
        {opcion: "Personero(a) Estudiantil", count: 5},
        {opcion: "Consejo de Padres de Familia", count: 4},
        {opcion: "Contralor(a) Estudiantil", count: 4},
        {opcion: "Comisión de Evaluación y Promoción", count: 4}
      ],
      comun: "El Grupo 5 presenta el nivel de consenso más alto observado en esta pregunta entre todos los grupos analizados: el Gobierno Escolar en su conjunto, el Consejo Directivo, el Consejo Académico, el Consejo Estudiantil y el Comité Escolar de Convivencia son adoptados por las seis instituciones, sin excepción. Con una adopción del 83,3% (5 de 6 IE) se ubica el/la Personero(a) Estudiantil; y con una adopción del 66,7% (4 de 6 IE), el Consejo de Padres de Familia, el/la Contralor(a) Estudiantil y la Comisión de Evaluación y Promoción. Este patrón confirma al Grupo 5 como el de estructura de gobierno escolar más homogénea de todos los grupos analizados hasta el momento.",
      particularidades: "El Instituto Técnico IPC Andrés Rosa vuelve a presentar, como en la pregunta análoga de la Sesión 2, la estructura de equipos más extensa del grupo, complementando el catálogo reglamentario con seis opciones \"Otro\" adicionales, entre ellas una Mesa de Diálogo Currículo-Comunidad y un Comité Curricular Institucional. Jairo Morera Lizcano, en el extremo opuesto, es la única institución del grupo que no reporta ni Personero(a) ni Contralor(a) Estudiantil entre sus equipos de participación, apoyándose en cambio en el acompañamiento de orientación escolar y en un proyecto de democracia propio."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Mecanismos de seguimiento",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para garantizar las acciones encaminadas a promover gobiernos educativos democráticos?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Autoevaluación institucional anual (Guía 34 – MEN)", count: 6},
        {opcion: "Reuniones periódicas con entes de gobierno escolar", count: 5},
        {opcion: "Elecciones estudiantiles", count: 5},
        {opcion: "Elección de docentes para consejo directivo", count: 4},
        {opcion: "Elección de docentes para consejo académico", count: 4},
        {opcion: "Plan de Mejoramiento Institucional (PMI)", count: 4},
        {opcion: "Actualización y ajuste permanente del PEI", count: 4},
        {opcion: "Valoración por instancias del Gobierno Escolar", count: 4}
      ],
      comun: "La autoevaluación institucional anual es el único mecanismo de seguimiento que las seis instituciones del grupo, sin excepción, reportan haber implementado para garantizar los gobiernos educativos democráticos. Con una adopción del 83,3% (5 de 6 IE) se ubican las reuniones periódicas con entes de gobierno escolar y las elecciones estudiantiles; y con una adopción del 66,7% (4 de 6 IE), la elección de docentes para los consejos directivo y académico, el Plan de Mejoramiento Institucional, la actualización del PEI y la valoración por instancias del Gobierno Escolar. Este patrón muestra un núcleo de seguimiento razonablemente homogéneo, aunque con más variación que la observada en la pregunta análoga de equipos de trabajo.",
      particularidades: "Santa Teresa y Jairo Morera Lizcano son las dos instituciones del grupo que más se apartan del catálogo estándar de mecanismos, complementando su selección reducida con descripciones propias extensas en la opción \"Otro\": Santa Teresa detalla actas de reuniones, cronogramas, planes de acción, registros de participación y una estrategia nombrada de \"Gobierno Alterno\" sin equivalente en ninguna otra institución del grupo; y Jairo Morera Lizcano vincula su seguimiento con un \"Proyecto de Eduderechos\" propio. El Instituto Técnico IPC Andrés Rosa y San Antonio de Anaconia, por su parte, son las únicas instituciones en reportar el seguimiento al componente de Cultura Institucional y Clima Escolar como mecanismo explícito."
    }
  ],
  conclusiones: "El análisis conjunto de las tres sesiones de trabajo del Grupo 5 permite concluir que las seis instituciones educativas (Santa Teresa, Escuela Normal Superior, Instituto Técnico IPC Andrés Rosa, Juan de Cabrera, Jairo Morera Lizcano y San Antonio de Anaconia) comparten una trayectoria de fortalecimiento de proyectos pedagógicos contextualizados y una estructura de gobierno escolar particularmente homogénea y consolidada —con el Gobierno Escolar, sus tres consejos principales y el Comité Escolar de Convivencia adoptados por la totalidad del grupo—, junto con una tensión de fondo, compartida por las seis instituciones aunque resuelta de manera distinta en cada una, entre los referentes curriculares nacionales y las particularidades sociales, culturales y productivas del territorio. Esta tensión llevó al Grupo 5 a presentar la valoración más dividida sobre la pertinencia curricular observada hasta ahora: mientras tres instituciones (Santa Teresa, Instituto Técnico IPC Andrés Rosa y Juan de Cabrera) afirman que sus currículos sí son pertinentes, Jairo Morera Lizcano concluye lo contrario de manera explícita, y ninguna otra institución en ningún grupo analizado había llegado a esa conclusión. Dos hallazgos particulares merecen destacarse: la disminución de la tasa de natalidad reportada por el Instituto Técnico IPC Andrés Rosa como causa de la contracción de su oferta de preescolar, un fenómeno demográfico y no institucional; y el conjunto de herramientas de participación directa y transparencia —mesas de co-creación del PEI, rendición pública de cuentas, presupuestos participativos escolares— reportado por San Antonio de Anaconia, sin equivalente en ninguna otra institución analizada. En cuanto a la valoración general del Foro, de las seis instituciones del grupo, la Escuela Normal Superior y el Instituto Técnico IPC Andrés Rosa registraron formalmente su valoración de la jornada, con calificaciones de 5,0/5,0 y 4,8/5,0 respectivamente; las cuatro instituciones restantes (Santa Teresa, Juan de Cabrera, Jairo Morera Lizcano y San Antonio de Anaconia) aún no habían registrado su valoración a la fecha de consolidación de este informe. Se recomienda a la Secretaría de Educación gestionar el diligenciamiento pendiente de este instrumento con las cuatro instituciones señaladas, y dar seguimiento particular a las condiciones de infraestructura tecnológica reportadas por Jairo Morera Lizcano.",
  proyectoNombre: "[Nombre de quien consolida el Informe Consolidado]",
  proyectoCargo: "[Cargo]",
  fechaRealizacion: "7 de septiembre de 2026"
};

/*
 * Genera (o reescribe, si ya existe) el Informe de Síntesis Grupal del
 * Grupo 5, dejándolo en la raíz de la carpeta "Grupo G5" (mismo
 * criterio de idempotencia que el resto de documentos de grupo).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El enlace del documento queda en "Ver
 * registros de ejecución".
 */
function generarInformeSintesisGrupo5FEM(){
  const carpetas=crearEstructuraCarpetasGrupoFEM_("G5");
  const nombreDoc="Informe de Síntesis - Grupo G5 FEM 2026";
  const existentesIt=carpetas.grupoFolder.getFilesByName(nombreDoc);
  const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
  while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

  const archivoDoc=generarInformeSintesisGrupoFEM_(DATOS_SINTESIS_GRUPO_5_FEM_, archivoExistente?archivoExistente.getId():null);
  if(!archivoExistente){
    carpetas.grupoFolder.addFile(archivoDoc);
    try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
  }

  Logger.log("========================================");
  Logger.log("INFORME DE SÍNTESIS — GRUPO 5 — RESULTADO");
  Logger.log("Documento: "+archivoDoc.getUrl());
  Logger.log("========================================");

  return archivoDoc.getUrl();
}

/*
 * INFORME DE SÍNTESIS GRUPAL — GRUPO 6.
 *
 * Mismo criterio que los grupos anteriores: prosa redactada a mano
 * leyendo directamente el Doc editable ya generado de cada IE del
 * grupo, y tallas de las 4 preguntas mixtas contadas a mano sobre las
 * opciones realmente seleccionadas.
 *
 * Las 6 IE del catálogo (Oliverio Lara Borrero, Agustín Codazzi, El
 * Limonar, Rodrigo Lara Bonilla, El Caguán, San Luis Beltrán) SÍ habían
 * remitido su Informe Ejecutivo a la fecha de redacción (7 de
 * septiembre de 2026), por lo que el denominador de las 4 preguntas
 * mixtas es 6.
 *
 * De esas 6 IE, El Limonar (4.8/5) y Rodrigo Lara Bonilla (4.8/5)
 * habían registrado su Valoración del Foro a la fecha de redacción.
 */
const DATOS_SINTESIS_GRUPO_6_FEM_ = {
  grupo: "G6",
  tituloInforme: "Informe de Síntesis Grupal — Elementos Comunes y Particularidades Institucionales",
  instituciones: [
    "OLIVERIO LARA BORRERO",
    "AGUSTIN CODAZZI",
    "EL LIMONAR",
    "RODRIGO LARA BONILLA",
    "EL CAGUAN",
    "SAN LUIS BELTRAN"
  ],
  responsableInforme: "[Nombre de quien consolida el informe de síntesis del Grupo 6 — SEM Neiva]",
  fechaPresentacion: "7 de septiembre de 2026",
  secciones: [
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta orientadora",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en el logro de los retos y propósitos planteados en el FEM2025?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del Grupo 6 (Oliverio Lara Borrero, Agustín Codazzi, El Limonar, Rodrigo Lara Bonilla, El Caguán y San Luis Beltrán) coinciden en reportar avances en la contextualización de sus prácticas y proyectos pedagógicos: los centros de interés PILEO y los Desafíos CRESE en Oliverio Lara Borrero, el fortalecimiento curricular y la articulación con el SENA en Agustín Codazzi, la ampliación de la educación inicial en El Limonar, el Proyecto Ambiental Escolar y la atención inclusiva mediante PIAR en Rodrigo Lara Bonilla, los proyectos de turismo, emprendimiento y gastronomía local en El Caguán, y la huerta escolar y el PRAE en San Luis Beltrán. De manera igualmente compartida, cuatro de las seis instituciones (Agustín Codazzi, Rodrigo Lara Bonilla, El Caguán y San Luis Beltrán) identifican las limitaciones de infraestructura y de recursos económicos como el obstáculo estructural más persistente para sostener estos avances en el tiempo, mientras que tres instituciones (Oliverio Lara Borrero, Agustín Codazzi y San Luis Beltrán) señalan de manera explícita la necesidad de fortalecer la articulación y el compromiso de las familias como condición para consolidarlos.",
      particularidades: "El Caguán presenta la particularidad más crítica del grupo en esta pregunta, al señalar la inestabilidad de la figura de la rectoría como un factor que ha afectado directamente la continuidad de procesos institucionales clave —la eficiencia en el manejo de recursos, la toma de decisiones y el seguimiento de las acciones propuestas—, un problema de gobernanza institucional no reportado por ninguna otra institución del grupo. El Limonar, por su parte, es la única institución en fundamentar sus avances explícitamente en el marco normativo de la educación inicial —el Documento N.° 20 del MEN, el artículo 5 de la Ley 115 de 1994 y los propósitos del MEN de 2014— como referente estructurante de su quehacer pedagógico."
    },
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta 2",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en la implementación de los nuevos grados del nivel de preescolar (jardín, prejardín)?",
      tipo: "cualitativo",
      comun: "Cuatro de las seis instituciones del grupo (Oliverio Lara Borrero, Agustín Codazzi, El Limonar y Rodrigo Lara Bonilla) ya tienen en funcionamiento el grado jardín, con distintos niveles de consolidación del prejardín, mientras que las otras dos presentan las situaciones más críticas del grupo en esta pregunta: El Caguán reconoce no haber logrado ningún avance en la implementación de jardín o prejardín por falta de infraestructura, espacios, materiales y mobiliario adecuados; y San Luis Beltrán reporta que, pese a haber gestionado una docente específica para estos grados, no logró alcanzar la matrícula mínima necesaria para sostener la iniciativa, atribuyendo esta situación, en parte, a que muchas familias prefieren mantener a sus hijos en los hogares de bienestar familiar por los beneficios de alimentación que estos ofrecen. Entre las cuatro instituciones que sí operan el nivel, persiste la limitación de infraestructura y dotación: Oliverio Lara Borrero necesita continuar adecuando y dotando los espacios ya abiertos, Agustín Codazzi señala la sede Emayá como prioritaria, y Rodrigo Lara Bonilla reconoce \"grandes vacíos en dotación de material didáctico especializado\" pese a haber ampliado su cobertura.",
      particularidades: "El Caguán aporta a este informe el relato más ilustrativo de todos los grupos analizados sobre las consecuencias concretas de la falta de infraestructura: una docente asignada inicialmente a prejardín y jardín en la sede principal tuvo que ser reasignada a otra sede sin el grado de transición y sin infraestructura adecuada, donde finalmente se adecuó \"un salón improvisado\" para atender a sus estudiantes. El Limonar, por su parte, es la única institución del grupo en describir con precisión técnica su modelo pedagógico de educación inicial —Aprendizaje Significativo con enfoque humanista, organizado en cuatro periodos académicos articulados por hilos conductores—, un nivel de formalización curricular no replicado por las demás instituciones."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que los currículos actuales que se desarrollan en las instituciones educativas son pertinentes con sus realidades territoriales (sociales, culturales, productivas)? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Cinco de las seis instituciones del grupo (Oliverio Lara Borrero, Agustín Codazzi, El Limonar, Rodrigo Lara Bonilla y El Caguán) coinciden en describir su currículo como pertinente de manera parcial o en proceso de consolidación, mientras que San Luis Beltrán plantea la pregunta en términos distintos, condicionando la pertinencia curricular a una participación más activa y real de toda la comunidad educativa más que a un juicio directo sobre los contenidos. Como elemento compartido, tres instituciones fundamentan su pertinencia en proyectos culturales o identitarios propios y nombrados: el San Pedrito y el proyecto Palabrerío en Oliverio Lara Borrero, la experiencia etnoeducativa de la sede La Gabriela —con población indígena, en concordancia con el Decreto 804 de 1995— en El Caguán, y el proyecto Lectuvalores en El Limonar. El Caguán es, de las seis, la única institución que afirma sin matices que su currículo sí es pertinente con la realidad del territorio, respaldando esa afirmación con ejemplos concretos y nombrados de dos de sus sedes.",
      particularidades: "El Caguán presenta la particularidad etnoeducativa más singular de todo el grupo, al describir cómo su sede La Gabriela, con población indígena, desarrolla procesos educativos que reconocen su identidad y sus saberes propios en concordancia con el Decreto 804 de 1995, una dimensión de pertinencia étnica no reportada por ninguna otra institución del grupo. Rodrigo Lara Bonilla, por su parte, es la única institución en fundamentar su pertinencia curricular en una oferta de especialización técnica nombrada y diferenciada —Turismo y Diseño Gráfico— articulada con la intensificación en inglés."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se han implementado para lograr currículos más pertinentes al territorio?",
      tipo: "cualitativo",
      comun: "Cinco de las seis instituciones del grupo (Oliverio Lara Borrero, Agustín Codazzi, El Limonar, Rodrigo Lara Bonilla y El Caguán) coinciden en reportar la revisión y actualización periódica de los planes de área y de estudio —generalmente durante la semana institucional de inicio de año— como la acción central para lograr currículos más pertinentes, mientras que San Luis Beltrán concentra sus acciones, en cambio, en el fortalecimiento de los espacios de participación de la comunidad como vía indirecta hacia la pertinencia curricular. Como segunda línea compartida por cuatro de las seis instituciones, el fortalecimiento de proyectos pedagógicos transversales e interdisciplinarios aparece de manera explícita en Oliverio Lara Borrero (Oliverio TV, Palabrerío, PILEO), Agustín Codazzi, El Limonar (proyectos transversales, Lectuvalores) y Rodrigo Lara Bonilla (liderazgo en sostenibilidad ambiental desde Ciencias Naturales).",
      particularidades: "El Caguán es la única institución del grupo en fundamentar explícitamente una de sus acciones curriculares en el principio de la libertad de cátedra, señalando que \"en la práctica cotidiana y bajo la figura de la libertad de cátedra, cada docente adecua los contenidos y competencias a las necesidades del aula\", una autonomía docente individual formulada en estos términos por ninguna otra institución del grupo. Oliverio Lara Borrero, por su parte, es la única institución en vincular explícitamente su transformación curricular con un canal de comunicación audiovisual propio y nombrado (Oliverio TV) como estrategia pedagógica."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 3",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar estas acciones?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Consejo Académico", count: 6},
        {opcion: "Comisión de Evaluación y Promoción", count: 6},
        {opcion: "Comités de área", count: 5},
        {opcion: "Consejo Directivo", count: 5},
        {opcion: "Consejo Estudiantil", count: 5},
        {opcion: "Personero(a) Estudiantil", count: 5},
        {opcion: "Equipo de Preescolar", count: 4},
        {opcion: "Consejo de Padres de Familia", count: 4}
      ],
      comun: "El Consejo Académico y la Comisión de Evaluación y Promoción son los dos equipos de trabajo que las seis instituciones del grupo, sin excepción, reportan haber conformado para esta pregunta. Con una adopción del 83,3% (5 de 6 IE) se ubican los comités de área, el Consejo Directivo, el Consejo Estudiantil y el/la Personero(a) Estudiantil; y con una adopción del 66,7% (4 de 6 IE), el Equipo de Preescolar y el Consejo de Padres de Familia. Este patrón muestra una base reglamentaria muy sólida y ampliamente compartida por casi todo el grupo, con Rodrigo Lara Bonilla como la única excepción sistemática en la mayoría del catálogo.",
      particularidades: "El Limonar presenta, con quince equipos reportados, la estructura más extensa y diferenciada del grupo, incluyendo un Equipo de Seguimiento a Egresados sin equivalente en ninguna otra institución. Rodrigo Lara Bonilla, en el extremo opuesto, es la institución con la selección más reducida del grupo —apenas cinco equipos—, sin reportar Consejo Directivo, comités de área, Consejo Estudiantil ni Personero(a) Estudiantil para esta pregunta, una discrepancia notable frente al resto del grupo."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 4",
      enunciado: "¿Cómo se están articulando estos equipos de trabajo para lograr currículos más pertinentes territorialmente?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del grupo coinciden en describir la articulación entre sus equipos de trabajo mediante reuniones periódicas, comités de área y el Consejo Académico como instancia orientadora central, apoyada en actas, circulares y agendas de planeación. Agustín Codazzi es la única institución del grupo en formular una autocrítica explícita sobre esta articulación, reconociendo que \"aún existe cierta fragmentación\" y que se requiere \"fortalecer una planeación común y mayor coordinación entre proyectos, áreas, comités y estamentos\", mientras que las demás instituciones describen sus mecanismos en términos más consolidados.",
      particularidades: "El Caguán es la única institución del grupo en explicar que sus equipos de área se conforman \"teniendo en cuenta los perfiles y fortalezas de los maestros\", un criterio de conformación basado en talento individual docente no mencionado por ninguna otra institución. Oliverio Lara Borrero, por su parte, es la única institución en nombrar una Red de Aprendizaje (RIA) específica para preescolar y primaria como parte de su estructura de articulación curricular."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 5",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para que dichas acciones se cumplan?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Consejo Académico", count: 6},
        {opcion: "Comités de área o núcleos de formación", count: 6},
        {opcion: "Análisis de resultados de Pruebas Saber e ICFES", count: 5},
        {opcion: "Autoevaluación institucional anual", count: 5},
        {opcion: "Articulación con el PEI", count: 4},
        {opcion: "Actas de reunión de área", count: 4},
        {opcion: "Reportes al MEN / Planes de Mejoramiento Institucional", count: 4}
      ],
      comun: "El Consejo Académico y los comités de área son los dos mecanismos de seguimiento que las seis instituciones del grupo, sin excepción, reportan haber implementado. Con una adopción del 83,3% (5 de 6 IE) se ubican el análisis de resultados de Pruebas Saber e ICFES y la autoevaluación institucional anual; y con una adopción del 66,7% (4 de 6 IE), la articulación con el PEI, las actas de reunión de área y los reportes al MEN mediante el PMI. Este patrón confirma un núcleo de seguimiento razonablemente homogéneo en el grupo, con Rodrigo Lara Bonilla como la institución que reporta el catálogo más reducido.",
      particularidades: "El Limonar presenta, con catorce mecanismos reportados, el catálogo de seguimiento más extenso del grupo, siendo la única institución en incluir la evaluación por matriz DOFA como herramienta formal. San Luis Beltrán, por su parte, es la única institución del grupo en reportar los propios Foros Educativos Municipales como mecanismo adicional de seguimiento curricular, vinculando explícitamente el ejercicio del FEM con su seguimiento institucional interno."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que la toma de decisiones en las instituciones educativas actualmente es participativa y democrática? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Las seis instituciones del grupo coinciden en reconocer la existencia de estructuras formales de gobierno escolar democrático. Cuatro de ellas —Oliverio Lara Borrero, Agustín Codazzi, Rodrigo Lara Bonilla y San Luis Beltrán— matizan esa existencia formal señalando una brecha explícita entre los mecanismos disponibles y su incidencia real: Oliverio Lara Borrero observa que \"frecuentemente las decisiones se comunican cuando ya han sido tomadas, reduciendo el margen para la concertación\"; Rodrigo Lara Bonilla reconoce \"notable apatía y ausencia de compromiso en la asistencia a los estamentos institucionales\"; y San Luis Beltrán reporta una \"percepción dividida entre las mesas de trabajo\", con la mayoría concluyendo que la participación sigue siendo \"limitada e incompleta\". El Limonar y El Caguán, en cambio, describen su ejercicio democrático en términos más consolidados y sin una autocrítica equivalente, aunque ambas reconocen aspectos puntuales por fortalecer, como la comunicación institucional oficial señalada por El Caguán.",
      particularidades: "San Luis Beltrán es la única institución del grupo en documentar explícitamente una \"percepción dividida entre las mesas de trabajo\" sobre esta misma pregunta, evidenciando que el desacuerdo sobre el estado de la democracia escolar no se da solo entre instituciones sino también al interior de una misma comunidad educativa. Oliverio Lara Borrero, por su parte, es la única institución en formular de manera explícita la distinción conceptual entre \"escuchar\" y \"tener capacidad de incidir y decidir\" como el núcleo del reto democrático pendiente."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se están implementando para canalizar y fortalecer la participación de la comunidad educativa?",
      tipo: "cualitativo",
      comun: "Cinco de las seis instituciones del grupo (Oliverio Lara Borrero, Agustín Codazzi, Rodrigo Lara Bonilla, El Caguán y San Luis Beltrán) coinciden en reportar el fortalecimiento del Gobierno Escolar y de sus procesos electorales como acción central para canalizar la participación, mientras que El Limonar concentra sus acciones, en cambio, en la Escuela de Padres y en un conjunto notable de convenios interinstitucionales externos (UIS, San Jorge, Comfamiliar, Alive, Banco de la República, SENA). Como segunda línea compartida por cuatro de las seis instituciones (Agustín Codazzi, El Limonar, El Caguán y San Luis Beltrán), el diálogo directo con las familias mediante la Escuela de Padres y los espacios de entrega de informes académicos aparece como estrategia concreta de vinculación familiar.",
      particularidades: "El Limonar presenta la particularidad más singular del grupo en esta pregunta, al fundamentar su fortalecimiento de la participación en una red de alianzas interinstitucionales excepcionalmente amplia y diversa —que incluye una universidad (UIS), una entidad de salud (San Jorge), una caja de compensación (Comfamiliar), el Banco de la República y el SENA—, un nivel de articulación externa nombrada muy superior al de cualquier otra institución del grupo. Oliverio Lara Borrero, por su parte, es la única institución en proyectar la creación de una \"escuela de liderazgo\" propia como estrategia futura para el fortalecimiento de habilidades democráticas estudiantiles."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Equipos de trabajo",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar las estrategias y mecanismos de participación escolar?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Consejo de Padres de Familia", count: 6},
        {opcion: "Comisión de Evaluación y Promoción", count: 6},
        {opcion: "Gobierno Escolar", count: 5},
        {opcion: "Consejo Directivo", count: 5},
        {opcion: "Consejo Académico", count: 5},
        {opcion: "Personero(a) Estudiantil", count: 5},
        {opcion: "Contralor(a) Estudiantil", count: 5},
        {opcion: "Comité Escolar de Convivencia", count: 5}
      ],
      comun: "El Consejo de Padres de Familia y la Comisión de Evaluación y Promoción son los dos equipos que las seis instituciones del grupo, sin excepción, reportan haber conformado para las estrategias de participación escolar. Con una adopción del 83,3% (5 de 6 IE) se ubican el Gobierno Escolar en su conjunto, el Consejo Directivo, el Consejo Académico, el/la Personero(a) y Contralor(a) Estudiantil, y el Comité Escolar de Convivencia. Este patrón muestra una estructura de gobierno escolar homogénea entre cinco de las seis instituciones, con El Caguán como la única excepción marcada del grupo.",
      particularidades: "El Caguán constituye, con notable diferencia, la particularidad más marcada del grupo en esta pregunta: es la única institución que selecciona apenas dos equipos del catálogo reglamentario, complementando con dos opciones \"Otro\" que ni siquiera especifica, pese a que el texto cualitativo de su propio informe describe una estructura de participación considerablemente más rica y articulada —incluyendo comités de proyectos transversales y Escuela de Padres— que no llegó a reflejarse en las casillas seleccionadas, una inconsistencia entre la narrativa y el catálogo marcado que no se observa en las demás instituciones del grupo."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Mecanismos de seguimiento",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para garantizar las acciones encaminadas a promover gobiernos educativos democráticos?",
      tipo: "mixto",
      totalIE: 6,
      tally: [
        {opcion: "Reuniones periódicas con entes de gobierno escolar", count: 6},
        {opcion: "Elecciones estudiantiles", count: 5},
        {opcion: "Elección de docentes para consejo directivo", count: 5},
        {opcion: "Autoevaluación institucional anual (Guía 34 – MEN)", count: 5},
        {opcion: "Plan de Mejoramiento Institucional (PMI)", count: 5},
        {opcion: "Actualización y ajuste permanente del PEI", count: 5},
        {opcion: "Elección de docentes para consejo académico", count: 4}
      ],
      comun: "Las reuniones periódicas con los distintos entes de gobierno escolar constituyen el único mecanismo de seguimiento que las seis instituciones del grupo, sin excepción, reportan haber implementado. Con una adopción del 83,3% (5 de 6 IE) se ubican las elecciones estudiantiles, la elección de docentes para el consejo directivo, la autoevaluación institucional anual, el Plan de Mejoramiento Institucional y la actualización permanente del PEI. Este patrón confirma, en paralelo con la pregunta análoga de equipos de trabajo, que El Caguán es sistemáticamente la institución que más se aparta del comportamiento homogéneo del resto del grupo.",
      particularidades: "El Caguán presenta, de nuevo, el catálogo más reducido y menos formalizado del grupo, describiendo en cambio en su respuesta cualitativa un conjunto rico de mecanismos propios —Personería, Contraloría, Consejo Estudiantil, Consejo Directivo y Académico, acompañamiento directo de la Secretaría de Educación mediante mesas de concertación, y \"mecanismos de control ciudadano directo\" ejercidos por padres y estudiantes mediante asambleas y pronunciamientos públicos—, una descripción de veeduría comunitaria activa sin equivalente en ninguna otra institución del grupo, aunque no reflejada en el catálogo estándar de casillas marcadas. El Limonar, por su parte, es la única institución en reportar el seguimiento al componente de Cultura Institucional y Clima Escolar como mecanismo explícito."
    }
  ],
  conclusiones: "El análisis conjunto de las tres sesiones de trabajo del Grupo 6 permite concluir que las seis instituciones educativas (Oliverio Lara Borrero, Agustín Codazzi, El Limonar, Rodrigo Lara Bonilla, El Caguán y San Luis Beltrán) comparten una trayectoria de fortalecimiento de proyectos pedagógicos contextualizados y una estructura de gobierno escolar razonablemente homogénea, sustentada de manera universal en el Consejo de Padres de Familia, la Comisión de Evaluación y Promoción y las reuniones periódicas con los entes de gobierno escolar, junto con retos comunes en materia de infraestructura, dotación y participación real de las familias más allá de la existencia formal de espacios democráticos. Tres hallazgos particulares merecen atención específica de la Secretaría de Educación Municipal: la inestabilidad de la rectoría reportada por El Caguán como factor que ha afectado la continuidad de sus procesos institucionales, incluyendo el manejo de recursos y el seguimiento de acciones; la experiencia relatada por la misma institución de una docente de primera infancia reasignada por falta de infraestructura hasta terminar atendiendo a sus estudiantes en un \"salón improvisado\"; y el caso de San Luis Beltrán, donde el esfuerzo por abrir prejardín y jardín no logró sostenerse por falta de matrícula suficiente, en parte por la competencia de los hogares de bienestar familiar. En cuanto a la valoración general del Foro, de las seis instituciones del grupo, El Limonar y Rodrigo Lara Bonilla registraron formalmente su valoración de la jornada, ambas con 4,8 sobre 5,0; las cuatro instituciones restantes (Oliverio Lara Borrero, Agustín Codazzi, El Caguán y San Luis Beltrán) aún no habían registrado su valoración a la fecha de consolidación de este informe. Se recomienda a la Secretaría de Educación gestionar el diligenciamiento pendiente de este instrumento con las cuatro instituciones señaladas, y dar seguimiento prioritario a la situación de gobernanza institucional reportada por El Caguán.",
  proyectoNombre: "[Nombre de quien consolida el Informe Consolidado]",
  proyectoCargo: "[Cargo]",
  fechaRealizacion: "7 de septiembre de 2026"
};

/*
 * Genera (o reescribe, si ya existe) el Informe de Síntesis Grupal del
 * Grupo 6, dejándolo en la raíz de la carpeta "Grupo G6" (mismo
 * criterio de idempotencia que el resto de documentos de grupo).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El enlace del documento queda en "Ver
 * registros de ejecución".
 */
function generarInformeSintesisGrupo6FEM(){
  const carpetas=crearEstructuraCarpetasGrupoFEM_("G6");
  const nombreDoc="Informe de Síntesis - Grupo G6 FEM 2026";
  const existentesIt=carpetas.grupoFolder.getFilesByName(nombreDoc);
  const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
  while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

  const archivoDoc=generarInformeSintesisGrupoFEM_(DATOS_SINTESIS_GRUPO_6_FEM_, archivoExistente?archivoExistente.getId():null);
  if(!archivoExistente){
    carpetas.grupoFolder.addFile(archivoDoc);
    try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
  }

  Logger.log("========================================");
  Logger.log("INFORME DE SÍNTESIS — GRUPO 6 — RESULTADO");
  Logger.log("Documento: "+archivoDoc.getUrl());
  Logger.log("========================================");

  return archivoDoc.getUrl();
}

/*
 * INFORME DE SÍNTESIS MUNICIPAL — CONSOLIDADO GRUPOS G1 A G6.
 *
 * Consolida en un solo documento las once preguntas de las tres
 * sesiones de trabajo para las 34 instituciones educativas (de las 36
 * del catálogo municipal) que remitieron su Informe Ejecutivo, más un
 * capítulo especial sobre la percepción de la comunidad educativa
 * recogida mediante firma por código QR.
 *
 * Metodología: la prosa de "comun"/"particularidades" de cada una de
 * las 11 preguntas se redactó a partir de los seis Informes de
 * Síntesis Grupal ya generados (DATOS_SINTESIS_GRUPO_1_FEM_ a
 * DATOS_SINTESIS_GRUPO_6_FEM_), identificando los patrones que se
 * repiten en el conjunto de los seis grupos y elevando, como
 * particularidades municipales, los 2-3 hallazgos institucionales más
 * significativos de todo el Foro (no uno por grupo). Las 4 tablas de
 * preguntas mixtas agregan, opción por opción, los conteos ya
 * publicados en los seis informes grupales (que a su vez surgieron de
 * la lectura directa y el conteo manual de los Informes Ejecutivos
 * reales); como cada informe grupal reportó solo las opciones de mayor
 * adopción dentro de su grupo, los porcentajes de las opciones de
 * menor adopción son un piso mínimo, mientras que el ordenamiento de
 * las opciones más adoptadas es plenamente confiable (así se advierte
 * en la Presentación del propio documento).
 *
 * PERCEPCIÓN QR: de las 34 IE con informe, solo 4 usaron firma por
 * código QR con encuesta de percepción demográfica: Liceo de Santa
 * Librada (G1, 69 firmantes), Atanasio Girardot (G4, 64), El Limonar
 * (G6, 101) y Rodrigo Lara Bonilla (G6, 52) — 286 firmantes en total,
 * cifra prácticamente coincidente con los 287 solicitados. El resto de
 * instituciones registró su asistencia mediante PDF escaneado, sin
 * encuesta de percepción digital.
 */
const DATOS_SINTESIS_MUNICIPAL_FEM_ = {
  tituloInforme: "Informe de Síntesis Municipal — Foro Educativo Institucional Neiva 2026",
  responsableInforme: "[Nombre de quien consolida el Informe de Síntesis Municipal — SEM Neiva]",
  fechaPresentacion: "8 de septiembre de 2026",
  totalConInforme: "34 de 36",
  grupos: [
    {
      grupo: "G1",
      instituciones: ["AIPECITO","CHAPINERO","I.E. CLARETIANO GUSTAVO TORRES PARRA","INEM JULIAM MOTTA SALAS","LICEO DE SANTA LIBRADA","PROMOCION SOCIAL"],
      sinInforme: []
    },
    {
      grupo: "G2",
      instituciones: ["MARIA CRISTINA ARANGO DE PASTRANA.","LUIS IGNACIO ANDRADE","GABRIEL GARCIA MARQUEZ","EDUARDO SANTOS","MARIA AUXILIADORA FORTALECILLAS","JAIRO MOSQUERA MORENO"],
      sinInforme: ["MARIA CRISTINA ARANGO DE PASTRANA."]
    },
    {
      grupo: "G3",
      instituciones: ["TECNICO SUPERIOR","DEPARTAMENTAL TIERRA DE PROMISION","SANTA LIBRADA","RICARDO BORRERO ALVAREZ","ANGEL MARIA PAREDES","CEINAR"],
      sinInforme: []
    },
    {
      grupo: "G4",
      instituciones: ["JOSE EUSTASIO RIVERA","ATANASIO GIRARDOT","MISAEL PASTRANA BORRERO","HUMBERTO TAFUR CHARRY","ENRIQUE OLAYA HERRERA","ROBERTO DURAN ALVIRA"],
      sinInforme: ["HUMBERTO TAFUR CHARRY"]
    },
    {
      grupo: "G5",
      instituciones: ["SANTA TERESA","ESCUELA NORMAL SUPERIOR","INSTITUTO TECNICO IPC ANDRES ROSA","JUAN DE CABRERA","JAIRO MORERA LIZCANO","SAN ANTONIO DE ANACONIA"],
      sinInforme: []
    },
    {
      grupo: "G6",
      instituciones: ["OLIVERIO LARA BORRERO","AGUSTIN CODAZZI","EL LIMONAR","RODRIGO LARA BONILLA","EL CAGUAN","SAN LUIS BELTRAN"],
      sinInforme: []
    }
  ],
  participacion: {
    totalParticipantes: 3196,
    porGrupo: [
      {grupo: "G1", totalIE: 6, totalParticipantes: 582},
      {grupo: "G2", totalIE: 5, totalParticipantes: 338},
      {grupo: "G3", totalIE: 6, totalParticipantes: 771},
      {grupo: "G4", totalIE: 5, totalParticipantes: 430},
      {grupo: "G5", totalIE: 6, totalParticipantes: 494},
      {grupo: "G6", totalIE: 6, totalParticipantes: 581}
    ]
  },
  secciones: [
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta orientadora",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en el logro de los retos y propósitos planteados en el FEM2025?",
      tipo: "cualitativo",
      comun: "Las 34 instituciones educativas de Neiva que remitieron su Informe Ejecutivo, agrupadas en los seis grupos de trabajo del Foro (G1 a G6), coinciden de manera prácticamente unánime en un primer hallazgo: el avance hacia los retos del FEM2025 se sostiene sobre todo en proyectos pedagógicos transversales y experiencias significativas propias de cada institución, más que en una transformación curricular estructural uniforme. Esto se expresa mediante iniciativas nombradas y distintas en cada caso —huertas escolares y alianzas con el SENA en el Grupo 1, el fortalecimiento de la primera infancia y los centros de interés en el Grupo 2, el hilo conductor institucional y las alianzas con el SENA y universidades en el Grupo 3, la educación inclusiva mediante PIAR y DUA en el Grupo 4, los macroproyectos y los Proyectos Pedagógicos de Aula en el Grupo 5, y los centros de interés PILEO/CRESE y los proyectos de turismo y emprendimiento en el Grupo 6—, pero conducen todas a la misma conclusión de fondo: el progreso pedagógico depende, en gran medida, del esfuerzo individual de cada institución y de sus equipos docentes. Un segundo hallazgo, igualmente transversal a los seis grupos, es que las limitaciones de infraestructura, dotación y conectividad constituyen el obstáculo estructural más persistente y más ampliamente repetido de todo el Foro, señalado por la mayoría de las instituciones de cada uno de los seis grupos, sin excepción. Un tercer hallazgo compartido es la necesidad, expresada de forma recurrente en al menos la mitad de las instituciones de casi todos los grupos, de fortalecer la corresponsabilidad y la participación activa de las familias como condición para sostener en el tiempo los avances alcanzados.",
      particularidades: "Entre las particularidades más críticas de las 34 instituciones, dos ameritan especial atención de la Secretaría de Educación Municipal por señalar problemas que trascienden lo estrictamente pedagógico: la IE El Caguán (Grupo 6) reportó que la inestabilidad de la figura de la rectoría ha afectado directamente la continuidad de sus procesos institucionales —incluyendo el manejo de recursos, la toma de decisiones y el seguimiento de las acciones—; y la IE Ricardo Borrero Álvarez (Grupo 3) documentó problemáticas concretas de salubridad e inseguridad (una infestación de palomas que deteriora su infraestructura, además de inconformidad con el Programa de Alimentación Escolar, el transporte escolar y episodios de microtráfico e inseguridad). En un registro distinto, la IE Eduardo Santos (Grupo 2) fue la única en formular un señalamiento directo de \"abandono\" hacia la Secretaría de Educación y la Alcaldía por falta de respuesta oportuna a sus exigencias institucionales, mientras que Roberto Durán Alvira (Grupo 4) se distinguió en el sentido opuesto, siendo la única institución de todo el Foro en cuantificar explícitamente un nivel de avance del 80% frente a las metas del FEM2025 gracias a su macroproyecto \"Cultivando saberes y conexiones rurales\"."
    },
    {
      sesion: "SESIÓN 1",
      pregunta: "Pregunta 2",
      enunciado: "¿Cómo hemos avanzado, desde nuestra institución educativa, en la implementación de los nuevos grados del nivel de preescolar (jardín, prejardín)?",
      tipo: "cualitativo",
      comun: "En los seis grupos se repite el mismo patrón: la implementación del grado jardín está mucho más consolidada que la del prejardín, y en al menos ocho de las 34 instituciones analizadas —entre ellas el INEM Juliam Motta Salas (G1), Ceinar (G3), José Eustasio Rivera, Agustín Codazzi y Enrique Olaya Herrera (G4), Juan de Cabrera (G5), y El Caguán y San Luis Beltrán (G6)— la causa principal de la baja demanda o de la nula apertura de prejardín es la misma: la competencia de los hogares comunitarios del ICBF y de los centros de cuidado privados, que ofrecen alimentación y jornada completa que la oferta escolar oficial no logra igualar. Un segundo hallazgo, igualmente generalizado en las seis mesas de trabajo, es que incluso en las instituciones donde el nivel ya opera, persisten limitaciones de infraestructura, dotación y personal formado que condicionan la calidad de la atención brindada.",
      particularidades: "Tres instituciones documentan de manera especialmente concreta las consecuencias de estas limitaciones: El Caguán (G6) relató el caso de una docente de primera infancia reasignada por falta de infraestructura hasta terminar atendiendo a sus estudiantes en un \"salón improvisado\"; Juan de Cabrera (G5) reconoció no haber logrado ninguna apertura efectiva pese a haber realizado la oferta; y San Luis Beltrán (G6) reportó que, tras gestionar una docente específica para estos grados, no logró sostener la matrícula mínima necesaria. En el extremo opuesto, el Instituto Técnico IPC Andrés Rosa (G5) aportó el único hallazgo demográfico —y no institucional— de todo el Foro: identificó una disminución en la tasa de natalidad del sector como causa de la contracción de su oferta de preescolar."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que los currículos actuales que se desarrollan en las instituciones educativas son pertinentes con sus realidades territoriales (sociales, culturales, productivas)? ¿Por qué?",
      tipo: "cualitativo",
      comun: "La pertinencia curricular es calificada por la inmensa mayoría de las 34 instituciones como parcial o en proceso de consolidación: los currículos se reconocen alineados con los lineamientos del Ministerio de Educación Nacional, pero todavía insuficientemente conectados con las realidades sociales, culturales y productivas específicas de cada territorio. Este diagnóstico compartido se sostiene en dos tensiones que se repiten en los seis grupos: la presión de las pruebas estandarizadas (Saber, ICFES, DBA) como factor que limita la flexibilización curricular necesaria para incorporar el contexto local, señalada explícitamente en los Grupos 2 y 5; y la persistencia de una \"concepción curricular homogénea\" —en palabras de San Antonio de Anaconia (G5)— que no siempre reconoce las particularidades de cada comunidad, especialmente en las instituciones rurales del Foro.",
      particularidades: "El hallazgo más contundente de todo el Foro en esta pregunta proviene de Jairo Morera Lizcano (G5): es la única de las 34 instituciones en concluir, sin matices, que su currículo \"no es pertinente con la realidad territorial\", atribuyendo la sostenibilidad de cualquier enseñanza contextualizada casi enteramente a la vocación individual de sus docentes y no a una política institucional o estatal. En el extremo étnico, El Caguán (G6) es la única institución en documentar un componente etnoeducativo explícito, con su sede La Gabriela atendiendo población indígena en concordancia con el Decreto 804 de 1995 — una dimensión de pertinencia curricular por identidad étnica sin equivalente en ninguna otra de las 34 instituciones del Foro."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se han implementado para lograr currículos más pertinentes al territorio?",
      tipo: "cualitativo",
      comun: "El fortalecimiento de proyectos pedagógicos transversales propios —con nombre y sello institucional en cada caso (huertas y PRAE en el Grupo 1, centros de interés y proyectos de vida en el Grupo 2, hilo conductor y CEINARTE en el Grupo 3, proyectos culturales y ambientales en el Grupo 4, Proyectos Pedagógicos de Aula y macroproyectos en el Grupo 5, y Lectuvalores, Oliverio TV o Palabrerío en el Grupo 6)— es, de lejos, la acción curricular más repetida en las 34 instituciones del Foro. La revisión y actualización periódica de los planes de área y de estudio, generalmente durante la semana institucional de inicio de año, es la segunda línea de acción más ampliamente compartida. Como tercer eje, al menos una institución en cinco de los seis grupos reporta alianzas formales con el SENA como estrategia concreta de pertinencia técnico-laboral, consolidando a esta entidad como el aliado externo más recurrente de todo el Foro Educativo Institucional.",
      particularidades: "El Caguán (Grupo 6) es la única institución de las 34 en fundamentar explícitamente una de sus acciones curriculares en el principio de la libertad de cátedra docente individual. Ricardo Borrero Álvarez (Grupo 3), por su parte, es la única en reportar una estrategia de apropiación estudiantil del currículo mediante roles asignados a los propios estudiantes como líderes y guardianes ambientales, de convivencia, de TIC y de rendimiento académico."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 3",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar estas acciones?",
      tipo: "mixto",
      totalIE: 34,
      tally: [
        {opcion: "Consejo Académico", count: 33},
        {opcion: "Comités de área", count: 31},
        {opcion: "Consejo Directivo", count: 24},
        {opcion: "Consejo Estudiantil", count: 19},
        {opcion: "Comisión de Evaluación y Promoción", count: 18},
        {opcion: "Personero(a) Estudiantil", count: 16},
        {opcion: "Consejo de Padres de Familia", count: 13},
        {opcion: "Equipo de Preescolar", count: 12}
      ],
      comun: "El Consejo Académico es la instancia que con mayor extensión lidera las acciones hacia currículos más pertinentes en todo el municipio, reportado por 33 de las 34 instituciones (97,1%). Le siguen los comités de área, presentes en 31 instituciones (91,2%), y el Consejo Directivo, en 24 (70,6%). Con una adopción minoritaria pero relevante se ubican el Consejo Estudiantil (19 IE, 55,9%), la Comisión de Evaluación y Promoción (18 IE, 52,9%), el/la Personero(a) Estudiantil (16 IE, 47,1%), el Consejo de Padres de Familia (13 IE, 38,2%) y el Equipo de Preescolar (12 IE, 35,3%). Este patrón, consistente en los seis grupos de trabajo, confirma al Consejo Académico y a los comités de área como el núcleo verdaderamente universal de la gestión curricular del Foro, mientras que las instancias de representación estudiantil y familiar, aunque mayoritarias, presentan una adopción más desigual entre instituciones y grupos.",
      particularidades: "AIPECITO (Grupo 1) sigue siendo, a escala municipal, la institución más atípica en su relación con el catálogo estándar de equipos de trabajo: en las cuatro preguntas mixtas del Foro respondió sistemáticamente por fuera de las categorías sugeridas, apoyándose en estructuras propias de corte comunitario-territorial (veeduría del PAE, comités de riesgo y medio ambiente) en lugar de las instancias reglamentarias de representación estudiantil que sí reportó la inmensa mayoría de las demás 33 instituciones. En el extremo de mayor complejidad organizativa, el Instituto Técnico IPC Andrés Rosa (Grupo 5) reportó, con cerca de veinte equipos de trabajo distintos, la estructura curricular más extensa de todo el municipio."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 4",
      enunciado: "¿Cómo se están articulando estos equipos de trabajo para lograr currículos más pertinentes territorialmente?",
      tipo: "cualitativo",
      comun: "En los seis grupos, la articulación entre los equipos de trabajo institucional se apoya, de manera prácticamente unánime, en reuniones periódicas, comités de área y el Consejo Académico como instancia orientadora central, más que en mecanismos formales creados específicamente para este fin. Una proporción significativa de instituciones en casi todos los grupos —el Liceo de Santa Librada y el INEM en el Grupo 1, varias del Grupo 2, Santa Librada y Ricardo Borrero Álvarez en el Grupo 3, la Escuela Normal Superior y Jairo Morera Lizcano en el Grupo 5, y Agustín Codazzi en el Grupo 6— reconoce de manera explícita que esta articulación resulta todavía insuficiente, dependiente de iniciativas individuales más que de una política institucional sistemática.",
      particularidades: "Atanasio Girardot (Grupo 4) constituye el hallazgo más severo de todo el Foro en esta pregunta: es la única de las 34 instituciones en admitir que su Consejo Académico no fue convocado ni una sola vez durante todo el año escolar 2026, dejando la reflexión curricular en manos exclusivas de comités de área específicos — un vacío de gobierno curricular sin equivalente en ninguna otra institución del municipio."
    },
    {
      sesion: "SESIÓN 2",
      pregunta: "Pregunta 5",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para que dichas acciones se cumplan?",
      tipo: "mixto",
      totalIE: 34,
      tally: [
        {opcion: "Autoevaluación institucional anual", count: 28},
        {opcion: "Consejo Académico", count: 26},
        {opcion: "Comités de área o núcleos de formación", count: 26},
        {opcion: "Actas de reunión de área", count: 25},
        {opcion: "Análisis de resultados de Pruebas Saber e ICFES", count: 22},
        {opcion: "Articulación con el PEI", count: 18},
        {opcion: "Plataformas de gestión académica (JIGRA)", count: 17},
        {opcion: "Reportes al MEN / Planes de Mejoramiento Institucional", count: 13}
      ],
      comun: "La autoevaluación institucional anual es el mecanismo de seguimiento a las acciones curriculares más ampliamente adoptado en las 34 instituciones del Foro, presente en 28 de ellas (82,4%). Le siguen, con una adopción muy similar y cercana a las tres cuartas partes del municipio, el Consejo Académico y los comités de área (26 IE cada uno, 76,5%) y las actas de reunión de área (25 IE, 73,5%). El análisis de resultados de Pruebas Saber e ICFES se reporta en 22 instituciones (64,7%), la articulación con el PEI en 18 (52,9%) y las plataformas de gestión académica JIGRA en 17 (50,0%). Este patrón confirma que el seguimiento curricular del municipio se apoya, sobre todo, en instrumentos internos de autoevaluación y en el análisis de resultados académicos externos, más que en mecanismos de seguimiento comunitario o de participación directa de las familias.",
      particularidades: "Gabriel García Márquez (Grupo 2) es la institución con el seguimiento más débil de todo el municipio en esta pregunta: seleccionó un único mecanismo del catálogo estándar y reconoció explícitamente carecer de una acción de seguimiento sistemática más allá de ese instrumento aislado. En el extremo opuesto, el Instituto Técnico IPC Andrés Rosa (Grupo 5) y El Limonar (Grupo 6) reportaron los catálogos de seguimiento más extensos y diversificados del Foro, incluyendo ambos, de manera excepcional, la evaluación por matriz DOFA como herramienta formal de análisis institucional."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 1",
      enunciado: "¿Consideran que la toma de decisiones en las instituciones educativas actualmente es participativa y democrática? ¿Por qué?",
      tipo: "cualitativo",
      comun: "Sin una sola excepción entre las 34 instituciones del Foro, todas reconocen la existencia de estructuras formales de gobierno escolar democrático (Consejo Directivo, Consejo Académico, Consejo Estudiantil, Consejo de Padres) y todas, también sin excepción, matizan esa existencia formal señalando límites concretos a su incidencia real sobre las decisiones institucionales. La baja participación de las familias es, con diferencia, el obstáculo específico más repetido de todo el Foro, mencionado de manera explícita por una mayoría de instituciones en cada uno de los seis grupos. Un segundo obstáculo ampliamente compartido es la percepción de que las decisiones de fondo se concentran en los equipos directivos o de gestión, señalada de forma recurrente —aunque con palabras propias en cada caso— en instituciones de los Grupos 1, 4, 5 y 6.",
      particularidades: "Ricardo Borrero Álvarez (Grupo 3) presenta la denuncia más directa de todo el Foro sobre opacidad en la comunicación de decisiones, al señalar que las decisiones clave \"se centralizan en el Equipo de Gestión sin comunicarse oportunamente a las bases\". San Luis Beltrán (Grupo 6), por su parte, es la única institución en documentar explícitamente una \"percepción dividida entre las mesas de trabajo\" sobre esta misma pregunta, evidenciando que el desacuerdo sobre el estado de la democracia escolar puede darse también al interior de una misma comunidad educativa y no solo entre instituciones."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Pregunta 2",
      enunciado: "¿Qué acciones se están implementando para canalizar y fortalecer la participación de la comunidad educativa?",
      tipo: "cualitativo",
      comun: "El fortalecimiento del Gobierno Escolar y de sus procesos electorales —elección de representantes, personeros y contralores— es la acción más repetida en las 34 instituciones del Foro para canalizar la participación de la comunidad educativa. La apertura de espacios de diálogo directo con las familias, principalmente mediante la Escuela de Padres, constituye la segunda línea de acción compartida por la mayoría de instituciones en los seis grupos. Un tercer grupo, más reducido pero presente en instituciones de casi todos los grupos, reporta alianzas interinstitucionales externas —con el SENA, el ICBF, la Policía de Infancia, universidades regionales o entidades gubernamentales— como estrategia complementaria de fortalecimiento comunitario.",
      particularidades: "San Antonio de Anaconia (Grupo 5) presenta, con diferencia, las herramientas de participación más innovadoras de todo el Foro: mesas comunitarias de co-creación del PEI, jornadas de rendición pública de cuentas pedagógicas, murales y boletines informativos comunitarios, y presupuestos participativos escolares. José Eustasio Rivera (Grupo 4), por su parte, es la única institución en reportar una veeduría comunitaria formalizada con un objeto de vigilancia concreto: un comité de acompañamiento de padres que hace seguimiento directo a la construcción de infraestructura de una de sus sedes."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Equipos de trabajo",
      enunciado: "¿Qué equipos de trabajo a nivel institucional se han conformado para liderar y desarrollar las estrategias y mecanismos de participación escolar?",
      tipo: "mixto",
      totalIE: 34,
      tally: [
        {opcion: "Gobierno Escolar", count: 30},
        {opcion: "Consejo Directivo", count: 30},
        {opcion: "Comité Escolar de Convivencia", count: 30},
        {opcion: "Consejo Académico", count: 29},
        {opcion: "Consejo de Padres de Familia", count: 29},
        {opcion: "Personero(a) Estudiantil", count: 27},
        {opcion: "Comisión de Evaluación y Promoción", count: 26},
        {opcion: "Consejo Estudiantil", count: 25},
        {opcion: "Contralor(a) Estudiantil", count: 25}
      ],
      comun: "El Gobierno Escolar en su conjunto, el Consejo Directivo y el Comité Escolar de Convivencia son los tres equipos de participación con mayor adopción en todo el municipio, cada uno reportado por 30 de las 34 instituciones (88,2%). Les siguen el Consejo Académico y el Consejo de Padres de Familia, cada uno en 29 instituciones (85,3%); el/la Personero(a) Estudiantil, en 27 (79,4%); la Comisión de Evaluación y Promoción, en 26 (76,5%); y el Consejo Estudiantil y el/la Contralor(a) Estudiantil, cada uno en 25 instituciones (73,5%). Este patrón evidencia una estructura de gobierno escolar notablemente homogénea y consolidada en las 34 instituciones del Foro, con las instancias reglamentarias de la Ley 115 alcanzando niveles de adopción cercanos o superiores al 80% en la práctica totalidad de los equipos evaluados.",
      particularidades: "El Caguán (Grupo 6) presenta, con diferencia, la estructura de participación más débil de todo el municipio en esta pregunta: seleccionó apenas dos equipos del catálogo reglamentario, una inconsistencia notable frente a la descripción cualitativa mucho más rica que la misma institución ofreció en su propio informe. En el otro extremo, Santa Librada (Grupo 3) y el Instituto Técnico IPC Andrés Rosa (Grupo 5) presentan algunas de las estructuras de participación más completas y diversificadas del Foro, complementando el catálogo reglamentario con equipos propios adicionales."
    },
    {
      sesion: "SESIÓN 3",
      pregunta: "Mecanismos de seguimiento",
      enunciado: "¿Qué mecanismos de seguimiento se están implementando para garantizar las acciones encaminadas a promover gobiernos educativos democráticos?",
      tipo: "mixto",
      totalIE: 34,
      tally: [
        {opcion: "Autoevaluación institucional anual (Guía 34 – MEN)", count: 31},
        {opcion: "Reuniones periódicas con entes de gobierno escolar", count: 28},
        {opcion: "Plan de Mejoramiento Institucional (PMI)", count: 27},
        {opcion: "Elecciones estudiantiles", count: 26},
        {opcion: "Elección de docentes para consejo directivo", count: 25},
        {opcion: "Elección de docentes para consejo académico", count: 23},
        {opcion: "Actualización y ajuste permanente del PEI", count: 23}
      ],
      comun: "La autoevaluación institucional anual conforme a la Guía 34 del MEN es, con 31 de las 34 instituciones (91,2%), el mecanismo de seguimiento a la democracia escolar más adoptado en todo el municipio. Le siguen las reuniones periódicas con los entes de gobierno escolar (28 IE, 82,4%), el Plan de Mejoramiento Institucional (27 IE, 79,4%), las elecciones estudiantiles (26 IE, 76,5%), la elección de docentes para el Consejo Directivo (25 IE, 73,5%) y, en un mismo nivel, la elección de docentes para el Consejo Académico y la actualización permanente del PEI (23 IE cada uno, 67,6%). Este patrón confirma que el seguimiento a la democracia escolar en las 34 instituciones del Foro se sustenta principalmente en los procesos electorales reglamentarios y en los instrumentos de autoevaluación dispuestos por la normativa nacional.",
      particularidades: "Atanasio Girardot (Grupo 4) constituye el hallazgo más crítico de todo el municipio en esta pregunta: es la única de las 34 instituciones en admitir explícitamente que \"no existen mecanismos formales, instrumentos consolidados o un seguimiento riguroso\" a la democracia escolar más allá de la votación inicial de representantes, una confesión de vacío institucional sin equivalente en ninguna otra institución del Foro. El Caguán (Grupo 6), por su parte, complementa este panorama al describir en su respuesta cualitativa —aunque sin reflejarlo en el catálogo de casillas marcadas— mecanismos propios de veeduría comunitaria directa, incluyendo asambleas y pronunciamientos públicos de padres y estudiantes."
    }
  ],
  percepcionQR: {
    totalFirmantes: 286,
    institucionesQR: [
      {nombre: "LICEO DE SANTA LIBRADA", grupo: "G1", firmantes: 69, totalParticipantes: 116},
      {nombre: "ATANASIO GIRARDOT", grupo: "G4", firmantes: 64, totalParticipantes: 87},
      {nombre: "EL LIMONAR", grupo: "G6", firmantes: 101, totalParticipantes: 112},
      {nombre: "RODRIGO LARA BONILLA", grupo: "G6", firmantes: 52, totalParticipantes: 75}
    ],
    demografia: [
      {categoria: "Mujeres adultas (mayores de 18 años)", count: 178},
      {categoria: "Hombres adultos (mayores de 18 años)", count: 80},
      {categoria: "Adolescentes mujeres (13 a 18 años)", count: 17},
      {categoria: "Adolescentes hombres (13 a 18 años)", count: 5},
      {categoria: "Niños y niñas (0 a 12 años)", count: 2},
      {categoria: "Prefirió no responder la edad", count: 3}
    ],
    fortalezasGenerales: [
      {opcion: "Trabajo colaborativo entre docentes", count: 4},
      {opcion: "Docentes con experiencias exitosas que pueden ser compartidas", count: 4},
      {opcion: "Participación activa de los estudiantes", count: 1},
      {opcion: "Liderazgo pedagógico de los directivos", count: 1},
      {opcion: "Reconocimiento de las necesidades del contexto", count: 1},
      {opcion: "Estrategias que han mejorado los aprendizajes", count: 1}
    ],
    oportunidadesGenerales: [
      {opcion: "Baja participación de estudiantes y familias", count: 4},
      {opcion: "Infraestructura o dotación que limita el aprendizaje", count: 3},
      {opcion: "Brechas de aprendizaje entre estudiantes, grados o sedes", count: 2},
      {opcion: "Bajo logro de aprendizajes fundamentales", count: 1},
      {opcion: "Desarticulación entre PEI, currículo y práctica de aula", count: 1},
      {opcion: "Necesidades de formación y acompañamiento docente", count: 1}
    ],
    analisisCuantitativo: "De los 3.196 participantes declarados en las 34 instituciones del Foro Educativo Institucional 2026, 286 personas —cifra prácticamente coincidente con las 287 reportadas— firmaron su asistencia mediante código QR con encuesta de percepción digital, en cuatro instituciones de tres de los seis grupos de trabajo: el Liceo de Santa Librada (Grupo 1, 69 firmantes de 116 participantes), Atanasio Girardot (Grupo 4, 64 de 87), El Limonar (Grupo 6, 101 de 112) y Rodrigo Lara Bonilla (Grupo 6, 52 de 75). Aunque esta muestra cubre solo 4 de las 34 instituciones (11,8%) y en torno al 9% del total de participantes del Foro, constituye la única fuente de datos demográficos y de percepción individual disponible en todo el municipio, y sus hallazgos —como se detalla más adelante— resultan coherentes con la síntesis cualitativa de los seis grupos de trabajo. En términos demográficos, el 90,2% de los 286 firmantes (258 personas) son adultos mayores de 18 años (178 mujeres y 80 hombres), el 7,7% (22 personas) son adolescentes entre 13 y 18 años (17 mujeres y 5 hombres), apenas el 0,7% (2 personas) son niños o niñas entre 0 y 12 años, y un 1,0% (3 personas) prefirió no responder la pregunta de edad. Esta composición demográfica —abrumadoramente adulta y mayoritariamente docente— es coherente con la caracterización general de participantes de todo el Foro, en la que los docentes representan la proporción más alta de asistentes en la práctica totalidad de las 34 instituciones. En cuanto a las fortalezas identificadas, dos opciones fueron señaladas por la totalidad de las cuatro instituciones con datos de percepción (100%): el trabajo colaborativo entre docentes y la existencia de docentes con experiencias exitosas que pueden ser compartidas. Con una adopción minoritaria se ubican la participación activa de los estudiantes, el liderazgo pedagógico de los directivos, el reconocimiento de las necesidades del contexto y las estrategias que han mejorado los aprendizajes, cada una señalada como fortaleza principal por una sola de las cuatro instituciones. En cuanto a las oportunidades de mejoramiento, un hallazgo es igualmente unánime: la baja participación de estudiantes y familias fue señalada por las cuatro instituciones (100%), seguida por la infraestructura o dotación que limita el aprendizaje, señalada por tres de las cuatro (75%). Con una adopción minoritaria se ubican las brechas de aprendizaje entre estudiantes, grados o sedes (2 de 4, 50%), el bajo logro de aprendizajes fundamentales, la desarticulación entre PEI, currículo y práctica de aula, y las necesidades de formación y acompañamiento docente (cada una en 1 de las 4 instituciones).",
    analisisCualitativo: "El análisis por grupo etario, posible en las tres instituciones que registraron respuestas de adolescentes (Liceo de Santa Librada, Atanasio Girardot y El Limonar) y en las dos que registraron respuestas de niños y niñas (Liceo de Santa Librada y El Limonar, con apenas un caso cada una), aporta un matiz relevante a la lectura general. Entre los adultos —el 90% de la muestra— el patrón es enteramente consistente con el hallazgo general: las cuatro instituciones coinciden en valorar el trabajo colaborativo docente y las experiencias exitosas compartidas como sus principales fortalezas, y en señalar la baja participación de familias como su principal reto, acompañada en tres de los cuatro casos por las limitaciones de infraestructura y dotación. Entre los adolescentes, el patrón es más heterogéneo y arroja un hallazgo particularmente relevante: mientras el Liceo de Santa Librada y El Limonar mantienen entre sus jóvenes una valoración positiva del trabajo colaborativo docente, los adolescentes de Atanasio Girardot señalan justamente lo contrario —el \"escaso trabajo colaborativo entre docentes\"— como una de sus principales oportunidades de mejora, una divergencia generacional dentro de una misma institución que sugiere que la percepción positiva de los adultos sobre su propio trabajo en equipo no necesariamente es compartida por sus estudiantes adolescentes, y que amerita ser explorada con mayor profundidad por la propia institución. Los dos registros disponibles de percepción infantil, aunque estadísticamente insuficientes para generalizar, coinciden en valorar el liderazgo pedagógico de los directivos como una fortaleza reconocida incluso por los niños y niñas más pequeños de la comunidad educativa. En el plano cualitativo de las respuestas de texto libre, los dos únicos comentarios registrados en las secciones de oportunidades de mejoramiento (Liceo de Santa Librada y Rodrigo Lara Bonilla) convergen exactamente en el mismo señalamiento: la necesidad de un mayor compromiso y participación de las familias en las actividades institucionales. Esta coincidencia, sumada a que la baja participación de estudiantes y familias fue la oportunidad de mejora unánime entre las cuatro instituciones con datos de percepción digital, y a que este mismo hallazgo emergió de manera independiente y recurrente en el análisis cualitativo de los 34 informes ejecutivos que sustenta las once secciones anteriores de este documento, permite afirmar con un alto grado de confianza que la participación insuficiente de las familias es, con diferencia, el reto más extendido y mejor triangulado de todo el Foro Educativo Institucional Neiva 2026 — una conclusión que converge de manera consistente tanto en el análisis cualitativo de los seis grupos de trabajo como en el análisis cuantitativo de la muestra poblacional de firmantes QR aquí presentada."
  },
  conclusiones: "El Foro Educativo Institucional — Neiva 2026 congregó a las 36 instituciones educativas oficiales del municipio organizadas en seis grupos de trabajo, de las cuales 34 remitieron su Informe Ejecutivo a la fecha de consolidación de este documento (Humberto Tafur Charry, del Grupo 4, y María Cristina Arango de Pastrana, del Grupo 2, no habían enviado su informe). El análisis conjunto de las tres sesiones de trabajo desarrolladas por estas 34 instituciones permite concluir que el municipio comparte una trayectoria sólida de avances en la contextualización curricular mediante proyectos pedagógicos propios y transversales, en el fortalecimiento de la educación inicial —aunque con una implementación de prejardín sistemáticamente rezagada frente a la de jardín, en buena medida por la competencia de los hogares comunitarios del ICBF—, y en la consolidación de una estructura de gobierno escolar notablemente homogénea, con las instancias reglamentarias de la Ley 115 alcanzando niveles de adopción entre el 70% y el 97% en la práctica totalidad de los equipos y mecanismos evaluados. Frente a estos avances, dos retos estructurales atraviesan la totalidad del municipio sin excepción: las limitaciones de infraestructura, dotación y conectividad, señaladas como el obstáculo más persistente en los seis grupos de trabajo; y la participación real de las familias, que —según confirma de manera convergente tanto el análisis cualitativo de los 34 informes ejecutivos como el análisis cuantitativo de los 286 firmantes por código QR presentado en el capítulo especial de este documento— constituye el reto mejor documentado y más extendido de todo el Foro. Cuatro hallazgos institucionales particulares ameritan seguimiento prioritario y específico de la Secretaría de Educación Municipal: la IE Atanasio Girardot (Grupo 4), única institución del municipio en admitir no haber convocado su Consejo Académico durante todo el año escolar ni contar con mecanismos formales de seguimiento a la democracia escolar; la IE El Caguán (Grupo 6), que reportó la inestabilidad de su rectoría como factor que ha afectado la continuidad de sus procesos institucionales, incluyendo un caso documentado de una docente de primera infancia reasignada por falta de infraestructura hasta terminar atendiendo a sus estudiantes en un salón improvisado; la IE Ricardo Borrero Álvarez (Grupo 3), que denunció la centralización de decisiones clave sin comunicación oportuna a las bases, además de problemáticas concretas de salubridad e inseguridad; y la IE Jairo Morera Lizcano (Grupo 5), única institución del Foro en concluir sin matices que su currículo actual no es pertinente con la realidad territorial. En el extremo positivo, la IE Roberto Durán Alvira (Grupo 4) y la IE San Antonio de Anaconia (Grupo 5) se destacan como referentes replicables: la primera por su macroproyecto rural \"Cultivando saberes y conexiones rurales\", con un nivel de avance estimado del 80% frente a las metas del FEM2025; y la segunda por un conjunto de herramientas de participación directa y transparencia (presupuestos participativos escolares, rendición pública de cuentas, mesas de co-creación del PEI) sin equivalente en ninguna otra institución del municipio. En cuanto a la valoración general de la jornada, de las 34 instituciones con informe remitido, únicamente ocho registraron formalmente su calificación del Foro a la fecha de este documento: Fortalecillas (5,0/5, Grupo 2), Departamental Tierra de Promisión (5,0/5, Grupo 3), José Eustasio Rivera (4,5/5, Grupo 4), Escuela Normal Superior (5,0/5, Grupo 5), Instituto Técnico IPC Andrés Rosa (4,8/5, Grupo 5), El Limonar (4,8/5, Grupo 6) y Rodrigo Lara Bonilla (4,8/5, Grupo 6); las 26 instituciones restantes aún no habían diligenciado este instrumento. Se recomienda a la Secretaría de Educación Municipal: (1) gestionar el diligenciamiento pendiente de la valoración general y del Informe Ejecutivo con las instituciones señaladas; (2) diseñar una estrategia diferenciada de fortalecimiento de la participación familiar, dado que este es el hallazgo más robusto y mejor triangulado de todo el Foro; (3) brindar acompañamiento específico a Atanasio Girardot y El Caguán en materia de gobierno escolar y continuidad institucional; y (4) ampliar el uso del instrumento de percepción por código QR a la totalidad de las instituciones en futuras versiones del Foro, dado que la muestra de 286 firmantes analizada en el capítulo especial de este documento, aunque valiosa, cubrió apenas 4 de las 36 instituciones del municipio.",
  proyectoNombre: "[Nombre de quien consolida el Informe Consolidado]",
  proyectoCargo: "[Cargo]",
  fechaRealizacion: "8 de septiembre de 2026"
};

/*
 * Genera (o reescribe, si ya existe) el Informe de Síntesis Municipal
 * consolidado del Foro Educativo Institucional Neiva 2026, dejándolo
 * en la raíz de la carpeta general del Foro (no en una carpeta de
 * grupo específica, ya que este documento cubre a los 6 grupos).
 *
 * Uso: desde el editor de Apps Script, seleccionar esta función y
 * presionar "Ejecutar". El enlace del documento queda en "Ver
 * registros de ejecución".
 */
function generarInformeSintesisMunicipalFEM(){
  const carpetaRaiz=DriveApp.getFolderById(DRIVE_CARPETA_FEM_ID);
  const nombreDoc="Informe de Síntesis Municipal - Foro Educativo Institucional Neiva 2026";
  const existentesIt=carpetaRaiz.getFilesByName(nombreDoc);
  const archivoExistente=existentesIt.hasNext() ? existentesIt.next() : null;
  while(existentesIt.hasNext()) existentesIt.next().setTrashed(true);

  const archivoDoc=generarInformeSintesisMunicipalFEM_(DATOS_SINTESIS_MUNICIPAL_FEM_, archivoExistente?archivoExistente.getId():null);
  if(!archivoExistente){
    carpetaRaiz.addFile(archivoDoc);
    try{ DriveApp.getRootFolder().removeFile(archivoDoc); }catch(e){}
  }

  Logger.log("========================================");
  Logger.log("INFORME DE SÍNTESIS MUNICIPAL — RESULTADO");
  Logger.log("Documento: "+archivoDoc.getUrl());
  Logger.log("========================================");

  return archivoDoc.getUrl();
}

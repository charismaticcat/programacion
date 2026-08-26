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

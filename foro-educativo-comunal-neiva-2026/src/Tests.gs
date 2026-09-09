/**
 * Tests.gs — Foro Educativo Comunal Neiva 2026
 *
 * Pruebas manuales ejecutables desde el editor de Apps Script (metodología
 * de Pruebas.js en FEI 3.1 — docs/01-auditoria-fei-3.1.md §2.7 — sin
 * reutilizar contenido específico de 2026: aquí todo dato de prueba está
 * marcado explícitamente como "PRUEBA", nunca se inventan datos reales de
 * grupos/IE de Neiva).
 *
 * Uso: seleccionar la función en el editor de Apps Script y ejecutar con
 * "Ejecutar". La primera vez, Google pedirá autorizar los permisos.
 */

/** Fase 0: inicializa el spreadsheet + todas las hojas base (equivalente a la primera llamada a getConfig()). */
function testInicializarProyecto() {
  var config = getConfig();
  Logger.log("Spreadsheet del proyecto: " + config.SPREADSHEET_ID);
  Logger.log("Abrir: https://docs.google.com/spreadsheets/d/" + config.SPREADSHEET_ID);
  asegurarEstructuraDriveComunal_();
  Logger.log("Estructura de Drive asegurada. Carpeta raíz: " + getConfig().CARPETA_DRIVE_ID);
}

/**
 * Crea un GRUPO DE PRUEBA con 2 IE ficticias en GruposComunal, claramente
 * marcadas como prueba (nunca usar en producción). Útil para probar el
 * flujo completo sin depender de que la SEM ya haya cargado los grupos
 * reales.
 */
function testCrearGrupoDePrueba() {
  var hoja = obtenerHoja_(HOJA_GRUPOS_COMUNAL_, cabecerasGruposComunal_());
  var idGrupo = "GRUPO-PRUEBA";
  var filas = leerFilasComoObjetos_(hoja);
  var yaExiste = filas.some(function (f) {
    return String(f.ID_GRUPO).trim() === idGrupo;
  });
  if (yaExiste) {
    Logger.log("GRUPO-PRUEBA ya existe, no se duplica.");
    return;
  }
  hoja.appendRow([idGrupo, "Grupo de prueba", "IE-PRUEBA-1", "I.E. Institución de Prueba 1", "Comuna de prueba", "SI"]);
  hoja.appendRow([idGrupo, "Grupo de prueba", "IE-PRUEBA-2", "I.E. Institución de Prueba 2", "Comuna de prueba", "SI"]);
  Logger.log("GRUPO-PRUEBA creado con 2 IE de prueba.");
}

/** Genera accesos (TOKEN+código) para todos los grupos que aún no lo tengan, incluido GRUPO-PRUEBA. */
function testGenerarAccesos() {
  var resultado = generarAccesosGrupo();
  Logger.log(JSON.stringify(resultado));
}

/** Imprime el TOKEN/CODIGO/URL del grupo de prueba, para probar el flujo desde el navegador. */
function testMostrarAccesoDePrueba() {
  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", "GRUPO-PRUEBA");
  if (fila === -1) {
    Logger.log("No hay acceso para GRUPO-PRUEBA todavía. Ejecute testCrearGrupoDePrueba() y luego testGenerarAccesos().");
    return;
  }
  var acceso = leerFilaComoObjeto_(hoja, fila, mapa);
  Logger.log("TOKEN: " + acceso.TOKEN);
  Logger.log("CODIGO_ACCESO: " + acceso.CODIGO_ACCESO);
  Logger.log("URL_ACCESO: " + acceso.URL_ACCESO);
}

/** Simula el flujo completo end-to-end sobre GRUPO-PRUEBA: acceso, participación, Sesión 1, ConectaEduca, informe. */
function testFlujoCompletoGrupoPrueba() {
  testCrearGrupoDePrueba();
  testGenerarAccesos();

  var hoja = obtenerHoja_(HOJA_ACCESOS_GRUPO_, cabecerasAccesosGrupo_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "ID_GRUPO", "GRUPO-PRUEBA");
  var acceso = leerFilaComoObjeto_(hoja, fila, mapa);
  var dispositivoId = "DISPOSITIVO-PRUEBA-1";

  var validacion = validarAccesoGrupo(acceso.TOKEN, acceso.CODIGO_ACCESO, dispositivoId, false);
  Logger.log("validarAccesoGrupo: " + JSON.stringify(validacion));
  if (!validacion.ok) throw new Error("Falló la validación de acceso de prueba.");

  var tokenSesion = validacion.tokenSesion;

  Logger.log(JSON.stringify(registrarParticipante("GRUPO-PRUEBA", "IE-PRUEBA-1", "Persona de Prueba", "Docente", "Participante", "", dispositivoId)));

  Logger.log(JSON.stringify(guardarConsentimientoGrupo("GRUPO-PRUEBA", tokenSesion, dispositivoId)));

  Logger.log(JSON.stringify(guardarResponsableEnvio("GRUPO-PRUEBA", tokenSesion, dispositivoId, "PRINCIPAL", {
    nombre: "Responsable de Prueba", idIE: "IE-PRUEBA-1", rolForo: "Líder (Rector/Rectora)", correo: "prueba@example.org"
  })));

  var guardado = guardarSesion1("GRUPO-PRUEBA", tokenSesion, dispositivoId, {
    REFLEXIONES: "Texto de prueba — reflexiones.",
    CONCLUSIONES: "Texto de prueba — conclusiones.",
    PROPUESTAS_IE: "Texto de prueba — propuestas IE.",
    EXPERIENCIAS: "Texto de prueba — experiencias.",
    RETOS: "Texto de prueba — retos.",
    APORTES_TERRITORIALES: "Texto de prueba — aportes territoriales.",
    CONVERGENCIAS: "Texto de prueba — convergencias.",
    APUESTAS: "Texto de prueba — apuestas.",
    DESAFIOS: "Texto de prueba — desafíos.",
    IDENTIDAD: "Texto de prueba — identidad.",
    PRIORIDADES: "Texto de prueba — prioridades.",
    PROPUESTAS_COLECTIVAS: "Texto de prueba — propuestas colectivas.",
    ACUERDOS: "Texto de prueba — acuerdos.",
    RUTA: "Texto de prueba — ruta de trabajo."
  });
  Logger.log("guardarSesion1: " + JSON.stringify(guardado));

  Logger.log(JSON.stringify(enviarSesion1Definitiva("GRUPO-PRUEBA", tokenSesion, dispositivoId)));

  Logger.log(JSON.stringify(
    guardarActorConectaEduca("GRUPO-PRUEBA", tokenSesion, dispositivoId, {
      ACTOR: "Actor de prueba",
      TIPO_ACTOR: "Educación superior",
      AREA: "Área de prueba",
      OPORTUNIDAD: "Oportunidad de prueba"
    })
  ));

  Logger.log(JSON.stringify(enviarSesion2Definitiva("GRUPO-PRUEBA", tokenSesion, dispositivoId)));

  var informe = generarInformeCompletoGrupo("GRUPO-PRUEBA", tokenSesion, dispositivoId);
  Logger.log("generarInformeCompletoGrupo: " + JSON.stringify(informe));

  liberarSesionGrupo_("GRUPO-PRUEBA", dispositivoId, tokenSesion);
  Logger.log("Flujo de prueba completado.");
}

/** Envío de grupo — recorrido de prueba: una asignación al azar de grupo real por cada dirección de la Alcaldía. */
function testEnviarGrupoRecorridoPrueba() {
  Logger.log(JSON.stringify(enviarGrupoRecorridoPrueba()));
}

/** Mismo recorrido de prueba, un único envío adicional de verificación. */
function testEnviarGrupoRecorridoPruebaAutor() {
  Logger.log(JSON.stringify(enviarGrupoRecorridoPruebaA("jhonefrainsanchez@gmail.com")));
}

/** Borra el GRUPO-PRUEBA y sus datos asociados (Sesión 1, ConectaEduca, participación, acceso, informe). */
function testLimpiarDatosDePrueba() {
  ["GruposComunal", "AccesosGrupo", "ParticipacionComunal", "Sesion1Comunal", "ConectaEduca", "InformesComunal", "EnviosDiferidosComunal", "ResponsablesComunal"].forEach(
    function (nombreHoja) {
      var ss = abrirSpreadsheet_();
      var hoja = ss.getSheetByName(nombreHoja);
      if (!hoja) return;
      var mapa = obtenerMapaCabeceras_(hoja);
      if (!mapa["ID_GRUPO"]) return;
      var filas = buscarFilasPorColumna_(hoja, mapa, "ID_GRUPO", "GRUPO-PRUEBA");
      filas.sort(function (a, b) {
        return b - a;
      }).forEach(function (f) {
        hoja.deleteRow(f);
      });
    }
  );
  Logger.log("Datos de GRUPO-PRUEBA eliminados.");
}

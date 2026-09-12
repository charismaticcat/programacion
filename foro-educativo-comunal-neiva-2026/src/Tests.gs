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

  Logger.log(JSON.stringify(guardarParticipacionEstamentoIE("GRUPO-PRUEBA", tokenSesion, dispositivoId, "IE-PRUEBA-1", {
    RECTOR: 1, COORDINADOR: 1, DOCENTES: 5, TUTOR_PTA: 0, ORIENTADOR: 1, ESTUDIANTES: 20,
    PADRES: 8, ADMINISTRATIVOS: 2, EGRESADOS: 0, SECTOR: 0, OTROS: 0
  })));
  Logger.log("obtenerParticipacionEstamentoGrupo: " + JSON.stringify(obtenerParticipacionEstamentoGrupo("GRUPO-PRUEBA")));

  Logger.log(JSON.stringify(actualizarRectorIE("GRUPO-PRUEBA", tokenSesion, dispositivoId, "IE-PRUEBA-1", "Rector de Prueba")));

  // Sesión de preparación (pre-socialización) — una IE del grupo de prueba diligencia y envía sus aportes.
  Logger.log("obtenerPreparacionIE (antes de guardar): " + JSON.stringify(obtenerPreparacionIE("GRUPO-PRUEBA", "IE-PRUEBA-1")));
  Logger.log(JSON.stringify(guardarPreparacionIE("GRUPO-PRUEBA", tokenSesion, dispositivoId, "IE-PRUEBA-1", "Responsable de Prueba", {
    P1: "Texto de prueba — P1.", P2: "Texto de prueba — P2.", P3: "Texto de prueba — P3.",
    P4: "Texto de prueba — P4.", P5: "Texto de prueba — P5.", P6: "Texto de prueba — P6."
  })));
  Logger.log(JSON.stringify(marcarPreparacionEnviada("GRUPO-PRUEBA", tokenSesion, dispositivoId, "IE-PRUEBA-1")));
  Logger.log("obtenerPreparacionesEnviadasGrupo: " + JSON.stringify(obtenerPreparacionesEnviadasGrupo("GRUPO-PRUEBA")));

  Logger.log(JSON.stringify(guardarResponsableEnvio("GRUPO-PRUEBA", tokenSesion, dispositivoId, "PRINCIPAL", {
    nombre: "Responsable de Prueba", idIE: "IE-PRUEBA-1", rolForo: "Líder (Rector/Rectora)", correo: "prueba@example.org"
  })));

  // Párrafo de prueba de ~58 palabras (Sesion1.gs exige entre 50 y 400 palabras en los campos con rango).
  var TEXTO_PRUEBA_50_PALABRAS_ =
    "Este es un texto de prueba generado únicamente para validar el flujo completo de la aplicación " +
    "del Foro Educativo Comunal Neiva 2026, sin ningún contenido real de ningún grupo o institución " +
    "educativa; sirve exclusivamente para comprobar que el guardado, el envío y la validación de " +
    "longitud mínima y máxima de palabras funcionan correctamente en este entorno de pruebas.";

  var guardado = guardarSesion1("GRUPO-PRUEBA", tokenSesion, dispositivoId, {
    REFLEXIONES: TEXTO_PRUEBA_50_PALABRAS_,
    DESAFIOS: TEXTO_PRUEBA_50_PALABRAS_,
    APUESTAS: TEXTO_PRUEBA_50_PALABRAS_,
    CONCLUSIONES: TEXTO_PRUEBA_50_PALABRAS_,
    PRIORIDADES: "Texto de prueba — prioridades.",
    PROPUESTAS_COLECTIVAS: "Texto de prueba — propuestas colectivas.",
    ACUERDOS: "Texto de prueba — acuerdos.",
    RUTA: "Texto de prueba — ruta de trabajo.",
    NECESIDADES_ARTICULACION_GRUPO: "Texto de prueba — necesidades de articulación.",
    OPORTUNIDADES_GRUPO: "Texto de prueba — oportunidades identificadas.",
    PRIORIDADES_CE: TEXTO_PRUEBA_50_PALABRAS_,
    ACUERDOS_CE: TEXTO_PRUEBA_50_PALABRAS_,
    PROPUESTAS_CE: TEXTO_PRUEBA_50_PALABRAS_,
    RUTA_CE: TEXTO_PRUEBA_50_PALABRAS_
  });
  Logger.log("guardarSesion1: " + JSON.stringify(guardado));

  Logger.log(JSON.stringify(enviarSesion1Definitiva("GRUPO-PRUEBA", tokenSesion, dispositivoId)));

  Logger.log(JSON.stringify(
    guardarActorConectaEduca("GRUPO-PRUEBA", tokenSesion, dispositivoId, {
      ACTOR: "Actor de prueba",
      TIPO_ACTOR: "Educación superior",
      AREA: "Educación y pedagogía",
      IE_INTERESADAS: "I.E. Institución de Prueba 1"
    })
  ));

  Logger.log(JSON.stringify(enviarSesion2Definitiva("GRUPO-PRUEBA", tokenSesion, dispositivoId)));

  // La valoración del Foro es condición para generar el informe (Valoracion.gs) — probar primero que
  // generarInformeCompletoGrupo la exige, luego enviarla y generar el informe de verdad.
  var sinValorar = generarInformeCompletoGrupo("GRUPO-PRUEBA", tokenSesion, dispositivoId);
  Logger.log("generarInformeCompletoGrupo sin valorar (debe fallar): " + JSON.stringify(sinValorar));
  if (sinValorar.ok) throw new Error("generarInformeCompletoGrupo no debería permitir generar el informe sin valoración.");

  Logger.log(JSON.stringify(guardarValoracionGrupo("GRUPO-PRUEBA", tokenSesion, dispositivoId, {
    p1: 5, p2: 4, p3: 5, p4: 4, p5: "Sugerencia de prueba.",
    mejoraP1: "", mejoraP2: "Mejora de prueba.", mejoraP3: "", mejoraP4: ""
  })));

  var informe = generarInformeCompletoGrupo("GRUPO-PRUEBA", tokenSesion, dispositivoId);
  Logger.log("generarInformeCompletoGrupo: " + JSON.stringify(informe));
  if (informe.ok) {
    Logger.log(JSON.stringify(marcarInformeDescargado("GRUPO-PRUEBA", tokenSesion, dispositivoId)));
    Logger.log(JSON.stringify(enviarInformeSiCorresponde("GRUPO-PRUEBA", tokenSesion, dispositivoId)));
  }

  liberarSesionGrupo_("GRUPO-PRUEBA", dispositivoId, tokenSesion);
  Logger.log("Flujo de prueba completado.");
}

/** Importa los logos reales de las 36 IE (con grupo asignado) desde FEI 3.1 — solo lectura, ejecutar una vez. */
function testImportarLogosIE() {
  Logger.log(JSON.stringify(importarLogosIEDesdeFEI31()));
}

/** Envío de grupo — recorrido de prueba: una asignación al azar de grupo real por cada dirección de la Alcaldía. */
function testEnviarGrupoRecorridoPrueba() {
  Logger.log(JSON.stringify(enviarGrupoRecorridoPrueba()));
}

/** Mismo recorrido de prueba, un único envío adicional de verificación. */
function testEnviarGrupoRecorridoPruebaAutor() {
  Logger.log(JSON.stringify(enviarGrupoRecorridoPruebaA("jhonefrainsanchez@gmail.com")));
}

/**
 * Verifica los permisos de envío de correo de la cuenta que ejecuta este
 * proyecto (spec del usuario: "verificar permisos para envío de email
 * desde email que se usó en 3.1 para envíos") — FEI 3.1 enviaba siempre
 * como REMITENTE_FEM = "calidadeducacion@alcaldianeiva.gov.co"
 * (Código.js, misma validación de alias que remitenteValido_ aquí).
 * Ejecutar esta función manualmente desde el editor de Apps Script (Ver →
 * Registros, o Ejecución → Ver registro de ejecución) para confirmar si
 * esta cuenta puede enviar con esa misma identidad antes de configurarla
 * en ConfiguracionComunal (CORREO_REMITENTE).
 */
function testVerificarPermisosCorreo() {
  var CORREO_USADO_EN_3_1 = "calidadeducacion@alcaldianeiva.gov.co";
  var cuenta = Session.getEffectiveUser().getEmail();
  var aliases = GmailApp.getAliases();
  var disponible = cuenta.toLowerCase() === CORREO_USADO_EN_3_1.toLowerCase() ||
    aliases.map(function (a) { return a.toLowerCase(); }).indexOf(CORREO_USADO_EN_3_1.toLowerCase()) !== -1;
  var cuotaRestante = MailApp.getRemainingDailyQuota();
  var resultado = {
    cuentaEjecutando: cuenta,
    aliasesDisponibles: aliases,
    correoUsadoEn3_1: CORREO_USADO_EN_3_1,
    puedeEnviarComoEn3_1: disponible,
    cuotaDiariaRestante: cuotaRestante,
    recomendacion: disponible
      ? "Esta cuenta SÍ puede enviar como " + CORREO_USADO_EN_3_1 + " — configure CORREO_REMITENTE con ese " +
        "valor en ConfiguracionComunal para mantener la misma identidad de envío que FEI 3.1."
      : "Esta cuenta NO tiene " + CORREO_USADO_EN_3_1 + " como alias — deje CORREO_REMITENTE vacío (usará " +
        cuenta + ") o configure/verifique un alias válido para esa cuenta en Gmail antes de usar ese valor."
  };
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}

/**
 * Corrige, en los registros YA guardados de ParticipacionComunal, la
 * columna ROL_FORO vacía (spec del usuario: "cruzar estamento con rol en
 * reporte de asistencia, actualmente aparece rol en blanco") —
 * registrarAsistenciaPublica (Asistencia.gs) ya guarda el rol cruzado con
 * el estamento para cualquier firma nueva; esta función es solo para
 * poner al día las firmas que ya existían antes de ese cambio. Ejecutar
 * una sola vez, manualmente, desde el editor de Apps Script.
 */
function corregirRolForoVacioConEstamento() {
  var hoja = obtenerHoja_(HOJA_PARTICIPACION_COMUNAL_, cabecerasParticipacionComunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { ok: true, corregidos: 0 };
  var rango = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn());
  var valores = rango.getValues();
  var corregidos = 0;
  valores.forEach(function (fila) {
    var estamento = String(fila[mapa["ESTAMENTO"] - 1] || "").trim();
    var rolForo = String(fila[mapa["ROL_FORO"] - 1] || "").trim();
    if (!rolForo && estamento) {
      fila[mapa["ROL_FORO"] - 1] = estamento;
      corregidos++;
    }
  });
  if (corregidos) rango.setValues(valores);
  Logger.log("corregirRolForoVacioConEstamento: " + corregidos + " fila(s) corregida(s).");
  return { ok: true, corregidos: corregidos };
}

/** Borra el GRUPO-PRUEBA y sus datos asociados (Sesión 1, ConectaEduca, participación, acceso, informe). */
function testLimpiarDatosDePrueba() {
  ["GruposComunal", "AccesosGrupo", "ParticipacionComunal", "Sesion1Comunal", "ConectaEduca", "InformesComunal", "EnviosDiferidosComunal", "ResponsablesComunal", "ValoracionComunal", "ParticipacionEstamentoIE", "PreparacionIE", "InvitadosPreparacion"].forEach(
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

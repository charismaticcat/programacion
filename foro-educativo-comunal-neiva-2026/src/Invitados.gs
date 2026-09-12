/**
 * Invitados.gs — Foro Educativo Comunal Neiva 2026
 *
 * Acceso de invitados (estudiantes actuales, egresados(as) o padres/
 * madres/acudientes — "adultos responsables"): NO requiere el código de
 * acceso del grupo (spec: "crear perfil de estudiante y de padre de
 * familia (no necesita código) solo accede dando click en soy
 * invitado"). Eligen su institución educativa (de todas las IE reales,
 * no solo las de un grupo) y su tipo de invitado, y entran DIRECTAMENTE
 * a la Sesión de preparación de esa IE para enviar sus aportes; una vez
 * enviados, no tienen más acceso a la aplicación (spec: "envía aportes y
 * no tiene más acceso").
 *
 * Es una sesión de privilegio mínimo y de un solo uso: el token de
 * invitado NUNCA pasa por sesionActivaPorIdGrupo_ (el mecanismo de
 * sesión completa del grupo, ligado al código de acceso) — solo
 * autoriza, mediante sesionInvitadoValida_(), guardar y enviar SU PROPIO
 * aporte de preparación de la IE concreta que eligieron al entrar. No da
 * acceso a Participación, Sesión 1, Sesión 2, ni a ninguna otra
 * escritura del grupo.
 *
 * El aporte de cada invitado se guarda en AportesInvitadosPreparacion,
 * una hoja PROPIA e independiente de PreparacionIE (la oficial de cada
 * IE, ver Preparacion.gs): cada envío de invitado es su propia fila
 * (clave TOKEN_INVITADO), nunca se fusiona con la respuesta
 * institucional. Esto corrige dos problemas del diseño anterior: (a) un
 * invitado ya no sobrescribe/comparte la fila que la propia IE está
 * construyendo, y (b) las preguntas del invitado ya NO se prellenan con
 * el resumen sugerido del Informe Ejecutivo (spec del usuario: "No se
 * tiene en cuenta el informe de cada IE, es construcción libre") — en su
 * lugar, preguntasPreparacionInvitado_(tipoInvitado, rolEstudiante) da,
 * para cada pregunta (mismo título que la versión oficial, sin modificar
 * su redacción), una explicación con un ejemplo de respuesta afirmativa
 * Y uno de respuesta negativa (para no sugerir una sola "respuesta
 * correcta"), distinta para estudiantes actuales, egresados(as) y
 * adultos responsables — nunca remite al Informe Ejecutivo, que el
 * invitado no tiene por qué conocer.
 *
 * Todo el texto de instrucciones de esta pantalla se redacta en plural
 * ("ustedes"), ya que un mismo dispositivo suele pasar de mano en mano
 * entre varios estudiantes/egresados(as) o adultos responsables que
 * entran uno tras otro (spec del usuario, lote 18).
 */

var HOJA_INVITADOS_ = "InvitadosPreparacion";
var HOJA_APORTES_INVITADOS_ = "AportesInvitadosPreparacion";
var HOJA_DECLARACION_INVITADOS_IE_ = "DeclaracionInvitadosIE";

function cabecerasInvitadosPreparacion_() {
  return ["TOKEN_INVITADO", "ID_GRUPO", "ID_IE", "TIPO_INVITADO", "DISPOSITIVO_ID", "FECHA_INGRESO", "ENVIADO"];
}

function cabecerasAportesInvitadosPreparacion_() {
  return [
    "TOKEN_INVITADO", "ID_GRUPO", "ID_IE", "TIPO_INVITADO",
    // Caracterización (spec del usuario) — nunca se muestra en pantalla
    // durante el Foro, solo se usa para el informe consolidado final
    // (ver obtenerCaracterizacionInvitadosGrupo_ e Informes.gs). EDAD
    // guarda un RANGO de edad (de 10 en 10), no un número puntual.
    // ROL_ESTUDIANTE/ANIOS_ESTUDIANDO (estudiante actual) y
    // ANIO_GRADUACION/PROFESION_ACTUAL_GRADUADO (egresado/a) son propios
    // del perfil estudiante; VINCULO_IE/ROL_IE_ACUDIENTE, del perfil
    // adulto responsable.
    "NOMBRE", "EDAD", "SEXO", "ROL_ESTUDIANTE", "ANIOS_ESTUDIANDO",
    "ANIO_GRADUACION", "PROFESION_ACTUAL_GRADUADO",
    "VINCULO_IE", "ROL_IE_ACUDIENTE",
    "P1", "P2", "P3", "P4", "P5", "P6",
    "ENVIADO", "FECHA_ENVIO", "ULTIMA_ACTUALIZACION"
  ];
}

/**
 * Estimado agregado, declarado por los propios invitados ANTES de
 * registrarse individualmente ("¿cuántos estudiantes/egresados(as) traen
 * de cada institución?") — pantalla "Instituciones que representan",
 * entre "Su grupo" y "Su institución educativa". Es solo un estimado
 * informativo (no bloquea nada ni reemplaza el registro individual de
 * cada quien), guardado por IE con upsert simple.
 */
function cabecerasDeclaracionInvitadosIE_() {
  return ["CLAVE", "ID_GRUPO", "ID_IE", "CANTIDAD_ESTUDIANTES", "CANTIDAD_EGRESADOS", "ULTIMA_ACTUALIZACION"];
}

function _claveDeclaracionInvitadosIE_(idGrupo, idIE) {
  return String(idGrupo || "").trim() + "|" + String(idIE || "").trim();
}

/**
 * Guarda el estimado agregado de hasta 6 instituciones (las del grupo
 * elegido) que un invitado marcó en la pantalla "Instituciones que
 * representan". Nunca escribe una IE que no pertenezca al grupo.
 */
function guardarInstitucionesRepresentadasInvitado(idGrupo, declaraciones) {
  idGrupo = String(idGrupo || "").trim();
  if (!idGrupo) return { ok: false, mensaje: "Falta el grupo." };

  var institucionesGrupo = obtenerInstitucionesDelGrupo(idGrupo);
  var idsValidos = {};
  institucionesGrupo.forEach(function (ie) { idsValidos[ie.idIE] = true; });

  var guardadas = 0;
  (Array.isArray(declaraciones) ? declaraciones : []).forEach(function (d) {
    var idIE = String((d && d.idIE) || "").trim();
    if (!idIE || !idsValidos[idIE]) return;
    var estudiantes = Math.max(0, Math.round(Number(d.cantidadEstudiantes) || 0));
    var egresados = Math.max(0, Math.round(Number(d.cantidadEgresados) || 0));
    conLock_(function () {
      upsertFila_(
        HOJA_DECLARACION_INVITADOS_IE_,
        cabecerasDeclaracionInvitadosIE_(),
        "CLAVE",
        _claveDeclaracionInvitadosIE_(idGrupo, idIE),
        {
          ID_GRUPO: idGrupo,
          ID_IE: idIE,
          CANTIDAD_ESTUDIANTES: estudiantes,
          CANTIDAD_EGRESADOS: egresados,
          ULTIMA_ACTUALIZACION: new Date()
        }
      );
      return null;
    }, 10000);
    guardadas++;
  });

  return { ok: true, guardadas: guardadas };
}

/** El estimado agregado (por IE) ya declarado para un grupo — usado por Informes.gs. */
function obtenerDeclaracionInvitadosGrupo_(idGrupo) {
  idGrupo = String(idGrupo || "").trim();
  var hoja = obtenerHoja_(HOJA_DECLARACION_INVITADOS_IE_, cabecerasDeclaracionInvitadosIE_());
  var filas = leerFilasComoObjetos_(hoja);
  var porIE = {};
  filas.forEach(function (f) {
    if (String(f.ID_GRUPO || "").trim() !== idGrupo) return;
    var idIE = String(f.ID_IE || "").trim();
    var estudiantes = Number(f.CANTIDAD_ESTUDIANTES || 0);
    var egresados = Number(f.CANTIDAD_EGRESADOS || 0);
    if (!estudiantes && !egresados) return;
    porIE[idIE] = { cantidadEstudiantes: estudiantes, cantidadEgresados: egresados };
  });
  return porIE;
}

/**
 * Ayuda de cada pregunta para el perfil "estudiante actual": un ejemplo
 * de respuesta afirmativa y uno de respuesta negativa — ninguno es "la
 * respuesta correcta", solo referencias de cómo se ve una respuesta
 * completa en cada sentido — y la instrucción de escribir "no sabemos"
 * en vez de dejar la pregunta en blanco cuando no se tiene información
 * sobre el tema. Redactado en plural ("ustedes"), ya que varios
 * estudiantes suelen usar el mismo dispositivo, uno tras otro.
 */
var AYUDA_PREPARACION_INVITADO_ESTUDIANTE_ = {
  P1: "Cuenten, con sus propias palabras, qué cambios o mejoras han notado en su institución en los últimos " +
    "años: nuevas actividades, mejoras en la enseñanza, proyectos que se hayan logrado, etc. Ejemplo: “Sí, " +
    "hemos notado que ahora hay más actividades deportivas y que los profesores usan más la tecnología en " +
    "las clases.” También puede ser: “No, no hemos notado cambios grandes, las clases siguen igual que " +
    "antes.” Si no saben sobre este tema, escriban que no saben — no dejen la pregunta en blanco.",
  P2: "Cuenten si conocen o han visto avances en la atención a los niños y niñas más pequeños (jardín, " +
    "prejardín) en su institución. Ejemplo: “Sí, sabemos que abrieron un salón nuevo para los niños " +
    "pequeños este año.” También puede ser: “No, en nuestra institución no hay ni jardín ni prejardín.” Si " +
    "no saben sobre este tema, escríbanlo así — no dejen la pregunta en blanco.",
  P3: "¿Sienten que lo que se enseña en su institución tiene que ver con la vida real de su barrio o " +
    "comunidad? ¿Por qué? Ejemplo: “Sí, porque en la clase de sociales hablamos de los problemas de " +
    "nuestro barrio.” También puede ser: “No, porque las clases no hablan de lo que pasa aquí en la " +
    "comuna.” Si no saben qué responder, díganlo con sus palabras — no dejen la pregunta en blanco.",
  P4: "Mencionen alguna actividad, proyecto o clase que haya conectado lo aprendido en la institución con " +
    "la comunidad o el entorno. Ejemplo: “Sí, hicimos un proyecto de reciclaje con los vecinos del " +
    "barrio.” También puede ser: “No recordamos ninguna actividad así.” Si no saben, escriban que no " +
    "saben — no dejen la pregunta en blanco.",
  P5: "¿Conocen algún grupo, comité o equipo (de padres, estudiantes, profesores) que trabaje por mejorar " +
    "la relación entre la institución y la comunidad? Cuéntennos cuál y qué hace. Ejemplo: “Sí, el " +
    "Consejo Estudiantil organiza actividades con la comunidad.” También puede ser: “No conocemos ningún " +
    "grupo así en nuestra institución.” Si no saben, díganlo — no dejen la pregunta en blanco.",
  P6: "¿Sienten que en su institución las decisiones se toman escuchando a estudiantes y familias? Cuenten " +
    "un ejemplo. Ejemplo: “Sí, porque elegimos al personero y al Consejo Estudiantil entre todos.” También " +
    "puede ser: “No, porque no hay elecciones de personero ni se nos pregunta nuestra opinión.” Si no " +
    "saben, escríbanlo así — no dejen la pregunta en blanco."
};

/**
 * Igual que AYUDA_PREPARACION_INVITADO_ESTUDIANTE_, para el perfil
 * "egresado(a)" — mismas preguntas y ejemplos, pero aclarando que,
 * aunque ya no estudien allí, también pueden responder desde el
 * recuerdo de su época como estudiantes (spec del usuario: agregar una
 * opción que denote "ya no estudio en la IE, pero recuerdo algo",
 * ejemplo "Recuerdo que…").
 */
var AYUDA_PREPARACION_INVITADO_GRADUADO_ = (function () {
  var NOTA_RECUERDO_ =
    " Si ya no estudian en la institución pero recuerdan algo relacionado con esta pregunta de su época " +
    "como estudiantes, también pueden responder así: “Recuerdo que…” y contar lo que recuerdan.";
  var resultado = {};
  Object.keys(AYUDA_PREPARACION_INVITADO_ESTUDIANTE_).forEach(function (clave) {
    resultado[clave] = AYUDA_PREPARACION_INVITADO_ESTUDIANTE_[clave] + NOTA_RECUERDO_;
  });
  return resultado;
})();

/** Igual que las anteriores, para el perfil "adultos responsables de un(a) estudiante" (plural, "ustedes"). */
var AYUDA_PREPARACION_INVITADO_ADULTO_ = {
  P1: "Cuenten, como adultos responsables, qué cambios o mejoras han notado en la institución educativa " +
    "en los últimos años: nuevas actividades, mejoras en la comunicación con las familias, proyectos que " +
    "se hayan logrado, etc. Ejemplo: “Sí, hemos notado que ahora nos informan más seguido sobre las " +
    "actividades del colegio.” También puede ser: “No, no hemos notado cambios importantes.” Si no saben " +
    "sobre este tema, escriban que no saben — no dejen la pregunta en blanco.",
  P2: "Cuenten si conocen o han visto avances en la atención a los niños y niñas más pequeños (jardín, " +
    "prejardín) en la institución. Ejemplo: “Sí, sabemos que este año abrieron un grado nuevo para los " +
    "más pequeños.” También puede ser: “No, en esta institución no hay jardín ni prejardín.” Si no saben " +
    "sobre este tema, escríbanlo así — no dejen la pregunta en blanco.",
  P3: "¿Sienten que lo que se enseña en la institución tiene que ver con la vida real del barrio o la " +
    "comunidad? ¿Por qué? Ejemplo: “Sí, porque los proyectos que hacen se relacionan con lo que vivimos " +
    "en el barrio.” También puede ser: “No, sentimos que las clases no tienen relación con nuestra " +
    "realidad.” Si no saben qué responder, díganlo con sus palabras — no dejen la pregunta en blanco.",
  P4: "Mencionen alguna actividad, proyecto o iniciativa que haya conectado lo que se enseña en la " +
    "institución con la comunidad o las familias. Ejemplo: “Sí, participamos en una jornada de aseo del " +
    "barrio organizada por el colegio.” También puede ser: “No conocemos ninguna actividad así.” Si no " +
    "saben, escríbanlo — no dejen la pregunta en blanco.",
  P5: "¿Conocen algún grupo, comité o equipo (de padres, estudiantes, profesores) que trabaje por mejorar " +
    "la relación entre la institución y la comunidad? Cuéntennos cuál y qué hace. Ejemplo: “Sí, el " +
    "Consejo de Padres organiza reuniones para hablar de las necesidades del colegio.” También puede ser: " +
    "“No conocemos ningún grupo así.” Si no saben, díganlo — no dejen la pregunta en blanco.",
  P6: "¿Sienten que en la institución las decisiones se toman escuchando a las familias y estudiantes? " +
    "Cuenten un ejemplo. Ejemplo: “Sí, porque nos consultan en las reuniones del Consejo de Padres.” " +
    "También puede ser: “No, sentimos que las decisiones se toman sin consultarnos.” Si no saben, " +
    "escríbanlo así — no dejen la pregunta en blanco."
};

/**
 * Las 6 preguntas de preparación para el invitado, con el MISMO título
 * que PREGUNTAS_PREPARACION_ (Preparacion.gs — nunca se modifica su
 * redacción) pero con la ayuda propia de su perfil: estudiante actual,
 * egresado(a) o adulto responsable.
 */
function preguntasPreparacionInvitado_(tipoInvitado, rolEstudiante) {
  var tipo = String(tipoInvitado || "").toUpperCase();
  var esGraduado = String(rolEstudiante || "").toUpperCase() === "GRADUADO";
  var ayudas = tipo === "ESTUDIANTE"
    ? (esGraduado ? AYUDA_PREPARACION_INVITADO_GRADUADO_ : AYUDA_PREPARACION_INVITADO_ESTUDIANTE_)
    : AYUDA_PREPARACION_INVITADO_ADULTO_;
  return PREGUNTAS_PREPARACION_.map(function (p) {
    return { clave: p.clave, titulo: p.titulo, ayuda: ayudas[p.clave] || "" };
  });
}

/**
 * Inicia una sesión de invitado para una IE concreta — sin código de
 * acceso. `tipoInvitado` es "ESTUDIANTE" (incluye estudiante actual Y
 * egresado(a) — se distinguen entre sí por ROL_ESTUDIANTE, elegido
 * desde el primer paso del flujo) o "ACUDIENTE". El grupo al que
 * pertenece la IE se resuelve automáticamente (obtenerGrupoDeInstitucion_,
 * Instituciones.gs), igual que ya hace el resto de la app.
 */
function iniciarAccesoInvitado(idIE, tipoInvitado, dispositivoId) {
  idIE = String(idIE || "").trim();
  tipoInvitado = String(tipoInvitado || "").trim().toUpperCase();
  if (tipoInvitado !== "ESTUDIANTE" && tipoInvitado !== "ACUDIENTE") {
    return { ok: false, mensaje: "Seleccionen quiénes son." };
  }
  if (!idIE) return { ok: false, mensaje: "Seleccionen su institución educativa." };

  var grupoInfo = obtenerGrupoDeInstitucion_(idIE);
  if (!grupoInfo) return { ok: false, mensaje: "No se encontró esa institución educativa." };

  var instituciones = obtenerInstitucionesDelGrupo(grupoInfo.idGrupo);
  var ie = instituciones.find(function (i) { return i.idIE === idIE; });
  if (!ie) return { ok: false, mensaje: "Esa institución no está activa en ningún grupo todavía." };

  var tokenInvitado = Utilities.getUuid();
  var hoja = obtenerHoja_(HOJA_INVITADOS_, cabecerasInvitadosPreparacion_());
  hoja.appendRow([
    tokenInvitado, grupoInfo.idGrupo, idIE, tipoInvitado, String(dispositivoId || "").trim(), new Date(), "NO"
  ]);

  return {
    ok: true,
    tokenInvitado: tokenInvitado,
    idGrupo: grupoInfo.idGrupo,
    grupo: grupoInfo.grupo,
    idIE: idIE,
    institucion: ie.institucion,
    logoId: ie.logoId || ""
  };
}

/** Verifica que el token de invitado corresponda a esa IE y a ese dispositivo. Devuelve la fila o null. */
function sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId) {
  var hoja = obtenerHoja_(HOJA_INVITADOS_, cabecerasInvitadosPreparacion_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "TOKEN_INVITADO", String(tokenInvitado || "").trim());
  if (fila === -1) return null;
  var obj = leerFilaComoObjeto_(hoja, fila, mapa);
  if (String(obj.ID_IE || "").trim() !== String(idIE || "").trim()) return null;
  if (String(obj.DISPOSITIVO_ID || "").trim() !== String(dispositivoId || "").trim()) return null;
  return obj;
}

/**
 * Guarda los datos de caracterización de un invitado (nombre, rango de
 * edad, sexo y, según el perfil: años estudiando en la IE si es
 * estudiante actual; año de graduación y profesión actual si es
 * egresado(a); o vínculo y rol dentro de la IE si es adulto responsable)
 * — en SU PROPIA fila de AportesInvitadosPreparacion (misma clave
 * TOKEN_INVITADO que sus respuestas P1-P6, puede crearse antes de que
 * existan). Estos datos NUNCA se muestran en pantalla durante el Foro
 * (spec del usuario) — solo se usan al final, en el informe consolidado
 * del grupo (ver obtenerCaracterizacionInvitadosGrupo_).
 */
function guardarCaracterizacionInvitado(tokenInvitado, idIE, dispositivoId, datos) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión de invitado ya no es válida." };

  datos = datos || {};
  var nombre = String(datos.nombre || "").trim();
  if (!nombre) return { ok: false, mensaje: "Ingresen el nombre completo." };

  var tipoInvitado = String(sesion.TIPO_INVITADO || "").toUpperCase();
  var campos = {
    ID_GRUPO: sesion.ID_GRUPO,
    ID_IE: idIE,
    TIPO_INVITADO: tipoInvitado,
    NOMBRE: nombre,
    EDAD: String(datos.rangoEdad || "").trim(),
    SEXO: String(datos.sexo || "").trim(),
    ULTIMA_ACTUALIZACION: new Date()
  };
  if (tipoInvitado === "ESTUDIANTE") {
    var rolEstudiante = String(datos.rolEstudiante || "").trim().toUpperCase();
    campos.ROL_ESTUDIANTE = rolEstudiante;
    if (rolEstudiante === "GRADUADO") {
      campos.ANIO_GRADUACION = String(datos.anioGraduacion || "").trim();
      campos.PROFESION_ACTUAL_GRADUADO = String(datos.profesionActual || "").trim();
      campos.ANIOS_ESTUDIANDO = "";
    } else {
      campos.ANIOS_ESTUDIANDO = String(datos.aniosEstudiando || "").trim();
      campos.ANIO_GRADUACION = "";
      campos.PROFESION_ACTUAL_GRADUADO = "";
    }
  } else {
    campos.VINCULO_IE = String(datos.vinculoIE || "").trim();
    campos.ROL_IE_ACUDIENTE = String(datos.rolIEAcudiente || "").trim();
  }

  return conLock_(function () {
    upsertFila_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_(), "TOKEN_INVITADO", String(tokenInvitado || "").trim(), campos);
    return { ok: true };
  }, 10000);
}

/**
 * Preparación del invitado, lista para pintar en pantalla: preguntas en
 * blanco por defecto (construcción libre, NUNCA se prellenan con el
 * resumen sugerido del Informe Ejecutivo) salvo que el propio invitado ya
 * haya guardado un avance con este mismo token, caso en el que se
 * recupera SU propio borrador. La ayuda de cada pregunta depende también
 * de ROL_ESTUDIANTE (estudiante actual vs. egresado/a), ya guardado por
 * guardarCaracterizacionInvitado antes de llegar aquí.
 */
function obtenerPreparacionIEInvitado(tokenInvitado, idIE, dispositivoId) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return null;

  var hoja = obtenerHoja_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "TOKEN_INVITADO", String(tokenInvitado || "").trim());
  var guardado = fila === -1 ? null : leerFilaComoObjeto_(hoja, fila, mapa);

  var preguntas = preguntasPreparacionInvitado_(sesion.TIPO_INVITADO, guardado ? guardado.ROL_ESTUDIANTE : "");
  var respuestas = {};
  preguntas.forEach(function (p) {
    respuestas[p.clave] = guardado ? String(guardado[p.clave] || "") : "";
  });

  return {
    preguntas: preguntas,
    respuestas: respuestas,
    enviado: guardado ? String(guardado.ENVIADO || "") === "SI" : false
  };
}

/**
 * Guarda las respuestas del invitado — en SU PROPIA fila de
 * AportesInvitadosPreparacion (clave TOKEN_INVITADO), independiente de la
 * fila oficial de la IE en PreparacionIE.
 */
function guardarPreparacionIEInvitado(tokenInvitado, idIE, tipoInvitado, dispositivoId, respuestas) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión de invitado ya no es válida." };

  var datos = {
    ID_GRUPO: sesion.ID_GRUPO,
    ID_IE: idIE,
    TIPO_INVITADO: String(sesion.TIPO_INVITADO || tipoInvitado || "").toUpperCase(),
    ULTIMA_ACTUALIZACION: new Date()
  };
  PREGUNTAS_PREPARACION_.forEach(function (p) {
    datos[p.clave] = String((respuestas || {})[p.clave] || "");
  });

  return conLock_(function () {
    upsertFila_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_(), "TOKEN_INVITADO", String(tokenInvitado || "").trim(), datos);
    return { ok: true };
  }, 10000);
}

/**
 * Envío definitivo de un invitado: marca su propio aporte como enviado
 * (AportesInvitadosPreparacion) y cierra su sesión de invitado (ENVIADO=SI
 * en InvitadosPreparacion) — a partir de aquí sesionInvitadoValida_ lo
 * sigue reconociendo (para poder mostrarle la confirmación), pero
 * cualquier intento de volver a guardar o enviar debe rechazarse: ya
 * cumplió su única tarea. Exige TODAS las preguntas respondidas (spec del
 * usuario: "se deben responder todas las preguntas obligatoriamente"),
 * no solo una como antes.
 */
function marcarPreparacionEnviadaInvitado(tokenInvitado, idIE, dispositivoId) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión de invitado ya no es válida." };
  if (String(sesion.ENVIADO || "") === "SI") {
    return { ok: false, mensaje: "Ya enviaron sus aportes con esta sesión de invitado." };
  }

  var hoja = obtenerHoja_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "TOKEN_INVITADO", String(tokenInvitado || "").trim());
  var tieneTodas = false;
  if (fila !== -1) {
    var guardado = leerFilaComoObjeto_(hoja, fila, mapa);
    tieneTodas = PREGUNTAS_PREPARACION_.every(function (p) {
      return String(guardado[p.clave] || "").trim() !== "";
    });
  }
  if (!tieneTodas) {
    return { ok: false, mensaje: "Respondan todas las preguntas antes de enviar." };
  }

  return conLock_(function () {
    upsertFila_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_(), "TOKEN_INVITADO", String(tokenInvitado || "").trim(), {
      ENVIADO: "SI",
      FECHA_ENVIO: new Date()
    });
    upsertFila_(HOJA_INVITADOS_, cabecerasInvitadosPreparacion_(), "TOKEN_INVITADO", String(tokenInvitado || "").trim(), {
      ENVIADO: "SI"
    });
    return { ok: true };
  }, 10000);
}

/**
 * Todos los aportes de invitados ya enviados en el grupo, agrupados por
 * IE — para la sección de invitados al final de la pantalla de
 * selección de IE de la Sesión de preparación (spec: "los aportes de
 * estudiantes deben aparecer con el mismo formato que las IE"; ya NO se
 * muestra dentro de la pantalla de cada IE individual, solo aquí, en el
 * listado agregado). Mismo formato de datos que
 * obtenerPreparacionesEnviadasGrupo (Preparacion.gs) — un párrafo con
 * título por cada pregunta diligenciada — para poder reutilizar el mismo
 * estilo visual que los aportes institucionales.
 */
function obtenerAportesInvitadosGrupo(idGrupo) {
  idGrupo = String(idGrupo || "").trim();
  var hoja = obtenerHoja_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_());
  var filas = leerFilasComoObjetos_(hoja);
  var porIE = {};
  filas.forEach(function (f) {
    if (String(f.ID_GRUPO || "").trim() !== idGrupo) return;
    if (String(f.ENVIADO || "") !== "SI") return;
    var secciones = [];
    PREGUNTAS_PREPARACION_.forEach(function (p) {
      var texto = String(f[p.clave] || "").trim();
      if (texto) secciones.push({ clave: p.clave, titulo: p.titulo, texto: texto });
    });
    if (!secciones.length) return;
    var idIE = String(f.ID_IE || "").trim();
    if (!porIE[idIE]) porIE[idIE] = [];
    porIE[idIE].push({ tipoInvitado: String(f.TIPO_INVITADO || "").toUpperCase(), secciones: secciones });
  });

  return obtenerInstitucionesDelGrupo(idGrupo)
    .filter(function (ie) { return !!porIE[ie.idIE]; })
    .map(function (ie) {
      // Estudiantes primero, luego adultos responsables — orden estable de lectura.
      var aportes = porIE[ie.idIE].sort(function (a, b) {
        if (a.tipoInvitado === b.tipoInvitado) return 0;
        return a.tipoInvitado === "ESTUDIANTE" ? -1 : 1;
      });
      return { idIE: ie.idIE, institucion: ie.institucion, aportes: aportes };
    });
}

/**
 * Caracterización de todos los invitados (estudiantes actuales,
 * egresados(as) y adultos responsables) que ya enviaron sus aportes en
 * el grupo — SOLO para el informe consolidado final (Informes.gs), spec
 * del usuario: "los nombres de los estudiantes y padres de familia
 * aparecerán únicamente al final del foro en un informe consolidado". En
 * ningún otro lugar de la aplicación se muestra el nombre de un
 * invitado.
 */
function obtenerCaracterizacionInvitadosGrupo_(idGrupo) {
  idGrupo = String(idGrupo || "").trim();
  var hoja = obtenerHoja_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_());
  var filas = leerFilasComoObjetos_(hoja);
  var institucionesPorId = {};
  obtenerInstitucionesDelGrupo(idGrupo).forEach(function (ie) { institucionesPorId[ie.idIE] = ie.institucion; });

  var personas = [];
  var totalEstudiantes = 0;
  var totalGraduados = 0;
  var totalAdultos = 0;
  filas.forEach(function (f) {
    if (String(f.ID_GRUPO || "").trim() !== idGrupo) return;
    if (String(f.ENVIADO || "") !== "SI") return;
    var tipo = String(f.TIPO_INVITADO || "").toUpperCase();
    var persona = {
      nombre: String(f.NOMBRE || "").trim() || "(sin nombre registrado)",
      institucion: institucionesPorId[String(f.ID_IE || "").trim()] || "",
      rangoEdad: String(f.EDAD || "").trim(),
      sexo: String(f.SEXO || "").trim(),
      tipoInvitado: tipo
    };
    if (tipo === "ESTUDIANTE") {
      var rolEstudiante = String(f.ROL_ESTUDIANTE || "").toUpperCase();
      persona.rol = rolEstudiante === "GRADUADO" ? "Graduado(a)" : "Estudiante";
      if (rolEstudiante === "GRADUADO") {
        persona.anioGraduacion = String(f.ANIO_GRADUACION || "").trim();
        persona.profesionActual = String(f.PROFESION_ACTUAL_GRADUADO || "").trim();
        totalGraduados++;
      } else {
        persona.aniosEstudiando = String(f.ANIOS_ESTUDIANDO || "").trim();
        totalEstudiantes++;
      }
    } else {
      persona.vinculoIE = String(f.VINCULO_IE || "").trim();
      persona.rolIE = String(f.ROL_IE_ACUDIENTE || "").trim();
      totalAdultos++;
    }
    personas.push(persona);
  });

  return {
    personas: personas,
    totalEstudiantes: totalEstudiantes,
    totalGraduados: totalGraduados,
    totalAdultos: totalAdultos
  };
}

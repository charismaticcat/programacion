/**
 * Invitados.gs — Foro Educativo Comunal Neiva 2026
 *
 * Acceso de invitado (estudiante o padre/madre de familia — acudiente):
 * NO requiere el código de acceso del grupo (spec: "crear perfil de
 * estudiante y de padre de familia (no necesita código) solo accede
 * dando click en soy invitado"). Elige su institución educativa (de
 * todas las IE reales, no solo las de un grupo) y su tipo de invitado, y
 * entra DIRECTAMENTE a la Sesión de preparación de esa IE para enviar
 * sus aportes; una vez enviados, no tiene más acceso a la aplicación
 * (spec: "envía aportes y no tiene más acceso").
 *
 * Es una sesión de privilegio mínimo y de un solo uso: el token de
 * invitado NUNCA pasa por sesionActivaPorIdGrupo_ (el mecanismo de
 * sesión completa del grupo, ligado al código de acceso) — solo
 * autoriza, mediante sesionInvitadoValida_(), guardar y enviar SU PROPIO
 * aporte de preparación de la IE concreta que el invitado eligió al
 * entrar. No da acceso a Participación, Sesión 1, Sesión 2, ni a ninguna
 * otra escritura del grupo.
 *
 * El aporte del invitado se guarda en AportesInvitadosPreparacion, una
 * hoja PROPIA e independiente de PreparacionIE (la oficial de cada IE,
 * ver Preparacion.gs): cada envío de invitado es su propia fila (clave
 * TOKEN_INVITADO), nunca se fusiona con la respuesta institucional. Esto
 * corrige dos problemas del diseño anterior: (a) un invitado ya no
 * sobrescribe/comparte la fila que la propia IE está construyendo, y (b)
 * las preguntas del invitado ya NO se prellenan con el resumen sugerido
 * del Informe Ejecutivo (spec del usuario: "No se tiene en cuenta el
 * informe de cada IE, es construcción libre") — en su lugar,
 * preguntasPreparacionInvitado_(tipoInvitado) da, para cada pregunta
 * (mismo título que la versión oficial, sin modificar su redacción), una
 * explicación con un ejemplo de respuesta afirmativa Y uno de respuesta
 * negativa (para no sugerir una sola "respuesta correcta"), distinta para
 * el perfil "estudiante o egresado(a)" y para "adulto responsable de
 * un(a) estudiante" — nunca remite al Informe Ejecutivo, que el invitado
 * no tiene por qué conocer.
 */

var HOJA_INVITADOS_ = "InvitadosPreparacion";
var HOJA_APORTES_INVITADOS_ = "AportesInvitadosPreparacion";

function cabecerasInvitadosPreparacion_() {
  return ["TOKEN_INVITADO", "ID_GRUPO", "ID_IE", "TIPO_INVITADO", "DISPOSITIVO_ID", "FECHA_INGRESO", "ENVIADO"];
}

function cabecerasAportesInvitadosPreparacion_() {
  return [
    "TOKEN_INVITADO", "ID_GRUPO", "ID_IE", "TIPO_INVITADO",
    "P1", "P2", "P3", "P4", "P5", "P6",
    "ENVIADO", "FECHA_ENVIO", "ULTIMA_ACTUALIZACION"
  ];
}

/**
 * Ayuda de cada pregunta para el perfil "estudiante o egresado(a)": un
 * ejemplo de respuesta afirmativa y uno de respuesta negativa — ninguno
 * es "la respuesta correcta", solo referencias de cómo se ve una
 * respuesta completa en cada sentido — y la instrucción de escribir "no
 * sé" en vez de dejar la pregunta en blanco cuando no se tiene
 * información sobre el tema.
 */
var AYUDA_PREPARACION_INVITADO_ESTUDIANTE_ = {
  P1: "Cuenta, con tus propias palabras, qué cambios o mejoras has notado en tu institución en los últimos " +
    "años: nuevas actividades, mejoras en la enseñanza, proyectos que se hayan logrado, etc. Ejemplo: “Sí, " +
    "he notado que ahora hay más actividades deportivas y que los profesores usan más la tecnología en las " +
    "clases.” También puede ser: “No, no he notado cambios grandes, las clases siguen igual que antes.” Si " +
    "no sabes sobre este tema, escribe que no sabes — no dejes la pregunta en blanco.",
  P2: "Cuenta si conoces o has visto avances en la atención a los niños y niñas más pequeños (jardín, " +
    "prejardín) en tu institución. Ejemplo: “Sí, sé que abrieron un salón nuevo para los niños pequeños " +
    "este año.” También puede ser: “No, en mi institución no hay ni jardín ni prejardín.” Si no sabes sobre " +
    "este tema, escríbelo así — no dejes la pregunta en blanco.",
  P3: "¿Sientes que lo que se enseña en tu institución tiene que ver con la vida real de tu barrio o " +
    "comunidad? ¿Por qué? Ejemplo: “Sí, porque en la clase de sociales hablamos de los problemas de " +
    "nuestro barrio.” También puede ser: “No, porque las clases no hablan de lo que pasa aquí en la " +
    "comuna.” Si no sabes qué responder, dilo con tus palabras — no dejes la pregunta en blanco.",
  P4: "Menciona alguna actividad, proyecto o clase que haya conectado lo aprendido en la institución con " +
    "la comunidad o el entorno. Ejemplo: “Sí, hicimos un proyecto de reciclaje con los vecinos del barrio.” " +
    "También puede ser: “No recuerdo ninguna actividad así.” Si no sabes, escribe que no sabes — no dejes " +
    "la pregunta en blanco.",
  P5: "¿Conoces algún grupo, comité o equipo (de padres, estudiantes, profesores) que trabaje por mejorar " +
    "la relación entre la institución y la comunidad? Cuéntanos cuál y qué hace. Ejemplo: “Sí, el Consejo " +
    "Estudiantil organiza actividades con la comunidad.” También puede ser: “No conozco ningún grupo así " +
    "en mi institución.” Si no sabes, dilo — no dejes la pregunta en blanco.",
  P6: "¿Sientes que en tu institución las decisiones se toman escuchando a estudiantes y familias? Cuenta " +
    "un ejemplo. Ejemplo: “Sí, porque elegimos al personero y al Consejo Estudiantil entre todos.” También " +
    "puede ser: “No, porque no hay elecciones de personero ni se nos pregunta nuestra opinión.” Si no " +
    "sabes, escríbelo así — no dejes la pregunta en blanco."
};

/** Igual que AYUDA_PREPARACION_INVITADO_ESTUDIANTE_, para el perfil "adulto responsable de un(a) estudiante". */
var AYUDA_PREPARACION_INVITADO_ADULTO_ = {
  P1: "Cuente, como adulto responsable, qué cambios o mejoras ha notado en la institución educativa en los " +
    "últimos años: nuevas actividades, mejoras en la comunicación con las familias, proyectos que se hayan " +
    "logrado, etc. Ejemplo: “Sí, he notado que ahora nos informan más seguido sobre las actividades del " +
    "colegio.” También puede ser: “No, no he notado cambios importantes.” Si no sabe sobre este tema, " +
    "escriba que no sabe — no deje la pregunta en blanco.",
  P2: "Cuente si conoce o ha visto avances en la atención a los niños y niñas más pequeños (jardín, " +
    "prejardín) en la institución. Ejemplo: “Sí, sé que este año abrieron un grado nuevo para los más " +
    "pequeños.” También puede ser: “No, en esta institución no hay jardín ni prejardín.” Si no sabe sobre " +
    "este tema, escríbalo así — no deje la pregunta en blanco.",
  P3: "¿Siente que lo que se enseña en la institución tiene que ver con la vida real del barrio o la " +
    "comunidad? ¿Por qué? Ejemplo: “Sí, porque los proyectos que hacen se relacionan con lo que vivimos en " +
    "el barrio.” También puede ser: “No, siento que las clases no tienen relación con nuestra realidad.” Si " +
    "no sabe qué responder, dígalo con sus palabras — no deje la pregunta en blanco.",
  P4: "Mencione alguna actividad, proyecto o iniciativa que haya conectado lo que se enseña en la " +
    "institución con la comunidad o las familias. Ejemplo: “Sí, participamos en una jornada de aseo del " +
    "barrio organizada por el colegio.” También puede ser: “No conozco ninguna actividad así.” Si no sabe, " +
    "escríbalo — no deje la pregunta en blanco.",
  P5: "¿Conoce algún grupo, comité o equipo (de padres, estudiantes, profesores) que trabaje por mejorar la " +
    "relación entre la institución y la comunidad? Cuéntenos cuál y qué hace. Ejemplo: “Sí, el Consejo de " +
    "Padres organiza reuniones para hablar de las necesidades del colegio.” También puede ser: “No conozco " +
    "ningún grupo así.” Si no sabe, dígalo — no deje la pregunta en blanco.",
  P6: "¿Siente que en la institución las decisiones se toman escuchando a las familias y estudiantes? " +
    "Cuente un ejemplo. Ejemplo: “Sí, porque nos consultan en las reuniones del Consejo de Padres.” También " +
    "puede ser: “No, siento que las decisiones se toman sin consultarnos.” Si no sabe, escríbalo así — no " +
    "deje la pregunta en blanco."
};

/**
 * Las 6 preguntas de preparación para el invitado, con el MISMO título
 * que PREGUNTAS_PREPARACION_ (Preparacion.gs — nunca se modifica su
 * redacción) pero con la ayuda propia de su perfil.
 */
function preguntasPreparacionInvitado_(tipoInvitado) {
  var ayudas = String(tipoInvitado || "").toUpperCase() === "ESTUDIANTE"
    ? AYUDA_PREPARACION_INVITADO_ESTUDIANTE_
    : AYUDA_PREPARACION_INVITADO_ADULTO_;
  return PREGUNTAS_PREPARACION_.map(function (p) {
    return { clave: p.clave, titulo: p.titulo, ayuda: ayudas[p.clave] || "" };
  });
}

/**
 * Inicia una sesión de invitado para una IE concreta — sin código de
 * acceso. `tipoInvitado` es "ESTUDIANTE" o "ACUDIENTE". El grupo al que
 * pertenece la IE se resuelve automáticamente (obtenerGrupoDeInstitucion_,
 * Instituciones.gs), igual que ya hace el resto de la app.
 */
function iniciarAccesoInvitado(idIE, tipoInvitado, dispositivoId) {
  idIE = String(idIE || "").trim();
  tipoInvitado = String(tipoInvitado || "").trim().toUpperCase();
  if (tipoInvitado !== "ESTUDIANTE" && tipoInvitado !== "ACUDIENTE") {
    return { ok: false, mensaje: "Selecciona si eres estudiante o acudiente." };
  }
  if (!idIE) return { ok: false, mensaje: "Selecciona tu institución educativa." };

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
 * Preparación del invitado, lista para pintar en pantalla: preguntas en
 * blanco por defecto (construcción libre, NUNCA se prellenan con el
 * resumen sugerido del Informe Ejecutivo) salvo que el propio invitado ya
 * haya guardado un avance con este mismo token, caso en el que se
 * recupera SU propio borrador.
 */
function obtenerPreparacionIEInvitado(tokenInvitado, idIE, dispositivoId) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return null;

  var hoja = obtenerHoja_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "TOKEN_INVITADO", String(tokenInvitado || "").trim());
  var guardado = fila === -1 ? null : leerFilaComoObjeto_(hoja, fila, mapa);

  var preguntas = preguntasPreparacionInvitado_(sesion.TIPO_INVITADO);
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
 * cumplió su única tarea.
 */
function marcarPreparacionEnviadaInvitado(tokenInvitado, idIE, dispositivoId) {
  var sesion = sesionInvitadoValida_(tokenInvitado, idIE, dispositivoId);
  if (!sesion) return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Esta sesión de invitado ya no es válida." };
  if (String(sesion.ENVIADO || "") === "SI") {
    return { ok: false, mensaje: "Ya enviaste tus aportes con esta sesión de invitado." };
  }

  var hoja = obtenerHoja_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var fila = buscarFilaPorColumna_(hoja, mapa, "TOKEN_INVITADO", String(tokenInvitado || "").trim());
  var tieneAlgo = false;
  if (fila !== -1) {
    var guardado = leerFilaComoObjeto_(hoja, fila, mapa);
    tieneAlgo = PREGUNTAS_PREPARACION_.some(function (p) {
      return String(guardado[p.clave] || "").trim() !== "";
    });
  }
  if (!tieneAlgo) {
    return { ok: false, mensaje: "Responde al menos una pregunta antes de enviar." };
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

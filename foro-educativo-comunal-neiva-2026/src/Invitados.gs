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
 * PREGUNTAS_PREPARACION_INVITADO_ da, para cada pregunta (mismo título
 * que la versión oficial, sin modificar su redacción), una explicación y
 * un ejemplo pensados para alguien sin acceso al Informe Ejecutivo.
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
 * Mismas 6 preguntas que PREGUNTAS_PREPARACION_ (Preparacion.gs) — MISMO
 * título, sin modificar su redacción (spec: "no se modifican las
 * preguntas") — pero con una ayuda distinta: en vez de remitir al
 * Informe Ejecutivo de la IE (que el invitado no tiene por qué conocer),
 * explica la pregunta en lenguaje sencillo y da un ejemplo de respuesta
 * (spec: "se explican y se dan ejemplos para ayuda").
 */
var PREGUNTAS_PREPARACION_INVITADO_ = [
  {
    clave: "P1",
    titulo: "Avances en el logro de retos y propósitos del SEM 2025",
    ayuda: "Cuenta, con tus propias palabras, qué cambios o mejoras has notado en tu institución en los " +
      "últimos años: nuevas actividades, mejoras en la enseñanza, proyectos que se hayan logrado, etc. " +
      "Ejemplo: “He notado que ahora hay más actividades deportivas y que los profesores usan más la " +
      "tecnología en las clases.”"
  },
  {
    clave: "P2",
    titulo: "Implementación de niveles de preescolar (Jardín, Pre-jardín)",
    ayuda: "Cuenta si conoces o has visto avances en la atención a los niños y niñas más pequeños (jardín, " +
      "prejardín) en tu institución. Ejemplo: “Sé que abrieron un salón nuevo para los niños pequeños este " +
      "año.” Si no tienes información sobre este tema, puedes dejarlo en blanco."
  },
  {
    clave: "P3",
    titulo: "Pertinencia curricular con las realidades de la comunidad",
    ayuda: "¿Sientes que lo que se enseña en tu institución tiene que ver con la vida real de tu barrio o " +
      "comunidad? ¿Por qué? Ejemplo: “Sí, porque en la clase de sociales hablamos de los problemas de " +
      "nuestro barrio.”"
  },
  {
    clave: "P4",
    titulo: "Acciones pedagógicas para articular el currículo con la comunidad",
    ayuda: "Menciona alguna actividad, proyecto o clase que haya conectado lo aprendido en la institución " +
      "con la comunidad o el entorno. Ejemplo: “Hicimos un proyecto de reciclaje con los vecinos del " +
      "barrio.”"
  },
  {
    clave: "P5",
    titulo: "Equipos de trabajo para articular con la comunidad",
    ayuda: "¿Conoces algún grupo, comité o equipo (de padres, estudiantes, profesores) que trabaje por " +
      "mejorar la relación entre la institución y la comunidad? Cuéntanos cuál y qué hace. Ejemplo: “El " +
      "Consejo de Padres organiza reuniones para hablar de las necesidades del colegio.”"
  },
  {
    clave: "P6",
    titulo: "Democracia institucional",
    ayuda: "¿Sientes que en tu institución las decisiones se toman escuchando a estudiantes y familias? " +
      "Cuenta un ejemplo. Ejemplo: “Sí, porque elegimos al personero y al Consejo Estudiantil entre todos.”"
  }
];

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

  var respuestas = {};
  PREGUNTAS_PREPARACION_INVITADO_.forEach(function (p) {
    respuestas[p.clave] = guardado ? String(guardado[p.clave] || "") : "";
  });

  return {
    preguntas: PREGUNTAS_PREPARACION_INVITADO_,
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
  PREGUNTAS_PREPARACION_INVITADO_.forEach(function (p) {
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
    tieneAlgo = PREGUNTAS_PREPARACION_INVITADO_.some(function (p) {
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
 * Aportes ya enviados por estudiantes y acudientes invitados para una IE
 * del grupo — para mostrarse DURANTE la etapa de preparación (spec:
 * "deben aparecer los aportes de acudientes y de estudiantes"), como
 * insumo de referencia para quien construye libremente la respuesta
 * oficial de la IE. Un párrafo con título por cada pregunta que el
 * invitado haya diligenciado (las vacías se omiten).
 */
function obtenerAportesInvitadosIE(idGrupo, idIE) {
  idGrupo = String(idGrupo || "").trim();
  idIE = String(idIE || "").trim();
  var hoja = obtenerHoja_(HOJA_APORTES_INVITADOS_, cabecerasAportesInvitadosPreparacion_());
  var filas = leerFilasComoObjetos_(hoja);
  var resultado = [];
  filas.forEach(function (f) {
    if (String(f.ID_GRUPO || "").trim() !== idGrupo) return;
    if (String(f.ID_IE || "").trim() !== idIE) return;
    if (String(f.ENVIADO || "") !== "SI") return;
    var secciones = [];
    PREGUNTAS_PREPARACION_INVITADO_.forEach(function (p) {
      var texto = String(f[p.clave] || "").trim();
      if (texto) secciones.push({ clave: p.clave, titulo: p.titulo, texto: texto });
    });
    if (!secciones.length) return;
    resultado.push({
      tipoInvitado: String(f.TIPO_INVITADO || "").toUpperCase(),
      secciones: secciones
    });
  });
  // Estudiantes primero, luego acudientes — orden estable de lectura.
  resultado.sort(function (a, b) {
    if (a.tipoInvitado === b.tipoInvitado) return 0;
    return a.tipoInvitado === "ESTUDIANTE" ? -1 : 1;
  });
  return resultado;
}

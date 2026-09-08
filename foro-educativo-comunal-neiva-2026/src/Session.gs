/**
 * Session.gs — Foro Educativo Comunal Neiva 2026
 *
 * Control de sesión multi-dispositivo por GRUPO, adaptado tal cual del
 * módulo de sesión de FEI 3.1 (docs/01-auditoria-fei-3.1.md §1.2/§5.3/§5.4
 * — reclamarSesionCodigo_, mantenerSesionCodigo_, liberarSesionCodigo_,
 * transferirResponsablePrincipalFEM, sesionActivaPorIdForo_). Cambios
 * frente a 3.1:
 *   - La clave de sesión y el cupo de dispositivos son por ID_GRUPO, no
 *     por IE/ID_FORO.
 *   - sesionActivaPorIdGrupo_ ya nace tratando el valor guardado como
 *     arreglo (el bug histórico de 3.1, corregido en su momento sobre la
 *     marcha, documentado en la auditoría §7, no se reproduce aquí).
 *   - El estado vive en PropertiesService (efímero, no en Sheets) —
 *     mismo criterio que 3.1: no genera contención de escritura ni queda
 *     auditado como fila.
 * No hay expiración por inactividad (decisión deliberada, igual que 3.1):
 * una sesión vive hasta que se libera explícitamente o se fuerza un
 * takeover con `forzar:true`.
 */

function obtenerClaveSesionGrupo_(idGrupo) {
  return "FEC_SESION_GRUPO_" + Utilities.base64EncodeWebSafe(String(idGrupo || "").trim());
}

/** Lee el arreglo de cupos activos, tolerante a propiedad ausente o corrupta. */
function leerSesionesActivas_(props, clave) {
  var guardado = props.getProperty(clave);
  if (!guardado) return [];
  try {
    var parsed = JSON.parse(guardado);
    if (Array.isArray(parsed)) return parsed;
    if (parsed && parsed.deviceId) return [parsed];
    return [];
  } catch (e) {
    return [];
  }
}

function maxSesionesSimultaneasGrupo_() {
  return Number(getConfig().MAX_SESIONES_SIMULTANEAS_GRUPO) || 4;
}

/**
 * Reclama un cupo de dispositivo para el grupo. El primer dispositivo en
 * reclamar un cupo queda esPrincipal=true de forma permanente (salvo
 * transferencia explícita) — es quien puede enviar Sesión 1/2 de forma
 * definitiva; los demás son colaboradores.
 */
function reclamarSesionGrupo_(idGrupo, dispositivoId, forzar) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var props = PropertiesService.getScriptProperties();
    var clave = obtenerClaveSesionGrupo_(idGrupo);
    var ahora = Date.now();
    var sesiones = leerSesionesActivas_(props, clave);
    var maxSesiones = maxSesionesSimultaneasGrupo_();

    var yaHabiaPrincipal = sesiones.some(function (s) {
      return s.esPrincipal;
    });

    var existente = sesiones.find(function (s) {
      return s.deviceId === dispositivoId;
    });
    if (existente) {
      existente.ultimaActividad = ahora;
      if (existente.esPrincipal === undefined && !yaHabiaPrincipal) existente.esPrincipal = true;
      props.setProperty(clave, JSON.stringify(sesiones));
      return { ok: true, tokenSesion: existente.tokenSesion, esPrincipal: !!existente.esPrincipal };
    }

    if (sesiones.length >= maxSesiones) {
      if (!forzar) {
        return {
          ok: false,
          codigo: "SESION_YA_ABIERTA",
          mensaje:
            "Ya hay " + maxSesiones + " dispositivos conectados con este código de acceso, el máximo " +
            "permitido por grupo. Si desea continuar en este dispositivo, se cerrará la conexión del " +
            "dispositivo con menos actividad reciente."
        };
      }
      sesiones.sort(function (a, b) {
        return (a.ultimaActividad || 0) - (b.ultimaActividad || 0);
      });
      var indiceExpulsar = sesiones.findIndex(function (s) {
        return !s.esPrincipal;
      });
      sesiones.splice(indiceExpulsar === -1 ? 0 : indiceExpulsar, 1);
    }

    var tokenSesion = Utilities.getUuid();
    var esPrimeraSesion = !sesiones.some(function (s) {
      return s.esPrincipal;
    });
    sesiones.push({
      deviceId: dispositivoId,
      tokenSesion: tokenSesion,
      ultimaActividad: ahora,
      esPrincipal: esPrimeraSesion
    });
    props.setProperty(clave, JSON.stringify(sesiones));
    return { ok: true, tokenSesion: tokenSesion, esPrincipal: esPrimeraSesion };
  } catch (e) {
    return { ok: false, codigo: "LOCK_SESION_ERROR", mensaje: "No fue posible asegurar la sesión de acceso. Intente nuevamente." };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/** Heartbeat periódico del cliente para mantener el cupo activo. */
function mantenerSesionGrupo(idGrupo, dispositivoId, tokenSesion) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var props = PropertiesService.getScriptProperties();
    var clave = obtenerClaveSesionGrupo_(idGrupo);
    var sesiones = leerSesionesActivas_(props, clave);
    var mia = sesiones.find(function (s) {
      return s.deviceId === dispositivoId && s.tokenSesion === tokenSesion;
    });
    if (!mia) {
      return { ok: false, codigo: "SESION_NO_AUTORIZADA", mensaje: "Este dispositivo ya no tiene un cupo activo en esta sesión." };
    }
    mia.ultimaActividad = Date.now();
    props.setProperty(clave, JSON.stringify(sesiones));
    return { ok: true, esPrincipal: !!mia.esPrincipal };
  } catch (e) {
    return { ok: false, codigo: "HEARTBEAT_ERROR" };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/** El responsable principal cede su rol a otro colaborador conectado (el de actividad más reciente). */
function transferirResponsablePrincipalGrupo(idGrupo, dispositivoId, tokenSesion) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var props = PropertiesService.getScriptProperties();
    var clave = obtenerClaveSesionGrupo_(idGrupo);
    var sesiones = leerSesionesActivas_(props, clave);
    var mia = sesiones.find(function (s) {
      return s.deviceId === dispositivoId && s.tokenSesion === tokenSesion;
    });
    if (!mia) return { ok: false, mensaje: "Este dispositivo ya no tiene un cupo activo en esta sesión." };
    if (!mia.esPrincipal) return { ok: false, mensaje: "Este dispositivo no es el responsable principal." };
    var otras = sesiones.filter(function (s) {
      return s.deviceId !== dispositivoId;
    });
    if (!otras.length) return { ok: false, mensaje: "No hay otro colaborador conectado para transferir el control." };
    otras.sort(function (a, b) {
      return (b.ultimaActividad || 0) - (a.ultimaActividad || 0);
    });
    var nuevoPrincipalId = otras[0].deviceId;
    sesiones.forEach(function (s) {
      s.esPrincipal = s.deviceId === nuevoPrincipalId;
    });
    props.setProperty(clave, JSON.stringify(sesiones));
    return { ok: true, nuevoPrincipalDispositivoId: nuevoPrincipalId };
  } catch (e) {
    return { ok: false, mensaje: "No fue posible transferir el control. Intente nuevamente." };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/** Libera explícitamente el cupo de un dispositivo (cierre de sesión / cambio de pantalla final). */
function liberarSesionGrupo_(idGrupo, dispositivoId, tokenSesion) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var props = PropertiesService.getScriptProperties();
    var clave = obtenerClaveSesionGrupo_(idGrupo);
    var sesiones = leerSesionesActivas_(props, clave);
    var restantes = sesiones.filter(function (s) {
      return !(s.deviceId === dispositivoId && s.tokenSesion === tokenSesion);
    });
    if (restantes.length) props.setProperty(clave, JSON.stringify(restantes));
    else props.deleteProperty(clave);
    return { ok: true };
  } catch (e) {
    return { ok: false };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/**
 * Guardia usada antes de cualquier escritura crítica (Sesion1.gs,
 * ConectaEduca.gs, Informes.gs): confirma que dispositivo+tokenSesion
 * siguen teniendo cupo activo. A diferencia de 3.1, nace ya tratando el
 * valor guardado como arreglo (ver bug documentado en la auditoría §7).
 */
function sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion) {
  var props = PropertiesService.getScriptProperties();
  var clave = obtenerClaveSesionGrupo_(idGrupo);
  var sesiones = leerSesionesActivas_(props, clave);
  return sesiones.some(function (s) {
    return s.deviceId === String(dispositivoId || "") && s.tokenSesion === String(tokenSesion || "");
  });
}

/** true si dispositivo+tokenSesion corresponden al responsable principal del grupo. */
function esPrincipalDeGrupo_(idGrupo, dispositivoId, tokenSesion) {
  var props = PropertiesService.getScriptProperties();
  var sesiones = leerSesionesActivas_(props, obtenerClaveSesionGrupo_(idGrupo));
  var mia = sesiones.find(function (s) {
    return s.deviceId === String(dispositivoId || "") && s.tokenSesion === String(tokenSesion || "");
  });
  return !!(mia && mia.esPrincipal);
}

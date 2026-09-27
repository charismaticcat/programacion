/**
 * AuthService.gs — Autenticación y autorización (Parte H).
 *
 * Super Admin: email + contraseña (hash HMAC-SHA256 iterado con sal y "pepper"),
 * sesión del lado del servidor (CacheService), bloqueo por intentos, 2FA opcional
 * por email y permisos por rol verificados en CADA llamada.
 *
 * Formulario público: token aleatorio por solicitud; en la hoja solo se guarda su hash.
 */

const ROLE_PERMISSIONS = Object.freeze({
  SUPER_ADMIN: ['*'],
  ADMIN: ['dashboard', 'clients.read', 'clients.write', 'clients.create', 'crm.read', 'crm.write',
    'sales.read', 'sales.write', 'files.read', 'files.write', 'reports', 'duplicates',
    'solicitudes.read', 'solicitudes.write', 'catalog.write'],
  VENDEDOR: ['dashboard', 'clients.read', 'clients.write', 'clients.create', 'crm.read', 'crm.write',
    'sales.read', 'sales.write', 'files.read', 'files.write', 'reports', 'solicitudes.read'],
  LECTURA: ['dashboard', 'clients.read', 'crm.read', 'sales.read', 'files.read', 'solicitudes.read']
});
// Permisos exclusivos de SUPER_ADMIN: audit.read, users.manage, config.write, import, export

const Auth = {
  // ================= Contraseñas =================
  _pepper() {
    let p = Props.get('AUTH_PEPPER');
    if (!p) {
      p = Utils.randomToken();
      Props.set('AUTH_PEPPER', p);
    }
    return p;
  },
  hashPassword(password, salt, iterations) {
    const key = Utilities.newBlob(salt + ':' + Auth._pepper()).getBytes();
    let h = Utilities.computeHmacSha256Signature(Utilities.newBlob(String(password)).getBytes(), key);
    for (let i = 0; i < iterations; i++) {
      h = Utilities.computeHmacSha256Signature(h, key);
    }
    return Utilities.base64Encode(h);
  },
  validatePasswordStrength(password) {
    const p = String(password || '');
    if (p.length < CONFIG.PASSWORD_MIN_LENGTH) {
      throw new AppError('WEAK_PASSWORD', 'La contraseña debe tener al menos ' + CONFIG.PASSWORD_MIN_LENGTH + ' caracteres.');
    }
    if (!/[A-Za-z]/.test(p) || !/\d/.test(p)) {
      throw new AppError('WEAK_PASSWORD', 'La contraseña debe combinar letras y números.');
    }
    if (p.length > 200) throw new AppError('WEAK_PASSWORD', 'La contraseña es demasiado larga.');
  },
  /** Devuelve los campos a guardar en USUARIOS_ADMIN para una contraseña. */
  passwordFields(password) {
    Auth.validatePasswordStrength(password);
    const salt = Utils.randomToken().slice(0, 32);
    const iter = CONFIG.PASSWORD_ITERATIONS;
    return { PASSWORD_SALT: salt, PASSWORD_ITER: String(iter), PASSWORD_HASH: Auth.hashPassword(password, salt, iter) };
  },
  /** Genera una contraseña temporal legible (se muestra una sola vez al Super Admin). */
  temporaryPassword() {
    const alphabet = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
    const hex = Utils.randomToken();
    let out = '';
    for (let i = 0; i < 12; i++) out += alphabet[parseInt(hex.substr(i * 2, 2), 16) % alphabet.length];
    return out.slice(0, 5) + '-' + out.slice(5, 9) + '-' + out.slice(9) + '7';
  },

  // ================= Usuarios =================
  findUserByEmail(email) {
    email = Utils.normalizeEmail(email);
    if (!email) return null;
    const rows = SheetService.findAllBy('USUARIOS_ADMIN', 'EMAIL', email);
    return rows[0] || null;
  },
  publicUser(u) {
    return {
      USUARIO_ID: u.USUARIO_ID, EMAIL: u.EMAIL, NOMBRE: u.NOMBRE, ROL: u.ROL,
      ACTIVO: u.ACTIVO, DEBE_CAMBIAR_PASSWORD: u.DEBE_CAMBIAR_PASSWORD,
      ULTIMO_ACCESO: u.ULTIMO_ACCESO, BLOQUEADO_HASTA: u.BLOQUEADO_HASTA, VERSION: u.VERSION
    };
  },
  hasActiveSuperAdmin() {
    try {
      return SheetService.readAll('USUARIOS_ADMIN').some(u => u.ROL === 'SUPER_ADMIN' && u.ACTIVO);
    } catch (e) {
      return false;
    }
  },

  // ================= Login =================
  login(email, password, userAgent) {
    email = Utils.normalizeEmail(email);
    Ctx.set('ANONIMO:login', '', 'ADMIN');
    if (!Utils.isValidEmail(email) || !password) throw new AppError('BAD_CREDENTIALS', 'Email o contraseña incorrectos.');
    RateLimit.hit('login:' + Utils.sha256Hex(email), 10, 900, 'Demasiados intentos. Espere 15 minutos.');
    RateLimit.hit('login:global', 200, 600, 'Demasiados intentos. Espere unos minutos.');

    const user = Auth.findUserByEmail(email);
    const fail = (detail) => {
      Audit.log({ accion: 'LOGIN', entidad: 'USUARIOS_ADMIN', entidadId: user ? user.USUARIO_ID : '', resultado: 'DENEGADO', detalle: detail + ' · ' + email });
      throw new AppError('BAD_CREDENTIALS', 'Email o contraseña incorrectos.');
    };
    if (!user || !user.ACTIVO) {
      // Mismo costo de cálculo que un usuario real, para no revelar qué emails existen.
      Auth.hashPassword(password, 'no-user-salt', CONFIG.PASSWORD_ITERATIONS);
      fail('usuario inexistente o inactivo');
    }

    if (user.BLOQUEADO_HASTA && user.BLOQUEADO_HASTA > Utils.nowIso()) {
      Audit.log({ accion: 'LOGIN', entidad: 'USUARIOS_ADMIN', entidadId: user.USUARIO_ID, resultado: 'DENEGADO', detalle: 'cuenta bloqueada' });
      throw new AppError('LOCKED', 'Cuenta bloqueada temporalmente por intentos fallidos. Intente más tarde.');
    }
    const hash = Auth.hashPassword(password, user.PASSWORD_SALT, Number(user.PASSWORD_ITER) || CONFIG.PASSWORD_ITERATIONS);
    if (!Utils.safeEqual(hash, user.PASSWORD_HASH)) {
      const attempts = Number(user.INTENTOS_FALLIDOS || 0) + 1;
      const patch = { INTENTOS_FALLIDOS: attempts };
      if (attempts >= CONFIG.LOGIN_MAX_ATTEMPTS) {
        const until = new Date(Date.now() + CONFIG.LOGIN_LOCK_MINUTES * 60000);
        patch.BLOQUEADO_HASTA = Utilities.formatDate(until, CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
        patch.INTENTOS_FALLIDOS = 0;
      }
      SheetService.update('USUARIOS_ADMIN', user.USUARIO_ID, patch);
      fail('contraseña incorrecta');
    }
    if (Number(user.INTENTOS_FALLIDOS) || user.BLOQUEADO_HASTA) {
      SheetService.update('USUARIOS_ADMIN', user.USUARIO_ID, { INTENTOS_FALLIDOS: 0, BLOQUEADO_HASTA: '' });
    }

    if (Settings.get('ADMIN_2FA')) {
      const challenge = Utils.randomToken();
      const code = String(parseInt(Utils.randomToken().slice(0, 8), 16) % 1000000);
      const padded = ('000000' + code).slice(-6);
      CacheService.getScriptCache().put('a2:' + Utils.sha256Hex(challenge), JSON.stringify({
        uid: user.USUARIO_ID, code: Utils.sha256Hex(padded + challenge), tries: 0
      }), CONFIG.TWO_FA_TTL_SEC);
      MailApp.sendEmail({
        to: user.EMAIL,
        subject: CONFIG.ADMIN_APPLICATION_NAME + ' · código de acceso',
        body: 'Su código de acceso es: ' + padded + '\n\nVence en 10 minutos. Si usted no intentó ingresar, cambie su contraseña.'
      });
      Audit.log({ accion: 'LOGIN_2FA_ENVIADO', entidad: 'USUARIOS_ADMIN', entidadId: user.USUARIO_ID });
      return { twoFactor: true, challenge: challenge };
    }
    return Auth._completeLogin(user, userAgent);
  },

  verifyTwoFactor(challenge, code, userAgent) {
    Ctx.set('ANONIMO:login', '', 'ADMIN');
    const cache = CacheService.getScriptCache();
    const key = 'a2:' + Utils.sha256Hex(String(challenge || ''));
    const raw = cache.get(key);
    if (!raw) throw new AppError('EXPIRED', 'El código venció. Inicie sesión de nuevo.');
    const data = JSON.parse(raw);
    if (!Utils.safeEqual(Utils.sha256Hex(String(code || '').trim() + challenge), data.code)) {
      data.tries++;
      if (data.tries >= 5) cache.remove(key);
      else cache.put(key, JSON.stringify(data), CONFIG.TWO_FA_TTL_SEC);
      Audit.log({ accion: 'LOGIN_2FA', entidad: 'USUARIOS_ADMIN', entidadId: data.uid, resultado: 'DENEGADO' });
      throw new AppError('BAD_CODE', 'Código incorrecto.');
    }
    cache.remove(key);
    const user = SheetService.getById('USUARIOS_ADMIN', data.uid);
    if (!user || !user.ACTIVO) throw new AppError('BAD_CREDENTIALS', 'Usuario no disponible.');
    return Auth._completeLogin(user, userAgent);
  },

  _completeLogin(user, userAgent) {
    const token = Utils.randomToken();
    const now = Date.now();
    const session = {
      uid: user.USUARIO_ID, email: user.EMAIL, rol: user.ROL, nombre: user.NOMBRE,
      created: now, last: now, mustChange: !!user.DEBE_CAMBIAR_PASSWORD
    };
    const cache = CacheService.getScriptCache();
    const key = 'as:' + Utils.sha256Hex(token);
    cache.put(key, JSON.stringify(session), CONFIG.SESSION_TTL_SEC);
    // Índice de sesiones del usuario (para revocarlas si se desactiva).
    const listKey = 'au:' + user.USUARIO_ID;
    const list = Utils.safeJsonParse(cache.get(listKey), []).slice(-9);
    list.push(key);
    cache.put(listKey, JSON.stringify(list), CONFIG.SESSION_TTL_SEC);

    Ctx.set('ADMIN:' + user.EMAIL, user.ROL, 'ADMIN', user);
    SheetService.update('USUARIOS_ADMIN', user.USUARIO_ID, { ULTIMO_ACCESO: Utils.nowIso() });
    Audit.log({ accion: 'LOGIN', entidad: 'USUARIOS_ADMIN', entidadId: user.USUARIO_ID, detalle: Utils.truncate(userAgent || '', 200) });
    return { token: token, user: Auth.sessionUser(session) };
  },

  sessionUser(s) {
    return {
      id: s.uid, email: s.email, nombre: s.nombre, rol: s.rol, mustChange: !!s.mustChange,
      permissions: ROLE_PERMISSIONS[s.rol] || []
    };
  },

  logout(token) {
    if (!Auth._isTokenFormat(token)) return;
    CacheService.getScriptCache().remove('as:' + Utils.sha256Hex(token));
  },

  revokeUserSessions(userId) {
    const cache = CacheService.getScriptCache();
    const listKey = 'au:' + userId;
    const list = Utils.safeJsonParse(cache.get(listKey), []);
    if (list.length) cache.removeAll(list);
    cache.remove(listKey);
    cache.remove('auc:' + userId);
  },

  _isTokenFormat(token) {
    return typeof token === 'string' && /^[a-f0-9]{64}$/.test(token);
  },

  /**
   * Valida la sesión y el permiso. Lanza AppError si no corresponde.
   * @return {Object} usuario de sesión
   */
  requireAdmin(token, permission) {
    Ctx.set('ANONIMO', '', 'ADMIN');
    if (!Auth._isTokenFormat(token)) throw new AppError('UNAUTHENTICATED', 'Sesión no válida. Inicie sesión.');
    const cache = CacheService.getScriptCache();
    const key = 'as:' + Utils.sha256Hex(token);
    const raw = cache.get(key);
    if (!raw) throw new AppError('UNAUTHENTICATED', 'Su sesión terminó. Inicie sesión de nuevo.');
    const s = JSON.parse(raw);
    const now = Date.now();
    if (now - s.last > CONFIG.SESSION_IDLE_SEC * 1000 || now - s.created > CONFIG.SESSION_TTL_SEC * 1000) {
      cache.remove(key);
      throw new AppError('UNAUTHENTICATED', 'Su sesión terminó por inactividad. Inicie sesión de nuevo.');
    }
    // Verifica que el usuario siga activo y con el mismo rol (caché 60 s).
    let u = Utils.safeJsonParse(cache.get('auc:' + s.uid), null);
    if (!u) {
      const rec = SheetService.getById('USUARIOS_ADMIN', s.uid);
      u = rec ? { activo: rec.ACTIVO, rol: rec.ROL, email: rec.EMAIL, nombre: rec.NOMBRE, mustChange: rec.DEBE_CAMBIAR_PASSWORD } : { activo: false };
      cache.put('auc:' + s.uid, JSON.stringify(u), 60);
    }
    if (!u.activo) {
      cache.remove(key);
      throw new AppError('UNAUTHENTICATED', 'Usuario desactivado.');
    }
    s.rol = u.rol;
    s.nombre = u.nombre;
    s.mustChange = !!u.mustChange;
    s.last = now;
    const remaining = Math.max(60, Math.floor((CONFIG.SESSION_TTL_SEC * 1000 - (now - s.created)) / 1000));
    cache.put(key, JSON.stringify(s), remaining);

    const user = Auth.sessionUser(s);
    Ctx.set('ADMIN:' + s.email, s.rol, 'ADMIN', user);
    if (s.mustChange && permission !== 'self') {
      throw new AppError('PASSWORD_CHANGE_REQUIRED', 'Debe cambiar su contraseña antes de continuar.');
    }
    if (permission && permission !== 'self' && !Auth.can(user, permission)) {
      Audit.log({ accion: 'ACCESO', entidad: 'PERMISO', entidadId: permission, resultado: 'DENEGADO', detalle: 'Permiso insuficiente' });
      throw new AppError('FORBIDDEN', 'No tiene permiso para realizar esta acción.');
    }
    return user;
  },

  can(user, permission) {
    const perms = ROLE_PERMISSIONS[user.rol] || [];
    return perms.indexOf('*') >= 0 || perms.indexOf(permission) >= 0;
  },

  /** VENDEDOR solo ve y edita sus clientes asignados. */
  isScopedToOwn(user) {
    return user && user.rol === 'VENDEDOR';
  },
  canAccessClient(user, cliente) {
    if (!cliente) return false;
    if (!Auth.isScopedToOwn(user)) return true;
    return String(cliente.RESPONSABLE_EMAIL || '').toLowerCase() === user.email;
  },
  requireClientAccess(user, cliente) {
    if (!cliente) throw new AppError('NOT_FOUND', 'Cliente no encontrado.');
    if (!Auth.canAccessClient(user, cliente)) {
      Audit.log({ accion: 'ACCESO', entidad: 'CLIENTES', entidadId: cliente.CLIENTE_ID, clienteId: cliente.CLIENTE_ID, resultado: 'DENEGADO', detalle: 'Cliente no asignado' });
      throw new AppError('FORBIDDEN', 'Este cliente no está asignado a usted.');
    }
  },

  changePassword(user, currentPassword, newPassword) {
    const rec = SheetService.getById('USUARIOS_ADMIN', user.id);
    if (!rec) throw new AppError('NOT_FOUND', 'Usuario no encontrado.');
    const hash = Auth.hashPassword(currentPassword, rec.PASSWORD_SALT, Number(rec.PASSWORD_ITER) || CONFIG.PASSWORD_ITERATIONS);
    if (!Utils.safeEqual(hash, rec.PASSWORD_HASH)) throw new AppError('BAD_CREDENTIALS', 'La contraseña actual no es correcta.');
    if (currentPassword === newPassword) throw new AppError('WEAK_PASSWORD', 'La nueva contraseña debe ser distinta.');
    const fields = Auth.passwordFields(newPassword);
    fields.DEBE_CAMBIAR_PASSWORD = false;
    SheetService.update('USUARIOS_ADMIN', rec.USUARIO_ID, fields);
    CacheService.getScriptCache().remove('auc:' + rec.USUARIO_ID);
    Audit.log({ accion: 'CAMBIO_PASSWORD', entidad: 'USUARIOS_ADMIN', entidadId: rec.USUARIO_ID });
    return true;
  },

  // ================= Primer Super Admin =================
  /** Genera un código de configuración de un solo uso (se ejecuta desde el editor). */
  createSetupCode() {
    const code = Utils.randomToken().slice(0, 12).toUpperCase().replace(/(.{4})(?=.)/g, '$1-');
    Props.set('SETUP_CODE_HASH', Utils.sha256Hex(code));
    const exp = new Date(Date.now() + CONFIG.SETUP_CODE_TTL_HOURS * 3600000);
    Props.set('SETUP_CODE_EXP', String(exp.getTime()));
    return code;
  },
  setupStatus() {
    return {
      configured: !!Props.get('DB_SPREADSHEET_ID'),
      hasSuperAdmin: Auth.hasActiveSuperAdmin(),
      setupCodeActive: !!Props.get('SETUP_CODE_HASH') && Number(Props.get('SETUP_CODE_EXP')) > Date.now()
    };
  },
  /** Crea (o restablece) un SUPER_ADMIN usando el código de configuración. */
  createSuperAdminWithCode(code, email, nombre, password) {
    Ctx.set('SETUP', 'SISTEMA', 'ADMIN');
    RateLimit.hit('setupcode', 10, 3600, 'Demasiados intentos. Espere una hora.');
    const hash = Props.get('SETUP_CODE_HASH');
    const exp = Number(Props.get('SETUP_CODE_EXP'));
    const cleanCode = String(code || '').trim().toUpperCase();
    if (!hash || !exp || exp < Date.now() || !Utils.safeEqual(Utils.sha256Hex(cleanCode), hash)) {
      Audit.log({ accion: 'CREAR_SUPER_ADMIN', entidad: 'USUARIOS_ADMIN', resultado: 'DENEGADO', detalle: 'código inválido o vencido' });
      throw new AppError('BAD_CODE', 'Código de configuración inválido o vencido. Ejecute createSetupCode() en el editor.');
    }
    email = Utils.normalizeEmail(email);
    if (!Utils.isValidEmail(email)) throw new AppError('VALIDATION', 'Email no válido.');
    nombre = Utils.truncate(Utils.cleanSpaces(nombre), 120);
    if (!nombre) throw new AppError('VALIDATION', 'Escriba su nombre.');
    const fields = Auth.passwordFields(password);
    Ctx.set('SETUP:' + email, 'SISTEMA', 'ADMIN');
    const existing = Auth.findUserByEmail(email);
    let userId;
    if (existing) {
      const res = SheetService.update('USUARIOS_ADMIN', existing.USUARIO_ID, Object.assign({
        ROL: 'SUPER_ADMIN', ACTIVO: true, NOMBRE: nombre, DEBE_CAMBIAR_PASSWORD: false,
        INTENTOS_FALLIDOS: 0, BLOQUEADO_HASTA: ''
      }, fields));
      userId = existing.USUARIO_ID;
      Audit.logChange('RESTABLECER_SUPER_ADMIN', 'USUARIOS_ADMIN', userId, '', res);
      Auth.revokeUserSessions(userId);
    } else {
      const rec = SheetService.insert('USUARIOS_ADMIN', Object.assign({
        EMAIL: email, NOMBRE: nombre, ROL: 'SUPER_ADMIN', ACTIVO: true,
        DEBE_CAMBIAR_PASSWORD: false, INTENTOS_FALLIDOS: 0
      }, fields));
      userId = rec.USUARIO_ID;
      Audit.log({ accion: 'CREAR', entidad: 'USUARIOS_ADMIN', entidadId: userId, despues: { EMAIL: email, ROL: 'SUPER_ADMIN' } });
    }
    Props.remove('SETUP_CODE_HASH');
    Props.remove('SETUP_CODE_EXP');
    return { ok: true, email: email };
  },

  // ================= Formulario público =================
  /** Devuelve la solicitud del token, o lanza error. */
  requireDraft(token, opts) {
    Ctx.set('PUBLICO:anonimo', 'PUBLICO', 'PUBLICA');
    if (!Auth._isTokenFormat(token)) throw new AppError('BAD_TOKEN', 'Su sesión del formulario no es válida. Recargue la página.');
    const hash = Utils.sha256Hex(token);
    const cache = CacheService.getScriptCache();
    const cachedId = cache.get('pt:' + hash);
    let sol = cachedId ? SheetService.getById('SOLICITUDES', cachedId) : null;
    if (!sol || sol.TOKEN_HASH !== hash) {
      sol = SheetService.findAllBy('SOLICITUDES', 'TOKEN_HASH', hash)[0] || null;
      if (sol) cache.put('pt:' + hash, sol.SOLICITUD_ID, 21600);
    }
    if (!sol) throw new AppError('BAD_TOKEN', 'No encontramos su formulario. Puede iniciar uno nuevo.');
    Ctx.set('PUBLICO:' + sol.RADICADO, 'PUBLICO', 'PUBLICA');
    if (opts && opts.editable && sol.ESTADO !== 'BORRADOR') {
      throw new AppError('ALREADY_SUBMITTED', 'Este formulario ya fue enviado con el radicado ' + sol.RADICADO + '.');
    }
    return sol;
  },

  // ================= Ejecución del propietario =================
  /**
   * Permite ejecutar funciones de mantenimiento solo desde el editor (propietario)
   * o desde un activador (trigger) instalado por el propietario.
   */
  requireOwnerContext(e) {
    if (e && e.triggerUid) {
      const uid = String(e.triggerUid);
      const ok = ScriptApp.getProjectTriggers().some(t => t.getUniqueId() === uid);
      if (ok) {
        Ctx.set('SISTEMA:trigger', 'SISTEMA', 'SISTEMA');
        return;
      }
    }
    const active = Session.getActiveUser().getEmail();
    const effective = Session.getEffectiveUser().getEmail();
    if (active && effective && active === effective) {
      Ctx.set('SISTEMA:propietario', 'SISTEMA', 'SISTEMA');
      return;
    }
    throw new AppError('FORBIDDEN', 'Función reservada al propietario del script.');
  }
};

/** Límite simple de frecuencia con CacheService. */
const RateLimit = {
  hit(key, limit, windowSec, message) {
    const cache = CacheService.getScriptCache();
    const k = 'rl:' + key;
    const n = Number(cache.get(k) || 0);
    if (n >= limit) throw new AppError('RATE_LIMIT', message || 'Demasiadas solicitudes. Intente más tarde.');
    cache.put(k, String(n + 1), windowSec);
  }
};

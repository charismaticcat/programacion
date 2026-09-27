/**
 * AdminService.gs — Catálogos, productos web, usuarios, ajustes y exportación.
 */
const CatalogService = {
  CATALOGS: ['MUNICIPIOS', 'TIPOS_ESTABLECIMIENTO', 'MOTIVOS_CIERRE'],

  list(catalogo, includeInactive) {
    return SheetService.readAll('CATALOGOS')
      .filter(r => r.CATALOGO === catalogo && (includeInactive || r.ACTIVO))
      .sort((a, b) => a.ORDEN - b.ORDEN || String(a.ETIQUETA).localeCompare(String(b.ETIQUETA)))
      .map(r => { delete r._row; return r; });
  },

  all() {
    const out = {};
    CatalogService.CATALOGS.forEach(c => { out[c] = CatalogService.list(c, true); });
    return out;
  },

  save(user, data) {
    data = data || {};
    if (CatalogService.CATALOGS.indexOf(data.CATALOGO) < 0) throw new AppError('VALIDATION', 'Catálogo no válido.');
    if (data.CATALOGO === 'MUNICIPIOS') {
      throw new AppError('VALIDATION', 'Los municipios están fijados por decisión del negocio (D9): Neiva, Pitalito, Palermo y Rivera.');
    }
    const clean = Validation.cleanRecord('CATALOGOS', data);
    if (!clean.ETIQUETA && !data.ITEM_ID) throw new AppError('VALIDATION', 'Escriba la etiqueta.');
    if (data.ITEM_ID) {
      const res = SheetService.update('CATALOGOS', data.ITEM_ID, clean, { expectedVersion: data.VERSION });
      Audit.logChange('EDITAR', 'CATALOGOS', data.ITEM_ID, '', res);
      return ClientService._strip(res.record);
    }
    const rec = Object.assign({ CATALOGO: data.CATALOGO, CODIGO: Utils.slug(clean.ETIQUETA, 40), ORDEN: 100, ACTIVO: true }, clean);
    SheetService.insert('CATALOGOS', rec);
    Audit.log({ accion: 'CREAR', entidad: 'CATALOGOS', entidadId: rec.ITEM_ID, despues: rec });
    return ClientService._strip(rec);
  }
};

const ProductoWebService = {
  list(includeInactive) {
    return SheetService.readAll('PRODUCTOS_WEB').filter(p => includeInactive || p.ACTIVO)
      .sort((a, b) => a.ORDEN - b.ORDEN).map(p => { delete p._row; return p; });
  },
  save(user, data) {
    data = data || {};
    const clean = Validation.cleanRecord('PRODUCTOS_WEB', data);
    if (data.PRODUCTO_WEB_ID) {
      const res = SheetService.update('PRODUCTOS_WEB', data.PRODUCTO_WEB_ID, clean, { expectedVersion: data.VERSION });
      Audit.logChange('EDITAR', 'PRODUCTOS_WEB', data.PRODUCTO_WEB_ID, '', res);
      return ClientService._strip(res.record);
    }
    if (!clean.NOMBRE) throw new AppError('VALIDATION', 'Escriba el nombre del producto.');
    const rec = Object.assign({ ACTIVO: true, ORDEN: 100, ES_ADICIONAL: false, PRECIO_SUGERIDO: 0 }, clean);
    SheetService.insert('PRODUCTOS_WEB', rec);
    Audit.log({ accion: 'CREAR', entidad: 'PRODUCTOS_WEB', entidadId: rec.PRODUCTO_WEB_ID, despues: clean });
    return ClientService._strip(rec);
  }
};

const UserService = {
  list() {
    return SheetService.readAll('USUARIOS_ADMIN').map(Auth.publicUser);
  },
  /** Lista mínima para asignar responsables (sin datos sensibles). */
  assignable() {
    return SheetService.readAll('USUARIOS_ADMIN').filter(u => u.ACTIVO && u.ROL !== 'LECTURA')
      .map(u => ({ email: u.EMAIL, nombre: u.NOMBRE, rol: u.ROL }));
  },
  /** Crea un usuario con contraseña temporal (se muestra una sola vez). */
  create(user, data) {
    data = data || {};
    const email = Utils.normalizeEmail(data.EMAIL);
    if (!Utils.isValidEmail(email)) throw new AppError('VALIDATION', 'Email no válido.');
    if (Auth.findUserByEmail(email)) throw new AppError('VALIDATION', 'Ya existe un usuario con ese email.');
    const rol = Validation.requireEnum(data.ROL, 'ROL', 'Rol');
    const nombre = Validation.requireText(data.NOMBRE, 'nombre', 120);
    const temp = Auth.temporaryPassword();
    const rec = Object.assign({
      EMAIL: email, NOMBRE: nombre, ROL: rol, ACTIVO: true, DEBE_CAMBIAR_PASSWORD: true, INTENTOS_FALLIDOS: 0
    }, Auth.passwordFields(temp));
    SheetService.insert('USUARIOS_ADMIN', rec);
    Audit.log({ accion: 'CREAR', entidad: 'USUARIOS_ADMIN', entidadId: rec.USUARIO_ID, despues: { EMAIL: email, ROL: rol } });
    return { user: Auth.publicUser(rec), temporaryPassword: temp };
  },
  update(user, id, data) {
    data = data || {};
    const target = SheetService.getById('USUARIOS_ADMIN', id);
    if (!target) throw new AppError('NOT_FOUND', 'Usuario no encontrado.');
    const patch = {};
    if (data.NOMBRE !== undefined) patch.NOMBRE = Validation.requireText(data.NOMBRE, 'nombre', 120);
    if (data.ROL !== undefined) patch.ROL = Validation.requireEnum(data.ROL, 'ROL', 'Rol');
    if (data.ACTIVO !== undefined) patch.ACTIVO = Validation.bool(data.ACTIVO);
    if (target.USUARIO_ID === user.id && (patch.ACTIVO === false || (patch.ROL && patch.ROL !== 'SUPER_ADMIN'))) {
      throw new AppError('VALIDATION', 'No puede quitarse a sí mismo el rol de Super Admin ni desactivarse.');
    }
    if ((patch.ACTIVO === false || (patch.ROL && patch.ROL !== 'SUPER_ADMIN')) && target.ROL === 'SUPER_ADMIN') {
      const others = SheetService.readAll('USUARIOS_ADMIN').filter(u => u.ROL === 'SUPER_ADMIN' && u.ACTIVO && u.USUARIO_ID !== id);
      if (!others.length) throw new AppError('VALIDATION', 'Debe existir al menos un Super Admin activo.');
    }
    const res = SheetService.update('USUARIOS_ADMIN', id, patch, { expectedVersion: data.VERSION });
    Audit.logChange('EDITAR', 'USUARIOS_ADMIN', id, '', res);
    Auth.revokeUserSessions(id);
    return Auth.publicUser(res.record);
  },
  resetPassword(user, id) {
    const target = SheetService.getById('USUARIOS_ADMIN', id);
    if (!target) throw new AppError('NOT_FOUND', 'Usuario no encontrado.');
    const temp = Auth.temporaryPassword();
    const fields = Auth.passwordFields(temp);
    fields.DEBE_CAMBIAR_PASSWORD = true;
    fields.INTENTOS_FALLIDOS = 0;
    fields.BLOQUEADO_HASTA = '';
    SheetService.update('USUARIOS_ADMIN', id, fields);
    Auth.revokeUserSessions(id);
    Audit.log({ accion: 'RESTABLECER_PASSWORD', entidad: 'USUARIOS_ADMIN', entidadId: id });
    return { temporaryPassword: temp };
  }
};

const ConfigService = {
  get() {
    return { settings: Settings.all(), catalogos: CatalogService.all(), productosWeb: ProductoWebService.list(true) };
  },
  save(user, key, value) {
    const def = EDITABLE_SETTINGS[key];
    if (!def) throw new AppError('VALIDATION', 'Ajuste no editable.');
    let v = value;
    if (def.type === 'n') {
      v = Number(value);
      if (!isFinite(v) || v <= 0) throw new AppError('VALIDATION', 'Escriba un número válido.');
      if (key === 'MAX_FILE_SIZE_MB' && v > 20) throw new AppError('VALIDATION', 'Máximo 20 MB (límite práctico de Apps Script).');
      v = String(v);
    } else if (def.type === 'b') {
      v = Validation.bool(value) ? 'SI' : 'NO';
    } else if (def.type === 'list') {
      const items = String(value || '').split(',').map(s => s.trim()).filter(Boolean);
      if (key === 'REPORT_EDITORS' && items.some(e => !Utils.isValidEmail(e))) throw new AppError('VALIDATION', 'Hay emails no válidos.');
      if (key === 'REQUIRED_FIELDS') {
        const allowed = ['razonSocial', 'medioContacto', 'municipio', 'direccion', 'email', 'nombrePropietario', 'TRATAMIENTO_DATOS'];
        if (items.some(i => allowed.indexOf(i) < 0)) throw new AppError('VALIDATION', 'Campos permitidos: ' + allowed.join(', '));
        if (items.indexOf('TRATAMIENTO_DATOS') < 0) items.push('TRATAMIENTO_DATOS');
      }
      v = items.join(',');
    } else {
      v = Validation.short(value, 500);
      if (key === 'DEFAULT_ASSIGNEE' && v) {
        const u = Auth.findUserByEmail(v);
        if (!u || !u.ACTIVO) throw new AppError('VALIDATION', 'El responsable debe ser un usuario activo.');
        v = u.EMAIL;
      }
      if (key === 'PUBLIC_FORM_URL' && v && Validation.url(v) === null) throw new AppError('VALIDATION', 'La URL debe empezar por https://');
    }
    const before = Settings.get(key);
    Lock.run(() => {
      const row = SheetService.findRow('CONFIGURACION', key);
      const rec = { CLAVE: key, VALOR: v, DESCRIPCION: def.label, ACTUALIZADO_EN: Utils.nowIso(), ACTUALIZADO_POR: user.email };
      if (row < 0) SheetService.insert('CONFIGURACION', rec);
      else SheetService.update('CONFIGURACION', key, rec);
    });
    Settings.invalidate();
    Audit.log({ accion: 'EDITAR', entidad: 'CONFIGURACION', entidadId: key, campos: ['VALOR'], antes: { VALOR: before }, despues: { VALOR: v } });
    return Settings.all();
  }
};

const ExportService = {
  COLUMNS: ['CLIENTE_ID', 'RAZON_SOCIAL', 'NOMBRE_COMERCIAL', 'NOMBRE_PROPIETARIO', 'TELEFONO', 'EMAIL', 'MUNICIPIO', 'BARRIO',
    'ESTADO_CRM', 'ESTADO_FORMULARIO', 'RESPONSABLE_EMAIL', 'ORIGEN', 'CIIU_CODIGO', 'TAMANO_EMPRESA', 'PRIORIDAD', 'NO_CONTACTAR'],
  clientsCsv(user, params) {
    const res = SearchIndex.adminQuery(user, Object.assign({}, params, { page: 1, pageSize: 100, returnAll: true }));
    const rows = res.all || [];
    if (rows.length > 40000) throw new AppError('LIMIT', 'Demasiados registros para exportar.');
    const lines = [ExportService.COLUMNS.join(';')];
    rows.forEach(o => lines.push(ExportService.COLUMNS.map(c => Utils.csvEscape(typeof o[c] === 'boolean' ? (o[c] ? 'SI' : 'NO') : o[c])).join(';')));
    Audit.log({ accion: 'EXPORTAR', entidad: 'CLIENTES', detalle: rows.length + ' registros · filtros: ' + JSON.stringify(params && params.filters || {}) + ' q=' + (params && params.q || '') });
    return {
      name: 'clientes_' + Utils.today() + '.csv', mime: 'text/csv',
      base64: Utilities.base64Encode('﻿' + lines.join('\r\n'), Utilities.Charset.UTF_8), count: rows.length
    };
  }
};

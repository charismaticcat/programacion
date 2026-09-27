/**
 * ClientService.gs — Clientes, sus datos públicos (contactos, sedes, catálogo…),
 * archivos, duplicados y fusiones. Todas las funciones reciben el usuario de
 * sesión y verifican acceso en el servidor.
 */
const CHILD_TABLES = ['CONTACTOS', 'SEDES', 'SERVICIOS', 'PRODUCTOS', 'TESTIMONIOS', 'CONTENIDO_WEB'];

const ClientService = {
  nextClientId() {
    return 'CLI-' + Seq.pad(Seq.next('CLIENTE'), 6);
  },

  get(id) {
    return SheetService.getById('CLIENTES', id);
  },

  /** Cliente accesible por el usuario (o error). */
  requireAccessible(user, id) {
    const c = ClientService.get(id);
    Auth.requireClientAccess(user, c);
    return c;
  },

  displayName(c) {
    return c ? (c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL || c.CLIENTE_ID) : '';
  },

  isPhoneShared(phone, excludeId) {
    if (!phone) return false;
    return SearchIndex.records().some(o => o.TELEFONO === phone && o.CLIENTE_ID !== excludeId);
  },

  /** Ficha 360°: datos públicos del cliente + CRM interno (separados). */
  getFull(user, id) {
    const c = ClientService.requireAccessible(user, id);
    const by = (table) => SheetService.findAllBy(table, 'CLIENTE_ID', id).map(r => { delete r._row; return r; });
    const solicitudes = SheetService.findAllBy('SOLICITUDES', 'CLIENTE_ID', id).map(s => ({
      SOLICITUD_ID: s.SOLICITUD_ID, RADICADO: s.RADICADO, ESTADO: s.ESTADO, TIPO_VINCULACION: s.TIPO_VINCULACION,
      ENVIADA_EN: s.ENVIADA_EN, CAMBIOS_ESTADO: s.CAMBIOS_ESTADO, REPORTE_ESTADO: s.REPORTE_ESTADO,
      OBSERVACIONES_ADMIN: s.OBSERVACIONES_ADMIN, VERSION: s.VERSION
    }));
    const archivos = by('ARCHIVOS').filter(a => a.ESTADO !== 'PAPELERA').map(a => ({
      ARCHIVO_ID: a.ARCHIVO_ID, CATEGORIA: a.CATEGORIA, ENTIDAD: a.ENTIDAD, ENTIDAD_ID: a.ENTIDAD_ID,
      NOMBRE_ORIGINAL: a.NOMBRE_ORIGINAL, NOMBRE_DRIVE: a.NOMBRE_DRIVE, MIME_TYPE: a.MIME_TYPE,
      TAMANO_BYTES: a.TAMANO_BYTES, ESTADO: a.ESTADO, MOTIVO_RECHAZO: a.MOTIVO_RECHAZO,
      CREADO_EN: a.CREADO_EN, VERSION: a.VERSION, DRIVE_URL: DriveService.fileUrl(a.DRIVE_FILE_ID)
    }));
    const dups = SheetService.findAllBy('DUPLICADOS', 'CLIENTE_A', id)
      .concat(SheetService.findAllBy('DUPLICADOS', 'CLIENTE_B', id))
      .filter(d => d.ESTADO === 'PENDIENTE')
      .map(d => ({ DUP_ID: d.DUP_ID, OTRO: d.CLIENTE_A === id ? d.CLIENTE_B : d.CLIENTE_A, REGLA: d.REGLA, PUNTAJE: d.PUNTAJE }));
    const cliente = Object.assign({}, c);
    delete cliente._row;
    return {
      cliente: cliente,
      publico: {
        contactos: by('CONTACTOS'),
        sedes: by('SEDES').sort((a, b) => a.ORDEN - b.ORDEN),
        servicios: by('SERVICIOS').sort((a, b) => a.ORDEN - b.ORDEN),
        productos: by('PRODUCTOS').sort((a, b) => a.ORDEN - b.ORDEN),
        testimonios: by('TESTIMONIOS'),
        contenido: by('CONTENIDO_WEB'),
        archivos: archivos,
        autorizaciones: by('AUTORIZACIONES')
      },
      crm: {
        oportunidades: by('OPORTUNIDADES'),
        actividades: by('ACTIVIDADES').sort((a, b) => String(b.FECHA_HORA).localeCompare(String(a.FECHA_HORA))),
        seguimientos: by('SEGUIMIENTOS').sort((a, b) => String(a.FECHA_VENCIMIENTO).localeCompare(String(b.FECHA_VENCIMIENTO))),
        propuestas: by('PROPUESTAS'),
        ventas: by('VENTAS'),
        pagos: by('PAGOS'),
        proyectos: by('PROYECTOS_WEB')
      },
      solicitudes: solicitudes,
      reportes: by('REPORTES').sort((a, b) => b.VERSION_REPORTE - a.VERSION_REPORTE),
      duplicados: dups,
      driveFolderUrl: DriveService.folderUrl(c.DRIVE_FOLDER_ID),
      contentKeys: CONTENT_KEYS
    };
  },

  /** Alta manual desde el admin. */
  create(user, data) {
    const clean = Validation.cleanRecord('CLIENTES', data || {});
    if (!clean.RAZON_SOCIAL && !clean.NOMBRE_COMERCIAL && !clean.NOMBRE_PROPIETARIO) {
      throw new AppError('VALIDATION', 'Escriba al menos la razón social, el nombre comercial o el propietario.');
    }
    const rec = Object.assign({
      CLIENTE_ID: ClientService.nextClientId(),
      ORIGEN: 'ADMIN',
      ESTADO_CRM: 'NUEVO',
      ESTADO_FORMULARIO: 'NO_INICIADO',
      PRIORIDAD: 'MEDIA',
      RESPONSABLE_EMAIL: Auth.isScopedToOwn(user) ? user.email : '',
      ACTIVIDAD_FUENTE: clean.ACTIVIDAD ? 'ADMIN' : ''
    }, clean);
    rec.MUNICIPIO_COD = Validation.municipioCode(rec.MUNICIPIO);
    rec.TELEFONO_COMPARTIDO = ClientService.isPhoneShared(rec.TELEFONO, rec.CLIENTE_ID);
    SheetService.insert('CLIENTES', rec);
    SearchIndex.upsert(rec);
    Audit.log({ accion: 'CREAR', entidad: 'CLIENTES', entidadId: rec.CLIENTE_ID, clienteId: rec.CLIENTE_ID, despues: clean });
    const dups = ClientService.detectDuplicates(rec);
    return { cliente: rec, duplicados: dups };
  },

  /** Edición de cualquier campo editable (D11). */
  update(user, id, patch, expectedVersion) {
    const c = ClientService.requireAccessible(user, id);
    const clean = Validation.cleanRecord('CLIENTES', patch || {});
    if (Object.prototype.hasOwnProperty.call(clean, 'MUNICIPIO')) clean.MUNICIPIO_COD = Validation.municipioCode(clean.MUNICIPIO);
    if (Object.prototype.hasOwnProperty.call(clean, 'ACTIVIDAD') && clean.ACTIVIDAD !== c.ACTIVIDAD) clean.ACTIVIDAD_FUENTE = 'ADMIN';
    if (Object.prototype.hasOwnProperty.call(clean, 'TELEFONO')) clean.TELEFONO_COMPARTIDO = ClientService.isPhoneShared(clean.TELEFONO, id);
    const res = SheetService.update('CLIENTES', id, clean, { expectedVersion: expectedVersion });
    if (res.changed.length) {
      SearchIndex.upsert(res.record);
      Audit.logChange('EDITAR', 'CLIENTES', id, id, res);
    }
    return { cliente: ClientService._strip(res.record), changed: res.changed };
  },

  assign(user, id, email) {
    if (Auth.isScopedToOwn(user)) throw new AppError('FORBIDDEN', 'Solo un administrador puede reasignar clientes.');
    const c = ClientService.get(id);
    if (!c) throw new AppError('NOT_FOUND', 'Cliente no encontrado.');
    email = Utils.normalizeEmail(email);
    if (email) {
      const u = Auth.findUserByEmail(email);
      if (!u || !u.ACTIVO) throw new AppError('VALIDATION', 'El responsable debe ser un usuario activo del CRM.');
    }
    const res = SheetService.update('CLIENTES', id, { RESPONSABLE_EMAIL: email });
    SearchIndex.upsert(res.record);
    Audit.logChange('ASIGNAR_RESPONSABLE', 'CLIENTES', id, id, res);
    // Las oportunidades abiertas siguen al responsable del cliente.
    SheetService.findAllBy('OPORTUNIDADES', 'CLIENTE_ID', id).filter(o => o.ABIERTA).forEach(o => {
      const r = SheetService.update('OPORTUNIDADES', o.OPORTUNIDAD_ID, { RESPONSABLE_EMAIL: email });
      Audit.logChange('ASIGNAR_RESPONSABLE', 'OPORTUNIDADES', o.OPORTUNIDAD_ID, id, r);
    });
    return ClientService._strip(res.record);
  },

  /** Recalcula ESTADO_CRM desde sus oportunidades (copia controlada para filtrar). */
  syncEstadoCrm(clienteId) {
    const opps = SheetService.findAllBy('OPORTUNIDADES', 'CLIENTE_ID', clienteId);
    if (!opps.length) return;
    const byDate = (a, b) => String(b.FECHA_CAMBIO_ETAPA || b.CREADO_EN).localeCompare(String(a.FECHA_CAMBIO_ETAPA || a.CREADO_EN));
    const open = opps.filter(o => o.ABIERTA).sort(byDate);
    const pick = open[0] || opps.sort(byDate)[0];
    const c = ClientService.get(clienteId);
    if (c && c.ESTADO_CRM !== pick.ETAPA) {
      const res = SheetService.update('CLIENTES', clienteId, { ESTADO_CRM: pick.ETAPA });
      SearchIndex.upsert(res.record);
    }
  },

  // ---------- Tablas hijas (datos públicos del cliente) ----------
  saveChild(user, table, record) {
    if (CHILD_TABLES.indexOf(table) < 0) throw new AppError('FORBIDDEN', 'Tabla no editable.');
    record = record || {};
    const key = SchemaUtil.table(table).key;
    if (record[key]) {
      const current = SheetService.getById(table, record[key]);
      if (!current) throw new AppError('NOT_FOUND', 'Registro no encontrado.');
      ClientService.requireAccessible(user, current.CLIENTE_ID);
      const clean = Validation.cleanRecord(table, record);
      const res = SheetService.update(table, record[key], clean, { expectedVersion: record.VERSION });
      Audit.logChange('EDITAR', table, record[key], current.CLIENTE_ID, res);
      return ClientService._strip(res.record);
    }
    const clienteId = String(record.CLIENTE_ID || '');
    ClientService.requireAccessible(user, clienteId);
    const clean = Validation.cleanRecord(table, record);
    const rec = Object.assign({ CLIENTE_ID: clienteId, SOLICITUD_ID: '' }, clean);
    if (table === 'CONTENIDO_WEB') {
      const k = String(record.CAMPO_CLAVE || '');
      if (!CONTENT_KEYS[k]) throw new AppError('VALIDATION', 'Campo de contenido no válido.');
      rec.CAMPO_CLAVE = k;
    }
    if (table === 'SERVICIOS' && record.ACTIVO === undefined) rec.ACTIVO = true;
    if (table === 'PRODUCTOS' && record.DISPONIBLE === undefined) rec.DISPONIBLE = true;
    SheetService.insert(table, rec);
    Audit.log({ accion: 'CREAR', entidad: table, entidadId: rec[key], clienteId: clienteId, despues: clean });
    return ClientService._strip(rec);
  },

  deleteChild(user, table, id) {
    if (CHILD_TABLES.indexOf(table) < 0) throw new AppError('FORBIDDEN', 'Tabla no editable.');
    const current = SheetService.getById(table, id);
    if (!current) throw new AppError('NOT_FOUND', 'Registro no encontrado.');
    ClientService.requireAccessible(user, current.CLIENTE_ID);
    SheetService.softDelete(table, id);
    Audit.log({ accion: 'ELIMINAR', entidad: table, entidadId: id, clienteId: current.CLIENTE_ID });
    return true;
  },

  // ---------- Archivos (admin) ----------
  ensureClientFolder(c) {
    const folder = DriveService.ensureFolder(c.DRIVE_FOLDER_ID, c.CLIENTE_ID + '_' + Utils.slug(ClientService.displayName(c)));
    if (folder.getId() !== c.DRIVE_FOLDER_ID) {
      const res = SheetService.update('CLIENTES', c.CLIENTE_ID, { DRIVE_FOLDER_ID: folder.getId() });
      Audit.logChange('EDITAR', 'CLIENTES', c.CLIENTE_ID, c.CLIENTE_ID, res, 'Carpeta Drive creada');
    }
    return folder.getId();
  },

  uploadFile(user, clienteId, meta, base64) {
    const c = ClientService.requireAccessible(user, clienteId);
    meta = meta || {};
    const category = ENUMS.ARCHIVO_CATEGORIA.indexOf(meta.category) >= 0 ? meta.category : 'OTRO';
    if (category === 'REPORTE' || category === 'SNAPSHOT') throw new AppError('VALIDATION', 'Categoría reservada.');
    const bytes = FileHelper.decode(base64);
    const v = DriveService.validateUpload(meta.fileName, meta.mimeType, category, bytes);
    const folderId = ClientService.ensureClientFolder(c);
    const archivoId = Utils.newId('ARC');
    const driveName = category + '_' + archivoId + '_' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMddHHmmss') + '.' + v.ext;
    const file = DriveService.saveFile(folderId, category, driveName, bytes, v.mime);
    const entidad = ENUMS.ENTIDAD_ARCHIVO.indexOf(meta.entidad) >= 0 ? meta.entidad : 'CLIENTE';
    const rec = {
      ARCHIVO_ID: archivoId, CLIENTE_ID: clienteId, SOLICITUD_ID: '', ENTIDAD: entidad,
      ENTIDAD_ID: Validation.short(meta.entidadId, 40), TEMP_REF: '', CATEGORIA: category,
      NOMBRE_ORIGINAL: Validation.short(meta.fileName, 200), NOMBRE_DRIVE: driveName, DRIVE_FILE_ID: file.getId(),
      DRIVE_FOLDER_ID: folderId, MIME_TYPE: v.mime, TAMANO_BYTES: bytes.length,
      SHA256: Utils.sha256HexBytes(bytes), ESTADO: 'APROBADO', MOTIVO_RECHAZO: ''
    };
    SheetService.insert('ARCHIVOS', rec);
    Audit.log({ accion: 'SUBIR_ARCHIVO', entidad: 'ARCHIVOS', entidadId: archivoId, clienteId: clienteId, despues: { CATEGORIA: category, NOMBRE: rec.NOMBRE_ORIGINAL, TAMANO: rec.TAMANO_BYTES } });
    // Asociar imagen al ítem si corresponde.
    if (rec.ENTIDAD_ID && (entidad === 'SERVICIO' || entidad === 'PRODUCTO' || entidad === 'TESTIMONIO')) {
      const table = entidad === 'SERVICIO' ? 'SERVICIOS' : (entidad === 'PRODUCTO' ? 'PRODUCTOS' : 'TESTIMONIOS');
      const field = entidad === 'TESTIMONIO' ? 'FOTO_ARCHIVO_ID' : 'IMAGEN_ARCHIVO_ID';
      const item = SheetService.getById(table, rec.ENTIDAD_ID);
      if (item && item.CLIENTE_ID === clienteId) {
        const r = SheetService.update(table, rec.ENTIDAD_ID, { [field]: archivoId });
        Audit.logChange('EDITAR', table, rec.ENTIDAD_ID, clienteId, r);
      }
    }
    return { ARCHIVO_ID: archivoId, CATEGORIA: category, NOMBRE_ORIGINAL: rec.NOMBRE_ORIGINAL, TAMANO_BYTES: rec.TAMANO_BYTES, ESTADO: rec.ESTADO };
  },

  _requireFile(user, archivoId) {
    const a = SheetService.getById('ARCHIVOS', archivoId);
    if (!a) throw new AppError('NOT_FOUND', 'Archivo no encontrado.');
    if (a.CLIENTE_ID) ClientService.requireAccessible(user, a.CLIENTE_ID);
    else if (Auth.isScopedToOwn(user)) throw new AppError('FORBIDDEN', 'Archivo no asignado.');
    return a;
  },

  filePreview(user, archivoId) {
    const a = ClientService._requireFile(user, archivoId);
    return { ARCHIVO_ID: a.ARCHIVO_ID, MIME_TYPE: a.MIME_TYPE, dataUrl: DriveService.preview(a.DRIVE_FILE_ID, a.MIME_TYPE, a.TAMANO_BYTES) };
  },

  fileDownload(user, archivoId) {
    const a = ClientService._requireFile(user, archivoId);
    Audit.log({ accion: 'DESCARGAR_ARCHIVO', entidad: 'ARCHIVOS', entidadId: archivoId, clienteId: a.CLIENTE_ID });
    const d = DriveService.download(a.DRIVE_FILE_ID);
    d.name = a.NOMBRE_ORIGINAL || d.name;
    return d;
  },

  setFileStatus(user, archivoId, estado, motivo) {
    const a = ClientService._requireFile(user, archivoId);
    if (['PENDIENTE', 'APROBADO', 'RECHAZADO', 'PAPELERA'].indexOf(estado) < 0) throw new AppError('VALIDATION', 'Estado no válido.');
    const res = SheetService.update('ARCHIVOS', archivoId, { ESTADO: estado, MOTIVO_RECHAZO: estado === 'RECHAZADO' ? Validation.short(motivo, 200) : '' });
    if (estado === 'PAPELERA') DriveService.trash(a.DRIVE_FILE_ID);
    Audit.logChange(estado === 'PAPELERA' ? 'ELIMINAR_ARCHIVO' : 'REVISAR_ARCHIVO', 'ARCHIVOS', archivoId, a.CLIENTE_ID, res);
    return { ARCHIVO_ID: archivoId, ESTADO: estado };
  },

  /** Archivos pendientes de revisión (bandeja de archivos). */
  pendingFiles(user) {
    const rows = SheetService.readColumns('ARCHIVOS', ['ARCHIVO_ID', 'CLIENTE_ID', 'SOLICITUD_ID', 'CATEGORIA', 'NOMBRE_ORIGINAL', 'MIME_TYPE', 'TAMANO_BYTES', 'ESTADO', 'CREADO_EN']);
    const names = SearchIndex.nameMap();
    const idx = Auth.isScopedToOwn(user) ? SearchIndex.records().reduce((m, o) => { m[o.CLIENTE_ID] = o; return m; }, {}) : null;
    return rows.filter(r => r.ESTADO === 'PENDIENTE' && r.CLIENTE_ID)
      .filter(r => !idx || (idx[r.CLIENTE_ID] && Auth.canAccessClient(user, idx[r.CLIENTE_ID])))
      .map(r => { r.CLIENTE_NOMBRE = names[r.CLIENTE_ID] || ''; delete r._row; return r; })
      .sort((a, b) => String(b.CREADO_EN).localeCompare(String(a.CREADO_EN)))
      .slice(0, 300);
  },

  // ---------- Duplicados y conciliación ----------
  /**
   * Busca posibles duplicados de `cliente` en el índice y los registra en DUPLICADOS.
   * Reglas (§5.6): mismo email 90 · teléfono + nombre similar 85 · nombre similar + municipio 70
   *                · mismo teléfono no compartido 60 · mismo teléfono compartido 20 (no se registra)
   */
  detectDuplicates(cliente) {
    const id = cliente.CLIENTE_ID;
    const name = [cliente.RAZON_SOCIAL, cliente.NOMBRE_COMERCIAL, cliente.NOMBRE_PROPIETARIO].filter(Boolean).join(' ');
    const nameNorm = Utils.stripLegalSuffixes(Utils.normalizeText(name));
    const tokens = nameNorm.split(' ').filter(t => t.length >= 4);
    const candidates = [];
    SearchIndex.records().forEach(o => {
      if (o.CLIENTE_ID === id || o.FUSIONADO_EN) return;
      let score = 0;
      let rule = '';
      const sameEmail = cliente.EMAIL && o.EMAIL && cliente.EMAIL === o.EMAIL;
      const samePhone = cliente.TELEFONO && o.TELEFONO && cliente.TELEFONO === o.TELEFONO;
      let sim = 0;
      if (name && (samePhone || tokens.some(t => SearchIndex.nameKey(o).indexOf(t) >= 0))) {
        sim = Math.max(
          Utils.similarity(name, o.RAZON_SOCIAL),
          o.NOMBRE_COMERCIAL ? Utils.similarity(name, o.NOMBRE_COMERCIAL) : 0,
          cliente.RAZON_SOCIAL ? Utils.similarity(cliente.RAZON_SOCIAL, o.RAZON_SOCIAL) : 0
        );
      }
      if (sameEmail) { score = 90; rule = 'MISMO_EMAIL'; }
      else if (samePhone && sim >= 0.6) { score = 85; rule = 'MISMO_TELEFONO_Y_NOMBRE'; }
      else if (sim >= 0.85 && cliente.MUNICIPIO && o.MUNICIPIO === cliente.MUNICIPIO) { score = 70; rule = 'NOMBRE_SIMILAR_MISMO_MUNICIPIO'; }
      else if (samePhone && !o.TELEFONO_COMPARTIDO) { score = 60; rule = 'MISMO_TELEFONO'; }
      if (score >= 60) candidates.push({ CLIENTE_ID: o.CLIENTE_ID, REGLA: rule, PUNTAJE: score, ORIGEN: o.ORIGEN });
    });
    candidates.sort((a, b) => b.PUNTAJE - a.PUNTAJE);
    const top = candidates.slice(0, 5);
    const existing = SheetService.findAllBy('DUPLICADOS', 'CLIENTE_B', id);
    const rows = top.filter(cnd => !existing.some(e => e.CLIENTE_A === cnd.CLIENTE_ID)).map(cnd => ({
      CLIENTE_A: cnd.CLIENTE_ID, CLIENTE_B: id, REGLA: cnd.REGLA, PUNTAJE: cnd.PUNTAJE, ESTADO: 'PENDIENTE'
    }));
    if (rows.length) {
      SheetService.insertMany('DUPLICADOS', rows);
      rows.forEach(r => Audit.log({ accion: 'DUPLICADO_DETECTADO', entidad: 'DUPLICADOS', entidadId: r.DUP_ID, clienteId: id, detalle: r.REGLA + ' ' + r.PUNTAJE + ' con ' + r.CLIENTE_A }));
    }
    return top;
  },

  listDuplicates(user, estado) {
    if (Auth.isScopedToOwn(user)) throw new AppError('FORBIDDEN', 'Solo administradores.');
    const names = SearchIndex.nameMap();
    return SheetService.readAll('DUPLICADOS')
      .filter(d => !estado || d.ESTADO === estado)
      .sort((a, b) => b.PUNTAJE - a.PUNTAJE || String(b.CREADO_EN).localeCompare(String(a.CREADO_EN)))
      .slice(0, 300)
      .map(d => { delete d._row; d.NOMBRE_A = names[d.CLIENTE_A] || ''; d.NOMBRE_B = names[d.CLIENTE_B] || ''; return d; });
  },

  compareDuplicate(user, dupId) {
    const d = SheetService.getById('DUPLICADOS', dupId);
    if (!d) throw new AppError('NOT_FOUND', 'Registro no encontrado.');
    const a = ClientService.get(d.CLIENTE_A);
    const b = ClientService.get(d.CLIENTE_B);
    if (!a || !b) throw new AppError('NOT_FOUND', 'Uno de los clientes ya no existe.');
    delete a._row;
    delete b._row;
    return { dup: d, a: a, b: b, fields: SchemaUtil.editableColumns('CLIENTES') };
  },

  /**
   * Resuelve un posible duplicado.
   * decision: 'ES_EL_MISMO' (fusiona B en A con los valores elegidos) | 'NO_ES_EL_MISMO'
   * choices: {CAMPO: 'A'|'B'} para campos editables.
   */
  resolveDuplicate(user, dupId, decision, choices) {
    if (Auth.isScopedToOwn(user)) throw new AppError('FORBIDDEN', 'Solo administradores.');
    const d = SheetService.getById('DUPLICADOS', dupId);
    if (!d || d.ESTADO !== 'PENDIENTE') throw new AppError('NOT_FOUND', 'Este caso ya fue resuelto.');
    if (decision === 'NO_ES_EL_MISMO') {
      const res = SheetService.update('DUPLICADOS', dupId, { ESTADO: 'NO_ES_EL_MISMO', RESUELTO_POR: user.email, RESUELTO_EN: Utils.nowIso() });
      Audit.logChange('RESOLVER_DUPLICADO', 'DUPLICADOS', dupId, d.CLIENTE_B, res);
      return { merged: false };
    }
    if (decision !== 'ES_EL_MISMO') throw new AppError('VALIDATION', 'Decisión no válida.');
    const a = ClientService.get(d.CLIENTE_A);
    const b = ClientService.get(d.CLIENTE_B);
    if (!a || !b) throw new AppError('NOT_FOUND', 'Uno de los clientes ya no existe.');
    return Lock.run(() => {
      choices = Utils.isPlainObject(choices) ? choices : {};
      const patch = {};
      SchemaUtil.editableColumns('CLIENTES').forEach(f => {
        const pickB = choices[f] === 'B' || (!choices[f] && !a[f] && b[f]);
        if (pickB) patch[f] = b[f];
      });
      if (patch.MUNICIPIO !== undefined) patch.MUNICIPIO_COD = Validation.municipioCode(patch.MUNICIPIO);
      if (!a.RESPONSABLE_EMAIL && b.RESPONSABLE_EMAIL) patch.RESPONSABLE_EMAIL = b.RESPONSABLE_EMAIL;
      if (!a.DRIVE_FOLDER_ID && b.DRIVE_FOLDER_ID) patch.DRIVE_FOLDER_ID = b.DRIVE_FOLDER_ID;
      if (a.ESTADO_FORMULARIO === 'NO_INICIADO' && b.ESTADO_FORMULARIO !== 'NO_INICIADO') patch.ESTADO_FORMULARIO = b.ESTADO_FORMULARIO;
      const resA = SheetService.update('CLIENTES', a.CLIENTE_ID, patch);
      const moved = {};
      SchemaUtil.clientScopedTables().forEach(t => {
        const n = SheetService.reassignClient(t, b.CLIENTE_ID, a.CLIENTE_ID);
        if (n) moved[t] = n;
      });
      const resB = SheetService.update('CLIENTES', b.CLIENTE_ID, { FUSIONADO_EN: a.CLIENTE_ID, ELIMINADO: true });
      SheetService.update('DUPLICADOS', dupId, { ESTADO: 'ES_EL_MISMO', RESUELTO_POR: user.email, RESUELTO_EN: Utils.nowIso() });
      // Otros casos pendientes con B quedan resueltos.
      SheetService.findAllBy('DUPLICADOS', 'CLIENTE_B', b.CLIENTE_ID).concat(SheetService.findAllBy('DUPLICADOS', 'CLIENTE_A', b.CLIENTE_ID))
        .filter(x => x.ESTADO === 'PENDIENTE' && x.DUP_ID !== dupId)
        .forEach(x => SheetService.update('DUPLICADOS', x.DUP_ID, { ESTADO: 'NO_ES_EL_MISMO', RESUELTO_POR: 'SISTEMA:fusion', RESUELTO_EN: Utils.nowIso() }));
      SearchIndex.upsert(resA.record);
      SearchIndex.upsert(Object.assign({}, b, { ELIMINADO: true }));
      ClientService.syncEstadoCrm(a.CLIENTE_ID);
      Audit.log({
        accion: 'FUSIONAR', entidad: 'CLIENTES', entidadId: a.CLIENTE_ID, clienteId: a.CLIENTE_ID,
        campos: resA.changed, antes: resA.before, despues: resA.after,
        detalle: 'Fusionado ' + b.CLIENTE_ID + ' en ' + a.CLIENTE_ID + ' · filas movidas: ' + JSON.stringify(moved)
      });
      Audit.logChange('FUSIONADO', 'CLIENTES', b.CLIENTE_ID, b.CLIENTE_ID, resB);
      return { merged: true, clienteId: a.CLIENTE_ID, moved: moved };
    });
  },

  _strip(o) {
    if (!o) return o;
    const c = Object.assign({}, o);
    delete c._row;
    return c;
  }
};

/** Decodificación de archivos recibidos en base64 (data URL o base64 puro). */
const FileHelper = {
  decode(base64) {
    if (typeof base64 !== 'string' || !base64) throw new AppError('BAD_FILE', 'Archivo vacío.');
    const clean = base64.replace(/^data:[^;]+;base64,/, '').replace(/\s/g, '');
    // Tope previo a decodificar: 20 MB * 4/3 de base64.
    if (clean.length > 28000000) throw new AppError('FILE_TOO_LARGE', 'El archivo es demasiado grande.');
    if (!/^[A-Za-z0-9+/]+=*$/.test(clean)) throw new AppError('BAD_FILE', 'Archivo dañado o incompleto.');
    return Utilities.base64Decode(clean);
  }
};

/**
 * FormService.gs — Lógica del formulario público.
 *
 * Flujo: bienvenida → (búsqueda opcional en la base) → datos → … → autorizaciones
 * → revisión → envío. El formulario no exige identificarse (D10); la búsqueda es
 * una ayuda para precargar datos (Parte F).
 *
 * Seguridad:
 * - Cada función recibe el token del borrador; nunca IDs de clientes.
 * - La búsqueda devuelve referencias opacas y datos enmascarados; solo se
 *   precargan datos del registro que el usuario eligió de SUS resultados.
 * - Por nombre no se precargan teléfono ni email (solo pistas enmascaradas).
 */
const PUBLIC_CATEGORIES = ['LOGO', 'FOTO_NEGOCIO', 'FOTO_EQUIPO', 'FOTO_PRODUCTO', 'FOTO_SERVICIO', 'FOTO_TESTIMONIO', 'DOCUMENTO', 'OTRO'];

const FormService = {
  bootstrap() {
    const tipos = CatalogService.list('TIPOS_ESTABLECIMIENTO').map(t => t.ETIQUETA);
    return {
      appName: Settings.get('APP_NAME'),
      responsable: Settings.get('RESPONSABLE_TRATAMIENTO'),
      contactoPrivacidad: Settings.get('CONTACTO_PRIVACIDAD'),
      municipios: CONFIG.MUNICIPIOS.map(m => m.name),
      tiposEstablecimiento: tipos,
      required: Settings.get('REQUIRED_FIELDS'),
      maxFileMb: Math.round(Settings.maxFileBytes() / 1048576),
      maxFiles: CONFIG.MAX_FILES_PER_SOLICITUD,
      limits: {
        sedes: CONFIG.MAX_SEDES, servicios: CONFIG.MAX_SERVICIOS, productos: CONFIG.MAX_PRODUCTOS,
        testimonios: CONFIG.MAX_TESTIMONIOS, testimonioPalabras: CONFIG.TESTIMONIO_MAX_PALABRAS
      },
      autorizaciones: AUTH_TEXTS.map(a => ({ tipo: a.tipo, titulo: a.titulo, obligatoria: a.obligatoria, texto: authTextFor_(a.tipo) })),
      autorizacionVersion: CONFIG.AUTORIZACION_VERSION,
      canales: ENUMS.CANAL
    };
  },

  startDraft(meta) {
    meta = meta || {};
    Ctx.set('PUBLICO:nuevo', 'PUBLICO', 'PUBLICA');
    if (meta.hp) {
      Audit.log({ accion: 'SPAM_BLOQUEADO', entidad: 'SOLICITUDES', resultado: 'DENEGADO', detalle: 'honeypot' });
      throw new AppError('BLOCKED', 'No pudimos iniciar el formulario. Recargue la página.');
    }
    RateLimit.hit('drafts', CONFIG.RATE_DRAFTS_PER_10MIN, 600, 'Estamos recibiendo muchas solicitudes. Intente en unos minutos.');
    const year = Utils.year();
    const radicado = 'WEB-' + year + '-' + Seq.pad(Seq.next('RADICADO_' + year), 6);
    const token = Utils.randomToken();
    Ctx.set('PUBLICO:' + radicado, 'PUBLICO', 'PUBLICA');
    const sol = SheetService.insert('SOLICITUDES', {
      RADICADO: radicado, CLIENTE_ID: '', CLIENTE_BASE_ID: '', TIPO_VINCULACION: 'NINGUNA', ESTADO: 'BORRADOR',
      PASO_ACTUAL: 0, PORCENTAJE: 0, TOKEN_HASH: Utils.sha256Hex(token), DATOS_JSON: '{}', CAMBIOS_ESTADO: 'NINGUNO',
      REPORTE_ESTADO: 'NO_APLICA', INICIADA_EN: Utils.nowIso(), USER_AGENT: Utils.truncate(meta.ua || '', 300)
    });
    CacheService.getScriptCache().put('pt:' + sol.TOKEN_HASH, sol.SOLICITUD_ID, 21600);
    Audit.log({ accion: 'CREAR', entidad: 'SOLICITUDES', entidadId: sol.SOLICITUD_ID, detalle: radicado });
    return { token: token, radicado: radicado };
  },

  _readData(sol) {
    if (sol.DATOS_FILE_ID) return DriveService.readJson(sol.DATOS_FILE_ID) || {};
    return Utils.isPlainObject(sol.DATOS_JSON) ? sol.DATOS_JSON : {};
  },

  _publicFiles(sol) {
    return SheetService.findAllBy('ARCHIVOS', 'SOLICITUD_ID', sol.SOLICITUD_ID)
      .filter(a => a.ESTADO !== 'PAPELERA')
      .map(a => ({ id: a.ARCHIVO_ID, category: a.CATEGORIA, name: a.NOMBRE_ORIGINAL, size: a.TAMANO_BYTES, tempRef: a.TEMP_REF, mime: a.MIME_TYPE }));
  },

  resume(token) {
    const sol = Auth.requireDraft(token);
    if (sol.ESTADO !== 'BORRADOR') {
      return { submitted: true, radicado: sol.RADICADO, enviadaEn: sol.ENVIADA_EN };
    }
    let vinculado = null;
    if (sol.CLIENTE_BASE_ID) {
      const c = ClientService.get(sol.CLIENTE_BASE_ID);
      if (c) vinculado = { nombre: c.RAZON_SOCIAL, tipo: sol.TIPO_VINCULACION };
    }
    return {
      submitted: false, radicado: sol.RADICADO, paso: sol.PASO_ACTUAL, data: FormService._readData(sol),
      files: FormService._publicFiles(sol), vinculado: vinculado
    };
  },

  // ---------- Búsqueda en la base (Parte F) ----------
  search(token, type, value, municipio) {
    const sol = Auth.requireDraft(token, { editable: true });
    RateLimit.hit('search:' + sol.SOLICITUD_ID, CONFIG.RATE_SEARCH_PER_DRAFT, 3600, 'Alcanzó el máximo de búsquedas. Continúe llenando el formulario sin buscar.');
    RateLimit.hit('search:global', CONFIG.RATE_SEARCH_PER_MIN_GLOBAL, 60, 'El buscador está ocupado. Intente en un minuto o continúe sin buscar.');
    if (['telefono', 'email', 'nombre'].indexOf(type) < 0) throw new AppError('VALIDATION', 'Tipo de búsqueda no válido.');
    const found = SearchIndex.publicSearch(type, String(value || '').slice(0, 150), municipio);
    Audit.log({ accion: 'BUSQUEDA_PUBLICA', entidad: 'SOLICITUDES', entidadId: sol.SOLICITUD_ID, detalle: type + ' · ' + found.length + ' resultado(s)' });
    if (!found.length) return { items: [], tooMany: false };
    if (found.length > CONFIG.SEARCH_MAX_RESULTS) {
      return { items: [], tooMany: true, count: found.length };
    }
    const allowed = {};
    const items = found.map(o => {
      const ref = Utils.randomToken().slice(0, 16);
      allowed[ref] = { id: o.CLIENTE_ID, type: type };
      return {
        ref: ref,
        nombre: o.NOMBRE_COMERCIAL || o.RAZON_SOCIAL,
        razonSocial: o.RAZON_SOCIAL,
        municipio: o.MUNICIPIO,
        barrio: o.BARRIO,
        telefono: Utils.maskPhone(o.TELEFONO),
        email: Utils.maskEmail(o.EMAIL)
      };
    });
    const cache = CacheService.getScriptCache();
    const prev = Utils.safeJsonParse(cache.get('ps:' + sol.SOLICITUD_ID), {});
    cache.put('ps:' + sol.SOLICITUD_ID, JSON.stringify(Object.assign(prev, allowed)), 3600);
    return { items: items, tooMany: false };
  },

  /** El usuario elige un resultado de su búsqueda: se precargan sus datos. */
  select(token, ref) {
    const sol = Auth.requireDraft(token, { editable: true });
    const allowed = Utils.safeJsonParse(CacheService.getScriptCache().get('ps:' + sol.SOLICITUD_ID), {});
    const hit = allowed[String(ref || '')];
    if (!hit) {
      Audit.log({ accion: 'SELECCION_INVALIDA', entidad: 'SOLICITUDES', entidadId: sol.SOLICITUD_ID, resultado: 'DENEGADO' });
      throw new AppError('BAD_REF', 'La búsqueda venció. Busque de nuevo.');
    }
    const c = ClientService.get(hit.id);
    if (!c || c.FUSIONADO_EN) throw new AppError('NOT_FOUND', 'Ese registro ya no está disponible.');
    const tipo = hit.type === 'telefono' ? 'BUSQUEDA_TELEFONO' : (hit.type === 'email' ? 'BUSQUEDA_EMAIL' : 'BUSQUEDA_NOMBRE');
    const res = SheetService.update('SOLICITUDES', sol.SOLICITUD_ID, { CLIENTE_BASE_ID: c.CLIENTE_ID, TIPO_VINCULACION: tipo });
    Audit.logChange('VINCULAR_BASE', 'SOLICITUDES', sol.SOLICITUD_ID, c.CLIENTE_ID, res);

    const full = hit.type !== 'nombre';
    const isMobile = /^3\d{9}$/.test(c.TELEFONO);
    return {
      prefill: {
        negocio: {
          razonSocial: c.RAZON_SOCIAL, nombrePropietario: c.NOMBRE_PROPIETARIO, nombreComercial: c.NOMBRE_COMERCIAL,
          tipoEstablecimiento: c.TIPO_ESTABLECIMIENTO, actividad: c.ACTIVIDAD
        },
        contacto: {
          whatsapp: full && isMobile ? c.TELEFONO : '',
          telefonoFijo: full && !isMobile ? c.TELEFONO : '',
          email: full ? c.EMAIL : ''
        },
        sede: { direccion: c.DIRECCION, barrio: c.BARRIO, municipio: c.MUNICIPIO }
      },
      hints: full ? null : { telefono: Utils.maskPhone(c.TELEFONO), email: Utils.maskEmail(c.EMAIL) },
      vinculado: { nombre: c.RAZON_SOCIAL, tipo: tipo }
    };
  },

  unlink(token) {
    const sol = Auth.requireDraft(token, { editable: true });
    const res = SheetService.update('SOLICITUDES', sol.SOLICITUD_ID, { CLIENTE_BASE_ID: '', TIPO_VINCULACION: 'NINGUNA' });
    Audit.logChange('DESVINCULAR_BASE', 'SOLICITUDES', sol.SOLICITUD_ID, sol.CLIENTE_BASE_ID, res);
    return true;
  },

  // ---------- Borrador ----------
  saveDraft(token, raw, step) {
    const sol = Auth.requireDraft(token, { editable: true });
    const cleaned = Validation.cleanForm(raw, false);
    const json = JSON.stringify(cleaned.data);
    const patch = {
      PASO_ACTUAL: Math.max(0, Math.min(20, Number(step) || 0)),
      PORCENTAJE: Validation.progress(cleaned.data)
    };
    if (json.length <= CONFIG.DRAFT_JSON_MAX_CHARS) {
      patch.DATOS_JSON = json;
      patch.DATOS_FILE_ID = '';
    } else {
      const folderId = FormService._ensureFolder(sol, cleaned.data);
      if (sol.DATOS_FILE_ID) DriveService.updateJson(sol.DATOS_FILE_ID, cleaned.data);
      else patch.DATOS_FILE_ID = DriveService.writeJson(folderId, 'BORRADOR_' + sol.RADICADO + '.json', cleaned.data);
      patch.DATOS_JSON = '{}';
    }
    SheetService.update('SOLICITUDES', sol.SOLICITUD_ID, patch);
    return { savedAt: Utils.nowIso(), porcentaje: patch.PORCENTAJE, warnings: cleaned.errors };
  },

  _ensureFolder(sol, data) {
    const name = sol.RADICADO + '_' + Utils.slug((data && data.negocio && (data.negocio.nombreComercial || data.negocio.razonSocial)) || 'SIN-NOMBRE');
    const folder = DriveService.ensureFolder(sol.DRIVE_FOLDER_ID, name);
    if (folder.getId() !== sol.DRIVE_FOLDER_ID) {
      SheetService.update('SOLICITUDES', sol.SOLICITUD_ID, { DRIVE_FOLDER_ID: folder.getId() });
      sol.DRIVE_FOLDER_ID = folder.getId();
    }
    return folder.getId();
  },

  // ---------- Archivos ----------
  upload(token, meta, base64) {
    const sol = Auth.requireDraft(token, { editable: true });
    meta = meta || {};
    const category = String(meta.category || '');
    if (PUBLIC_CATEGORIES.indexOf(category) < 0) throw new AppError('VALIDATION', 'Categoría de archivo no válida.');
    RateLimit.hit('upl:' + sol.SOLICITUD_ID, CONFIG.RATE_UPLOADS_PER_DRAFT, 3600, 'Alcanzó el máximo de cargas por hora.');
    const existing = SheetService.findAllBy('ARCHIVOS', 'SOLICITUD_ID', sol.SOLICITUD_ID).filter(a => a.ESTADO !== 'PAPELERA');
    if (existing.length >= CONFIG.MAX_FILES_PER_SOLICITUD) throw new AppError('LIMIT', 'Alcanzó el máximo de ' + CONFIG.MAX_FILES_PER_SOLICITUD + ' archivos.');
    const bytes = FileHelper.decode(base64);
    const total = existing.reduce((s, a) => s + Number(a.TAMANO_BYTES || 0), 0) + bytes.length;
    if (total > CONFIG.MAX_TOTAL_MB_SOLICITUD * 1048576) throw new AppError('LIMIT', 'Superó el espacio total de ' + CONFIG.MAX_TOTAL_MB_SOLICITUD + ' MB.');
    const v = DriveService.validateUpload(meta.fileName, meta.mimeType, category, bytes);
    const sha = Utils.sha256HexBytes(bytes);
    const tempRef = /^[A-Za-z0-9-]{1,40}$/.test(String(meta.tempRef || '')) ? String(meta.tempRef) : '';
    const dup = existing.find(a => a.SHA256 === sha && a.CATEGORIA === category && a.TEMP_REF === tempRef);
    if (dup) {
      return { id: dup.ARCHIVO_ID, category: dup.CATEGORIA, name: dup.NOMBRE_ORIGINAL, size: dup.TAMANO_BYTES, tempRef: dup.TEMP_REF, mime: dup.MIME_TYPE, duplicate: true };
    }
    const folderId = FormService._ensureFolder(sol, FormService._readData(sol));
    const archivoId = Utils.newId('ARC');
    const driveName = category + '_' + archivoId + '_' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMddHHmmss') + '.' + v.ext;
    const file = DriveService.saveFile(folderId, category, driveName, bytes, v.mime);

    // Logo y foto por ítem son únicos: el anterior va a la papelera.
    const replaces = existing.filter(a => a.CATEGORIA === category && (category === 'LOGO' || (tempRef && a.TEMP_REF === tempRef &&
      ['FOTO_PRODUCTO', 'FOTO_SERVICIO', 'FOTO_TESTIMONIO'].indexOf(category) >= 0)));
    replaces.forEach(a => {
      SheetService.update('ARCHIVOS', a.ARCHIVO_ID, { ESTADO: 'PAPELERA' });
      DriveService.trash(a.DRIVE_FILE_ID);
    });

    const entidad = category === 'FOTO_PRODUCTO' ? 'PRODUCTO' : (category === 'FOTO_SERVICIO' ? 'SERVICIO' : (category === 'FOTO_TESTIMONIO' ? 'TESTIMONIO' : 'CLIENTE'));
    const rec = {
      ARCHIVO_ID: archivoId, CLIENTE_ID: sol.CLIENTE_BASE_ID || '', SOLICITUD_ID: sol.SOLICITUD_ID,
      ENTIDAD: entidad, ENTIDAD_ID: '', TEMP_REF: tempRef, CATEGORIA: category,
      NOMBRE_ORIGINAL: Validation.short(meta.fileName, 200), NOMBRE_DRIVE: driveName, DRIVE_FILE_ID: file.getId(),
      DRIVE_FOLDER_ID: folderId, MIME_TYPE: v.mime, TAMANO_BYTES: bytes.length, SHA256: sha,
      ESTADO: 'PENDIENTE', MOTIVO_RECHAZO: ''
    };
    SheetService.insert('ARCHIVOS', rec);
    Audit.log({ accion: 'SUBIR_ARCHIVO', entidad: 'ARCHIVOS', entidadId: archivoId, detalle: category + ' · ' + bytes.length + ' bytes' });
    return { id: archivoId, category: category, name: rec.NOMBRE_ORIGINAL, size: bytes.length, tempRef: tempRef, mime: v.mime };
  },

  removeFile(token, archivoId) {
    const sol = Auth.requireDraft(token, { editable: true });
    const a = SheetService.getById('ARCHIVOS', String(archivoId || ''));
    if (!a || a.SOLICITUD_ID !== sol.SOLICITUD_ID) throw new AppError('NOT_FOUND', 'Archivo no encontrado.');
    SheetService.update('ARCHIVOS', a.ARCHIVO_ID, { ESTADO: 'PAPELERA' });
    DriveService.trash(a.DRIVE_FILE_ID);
    Audit.log({ accion: 'ELIMINAR_ARCHIVO', entidad: 'ARCHIVOS', entidadId: a.ARCHIVO_ID });
    return true;
  },

  filePreview(token, archivoId) {
    const sol = Auth.requireDraft(token);
    const a = SheetService.getById('ARCHIVOS', String(archivoId || ''));
    if (!a || a.SOLICITUD_ID !== sol.SOLICITUD_ID || a.ESTADO === 'PAPELERA') throw new AppError('NOT_FOUND', 'Archivo no encontrado.');
    return { id: a.ARCHIVO_ID, dataUrl: DriveService.preview(a.DRIVE_FILE_ID, a.MIME_TYPE, a.TAMANO_BYTES) };
  },

  // ---------- Envío ----------
  /** Campos principales del cliente a partir del formulario. */
  coreFromForm(data) {
    const s0 = data.sedes[0] || {};
    return {
      NOMBRE_PROPIETARIO: data.negocio.nombrePropietario,
      RAZON_SOCIAL: data.negocio.razonSocial,
      NOMBRE_COMERCIAL: data.negocio.nombreComercial,
      TIPO_ESTABLECIMIENTO: data.negocio.tipoEstablecimiento,
      ACTIVIDAD: data.negocio.actividad,
      TELEFONO: data.contacto.whatsapp || data.contacto.telefonoFijo,
      EMAIL: data.contacto.email,
      MUNICIPIO: s0.municipio || '',
      DIRECCION: s0.direccion || '',
      BARRIO: s0.barrio || ''
    };
  },

  submit(token, raw, meta) {
    meta = meta || {};
    let sol = Auth.requireDraft(token);
    if (sol.ESTADO !== 'BORRADOR') {
      return { radicado: sol.RADICADO, enviadaEn: sol.ENVIADA_EN, already: true };
    }
    const secs = (Date.now() - (Utils.parseIsoDate(sol.INICIADA_EN) || new Date()).getTime()) / 1000;
    if (secs < CONFIG.MIN_SECONDS_BEFORE_SUBMIT) {
      throw new AppError('TOO_FAST', 'Revise la información antes de enviar.');
    }
    const cleaned = Validation.cleanForm(raw, true);
    if (cleaned.errors.length) {
      throw new AppError('VALIDATION', 'Revise los campos marcados.', cleaned.errors);
    }
    const data = cleaned.data;

    // 1) Reclamar el envío dentro del bloqueo (evita doble envío).
    const claim = Lock.run(() => {
      const fresh = SheetService.getById('SOLICITUDES', sol.SOLICITUD_ID);
      if (fresh.ESTADO !== 'BORRADOR') return { already: true, sol: fresh };
      SheetService.update('SOLICITUDES', sol.SOLICITUD_ID, { ESTADO: 'ENVIADA', ENVIADA_EN: Utils.nowIso() });
      return { already: false, sol: SheetService.getById('SOLICITUDES', sol.SOLICITUD_ID) };
    });
    if (claim.already) return { radicado: claim.sol.RADICADO, enviadaEn: claim.sol.ENVIADA_EN, already: true };
    sol = claim.sol;

    const created = [];
    try {
      const result = FormService._materialize(sol, data, meta, created);
      Audit.log({ accion: 'ENVIAR_FORMULARIO', entidad: 'SOLICITUDES', entidadId: sol.SOLICITUD_ID, clienteId: result.clienteId, detalle: sol.RADICADO });
      return { radicado: sol.RADICADO, enviadaEn: sol.ENVIADA_EN, already: false };
    } catch (e) {
      // Deshacer lo creado y devolver la solicitud a borrador para reintentar.
      created.forEach(c => {
        try { SheetService.softDelete(c.table, c.id); } catch (x) { console.error('ROLLBACK', c, x && x.message); }
      });
      try { SheetService.update('SOLICITUDES', sol.SOLICITUD_ID, { ESTADO: 'BORRADOR', ENVIADA_EN: '' }); } catch (x) { console.error('ROLLBACK_SOL', x && x.message); }
      throw e;
    }
  },

  _materialize(sol, data, meta, created) {
    const core = FormService.coreFromForm(data);
    const track = (table, rows) => {
      const key = SchemaUtil.table(table).key;
      rows.forEach(r => created.push({ table: table, id: r[key] }));
    };
    const folderId = FormService._ensureFolder(sol, data);
    DriveService.rename(folderId, sol.RADICADO + '_' + Utils.slug(data.negocio.nombreComercial || data.negocio.razonSocial || 'SIN-NOMBRE'));
    const defaultAssignee = Settings.get('DEFAULT_ASSIGNEE');

    // 2) Cliente: existente (vinculado) o nuevo.
    let cliente;
    let pending = [];
    let isNew = false;
    const base = sol.CLIENTE_BASE_ID ? ClientService.get(sol.CLIENTE_BASE_ID) : null;
    if (base) {
      const auto = {};
      Object.keys(core).forEach(k => {
        const proposed = core[k];
        if (!proposed) return;
        const current = base[k];
        if (!current) auto[k] = proposed;
        else if (Utils.normalizeText(current) !== Utils.normalizeText(proposed)) {
          pending.push({ campo: k, actual: current, propuesto: proposed });
        }
      });
      if (auto.MUNICIPIO) auto.MUNICIPIO_COD = Validation.municipioCode(auto.MUNICIPIO);
      if (auto.ACTIVIDAD) auto.ACTIVIDAD_FUENTE = 'CLIENTE';
      if (auto.TELEFONO) auto.TELEFONO_COMPARTIDO = ClientService.isPhoneShared(auto.TELEFONO, base.CLIENTE_ID);
      auto.ESTADO_FORMULARIO = 'ENVIADO';
      auto.DRIVE_FOLDER_ID = folderId;
      if (!base.RESPONSABLE_EMAIL && defaultAssignee) auto.RESPONSABLE_EMAIL = defaultAssignee;
      const res = SheetService.update('CLIENTES', base.CLIENTE_ID, auto);
      Audit.logChange('EDITAR', 'CLIENTES', base.CLIENTE_ID, base.CLIENTE_ID, res, 'Datos completados desde el formulario ' + sol.RADICADO);
      cliente = res.record;
    } else {
      isNew = true;
      cliente = Object.assign({
        CLIENTE_ID: ClientService.nextClientId(), ORIGEN: 'FORMULARIO', ESTADO_CRM: 'NUEVO', ESTADO_FORMULARIO: 'ENVIADO',
        PRIORIDAD: 'MEDIA', RESPONSABLE_EMAIL: defaultAssignee, DRIVE_FOLDER_ID: folderId
      }, core);
      cliente.MUNICIPIO_COD = Validation.municipioCode(cliente.MUNICIPIO);
      cliente.ACTIVIDAD_FUENTE = cliente.ACTIVIDAD ? 'CLIENTE' : '';
      cliente.TELEFONO_COMPARTIDO = ClientService.isPhoneShared(cliente.TELEFONO, cliente.CLIENTE_ID);
      SheetService.insert('CLIENTES', cliente);
      created.push({ table: 'CLIENTES', id: cliente.CLIENTE_ID });
      Audit.log({ accion: 'CREAR', entidad: 'CLIENTES', entidadId: cliente.CLIENTE_ID, clienteId: cliente.CLIENTE_ID, despues: core, detalle: 'Desde formulario ' + sol.RADICADO });
    }
    const clienteId = cliente.CLIENTE_ID;
    const sid = sol.SOLICITUD_ID;

    // 3) Archivos del borrador: se asocian al cliente y a sus ítems.
    const files = SheetService.findAllBy('ARCHIVOS', 'SOLICITUD_ID', sid).filter(a => a.ESTADO !== 'PAPELERA');
    const fileFor = (category, ref) => files.find(f => f.CATEGORIA === category && f.TEMP_REF === ref);

    const servicios = data.servicios.map((s, i) => {
      const f = fileFor('FOTO_SERVICIO', s.ref);
      return { SERVICIO_ID: Utils.newId('SRV'), CLIENTE_ID: clienteId, SOLICITUD_ID: sid, NOMBRE: s.nombre, DESCRIPCION: s.descripcion,
        PRECIO: s.precio, MOSTRAR_PRECIO: s.mostrarPrecio, REQUIERE_CITA: s.requiereCita, IMAGEN_ARCHIVO_ID: f ? f.ARCHIVO_ID : '', ORDEN: i + 1, ACTIVO: true, _ref: s.ref };
    });
    const productos = data.productos.map((p, i) => {
      const f = fileFor('FOTO_PRODUCTO', p.ref);
      return { PRODUCTO_ID: Utils.newId('PRD'), CLIENTE_ID: clienteId, SOLICITUD_ID: sid, NOMBRE: p.nombre, DESCRIPCION: p.descripcion,
        CATEGORIA: p.categoria, PRECIO: p.precio, MOSTRAR_PRECIO: p.mostrarPrecio, DISPONIBLE: p.disponible, IMAGEN_ARCHIVO_ID: f ? f.ARCHIVO_ID : '', ORDEN: i + 1, _ref: p.ref };
    });
    const testimonios = data.testimonios.map(t => {
      const f = fileFor('FOTO_TESTIMONIO', t.ref);
      return { TESTIMONIO_ID: Utils.newId('TES'), CLIENTE_ID: clienteId, SOLICITUD_ID: sid, AUTOR_NOMBRE: t.autor, TEXTO: t.texto,
        FOTO_ARCHIVO_ID: f ? f.ARCHIVO_ID : '', PERMISO_USO_FOTO: t.permisoFoto, PERMISO_USO_NOMBRE: t.permisoNombre, APROBADO_ADMIN: false, _ref: t.ref };
    });
    const sedes = data.sedes.map((s, i) => ({
      CLIENTE_ID: clienteId, SOLICITUD_ID: sid, NOMBRE_SEDE: s.nombre || (i === 0 ? 'Sede principal' : 'Sede ' + (i + 1)),
      ES_PRINCIPAL: i === 0, DIRECCION: s.direccion, BARRIO: s.barrio, MUNICIPIO: s.municipio, REFERENCIA: s.referencia,
      GOOGLE_MAPS_URL: s.googleMapsUrl, TELEFONO: s.telefono, WHATSAPP: s.whatsapp, HORARIO_SEMANA: s.horarioSemana,
      HORARIO_FIN_SEMANA: s.horarioFinSemana, HORARIOS_ESPECIALES: s.horariosEspeciales, ATIENDE_EN_SITIO: s.atiendeEnSitio, ORDEN: i + 1
    })).filter((s, i) => i === 0 || s.DIRECCION || s.NOMBRE_SEDE);
    const contacto = data.contacto;
    const contactos = [{
      CLIENTE_ID: clienteId, SOLICITUD_ID: sid, NOMBRE: data.negocio.nombrePropietario || data.autorizaciones.firmanteNombre,
      CARGO: data.negocio.nombrePropietario ? 'Propietario' : '', ROL: data.negocio.nombrePropietario ? 'PROPIETARIO' : 'OTRO',
      TELEFONO_FIJO: contacto.telefonoFijo, WHATSAPP: contacto.whatsapp, EMAIL: contacto.email, ES_PRINCIPAL: true,
      PREFERENCIA_CANAL: contacto.canalPreferido, HORARIO_CONTACTO: contacto.horarioContacto, REDES_SOCIALES: contacto.redes
    }];
    const contentValues = {
      nombre_pagina_1: data.negocio.nombrePagina1, nombre_pagina_2: data.negocio.nombrePagina2,
      tipo_establecimiento_alt: data.negocio.tipoEstablecimientoAlt, mensaje: data.marca.mensaje,
      publico_objetivo: data.marca.publicoObjetivo, diferenciales: data.marca.diferenciales, colores: data.marca.colores,
      referencias: data.marca.referencias, dominio_actual: data.marca.dominioActual,
      domicilios_ofrece: data.domicilios.ofrece, domicilios_costo: data.domicilios.costo,
      agenda_citas: data.adicionales.agendaCitas, agenda_version_pago: data.adicionales.agendaVersionPago,
      bd_clientes: data.adicionales.bdClientes, bd_version_pago: data.adicionales.bdVersionPago
    };
    // Si el cliente ya tenía contenido, el nuevo reemplaza al anterior (se conserva en auditoría).
    const previousContent = SheetService.findAllBy('CONTENIDO_WEB', 'CLIENTE_ID', clienteId);
    const contenido = Object.keys(contentValues).filter(k => contentValues[k]).map(k => ({
      CLIENTE_ID: clienteId, SOLICITUD_ID: sid, CAMPO_CLAVE: k, VALOR: contentValues[k]
    }));
    previousContent.filter(p => contenido.some(n => n.CAMPO_CLAVE === p.CAMPO_CLAVE)).forEach(p => {
      SheetService.softDelete('CONTENIDO_WEB', p.CONTENIDO_ID);
    });
    const now = Utils.nowIso();
    const autorizaciones = ENUMS.AUT_TIPO.map(t => ({
      CLIENTE_ID: clienteId, SOLICITUD_ID: sid, TIPO: t, OTORGADA: !!data.autorizaciones[t],
      TEXTO_VERSION: CONFIG.AUTORIZACION_VERSION, TEXTO_HASH: Utils.sha256Hex(authTextFor_(t)), FECHA_HORA: now,
      FIRMANTE_NOMBRE: data.autorizaciones.firmanteNombre, USER_AGENT: Utils.truncate(meta.ua || sol.USER_AGENT || '', 300)
    }));

    const strip = (rows) => rows.map(r => { const o = Object.assign({}, r); delete o._ref; return o; });
    [['CONTACTOS', contactos], ['SEDES', sedes], ['SERVICIOS', strip(servicios)], ['PRODUCTOS', strip(productos)],
      ['TESTIMONIOS', strip(testimonios)], ['CONTENIDO_WEB', contenido], ['AUTORIZACIONES', autorizaciones]].forEach(pair => {
      if (!pair[1].length) return;
      SheetService.insertMany(pair[0], pair[1]);
      track(pair[0], pair[1]);
    });

    // Asociar archivos al cliente y a sus ítems.
    const itemId = (list, key, ref) => { const it = list.find(x => x._ref === ref); return it ? it[key] : ''; };
    files.forEach(f => {
      const patch = { CLIENTE_ID: clienteId };
      if (f.CATEGORIA === 'FOTO_SERVICIO') patch.ENTIDAD_ID = itemId(servicios, 'SERVICIO_ID', f.TEMP_REF);
      if (f.CATEGORIA === 'FOTO_PRODUCTO') patch.ENTIDAD_ID = itemId(productos, 'PRODUCTO_ID', f.TEMP_REF);
      if (f.CATEGORIA === 'FOTO_TESTIMONIO') patch.ENTIDAD_ID = itemId(testimonios, 'TESTIMONIO_ID', f.TEMP_REF);
      SheetService.update('ARCHIVOS', f.ARCHIVO_ID, patch);
    });

    // 4) Snapshot congelado en Drive.
    const snapshot = { radicado: sol.RADICADO, enviadaEn: sol.ENVIADA_EN, clienteId: clienteId, vinculacion: sol.TIPO_VINCULACION, datos: data, correccionesPropuestas: pending };
    const snapshotId = DriveService.writeJson(folderId, 'SOLICITUD_' + sol.RADICADO + '.json', snapshot);

    // 5) CRM: oportunidad, actividad y seguimiento (información interna, separada de la pública).
    CRMService.onFormSubmitted(cliente, sol, isNew, created);

    // 6) Cerrar la solicitud.
    SheetService.update('SOLICITUDES', sid, {
      CLIENTE_ID: clienteId, DATOS_JSON: JSON.stringify(data).length <= CONFIG.DRAFT_JSON_MAX_CHARS ? JSON.stringify(data) : '{}',
      CAMBIOS_PROPUESTOS_JSON: pending.length ? JSON.stringify(pending) : '', CAMBIOS_ESTADO: pending.length ? 'PENDIENTE' : 'NINGUNO',
      SNAPSHOT_FILE_ID: snapshotId, REPORTE_ESTADO: 'PENDIENTE', PORCENTAJE: 100, DRIVE_FOLDER_ID: folderId,
      USER_AGENT: Utils.truncate(meta.ua || sol.USER_AGENT || '', 300)
    });

    if (isNew) ClientService.detectDuplicates(cliente);
    SearchIndex.upsert(ClientService.get(clienteId));
    return { clienteId: clienteId };
  }
};

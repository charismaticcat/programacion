/**
 * ReportService.gs — Reporte editable por cliente (D13).
 *
 * Google Doc ordenado como las secciones de una página web, pensado para copiar
 * y pegar en el software de diseño (aún no definido, D16). Junto al Doc se guarda
 * un JSON con la misma estructura. Solo usa datos PÚBLICOS del cliente: nunca
 * notas internas, oportunidades ni ventas.
 *
 * Regenerar nunca sobrescribe: crea una versión nueva. Si la anterior fue
 * editada a mano, se exige confirmación (force).
 */
const ReportService = {
  MAX_INLINE_IMAGES: 12,
  MAX_INLINE_IMAGE_BYTES: 3 * 1024 * 1024,

  /** Reúne los datos públicos del cliente en la estructura del reporte. */
  collect(clienteId) {
    const c = ClientService.get(clienteId);
    if (!c) throw new AppError('NOT_FOUND', 'Cliente no encontrado.');
    const by = (t) => SheetService.findAllBy(t, 'CLIENTE_ID', clienteId);
    const contactos = by('CONTACTOS');
    const principal = contactos.find(x => x.ES_PRINCIPAL) || contactos[0] || {};
    const contenido = {};
    by('CONTENIDO_WEB').sort((a, b) => String(a.ACTUALIZADO_EN).localeCompare(String(b.ACTUALIZADO_EN)))
      .forEach(x => { contenido[x.CAMPO_CLAVE] = x.VALOR; });
    const archivos = by('ARCHIVOS').filter(a => a.ESTADO === 'APROBADO' || a.ESTADO === 'PENDIENTE');
    const autorizaciones = {};
    by('AUTORIZACIONES').sort((a, b) => String(a.FECHA_HORA).localeCompare(String(b.FECHA_HORA)))
      .forEach(a => { autorizaciones[a.TIPO] = a.OTORGADA && !a.REVOCADA_EN; });
    const fileById = {};
    archivos.forEach(a => { fileById[a.ARCHIVO_ID] = a; });
    const img = (id) => (id && fileById[id] ? { archivoId: id, nombre: fileById[id].NOMBRE_ORIGINAL, url: DriveService.fileUrl(fileById[id].DRIVE_FILE_ID) } : null);
    const solicitudes = SheetService.findAllBy('SOLICITUDES', 'CLIENTE_ID', clienteId).filter(s => s.ESTADO !== 'BORRADOR')
      .sort((a, b) => String(b.ENVIADA_EN).localeCompare(String(a.ENVIADA_EN)));
    const allowImages = !!autorizaciones.USO_IMAGENES;
    return {
      ficha: {
        clienteId: c.CLIENTE_ID, radicado: solicitudes.length ? solicitudes[0].RADICADO : '',
        fecha: Utils.today(), carpetaDrive: DriveService.folderUrl(c.DRIVE_FOLDER_ID),
        nombrePagina1: contenido.nombre_pagina_1 || '', nombrePagina2: contenido.nombre_pagina_2 || ''
      },
      encabezado: {
        nombre: c.NOMBRE_COMERCIAL || c.RAZON_SOCIAL, razonSocial: c.RAZON_SOCIAL,
        logo: img((archivos.find(a => a.CATEGORIA === 'LOGO') || {}).ARCHIVO_ID)
      },
      portada: {
        mensaje: contenido.mensaje || '', tipoEstablecimiento: c.TIPO_ESTABLECIMIENTO,
        whatsappUrl: principal.WHATSAPP ? 'https://wa.me/57' + principal.WHATSAPP : ''
      },
      quienesSomos: {
        propietario: c.NOMBRE_PROPIETARIO, tipoEstablecimiento: c.TIPO_ESTABLECIMIENTO,
        tipoEstablecimientoAlt: contenido.tipo_establecimiento_alt || '', actividad: c.ACTIVIDAD,
        publicoObjetivo: contenido.publico_objetivo || '', diferenciales: contenido.diferenciales || '',
        fotosEquipo: archivos.filter(a => a.CATEGORIA === 'FOTO_EQUIPO').map(a => img(a.ARCHIVO_ID))
      },
      servicios: by('SERVICIOS').filter(s => s.ACTIVO !== false).sort((a, b) => a.ORDEN - b.ORDEN).map(s => ({
        nombre: s.NOMBRE, descripcion: s.DESCRIPCION, precio: s.MOSTRAR_PRECIO ? Utils.formatPriceText(s.PRECIO) : '',
        precioInterno: Utils.formatPriceText(s.PRECIO), requiereCita: s.REQUIERE_CITA, imagen: img(s.IMAGEN_ARCHIVO_ID)
      })),
      catalogo: {
        productos: by('PRODUCTOS').sort((a, b) => a.ORDEN - b.ORDEN).map(p => ({
          nombre: p.NOMBRE, descripcion: p.DESCRIPCION, categoria: p.CATEGORIA,
          precio: p.MOSTRAR_PRECIO ? Utils.formatPriceText(p.PRECIO) : '', precioInterno: Utils.formatPriceText(p.PRECIO),
          disponible: p.DISPONIBLE, imagen: img(p.IMAGEN_ARCHIVO_ID)
        })),
        domicilios: contenido.domicilios_ofrece || '', costoDomicilio: Utils.formatPriceText(contenido.domicilios_costo || '')
      },
      testimonios: by('TESTIMONIOS').map(t => ({
        autor: t.PERMISO_USO_NOMBRE ? t.AUTOR_NOMBRE : 'Cliente', texto: t.TEXTO,
        foto: t.PERMISO_USO_FOTO ? img(t.FOTO_ARCHIVO_ID) : null, aprobadoAdmin: t.APROBADO_ADMIN,
        publicable: !!autorizaciones.USO_FOTOS_TESTIMONIOS || !t.FOTO_ARCHIVO_ID
      })),
      galeria: archivos.filter(a => a.CATEGORIA === 'FOTO_NEGOCIO').map(a => img(a.ARCHIVO_ID)),
      ubicacion: by('SEDES').sort((a, b) => a.ORDEN - b.ORDEN).map(s => ({
        nombre: s.NOMBRE_SEDE, principal: s.ES_PRINCIPAL, direccion: s.DIRECCION, barrio: s.BARRIO, municipio: s.MUNICIPIO,
        referencia: s.REFERENCIA, mapa: s.GOOGLE_MAPS_URL, telefono: s.TELEFONO, whatsapp: s.WHATSAPP,
        horarioSemana: s.HORARIO_SEMANA, horarioFinSemana: s.HORARIO_FIN_SEMANA, horariosEspeciales: s.HORARIOS_ESPECIALES,
        atiendeEnSitio: s.ATIENDE_EN_SITIO
      })),
      contacto: {
        telefonoFijo: principal.TELEFONO_FIJO || '', whatsapp: principal.WHATSAPP || '', email: principal.EMAIL || c.EMAIL,
        redes: principal.REDES_SOCIALES || '', colores: contenido.colores || '', referencias: contenido.referencias || '',
        dominioActual: contenido.dominio_actual || ''
      },
      funcionesAdicionales: {
        agendaCitas: contenido.agenda_citas || '', agendaVersionPago: contenido.agenda_version_pago || '',
        baseDatosClientes: contenido.bd_clientes || '', baseDatosVersionPago: contenido.bd_version_pago || ''
      },
      permisos: {
        usoImagenes: !!autorizaciones.USO_IMAGENES, publicacionContenido: !!autorizaciones.PUBLICACION_CONTENIDO,
        testimonios: !!autorizaciones.USO_FOTOS_TESTIMONIOS, tratamientoDatos: !!autorizaciones.TRATAMIENTO_DATOS
      },
      _allowImages: allowImages,
      _cliente: c
    };
  },

  /**
   * Genera una nueva versión del reporte.
   * @param {Object|null} user usuario admin (null = activador automático)
   */
  generate(user, clienteId, solicitudId, force) {
    const c = user ? ClientService.requireAccessible(user, clienteId) : ClientService.get(clienteId);
    if (!c) throw new AppError('NOT_FOUND', 'Cliente no encontrado.');
    const previous = SheetService.findAllBy('REPORTES', 'CLIENTE_ID', clienteId).sort((a, b) => b.VERSION_REPORTE - a.VERSION_REPORTE);
    const last = previous.find(r => r.ESTADO === 'GENERADO');
    if (last && !force && user) {
      try {
        const updated = DriveApp.getFileById(last.DOC_ID).getLastUpdated();
        const genAt = Utils.parseIsoDate(last.GENERADO_EN);
        if (genAt && updated.getTime() - genAt.getTime() > 120000) {
          throw new AppError('REPORT_EDITED', 'La versión ' + last.VERSION_REPORTE + ' fue editada a mano. Se creará una versión nueva y la anterior se conservará. ¿Continuar?');
        }
      } catch (e) {
        if (e instanceof AppError) throw e;
      }
    }
    const version = previous.length ? Number(previous[0].VERSION_REPORTE) + 1 : 1;
    const data = ReportService.collect(clienteId);
    const folderId = ClientService.ensureClientFolder(ClientService.get(clienteId));
    const docName = 'REPORTE_' + clienteId + '_v' + version + '_' + Utils.slug(data.encabezado.nombre, 30);
    const doc = DocumentApp.create(docName);
    // ID y URL se leen antes de cerrar: un Document cerrado ya no se puede consultar.
    const docId = doc.getId();
    const docUrl = doc.getUrl();
    const docFile = DriveApp.getFileById(docId);
    docFile.moveTo(DriveService.folderForCategory(folderId, 'REPORTE'));
    ReportService._render(doc, data, version);
    doc.saveAndClose();

    const json = Object.assign({}, data);
    delete json._cliente;
    delete json._allowImages;
    json.version = version;
    const jsonId = DriveService.writeJson(folderId, 'DATOS_' + clienteId + '_v' + version + '.json', json);

    Settings.get('REPORT_EDITORS').forEach(email => {
      try { docFile.addEditor(email); } catch (e) { console.warn('No se pudo compartir con', email, e && e.message); }
    });

    const rec = {
      CLIENTE_ID: clienteId, SOLICITUD_ID: solicitudId || '', VERSION_REPORTE: version, DOC_ID: docId,
      DOC_URL: docUrl, JSON_FILE_ID: jsonId, GENERADO_POR: user ? 'ADMIN:' + user.email : 'SISTEMA',
      GENERADO_EN: Utils.nowIso(), ESTADO: 'GENERADO', ERROR: ''
    };
    SheetService.insert('REPORTES', rec);
    previous.filter(r => r.ESTADO === 'GENERADO').forEach(r => SheetService.update('REPORTES', r.REPORTE_ID, { ESTADO: 'REEMPLAZADO' }));
    Audit.log({ accion: 'GENERAR_REPORTE', entidad: 'REPORTES', entidadId: rec.REPORTE_ID, clienteId: clienteId, detalle: 'Versión ' + version });
    return ClientService._strip(rec);
  },

  /** Construye el documento. */
  _render(doc, d, version) {
    const body = doc.getBody();
    body.clear();
    body.setMarginTop(48).setMarginBottom(48).setMarginLeft(56).setMarginRight(56);
    const H = DocumentApp.ParagraphHeading;
    const title = body.appendParagraph(d.encabezado.nombre || 'Cliente');
    title.setHeading(H.TITLE);
    body.appendParagraph('Reporte para diseño web · versión ' + version + ' · ' + d.ficha.fecha)
      .setHeading(H.SUBTITLE);
    body.appendParagraph('Cada bloque está listo para copiar y pegar en la sección correspondiente de la página. ' +
      'Los campos vacíos se omiten. Las imágenes originales están en la carpeta de Drive del cliente.')
      .editAsText().setItalic(true);

    const section = (text) => body.appendParagraph(text).setHeading(H.HEADING1);
    const sub = (text) => body.appendParagraph(text).setHeading(H.HEADING2);
    const field = (label, value) => {
      if (value === undefined || value === null || value === '' || value === false) return;
      const v = value === true ? 'Sí' : String(value);
      body.appendParagraph(label.toUpperCase()).editAsText().setBold(true).setFontSize(9);
      body.appendParagraph(v).editAsText().setBold(false).setFontSize(11);
    };
    const link = (label, url) => {
      if (!url) return;
      const p = body.appendParagraph(label);
      p.setLinkUrl(url);
    };
    const image = (ref, width) => {
      if (!ref) return;
      let inserted = false;
      if (ReportService._imageCount < ReportService.MAX_INLINE_IMAGES) {
        try {
          const f = SheetService.getById('ARCHIVOS', ref.archivoId);
          if (f && f.MIME_TYPE.indexOf('image/') === 0 && Number(f.TAMANO_BYTES) <= ReportService.MAX_INLINE_IMAGE_BYTES) {
            const blob = DriveApp.getFileById(f.DRIVE_FILE_ID).getBlob();
            const im = body.appendParagraph('').appendInlineImage(blob);
            const w = im.getWidth();
            const h = im.getHeight();
            if (w > width) { im.setWidth(width); im.setHeight(Math.round(h * width / w)); }
            ReportService._imageCount++;
            inserted = true;
          }
        } catch (e) {
          console.warn('Imagen no insertada', ref.archivoId, e && e.message);
        }
      }
      link((inserted ? 'Original: ' : 'Imagen: ') + (ref.nombre || ref.archivoId), ref.url);
    };
    ReportService._imageCount = 0;

    section('0 · Ficha de trabajo');
    field('ID cliente', d.ficha.clienteId);
    field('Radicado', d.ficha.radicado);
    field('Nombre de página deseado (opción 1)', d.ficha.nombrePagina1);
    field('Nombre de página deseado (opción 2)', d.ficha.nombrePagina2);
    link('Abrir carpeta de Drive del cliente', d.ficha.carpetaDrive);

    section('1 · Encabezado / menú');
    field('Nombre a mostrar', d.encabezado.nombre);
    if (d.encabezado.razonSocial !== d.encabezado.nombre) field('Razón social', d.encabezado.razonSocial);
    if (d.encabezado.logo) { sub('Logo'); image(d.encabezado.logo, 180); }

    section('2 · Portada');
    field('Mensaje principal (qué quiere transmitir)', d.portada.mensaje);
    field('Tipo de establecimiento', d.portada.tipoEstablecimiento);
    field('Botón de WhatsApp', d.portada.whatsappUrl);

    section('3 · Quiénes somos');
    field('Propietario', d.quienesSomos.propietario);
    field('Tipo de establecimiento', d.quienesSomos.tipoEstablecimiento);
    field('Tipo de establecimiento (opción 2)', d.quienesSomos.tipoEstablecimientoAlt);
    field('Actividad', d.quienesSomos.actividad);
    field('Público objetivo', d.quienesSomos.publicoObjetivo);
    field('Qué lo hace diferente', d.quienesSomos.diferenciales);
    if (d.quienesSomos.fotosEquipo.length) {
      sub('Fotos del equipo');
      d.quienesSomos.fotosEquipo.forEach(f => image(f, 220));
    }

    if (d.servicios.length) {
      section('4 · Servicios');
      d.servicios.forEach((s, i) => {
        sub((i + 1) + '. ' + s.nombre);
        field('Descripción', s.descripcion);
        field('Precio (se muestra)', s.precio);
        if (!s.precio && s.precioInterno) field('Precio (no mostrar en la web)', s.precioInterno);
        field('Requiere cita', s.requiereCita);
        image(s.imagen, 200);
      });
    }

    if (d.catalogo.productos.length || d.catalogo.domicilios) {
      section('5 · Catálogo / productos');
      d.catalogo.productos.forEach((p, i) => {
        sub((i + 1) + '. ' + p.nombre);
        field('Categoría', p.categoria);
        field('Descripción', p.descripcion);
        field('Costo (se muestra)', p.precio);
        if (!p.precio && p.precioInterno) field('Costo (no mostrar en la web)', p.precioInterno);
        if (!p.disponible) field('Disponibilidad', 'No disponible por ahora');
        image(p.imagen, 200);
      });
      field('Domicilios', d.catalogo.domicilios === 'SI' ? 'Sí ofrece domicilios' : (d.catalogo.domicilios === 'NO' ? 'No ofrece domicilios' : ''));
      field('Costo del domicilio', d.catalogo.costoDomicilio);
    }

    if (d.testimonios.length) {
      section('6 · Testimonios');
      d.testimonios.forEach(t => {
        body.appendParagraph('“' + t.texto + '”').editAsText().setItalic(true);
        body.appendParagraph('— ' + t.autor);
        if (t.foto) image(t.foto, 120);
        if (!t.aprobadoAdmin) body.appendParagraph('(Pendiente de aprobación interna)').editAsText().setFontSize(9);
      });
    }

    if (d.galeria.length) {
      section('7 · Galería');
      d.galeria.forEach(f => image(f, 240));
    }

    if (d.ubicacion.length) {
      section('8 · Ubicación y horarios');
      d.ubicacion.forEach(s => {
        sub(s.nombre + (s.principal ? ' (principal)' : ''));
        field('Dirección', [s.direccion, s.barrio, s.municipio].filter(Boolean).join(', '));
        field('Punto de referencia', s.referencia);
        field('Google Maps', s.mapa);
        field('Teléfono', s.telefono);
        field('WhatsApp', s.whatsapp);
        field('Horario entre semana', s.horarioSemana);
        field('Horario fin de semana', s.horarioFinSemana);
        field('Horarios especiales', s.horariosEspeciales);
        field('Atiende en sitio', s.atiendeEnSitio);
      });
    }

    section('9 · Contacto / pie de página');
    field('Teléfono fijo', d.contacto.telefonoFijo);
    field('WhatsApp', d.contacto.whatsapp);
    field('Email', d.contacto.email);
    field('Redes sociales', d.contacto.redes);
    field('Colores de marca', d.contacto.colores);
    field('Páginas de referencia', d.contacto.referencias);
    field('Dominio o página actual', d.contacto.dominioActual);

    section('10 · Funciones adicionales');
    field('Agenda de citas', d.funcionesAdicionales.agendaCitas);
    field('Agenda de citas: versión de pago', d.funcionesAdicionales.agendaVersionPago);
    field('Base de datos de clientes', d.funcionesAdicionales.baseDatosClientes);
    field('Base de datos de clientes: versión de pago', d.funcionesAdicionales.baseDatosVersionPago);

    section('11 · Permisos: qué se puede publicar');
    body.appendParagraph((d.permisos.usoImagenes ? '✔' : '✘') + ' Uso de logo y fotografías en la página');
    body.appendParagraph((d.permisos.publicacionContenido ? '✔' : '✘') + ' Publicación de textos, servicios, precios y horarios');
    body.appendParagraph((d.permisos.testimonios ? '✔' : '✘') + ' Publicación de testimonios con nombre/foto de terceros');
    if (!d.permisos.usoImagenes || !d.permisos.publicacionContenido) {
      body.appendParagraph('Atención: falta alguna autorización. Confirme con el cliente antes de publicar.').editAsText().setBold(true);
    }

    section('12 · Notas de Webpaya');
    body.appendParagraph('Espacio libre para el diseñador o el administrador.');
    body.appendParagraph('');
  },

  /** Activador: genera los reportes pendientes (máximo ~4,5 min por ejecución). */
  processPending() {
    const start = Date.now();
    const pending = SheetService.readColumns('SOLICITUDES', ['SOLICITUD_ID', 'CLIENTE_ID', 'REPORTE_ESTADO', 'ESTADO'])
      .filter(s => s.REPORTE_ESTADO === 'PENDIENTE' && s.CLIENTE_ID && s.ESTADO !== 'BORRADOR');
    let done = 0;
    for (let i = 0; i < pending.length; i++) {
      if (Date.now() - start > 240000) break;
      const s = pending[i];
      try {
        ReportService.generate(null, s.CLIENTE_ID, s.SOLICITUD_ID, true);
        SheetService.update('SOLICITUDES', s.SOLICITUD_ID, { REPORTE_ESTADO: 'GENERADO' });
        done++;
      } catch (e) {
        console.error('REPORTE_ERROR', s.SOLICITUD_ID, e && e.message);
        SheetService.update('SOLICITUDES', s.SOLICITUD_ID, { REPORTE_ESTADO: 'ERROR' });
        Audit.log({ accion: 'GENERAR_REPORTE', entidad: 'SOLICITUDES', entidadId: s.SOLICITUD_ID, clienteId: s.CLIENTE_ID, resultado: 'ERROR', detalle: e && e.message });
      }
    }
    return { pendientes: pending.length, generados: done };
  },

  pdf(user, reporteId) {
    const r = SheetService.getById('REPORTES', reporteId);
    if (!r) throw new AppError('NOT_FOUND', 'Reporte no encontrado.');
    ClientService.requireAccessible(user, r.CLIENTE_ID);
    const blob = DriveApp.getFileById(r.DOC_ID).getAs('application/pdf');
    Audit.log({ accion: 'DESCARGAR_REPORTE', entidad: 'REPORTES', entidadId: reporteId, clienteId: r.CLIENTE_ID });
    return { name: 'REPORTE_' + r.CLIENTE_ID + '_v' + r.VERSION_REPORTE + '.pdf', mime: 'application/pdf', base64: Utilities.base64Encode(blob.getBytes()) };
  },

  json(user, reporteId) {
    const r = SheetService.getById('REPORTES', reporteId);
    if (!r) throw new AppError('NOT_FOUND', 'Reporte no encontrado.');
    ClientService.requireAccessible(user, r.CLIENTE_ID);
    const d = DriveService.download(r.JSON_FILE_ID);
    return { name: 'DATOS_' + r.CLIENTE_ID + '_v' + r.VERSION_REPORTE + '.json', mime: 'application/json', base64: d.base64 };
  }
};

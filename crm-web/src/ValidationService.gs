/**
 * ValidationService.gs — Validación y limpieza en el servidor.
 * El navegador valida para ayudar al usuario; aquí se valida de verdad.
 */
const Validation = {
  STEPS: ['negocio', 'contacto', 'sedes', 'marca', 'servicios', 'productos', 'archivos', 'testimonios', 'adicionales', 'autorizaciones'],

  short(v, max) {
    return Utils.truncate(Utils.cleanSpaces(v), max || CONFIG.TEXT_SHORT_MAX);
  },
  long(v, max) {
    const s = String(v == null ? '' : v).replace(/\r\n?/g, '\n').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
    return Utils.truncate(s.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim(), max || CONFIG.TEXT_LONG_MAX);
  },
  bool(v) {
    return v === true || Utils.toBool(v);
  },
  siNo(v) {
    const s = String(v || '').toUpperCase();
    return s === 'SI' || s === 'NO' ? s : '';
  },
  ref(v) {
    const s = String(v || '');
    return /^[A-Za-z0-9-]{1,40}$/.test(s) ? s : 'r' + Utils.randomToken().slice(0, 10);
  },
  municipioName(v) {
    const n = Utils.normalizeText(v);
    const m = CONFIG.MUNICIPIOS.find(x => x.name === n || x.code === String(v || '').trim());
    return m ? m.name : '';
  },
  municipioCode(name) {
    const m = CONFIG.MUNICIPIOS.find(x => x.name === name);
    return m ? m.code : '';
  },
  url(v) {
    const s = Validation.short(v, 500);
    if (!s) return '';
    return /^https?:\/\/[^\s"'<>]+$/i.test(s) ? s : null;
  },

  /**
   * Limpia los datos del formulario público.
   * @param {Object} raw datos del navegador
   * @param {boolean} forSubmit aplica campos obligatorios
   * @return {{data:Object, errors:Array<{step,field,message}>}}
   */
  cleanForm(raw, forSubmit) {
    raw = Utils.isPlainObject(raw) ? raw : {};
    const errors = [];
    const err = (step, field, message) => errors.push({ step: step, field: field, message: message });
    const S = Validation.short;
    const L = Validation.long;
    const B = Validation.bool;

    const n = Utils.isPlainObject(raw.negocio) ? raw.negocio : {};
    const negocio = {
      razonSocial: S(n.razonSocial),
      nombrePropietario: S(n.nombrePropietario, 150),
      nombreComercial: S(n.nombreComercial),
      tipoEstablecimiento: S(n.tipoEstablecimiento, 120),
      tipoEstablecimientoAlt: S(n.tipoEstablecimientoAlt, 120),
      actividad: L(n.actividad, 1000),
      nombrePagina1: S(n.nombrePagina1, 120),
      nombrePagina2: S(n.nombrePagina2, 120)
    };

    const c = Utils.isPlainObject(raw.contacto) ? raw.contacto : {};
    const phoneField = (value, step, field, label) => {
      const d = Utils.digits(value);
      if (!d) return '';
      const p = Utils.normalizePhone(d);
      if (!p.value) {
        err(step, field, label + ': escriba 7 dígitos (fijo) o 10 dígitos (celular).');
        return '';
      }
      return p.value;
    };
    const email = Utils.normalizeEmail(c.email);
    if (email && !Utils.isValidEmail(email)) err('contacto', 'email', 'El email no es válido.');
    const contacto = {
      telefonoFijo: phoneField(c.telefonoFijo, 'contacto', 'telefonoFijo', 'Teléfono fijo'),
      whatsapp: phoneField(c.whatsapp, 'contacto', 'whatsapp', 'WhatsApp'),
      email: email && Utils.isValidEmail(email) ? Utils.truncate(email, 150) : '',
      canalPreferido: ENUMS.CANAL.indexOf(String(c.canalPreferido || '')) >= 0 ? c.canalPreferido : '',
      horarioContacto: S(c.horarioContacto, 120),
      redes: L(c.redes, 1000)
    };

    const sedesRaw = Array.isArray(raw.sedes) ? raw.sedes.slice(0, CONFIG.MAX_SEDES) : [];
    const sedes = [];
    sedesRaw.forEach((s, i) => {
      if (!Utils.isPlainObject(s)) return;
      const municipioRaw = S(s.municipio, 60);
      const municipio = Validation.municipioName(municipioRaw);
      if (municipioRaw && !municipio) err('sedes', 'sedes.' + i + '.municipio', 'Sede ' + (i + 1) + ': el municipio debe ser Neiva, Pitalito, Palermo o Rivera.');
      const maps = Validation.url(s.googleMapsUrl);
      if (maps === null) err('sedes', 'sedes.' + i + '.googleMapsUrl', 'Sede ' + (i + 1) + ': el enlace de Google Maps debe empezar por https://');
      const sede = {
        ref: Validation.ref(s.ref),
        nombre: S(s.nombre, 120),
        direccion: S(s.direccion, 200),
        barrio: S(s.barrio, 120),
        municipio: municipio,
        referencia: S(s.referencia, 200),
        googleMapsUrl: maps || '',
        telefono: phoneField(s.telefono, 'sedes', 'sedes.' + i + '.telefono', 'Sede ' + (i + 1) + ' teléfono'),
        whatsapp: phoneField(s.whatsapp, 'sedes', 'sedes.' + i + '.whatsapp', 'Sede ' + (i + 1) + ' WhatsApp'),
        horarioSemana: S(s.horarioSemana, 200),
        horarioFinSemana: S(s.horarioFinSemana, 200),
        horariosEspeciales: L(s.horariosEspeciales, 1000),
        atiendeEnSitio: B(s.atiendeEnSitio)
      };
      const hasContent = ['nombre', 'direccion', 'barrio', 'municipio', 'referencia', 'googleMapsUrl', 'telefono', 'whatsapp',
        'horarioSemana', 'horarioFinSemana', 'horariosEspeciales'].some(k => sede[k]);
      if (hasContent || i === 0) sedes.push(sede);
    });

    const m = Utils.isPlainObject(raw.marca) ? raw.marca : {};
    const marca = {
      mensaje: L(m.mensaje, 2000),
      publicoObjetivo: L(m.publicoObjetivo, 1000),
      diferenciales: L(m.diferenciales, 1000),
      colores: S(m.colores, 200),
      referencias: L(m.referencias, 1000),
      dominioActual: S(m.dominioActual, 200)
    };

    const items = (list, max, step, label, mapper) => {
      const arr = Array.isArray(list) ? list : [];
      if (arr.length > max) err(step, step, 'Máximo ' + max + ' ' + label + '.');
      const out = [];
      arr.slice(0, max).forEach((it, i) => {
        if (!Utils.isPlainObject(it)) return;
        const r = mapper(it, i);
        if (r) out.push(r);
      });
      return out;
    };

    const servicios = items(raw.servicios, CONFIG.MAX_SERVICIOS, 'servicios', 'servicios', (s, i) => {
      const r = {
        ref: Validation.ref(s.ref), nombre: S(s.nombre, 150), descripcion: L(s.descripcion, 1500),
        precio: S(s.precio, 60), mostrarPrecio: B(s.mostrarPrecio), requiereCita: B(s.requiereCita)
      };
      if (!r.nombre && !r.descripcion && !r.precio) return null;
      if (!r.nombre) err('servicios', 'servicios.' + i + '.nombre', 'Servicio ' + (i + 1) + ': escriba el nombre.');
      return r;
    });

    const productos = items(raw.productos, CONFIG.MAX_PRODUCTOS, 'productos', 'productos', (p, i) => {
      const r = {
        ref: Validation.ref(p.ref), nombre: S(p.nombre, 150), descripcion: L(p.descripcion, 1500),
        categoria: S(p.categoria, 100), precio: S(p.precio, 60), mostrarPrecio: B(p.mostrarPrecio),
        disponible: p.disponible === undefined ? true : B(p.disponible)
      };
      if (!r.nombre && !r.descripcion && !r.precio && !r.categoria) return null;
      if (!r.nombre) err('productos', 'productos.' + i + '.nombre', 'Producto ' + (i + 1) + ': escriba el nombre.');
      return r;
    });

    const d = Utils.isPlainObject(raw.domicilios) ? raw.domicilios : {};
    const domicilios = { ofrece: Validation.siNo(d.ofrece), costo: S(d.costo, 60) };

    const testimonios = items(raw.testimonios, CONFIG.MAX_TESTIMONIOS, 'testimonios', 'testimonios', (t, i) => {
      const r = {
        ref: Validation.ref(t.ref), autor: S(t.autor, 120), texto: L(t.texto, 400),
        permisoFoto: B(t.permisoFoto), permisoNombre: B(t.permisoNombre)
      };
      if (!r.autor && !r.texto) return null;
      if (!r.texto) err('testimonios', 'testimonios.' + i + '.texto', 'Testimonio ' + (i + 1) + ': escriba el texto.');
      if (Utils.wordCount(r.texto) > CONFIG.TESTIMONIO_MAX_PALABRAS) {
        err('testimonios', 'testimonios.' + i + '.texto', 'Testimonio ' + (i + 1) + ': máximo ' + CONFIG.TESTIMONIO_MAX_PALABRAS + ' palabras.');
      }
      return r;
    });

    const a = Utils.isPlainObject(raw.adicionales) ? raw.adicionales : {};
    const adicionales = {
      agendaCitas: Validation.siNo(a.agendaCitas),
      agendaVersionPago: Validation.siNo(a.agendaVersionPago),
      bdClientes: Validation.siNo(a.bdClientes),
      bdVersionPago: Validation.siNo(a.bdVersionPago)
    };

    const au = Utils.isPlainObject(raw.autorizaciones) ? raw.autorizaciones : {};
    const autorizaciones = { firmanteNombre: S(au.firmanteNombre, 150) };
    ENUMS.AUT_TIPO.forEach(t => { autorizaciones[t] = B(au[t]); });

    const data = {
      negocio: negocio, contacto: contacto, sedes: sedes, marca: marca, servicios: servicios,
      productos: productos, domicilios: domicilios, testimonios: testimonios,
      adicionales: adicionales, autorizaciones: autorizaciones
    };

    if (forSubmit) {
      const required = Settings.get('REQUIRED_FIELDS');
      const sede0 = sedes[0] || {};
      if (required.indexOf('razonSocial') >= 0 && !negocio.razonSocial) {
        err('negocio', 'razonSocial', 'Escriba la razón social o el nombre del establecimiento.');
      }
      if (required.indexOf('medioContacto') >= 0 && !contacto.whatsapp && !contacto.telefonoFijo) {
        err('contacto', 'whatsapp', 'Escriba al menos un teléfono: WhatsApp o fijo.');
      }
      if (required.indexOf('nombrePropietario') >= 0 && !negocio.nombrePropietario) {
        err('negocio', 'nombrePropietario', 'Escriba el nombre del propietario.');
      }
      if (required.indexOf('email') >= 0 && !contacto.email) err('contacto', 'email', 'Escriba un email.');
      if (required.indexOf('municipio') >= 0 && !sede0.municipio) err('sedes', 'sedes.0.municipio', 'Seleccione el municipio.');
      if (required.indexOf('direccion') >= 0 && !sede0.direccion) err('sedes', 'sedes.0.direccion', 'Escriba la dirección.');
      // La autorización de tratamiento de datos es obligatoria por ley (siempre).
      if (!autorizaciones.TRATAMIENTO_DATOS) {
        err('autorizaciones', 'TRATAMIENTO_DATOS', 'Debe aceptar el tratamiento de datos personales para enviar.');
      }
      if (!autorizaciones.firmanteNombre) {
        err('autorizaciones', 'firmanteNombre', 'Escriba el nombre de quien autoriza.');
      }
    }
    return { data: data, errors: errors };
  },

  /** Porcentaje aproximado de avance del formulario. */
  progress(data) {
    const checks = [
      !!data.negocio.razonSocial, !!data.negocio.tipoEstablecimiento, !!(data.contacto.whatsapp || data.contacto.telefonoFijo),
      !!(data.sedes[0] && data.sedes[0].direccion), !!data.marca.mensaje, data.servicios.length > 0 || data.productos.length > 0,
      data.testimonios.length > 0, !!data.adicionales.agendaCitas, !!data.autorizaciones.TRATAMIENTO_DATOS
    ];
    return Math.round(100 * checks.filter(Boolean).length / checks.length);
  },

  /**
   * Limpia un parche de edición del admin para una tabla: solo columnas editables,
   * tipos correctos, listas válidas, teléfonos y emails normalizados.
   */
  cleanRecord(table, patch) {
    if (!Utils.isPlainObject(patch)) throw new AppError('VALIDATION', 'Datos inválidos.');
    const out = {};
    const errors = [];
    SchemaUtil.columns(table).forEach(col => {
      if (!col.editable || !Object.prototype.hasOwnProperty.call(patch, col.name)) return;
      let v = patch[col.name];
      const label = col.label;
      switch (col.type) {
        case 'n': {
          if (v === '' || v === null) { v = 0; break; }
          const num = Number(v);
          if (!isFinite(num)) { errors.push(label + ': número no válido.'); return; }
          v = num;
          break;
        }
        case 'm': {
          const num = Utils.parseMoney(v);
          if (num > 999999999999) { errors.push(label + ': valor demasiado grande.'); return; }
          v = num;
          break;
        }
        case 'b': v = Validation.bool(v); break;
        case 'd':
          v = Validation.short(v, 10);
          if (v && !Utils.isIsoDate(v)) { errors.push(label + ': fecha no válida (aaaa-mm-dd).'); return; }
          break;
        case 'dt':
          v = Validation.short(v, 19);
          if (v && !Utils.isIsoDateTime(v)) { errors.push(label + ': fecha y hora no válidas.'); return; }
          if (v && v.length === 16) v += ':00';
          break;
        case 't': v = Validation.long(v, 5000); break;
        case 'e': {
          v = Validation.short(v, 100);
          if (col.enumKey === 'cat:MUNICIPIOS') {
            const mun = Validation.municipioName(v);
            if (v && !mun) { errors.push(label + ': solo Neiva, Pitalito, Palermo o Rivera.'); return; }
            v = mun;
          } else if (col.enumKey.indexOf('ref:') === 0) {
            // referencia a otra tabla: se valida en el servicio correspondiente
          } else {
            const allowed = ENUMS[col.enumKey] || [];
            if (v && allowed.indexOf(v) < 0) { errors.push(label + ': valor no permitido.'); return; }
          }
          break;
        }
        default:
          v = Validation.short(v, 500);
      }
      if (/^(TELEFONO|TELEFONO_FIJO|WHATSAPP)$/.test(col.name) && v) {
        const p = Utils.normalizePhone(v);
        if (!p.value) { errors.push(label + ': escriba 7 o 10 dígitos.'); return; }
        v = p.value;
      }
      if (col.name === 'EMAIL' && v) {
        v = Utils.normalizeEmail(v);
        if (!Utils.isValidEmail(v)) { errors.push(label + ': email no válido.'); return; }
      }
      if (col.name === 'GOOGLE_MAPS_URL' && v && Validation.url(v) === null) {
        errors.push(label + ': debe empezar por https://');
        return;
      }
      out[col.name] = v;
    });
    if (errors.length) throw new AppError('VALIDATION', errors.join(' '), errors);
    return out;
  },

  requireText(v, label, max) {
    const s = Validation.short(v, max);
    if (!s) throw new AppError('VALIDATION', 'Falta: ' + label + '.');
    return s;
  },
  requireEnum(v, enumKey, label) {
    if ((ENUMS[enumKey] || []).indexOf(v) < 0) throw new AppError('VALIDATION', label + ': valor no permitido.');
    return v;
  },
  requireId(v, prefix) {
    const s = String(v || '');
    if (!/^[A-Z]{3}-[A-Z0-9]{4,20}$/.test(s) || (prefix && s.indexOf(prefix + '-') !== 0)) {
      throw new AppError('VALIDATION', 'Identificador no válido.');
    }
    return s;
  }
};

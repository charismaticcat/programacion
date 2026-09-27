/**
 * Batería de pruebas automáticas (Parte M).
 * Ejecuta el código real de src/*.gs sobre servicios de Google simulados.
 *   node tests/run-tests.js
 */
'use strict';
const assert = require('assert');
const { loadProject } = require('./harness/load-gas');
const F = require('./fixtures');

// ------------------------------------------------------------------ mini runner
const results = [];
let current = '';
function suite(name, fn) { current = name; fn(); }
function test(name, fn) {
  try {
    fn();
    results.push({ ok: true, name: current + ' › ' + name });
  } catch (e) {
    results.push({ ok: false, name: current + ' › ' + name, err: e });
  }
}

// ------------------------------------------------------------------ helpers
const SUPER = { email: 'admin@webpaya.test', pass: 'ClaveSegura2026' };

function freshSystem() {
  const p = loadProject({ echoErrors: false });
  p.asOwner();
  const log = p.call('setup');
  const codeLine = log.find(l => l.indexOf('CÓDIGO DE CONFIGURACIÓN') >= 0);
  const code = codeLine.match(/: ([0-9A-F-]{14}) /)[1];
  // Base existente simulada
  const ss = p.env.globals.SpreadsheetApp.create('CCH_PRUEBA');
  const sh = ss.getSheetByName('Hoja 1');
  sh.name = 'Hoja1';
  [F.HEADER].concat(F.ROWS).forEach((r, i) => { sh.rows[i] = r.slice(); });
  p.env.state.props.IMPORT_SOURCE_SPREADSHEET_ID = ss.getId();
  p.call('importBase');
  p.asAnonymous();
  const created = p.call('adm_createSuperAdmin', code, SUPER.email, 'Admin Prueba', SUPER.pass);
  assert.ok(created.ok, JSON.stringify(created));
  const login = p.call('adm_login', SUPER.email, SUPER.pass, 'test');
  assert.ok(login.ok, JSON.stringify(login));
  return { p, token: login.data.token, code };
}
const ok = (res) => { assert.ok(res && res.ok, 'Se esperaba ok: ' + JSON.stringify(res && res.error)); return res.data; };
const err = (res, code) => {
  assert.ok(res && !res.ok, 'Se esperaba error ' + code + ' y fue ok');
  if (code) assert.strictEqual(res.error.code, code, 'Código esperado ' + code + ' y fue ' + res.error.code + ': ' + res.error.message);
  return res.error;
};
function backdate(p, token) {
  // Simula que el usuario lleva más de 20 s en el formulario.
  p.run(`(function(){ const s = Auth.requireDraft(${JSON.stringify(token)}); SheetService.update('SOLICITUDES', s.SOLICITUD_ID, { INICIADA_EN: '2020-01-01T00:00:00' }); })()`);
}
function fullForm(over) {
  const d = {
    negocio: { razonSocial: 'CAFE DE PRUEBA NUEVO', nombrePropietario: 'Persona Prueba', nombreComercial: 'Café Prueba', tipoEstablecimiento: 'Cafetería', nombrePagina1: 'cafeprueba' },
    contacto: { whatsapp: '3001112233', telefonoFijo: '8765432', email: 'nuevo@example.com', canalPreferido: 'WHATSAPP' },
    sedes: [{ ref: 'sede1', nombre: 'Principal', direccion: 'Calle 10 # 1-1', barrio: 'Centro', municipio: 'NEIVA', horarioSemana: '8-6' }],
    marca: { mensaje: 'Café de origen del Huila' },
    servicios: [{ ref: 'srv1', nombre: 'Catas', precio: '25000', mostrarPrecio: true }],
    productos: [{ ref: 'prd1', nombre: 'Café molido', precio: 'Desde $20.000', mostrarPrecio: true }],
    domicilios: { ofrece: 'SI', costo: '5000' },
    testimonios: [{ ref: 'tes1', autor: 'Cliente Prueba', texto: 'Excelente café', permisoNombre: true }],
    adicionales: { agendaCitas: 'NO', bdClientes: 'SI', bdVersionPago: 'NO' },
    autorizaciones: { TRATAMIENTO_DATOS: true, USO_IMAGENES: true, PUBLICACION_CONTENIDO: true, firmanteNombre: 'Persona Prueba' }
  };
  return Object.assign(d, over || {});
}
const count = (p, table, pred) => p.rows(table).filter(pred || (() => true)).length;
const bigBase64 = (bytes) => Buffer.concat([F.JPEG_HEAD, Buffer.alloc(bytes)]).toString('base64');

// ================================================================== PÚBLICO
suite('Público', () => {
  const { p } = freshSystem();

  test('bootstrap no expone IDs ni configuración privada', () => {
    const b = ok(p.call('pub_bootstrap'));
    assert.deepStrictEqual(b.municipios, ['NEIVA', 'PITALITO', 'PALERMO', 'RIVERA']);
    const json = JSON.stringify(b);
    Object.values(p.env.state.props).forEach(v => { if (String(v).length > 10) assert.ok(json.indexOf(v) < 0, 'expone propiedad privada'); });
  });

  test('honeypot bloquea bots', () => { err(p.call('pub_start', { hp: 'http://spam' }), 'BLOCKED'); });

  test('cliente existente: búsqueda por teléfono y precarga completa', () => {
    const { token } = ok(p.call('pub_start', { ua: 'test' }));
    const r = ok(p.call('pub_search', token, 'telefono', '300 000 0001', ''));
    assert.strictEqual(r.items.length, 1);
    assert.ok(!('CLIENTE_ID' in r.items[0]) && r.items[0].ref, 'no expone CLIENTE_ID');
    assert.strictEqual(r.items[0].telefono, '300****001');
    const sel = ok(p.call('pub_select', token, r.items[0].ref));
    assert.strictEqual(sel.prefill.negocio.razonSocial, 'PANADERIA PRUEBA UNO');
    assert.strictEqual(sel.prefill.contacto.whatsapp, '3000000001');
    assert.strictEqual(sel.prefill.contacto.email, 'prueba.uno@example.com');
    assert.strictEqual(sel.prefill.sede.municipio, 'NEIVA');
    assert.strictEqual(sel.prefill.negocio.nombrePropietario, '', 'propietario en blanco (D1)');
  });

  test('búsqueda por nombre no revela teléfono/email completos', () => {
    const { token } = ok(p.call('pub_start', {}));
    const r = ok(p.call('pub_search', token, 'nombre', 'restaurante prueba', ''));
    assert.strictEqual(r.items.length, 1);
    const sel = ok(p.call('pub_select', token, r.items[0].ref));
    assert.strictEqual(sel.prefill.contacto.whatsapp, '');
    assert.strictEqual(sel.prefill.contacto.email, '');
    assert.ok(sel.hints.telefono.indexOf('****') > 0);
  });

  test('búsqueda por email de un email repetido no encuentra nada (D5)', () => {
    const { token } = ok(p.call('pub_start', {}));
    const r = ok(p.call('pub_search', token, 'email', 'compartido@example.com', ''));
    assert.strictEqual(r.items.length, 0);
  });

  test('manipulación: seleccionar una referencia inventada es rechazado', () => {
    const { token } = ok(p.call('pub_start', {}));
    err(p.call('pub_select', token, 'CLI-000001'), 'BAD_REF');
  });

  test('cliente nuevo: envío completo crea cliente, CRM, snapshot y reporte pendiente', () => {
    const { token, radicado } = ok(p.call('pub_start', { ua: 'test' }));
    assert.ok(/^WEB-\d{4}-\d{6}$/.test(radicado));
    ok(p.call('pub_upload', token, { category: 'LOGO', fileName: 'logo.png', mimeType: 'image/png' }, 'data:image/png;base64,' + F.PNG_1PX));
    const photo = ok(p.call('pub_upload', token, { category: 'FOTO_PRODUCTO', fileName: 'cafe.png', mimeType: 'image/png', tempRef: 'prd1' }, F.PNG_1PX));
    ok(p.call('pub_save', token, fullForm(), 5));
    backdate(p, token);
    const r = ok(p.call('pub_submit', token, fullForm(), { ua: 'test' }));
    assert.strictEqual(r.radicado, radicado);
    const cli = p.rows('CLIENTES').find(c => c.RAZON_SOCIAL === 'CAFE DE PRUEBA NUEVO');
    assert.ok(cli, 'cliente creado');
    assert.strictEqual(cli.ORIGEN, 'FORMULARIO');
    assert.strictEqual(cli.TELEFONO, '3001112233');
    assert.strictEqual(cli.ESTADO_FORMULARIO, 'ENVIADO');
    assert.strictEqual(count(p, 'OPORTUNIDADES', o => o.CLIENTE_ID === cli.CLIENTE_ID && o.ETAPA === 'NUEVO'), 1);
    assert.strictEqual(count(p, 'SEGUIMIENTOS', s => s.CLIENTE_ID === cli.CLIENTE_ID && s.TIPO === 'REVISAR_SOLICITUD'), 1);
    assert.strictEqual(count(p, 'AUTORIZACIONES', a => a.CLIENTE_ID === cli.CLIENTE_ID), 8);
    const prod = p.rows('PRODUCTOS').find(x => x.CLIENTE_ID === cli.CLIENTE_ID);
    assert.strictEqual(prod.IMAGEN_ARCHIVO_ID, photo.id, 'foto asociada al producto');
    const arch = p.rows('ARCHIVOS').filter(a => a.SOLICITUD_ID && a.CLIENTE_ID === cli.CLIENTE_ID && a.ESTADO !== 'PAPELERA');
    assert.strictEqual(arch.length, 2, 'archivos asociados al cliente por ID');
    const sol = p.rows('SOLICITUDES').find(s => s.RADICADO === radicado);
    assert.strictEqual(sol.ESTADO, 'ENVIADA');
    assert.strictEqual(sol.REPORTE_ESTADO, 'PENDIENTE');
    assert.ok(sol.SNAPSHOT_FILE_ID);
    const folder = p.env.state.drive.folders[sol.DRIVE_FOLDER_ID];
    assert.ok(folder.name.indexOf(radicado + '_') === 0, 'carpeta WEB-AAAA-NNNNNN_NOMBRE');
    const subs = folder.children().map(f => f.name).sort();
    assert.deepStrictEqual(subs, ['01_IDENTIDAD_VISUAL', '02_EQUIPO', '03_PRODUCTOS', '04_TESTIMONIOS', '05_DOCUMENTOS', '06_OTROS']);
    const logoFile = p.env.state.drive.files[arch.find(a => a.CATEGORIA === 'LOGO').DRIVE_FILE_ID];
    assert.strictEqual(logoFile.parent.name, '01_IDENTIDAD_VISUAL');
  });

  test('duplicado: cliente nuevo con email de la base queda en DUPLICADOS', () => {
    const { token } = ok(p.call('pub_start', {}));
    const form = fullForm({ negocio: { razonSocial: 'OTRO NOMBRE CINCO' }, contacto: { whatsapp: '3009998877', email: 'cinco@example.com' } });
    backdate(p, token);
    ok(p.call('pub_submit', token, form, {}));
    const newCli = p.rows('CLIENTES').find(c => c.RAZON_SOCIAL === 'OTRO NOMBRE CINCO');
    const dup = p.rows('DUPLICADOS').find(d => d.CLIENTE_B === newCli.CLIENTE_ID);
    assert.ok(dup, 'duplicado registrado');
    assert.strictEqual(dup.CLIENTE_A, 'CLI-000005');
    assert.strictEqual(dup.REGLA, 'MISMO_EMAIL');
  });

  test('corrección: cliente de la base propone cambios y se autocompletan vacíos', () => {
    const { token } = ok(p.call('pub_start', {}));
    const r = ok(p.call('pub_search', token, 'telefono', '3000000003', ''));
    ok(p.call('pub_select', token, r.items[0].ref));
    const form = fullForm({
      negocio: { razonSocial: 'TIENDA PRUEBA TRES', nombrePropietario: 'Dueño Tres', nombreComercial: 'La Tienda' },
      contacto: { whatsapp: '3104445566' },
      sedes: [{ ref: 's', direccion: 'CALLE 9 # 9-9', municipio: 'PALERMO' }]
    });
    backdate(p, token);
    ok(p.call('pub_submit', token, form, {}));
    const c = p.rows('CLIENTES').find(x => x.CLIENTE_ID === 'CLI-000003');
    assert.strictEqual(c.NOMBRE_PROPIETARIO, 'Dueño Tres', 'vacío autocompletado');
    assert.strictEqual(c.NOMBRE_COMERCIAL, 'La Tienda');
    assert.strictEqual(c.TELEFONO, '3000000003', 'cambio NO aplicado sin revisión');
    const sol = p.rows('SOLICITUDES').find(s => s.CLIENTE_ID === 'CLI-000003');
    assert.strictEqual(sol.CAMBIOS_ESTADO, 'PENDIENTE');
    assert.ok(sol.CAMBIOS_PROPUESTOS_JSON.some(x => x.campo === 'TELEFONO' && x.propuesto === '3104445566'));
  });

  test('archivo inválido: extensión, contenido falso y PDF en logo', () => {
    const { token } = ok(p.call('pub_start', {}));
    err(p.call('pub_upload', token, { category: 'FOTO_NEGOCIO', fileName: 'virus.exe', mimeType: 'application/octet-stream' }, F.PNG_1PX), 'BAD_FILE');
    err(p.call('pub_upload', token, { category: 'FOTO_NEGOCIO', fileName: 'falsa.png', mimeType: 'image/png' }, Buffer.from('<script>alert(1)</script>').toString('base64')), 'BAD_FILE');
    err(p.call('pub_upload', token, { category: 'LOGO', fileName: 'logo.pdf', mimeType: 'application/pdf' }, F.PDF_MIN), 'BAD_FILE');
    err(p.call('pub_upload', token, { category: 'FOTO_NEGOCIO', fileName: 'x.png', mimeType: 'image/png' }, 'no-es-base64!!'), 'BAD_FILE');
    ok(p.call('pub_upload', token, { category: 'DOCUMENTO', fileName: 'menu.pdf', mimeType: 'application/pdf' }, F.PDF_MIN));
  });

  test('archivo demasiado grande (> 10 MB)', () => {
    const { token } = ok(p.call('pub_start', {}));
    err(p.call('pub_upload', token, { category: 'FOTO_NEGOCIO', fileName: 'grande.jpg', mimeType: 'image/jpeg' }, bigBase64(10.5 * 1024 * 1024)), 'FILE_TOO_LARGE');
  });

  test('formulario incompleto: errores de validación por paso', () => {
    const { token } = ok(p.call('pub_start', {}));
    backdate(p, token);
    const e = err(p.call('pub_submit', token, { negocio: {}, contacto: { email: 'malo' }, autorizaciones: {} }, {}), 'VALIDATION');
    const fields = e.details.map(d => d.field);
    ['razonSocial', 'whatsapp', 'email', 'TRATAMIENTO_DATOS', 'firmanteNombre'].forEach(f => assert.ok(fields.indexOf(f) >= 0, 'falta error ' + f));
  });

  test('testimonio de más de 30 palabras es rechazado', () => {
    const { token } = ok(p.call('pub_start', {}));
    backdate(p, token);
    const long = new Array(35).fill('palabra').join(' ');
    const e = err(p.call('pub_submit', token, fullForm({ testimonios: [{ ref: 't', autor: 'A', texto: long }] }), {}), 'VALIDATION');
    assert.ok(e.details.some(d => d.step === 'testimonios'));
  });

  test('envío demasiado rápido (anti-bot)', () => {
    const { token } = ok(p.call('pub_start', {}));
    err(p.call('pub_submit', token, fullForm(), {}), 'TOO_FAST');
  });

  test('doble envío no duplica información', () => {
    const { token } = ok(p.call('pub_start', {}));
    backdate(p, token);
    const form = fullForm({ negocio: { razonSocial: 'DOBLE ENVIO PRUEBA' } });
    ok(p.call('pub_submit', token, form, {}));
    const second = ok(p.call('pub_submit', token, form, {}));
    assert.strictEqual(second.already, true);
    assert.strictEqual(count(p, 'CLIENTES', c => c.RAZON_SOCIAL === 'DOBLE ENVIO PRUEBA'), 1);
    err(p.call('pub_save', token, form, 3), 'ALREADY_SUBMITTED');
  });

  test('reanudar borrador devuelve datos y archivos propios, sin IDs de Drive', () => {
    const { token } = ok(p.call('pub_start', {}));
    ok(p.call('pub_upload', token, { category: 'FOTO_NEGOCIO', fileName: 'a.png', mimeType: 'image/png' }, F.PNG_1PX));
    ok(p.call('pub_save', token, fullForm({ negocio: { razonSocial: 'BORRADOR X' } }), 4));
    const r = ok(p.call('pub_resume', token));
    assert.strictEqual(r.data.negocio.razonSocial, 'BORRADOR X');
    assert.strictEqual(r.files.length, 1);
    assert.ok(JSON.stringify(r).indexOf('file_') < 0, 'no expone IDs de Drive');
  });

  test('manipulación: borrar archivo de otra solicitud es rechazado', () => {
    const a = ok(p.call('pub_start', {}));
    const b = ok(p.call('pub_start', {}));
    const f = ok(p.call('pub_upload', a.token, { category: 'FOTO_NEGOCIO', fileName: 'a.png', mimeType: 'image/png' }, F.PNG_1PX));
    err(p.call('pub_removeFile', b.token, f.id), 'NOT_FOUND');
    err(p.call('pub_resume', 'f'.repeat(64)), 'BAD_TOKEN');
  });

  test('inyección de fórmulas: se guarda como texto', () => {
    const { token } = ok(p.call('pub_start', {}));
    backdate(p, token);
    ok(p.call('pub_submit', token, fullForm({ negocio: { razonSocial: '=HYPERLINK("http://malo","x")' } }), {}));
    const c = p.rows('CLIENTES').find(x => x.RAZON_SOCIAL.indexOf('HYPERLINK') >= 0);
    assert.ok(c, 'guardado como texto (la simulación falla si una fórmula llega sin sanear)');
  });

  test('municipio fuera de la lista es rechazado', () => {
    const { token } = ok(p.call('pub_start', {}));
    backdate(p, token);
    const e = err(p.call('pub_submit', token, fullForm({ sedes: [{ ref: 's', municipio: 'BOGOTA' }] }), {}), 'VALIDATION');
    assert.ok(e.details.some(d => d.step === 'sedes'));
  });
});

// ================================================================== IMPORTACIÓN
suite('Importación (reglas D1–D9)', () => {
  const { p } = freshSystem();
  const rows = p.rows('CLIENTES').filter(c => c.ORIGEN === 'BASE_CCH');
  test('importa todas las filas con ID por número de fila', () => {
    assert.strictEqual(rows.length, F.ROWS.length);
    assert.strictEqual(rows[0].CLIENTE_ID, 'CLI-000001');
  });
  test('NIT no se guarda en ninguna hoja', () => {
    const all = JSON.stringify(Object.keys(p.env.state.spreadsheets).map(id => p.env.state.spreadsheets[id].sheets.filter(s => s.ss.name !== 'CCH_PRUEBA').map(s => s.rows)));
    assert.ok(all.indexOf('900000001') < 0);
  });
  test('propietario y nombre comercial en blanco', () => { rows.forEach(r => { assert.strictEqual(r.NOMBRE_PROPIETARIO, ''); assert.strictEqual(r.NOMBRE_COMERCIAL, ''); }); });
  test('emails repetidos y mal escritos quedan en blanco', () => {
    assert.strictEqual(rows[1].EMAIL, '');
    assert.strictEqual(rows[2].EMAIL, '');
    assert.strictEqual(rows[3].EMAIL, '');
    assert.strictEqual(rows[0].EMAIL, 'prueba.uno@example.com');
  });
  test('actividad vacía se completa con el CIIU', () => {
    assert.strictEqual(rows[1].ACTIVIDAD, 'Comercio al por menor de articulos de ferreteria');
    assert.strictEqual(rows[1].ACTIVIDAD_FUENTE, 'CIIU');
  });
  test('fijo de 7 dígitos lleva 608; inválido queda vacío', () => {
    assert.strictEqual(rows[1].TELEFONO, '6088700001');
    assert.strictEqual(rows[6].TELEFONO, '');
    assert.ok(rows[4].TELEFONO_COMPARTIDO && rows[5].TELEFONO_COMPARTIDO);
  });
  test('secuencia de clientes continúa después de la base', () => {
    assert.strictEqual(p.rows('SECUENCIAS').find(s => s.NOMBRE === 'CLIENTE').ULTIMO_VALOR, F.ROWS.length);
  });
  test('reimportar no duplica', () => {
    p.asOwner();
    p.call('importBase');
    assert.strictEqual(p.rows('CLIENTES').filter(c => c.ORIGEN === 'BASE_CCH').length, F.ROWS.length);
  });
});

// ================================================================== CRM
suite('CRM', () => {
  const { p, token } = freshSystem();

  test('login: contraseña incorrecta da mensaje genérico', () => {
    const e = err(p.call('adm_login', SUPER.email, 'otraClave123', ''), 'BAD_CREDENTIALS');
    const e2 = err(p.call('adm_login', 'noexiste@x.com', 'otraClave123', ''), 'BAD_CREDENTIALS');
    assert.strictEqual(e.message, e2.message);
  });

  test('clientes: búsqueda, filtros y paginación', () => {
    let r = ok(p.call('adm_searchClients', token, { q: 'ferreteria' }));
    assert.strictEqual(r.total, 1);
    r = ok(p.call('adm_searchClients', token, { filters: { municipio: 'NEIVA' } }));
    assert.strictEqual(r.total, 3);
    r = ok(p.call('adm_searchClients', token, { filters: { sinEmail: true } }));
    assert.strictEqual(r.total, 3);
    r = ok(p.call('adm_searchClients', token, { q: '3000000005' }));
    assert.strictEqual(r.total, 2);
    r = ok(p.call('adm_searchClients', token, { pageSize: 10, page: 1 }));
    assert.strictEqual(r.items.length, F.ROWS.length);
  });

  test('ficha: edición de campos con auditoría y control de versión', () => {
    const full = ok(p.call('adm_getClient', token, 'CLI-000001'));
    const r = ok(p.call('adm_updateClient', token, 'CLI-000001', { NOMBRE_PROPIETARIO: 'Propietario Uno', TELEFONO: '8712345', ORIGEN: 'ADMIN' }, full.cliente.VERSION));
    assert.deepStrictEqual(r.changed.sort(), ['NOMBRE_PROPIETARIO', 'TELEFONO']);
    assert.strictEqual(r.cliente.TELEFONO, '6088712345');
    assert.strictEqual(r.cliente.ORIGEN, 'BASE_CCH', 'campo no editable ignorado');
    err(p.call('adm_updateClient', token, 'CLI-000001', { NOMBRE_PROPIETARIO: 'X' }, full.cliente.VERSION), 'CONFLICT');
    assert.ok(p.rows('AUDITORIA').some(a => a.ACCION === 'EDITAR' && a.ENTIDAD_ID === 'CLI-000001'));
  });

  test('pipeline: crear oportunidad y cambiar etapa', () => {
    const o = ok(p.call('adm_createOpportunity', token, { CLIENTE_ID: 'CLI-000002', VALOR_ESTIMADO: 900000 }));
    assert.strictEqual(o.ETAPA, 'POR_CONTACTAR');
    const m = ok(p.call('adm_moveOpportunity', token, o.OPORTUNIDAD_ID, 'REUNION', o.VERSION, {}));
    assert.strictEqual(m.ETAPA, 'REUNION');
    assert.strictEqual(m.PROBABILIDAD, 50);
    assert.strictEqual(p.rows('CLIENTES').find(c => c.CLIENTE_ID === 'CLI-000002').ESTADO_CRM, 'REUNION');
    err(p.call('adm_moveOpportunity', token, o.OPORTUNIDAD_ID, 'GANADO', m.VERSION, {}), 'SALE_REQUIRED');
    err(p.call('adm_moveOpportunity', token, o.OPORTUNIDAD_ID, 'PERDIDO', m.VERSION, {}), 'REASON_REQUIRED');
    err(p.call('adm_moveOpportunity', token, o.OPORTUNIDAD_ID, 'INVENTADA', m.VERSION, {}), 'VALIDATION');
    assert.ok(p.rows('AUDITORIA').some(a => a.ACCION === 'CAMBIO_ETAPA' && a.ENTIDAD_ID === o.OPORTUNIDAD_ID));
  });

  test('Kanban: agrupa por las 10 etapas con totales', () => {
    const k = ok(p.call('adm_listOpportunities', token, {}));
    assert.strictEqual(k.stages.length, 10);
    assert.deepStrictEqual(k.stages.map(s => s.key), ['NUEVO', 'POR_CONTACTAR', 'CONTACTADO', 'INTERESADO', 'REUNION', 'PROPUESTA_ENVIADA', 'NEGOCIACION', 'GANADO', 'PERDIDO', 'NO_INTERESADO']);
    assert.strictEqual(k.totals.REUNION.count, 1);
    assert.strictEqual(k.totals.REUNION.value, 900000);
  });

  test('actividad: primera llamada saliente mueve a Contactado', () => {
    const r = ok(p.call('adm_logActivity', token, { CLIENTE_ID: 'CLI-000004', TIPO: 'LLAMADA', RESULTADO: 'CONTESTO', ASUNTO: 'Primera llamada' }));
    assert.ok(r.actividad.ACTIVIDAD_ID);
    assert.strictEqual(p.rows('CLIENTES').find(c => c.CLIENTE_ID === 'CLI-000004').ESTADO_CRM, 'CONTACTADO');
    ['WHATSAPP', 'EMAIL', 'REUNION', 'NOTA'].forEach(t => ok(p.call('adm_logActivity', token, { CLIENTE_ID: 'CLI-000004', TIPO: t, ASUNTO: t })));
    assert.strictEqual(count(p, 'ACTIVIDADES', a => a.CLIENTE_ID === 'CLI-000004' && a.TIPO !== 'CAMBIO_ETAPA'), 5);
  });

  test('actividad: "volver a llamar" exige seguimiento; NO CONTACTAR bloquea', () => {
    err(p.call('adm_logActivity', token, { CLIENTE_ID: 'CLI-000004', TIPO: 'LLAMADA', RESULTADO: 'VOLVER_A_LLAMAR', ASUNTO: 'x' }), 'FOLLOWUP_REQUIRED');
    const r = ok(p.call('adm_logActivity', token, { CLIENTE_ID: 'CLI-000004', TIPO: 'LLAMADA', RESULTADO: 'VOLVER_A_LLAMAR', ASUNTO: 'x', seguimiento: { FECHA_VENCIMIENTO: '2020-01-01', DESCRIPCION: 'Llamar' } }));
    assert.ok(r.seguimiento.SEGUIMIENTO_ID);
    ok(p.call('adm_updateClient', token, 'CLI-000007', { NO_CONTACTAR: true }));
    err(p.call('adm_logActivity', token, { CLIENTE_ID: 'CLI-000007', TIPO: 'LLAMADA', ASUNTO: 'x' }), 'DO_NOT_CONTACT');
    ok(p.call('adm_logActivity', token, { CLIENTE_ID: 'CLI-000007', TIPO: 'NOTA', ASUNTO: 'Pidió no ser contactado' }));
  });

  test('seguimientos: vencidos, hoy, próximos y completar', () => {
    const today = p.run('Utils.today()');
    ok(p.call('adm_saveFollowup', token, { CLIENTE_ID: 'CLI-000001', TIPO: 'LLAMAR', FECHA_VENCIMIENTO: today, DESCRIPCION: 'hoy' }));
    ok(p.call('adm_saveFollowup', token, { CLIENTE_ID: 'CLI-000001', TIPO: 'EMAIL', FECHA_VENCIMIENTO: '2099-01-01', DESCRIPCION: 'futuro' }));
    const venc = ok(p.call('adm_listFollowups', token, 'vencidos', 'all'));
    const hoy = ok(p.call('adm_listFollowups', token, 'hoy', 'all'));
    const prox = ok(p.call('adm_listFollowups', token, 'proximos', 'all'));
    assert.ok(venc.length >= 1 && hoy.length === 1 && prox.length === 1);
    ok(p.call('adm_completeFollowup', token, hoy[0].SEGUIMIENTO_ID, 'Hecho', 'COMPLETADO'));
    assert.strictEqual(ok(p.call('adm_listFollowups', token, 'hoy', 'all')).length, 0);
    err(p.call('adm_saveFollowup', token, { CLIENTE_ID: 'CLI-000001', FECHA_VENCIMIENTO: today, ASIGNADO_A: 'fantasma@x.com' }), 'VALIDATION');
  });

  test('propuesta: totales calculados en el servidor (ignora manipulación)', () => {
    const o = ok(p.call('adm_createOpportunity', token, { CLIENTE_ID: 'CLI-000005' }));
    const pr = ok(p.call('adm_saveProposal', token, { OPORTUNIDAD_ID: o.OPORTUNIDAD_ID, items: [{ descripcion: 'Página básica', cantidad: 1, precioUnitario: '$ 1.200.000' }, { descripcion: 'Agenda', cantidad: 2, precioUnitario: 150000 }], DESCUENTO: 100000, TOTAL: 1 }));
    assert.strictEqual(pr.SUBTOTAL, 1500000);
    assert.strictEqual(pr.TOTAL, 1400000);
    ok(p.call('adm_setProposalStatus', token, pr.PROPUESTA_ID, 'ENVIADA'));
    assert.strictEqual(p.rows('OPORTUNIDADES').find(x => x.OPORTUNIDAD_ID === o.OPORTUNIDAD_ID).ETAPA, 'PROPUESTA_ENVIADA');
  });

  test('venta: registra cierre, GANADO, proyecto web y pagos', () => {
    const o = p.rows('OPORTUNIDADES').find(x => x.CLIENTE_ID === 'CLI-000005');
    const pr = p.rows('PROPUESTAS').find(x => x.OPORTUNIDAD_ID === o.OPORTUNIDAD_ID);
    const r = ok(p.call('adm_registerSale', token, { OPORTUNIDAD_ID: o.OPORTUNIDAD_ID, PROPUESTA_ID: pr.PROPUESTA_ID, FORMA_PAGO: 'ANTICIPO_Y_SALDO', PAGO_INICIAL: 700000, METODO: 'NEQUI' }));
    assert.strictEqual(r.venta.VALOR_TOTAL, 1400000);
    assert.strictEqual(r.venta.ESTADO_PAGO, 'PARCIAL');
    assert.strictEqual(p.rows('OPORTUNIDADES').find(x => x.OPORTUNIDAD_ID === o.OPORTUNIDAD_ID).ETAPA, 'GANADO');
    assert.strictEqual(p.rows('PROPUESTAS').find(x => x.PROPUESTA_ID === pr.PROPUESTA_ID).ESTADO, 'ACEPTADA');
    assert.strictEqual(count(p, 'PROYECTOS_WEB', x => x.VENTA_ID === r.venta.VENTA_ID && x.ESTADO === 'BRIEF'), 1);
    ok(p.call('adm_registerPayment', token, { VENTA_ID: r.venta.VENTA_ID, VALOR: 700000, METODO: 'TRANSFERENCIA' }));
    assert.strictEqual(p.rows('VENTAS').find(v => v.VENTA_ID === r.venta.VENTA_ID).ESTADO_PAGO, 'PAGADA');
    err(p.call('adm_registerSale', token, { OPORTUNIDAD_ID: o.OPORTUNIDAD_ID, VALOR_TOTAL: 5 }), 'VALIDATION');
  });

  test('dashboard: indicadores coherentes', () => {
    const d = ok(p.call('adm_dashboard', token));
    assert.ok(d.clientes.total >= F.ROWS.length);
    assert.strictEqual(d.ventas.mesCantidad, 1);
    assert.strictEqual(d.embudo.length, 5);
    assert.ok(d.oportunidades.porEtapa.GANADO.count >= 1);
  });

  test('auditoría: registra creación, edición, etapa, actividad, seguimiento y venta', () => {
    const acciones = new Set(p.rows('AUDITORIA').map(a => a.ACCION));
    ['CREAR', 'EDITAR', 'CAMBIO_ETAPA', 'ACTIVIDAD', 'SEGUIMIENTO', 'VENTA', 'PAGO', 'LOGIN'].forEach(a => assert.ok(acciones.has(a), 'falta ' + a));
    const q = ok(p.call('adm_listAudit', token, { entidad: 'VENTAS' }, 1));
    assert.ok(q.items.length >= 1);
  });

  test('duplicados: fusión manual mueve datos y conserva el registro A', () => {
    ok(p.call('adm_createClient', token, { RAZON_SOCIAL: 'PANADERIA PRUEBA UNO', MUNICIPIO: 'NEIVA', TELEFONO: '3000000001' }));
    const dups = ok(p.call('adm_listDuplicates', token, 'PENDIENTE'));
    const d = dups.find(x => x.CLIENTE_A === 'CLI-000001');
    assert.ok(d, 'duplicado detectado');
    ok(p.call('adm_logActivity', token, { CLIENTE_ID: d.CLIENTE_B, TIPO: 'NOTA', ASUNTO: 'nota en B' }));
    const r = ok(p.call('adm_resolveDuplicate', token, d.DUP_ID, 'ES_EL_MISMO', {}));
    assert.strictEqual(r.clienteId, 'CLI-000001');
    assert.ok(p.rows('ACTIVIDADES').some(a => a.ASUNTO === 'nota en B' && a.CLIENTE_ID === 'CLI-000001'));
    assert.strictEqual(p.rows('CLIENTES').find(c => c.CLIENTE_ID === d.CLIENTE_B).FUSIONADO_EN, 'CLI-000001');
  });

  test('reporte: Google Doc versionado, sin información interna del CRM', () => {
    const cli = p.rows('CLIENTES').find(c => c.RAZON_SOCIAL === 'CLIENTE REPORTE') || null;
    assert.strictEqual(cli, null);
    // cliente del formulario
    p.asAnonymous();
    const { token: t } = ok(p.call('pub_start', {}));
    backdate(p, t);
    ok(p.call('pub_upload', t, { category: 'LOGO', fileName: 'logo.png', mimeType: 'image/png' }, F.PNG_1PX));
    ok(p.call('pub_submit', t, fullForm({ negocio: { razonSocial: 'CLIENTE REPORTE' } }), {}));
    const c = p.rows('CLIENTES').find(x => x.RAZON_SOCIAL === 'CLIENTE REPORTE');
    ok(p.call('adm_logActivity', token, { CLIENTE_ID: c.CLIENTE_ID, TIPO: 'NOTA', ASUNTO: 'NOTA_SECRETA_INTERNA', DETALLE: 'NO DEBE SALIR' }));
    p.asOwner();
    const t0 = p.env.state.triggers.find(x => x.fn === 'cronProcessReports');
    const res = p.call('cronProcessReports', { triggerUid: t0.uid });
    assert.strictEqual(res.generados, 1);
    p.asAnonymous();
    const rep = p.rows('REPORTES').find(r => r.CLIENTE_ID === c.CLIENTE_ID);
    assert.strictEqual(rep.VERSION_REPORTE, 1);
    const doc = p.env.state.docs[rep.DOC_ID];
    const text = doc._body.getText();
    assert.ok(text.indexOf('CLIENTE REPORTE') >= 0 && text.indexOf('Catas') >= 0);
    assert.ok(text.indexOf('NOTA_SECRETA_INTERNA') < 0 && text.indexOf('NO DEBE SALIR') < 0, 'no mezcla notas internas');
    assert.ok(doc._body.images >= 1, 'logo insertado');
    assert.strictEqual(p.env.state.drive.files[rep.DOC_ID].parent.name, '05_DOCUMENTOS');
    const r2 = ok(p.call('adm_generateReport', token, c.CLIENTE_ID, false));
    assert.strictEqual(r2.VERSION_REPORTE, 2);
    assert.strictEqual(p.rows('REPORTES').find(r => r.REPORTE_ID === rep.REPORTE_ID).ESTADO, 'REEMPLAZADO');
    const pdf = ok(p.call('adm_reportPdf', token, r2.REPORTE_ID));
    assert.ok(pdf.base64.length > 10);
  });

  test('solicitud: aplicar correcciones propuestas', () => {
    p.asAnonymous();
    const { token: t } = ok(p.call('pub_start', {}));
    const r = ok(p.call('pub_search', t, 'telefono', '3000000004', ''));
    ok(p.call('pub_select', t, r.items[0].ref));
    backdate(p, t);
    ok(p.call('pub_submit', t, fullForm({ negocio: { razonSocial: 'PELUQUERIA RENOVADA' }, contacto: { whatsapp: '3000000004' }, sedes: [{ ref: 's', municipio: 'RIVERA', direccion: 'AVENIDA 4 # 4-4' }] }), {}));
    const sol = p.rows('SOLICITUDES').find(s => s.CLIENTE_ID === 'CLI-000004' && s.ESTADO === 'ENVIADA');
    const det = ok(p.call('adm_getSolicitud', token, sol.SOLICITUD_ID));
    assert.ok(JSON.stringify(det).indexOf('TOKEN_HASH') < 0);
    assert.ok(det.cambios.some(c => c.campo === 'RAZON_SOCIAL'));
    ok(p.call('adm_applyCorrections', token, sol.SOLICITUD_ID, ['RAZON_SOCIAL'], false));
    assert.strictEqual(p.rows('CLIENTES').find(c => c.CLIENTE_ID === 'CLI-000004').RAZON_SOCIAL, 'PELUQUERIA RENOVADA');
    ok(p.call('adm_setSolicitudEstado', token, sol.SOLICITUD_ID, 'APROBADA', 'Todo bien'));
    assert.strictEqual(p.rows('CLIENTES').find(c => c.CLIENTE_ID === 'CLI-000004').ESTADO_FORMULARIO, 'APROBADO');
  });
});

// ================================================================== SEGURIDAD
suite('Seguridad', () => {
  const { p, token } = freshSystem();
  const u1 = ok(p.call('adm_createUser', token, { EMAIL: 'vendedor@webpaya.test', NOMBRE: 'Vendedor', ROL: 'VENDEDOR' }));
  const u2 = ok(p.call('adm_createUser', token, { EMAIL: 'lectura@webpaya.test', NOMBRE: 'Lectura', ROL: 'LECTURA' }));
  const u3 = ok(p.call('adm_createUser', token, { EMAIL: 'admin2@webpaya.test', NOMBRE: 'Admin', ROL: 'ADMIN' }));
  const loginAndChange = (email, temp) => {
    const t = ok(p.call('adm_login', email, temp, '')).token;
    err(p.call('adm_dashboard', t), 'PASSWORD_CHANGE_REQUIRED');
    ok(p.call('adm_changePassword', t, temp, 'NuevaClave2026x'));
    return ok(p.call('adm_login', email, 'NuevaClave2026x', '')).token;
  };
  const tv = loginAndChange('vendedor@webpaya.test', u1.temporaryPassword);
  const tl = loginAndChange('lectura@webpaya.test', u2.temporaryPassword);
  const ta = loginAndChange('admin2@webpaya.test', u3.temporaryPassword);
  ok(p.call('adm_assignClient', token, 'CLI-000001', 'vendedor@webpaya.test'));

  test('usuario no autenticado', () => {
    err(p.call('adm_dashboard', ''), 'UNAUTHENTICATED');
    err(p.call('adm_getClient', 'a'.repeat(64), 'CLI-000001'), 'UNAUTHENTICATED');
    err(p.call('adm_listUsers', 'x'), 'UNAUTHENTICATED');
  });

  test('usuario sin permisos', () => {
    err(p.call('adm_listUsers', tv), 'FORBIDDEN');
    err(p.call('adm_listAudit', ta, {}, 1), 'FORBIDDEN');
    err(p.call('adm_exportClients', ta, {}), 'FORBIDDEN');
    err(p.call('adm_updateClient', tl, 'CLI-000001', { PRIORIDAD: 'ALTA' }), 'FORBIDDEN');
    err(p.call('adm_logActivity', tl, { CLIENTE_ID: 'CLI-000001', TIPO: 'NOTA', ASUNTO: 'x' }), 'FORBIDDEN');
    ok(p.call('adm_getClient', tl, 'CLI-000002'));
  });

  test('vendedor solo ve sus clientes', () => {
    const r = ok(p.call('adm_searchClients', tv, {}));
    assert.deepStrictEqual(r.items.map(c => c.CLIENTE_ID), ['CLI-000001']);
    err(p.call('adm_getClient', tv, 'CLI-000002'), 'FORBIDDEN');
    ok(p.call('adm_getClient', tv, 'CLI-000001'));
    err(p.call('adm_assignClient', tv, 'CLI-000001', 'lectura@webpaya.test'), 'FORBIDDEN');
    assert.ok(p.rows('AUDITORIA').some(a => a.RESULTADO === 'DENEGADO' && a.CLIENTE_ID === 'CLI-000002'));
  });

  test('manipulación del frontend: tablas y campos no permitidos', () => {
    err(p.call('adm_saveChild', token, 'USUARIOS_ADMIN', { EMAIL: 'x@x.com' }), 'FORBIDDEN');
    err(p.call('adm_saveChild', token, 'CONTENIDO_WEB', { CLIENTE_ID: 'CLI-000001', CAMPO_CLAVE: 'inventado', VALOR: 'x' }), 'VALIDATION');
    const r = ok(p.call('adm_updateClient', token, 'CLI-000002', { CLIENTE_ID: 'CLI-999999', ESTADO_CRM: 'GANADO', FUSIONADO_EN: 'x', DRIVE_FOLDER_ID: 'y' }));
    assert.deepStrictEqual(r.changed, []);
    err(p.call('adm_updateClient', token, 'CLI-000002', { MUNICIPIO: 'CALI' }), 'VALIDATION');
    err(p.call('adm_updateClient', token, 'CLI-000002', { EMAIL: 'malo' }), 'VALIDATION');
  });

  test('acceso a funciones administrativas y de mantenimiento', () => {
    p.asAnonymous();
    // Las funciones del propietario lanzan error (google.script.run lo entrega al withFailureHandler).
    assert.throws(() => p.call('setup'), /reservada al propietario/);
    assert.throws(() => p.call('importBase'), /reservada al propietario/);
    assert.throws(() => p.call('cronProcessReports', { triggerUid: 'inventado' }), /reservada al propietario/);
    assert.throws(() => p.call('respond_', () => 1), /privadas/);
    // Toda función global expuesta es pública, administrativa o del propietario.
    const fs = require('fs');
    const path = require('path');
    const src = path.join(__dirname, '..', 'src');
    const exposed = [];
    fs.readdirSync(src).filter(f => f.endsWith('.gs')).forEach(f => {
      const re = /^function\s+([A-Za-z0-9_$]+)\s*\(/gm;
      let m;
      const code = fs.readFileSync(path.join(src, f), 'utf8');
      while ((m = re.exec(code))) if (!/_$/.test(m[1])) exposed.push(m[1]);
    });
    assert.ok(exposed.length > 60, 'se encontraron las funciones globales');
    const allowedOwner = ['doGet', 'setup', 'createSetupCode', 'installTriggers', 'importBase', 'rebuildIndex', 'cronProcessReports', 'cronDaily', 'cronContinueImport'];
    exposed.forEach(k => assert.ok(/^pub_|^adm_/.test(k) || allowedOwner.indexOf(k) >= 0, 'función expuesta inesperada: ' + k));
  });

  test('no se exponen contraseñas, tokens ni IDs privados', () => {
    const users = JSON.stringify(ok(p.call('adm_listUsers', token)));
    assert.ok(users.indexOf('PASSWORD_HASH') < 0 && users.indexOf('PASSWORD_SALT') < 0 && users.indexOf('PASSWORD_ITER') < 0);
    const me = JSON.stringify(ok(p.call('adm_me', token)));
    assert.ok(me.indexOf('PASSWORD_HASH') < 0 && me.indexOf('TOKEN_HASH') < 0);
    const audit = JSON.stringify(p.rows('AUDITORIA'));
    assert.ok(audit.indexOf(SUPER.pass) < 0 && audit.indexOf('NuevaClave2026x') < 0);
  });

  test('bloqueo tras 5 intentos fallidos', () => {
    for (let i = 0; i < 5; i++) err(p.call('adm_login', 'lectura@webpaya.test', 'mal' + i + 'aaaaaaa', ''), 'BAD_CREDENTIALS');
    err(p.call('adm_login', 'lectura@webpaya.test', 'NuevaClave2026x', ''), 'LOCKED');
  });

  test('desactivar usuario revoca su sesión', () => {
    ok(p.call('adm_updateUser', token, u3.user.USUARIO_ID, { ACTIVO: false }));
    err(p.call('adm_dashboard', ta), 'UNAUTHENTICATED');
  });

  test('sesión vence por inactividad', () => {
    const cache = p.env.state.cache;
    for (const [k, v] of cache) {
      if (k.indexOf('as:') === 0) { const s = JSON.parse(v.v); s.last -= 2 * 3600 * 1000; v.v = JSON.stringify(s); }
    }
    err(p.call('adm_dashboard', token), 'UNAUTHENTICATED');
  });

  test('código de configuración es de un solo uso', () => {
    const s = freshSystem();
    err(s.p.call('adm_createSuperAdmin', s.code, 'otro@x.com', 'Otro', 'ClaveSegura2026'), 'BAD_CODE');
  });
});

// ================================================================== CONTRATO CLIENTE ↔ SERVIDOR
suite('Contrato frontend ↔ backend', () => {
  const fs = require('fs');
  const path = require('path');
  const src = path.join(__dirname, '..', 'src');
  const read = (f) => fs.readFileSync(path.join(src, f), 'utf8');
  const defined = new Set();
  fs.readdirSync(src).filter(f => f.endsWith('.gs')).forEach(f => {
    const re = /^function\s+([A-Za-z0-9_$]+)\s*\(/gm;
    let m;
    const code = read(f);
    while ((m = re.exec(code))) defined.add(m[1]);
  });
  test('toda llamada rpc()/api() del HTML existe en el servidor', () => {
    const calls = new Set();
    [['Scripts.html', /rpc\('([a-z]+_[A-Za-z0-9]+)'/g], ['AdminScripts.html', /(?:rpc|api)\('([a-z]+_[A-Za-z0-9]+)'/g]].forEach(([f, re]) => {
      let m;
      const code = read(f);
      while ((m = re.exec(code))) calls.add(m[1]);
    });
    assert.ok(calls.size > 50, 'se detectaron las llamadas: ' + calls.size);
    calls.forEach(c => assert.ok(defined.has(c), 'el HTML llama a ' + c + ' que no existe'));
  });
  test('includes de plantillas existen', () => {
    ['Index.html', 'Admin.html'].forEach(f => {
      const re = /include_\('([A-Za-z]+)'\)/g;
      let m;
      const code = read(f);
      while ((m = re.exec(code))) assert.ok(fs.existsSync(path.join(src, m[1] + '.html')), 'falta ' + m[1] + '.html');
    });
  });
  test('ningún archivo usa innerHTML con datos', () => {
    ['Scripts.html', 'AdminScripts.html'].forEach(f => assert.ok(!/\.(innerHTML|outerHTML)\s*=|insertAdjacentHTML|document\.write/.test(read(f)), f + ' inserta HTML sin escapar'));
  });
});

// ================================================================== PÁGINAS
suite('Páginas (HtmlService)', () => {
  const { p } = freshSystem();
  test('doGet sirve el formulario y el panel con sus includes', () => {
    const pub = p.run('doGet({parameter:{}})');
    assert.ok(pub._html.indexOf('data-step="welcome"') >= 0 && pub._html.indexOf('<script>') >= 0 && pub._html.indexOf('<style>') >= 0);
    assert.ok(pub._html.indexOf('<?') < 0, 'sin scriptlets sin evaluar');
    const adm = p.run('doGet({parameter:{page:"admin"}})');
    assert.ok(adm._html.indexOf('id="loginForm"') >= 0 && adm._html.indexOf('adm_login') >= 0);
  });
});

// ------------------------------------------------------------------ reporte
const failed = results.filter(r => !r.ok);
results.forEach(r => console.log((r.ok ? '  ✔ ' : '  ✘ ') + r.name + (r.ok ? '' : '\n      ' + (r.err && r.err.stack ? r.err.stack.split('\n').slice(0, 3).join('\n      ') : r.err))));
console.log('\n' + (results.length - failed.length) + ' / ' + results.length + ' pruebas OK');
process.exit(failed.length ? 1 : 0);

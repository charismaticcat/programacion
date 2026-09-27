/**
 * Pruebas de interfaz con Playwright (Chromium) sobre el servidor local.
 *   NODE_PATH=$(npm root -g) node tests/ui-tests.js
 * Guarda capturas en tests/output/.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { chromium, devices } = require('playwright');
const { start, ADMIN } = require('./dev-server');
const F = require('./fixtures');

const OUT = path.join(__dirname, 'output');
fs.mkdirSync(OUT, { recursive: true });
const PNG_FILE = path.join(OUT, 'logo-prueba.png');
fs.writeFileSync(PNG_FILE, Buffer.from(F.PNG_1PX, 'base64'));
const BAD_FILE = path.join(OUT, 'malo.exe');
fs.writeFileSync(BAD_FILE, 'MZ no es imagen');

const PORT = 8799;
const BASE = 'http://localhost:' + PORT;
const results = [];

async function run(name, browser, opts, fn) {
  const context = await browser.newContext(opts || {});
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error' && !/Failed to load resource|ERR_INTERNET_DISCONNECTED/.test(m.text())) errors.push(m.text()); });
  page.on('dialog', d => d.accept());
  try {
    await fn(page, context);
    assert.deepStrictEqual(errors, [], 'errores de JavaScript en la página');
    results.push({ ok: true, name });
  } catch (e) {
    results.push({ ok: false, name, err: e });
    try { await page.screenshot({ path: path.join(OUT, 'FALLO-' + name.replace(/\W+/g, '_') + '.png'), fullPage: true }); } catch (x) { /* sin captura */ }
  }
  await context.close();
}

async function next(page) { await page.click('#btnNext'); }
async function stepIs(page, id) { await page.waitForSelector('.step.active[data-step="' + id + '"]', { timeout: 8000 }); }
async function post(pathname, body) {
  const r = await fetch(BASE + pathname, { method: 'POST', body: JSON.stringify(body) });
  return (await r.json()).result;
}

(async () => {
  const server = start(PORT);
  const browser = await chromium.launch();

  await run('Público: cliente existente, precarga, archivos, envío y confirmación', browser, { viewport: { width: 1280, height: 900 } }, async (page) => {
    await page.goto(BASE + '/');
    await stepIs(page, 'welcome');
    await page.screenshot({ path: path.join(OUT, '01-bienvenida.png'), fullPage: true });
    await next(page); // sin aceptar aviso
    await page.waitForSelector('.toast-msg.error');
    await page.check('#acceptNotice');
    await next(page);
    await stepIs(page, 'identificacion');
    await page.click('#optSearch');
    await page.fill('#searchValue', '3000000001');
    await page.click('#btnSearch');
    await page.waitForSelector('.result');
    await page.screenshot({ path: path.join(OUT, '02-busqueda.png'), fullPage: true });
    await page.click('.result button');
    await stepIs(page, 'negocio');
    assert.strictEqual(await page.inputValue('#f-razonSocial'), 'PANADERIA PRUEBA UNO');
    await page.fill('#f-nombrePropietario', 'Propietaria Prueba');
    await page.fill('#f-tipo1', 'Panadería');
    await page.screenshot({ path: path.join(OUT, '03-datos-precargados.png'), fullPage: true });
    await next(page);
    await stepIs(page, 'contacto');
    assert.strictEqual(await page.inputValue('#f-whatsapp'), '3000000001');
    await page.fill('#f-fijo', '12');
    await next(page);
    await page.waitForSelector('.field.invalid');
    await page.fill('#f-fijo', '8765432');
    await next(page);
    await stepIs(page, 'sedes');
    assert.strictEqual(await page.inputValue('#f-sedes-0-municipio'), 'NEIVA');
    await page.fill('#f-sedes-0-horarioSemana', 'Lun a vie 6 a.m. - 8 p.m.');
    await next(page);
    await stepIs(page, 'marca');
    await page.fill('#f-mensaje', 'Pan artesanal todos los días');
    await next(page);
    await stepIs(page, 'servicios');
    await page.click('[data-add=servicios]');
    await page.fill('#f-servicios-0-nombre', 'Tortas por encargo');
    await page.fill('#f-servicios-0-precio', '45000');
    await next(page);
    await stepIs(page, 'productos');
    await page.click('[data-add=productos]');
    await page.fill('#f-productos-0-nombre', 'Pan de bono');
    await page.fill('#f-productos-0-precio', '1500');
    await page.setInputFiles('#rep-productos .file-btn input', PNG_FILE);
    await page.waitForSelector('#rep-productos .thumb:not(.loading)');
    await next(page);
    await stepIs(page, 'archivos');
    await page.setInputFiles('#up-LOGO', PNG_FILE);
    await page.waitForSelector('.upload-block[data-category=LOGO] .thumb:not(.loading)');
    await page.setInputFiles('#up-FOTO_NEGOCIO', BAD_FILE);
    await page.waitForSelector('.toast-msg.error');
    await page.screenshot({ path: path.join(OUT, '04-archivos.png'), fullPage: true });
    await next(page);
    await stepIs(page, 'testimonios');
    await page.click('[data-add=testimonios]');
    await page.fill('#f-testimonios-0-autor', 'Cliente feliz');
    await page.fill('#f-testimonios-0-texto', 'El mejor pan del barrio');
    await next(page);
    await stepIs(page, 'adicionales');
    await page.check('input[name=agenda][value=SI]');
    await next(page);
    await stepIs(page, 'autorizaciones');
    await next(page);
    await stepIs(page, 'revision');
    assert.ok(await page.isVisible('#reviewErrors .notice.error'), 'la revisión muestra lo que falta');
    await page.click('#reviewErrors a');
    await stepIs(page, 'autorizaciones');
    await page.check('#aut-TRATAMIENTO_DATOS');
    await page.check('#aut-USO_IMAGENES');
    await page.fill('#f-firmante', 'Propietaria Prueba');
    await next(page);
    await stepIs(page, 'revision');
    await page.screenshot({ path: path.join(OUT, '05-revision.png'), fullPage: true });
    const token = await page.evaluate(() => localStorage.getItem('wp_form_token'));
    await post('/__test/backdate', { token });
    await next(page);
    await stepIs(page, 'confirmacion');
    const radicado = await page.textContent('#okRadicado');
    assert.ok(/^WEB-\d{4}-\d{6}$/.test(radicado));
    await page.screenshot({ path: path.join(OUT, '06-confirmacion.png'), fullPage: true });
    // Recargar muestra la confirmación, no un formulario editable (doble envío)
    await page.reload();
    await stepIs(page, 'confirmacion');
    const cli = (await post('/__test/rows', { table: 'CLIENTES' })).find(c => c.CLIENTE_ID === 'CLI-000001');
    assert.strictEqual(cli.NOMBRE_PROPIETARIO, 'Propietaria Prueba');
    assert.strictEqual(cli.ESTADO_FORMULARIO, 'ENVIADO');
  });

  await run('Público: pérdida de conexión guarda en el dispositivo y sincroniza al volver', browser, {}, async (page, context) => {
    await page.goto(BASE + '/');
    await page.check('#acceptNotice');
    await next(page);
    await stepIs(page, 'identificacion');
    await page.click('#optNew');
    await stepIs(page, 'negocio');
    await context.setOffline(true);
    await page.fill('#f-razonSocial', 'NEGOCIO SIN CONEXION');
    await page.waitForFunction(() => document.querySelector('#saveState').classList.contains('offline'), null, { timeout: 10000 });
    const backup = await page.evaluate(() => Object.keys(localStorage).filter(k => k.indexOf('wp_backup_') === 0).map(k => localStorage.getItem(k))[0]);
    assert.ok(backup && backup.indexOf('NEGOCIO SIN CONEXION') >= 0, 'copia local');
    await context.setOffline(false);
    await page.evaluate(() => window.dispatchEvent(new Event('online')));
    await page.waitForFunction(() => /Guardado/.test(document.querySelector('#saveState .txt').textContent), null, { timeout: 15000 });
    const sols = await post('/__test/rows', { table: 'SOLICITUDES' });
    assert.ok(sols.some(s => s.DATOS_JSON && s.DATOS_JSON.negocio && s.DATOS_JSON.negocio.razonSocial === 'NEGOCIO SIN CONEXION'));
  });

  await run('Público: celular (sin desbordamiento horizontal)', browser, Object.assign({}, devices['iPhone 12']), async (page) => {
    await page.goto(BASE + '/');
    await stepIs(page, 'welcome');
    assert.ok(await page.isVisible('#progressMobile'));
    assert.ok(!(await page.isVisible('.stepper')));
    await page.check('#acceptNotice');
    await next(page);
    await stepIs(page, 'identificacion');
    await page.click('#optNew');
    await stepIs(page, 'negocio');
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    assert.ok(overflow <= 1, 'desbordamiento horizontal: ' + overflow);
    await page.screenshot({ path: path.join(OUT, '07-movil.png'), fullPage: true });
  });

  await run('Super Admin: login, dashboard, clientes, ficha, Kanban, ventas y auditoría', browser, { viewport: { width: 1400, height: 950 } }, async (page) => {
    await page.goto(BASE + '/?page=admin');
    await page.waitForSelector('#viewLogin:not([hidden])');
    await page.fill('#loginEmail', ADMIN.email);
    await page.fill('#loginPass', 'incorrecta123');
    await page.click('#loginBtn');
    await page.waitForSelector('#loginMsg .notice.error');
    await page.fill('#loginPass', ADMIN.pass);
    await page.click('#loginBtn');
    await page.waitForSelector('#viewApp:not([hidden]) .tile');
    await page.waitForSelector('.hbar');
    await page.screenshot({ path: path.join(OUT, '10-dashboard.png'), fullPage: true });
    // Solicitudes
    await page.click('#nav button[data-view=solicitudes]');
    await page.waitForSelector('#viewBody .tbl tr.click');
    await page.click('#viewBody .tbl tr.click');
    await page.waitForSelector('.modal');
    await page.screenshot({ path: path.join(OUT, '11-solicitud.png'), fullPage: true });
    await page.click('.modal footer button:has-text("Cerrar")');
    // Clientes
    await page.click('#nav button[data-view=clientes]');
    await page.waitForSelector('#viewBody .tbl tr.click');
    await page.fill('#fc-q', 'drogueria');
    await page.waitForFunction(() => document.querySelectorAll('#viewBody .tbl tbody tr').length === 1);
    await page.selectOption('#fc-municipio', 'NEIVA');
    await page.waitForSelector('#viewBody .tbl tr.click');
    await page.screenshot({ path: path.join(OUT, '12-clientes.png'), fullPage: true });
    await page.click('#viewBody .tbl tr.click');
    await page.waitForSelector('.client-head');
    await page.screenshot({ path: path.join(OUT, '13-ficha.png'), fullPage: true });
    // Editar datos
    await page.click('.tabs button:has-text("Datos del cliente")');
    await page.fill('[data-col=NOMBRE_COMERCIAL]', 'Droguería La Prueba');
    await page.click('button:has-text("Guardar cambios")');
    await page.waitForSelector('.toast.ok');
    // Actividad
    await page.click('.client-head button:has-text("Registrar actividad")');
    await page.fill('#actAsunto', 'Llamada de presentación');
    await page.selectOption('#actRes', 'INTERESADO');
    await page.click('.modal footer button:has-text("Guardar")');
    await page.waitForSelector('.toast.ok');
    await page.click('.tabs button:has-text("Oportunidades y actividades")');
    await page.waitForSelector('.timeline li');
    // Propuesta y venta
    await page.click('.tabs button:has-text("Propuestas y ventas")');
    await page.click('button:has-text("+ Propuesta")');
    await page.fill('.modal input[placeholder=Descripción]', 'Página web básica');
    await page.fill('.modal input[placeholder=Precio]', '1200000');
    await page.click('.modal footer button:has-text("Guardar")');
    await page.waitForSelector('text=PRO-');
    await page.click('button:has-text("Registrar venta")');
    await page.click('.modal footer button:has-text("Registrar venta")');
    await page.waitForSelector('text=VEN-');
    await page.screenshot({ path: path.join(OUT, '14-ventas-cliente.png'), fullPage: true });
    // Reporte
    await page.click('.tabs button:has-text("Solicitudes y reporte")');
    await page.click('button:has-text("Generar nueva versión")');
    await page.waitForSelector('text=v1');
    // Kanban
    await page.click('#nav button[data-view=pipeline]');
    await page.waitForSelector('.kanban .kcol');
    assert.strictEqual(await page.locator('.kanban .kcol').count(), 10);
    const card = page.locator('.kcol[data-stage=NUEVO] .kcard').first();
    await card.locator('select').selectOption('CONTACTADO');
    await page.waitForSelector('.toast.ok');
    await page.waitForSelector('.kcol[data-stage=CONTACTADO] .kcard');
    await page.screenshot({ path: path.join(OUT, '15-kanban.png'), fullPage: true });
    // Seguimientos, ventas, auditoría, usuarios, configuración
    for (const v of ['seguimientos', 'actividades', 'ventas', 'proyectos', 'duplicados', 'archivos', 'auditoria', 'usuarios', 'configuracion']) {
      await page.click('#nav button[data-view=' + v + ']');
      await page.waitForFunction(() => !document.querySelector('#viewBody .loading'), null, { timeout: 8000 });
      assert.ok(!(await page.isVisible('#viewBody .notice.error')), 'error en vista ' + v);
    }
    await page.click('#nav button[data-view=auditoria]');
    await page.waitForSelector('#viewBody .tbl');
    await page.screenshot({ path: path.join(OUT, '16-auditoria.png'), fullPage: true });
  });

  await run('Super Admin: acceso sin sesión muestra el login', browser, {}, async (page) => {
    await page.goto(BASE + '/?page=admin');
    await page.waitForSelector('#viewLogin:not([hidden])');
    assert.ok(!(await page.isVisible('#viewApp')));
    await page.evaluate(() => sessionStorage.setItem('wp_admin_token', 'f'.repeat(64)));
    await page.reload();
    await page.waitForSelector('#viewLogin:not([hidden])');
  });

  await browser.close();
  server.close();
  const failed = results.filter(r => !r.ok);
  results.forEach(r => console.log((r.ok ? '  ✔ ' : '  ✘ ') + r.name + (r.ok ? '' : '\n      ' + String(r.err && r.err.stack || r.err).split('\n').slice(0, 4).join('\n      '))));
  console.log('\n' + (results.length - failed.length) + ' / ' + results.length + ' pruebas de interfaz OK · capturas en tests/output/');
  process.exit(failed.length ? 1 : 0);
})();

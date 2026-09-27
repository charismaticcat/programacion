/**
 * Servidor local de desarrollo: sirve Index.html / Admin.html y ejecuta las
 * llamadas google.script.run contra el código real (servicios simulados).
 *
 *   node tests/dev-server.js            → http://localhost:8787  y  /?page=admin
 *   Usuario: admin@webpaya.test / ClaveSegura2026
 *
 * Solo para desarrollo y pruebas: los datos viven en memoria.
 */
'use strict';
const http = require('http');
const { loadProject } = require('./harness/load-gas');
const F = require('./fixtures');

const PORT = Number(process.env.PORT || 8787);
const ADMIN = { email: 'admin@webpaya.test', pass: 'ClaveSegura2026', name: 'Admin Local' };

function createSystem() {
  const p = loadProject({ echoErrors: true });
  p.asOwner();
  const log = p.call('setup');
  const code = log.find(l => l.indexOf('CÓDIGO DE CONFIGURACIÓN') >= 0).match(/: ([0-9A-F-]{14}) /)[1];
  const ss = p.env.globals.SpreadsheetApp.create('CCH_PRUEBA');
  const sh = ss.getSheetByName('Hoja 1');
  sh.name = 'Hoja1';
  [F.HEADER].concat(F.ROWS).forEach((r, i) => { sh.rows[i] = r.slice(); });
  p.env.state.props.IMPORT_SOURCE_SPREADSHEET_ID = ss.getId();
  p.call('importBase');
  p.asAnonymous();
  p.call('adm_createSuperAdmin', code, ADMIN.email, ADMIN.name, ADMIN.pass);
  return p;
}

const SHIM = `<script>
(function () {
  function runner() {
    var ok = null, fail = null, proxy;
    proxy = new Proxy({}, {
      get: function (t, prop) {
        if (prop === 'withSuccessHandler') return function (fn) { ok = fn; return proxy; };
        if (prop === 'withFailureHandler') return function (fn) { fail = fn; return proxy; };
        if (prop === 'withUserObject') return function () { return proxy; };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          var s = ok, f = fail;
          fetch('/rpc', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fn: prop, args: args }) })
            .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
            .then(function (x) { if (x.ok) { if (s) s(x.j.result); } else if (f) f(new Error(x.j.message)); })
            .catch(function (e) { if (f) f(e); });
        };
      }
    });
    return proxy;
  }
  window.google = { script: {} };
  Object.defineProperty(window.google.script, 'run', { get: runner });
})();
</script>`;

function start(port, cb) {
  let p = createSystem();
  const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://localhost');
    if (req.method === 'GET' && (url.pathname === '/' || url.pathname === '/exec')) {
      const out = p.run('doGet(' + JSON.stringify({ parameter: Object.fromEntries(url.searchParams) }) + ')');
      // En Apps Script, addMetaTag('viewport') lo agrega HtmlService; aquí se inyecta igual.
      const html = out._html.replace('<head>', '<head><meta name="viewport" content="width=device-width, initial-scale=1">' + SHIM);
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(html);
      return;
    }
    if (req.method === 'POST' && (url.pathname === '/rpc' || url.pathname.indexOf('/__test/') === 0)) {
      let body = '';
      req.on('data', c => { body += c; });
      req.on('end', () => {
        try {
          const msg = JSON.parse(body || '{}');
          let result;
          if (url.pathname === '/rpc') {
            result = p.call(msg.fn, ...(msg.args || []));
          } else if (url.pathname === '/__test/backdate') {
            p.run(`(function(){ const s = Auth.requireDraft(${JSON.stringify(msg.token)}); SheetService.update('SOLICITUDES', s.SOLICITUD_ID, { INICIADA_EN: '2020-01-01T00:00:00' }); })()`);
            result = true;
          } else if (url.pathname === '/__test/reset') {
            p = createSystem();
            result = true;
          } else if (url.pathname === '/__test/rows') {
            result = p.rows(msg.table);
          }
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ result: result === undefined ? null : result }));
        } catch (e) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ message: e.message }));
        }
      });
      return;
    }
    res.writeHead(404);
    res.end('No encontrado');
  });
  server.listen(port, () => cb && cb(server));
  return server;
}

if (require.main === module) {
  start(PORT, () => {
    console.log('Formulario:  http://localhost:' + PORT + '/');
    console.log('Super Admin: http://localhost:' + PORT + '/?page=admin  (' + ADMIN.email + ' / ' + ADMIN.pass + ')');
  });
}

module.exports = { start, ADMIN };

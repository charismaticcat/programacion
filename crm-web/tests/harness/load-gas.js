/**
 * Carga todos los archivos .gs de src/ en un contexto aislado (como Apps Script)
 * con los servicios simulados de gas-mocks.js.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { createEnvironment } = require('./gas-mocks');

const SRC = path.join(__dirname, '..', '..', 'src');

function loadProject(options) {
  options = options || {};
  const env = createEnvironment(options);
  const htmlFiles = {};
  fs.readdirSync(SRC).filter(f => f.endsWith('.html')).forEach(f => { htmlFiles[f.replace(/\.html$/, '')] = fs.readFileSync(path.join(SRC, f), 'utf8'); });

  const HtmlService = {
    XFrameOptionsMode: { DEFAULT: 'DEFAULT', ALLOWALL: 'ALLOWALL' },
    createHtmlOutputFromFile(name) {
      if (!(name in htmlFiles)) throw new Error('No existe el archivo HTML ' + name);
      const content = htmlFiles[name];
      return { getContent: () => content };
    },
    createTemplateFromFile(name) {
      if (!(name in htmlFiles)) throw new Error('No existe el archivo HTML ' + name);
      const tpl = { _src: htmlFiles[name] };
      tpl.evaluate = () => {
        const html = tpl._src
          .replace(/<\?!=\s*([\s\S]*?)\s*;?\s*\?>/g, (m, expr) => String(context.__evalExpr(expr, tpl)))
          .replace(/<\?=\s*([\s\S]*?)\s*;?\s*\?>/g, (m, expr) => escapeHtml(String(context.__evalExpr(expr, tpl))));
        const out = {
          _html: html, title: '', getContent: () => html,
          setTitle: (t) => { out.title = t; return out; }, addMetaTag: () => out, setXFrameOptionsMode: () => out
        };
        return out;
      };
      return tpl;
    }
  };
  const escapeHtml = (s) => s.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const sandbox = Object.assign({}, env.globals, {
    HtmlService,
    Date, Math, JSON, Object, Array, String, Number, Boolean, RegExp, Error, TypeError, Promise, Map, Set,
    encodeURIComponent, decodeURIComponent, isFinite, parseInt, parseFloat, isNaN
  });
  const context = vm.createContext(sandbox);
  const files = fs.readdirSync(SRC).filter(f => f.endsWith('.gs')).sort();
  files.forEach(f => {
    const code = fs.readFileSync(path.join(SRC, f), 'utf8');
    vm.runInContext(code, context, { filename: f });
  });
  // Evaluador de expresiones de plantillas (<?!= include_('X') ?>, <?= pageTitle ?>)
  vm.runInContext('this.__evalExpr = function (expr, tpl) { with (tpl) { return eval(expr); } };', context);

  const api = {
    env, context, files,
    /** Llama una función global (como google.script.run). Serializa como Apps Script. */
    call(fn, ...args) {
      const f = context[fn];
      if (typeof f !== 'function') throw new Error('Función no expuesta: ' + fn);
      if (/_$/.test(fn)) throw new Error('google.script.run no permite funciones privadas: ' + fn);
      const safeArgs = JSON.parse(JSON.stringify(args));
      const res = f.apply(null, safeArgs);
      return res === undefined ? undefined : JSON.parse(JSON.stringify(res));
    },
    /** Evalúa código dentro del proyecto (para inspección en pruebas). */
    run(code) { return vm.runInContext(code, context); },
    asOwner() { env.state.activeUser = env.state.effectiveUser; },
    asAnonymous() { env.state.activeUser = ''; },
    sheet(table) {
      return vm.runInContext('SheetService.resetCache(); SheetService.sheet(' + JSON.stringify(table) + ')', context);
    },
    rows(table) {
      return JSON.parse(JSON.stringify(vm.runInContext('SheetService.resetCache(); SheetService.readAll(' + JSON.stringify(table) + ', {includeDeleted:true})', context)));
    }
  };
  return api;
}

module.exports = { loadProject, SRC };

/**
 * Genera docs/BASE_DE_DATOS.md a partir de src/Schema.gs (fuente única).
 *   node tests/gen-db-doc.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { loadProject } = require('./harness/load-gas');

const p = loadProject();
const schema = p.run('SCHEMA');
const control = p.run('CONTROL_COLUMNS');
const enums = p.run('ENUMS');
const contentKeys = p.run('CONTENT_KEYS');
const TYPES = { s: 'Texto', t: 'Texto largo', n: 'Número', m: 'Dinero (COP)', b: 'Sí/No', d: 'Fecha', dt: 'Fecha y hora', e: 'Lista', j: 'JSON', x: 'Secreto' };
const esc = (s) => String(s || '').replace(/\|/g, '\\|');

const out = [];
out.push('# Base de datos (Google Sheets)');
out.push('');
out.push('> Documento generado automáticamente desde `src/Schema.gs` con `node tests/gen-db-doc.js`. No lo edite a mano.');
out.push('');
out.push('## Archivos');
out.push('');
out.push('| Spreadsheet | Hojas |');
out.push('|---|---|');
out.push('| `CRM_WEB_DB` | ' + Object.keys(schema).filter(t => !schema[t].auditDb).map(t => '`' + t + '`').join(', ') + ' |');
out.push('| `CRM_WEB_AUDITORIA` | `AUDITORIA` (solo se agregan filas) |');
out.push('');
out.push('**Reglas comunes**');
out.push('');
out.push('- Fila 1 = encabezados. El código ubica las columnas **por nombre**, no por posición: se pueden agregar columnas al final, pero **no renombre, ordene ni borre filas** a mano.');
out.push('- Todas las celdas usan formato de texto plano (`@`): no se pierden ceros (`01112`) ni se convierten fechas.');
out.push('- Los textos que empiezan por `= + - @` se guardan con apóstrofo para impedir fórmulas.');
out.push('- Borrado lógico: `ELIMINADO = SI`. Sí/No se guarda como `SI`/`NO`.');
out.push('- Columnas **editables** = se pueden cambiar desde el Super Admin (decisión D11). Las demás las mantiene el sistema.');
out.push('');
out.push('**Columnas de control** (en las tablas marcadas con ✓):');
out.push('');
out.push('| Columna | Tipo | Descripción |');
out.push('|---|---|---|');
control.forEach(c => out.push('| `' + c[0] + '` | ' + TYPES[c[1]] + ' | ' + esc(c[3]) + ' |'));
out.push('');
Object.keys(schema).forEach(t => {
  const def = schema[t];
  out.push('## ' + t);
  out.push('');
  out.push(esc(def.desc) + (def.control ? ' Columnas de control: ✓.' : ''));
  out.push('');
  out.push('| Columna | Tipo | Editable | Descripción |');
  out.push('|---|---|---|---|');
  def.columns.forEach(c => {
    let desc = c[3] || c[2];
    if (c[5] && enums[c[5]]) desc += ' · Valores: ' + enums[c[5]].filter(Boolean).map(v => '`' + v + '`').join(', ');
    if (c[5] === 'cat:MUNICIPIOS') desc += ' · Valores: `NEIVA`, `PITALITO`, `PALERMO`, `RIVERA`';
    out.push('| `' + c[0] + '` | ' + TYPES[c[1]] + ' | ' + (c[4] ? 'Sí' : '') + ' | ' + esc(desc) + ' |');
  });
  out.push('');
});
out.push('## Claves de CONTENIDO_WEB');
out.push('');
out.push('| Clave | Pregunta | Sección del reporte |');
out.push('|---|---|---|');
Object.keys(contentKeys).forEach(k => out.push('| `' + k + '` | ' + esc(contentKeys[k].label) + ' | ' + contentKeys[k].section + ' |'));
out.push('');

const file = path.join(__dirname, '..', 'docs', 'BASE_DE_DATOS.md');
fs.writeFileSync(file, out.join('\n'));
console.log('Generado', file);

# Diagnóstico del código — Foro Educativo Neiva 2026

Generado el 2026-08-25 a partir de los archivos reales clonados con `clasp clone`
(scriptId `1CD6ccfWvr0gnnxQK6qHGLsN07RPeUpgEDvk4kZsjX3WlJZctlIH076mR`).

## ⚠️ Aviso operativo (no es un problema de código)

`Código.js` contiene una función `programarEnvioAccesos28Agosto()` que, de haber
sido ejecutada, deja programado un disparador (trigger) que envía por correo los
códigos de acceso a las 37 IE el **28 de agosto de 2026 a las 5:45 a.m.**. No es
posible confirmar desde este entorno si ese trigger quedó activo — se revisa en
el editor de Apps Script, en el ícono de reloj ("Activadores" / "Triggers").
Conviene revisarlo antes de esa fecha.

## 1. Qué hace cada archivo y flujo de datos

`doGet()` (Código.js) identifica la IE por el parámetro `?t=TOKEN` y sirve
`Index.html`, un único documento con 12 `<section class="pantalla">` que
`App.html` muestra/oculta vía JavaScript (SPA de una sola página, sin
recargas). El usuario ingresa un código → `validarAccesoIE()` en el servidor
lo valida contra la hoja `AccesosIE` → las respuestas de las 3 sesiones se
guardan primero en `localStorage` y se sincronizan a la hoja `AvancesForo`
(`guardarAvanceForo`, con upsert por `ID_FORO`) → al enviar definitivamente
(`enviarForoDefinitivo`) se generan gráficos de participación, un informe en
Google Docs/PDF (`generarInformeFEM`) y se envía por correo
(`enviarInformeFEM`).

| Archivo | Líneas | Rol |
|---|---|---|
| `Código.js` | 6411 | Backend: acceso/sesión, lectura de "Oficiales", AvancesForo, generación de accesos, correos, informe PDF, evidencias en Drive |
| `App.html` | 5745 | JS del cliente: las 3 sesiones, autenticación, autoguardado, evidencias, plenaria |
| `Index.html` | 2254 | Estructura HTML de las 12 pantallas; incluye `CSS` y `App` |
| `CSS.html` | 1828 | Estilos por sección |
| `appsscript.json` | 9 | `timeZone: America/Bogota`, `access: ANYONE_ANONYMOUS` ✅ |

## 2. Código muerto, duplicado o roto

- **🐞 Bug real:** `generarInformeFEM()` (Código.js:6236) llama a
  `obtenerTokenPorIdForo_()`, que no está definida en ningún archivo del
  proyecto. El PDF se crea correctamente en Drive, pero la función revienta
  justo después al intentar guardar `ID_INFORME`/`ID_PDF_INFORME` en
  `AccesosIE`. Pendiente de arreglar.
- **Función duplicada:** `diagnosticarOficiales()` definida dos veces
  (Código.js:725 y :749). La segunda pisa silenciosamente a la primera —
  la primera es código muerto.
- **Función duplicada:** `actualizarURLsAccesoIE()` definida dos veces
  (Código.js:4601 y :4862) con lógica distinta. Solo la segunda corre.
- **Bloque de servidor pegado por error en el cliente:** `App.html:2166-2370`
  (~205 líneas) contiene copias de `probarValidacion1234` y
  `probarCodigoIncorrecto1234` que usan `SpreadsheetApp`/`Logger.log` — APIs
  que no existen en el navegador. Muerto de nacimiento; seguro de borrar.
- **Llamada duplicada:** en el login (`App.html` ~5240),
  `aplicarDatosAccesoFEM()` se invoca dos veces seguidas.
- **`console.log` de diagnóstico temporal confirmado:** `App.html` ~5547,
  dentro de `subirEvidenciasFEM()`, marcado explícitamente
  `DIAGNOSTICO TEMPORAL`.
- **Funciones de prueba mezcladas con producción** (ejecutables manualmente
  desde el editor, no parte del flujo real): `probarInstituciones`,
  `probarGuardarAvanceForo`, `crearAccesoPrueba1234`, `probarValidacion1234`,
  `verificarCuentaEnvio`, `enviarCorreoPruebaIE1234`,
  `programarCorreoPruebaEn3Minutos`, `enviarCorreoPruebaActual`,
  `probarCodigoIncorrecto1234`, `probarCatalogoIE`.

No se encontraron arquitecturas completas en competencia (p. ej. dos flujos
paralelos de todo el formulario) — la mezcla es de restos puntuales.

## 3. ¿El tamaño es un problema real?

Google Apps Script no soporta `import`/`export` de Node.js, pero sí permite
varios archivos `.js` en un mismo proyecto: todos comparten un único espacio
de funciones globales, así que dividir es puramente organizativo (como
separar capítulos de un libro) y no cambia el comportamiento. Igual con los
HTML: `include('CSS')` / `include('App')` solo pegan texto antes de enviarlo
al navegador; se puede incluir `Sesion1.html`, `Sesion2.html`, etc. sin
afectar el resultado final ni el rendimiento.

El conteo de líneas exagera el tamaño real: buena parte del archivo usa un
estilo "una expresión por línea" que en formato compacto ocuparía la mitad.
El código ya está bien seccionado con comentarios `/* ===== SECCIÓN ===== */`
y funciones con nombres claros. No hay funciones gigantescas salvo
`generarAccesosIE` y `generarInformeFEM`, que son largas por lo que hacen.
El problema real es la comodidad de navegación en el editor con miles de
líneas, no una falla estructural.

### Opciones, de menor a mayor esfuerzo

1. **Mover las funciones de prueba** a `Pruebas.js` (o borrarlas). Bajo
   riesgo, reduce `Código.js` varios cientos de líneas.
2. **Dividir `Código.js` por responsabilidad**: `Config.js`,
   `Instituciones.js`, `Accesos.js`, `AvancesForo.js`, `Informes.js`,
   `Evidencias.js`. Riesgo bajo-medio, se valida con `clasp push` + pruebas.
3. **Dividir `App.html` en partials incluidos** por sesión/flujo. Mismo
   principio; conviene probar el orden de los `<script>` aunque en la
   práctica las funciones JS se izan y casi nunca importa.
4. **No tocar nada todavía** — opción legítima: el código funciona, dividir
   archivos no arregla bugs por sí solo, solo mejora la comodidad futura.

## 4. Recomendación

No hace falta reescribir nada. Orden sugerido:

1. Arreglar el bug de `obtenerTokenPorIdForo_` (rompe la generación de
   informes).
2. Revisar el trigger del 28 de agosto.
3. Limpieza de bajo riesgo: borrar el bloque muerto de `App.html`, unificar
   las funciones duplicadas, quitar el `console.log` de diagnóstico.
4. Solo si después de eso el archivo sigue siendo incómodo de editar,
   considerar dividir en varios `.js`/`.html` (opciones 1–3 de arriba).

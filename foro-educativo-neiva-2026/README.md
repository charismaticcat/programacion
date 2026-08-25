# Foro Educativo Institucional — Neiva 2026

Formulario web construido con Google Apps Script (backend + HTML/CSS/JS servidos
desde Google) para recolectar las respuestas de las 37 instituciones educativas
oficiales de Neiva en el Foro Educativo Municipal 2026.

> **Nota (actualizado 2026-08-25):** los 5 archivos reales ya están clonados
> en esta carpeta vía `clasp clone` y quedaron commiteados en el repo. El
> `.clasp.json` apunta al scriptId del proyecto real. Para bajar cambios
> hechos desde el navegador usa `clasp pull`; para subir cambios editados
> aquí usa `clasp push` (ver la sección de clasp más abajo).

## Estado del proyecto

Este documento describe el proyecto real de Google Apps Script llamado
**"Foro Educativo 3.1"** (scriptId `1CD6ccfWvr0gnnxQK6qHGLsN07RPeUpgEDvk4kZsjX3WlJZctlIH076mR`),
identificado como el que las instituciones usan realmente (a diferencia de otras
copias/versiones que existen en la misma cuenta de Google: "Foro 3.2", "Foro 4.0",
"foro 5", etc. — esas son ramas de desarrollo antiguas, no la que está en uso).

### Archivos

| Archivo | Tipo en Apps Script | Contenido |
|---|---|---|
| `Código.js` | Script | Backend: `doGet`, validación de acceso, guardado en Sheets, Drive, Gmail |
| `Index.html` | HTML | Página principal, incluye a `App` y `CSS` |
| `App.html` | HTML | Todo el JavaScript del cliente (formulario, sesiones, evidencias, plenaria) |
| `CSS.html` | HTML | Estilos |
| `appsscript.json` | Manifiesto | Configuración del proyecto (permisos, tipo de despliegue) |

### Correcciones ya aplicadas (2026-08-24/25)

- Se quitaron 5 usos del operador `...` (spread), sintaxis que el runtime de
  Apps Script de este proyecto rechazaba con `SyntaxError: Unexpected token '...'`
  y que impedía cargar el formulario por completo.
- Se agregó una validación en `subirEvidenciasFEM()` (cliente) para que, si la
  sesión no tiene un `idForo` válido, se muestre un mensaje claro en vez de un
  error técnico (`Cannot read properties of null`).
- Se corrigió `appsscript.json`: tenía `"access": "MYSELF"`, lo que significaba
  que **solo la cuenta propietaria podía usar la aplicación**, jamás terceros.
  Se cambió a `"access": "ANYONE_ANONYMOUS"`.
- Se creó una implementación (deployment) nueva de tipo Web App con esa
  configuración corregida.
- Quedan en el código dos bloques de diagnóstico temporal (`console.log`) en
  el flujo de login y de evidencias — se pueden quitar una vez que se confirme
  que todo funciona de forma estable con usuarios reales.

### Pendiente de verificar

- Confirmar con un usuario externo real (no el propietario, en incógnito o
  cuenta distinta) que la subida de evidencias a Drive funciona sin pedir
  autorización.
- Decidir si se sigue usando la URL `/dev` (siempre sirve el código más
  reciente del editor, pero antes exigía login de Google por el `access:
  MYSELF`) o la URL `/exec` de la implementación publicada nueva — con el
  fix de acceso público, ambas deberían funcionar para cualquier usuario,
  pero `/exec` es la práctica estándar recomendada para producción porque no
  cambia si alguien edita el código a medias.
- Revisar el dato corrupto en la hoja "Oficiales" del Google Sheet real
  (una fila con "Holman Steven Mujica Gutierrez" en la columna de
  institución, que rompe el conteo esperado de 37 IE oficiales).

---

## Cómo conectar este repositorio a tu cuenta de Google (clasp)

Este proyecto no vive en GitHub como código independiente — vive dentro de
Google Apps Script, asociado a tu cuenta de Google. `clasp` es la herramienta
oficial de Google que permite editar ese código desde un editor normal (VS
Code, este repo, etc.) y subir/bajar cambios hacia Apps Script.

**Importante: cada persona que use este repo debe hacer login con SU PROPIA
cuenta de Google.** El login de clasp no se comparte ni se sube a GitHub — es
un archivo que vive únicamente en tu computadora (`~/.clasprc.json`), fuera de
este repositorio, y `.gitignore` ya está configurado para nunca subirlo por
accidente.

### 1. Instalar Node.js

Si no lo tienes, descárgalo de [nodejs.org](https://nodejs.org) (versión LTS).
Verifica la instalación abriendo una terminal y escribiendo:

```
node --version
```

### 2. Instalar clasp

```
npm install -g @google/clasp
```

Verifica con:

```
clasp --version
```

### 3. Habilitar la API de Apps Script en tu cuenta

Antes de poder usar `clasp push`/`clasp pull`, tu cuenta de Google necesita
tener habilitada la API de Apps Script (es un interruptor de seguridad, se
activa una sola vez):

1. Entra a **https://script.google.com/home/usersettings**
2. Activa el interruptor **"Apps Script API"**

### 4. Iniciar sesión con clasp

```
clasp login
```

Esto abre tu navegador y pide iniciar sesión con Google. **Usa la cuenta que
es dueña del proyecto del Foro** (o que tenga acceso de edición sobre él).
Acepta los permisos que pide (son los necesarios para que clasp pueda leer y
escribir tu proyecto de Apps Script).

### 5. Conectar este repo al proyecto real

Este repo ya incluye un archivo `.clasp.json` apuntando al `scriptId` correcto
del proyecto real ("Foro Educativo 3.1"). Solo necesitas pararte dentro de la
carpeta del repo (`foro-educativo-neiva-2026/`) y ejecutar:

```
clasp status
```

Si ves listados los 5 archivos (`Código.js`, `App.html`, `Index.html`,
`CSS.html`, `appsscript.json`) sin errores, la conexión quedó bien.

### 6. Bajar el código real más reciente (por si algo cambió desde afuera)

```
clasp pull
```

Esto sobreescribe los archivos locales con lo que esté guardado en Apps
Script en este momento — úsalo antes de empezar a editar, para no perder
cambios hechos desde el navegador.

### 7. Subir tus cambios

Después de editar cualquier archivo en este repo:

```
clasp push
```

Esto sube el contenido al editor de Apps Script. **Ojo:** esto actualiza el
código "en el editor" (`@HEAD`), pero si la URL que usan las instituciones es
una implementación publicada (`/exec`) con una versión fija, necesitas además
crear una nueva versión de esa implementación para que el cambio llegue a
producción:

```
clasp deploy --description "descripción corta del cambio"
```

Esto crea una implementación nueva. La URL pública se mantiene si editas una
implementación existente desde el navegador (Implementar > Administrar
implementaciones); `clasp deploy` sin más argumentos crea una implementación
nueva con una URL de deployment distinta — para actualizar una implementación
existente en vez de crear una nueva, usa:

```
clasp deploy --deploymentId <ID_DE_LA_IMPLEMENTACION_EXISTENTE> --description "..."
```

### Comandos útiles adicionales

```
clasp open-script          # abre el proyecto en el editor web de Apps Script
clasp list-deployments     # lista todas las implementaciones y sus versiones
clasp list-scripts         # lista todos los proyectos de Apps Script de tu cuenta
clasp tail-logs             # muestra los registros de ejecución en vivo
```

---

## Prompt para continuar el trabajo con Claude

El código actual es funcional pero fue construido de forma incremental,
mezclando varias versiones y ramas de desarrollo, con dos archivos muy
grandes (`Código.js` ~6300 líneas, `App.html` ~5900 líneas) que concentran
casi toda la lógica. Antes de decidir si conviene reestructurarlo, vale la
pena evaluarlo con calma.

Copia y pega esto en una conversación nueva con Claude (una vez que los
archivos reales ya estén en este repo vía `clasp clone`/`clasp pull`) cuando
quieras retomar el trabajo:

```
Tengo un proyecto de Google Apps Script (formulario web para un Foro
Educativo Institucional) que ya funciona, pero fue construido de forma
incremental a lo largo de varias sesiones de trabajo con IA, sin mucha
planificación. El repositorio tiene 5 archivos: Código.js (backend, ~6300
líneas), App.html (todo el JavaScript del cliente, ~5900 líneas), Index.html,
CSS.html y appsscript.json.

Antes de tocar nada, ayúdame a entender el estado actual:

1. Lee los 5 archivos completos y hazme un resumen de qué hace cada uno y
   cómo se relacionan (flujo de datos entre backend y cliente).
2. Identifica si hay código muerto, funciones duplicadas, o lógica que ya
   no se usa (sé que hubo varias iteraciones previas, así que es probable
   que haya restos de versiones anteriores).
3. Evalúa si el tamaño actual de Código.js y App.html es un problema real
   para mantenerlo, o si es manejable así. Si crees que conviene dividir
   el código en módulos más pequeños, explícame las opciones concretas
   (recordando que Google Apps Script no soporta imports/exports de
   módulos como Node.js — todos los archivos .gs comparten un mismo
   espacio de funciones globales, y los archivos HTML se combinan con
   include()).
4. No asumas que hay que reescribir todo desde cero. Primero quiero
   entender qué tan grave es la situación actual y qué opciones tengo,
   con sus ventajas y desventajas, antes de decidir un plan.

Trátame como alguien no técnico que está aprendiendo sobre la marcha —
explica los términos técnicos que uses.
```

Esto le pide a Claude que **primero entienda y evalúe**, sin asumir de entrada
que hay que reescribir o subdividir — la decisión de si conviene dividir el
proyecto en módulos, o dejarlo como está, debe tomarse después de ver el
diagnóstico real, no antes.

# Auditoría técnica — Foro Educativo Institucional (FEI/FEM) 3.1 Neiva 2026
### Base para el diseño de "Foro Educativo Comunal Neiva 2026" (unidad principal = GRUPO)

Proyecto auditado: clon clasp en `FEI_3.1_ref/` — `Código.js` (backend, 135 funciones), `Pruebas.js`
(91 funciones, mezcla de pruebas reales y utilidades de producción/reparación), `App.html` (191
funciones, JS de frontend), `Index.html` (estructura/pantallas), `CSS.html` (estilos),
`appsscript.json` (manifest). Todo el código fue leído íntegramente (Read/Grep en bloques). Las citas
de código son textuales, con archivo y línea aproximada.

---

## 1. Resumen de arquitectura general

### 1.1 Organización general

- **Monolito de Apps Script "V8" (`runtimeVersion: "V8"`)**, un único `SPREADSHEET_ID` como base de
  datos (Google Sheets) más 2 spreadsheets satélite creados dinámicamente ("Análisis FEM 2026" y
  "Participación, Valoración y Percepción por Grupo — FEM 2026"). No hay separación en librerías ni en
  otros proyectos: todo vive en un solo Apps Script con 5 archivos de código/markup.
- **`doGet(e)` (Código.js:229)** es el único punto de entrada HTTP. Enruta por parámetro de query, sin
  router formal:
  - `?asistencia=ID_FORO` → `paginaAsistenciaQR_(idForoAsistencia)` (página pública mínima de firma QR).
  - `?valoracion=ID_FORO` → `paginaValoracionFEM_(idForoValoracion)` (página pública mínima de cierre/valoración).
  - `?t=TOKEN` (o `token`/`TOKEN`) → sirve `Index.html` como plantilla (`HtmlService.createTemplateFromFile("Index")`),
    resolviendo en el propio `doGet` el nombre de la IE, el logo y el título del encabezado a partir de la
    hoja `AccesosIE` (búsqueda lineal por TOKEN, sin índice).
  - Sin parámetros → misma plantilla `Index`, en "modo desarrollo" (sin token), usada solo para pruebas manuales.
  - El manifest (`appsscript.json`) despliega la web app como `"executeAs": "USER_DEPLOYING"` y
    `"access": "ANYONE_ANONYMOUS"` — cualquiera con el enlace puede ejecutar el script con los permisos
    de quien lo publicó, sin login de Google. Toda la seguridad depende exclusivamente del TOKEN+código
    a nivel de aplicación (ver 1.2), no de Google Identity.
- **Patrón de inclusión de HTML**: `include(nombre)` (Código.js:216) hace
  `HtmlService.createHtmlOutputFromFile(nombre).getContent()`, y se invoca desde `Index.html` con
  `<?!= include('CSS'); ?>` (línea 18) y `<?!= include('App'); ?>` (línea 2741). Es decir: **un único HTML
  servido** compuesto por `Index.html` (estructura + variables de plantilla `<? ?>`/`<?!= ?>`) +
  `CSS.html` (hoja de estilos completa, se incluye como texto plano) + `App.html` (todo el JS de
  frontend, ~9500 líneas en un solo `<script>`). No hay build step, bundlers ni módulos ES — es JS
  clásico con cientos de funciones globales.
- Variables de plantilla inyectadas por el servidor antes de cargar `App.html`: `TOKEN_ACCESO`,
  `NOMBRE_IE_ACCESO`, `LOGO_IE_URL`, `DATOS_ACCESO_FEM` (Index.html:2728-2741) — así el frontend sabe si
  está en "modo token" sin llamar al servidor primero.
- **17 pantallas** controladas por clases CSS `visible`/oculto y JS (`cambiarPantalla`, App.html:1191),
  no rutas de navegador: `pantallaAccesoFEM`, `pantallaBloqueoHorarioFEM`, `pantallaBienvenida`,
  `pantallaRecorrido`, `pantallaConsentimiento`, `pantallaCaracterizacion`, `pantallaResumen`,
  `pantallaSesion1`, `pantallaSesion2`, `pantallaSesion3`, `pantallaPreguntaSesion4`, `pantallaSesion4`
  (= "Sesión Propia"), `pantallaEvidencias`, `pantallaPlenaria`, `pantallaFinalFEM`, `pantallaCierreFEM`,
  `pantallaDespedidaFEM` (Index.html).
- **Guardado**: autoguardado local (`localStorage`, `guardarBorradorLocal`/`restaurarBorradorLocal` en
  App.html) + sincronización periódica al servidor (`sincronizarBorradorConServidor`,
  `programarSincronizacion`, `activarGuardadoAutomatico`) que llama a `guardarAvanceForo` (UPSERT, ver
  1.3). El servidor nunca genera el `ID_FORO`; siempre viene de `AccesosIE` (comentario explícito en
  Código.js:2135 "ID_FORO YA NO SE GENERA EN EL SERVIDOR").

### 1.2 Modelo de seguridad (acceso, dispositivo, sesión)

Cadena real de llamadas, de más externo a más interno:

1. **`doGet` → `validarAccesoIE(token, codigo, dispositivoId, forzar)`** (Código.js:2814, llamada desde
   el cliente vía `google.script.run`, no desde `doGet` directamente — `doGet` solo resuelve el nombre/
   logo de la IE por el token para pintar la pantalla de acceso).
   - Busca el `token` en `AccesosIE` (columna `TOKEN`), valida `codigo` contra `CODIGO_ACCESO` +
     `CODIGO_CONTINGENCIA_1..4` (hasta 5 códigos válidos por IE), exige `ID_FORO` no vacío, rechaza
     `ESTADO` = `BLOQUEADO`/`INACTIVO`, soporta bloqueo temporal por horario (`HABILITAR_DESDE`, una
     fecha futura en la celda bloquea el ingreso con código `BLOQUEADO_POR_HORARIO`), y valida que la IE
     exista en el catálogo `Oficiales` salvo `TIPO = "PRUEBA"` (fila de prueba controlada a mano).
   - Al final llama a `reclamarSesionCodigo_(token, codigo, dispositivoId, idForo, forzar)`.
2. **`reclamarSesionCodigo_`** (Código.js:2692) — control de dispositivo/sesión, con `LockService`:
   - Clave de sesión: `obtenerClaveSesionCodigo_` (Código.js:2662) → `"FEM_SESION_FORO_" +
     Utilities.base64EncodeWebSafe(idForo || token+"|"+codigo)`.
   - Estado guardado en **`PropertiesService.getScriptProperties()`** (no en Sheets) como un **array de
     hasta `MAX_SESIONES_SIMULTANEAS_IE = 4`** objetos `{deviceId, tokenSesion, ultimaActividad,
     esPrincipal}` — hasta 4 dispositivos (colaboradores) pueden estar conectados a la vez con el mismo
     código.
   - El **primer** dispositivo que reclama un cupo para un `idForo` queda `esPrincipal:true` de forma
     permanente (es el único que llega a Plenaria y puede enviar el informe definitivo); los siguientes
     son colaboradores.
   - Si ya hay 4 cupos y llega un 5º sin `forzar`, responde `{ok:false, codigo:"SESION_YA_ABIERTA"}`; con
     `forzar:true` expulsa al dispositivo de actividad más antigua que **no** sea el principal.
   - **No hay expiración por inactividad** ("Ya NO hay temporizador de inactividad", comentario
     Código.js:2653): una sesión vive indefinidamente hasta que se libera explícitamente o se fuerza un
     takeover.
3. **`mantenerSesionCodigo_`** (Código.js:2753) — heartbeat periódico desde el cliente
   (`iniciarLatidoSesionFEM_`, App.html:1558) que actualiza `ultimaActividad`; si el cupo ya no existe
   (otro dispositivo lo tomó) responde `SESION_NO_AUTORIZADA`.
4. **`liberarSesionCodigo_`** (2799) quita el cupo del dispositivo actual (o borra la propiedad si era el
   último). **`transferirResponsablePrincipalFEM`** (2777) permite que el principal ceda su rol a otro
   colaborador conectado.
5. **`sesionActivaPorIdForo_`** (8427) es la comprobación de guardia usada por `enviarRespuestasSesion` y
   `enviarForoDefinitivo` antes de escribir — vuelve a leer el mismo array de `PropertiesService` y
   verifica `deviceId`+`tokenSesion`. Tenía un bug crítico en producción corregido en este mismo archivo
   (documentado in-line, ver §5 y §7).
6. **`dispositivoId`** se genera en el cliente y se persiste en `localStorage`:
   `obtenerDispositivoIdFEM()` (App.html:9478) → clave `"foroEducativoNeiva2026_dispositivo"`, valor
   `"FEM-DEV-"+Date.now().toString(36)+"-"+random`. No hay ninguna verificación server-side de que el
   `dispositivoId` no esté falsificado — es un identificador de conveniencia, no un secreto.
7. **`LockService.getScriptLock()`** se usa como exclusión mutua en TODAS las funciones que escriben
   estado compartido: `reclamarSesionCodigo_`, `mantenerSesionCodigo_`, `transferirResponsablePrincipalFEM`,
   `liberarSesionCodigo_`, `registrarAsistenciaQR`, `guardarAvanceForo`, `generarAccesosIE`,
   `generarInformeFEM`, `finalizarFormularioFEM`, `guardarValoracionFEM`, `enviarRespuestasSesion`,
   `enviarForoDefinitivo` (timeouts de `waitLock` entre 10s y 30s según la función).

### 1.3 Modelo de persistencia (Sheets) y UPSERT

- **Un único spreadsheet principal** (`SPREADSHEET_ID`) con hojas fijas (`Oficiales`, `AvancesForo`,
  `AccesosIE`, `Participacion`, `AsistenciaQR`, `Valoración FEMI2026`, `EnviosInformeDiferidos`) **más
  una hoja dinámica por cada IE** que envía datos (`nombreHojaIE_(nombre)`, Código.js:5165 — sanitiza el
  nombre y lo trunca a 95 caracteres; `guardarEnHojaIE_`, 5203, crea/actualiza esa hoja con las mismas
  cabeceras de `AvancesForo`). Con 37 IE esto ya son 37 pestañas adicionales — ver riesgo de escalabilidad
  en §7 para un modelo por GRUPO con más instituciones.
- **UPSERT por `ID_FORO`**: `buscarFilaPorIdForo_` (Código.js:1690) busca la fila; `guardarAvanceForo`
  (2111) hace update in-place si existe, `appendRow` si no. Nunca borra ni duplica filas para el mismo
  `ID_FORO`.
- **Fusión de campos entre dispositivos concurrentes** (Código.js:2288-2328): como hasta 4 dispositivos
  pueden guardar el DOM completo del formulario, `guardarAvanceForo` **fusiona** el JSON entrante con el
  ya guardado campo por campo — un campo entrante solo sobreescribe al guardado si trae contenido real o
  si el campo no existía antes. Evita que el último `save()` de un dispositivo borre en silencio lo que
  otro dispositivo ya guardó.
- **Generación de identificadores**:
  - `ID_FORO` = `Utilities.getUuid()` (UUID completo, generado una sola vez por `generarAccesosIE`,
    Código.js:4675-4692, con verificación de unicidad en memoria `idsForoUsados`).
  - `TOKEN` = `Utilities.getUuid().replace(/-/g,"")` (Código.js:4651-4672).
  - `CODIGO_ACCESO`/contingencia = 5 caracteres alfanuméricos sin ambigüedad (`ABCDEFGHJKMNPQRSTUVWXYZ23456789`)
    prefijados `"FEM-"` (Código.js:4600-4648, función interna de `generarAccesosIE`). Existe una **segunda
    implementación** casi idéntica pero con formato distinto, `generarCodigoAcceso_` (Código.js:5157):
    `"FEM-"+PALABRA+4 caracteres` (con lista de 12 palabras) — dos generadores de código de acceso
    incompatibles entre sí conviviendo en el mismo archivo (ver §7).
- Cada dato "en vivo" del formulario se guarda además como **blob JSON completo** en la columna `DATOS`
  de `AvancesForo` (y de la hoja por IE) — `{"institucion":..., "campos":{"id":{"valor":...,"tipo":...}, ...}}`
  — que es la fuente que leen `obtenerDatosGuardadosPorIdForo_`, `generarInformeFEM`, y los compiladores
  de grupo.
- **Spreadsheets satélite**, creados perezosamente con `SpreadsheetApp.create(...)` y cuyo ID se guarda
  en `PropertiesService.getScriptProperties()` para no volver a crearlos:
  - `"Análisis FEM 2026 — Foro Educativo Institucional Neiva"` — clave de propiedad
    `CLAVE_PROP_SPREADSHEET_ANALISIS = "SPREADSHEET_ANALISIS_ID"` (Código.js:8558,
    `obtenerSpreadsheetAnalisisFEM_`, 8562). Contiene la hoja `"Respuestas Totales"`
    (`HOJA_ANALISIS_TOTALES`) que es una **copia analítica desnormalizada** de `AvancesForo` (todas las
    respuestas + participación + valoración en una sola fila por IE) y `"Gráficos"`
    (`HOJA_ANALISIS_GRAFICOS`) con gráficos nativos de Sheets.
  - `"Participación, Valoración y Percepción por Grupo — FEM 2026"` — clave `CLAVE_PROP_SPREADSHEET_GRUPOS_FEM_`
    (Pruebas.js:5331, `obtenerSpreadsheetGruposFEM_`). Contiene hojas `"Participación"`, `"Valoración"` y
    (por patrón) `"Percepción"`, una fila por IE agrupada y ordenada por G1-G6, con gráficos de columnas
    por grupo.

### 1.4 Organización de Google Drive

- **Carpeta raíz pública del FEM**: `DRIVE_CARPETA_FEM_ID = "1IqcFgQUSKocvGX3JwvNOu-xJzt0gfKc8"`
  (Código.js:61) — contiene una subcarpeta por IE (`crearCarpetaIE_`, 5251, "obtener o crear" por
  nombre) con el PDF final del informe, y (nuevo, orientado a grupo) una subcarpeta `"Grupo GN"` por cada
  grupo con dos subcarpetas hijas `"Informes enviados de grupo GN"` e `"Informes editables de grupo GN"`
  (`crearEstructuraCarpetasGrupoFEM_`, 8971). Los dos spreadsheets satélite también se mueven a esta
  carpeta tras crearse.
- **Carpeta privada de Docs editables**: `DRIVE_CARPETA_EDITABLES_FEM_ID = "1047FoYnCOsVakUl-7pCtvwpeceWM2DpB"`
  — solo el equipo de calidad tiene acceso; ahí va el Google Doc editable de cada informe (nunca
  compartido públicamente), separado del PDF (que sí se comparte "cualquiera con el enlace").
  Ver comentario explícito Código.js:62-71 sobre el incidente de seguridad que motivó esta separación.
- **IDs de logos y plantillas, todos hardcodeados como constantes** (Código.js:86-107):
  `TEMPLATE_INFORME_ID` (ya no se usa — ver §7), `LOGO_ENCABEZADO_ID`, `LOGO_PIE_ID`, `MARCO_ACCESO_ID`,
  `FONDO_SINTESIS_BANDA_SUPERIOR_ID`, `FONDO_SINTESIS_BANDA_INFERIOR_ID`, `DISENADOR_LOGO_ID`.
- **Logo por IE**: no es un ID hardcodeado por institución, sino una columna `LOGO_ID` en `AccesosIE`
  (poblada a mano/una vez, `vincularLogosIE()` en Pruebas.js), resuelta en tiempo real por
  `obtenerLogoIdPorNombreIE_` (Código.js:5296) con comparación de nombre normalizado.
- **`hacerPublicoSiEsPosible_`** (Código.js:5256) intenta `file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,
  DriveApp.Permission.VIEW)` envuelto en try/catch silencioso — patrón usado en todo archivo público (PDF,
  fotos de evidencia, carpetas de IE).

### 1.5 Generación de informes y envío por correo

- **`generarInformeFEM(idForo, datosCliente)`** (Código.js:7124) construye el informe **desde cero** con
  `DocumentApp.create(...)` (no copia ninguna plantilla — la plantilla original `TEMPLATE_INFORME_ID`
  resultó ser un `.docx` no nativo y se abandonó, ver comentario 7131-7141). Arma manualmente: encabezado
  con 3 logos en tabla sin bordes, portada, tabla de caracterización, secciones de Sesión 1/2/3 con
  gráficos, listado de asistencia (`agregarListadoAsistenciaAlInforme_`), perfil/percepción
  (`agregarPerfilYPercepcionAlInforme_`), sesión propia opcional, bloque de firmas dinámico según
  responsables. Aplica ajuste de "título huérfano" (`aplicarKeepWithNextATitulosInforme_`, usa el
  **Advanced Service `Docs v1`** habilitado en `appsscript.json`) y exporta a PDF con
  `DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF)` (línea 7668).
- **Gráficos**: no son gráficos nativos de Sheets/Docs sino **imágenes generadas con
  `Charts.newColumnChart()`/`newBarChart()`** (`construirGraficoColumnas_`, `construirGraficoBarrasHorizontal_`,
  Código.js:5957-5998) insertadas como `Blob` dentro del Doc con `body.appendImage(blob)`.
- **Informes de grupo/municipales** reutilizan el mismo motor `DocumentApp` con las mismas paletas y
  helpers (`titulo1_`, `subtitulo_`, `parrafo_`, `tabla_`) redefinidos función a función (código
  duplicado, no factorizado en un helper común — ver §7).
- **Envío de correo**: exclusivamente `GmailApp.sendEmail(...)` desde la cuenta que ejecuta el
  deployment, remitente forzado `REMITENTE_FEM = "calidadeducacion@alcaldianeiva.gov.co"` (validado
  contra `GmailApp.getAliases()`/`Session.getEffectiveUser()`), con copia fija a `COPIAS_INFORME_FEM`
  (3 correos). **Cuota diaria compartida**: `enviarInformeFEM` (7992) detecta el error de cuota agotada
  (`esErrorCuotaCorreoAgotada_`) y en vez de fallar, registra un envío diferido en la hoja
  `EnviosInformeDiferidos` para reintento posterior (`reintentarEnviosInformeDiferidosFEM`, Pruebas.js) —
  patrón de cola simple sobre Sheets, sin ningún servicio de colas real.

### 1.6 Lógica ya existente de GRUPO (la más relevante para el proyecto nuevo)

Existe un **catálogo fijo de 6 grupos (G1–G6) con 6 IE cada uno = 36 de las 37 IE oficiales**, definido
**por duplicado** en dos archivos (deben mantenerse sincronizados a mano, ver comentario explícito):

- `GRUPOS_INSTITUCIONES` en **App.html:211** (frontend, usado por `obtenerGrupoInstitucion`, 274, para
  autocompletar el campo "grupo" en Caracterización).
- `GRUPOS_INSTITUCIONES_FEM_` en **Código.js:8810** (backend, usado por `obtenerGrupoInstitucionCatalogoFEM_`,
  8823, como catálogo de respaldo).

El **grupo "real" de una IE** no se toma directamente del catálogo sino de la respuesta que la propia IE
guardó en Caracterización, con fallback al catálogo:

```js
// Código.js:8844
function obtenerGrupoRealDeIEFEM_(nombreIE, idForo){
  if(idForo){
    try{
      const datosGuardados=obtenerDatosGuardadosPorIdForo_(idForo);
      const grupoGuardado=String(datosGuardados?.campos?.grupo?.valor||"").trim();
      if(grupoGuardado) return grupoGuardado;
    }catch(error){ ... }
  }
  return obtenerGrupoInstitucionCatalogoFEM_(nombreIE);
}
```

A partir de ahí, la cadena de generación por grupo es:

1. **`mapaGruposFEM_()`** (Pruebas.js:5084) recorre `AccesosIE`, excluye `TIPO=PRUEBA`, y clasifica cada
   IE con informe (`ID_PDF_INFORME` no vacío) en `{G1..G6, SIN_GRUPO}` usando `obtenerGrupoRealDeIEFEM_`.
2. **`crearEstructuraCarpetasGrupoFEM_(grupo)`** (Código.js:8971) crea/reutiliza la carpeta de Drive del
   grupo.
3. **`organizarInformesPorGrupoFEM()`** (Pruebas.js:5126) **copia** (nunca mueve) el PDF y el Doc editable
   más reciente de cada IE del grupo a esas carpetas.
4. **`generarDocumentoCompiladoGrupoFEM_(grupo, listaIEsConSesiones, idDocExistente)`** (Código.js:9088)
   arma un Doc único por grupo con las respuestas textuales (Sesión 1/2/3 + Sesión Propia) de cada IE,
   una tras otra — es un "compilado", no un análisis.
5. **`generarDocumentoAnalisisGrupoFEM_(grupo, analisis, idDocExistente)`** (Código.js:9178) arma un
   segundo Doc con 9 secciones cuantitativas (participantes, edad, fortalezas/debilidades por edad,
   comparativo) reutilizando `tallyOpciones_`, `categoriaEdad_`, `calcularDemografiaAsistentes_`,
   `construirGraficoColumnas_`/`BarrasHorizontal_`, `top3Texto_` — el mismo "idioma" que el informe
   individual.
6. **`generarInformeSintesisGrupoFEM_(datosGrupo, idDocExistente)`** (Código.js:9496) — el informe
   "de verdad" de síntesis grupal: portada con bandas decorativas + logos de las 6 IE, secciones por
   pregunta con prosa **redactada a mano y hardcodeada** (no generada por código) sobre "elementos
   comunes" y "particularidades coyunturales", tablas/gráficos para las 4 preguntas de selección
   múltiple, conclusiones y firmas fijas. Los datos de entrada (`DATOS_SINTESIS_GRUPO_1_FEM_` …
   `DATOS_SINTESIS_GRUPO_6_FEM_`, definidos en Pruebas.js) son objetos literales con prosa 100% escrita
   a mano para el FEM 2026 — **no hay ninguna lógica de generación automática de esa prosa**: es contenido
   editorial, no reutilizable como código (ver §6).
7. **`generarInformeSintesisMunicipalFEM_(datosMunicipio, idDocExistente)`** (Código.js:9720) — mismo
   patrón a escala municipal (consolida los 6 informes grupales en uno), añade un capítulo especial de
   percepción QR agregando los datos de las 4 IE que usaron el QR. También alimentado por un objeto
   literal hardcodeado (`DATOS_SINTESIS_MUNICIPAL_FEM_`, en Pruebas.js).
8. **`generarSpreadsheetGrupoFEM()`** (Pruebas.js:5586) arma el spreadsheet satélite de grupo con hojas
   `"Participación"`, `"Valoración"`, `"Percepción"` (tablas + gráficos nativos de Sheets por grupo).
9. **`generarDocumentosAnalisisPorGrupoFEM()`** (Pruebas.js:5957) orquesta el paso 5 para los 6 grupos.

**Conclusión clave para el proyecto nuevo**: la infraestructura de "agrupar por grupo" (carpetas,
copiado de informes, compilados, hoja de análisis por grupo) **sí está bien resuelta y es reutilizable
como patrón**, pero (a) el catálogo de grupos está hardcodeado y duplicado, (b) el grupo se calcula desde
la IE hacia arriba (bottom-up) en vez de ser la entidad raíz del modelo de datos, y (c) los informes de
síntesis "de verdad" (los más elaborados) no son generados automáticamente sino textos redactados a mano
por sesión de trabajo — el nuevo proyecto, si quiere generación automática de síntesis, necesita diseñar
esa capa desde cero (o mantenerla como proceso editorial asistido, no 100% automático).

---

## 2. Matriz de auditoría por archivo/función

Leyenda de la columna "Equivalente propuesto": nombres exactos de la nueva estructura CLASP
(`Code.gs, Config.gs, Access.gs, Session.gs, Data.gs, Instituciones.gs, Grupos.gs, Sesion1.gs,
ConectaEduca.gs, Informes.gs, Drive.gs, Correo.gs, Utils.gs, Tests.gs, Index.html, CSS.html, JS.html,
Components.html, Modal.html, InformeStyles.html`). Donde el nombre del archivo nuevo no permite inferir
con certeza su alcance (`Sesion1.gs`, `ConectaEduca.gs`) se marca **[supuesto]** e indica la duda.

### 2.1 Código.js — Acceso, validación, sesión y dispositivo

| Archivo 3.1 | Función | Propósito | Hojas/Carpetas/Plantillas | Reutilizar/Adaptar/No usar | Equivalente propuesto | Observaciones |
|---|---|---|---|---|---|---|
| Código.js:229 | `doGet(e)` | Router único por query param (`asistencia`, `valoracion`, `t`) y resolución de nombre/logo de IE por token | `AccesosIE` | Adaptar | `Code.gs` | Debe enrutar por GRUPO en vez de IE; añadir un modo de entrada institucional dentro del grupo si se conserva el detalle por IE |
| Código.js:216 | `include(nombre)` | Incluye HTML como texto plano | — | Reutilizar tal cual | `Code.gs` | Ahora deberá incluir `JS.html`, `Components.html`, `Modal.html`, `InformeStyles.html` además de `CSS.html` |
| Código.js:195 | `abrirSpreadsheet_()` | Apertura de Sheet con reintentos (error transitorio de Google) | `SPREADSHEET_ID` | Reutilizar tal cual | `Data.gs` | Patrón sólido, sin cambios |
| Código.js:2662 | `obtenerClaveSesionCodigo_` | Construye clave de `ScriptProperties` para la sesión activa | — (PropertiesService) | Reutilizar, adaptar clave | `Session.gs` | Cambiar semántica de clave: por `idGrupo`+`código` en vez de por IE |
| Código.js:2681 | `leerSesionesActivas_` | Lee array de cupos activos, compatible con formato legado | — | Reutilizar tal cual | `Session.gs` | — |
| Código.js:2692 | `reclamarSesionCodigo_` | Reclama cupo de dispositivo (máx. `MAX_SESIONES_SIMULTANEAS_IE=4`), marca `esPrincipal` | — | Reutilizar, renombrar constante | `Session.gs` | Para GRUPO probablemente se necesiten más de 4 cupos simultáneos (varias IE del grupo participando a la vez) — revisar el límite |
| Código.js:2750/2753 | `mantenerSesionCodigo`/`_` | Heartbeat de sesión | — | Reutilizar tal cual | `Session.gs` | — |
| Código.js:2777 | `transferirResponsablePrincipalFEM` | Transferencia de rol "responsable principal" a otro colaborador conectado | — | Reutilizar, renombrar | `Session.gs` | — |
| Código.js:2798/2799 | `liberarSesionCodigo`/`_` | Libera cupo de dispositivo | — | Reutilizar tal cual | `Session.gs` | — |
| Código.js:2814 | `validarAccesoIE` | Valida token+código+dispositivo, estado, bloqueo por horario, IE oficial vs prueba | `AccesosIE`, `Oficiales` | Adaptar fuertemente | `Access.gs` | Renombrar a `validarAccesoGrupo`; el "código de acceso" pasa a ser por GRUPO, no por IE; conservar exactamente la lógica de códigos de contingencia y bloqueo por horario |
| Código.js:8427 | `sesionActivaPorIdForo_` | Verifica que dispositivo+tokenSesion siga con cupo antes de escribir | — | Reutilizar tal cual (con el bug ya corregido) | `Session.gs` | Ver bug histórico documentado en el propio código, §5/§7 |
| Código.js:4217 | `generarAccesosIE` | Genera TOKEN/CODIGO_ACCESO/ID_FORO por IE, valida "exactamente 37 IE" | `Oficiales`→`AccesosIE` | Adaptar fuertemente | `Access.gs` | El nuevo proyecto genera accesos por GRUPO, no por IE; quitar el candado "exactamente 37" y reemplazar por conteo de grupos configurado |
| Código.js:3881 | `actualizarURLsAccesoIE` | Recalcula `URL_ACCESO`/`LINK_ACCESO` si cambia el deployment | `AccesosIE` | Reutilizar, adaptar nombre de hoja | `Access.gs` | — |
| Código.js:5127 | `asegurarColumnasAccesosIE_` | Crea/asegura columnas de `AccesosIE` (25 columnas) | `AccesosIE` | Adaptar el listado de columnas | `Access.gs` | Ver columnas propuestas para `GruposComunal` en §3 |
| Código.js:5148 | `filaAccesoPorToken_` | Busca fila por TOKEN | `AccesosIE` | Reutilizar tal cual | `Access.gs` | — |
| Código.js:5157 | `generarCodigoAcceso_` | 2ª implementación de generación de código (formato distinto a la de `generarAccesosIE`) | — | **No reutilizar** (duplicado) | — | Unificar en un solo generador; ver riesgo en §7 |
| Código.js:5170 | `inicializarHojasIE` | Crea hojas fijas con cabeceras si faltan | `AccesosIE`,`Participacion` | Adaptar | `Data.gs` | — |
| Código.js:1773 | `obtenerAccesoPorIdForo_` | Busca fila de `AccesosIE` por `ID_FORO`, devuelve objeto simple | `AccesosIE` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:8420 | `obtenerAccesoPorIdForoRaw_` | Igual, pero devuelve también `hoja`/`fila`/`mapa` para escribir de vuelta | `AccesosIE` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:5241 | `obtenerEstadoSesiones_` | Estado S1/S2/S3 enviadas, por `idForo` | `AvancesForo` | Adaptar | `Data.gs` | — |
| Código.js:8449 | `validarEnvioFinal_` | Verifica que existan respuestas obligatorias antes del envío definitivo | — | Adaptar al cuestionario nuevo | `Sesion1.gs` **[supuesto]** | Lista de campos obligatorios hardcodeada — reescribir 100% para el nuevo cuestionario |
| Código.js:8464 | `enviarForoDefinitivo` | Envío definitivo: UPSERT + marca ENVIADO + dispara actualizaciones de participación/gráficos/análisis | `AvancesForo`,`AccesosIE`,hoja por IE | Adaptar fuertemente | `Data.gs`/`Grupos.gs` | Candado de "un solo envío" debe pasar de ser por IE a ser (según diseño) por IE dentro del grupo o por grupo completo |
| Código.js:8110 | `finalizarFormularioFEM` | Cierre de formulario tras generar el informe: limpia sesión activa, marca ENVIADO en `AccesosIE` | `AccesosIE` | Adaptar | `Data.gs` | — |

### 2.2 Código.js — Participación y asistencia (incluye QR)

| Archivo 3.1 | Función | Propósito | Hojas/Carpetas/Plantillas | Reutilizar/Adaptar/No usar | Equivalente propuesto | Observaciones |
|---|---|---|---|---|---|---|
| Código.js:5188 | `inicializarParticipacion_` | Crea/asegura hoja `Participacion` con columnas por rol | `Participacion` | Reutilizar, renombrar a `ParticipacionComunal` | `Data.gs` | — |
| Código.js:5196 | `actualizarParticipacion_` | UPSERT de fila de participación por `ID_FORO` | `Participacion` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:5203 | `guardarEnHojaIE_` | Crea/actualiza la **hoja dedicada por IE** (una pestaña por institución) | hoja dinámica por IE | **No reutilizar tal cual** | `Data.gs`/`Instituciones.gs` | Ver riesgo de escalabilidad §7: con más IE por grupo, una pestaña por IE no escala bien; considerar hoja única con filtro/tabla dinámica |
| Código.js:5227 | `reordenarHojasPorIE_` | Reordena alfabéticamente las pestañas por IE | — | No reutilizar (si se descarta la pestaña por IE) | — | — |
| Código.js:5412 | `asegurarHojaAsistenciaQR_` | Crea/asegura columnas de `AsistenciaQR` (21 columnas) | `AsistenciaQR` | Reutilizar, adaptar nombre | `Data.gs` | Base para `ConectaEduca.gs` si ese módulo cubre asistencia/percepción — ver nota en §2.5 |
| Código.js:5436-5495 | `formatearFechaHoraFirma_`, `formatearFechaFotoEvidencia_`, `formatearFechaLargaEs_`, `formatearFechaSubidaFoto_`, `construirPieFotoEvidencia_` | Grupo de helpers de formato de fecha/pie de foto en español | — | Reutilizar tal cual (agrupar) | `Utils.gs` | Helpers triviales, sin lógica de negocio |
| Código.js:5500 | `registrarAsistenciaQR` | Alta de firma QR con 18 parámetros, deduplicación por documento, validaciones de formato | `AsistenciaQR` | Reutilizar fuertemente | `Data.gs` | Patrón de deduplicación por documento (no por dispositivo) es sólido y reutilizable |
| Código.js:5608 | `obtenerAsistentesQR_` | Lee todas las firmas de un `idForo` | `AsistenciaQR` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:5645 | `eliminarAsistenciaQRPorCambioMetodo` | Borra firmas QR si la IE cambia de método de asistencia (QR↔PDF) | `AsistenciaQR` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:5673 | `eliminarAsistenciaPDFPorCambioMetodo` | Borra el PDF de asistencia subido manualmente si cambia de método | Drive | Reutilizar tal cual | `Drive.gs` | — |
| Código.js:5690 | `contarAsistentesQR` | Cuenta firmas por `idForo` (usado para refrescar contador en vivo) | `AsistenciaQR` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:5712 | `obtenerListaAsistentesQR` | Lista completa de asistentes (para exportar/mostrar) | `AsistenciaQR` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:5728 | `agregarListadoAsistenciaAlInforme_` | Inserta la tabla de asistencia en el Doc del informe | — (DocumentApp) | Reutilizar tal cual | `Informes.gs` | — |
| Código.js:5790 | `categoriaEdad_` | Clasifica rango de edad en niño/adolescente/adulto | — | Reutilizar tal cual | `Utils.gs` | — |
| Código.js:5803 | `calcularDemografiaAsistentes_` | Tabula sexo×categoría de edad | `AsistenciaQR` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:5820 | `tallyOpciones_` | Conteo genérico de opciones seleccionadas (fortalezas/dificultades) | — | Reutilizar tal cual | `Utils.gs` | Muy reutilizable para cualquier pregunta de selección múltiple del nuevo cuestionario |
| Código.js:6267 | `paginaAsistenciaQR_` | HTML standalone de firma QR (independiente del formulario principal) | — | Reutilizar fuertemente | `ConectaEduca.gs` **[supuesto]** o `Access.gs` | Candidato natural si `ConectaEduca.gs` es el módulo de participación ciudadana/asistencia — confirmar alcance con el usuario |
| Código.js:6926 | `subirEvidenciasFEM` | Sube foto de evidencia a Drive | Drive (carpeta IE) | Reutilizar, adaptar carpeta destino a GRUPO | `Drive.gs` | — |
| Código.js:6944 | `subirAsistenciaPDF` | Sube PDF de asistencia manual (método alterno al QR) | Drive (carpeta IE) | Reutilizar tal cual | `Drive.gs` | — |
| Código.js:8340 | `enviarRespuestasSesion` | Envío parcial de una sesión (S1/S2/S3) con validación de sesión activa | `AvancesForo` | Adaptar al nuevo cuestionario | `Sesion1.gs` **[supuesto]** | — |

### 2.3 Código.js — Generación de informes / PDF / gráficos

| Archivo 3.1 | Función | Propósito | Hojas/Carpetas/Plantillas | Reutilizar/Adaptar/No usar | Equivalente propuesto | Observaciones |
|---|---|---|---|---|---|---|
| Código.js:5857-5917 | `estilizarTituloInforme_`, `estilizarCuerpoInforme_`, `estilizarSubtituloInforme_`, `aplicarKeepWithNextATitulosInforme_` | Helpers de estilo de párrafo del Doc + arreglo de "título huérfano" vía Docs API v1 | — (Advanced Service `Docs`) | Reutilizar tal cual (agrupar) | `Informes.gs` | Requiere mantener habilitado el servicio avanzado `Docs v1` en el manifest nuevo |
| Código.js:5957 | `construirGraficoColumnas_` | Genera imagen de gráfico de columnas (`Charts` service) | — | Reutilizar tal cual | `Informes.gs` | — |
| Código.js:5983 | `construirGraficoBarrasHorizontal_` | Ídem, barras horizontales | — | Reutilizar tal cual | `Informes.gs` | — |
| Código.js:5998 | `agregarGraficoConOtro_` | Inserta gráfico + tabla + texto libre de "Otro" | — | Reutilizar tal cual | `Informes.gs` | — |
| Código.js:6027 | `top3Texto_` | Top 3 de un tally en texto | — | Reutilizar tal cual | `Utils.gs` | — |
| Código.js:6041 | `obtenerSugerenciasValoracion_` | Extrae sugerencias abiertas de valoración por IE | `Valoración FEMI2026` | Adaptar | `Data.gs` | — |
| Código.js:6063 | `agregarPerfilYPercepcionAlInforme_` | Sección de perfil demográfico + percepción en el informe individual | — (DocumentApp) | Reutilizar tal cual | `Informes.gs` | — |
| Código.js:6246/6258 | `capitalizarNombreIE_`, `nombreIESinPrefijoInstitucional_` | Formato de nombre de IE ("I.E. X" → "X") | — | Reutilizar tal cual | `Utils.gs` | — |
| Código.js:7124 | `generarInformeFEM` | **Función central**: construye el Doc completo desde cero con `DocumentApp.create`, exporta a PDF | `AvancesForo`, carpeta IE, carpeta editables, logos | Reutilizar la arquitectura, reescribir el contenido | `Informes.gs` | 557 líneas; separar en sub-funciones por sección facilitaría adaptarlo al cuestionario del Foro Comunal |
| Código.js:7974 | `obtenerPdfInformeMasRecienteFEM_` | Resuelve el PDF más reciente por nombre de archivo en la carpeta de la IE | Drive | Reutilizar tal cual | `Drive.gs` | — |
| Código.js:8510 | `actualizarGraficosParticipacion_` | Refresca gráficos nativos de Sheets con la participación total | `Participacion` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:8523 | `actualizarGraficoHojaIE_` | Refresca gráfico en la hoja propia de cada IE | hoja por IE | Adaptar/eliminar según §7 | `Data.gs` | — |
| Código.js:8562 | `obtenerSpreadsheetAnalisisFEM_` | Crea/reutiliza spreadsheet satélite "Análisis FEM 2026" vía `ScriptProperties` | — | Reutilizar el patrón | `Informes.gs`/`Grupos.gs` | Buen patrón: mover a `InformesComunal` como hoja dentro del spreadsheet principal, no spreadsheet aparte, si se prefiere un solo archivo |
| Código.js:8584/8592 | `obtenerCabecerasAnalisisTotales_`, `asegurarHojaAnalisisTotales_` | Cabeceras/estructura de `Respuestas Totales` | — | Adaptar | `Data.gs` | — |
| Código.js:8606 | `obtenerValoracionPorIdForo_` | Lee valoración guardada de una IE | `Valoración FEMI2026` | Reutilizar tal cual | `Data.gs` | — |
| Código.js:8638 | `actualizarAnalisisFEMIndividual_` | Vuelca todas las respuestas + participación + valoración de una IE a `Respuestas Totales` | `AvancesForo`→`Respuestas Totales` | Reutilizar el patrón | `Data.gs`/`Informes.gs` | — |
| Código.js:8749 | `reordenarHojasAnalisisFEM_` | Reordena pestañas del spreadsheet de análisis | — | Reutilizar tal cual | `Data.gs` | — |
| Código.js:8765 | `actualizarGraficosAnalisisFEM_` | Refresca gráficos nativos del spreadsheet de análisis | `Gráficos` | Reutilizar tal cual | `Informes.gs` | — |

### 2.4 Código.js — Envío de correo

| Archivo 3.1 | Función | Propósito | Hojas/Carpetas/Plantillas | Reutilizar/Adaptar/No usar | Equivalente propuesto | Observaciones |
|---|---|---|---|---|---|---|
| Código.js:3466 | `construirCorreoAccesoIE_` | Arma asunto/texto/HTML del correo de acceso (token+código) | — | Reutilizar, adaptar copy | `Correo.gs` | — |
| Código.js:3539 | `enviarAccesosTodasIE` | Envía el correo de acceso a las 37 IE | `AccesosIE` | Adaptar a "todas las IE del grupo"/"todos los grupos" | `Correo.gs` | — |
| Código.js:3786 | `programarEnvioAccesos28Agosto` | Trigger programado (fecha hardcodeada) para el envío masivo | `AccesosIE` | **No reutilizar** (fecha de un evento pasado) | — | Reemplazar por un trigger genérico configurable |
| Código.js:7732 | `construirCorreoInformeFEM_` | Arma el correo con el informe adjunto (PDF) + enlaces a carpeta/valoración | — | Reutilizar, adaptar copy | `Correo.gs` | — |
| Código.js:7821/7876 | `construirCorreoRecordatorioValoracionFEM_` / `enviarRecordatorioValoracionFEM_` | Recordatorio de "cierre formal" (Valoración del Foro) | `AccesosIE` | Reutilizar el patrón | `Correo.gs` | — |
| Código.js:7901/7953 | `construirCorreoRecordatorioVencimientoFEM_` / `enviarRecordatorioVencimientoFEM_` | Recordatorio de vencimiento de plazo | `AccesosIE` | Reutilizar el patrón | `Correo.gs` | — |
| Código.js:7992 | `enviarInformeFEM` | **Envío central del informe**: adjunta PDF, gestiona cuota agotada como envío diferido, correo adicional al responsable | `AccesosIE`, Drive | Reutilizar fuertemente | `Correo.gs` | Manejo de cuota agotada (try/catch específico + cola en Sheets) es el patrón más valioso a preservar tal cual |
| Código.js:7691/7701/7718 | `asegurarHojaEnviosDiferidosFEM_`, `registrarEnvioInformeDiferidoFEM_`, `esErrorCuotaCorreoAgotada_` | Cola de reintento de correos por cuota agotada | `EnviosInformeDiferidos` | Reutilizar tal cual | `Correo.gs`/`Data.gs` | — |
| Código.js:8175 | `guardarValoracionFEM` | Guarda la Valoración del Foro (cierre formal) | `Valoración FEMI2026` | Adaptar cuestionario | `Data.gs` | — |
| Código.js:8160 | `asegurarHojaValoracionFEM_` | Cabeceras de `Valoración FEMI2026` | `Valoración FEMI2026` | Adaptar nombre de hoja | `Data.gs` | — |
| Código.js:8146 | `textoTieneSentidoFEM_` | Heurística anti-texto-basura en respuestas abiertas | — | Reutilizar tal cual | `Utils.gs` | — |
| Código.js:8270 | `enviarComprobanteParticipacionFEM` | Envía comprobante de participación por correo | `AccesosIE` | Reutilizar el patrón | `Correo.gs` | — |
| Código.js:6912 | `guardarValoracionYConfirmarFEM` | Guarda valoración + dispara confirmación | `Valoración FEMI2026` | Adaptar | `Data.gs` | — |
| Código.js:6695 | `paginaValoracionFEM_` | HTML standalone de valoración/cierre (mismo patrón que asistencia QR) | — | Reutilizar fuertemente | `Access.gs`/`ConectaEduca.gs` **[supuesto]** | — |

### 2.5 Código.js y Pruebas.js — TODO lo relacionado a GRUPO / consolidación / análisis FEM

Esta es la sección más relevante para "Foro Educativo Comunal Neiva 2026": son las únicas piezas de 3.1
que ya piensan en una unidad por encima de la IE.

| Archivo 3.1 | Función | Propósito | Hojas/Carpetas/Plantillas | Reutilizar/Adaptar/No usar | Equivalente propuesto | Observaciones |
|---|---|---|---|---|---|---|
| App.html:211 | `GRUPOS_INSTITUCIONES` (const) | Catálogo fijo G1-G6 → 6 IE cada uno, **frontend** | — | **No reutilizar tal cual** | `Grupos.gs` (servir desde backend) | Duplicado con el de Código.js; el nuevo proyecto debe tener un solo catálogo, editable en `ConfiguracionComunal` o `GruposComunal`, no hardcodeado en 2 archivos |
| App.html:274 | `obtenerGrupoInstitucion` | Busca el grupo de una IE en el catálogo del frontend | — | No reutilizar tal cual (mover a backend) | `Grupos.gs` | — |
| Código.js:8810 | `GRUPOS_INSTITUCIONES_FEM_` (const) | Mismo catálogo, **backend** | — | **No reutilizar tal cual** | `GruposComunal` (hoja) | Reemplazar por datos en Sheet: en el modelo nuevo el GRUPO es la entidad raíz, con sus IE como colección hija, no un array hardcodeado |
| Código.js:8819 | `normalizarNombreParaGrupoFEM_` | Normaliza nombre de IE para comparar (sin tildes/mayúsculas) | — | Reutilizar tal cual | `Utils.gs` | — |
| Código.js:8823 | `obtenerGrupoInstitucionCatalogoFEM_` | Busca grupo de una IE en el catálogo fijo | — | Adaptar a consulta contra `GruposComunal` | `Grupos.gs` | — |
| Código.js:8844 | `obtenerGrupoRealDeIEFEM_` | Grupo real = respuesta guardada por la IE, con fallback al catálogo | `AvancesForo` | Adaptar (invertir prioridad: en el nuevo modelo el grupo se asigna, no se autodeclara) | `Grupos.gs` | Patrón "preferir dato vivo, fallback a catálogo" es útil pero la relación IE↔Grupo debería ser autoritativa desde `GruposComunal`, no autodeclarada en el formulario |
| Código.js:8865 | `obtenerRespuestasSesionesParaCompilado_` | Extrae preguntas/respuestas S1/S2/S3 de una IE en formato `[etiqueta,valor]` para compilar | `AvancesForo` (vía `DATOS`) | Adaptar al nuevo cuestionario | `Grupos.gs`/`Sesion1.gs` **[supuesto]** | — |
| Código.js:8906 | `obtenerRespuestasSesionesDesdeAnalisisFEM_` | Respaldo: reconstruye S1/S2/S3 desde `Respuestas Totales` si `AvancesForo` perdió la fila | `Respuestas Totales` | Adaptar (función de recuperación ante incidentes, ver §7) | `Grupos.gs` | Nace de un incidente real de pérdida de datos — considerar backup periódico en el proyecto nuevo en vez de este parche |
| Código.js:8956 | `crearOFolderHija_` | Obtiene/crea subcarpeta de Drive (genérico) | Drive | Reutilizar tal cual | `Drive.gs` | — |
| Código.js:8971 | `crearEstructuraCarpetasGrupoFEM_` | Crea carpeta del grupo + subcarpetas "enviados"/"editables" | Drive (`DRIVE_CARPETA_FEM_ID`) | **Reutilizar fuertemente** | `Drive.gs`/`Grupos.gs` | Pieza más directamente aprovechable: es la estructura de carpetas que el nuevo proyecto necesita desde el día 1, ya por grupo |
| Código.js:8998 | `extraerSesionesDesdeDocEditableFEM_` | Recupera S1/S2/S3 leyendo directamente el Doc ya generado de una IE (fuente más confiable) | Drive (carpeta editables) | Adaptar (parche de incidente, ver §7) | `Grupos.gs` | Patrón "el Doc es la fuente de verdad si Sheets falla" es interesante pero frágil (depende de parsear texto del Doc por títulos exactos) |
| Código.js:9088 | `generarDocumentoCompiladoGrupoFEM_` | Doc editable que compila S1/S2/S3 (+Sesión Propia) de todas las IE de un grupo | Drive (carpeta grupo) | **Reutilizar fuertemente** | `Grupos.gs`+`Informes.gs` | — |
| Código.js:9178 | `generarDocumentoAnalisisGrupoFEM_` | Doc de análisis cuantitativo (9 secciones) por grupo | Drive (carpeta grupo) | **Reutilizar fuertemente** | `Grupos.gs`+`Informes.gs` | — |
| Código.js:9496 | `generarInformeSintesisGrupoFEM_` | Motor de render del informe de síntesis grupal (portada, secciones, gráficos, firmas) | Drive, logos, bandas decorativas | Reutilizar el **motor**, no el contenido | `Grupos.gs`+`Informes.gs` | El contenido (prosa) viene de fuera como parámetro — ver fila siguiente sobre los datos hardcodeados |
| Pruebas.js:6233-7242 | `generarInformeSintesisGrupo1FEM` … `generarInformeSintesisGrupo6FEM` (6 funciones) + `DATOS_SINTESIS_GRUPO_1_FEM_`…`_6_` (constantes) | Invocan el motor anterior con prosa **redactada a mano** para cada uno de los 6 grupos del FEM 2026 | Drive (carpeta grupo) | **No reutilizar el contenido**; reutilizar solo el patrón de invocación | `Grupos.gs` (solo la función orquestadora genérica) | Miles de líneas de texto editorial específico de 2026 — ver §6 |
| Código.js:9720 | `generarInformeSintesisMunicipalFEM_` | Motor de render del informe de síntesis municipal (agrega los 6 grupales + capítulo QR) | Drive, logos, bandas | Reutilizar el motor | `Grupos.gs`+`Informes.gs` | Para el Foro Comunal este sería el nivel "comunal" (varios grupos) — renombrar conceptualmente |
| Pruebas.js:7529 | `generarInformeSintesisMunicipalFEM` | Invoca el motor anterior con `DATOS_SINTESIS_MUNICIPAL_FEM_` (prosa hardcodeada) | Drive | No reutilizar el contenido | `Grupos.gs` | — |
| Pruebas.js:5084 | `mapaGruposFEM_` | Clasifica todas las IE con informe en `{G1..G6,SIN_GRUPO}` | `AccesosIE` | **Reutilizar fuertemente** (adaptar a consultar `GruposComunal`) | `Grupos.gs` | Pieza central para cualquier vista/reporte "por grupo" |
| Pruebas.js:5126 | `organizarInformesPorGrupoFEM` | Copia PDFs/Docs de cada IE a la carpeta de su grupo | Drive | **Reutilizar fuertemente** | `Grupos.gs`+`Drive.gs` | — |
| Pruebas.js:5231 | `resolverSesionesIEFEM_` | Resuelve S1/S2/S3 de una IE probando 3 fuentes en cascada (Doc→vivo→respaldo) | Drive, `AvancesForo`, `Respuestas Totales` | Adaptar (o simplificar si el nuevo proyecto no sufre el incidente de pérdida de datos) | `Grupos.gs` | Complejidad nacida de un incidente puntual — no reproducir la causa raíz (ver §7) en vez de solo copiar el parche |
| Pruebas.js:5262 | `compilarRespuestasPorGrupoFEM` | Orquesta `generarDocumentoCompiladoGrupoFEM_` para los 6 grupos | — | Reutilizar el patrón | `Grupos.gs` | — |
| Pruebas.js:5331 | `obtenerSpreadsheetGruposFEM_` | Crea/reutiliza spreadsheet satélite de grupo | `ScriptProperties` | Reutilizar el patrón (o fusionar en el spreadsheet principal) | `Grupos.gs`/`Data.gs` | — |
| Pruebas.js:5348 | `hojaLimpiaGrupoFEM_` | Limpia/crea hoja (quita gráficos y contenido) | — | Reutilizar tal cual | `Utils.gs` | — |
| Pruebas.js:5356 | `etiquetaGrupoFEM_` | Etiqueta legible ("Grupo G1" / "Sin grupo asignado") | — | Reutilizar tal cual | `Utils.gs` | — |
| Pruebas.js:5358 | `construirHojaParticipacionGrupoFEM_` | Hoja "Participación" agregada por grupo, con gráficos por rol | `Participacion` (vía datos vivos) | **Reutilizar fuertemente** | `Grupos.gs` | — |
| Pruebas.js:5410 | `construirHojaValoracionGrupoFEM_` | Hoja "Valoración" agregada por grupo | `Valoración FEMI2026` | **Reutilizar fuertemente** | `Grupos.gs` | — |
| Pruebas.js:5477 | `escribirBloquePercepcionFEM_` | Bloque de percepción (fortalezas/dificultades) por grupo de asistentes | `AsistenciaQR` | Reutilizar el patrón | `Grupos.gs` | — |
| Pruebas.js:5543 | `construirHojaPercepcionGrupoFEM_` | Hoja "Percepción" agregada por grupo | `AsistenciaQR` | **Reutilizar fuertemente** | `Grupos.gs` | — |
| Pruebas.js:5586 | `generarSpreadsheetGrupoFEM` | Orquesta las 3 hojas anteriores | — | Reutilizar el patrón | `Grupos.gs` | — |
| Pruebas.js:5957 | `generarDocumentosAnalisisPorGrupoFEM` | Orquesta `generarDocumentoAnalisisGrupoFEM_` para los 6 grupos | Drive | Reutilizar el patrón | `Grupos.gs` | — |
| Pruebas.js:3029 | `reconstruirAnalisisFEM` | Reconstruye `Respuestas Totales` para todas las IE | `AvancesForo`→`Respuestas Totales` | Reutilizar el patrón | `Grupos.gs`/`Data.gs` | Útil como utilidad de mantenimiento en el nuevo proyecto también |
| Pruebas.js:5750 | `restaurarAvancesForoDesdeAnalisisFEM` | Restaura `AvancesForo` desde el respaldo de análisis (recuperación del incidente 2026-09-05) | `Respuestas Totales`→`AvancesForo` | No reutilizar tal cual (es un parche puntual) | — | Documentar como lección aprendida, ver §7 |
| Pruebas.js:5126 (comentario) / 5084 | — | (contexto) | — | — | — | — |

### 2.6 Frontend — App.html (agrupado por módulo) e Index.html/CSS.html

| Archivo 3.1 | Función/bloque | Propósito | Reutilizar/Adaptar/No usar | Equivalente propuesto | Observaciones |
|---|---|---|---|---|---|
| App.html:9478 | `obtenerDispositivoIdFEM` | Genera/persiste `dispositivoId` en `localStorage` | Reutilizar tal cual | `JS.html` | Pieza de seguridad crítica del lado cliente |
| App.html:7043 | `inicializarAccesoFEM` | Pantalla de acceso: muestra IE/logo resuelto por servidor, oculta el resto | Adaptar a "grupo" | `JS.html` | — |
| App.html:7300 | `validarCodigoAccesoFEM` | Llama a `validarAccesoIE` (`google.script.run`), maneja `forzar`/takeover | Adaptar nombre de llamada | `JS.html` | — |
| App.html:1558 | `iniciarLatidoSesionFEM_` | Heartbeat periódico al servidor (`mantenerSesionCodigo`) | Reutilizar tal cual | `JS.html` | — |
| App.html:1341/1354 | `esPrincipalActualFEM_`, `avisarFinColaboracionSesion4_` | Lógica de UI para diferenciar responsable principal vs colaborador | Reutilizar el patrón | `JS.html` | — |
| App.html:211-321 | `GRUPOS_INSTITUCIONES` + `obtenerGrupoInstitucion` | Catálogo de grupos y búsqueda, en el cliente | No reutilizar (mover a backend, ver 2.5) | `Grupos.gs` (server) | — |
| App.html:6294-6984 | Bloque de borrador local: `construirDatosBorrador`, `guardarBorradorLocal`, `aplicarDatosDeBorrador_`, `restaurarBorradorLocal`, `restaurarAvanceDesdeServidor_`, `programarSincronizacion`, `sincronizarBorradorConServidor`, `activarGuardadoAutomatico`, `mostrarEstadoGuardado` | Autoguardado local + sync periódica con `guardarAvanceForo` | **Reutilizar fuertemente** (patrón robusto) | `JS.html` | Uno de los módulos más maduros del frontend; aplicable tal cual al nuevo formulario |
| App.html:2430-3120 | Contadores de palabras/validación por pregunta: `contarPalabras`, `calcularColorBarraProgreso_`, `actualizarMarcaMinimoCumplido_`, `actualizarContadorPregunta1/2` | UX de mínimos de palabras por pregunta abierta | Reutilizar el patrón, adaptar mínimos | `Components.html`/`JS.html` | — |
| App.html:3378-5875 | Módulos de Sesión 1/2/3 (contadores, colapsos, validación, selección múltiple con "Otro") | Toda la lógica específica del cuestionario FEI 3.1 | **No reutilizar el contenido**, sí el patrón de UI (acordeón, validación, "Otro") | `Sesion1.gs`(server)/`Components.html`(UI) **[supuesto]** | Cuestionario 100% específico del FEM 2026 — el nuevo Foro Comunal tendrá preguntas propias |
| App.html:5874-6141 | Sesión Propia/4 (línea temática libre de la IE) | Módulo opcional de tema propio | Adaptar concepto | `Sesion1.gs` **[supuesto]** | — |
| App.html:8216-8778 | Módulo QR/Firmas/Evidencias: `actualizarQRAsistencia_`, `actualizarContadoresAsistenciaQR_`, `inicializarSelectorMetodoAsistencia_`, `inicializarFirmasEnVivo_`, `actualizarPreviaFotoEvidencia_`, `actualizarEstadoEvidencias` | Firma QR en vivo + evidencia fotográfica | **Reutilizar fuertemente** | `ConectaEduca.gs`(server) **[supuesto]** /`Components.html`(UI) | — |
| App.html:7742-8216 | `inicializarPlenaria`, `inicializarValoracionFEM_`, `mostrarModalCierreDefinitivoFEM_`, `mostrarPantallaDespedidaFEM_` | Cierre del formulario: plenaria, valoración final, despedida | Adaptar contenido | `Modal.html`/`JS.html` | — |
| App.html:9029-9235 | `renderizarPlenaria`, `renderCaracterizacionHTML`, `renderGraficoParticipacionHTML`, `escapeHtml` | Render de resumen HTML en la pantalla de Plenaria | Reutilizar el patrón | `Components.html` | — |
| Index.html (completo) | Estructura de 17 `<section id="pantalla...">` + variables de plantilla del servidor | Maquetación de todas las pantallas | Adaptar contenido, reutilizar patrón de plantilla | `Index.html` | Server-side templating (`<? ?>`/`<?!= ?>`) se conserva igual en Apps Script |
| CSS.html (completo, 4544 líneas) | Hoja de estilos única, paleta institucional (verde/amarillo), diseño responsive | Reutilizar base, adaptar marca/paleta | `CSS.html` | Sin lógica, solo revisar si la nueva identidad visual del Foro Comunal reutiliza la paleta verde/amarillo de la Alcaldía de Neiva |

### 2.7 Helpers menores agrupados (no ameritan fila individual)

| Grupo de funciones | Archivo | Propósito | Reutilizar/Adaptar | Equivalente propuesto |
|---|---|---|---|---|
| `normalizarCabeceraInstituciones_`, `buscarColumnaInstituciones_`, `obtenerInstituciones`, `normalizarNombreIE_`, `buscarInstitucionOficial_`, `obtenerInstitucionesJSON`, `obtenerSedesDeIE_`, `diagnosticarOficiales` | Código.js | Lectura tolerante del catálogo `Oficiales` (columnas por nombre, no por posición fija) | Reutilizar el patrón de lectura tolerante | `Instituciones.gs` |
| `obtenerMapaCabeceras_`, `mapaHoja_`, `obtenerCampoFormulario_`, `normalizarValorHoja_`, `extraerRespuestasSesiones_`, `buscarFilaPorIdForo_` | Código.js | Utilidades genéricas de lectura/escritura de Sheets por cabecera | Reutilizar tal cual | `Data.gs`/`Utils.gs` |
| `prepararHojaAvancesForo`, `obtenerAvanceForo`, `obtenerDatosGuardadosPorIdForo_` | Código.js | Estructura y lectura de `AvancesForo` | Adaptar cabeceras | `Data.gs` |
| `crearCarpetaIE_`, `hacerPublicoSiEsPosible_`, `validarCarpetaDrive_`, `urlPublicaLogoDrive_`, `obtenerLogoIdPorNombreIE_` | Código.js | Utilidades de Drive (carpetas, permisos, URLs públicas) | Reutilizar tal cual | `Drive.gs` |
| `numeroLocalizado_` | Código.js | `Number()` tolerante a coma decimal española | Reutilizar tal cual | `Utils.gs` |
| ~85 funciones de `Pruebas.js` no listadas arriba (pruebas unitarias reales: `probarInstituciones`, `probarGuardarAvanceForo`, `crearAccesoPrueba1234`, `probarValidacion1234`, `probarSesionesSimultaneasYFusionDatos`, etc., y utilidades de diagnóstico/reparación puntual: `repararEmailIEOficialesFEM`, `backfillFechaEnvioCorreoInformeDesdeGmailFEM`, `investigarEnvioInformeIE`, etc.) | Pruebas.js | Pruebas manuales ejecutadas desde el editor + scripts de reparación de incidentes puntuales de 2026 | Reutilizar la **metodología** de prueba manual, no el contenido (specífico de IE/incidentes de 2026) | `Tests.gs` |
| ~150 funciones de `App.html` de UI pura no listadas arriba (botones, acordeones, validación de campos específicos del cuestionario FEI, plenaria por sesión, etc.) | App.html | Interacción de UI del formulario FEI 3.1 | Reutilizar patrones de interacción (acordeón, contador de palabras, guardado automático); reescribir contenido | `JS.html`/`Components.html`/`Sesion1.gs` |

---

## 3. Inventario de hojas de Sheets detectadas en 3.1

### 3.1 Spreadsheet principal (`SPREADSHEET_ID`)

| Hoja (nombre real) | Constante | Columnas detectadas | Uso |
|---|---|---|---|
| `Oficiales` | `HOJA_OFICIALES` | Encabezados en **fila 5** (no fila 1): `INSTITUCIÓN/SEDE`, `CODIGO DANE`, `DIRECCIÓN`, `SECTOR`, `COMUNA`, `ZONA`, `E-MAIL INSTITUCIONAL`, `RECTOR (A)` | Catálogo oficial de 37 IE (mayúsculas=IE central, minúsculas=sedes), mantenido manualmente por la SEM |
| `AvancesForo` | `HOJA_AVANCES` | `ID_FORO, INSTITUCION, DANE, FECHA_INICIO, ULTIMA_ACTUALIZACION, ESTADO, S1_P1, S1_P2, S2_P1, S2_P2_ACCION_1..5, S2_P3, S2_P4, S2_P5, S3_P1, S3_P2_ACCION_1..5, S3_P3, S3_P4, SESION_PROPIA_TITULO, SESION_PROPIA_OBJETIVO, SESION_PROPIA_LINEAS_JSON, S1_ENVIADA, S2_ENVIADA, S3_ENVIADA, FECHA_ENVIO_S1/S2/S3, FECHA_ENVIO_DEFINITIVO, ID_INFORME, ID_PDF_INFORME, DATOS` (ver `obtenerCabecerasAvancesForo`, Código.js:1258) | Tabla maestra de avances del formulario, UPSERT por `ID_FORO`; `DATOS` guarda el JSON completo del formulario |
| `AccesosIE` | `HOJA_ACCESOS` | `ID_ACCESO, IE, DANE, CODIGO_ACCESO, CODIGO_CONTINGENCIA_1..4, TOKEN, URL_ACCESO, LINK_ACCESO, ID_FORO, ESTADO, TOKEN_SESION, DISPOSITIVO_ID, FECHA_GENERACION, FECHA_PRIMER_ACCESO, ULTIMA_ACTIVIDAD, FECHA_ENVIO, EMAIL_IE, EMAIL_RESPONSABLE, TIPO, S1_ENVIADA, S2_ENVIADA, S3_ENVIADA, ID_INFORME, ID_PDF_INFORME, LOGO_ID, HABILITAR_DESDE, ENVIADO_POR, FECHA_ENVIO_CORREO_INFORME` (Código.js:5131) | Tabla de control de acceso por IE: credenciales, estado, referencias a informe generado |
| `Participacion` | `HOJA_PARTICIPACION` | `IE, ID_FORO, FECHA, Rector(a), Coordinador(a), Docentes, Tutor PTA PFI/3.0, Orientador(a), Estudiantes, Padres/madres/acudientes, Personal administrativo, Egresados, Sector productivo, Otros, Total` (Código.js:5190) | Conteo de participantes declarados por IE, UPSERT por `ID_FORO` |
| `AsistenciaQR` | `HOJA_ASISTENCIA_QR` | `ID_FORO, IE, NOMBRE_COMPLETO, SEXO, EDAD, TIPO_ASISTENCIA, CARGO, ROL_FORO, JORNADA, SEDE, FORTALEZAS, FORTALEZA_OTRO, DIFICULTADES, DIFICULTAD_OTRO, NUMERO_DOCUMENTO, CORREO, TELEFONO, CONSENTIMIENTO, FECHA, HORA, DISPOSITIVO_ID` (Código.js:5416) | Registro de firmas de asistencia individuales vía QR, con encuesta de percepción |
| `Valoración FEMI2026` | `HOJA_VALORACION_FEM` | `ID_FORO, IE, FECHA, P1_DIALOGO_REFLEXION, P2_PARTICIPACION, P3_IDEAS_PROPUESTAS, P4_SATISFACCION_INSTRUMENTO, NOTA_PROMEDIO, P1_MEJORA, P2_MEJORA, P3_MEJORA, P4_MEJORA, P5_SUGERENCIAS` (Código.js:8164) | Encuesta de cierre formal del Foro por IE |
| `EnviosInformeDiferidos` | `HOJA_ENVIOS_DIFERIDOS_FEM` | `ID_FORO, FECHA_REGISTRO, REINTENTADO` (Código.js:7696) | Cola de reintento de correos de informe cuando se agota la cuota diaria de Gmail |
| *(dinámica)* una hoja por cada IE, nombre = `nombreHojaIE_(IE)` | — | Mismas cabeceras que `AvancesForo` (sin `DATOS` truncado) | Copia de trabajo por institución; se reordenan alfabéticamente tras `AvancesForo` (`reordenarHojasPorIE_`) |

### 3.2 Spreadsheet satélite "Análisis FEM 2026 — Foro Educativo Institucional Neiva"

ID guardado en `ScriptProperties["SPREADSHEET_ANALISIS_ID"]` (`CLAVE_PROP_SPREADSHEET_ANALISIS`).

| Hoja | Constante | Columnas | Uso |
|---|---|---|---|
| `Respuestas Totales` | `HOJA_ANALISIS_TOTALES` | Cabeceras de `AvancesForo` (sin `DATOS`) + `PART_RECTOR, PART_COORDINADOR, PART_DOCENTES, PART_TUTOR_PTA, PART_ORIENTADOR, PART_ESTUDIANTES, PART_PADRES, PART_ADMINISTRATIVOS, PART_EGRESADOS, PART_SECTOR, PART_OTROS, TOTAL_PARTICIPANTES, TOTAL_ASISTENTES_QR, VAL_NOTA_PROMEDIO, VAL_P1..P4, VAL_P1_MEJORA..P4_MEJORA, VAL_P5_SUGERENCIAS` (`obtenerCabecerasAnalisisTotales_`, Código.js:8584) | Vista desnormalizada de una fila por IE con todo (respuestas + participación + valoración) — usada como **respaldo** cuando `AvancesForo` pierde filas |
| `Gráficos` | `HOJA_ANALISIS_GRAFICOS` | 2 columnas (`Estamento`, `Participantes`) + gráficos nativos | Panel de gráficos de participación agregada |
| *(dinámica)* hoja por IE dentro de este spreadsheet (Código.js:8692) | — | Similar a `Respuestas Totales` filtrada por IE | Vista individual dentro del documento de análisis |

### 3.3 Spreadsheet satélite "Participación, Valoración y Percepción por Grupo — FEM 2026"

ID guardado en `ScriptProperties["..."]` (`CLAVE_PROP_SPREADSHEET_GRUPOS_FEM_`).

| Hoja | Columnas | Uso |
|---|---|---|
| `Participación` | `IE, Grupo` + 11 columnas de `ETIQUETAS_PARTICIPACION_ANALISIS_` + `Total`, más bloques auxiliares para gráficos (Pruebas.js:5360) | Participación agregada por grupo (G1-G6 + SIN_GRUPO), con un gráfico de columnas por rol comparando grupos |
| `Valoración` | `IE, Grupo, Estado, Nota promedio, P1 diálogo y reflexión, P2 participación, P3 ideas y propuestas, P4 satisfacción del instrumento, P1 mejora, P2 mejora, P3 mejora, P4 mejora, P5 sugerencias` (Pruebas.js:5412) | Valoración de cierre agregada por grupo |
| `Percepción` (por patrón `construirHojaPercepcionGrupoFEM_`, Pruebas.js:5543) | Bloques de conteo de fortalezas/dificultades por grupo de edad (`escribirBloquePercepcionFEM_`) | Percepción cualitativa agregada por grupo, alimentada por `AsistenciaQR` |

---

## 4. Inventario de carpetas/IDs de Drive y plantillas detectadas

| Constante | ID (Código.js) | Uso | Resolución |
|---|---|---|---|
| `SPREADSHEET_ID` | `1OiBPO8BEsa0TpmYGRfEu2I2tMpxIMKAJdr9WtTRd14Y` | Spreadsheet principal | Hardcodeado |
| `DRIVE_CARPETA_FEM_ID` | `1IqcFgQUSKocvGX3JwvNOu-xJzt0gfKc8` | Carpeta raíz pública (subcarpetas por IE y por grupo) | Hardcodeado |
| `DRIVE_CARPETA_EDITABLES_FEM_ID` | `1047FoYnCOsVakUl-7pCtvwpeceWM2DpB` | Carpeta privada de Docs editables (no compartidos) | Hardcodeado |
| `TEMPLATE_INFORME_ID` | `1Gtsccdbnlcyjl6TcDDjTOA7pAW3JQbHM` | Plantilla original del informe (Word `.docx` subido a Drive) | **Ya no se usa** — se conserva solo por referencia histórica (ver comentario Código.js:82-86) |
| `LOGO_ENCABEZADO_ID` | `1mFOOUZ5aFAuwM-JMxNUaDnPPznDlQ2bj` | Logo del Foro (FEM), encabezado/pie de informes y correos | Hardcodeado |
| `LOGO_PIE_ID` | `1Cmx7c3ec2gQCjRc8kcNeUbZt5LiURyD5` | Logo de la SEM | Hardcodeado |
| `MARCO_ACCESO_ID` | `1qKHFEoq61uBOn1tNusZxXcK8rutIxDAS` | Marco decorativo de la pantalla de acceso, referenciado como URL fija directo en `CSS.html` | Hardcodeado; debe estar público (`hacerPublicosLogosGlobales` en Pruebas.js) |
| `FONDO_SINTESIS_BANDA_SUPERIOR_ID` / `FONDO_SINTESIS_BANDA_INFERIOR_ID` | `113N96SZhqOkZclic-Kf-KKsVhqCGTbfN` / `15x2ulbI8iZLUluUxKWkAdcJHjrrP8I7-` | Bandas decorativas de portada de los informes de síntesis grupal/municipal | Hardcodeado |
| `DISENADOR_LOGO_ID` | `1BXkKDuSH_XhlLbdPtyYlXpJypbFH9f38` | Insignia del diseñador de la app, fija en todas las pantallas | Hardcodeado |
| Logo por IE (`LOGO_ID`) | *(variable)* | Columna `LOGO_ID` en `AccesosIE`, resuelta por `obtenerLogoIdPorNombreIE_` | **Dinámico** (vía Sheets, no hardcodeado) — patrón correcto a imitar para logos de IE/grupo en el proyecto nuevo |
| Spreadsheet "Análisis FEM 2026" | *(dinámico)* | `PropertiesService` (`SPREADSHEET_ANALISIS_ID`), creado con `SpreadsheetApp.create` si no existe | Dinámico |
| Spreadsheet "Participación... por Grupo" | *(dinámico)* | `PropertiesService` (`CLAVE_PROP_SPREADSHEET_GRUPOS_FEM_`), creado con `SpreadsheetApp.create` si no existe | Dinámico |
| Carpeta por IE | *(dinámica)* | `crearCarpetaIE_` — `getFoldersByName`/`createFolder` dentro de `DRIVE_CARPETA_FEM_ID` | Dinámico |
| Carpeta por grupo (`"Grupo GN"`, `"Informes enviados de grupo GN"`, `"Informes editables de grupo GN"`) | *(dinámica)* | `crearEstructuraCarpetasGrupoFEM_` — mismo patrón "obtener o crear" | Dinámico, **patrón directamente reutilizable para GRUPO como entidad raíz** |

**Patrón general observado**: assets fijos institucionales (logos, bandas, marco) están hardcodeados
como constantes; assets variables (por IE, por grupo, spreadsheets satélite) se resuelven dinámicamente
vía Sheets o `PropertiesService`. El proyecto nuevo debería llevar este segundo patrón un paso más allá y
mover también los IDs "fijos" a una hoja `ConfiguracionComunal` (clave/valor) en vez de constantes en
`Config.gs`, para permitir cambiarlos sin reeditar código.

---

## 5. Patrones de seguridad reutilizables (detalle técnico)

### 5.1 Generación de código de acceso

Dos generadores incompatibles coexisten (ver riesgo en §7):

```js
// Código.js:5157 — generarCodigoAcceso_ (NO se usa en generarAccesosIE)
function generarCodigoAcceso_(){
  const palabras=["LUNA","NUBE","RUTA","VOZ","AULA","FARO","NORTE","RIO","SOL","PUENTE","VIVA","VALLE"];
  const chars="ABCDEFGHJKMNPQRSTUVWXYZ23456789";
  let n=""; for(let i=0;i<4;i++) n+=chars.charAt(Math.floor(Math.random()*chars.length));
  return "FEM-"+palabras[Math.floor(Math.random()*palabras.length)]+n;
}
```

```js
// Código.js:4600-4648 — generador interno realmente usado por generarAccesosIE()
const caracteres = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";
let codigo = "";
do{
  codigo = "FEM-" + caracteres[...]+caracteres[...]+caracteres[...]+caracteres[...]+caracteres[...];
}while(codigosUsados[codigo]);
```

Ambos evitan caracteres ambiguos (`I,O,0,1` excluidos del alfabeto), y ambos verifican unicidad en
memoria (`codigosUsados`) durante la corrida — no contra toda la hoja histórica, así que solo garantizan
unicidad **dentro de la misma ejecución**.

### 5.2 Validación de token + código

```js
// Código.js:3164-3175
if (codigosValidos.indexOf(codigo) === -1) {
  return {
    ok: false,
    codigo: "CODIGO_INCORRECTO",
    mensaje: "El código es incorrecto. Verifique el código que se envió a su I.E. ..."
  };
}
```

`codigosValidos` (Código.js:3104-3128) es un arreglo de hasta 5 valores (`CODIGO_ACCESO` +
`CODIGO_CONTINGENCIA_1..4`) leídos de la misma fila que el `token` — es decir, el código nunca se busca
"a ciegas" en toda la hoja: primero se localiza la fila por `TOKEN` y luego se compara el código **solo**
contra esa fila. Esto hace que un código filtrado sin el token asociado no sirva para nada — el par
`(token, código)` es indivisible.

### 5.3 Control de un solo dispositivo (en realidad, hasta 4) activo

```js
// Código.js:2723-2739
if(sesiones.length >= MAX_SESIONES_SIMULTANEAS_IE){
  if(!forzar){
    return {
      ok:false, codigo:"SESION_YA_ABIERTA",
      mensaje:"Ya hay "+MAX_SESIONES_SIMULTANEAS_IE+" dispositivos conectados con este código de acceso, ..."
    };
  }
  sesiones.sort((a,b)=>(a.ultimaActividad||0)-(b.ultimaActividad||0));
  const indiceExpulsar = sesiones.findIndex(s=>!s.esPrincipal);
  sesiones.splice(indiceExpulsar===-1?0:indiceExpulsar, 1);
}
```

El "responsable principal" nunca es expulsado por un takeover; solo colaboradores. El estado vive en
`PropertiesService.getScriptProperties()` (no en Sheets), por lo que no genera contención de escritura en
la hoja ni queda auditado como fila — es efímero y se pierde si se limpian las propiedades del script.

### 5.4 Expiración de sesión

**No existe expiración por tiempo.** Cita textual del propio código (Código.js:2652-2659):

> "Ya NO hay temporizador de inactividad: una sesión reclamada por un dispositivo permanece activa
> indefinidamente (sin importar cuánto tiempo pase sin actividad) hasta que: a) ese mismo dispositivo la
> libera (liberarSesionCodigo_), o b) otro dispositivo la toma explícitamente con "forzar:true", tras la
> confirmación del usuario en pantalla."

Esto es una decisión de producto deliberada (documentada), no un descuido — pero implica que
`ScriptProperties` puede acumular sesiones "fantasma" de dispositivos que nunca liberaron su cupo (cerraron
la pestaña sin `liberarSesionCodigo_`) hasta que alguien haga un `forzar:true`.

### 5.5 `LockService` en escrituras concurrentes

Patrón uniforme en todas las funciones de escritura crítica:

```js
// Ejemplo — Código.js:5501 (registrarAsistenciaQR)
const lock=LockService.getScriptLock();
try{
  lock.waitLock(10000);
  ...
}catch(error){
  return {ok:false, mensaje:error.message};
}finally{
  try{ lock.releaseLock(); }catch(e){}
}
```

Timeouts observados: 10 s (`registrarAsistenciaQR`, `reclamarSesionCodigo_`, `mantenerSesionCodigo_`,
`liberarSesionCodigo_`), 15 s (`guardarAvanceForo`), 30 s (`generarAccesosIE`, `generarInformeFEM`,
`finalizarFormularioFEM`, `enviarForoDefinitivo`). El `finally` siempre libera el lock envuelto en su
propio try/catch para que un error al liberar no oculte el error real de negocio.

### 5.6 Protección contra duplicados / patrón UPSERT

Dos variantes del mismo patrón:

1. **UPSERT por `ID_FORO`** (`AvancesForo`, `Participacion`, hojas por IE): `buscarFilaPorIdForo_` /
   comparación manual de columna `ID_FORO`; si existe, `setValues` sobre la fila; si no, `appendRow`.
2. **Deduplicación por documento de identidad** (`AsistenciaQR`, Código.js:5567-5586): antes de insertar
   una firma nueva, se recorre toda la hoja buscando `(ID_FORO, NUMERO_DOCUMENTO)` igual — si ya existe,
   se devuelve la firma existente en vez de crear una fila nueva. Cita:

```js
// Código.js:5579-5585
for(let i=0;i<filas.length;i++){
  const mismoForo=String(filas[i][m.ID_FORO-1]||"").trim()===idForo;
  if(!mismoForo) continue;
  if(String(filas[i][m.NUMERO_DOCUMENTO-1]||"").trim()===documento){
    return {ok:true, yaRegistrado:true, textoFirma:textoFirmaDesdeFila_(filas[i])};
  }
}
```

Ambos patrones son lineales (`O(n)` filas leídas por operación, sin índice) — aceptable a la escala de
37 IE / cientos de asistentes, pero a revisar si el Foro Comunal multiplica el volumen de filas por
decenas de grupos (ver §7).

### 5.7 Fusión de campos concurrente (anti "última escritura gana")

```js
// Código.js:2311-2322
const camposFusionados = Object.assign({}, camposExistentes);
Object.keys(camposEntrantes).forEach(function(id) {
  const entrante = camposEntrantes[id];
  const tieneContenido = entrante && (
    entrante.tipo === "checkbox" ? true : String(entrante.valor ?? "").trim() !== ""
  );
  if (tieneContenido || !(id in camposFusionados)) {
    camposFusionados[id] = entrante;
  }
});
```

Este es, junto con el modelo de sesiones múltiples, el patrón más sofisticado de todo el proyecto 3.1 y
el más valioso de preservar tal cual: resuelve correctamente el caso de colaboración simultánea sin
sistema de CRDT ni websockets, apoyándose solo en Sheets + Lock + fusión de JSON.

---

## 6. Funciones NO necesarias para el nuevo proyecto

| Función/bloque | Archivo | Por qué no aplica |
|---|---|---|
| `GRUPOS_INSTITUCIONES` (App.html:211) y `GRUPOS_INSTITUCIONES_FEM_` (Código.js:8810) | App.html, Código.js | Catálogo fijo de 6 grupos × 6 IE de Neiva 2026 — el Foro Comunal define sus propios grupos/municipios, con cardinalidad distinta y probablemente configurable, no hardcodeada por duplicado en dos archivos |
| `programarEnvioAccesos28Agosto` (Código.js:3786) | Código.js | Trigger con fecha del 28 de agosto de 2026 hardcodeada, para un envío masivo ya ocurrido |
| `CORRECCION_EMAIL_POR_DANE_` (Código.js:136-174) y su uso en `repararEmailIEOficialesFEM`/`enviarAccesosSoloOficialesFEM` | Código.js, Pruebas.js | Tabla de 37 correos institucionales corregidos a mano por código DANE, específica de las IE de Neiva — no existe un equivalente genérico reutilizable |
| `DATOS_SINTESIS_GRUPO_1_FEM_` … `DATOS_SINTESIS_GRUPO_6_FEM_`, `DATOS_SINTESIS_MUNICIPAL_FEM_` y las funciones `generarInformeSintesisGrupo1FEM`…`Grupo6FEM`, `generarInformeSintesisMunicipalFEM` (Pruebas.js) | Pruebas.js | Miles de líneas de **prosa editorial redactada a mano** sobre el contenido real del FEM 2026 (qué dijo cada IE, qué grupo tuvo qué particularidad) — es contenido de un evento específico, no lógica de software |
| Todo el cuestionario de Sesión 1/2/3/Propia (preguntas orientadoras específicas del FEM2025/preescolar/currículo/gobierno escolar) en App.html y `obtenerRespuestasSesionesParaCompilado_`/`obtenerCabecerasAvancesForo` (Código.js) | App.html, Código.js | Preguntas específicas del ciclo FEM 2025→2026 de Neiva; el Foro Comunal necesitará su propio cuestionario, aunque puede reutilizar el *motor* de UI (contador de palabras, acordeones, "Otro") |
| `TEMPLATE_INFORME_ID` y cualquier código que lo referencie como "aún vigente" | Código.js | Ya deprecado en 3.1 mismo — el informe se genera 100% con `DocumentApp.create`, no copiando plantilla |
| Funciones de diagnóstico/reparación puntual de incidentes 2026: `restaurarAvancesForoDesdeAnalisisFEM`, `extraerSesionesDesdeDocEditableFEM_`, `resolverSesionesIEFEM_` (en su forma actual de 3 fuentes en cascada), `backfillFechaEnvioCorreoInformeDesdeGmailFEM`, `investigarEnvioInformeIE`, `diagnosticarCorreosInformeFEMPendientes` | Código.js, Pruebas.js | Parches reactivos a un incidente puntual de pérdida de filas en `AvancesForo` (2026-09-05) — el proyecto nuevo debería prevenir la causa (ver §7) en vez de heredar el parche |
| `crearIEsPruebaAdicionales`, `enviarTresCorreosIEsPruebaAdicionales`, `diagnosticarYLiberarIEsPruebaAdicionales`, `bloquearPruebaAnaYNelsonHasta9am`, y en general las ~40 funciones de `Pruebas.js` que crean/manipulan datos de prueba con nombres propios ("Ana y Nelson", "IE PRUEBA 1234") | Pruebas.js | Datos de prueba ad hoc del ciclo de pruebas de este proyecto concreto |
| `vincularLogosIE`, `hacerPublicosLogosGlobales` (una sola vez, ya ejecutadas) | Pruebas.js | Scripts de configuración inicial ya ejecutados sobre el Drive de este proyecto — no hay estado que migrar automáticamente |
| Validaciones/constantes específicas de asistencia atadas a roles del FEM (`ROLES_FORO_QR` con "Sistematización", "Relator(a)", etc.) | Código.js | Roles específicos de la metodología FEM — el Foro Comunal probablemente define sus propios roles de mesa de trabajo |

---

## 7. Riesgos y advertencias técnicas

| Riesgo | Evidencia | Recomendación para el proyecto nuevo |
|---|---|---|
| **Catálogo de grupos duplicado en 2 archivos** (`GRUPOS_INSTITUCIONES` en App.html:211 y `GRUPOS_INSTITUCIONES_FEM_` en Código.js:8810), con comentario explícito "MANTENER SINCRONIZADO con App.html" (Código.js:8807) | Código.js:8807-8817 | Fuente única de verdad: la hoja `GruposComunal`, consultada por RPC desde el frontend — nunca un array hardcodeado en el cliente |
| **Dos generadores de código de acceso incompatibles** (`generarCodigoAcceso_` en Código.js:5157, sin uso real, vs. el generador embebido en `generarAccesosIE`, Código.js:4600) | Código.js:5157-5162 y 4600-4648 | Un único generador en `Access.gs`, con test unitario que garantice que es el que realmente se invoca |
| **Bug crítico ya corregido pero ilustrativo de fragilidad del modelo de sesión**: `sesionActivaPorIdForo_` interpretaba el valor guardado como un objeto único cuando en realidad, desde que se soportan 4 dispositivos, siempre es un array — bloqueaba a todos los usuarios legítimos | Código.js:8427-8441 (comentario "BUG CRÍTICO CORREGIDO") | Al reescribir el modelo de sesión para GRUPO, cubrir con pruebas automatizadas (no solo manuales en `Pruebas.js`) los cambios de forma de dato en `PropertiesService` |
| **Incidente real de pérdida de datos** (2026-09-05): la hoja `AvancesForo` perdió filas de la mayoría de IE ya enviadas, y motivó 3 funciones de respaldo en cascada (`extraerSesionesDesdeDocEditableFEM_` → datos vivos → `obtenerRespuestasSesionesDesdeAnalisisFEM_`) | Código.js:8891-8905, 8979-8997; Pruebas.js:5199-5216 | Diseñar copias de seguridad periódicas explícitas (export a Drive, o versión de Sheets) desde el día 1, en vez de depender de una reconstrucción heurística post-incidente |
| **Un tab de Sheets por institución** (`nombreHojaIE_`, `guardarEnHojaIE_`) — con 37 IE ya son 37+ pestañas en el spreadsheet principal, más otras tantas en el spreadsheet de análisis | Código.js:5165-5213, 8692-8696 | Si el Foro Comunal cubre más municipios/IE que Neiva sola, este patrón escala mal (límite práctico de pestañas, lentitud de UI de Sheets); usar tablas únicas con columna `ID_GRUPO`/`ID_IE` y vistas filtradas en vez de una pestaña por entidad |
| **Búsquedas lineales sin índice** en casi toda función de lectura (`obtenerAccesoPorIdForo_`, `filaAccesoPorToken_`, deduplicación de `AsistenciaQR`, etc.) — recorren todas las filas de la hoja en cada llamada | Código.js (patrón repetido) | Aceptable a la escala actual (cientos de filas); si el volumen crece un orden de magnitud, considerar `CacheService` o mapas en memoria reconstruidos una vez por ejecución |
| **URL de despliegue hardcodeada** (`URL_WEBAPP_PRODUCCION`), con advertencia explícita de que `ScriptApp.getService().getUrl()` no sirve en el editor | Código.js:34-50 | Documentar el mismo procedimiento operativo (actualizar constante + volver a ejecutar el generador de accesos tras cada nuevo deployment) — o mover a `PropertiesService`/`ConfiguracionComunal` para no requerir republicar código |
| **`executeAs: USER_DEPLOYING` + `access: ANYONE_ANONYMOUS`** — todo el modelo de seguridad depende 100% de la lógica de aplicación (token+código); cualquier bug en `validarAccesoIE`/`reclamarSesionCodigo_` es una brecha directa, sin capa de identidad de Google como respaldo | appsscript.json | Mantener el mismo modelo si el requisito es "sin login", pero blindar con pruebas automatizadas la superficie de `Access.gs`/`Session.gs`; considerar rate-limiting básico (vía `CacheService`) contra fuerza bruta de código de 5 caracteres |
| **`dispositivoId` generado y confiado del lado cliente**, sin verificación server-side de unicidad real de dispositivo | App.html:9478-9486 | Aceptar como "identificador de conveniencia" (igual que 3.1), no como control de seguridad fuerte — documentarlo explícitamente para que el equipo del proyecto nuevo no le atribuya más garantía de la que tiene |
| **Código de generación de informes muy largo y monolítico** (`generarInformeFEM`, 557 líneas; `generarInformeSintesisGrupoFEM_` y `generarInformeSintesisMunicipalFEM_`, ~200-330 líneas cada una, con helpers de estilo redefinidos función a función en vez de compartidos) | Código.js:7124-7681, 9496-9695, 9720-9973 | Factorizar los helpers de estilo (`titulo1_`, `subtitulo_`, `parrafo_`, `tabla_`) en `Informes.gs` una sola vez, parametrizados, en vez de redefinirlos dentro de cada función generadora |
| **Contenido editorial (prosa de síntesis) embebido como código fuente** (`DATOS_SINTESIS_GRUPO_N_FEM_`) | Pruebas.js:6233-7615 | Si el Foro Comunal quiere síntesis por grupo, separar claramente "motor de render" (código) de "contenido" (dato editable, idealmente en un Doc o Sheet, no en un `.js`) desde el diseño inicial |
| **Falta de manifest de servicios avanzados explícitos más allá de Docs v1** | appsscript.json | Confirmar si el nuevo proyecto necesita además Drive Advanced Service o Sheets Advanced Service según las operaciones que se decidan (p. ej. batch updates); 3.1 se las arregla solo con los servicios base + `Docs v1` |
| **Falta de tipado / validación de esquema del JSON en la columna `DATOS`** — es un blob libre `{"campos":{"id":{"valor":...,"tipo":...}}}` sin ninguna validación de forma antes de guardarlo | Código.js (todo el flujo de `guardarAvanceForo`) | Aceptable para prototipar rápido; para el proyecto nuevo, documentar el esquema esperado de `DATOS` y validar campos mínimos server-side antes de aceptar el guardado, especialmente si varios GRUPOs comparten el mismo backend |

---

## 8. Conteo de líneas por archivo

```
   9973 Código.js
   7681 Pruebas.js
   9487 App.html
   2743 Index.html
   4544 CSS.html
     17 appsscript.json
  34445 total
```

Funciones de nivel superior contadas (`^function nombre(`):

```
135 funciones en Código.js
191 funciones en App.html
 91 funciones en Pruebas.js
```

Todo el contenido de los 6 archivos fue leído en su totalidad (lectura secuencial en bloques + `Grep`
dirigido por patrones de `getSheetByName`, `DriveApp`, `LockService`, `PropertiesService`, `GmailApp`,
constantes `_ID`/`HOJA_`, y las funciones citadas explícitamente en el encargo) para construir esta
auditoría.

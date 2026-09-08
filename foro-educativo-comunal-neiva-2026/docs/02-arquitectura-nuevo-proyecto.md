# Arquitectura propuesta — Foro Educativo Comunal Neiva 2026

> Fases 2-9 del plan de desarrollo (ver `Especificacion_Foro_Educativo_Comunal_Neiva_2026.docx`, sección
> 28). Basado en la auditoría de FEI 3.1 (`01-auditoria-fei-3.1.md`). Este documento se presenta para
> **validación antes de generar el código completo** (Fases 10-16), tal como exige la especificación.
>
> Proyectos involucrados:
> - **Foro Educativo 3.1** (`1CD6ccfWvr0gnnxQK6qHGLsN07RPeUpgEDvk4kZsjX3WlJZctlIH076mR`) — solo lectura,
>   referencia técnica. No se modifica.
> - **Foro comunal 1.0** (`1tqsSNT-3BiCkQhcfSCVmzMe9IeVIsDuFujf_rzMRygz3pDUbRP7fKy7Y`) — copia de 3.1 ya
>   creada por el usuario; es el proyecto de Apps Script donde se desplegará el código nuevo (independiente,
>   sin dependencias de ejecución con 3.1).

---

## Fase 2 — Arquitectura existente (resumen)

Ver `01-auditoria-fei-3.1.md` §1 completo. En una frase: monolito Apps Script V8, `doGet` único +
`include()` para componer `Index+CSS+App` en un solo HTML, seguridad 100% de aplicación (token+código,
sin login Google), Sheets como base de datos con UPSERT por `ID_FORO` y `LockService`, Drive organizado
con carpetas fijas + dinámicas, informes generados con `DocumentApp.create()` y exportados a PDF, correo
vía `GmailApp` con cola de reintento por cuota. Ya existe una capa de "grupo" (bottom-up, catálogo
hardcodeado) cuyo motor de carpetas/documentos por grupo es el activo más reutilizable para este proyecto.

## Fase 3 — Matriz de reutilización

Ver `01-auditoria-fei-3.1.md` §2 (matriz completa función por función) y §6 (qué no se reutiliza). Resumen
por bloque:

| Bloque | Decisión | Archivo nuevo |
|---|---|---|
| Fusión de campos concurrentes + guardado local/sync | Reutilizar tal cual (patrón más maduro de 3.1) | `Data.gs` / `JS.html` |
| Sesión múltiple por dispositivo + `LockService` | Reutilizar la arquitectura, adaptar clave (grupo, no IE) y el límite de cupos | `Session.gs` |
| Token + código + bloqueo por horario | Reutilizar fuertemente, adaptar a nivel GRUPO | `Access.gs` |
| Estructura de carpetas por grupo (`crearEstructuraCarpetasGrupoFEM_`) | Reutilizar fuertemente | `Drive.gs` / `Grupos.gs` |
| Motor de generación de Doc/PDF + gráficos (`Charts`, `Docs v1`) | Reutilizar el motor, contenido nuevo | `Informes.gs` |
| Envío de correo + cola de cuota agotada | Reutilizar fuertemente | `Correo.gs` |
| Catálogo de grupos hardcodeado (duplicado en 2 archivos) | **No reutilizar** — pasa a ser la hoja `GruposComunal`, fuente única | `Grupos.gs` (lee Sheet) |
| Hoja de Sheets por institución | **No reutilizar** — no escala | Tabla única con `ID_IE`/`ID_GRUPO` |
| Cuestionario Sesión 1/2/3 del FEI y prosa editorial de síntesis | **No reutilizar el contenido** | Cuestionario propio (Fase 5) |
| Parches de incidentes puntuales (pérdida de filas, doble generador de código) | **No reutilizar** | Diseñar sin la causa raíz (backups explícitos, un solo generador) |

---

## Fase 4 — Arquitectura del nuevo proyecto

### 4.1 Principio rector

La unidad raíz es **GRUPO**, no IE. Todo el modelo de datos, acceso y generación de informes cuelga de
`ID_GRUPO`. Las IE son atributos/miembros de un grupo (tabla `GruposComunal`), no la entidad de acceso.

### 4.2 Relación de identificadores

```
ID_FORO_COMUNAL  (uno solo, estable, para todo el evento — p.ej. "FEC-NEIVA-2026")
      └── ID_GRUPO   (uno por grupo: GRUPO 1, GRUPO 2, ...)
              ├── ID_IE (N por grupo, vía GruposComunal)
              ├── ID_PARTICIPANTE (N, vía ParticipacionComunal)
              ├── Sesion1Comunal (1 fila por ID_GRUPO — síntesis colectiva)
              ├── ConectaEduca (N filas por ID_GRUPO — un actor por fila)
              └── InformesComunal (1 fila por ID_GRUPO — DOC_ID + PDF_ID únicos)
```

`ID_FORO_COMUNAL` se define una sola vez en `ConfiguracionComunal` (no se genera en cada guardado, igual
que 3.1 nunca regenera `ID_FORO` — lección ya aprendida y documentada en la auditoría §1.1). `ID_GRUPO` se
genera una vez al poblar `GruposComunal` (p. ej. `GRUPO-01`) y es estable durante todo el evento.

### 4.3 Componentes del sistema (mapeo a archivos CLASP)

| Componente | Archivo | Responsabilidad | Hereda de 3.1 |
|---|---|---|---|
| Router HTTP | `Code.gs` | `doGet(e)`, `include()`, arma variables de plantilla (token, grupo si aplica) | `doGet`, `include` (adaptados) |
| Configuración | `Config.gs` | Lee `ConfiguracionComunal` (clave/valor) una vez por ejecución, cachea en memoria; expone `getConfig()` | Patrón nuevo (3.1 usaba constantes fijas) |
| Acceso | `Access.gs` | `validarAccesoGrupo`, `generarAccesosGrupo`, gestión de `AccesosGrupo` | `validarAccesoIE`, `generarAccesosIE` (adaptados a grupo) |
| Sesión/dispositivo | `Session.gs` | `reclamarSesionCodigo_`, `mantenerSesionCodigo`, `liberarSesionCodigo_`, `transferirResponsablePrincipal`, `sesionActivaPorIdGrupo_` | Reutilización directa de la arquitectura de `PropertiesService` + `LockService` |
| Datos genéricos | `Data.gs` | Apertura de Sheet con reintento, UPSERT genérico por ID, lectura tolerante por cabecera, fusión de campos concurrentes | `abrirSpreadsheet_`, `buscarFilaPorIdForo_`, patrón de fusión (§5.7 auditoría) |
| Instituciones | `Instituciones.gs` | `obtenerInstitucionesDelGrupo(idGrupo)` desde `GruposComunal` | `obtenerInstitucionesJSON` (adaptado, ya no es catálogo raíz) |
| Grupos | `Grupos.gs` | `obtenerGrupos()`, `obtenerGrupoPorId`, orquestación de generación de informe por grupo | `mapaGruposFEM_`, `crearEstructuraCarpetasGrupoFEM_`, `generarDocumentoCompiladoGrupoFEM_` (adaptados) |
| Sesión 1 | `Sesion1.gs` | `guardarSesion1(idGrupo, datos)` (UPSERT), `obtenerSesion1(idGrupo)` — convergencias, apuestas, desafíos, identidad, prioridades, propuestas, acuerdos, ruta | Patrón UPSERT de `guardarAvanceForo`, cuestionario 100% nuevo |
| ConectaEduca | `ConectaEduca.gs` | `guardarActorConectaEduca(idGrupo, registro)` (append), `listarConectaEduca(idGrupo)` | Sin equivalente directo — módulo nuevo, reutiliza solo utilidades genéricas |
| Informes | `Informes.gs` | `generarInformeGrupo(idGrupo)`: arma Doc con `DocumentApp.create`, gráficos (`Charts`), exporta PDF, UPSERT en `InformesComunal` | `generarInformeFEM`, `generarInformeSintesisGrupoFEM_` (motor reutilizado, contenido nuevo) |
| Drive | `Drive.gs` | Crea/asegura las 6 carpetas raíz + subcarpeta por grupo, helpers de "obtener o crear", `hacerPublicoSiEsPosible_` | `crearEstructuraCarpetasGrupoFEM_`, `crearOFolderHija_`, helpers de Drive (reutilización directa) |
| Correo | `Correo.gs` | Envío de accesos por grupo, envío de informe final, cola `EnviosDiferidosComunal` | `enviarInformeFEM`, cola de cuota agotada (reutilización directa) |
| Utilidades | `Utils.gs` | Formato de fecha, normalización de texto, `tallyOpciones_`, conteo de palabras | Helpers triviales de 3.1, agrupados tal cual |
| Pruebas | `Tests.gs` | Pruebas manuales ejecutables desde el editor (creación de accesos de prueba, validación end-to-end) | Metodología de `Pruebas.js`, contenido nuevo |
| Estructura | `Index.html` | 12 pantallas (ver Fase 8) + panel de firmantes persistente | Patrón de `<section id="pantalla...">` + plantilla `<? ?>` |
| Estilos | `CSS.html` | Paleta institucional, responsive | Base reutilizada de `CSS.html` de 3.1, marca a confirmar |
| Lógica cliente | `JS.html` | `dispositivoId`, guardado local + sync, heartbeat, validación de acceso, actualización del panel de firmantes | `App.html` — módulos de borrador local y latido de sesión (reutilización directa) |
| Componentes UI | `Components.html` | Acordeones, contador de palabras, tarjeta de IE, tarjeta de firmante | Patrones de UI de `App.html`, sin el cuestionario específico del FEI |
| Modales | `Modal.html` | "Ver todos los firmantes", confirmación de envío definitivo, aviso de toma de sesión (`forzar`) | Patrón de modales de cierre de 3.1 |
| Vista previa de informe | `InformeStyles.html` | Estilos para una vista previa en pantalla del informe antes de generarlo (opcional, a confirmar en Fase 8) | Sin equivalente directo en 3.1 |

### 4.4 Seguridad (heredada de 3.1, adaptada a GRUPO)

- Manifest igual que 3.1: `"executeAs": "USER_DEPLOYING"`, `"access": "ANYONE_ANONYMOUS"` (sin login de
  Google) — toda la seguridad es de aplicación.
- Acceso por **token + código a nivel de GRUPO** (no de IE): un grupo tiene un único `TOKEN` y hasta 4
  códigos (principal + 3 de contingencia), igual que 3.1 por IE.
- Sesión múltiple por dispositivo: `MAX_SESIONES_SIMULTANEAS_GRUPO = 4`, igual que 3.1 (confirmado con el
  usuario), pero definido como constante en `Config.gs`/`ConfiguracionComunal` para poder ajustarlo sin
  reeditar código si la operación real lo exige.
- `LockService` en toda escritura compartida (participación, Sesión 1, ConectaEduca, generación de
  informe) — igual patrón que 3.1 (§5.5 auditoría).
- Cliente nunca puede fijar por sí mismo grupo, IE, `ID_FORO_COMUNAL`, correo institucional ni estado
  definitivo — toda esa información se resuelve/valida en servidor a partir del `TOKEN` (igual regla que
  3.1, spec sección 22).
- **Un solo generador de código de acceso** (a diferencia de 3.1, que tenía dos incompatibles — ver
  auditoría §7) — evitar caracteres ambiguos, igual que 3.1.

---

## Fase 5 — Modelo de datos

### 5.1 Entidades

| Entidad | Clave | Pertenece a | Cardinalidad |
|---|---|---|---|
| ForoComunal (config) | `ID_FORO_COMUNAL` | — | 1 por evento (2026) |
| Grupo | `ID_GRUPO` | ForoComunal | N (uno por grupo territorial) |
| Institución (IE) | `ID_IE` | Grupo | N por grupo (miembro, no raíz) |
| Participante/Asistencia | `ID_PARTICIPANTE` | Grupo + IE | N por grupo |
| Sesión 1 (síntesis colectiva) | `ID_GRUPO` | Grupo | 1 fila por grupo (UPSERT) |
| ConectaEduca (actor) | `ID_GRUPO` + fila | Grupo | N filas por grupo |
| Informe | `ID_GRUPO` | Grupo | 1 fila por grupo (un único DOC_ID/PDF_ID) |

### 5.2 Estados del proceso (por grupo, spec sección 15)

`PENDIENTE → EN_PROCESO → SESION1_GUARDADA → SESION1_ENVIADA → SESION2_GUARDADA → SESION2_ENVIADA →
INFORME_GENERADO → ENVIADO` (con rama a `ERROR` en cualquier punto). Se guarda en `AccesosGrupo.ESTADO`
igual que 3.1 guardaba `ESTADO` en `AccesosIE`.

### 5.3 Campos de "socialización" de Sesión 1 (confirmado con el usuario)

Además de los 8 campos de síntesis colectiva (sección 10 de la spec), `Sesion1Comunal` añade 6 columnas
más para los insumos de socialización que menciona la sección 9 (Reflexiones, Conclusiones, Propuestas,
Experiencias, Retos, Aportes territoriales) — **un solo valor agregado por grupo, no por IE**: el
facilitador de la mesa consolida en un campo lo que las IE compartieron verbalmente, no se piden 37
formularios individuales (se respeta "el aplicativo no debe volver a solicitar todas las preguntas del
FEI"). Ver estructura final de `Sesion1Comunal` en la Fase 6.

---

## Fase 6 — Google Sheets

Spreadsheet único del proyecto nuevo (independiente de 3.1). Nombres de hoja exactamente los del mínimo
pedido en la especificación, más 2 hojas de soporte necesarias para replicar el modelo de seguridad de 3.1
(no listadas en el mínimo, pero requeridas — señaladas como **[añadida]**).

### `GruposComunal`
`ID_GRUPO | GRUPO | ID_IE | INSTITUCION | COMUNA | ACTIVO`
Una fila por IE dentro de un grupo (relación Grupo→IE desnormalizada, tal como pide la spec sección 6:
"la información Grupo → IE debe provenir de Google Sheets"). `obtenerInstitucionesDelGrupo(idGrupo)`
filtra por `ID_GRUPO`.

### `AccesosGrupo` **[añadida]**
`ID_ACCESO | ID_GRUPO | GRUPO | ID_FORO_COMUNAL | TOKEN | CODIGO_ACCESO | CODIGO_CONTINGENCIA_1..3 |
URL_ACCESO | ESTADO | HABILITAR_DESDE | EMAIL_RESPONSABLE_GRUPO | FECHA_GENERACION | ULTIMA_ACTIVIDAD |
SESION1_ENVIADA | SESION2_ENVIADA | FECHA_ENVIO_S1 | FECHA_ENVIO_S2 | FECHA_ENVIO_DEFINITIVO`
Equivalente a `AccesosIE` de 3.1 pero a nivel de grupo (auditoría §2.1). Necesaria para que
`validarAccesoGrupo`/`generarAccesosGrupo` tengan dónde vivir — sin esta hoja no hay forma de reproducir
el modelo de token+código exigido por la spec (sección 2, "Acceso mediante TOKEN + código").

### `ParticipacionComunal`
`ID_PARTICIPANTE | ID_GRUPO | ID_IE | NOMBRE | ROL | CORREO | CONFIRMACION_ASISTENCIA | FECHA | ESTADO`
(columnas base de la spec, sección 13/tabla, + `CORREO` y `CONFIRMACION_ASISTENCIA`/`ESTADO` pedidos
explícitamente en sección 7). Deduplicación por `(ID_GRUPO, ID_IE, NOMBRE)` — patrón de `AsistenciaQR`
(auditoría §5.6).

### `Sesion1Comunal`
`ID_GRUPO | REFLEXIONES | CONCLUSIONES | PROPUESTAS_IE | EXPERIENCIAS | RETOS | APORTES_TERRITORIALES |
CONVERGENCIAS | APUESTAS | DESAFIOS | IDENTIDAD | PRIORIDADES | PROPUESTAS_COLECTIVAS | ACUERDOS | RUTA`
Las 6 primeras columnas (tras `ID_GRUPO`) son los insumos de socialización de la sección 9 de la spec,
agregados por el grupo (no por IE — confirmado con el usuario, ver 5.3); las 8 restantes son la síntesis
colectiva de la sección 10/tabla mínima de la spec (`PROPUESTAS_COLECTIVAS` se renombra así, sin guion
bajo con `PROPUESTAS_IE`, para no confundir ambos campos de "propuestas"). UPSERT por `ID_GRUPO` — un
único registro editable colectivamente por todos los dispositivos conectados al grupo (mismo patrón de
fusión de campos que `AvancesForo`, auditoría §5.7).

### `ConectaEduca`
`ID_GRUPO | ACTOR | TIPO_ACTOR | AREA | NECESIDADES_ARTICULACION | OPORTUNIDAD | ALIANZA | IE_INTERESADAS
| CONEXIONES | OBSERVACIONES`
(columnas base de la spec sección 13 + `NECESIDADES_ARTICULACION` y `CONEXIONES`, pedidos explícitamente
en la lista de la sección 11 pero ausentes en la tabla mínima — añadidas para cubrir el requisito
completo). Una fila por actor/entidad registrada; no es UPSERT único como Sesión 1, es *append* con
posibilidad de editar/borrar la fila propia.

### `InformesComunal`
`ID_GRUPO | DOC_ID | PDF_ID | URL | FECHA | ESTADO`
(exactamente la tabla mínima de la spec). Regla fundamental: **una sola fila por grupo**, nunca una por
IE (spec sección 18).

### `ConfiguracionComunal`
Formato clave/valor (`CLAVE | VALOR | DESCRIPCION`), con al menos las claves de la spec sección 20:
`ID_FORO_COMUNAL, NOMBRE_FORO, SUBTITULO, FECHA, CARPETA_DRIVE_ID, SPREADSHEET_ID, PLANTILLA_INFORME,
LOGO_ENCABEZADO_ID, LOGO_PIE_ID, CORREO_REMITENTE, COPIAS_CORREO, TIEMPO_SESION, TIEMPO_MAXIMO_SOCIALIZACION,
MAX_SESIONES_SIMULTANEAS_GRUPO, ACTIVO`. Reemplaza el patrón de constantes hardcodeadas de `Config.js`
de 3.1 (recomendación explícita de la auditoría §4) — así los IDs de Drive/logos se pueden cambiar sin
tocar código.

### `EnviosDiferidosComunal` **[añadida]**
`ID_GRUPO | FECHA_REGISTRO | REINTENTADO` — cola de reintento de correo por cuota agotada, mismo patrón
que `EnviosInformeDiferidos` de 3.1 (auditoría §1.5/§5).

---

## Fase 7 — Google Drive

Estructura exacta pedida por la spec (sección 16), con el patrón "obtener o crear" de 3.1
(`crearOFolderHija_`) para cada nivel:

```
FORO EDUCATIVO COMUNAL NEIVA 2026/   (raíz, ID guardado en ConfiguracionComunal.CARPETA_DRIVE_ID)
├── 01_DATOS
├── 02_ASISTENCIA
├── 03_INFORMES_GRUPALES/
│   ├── GRUPO 1/
│   │   ├── Informe Grupo 1.docx      (privado, equipo de calidad — mismo criterio que
│   │   │                               DRIVE_CARPETA_EDITABLES_FEM_ID de 3.1, ver auditoría §1.4)
│   │   └── Informe Grupo 1.pdf       (compartido "cualquiera con el enlace", vía
│   │                                   hacerPublicoSiEsPosible_)
│   ├── GRUPO 2/ ...
├── 04_CONECTAEDUCA
├── 05_CONSOLIDADO_MUNICIPAL
└── 06_EVIDENCIAS
```

**Confirmado con el usuario**: se mantiene el mismo criterio de seguridad de 3.1 — el `.docx` vive en la
misma carpeta `GRUPO N` (tal como pide la spec, sección 16) pero **no** se comparte públicamente; solo el
`.pdf` se comparte "cualquiera con el enlace" vía `hacerPublicoSiEsPosible_`. Permisos por archivo, no por
carpeta.

---

## Fase 8 — Flujo frontend

Pantallas (spec sección 5), controladas por clase `visible`/oculta igual que 3.1 (sin router de
navegador):

```
1. pantallaInicio                 — portada, "Encuentro de voces que construyen territorio"
2. pantallaPresentacion           — presentación del Foro
3. pantallaMetodologia            — metodología del encuentro
4. pantallaAcceso                 — token + código de GRUPO (Access.gs)
5. pantallaSeleccionGrupo         — solo si el token no resuelve un único grupo; si el token ya
                                     identifica el grupo, se salta (igual que 3.1 resolvía la IE por
                                     token sin pedirla en pantalla)
6. pantallaParticipacion          — asistencia + firmantes (con panel permanente, ver 8.1)
7. pantallaSesion1                — socialización (guiada) + construcción colectiva (8 campos)
8. pantallaSesion2ConectaEduca    — registro de actores/entidades
9. pantallaRevisionCierre         — resumen antes de enviar
10. pantallaEnvioDefinitivo       — confirmación (modal), envío final + generación de informe
11. pantallaInformeGenerado       — enlace al informe del grupo (mismo para todas las IE)
12. pantallaDespedida             — cierre
```

### 8.1 Panel permanente de firmantes

Componente `sticky` en la parte superior de **todas** las pantallas desde `pantallaParticipacion` en
adelante (spec sección 8): contador + últimos firmantes + botón `[Ver todos]` (abre `Modal.html` con lista
completa: nombre, institución, rol). Se actualiza:
- al registrar una asistencia (respuesta local inmediata, optimista), y
- por *polling* periódico a `contarAsistentesGrupo(idGrupo)`/`listarFirmantesGrupo(idGrupo)` (mismo
  patrón que el latido de sesión, `JS.html`), para reflejar firmas de otros dispositivos del mismo grupo.

### 8.2 Guardado

Autoguardado local (`localStorage`) + sincronización periódica al servidor, reutilizando el módulo de
borrador de 3.1 tal cual (auditoría §2.6) — el candidato más maduro para reutilización directa.

---

## Fase 9 — Backend (firmas de funciones propuestas)

```js
// Code.gs
function doGet(e)
function include(nombre)

// Config.gs
function getConfig()                          // lee ConfiguracionComunal, cachea en memoria de ejecución

// Access.gs
function validarAccesoGrupo(token, codigo, dispositivoId, forzar)
function generarAccesosGrupo()                 // una vez, genera TOKEN/CODIGO por grupo en AccesosGrupo
function generarCodigoAcceso_()                // ÚNICO generador (a diferencia de 3.1)

// Session.gs
function reclamarSesionCodigo_(token, codigo, dispositivoId, idGrupo, forzar)
function mantenerSesionCodigo(token, codigo, dispositivoId, tokenSesion, idGrupo)
function liberarSesionCodigo_(token, codigo, dispositivoId, tokenSesion, idGrupo)
function transferirResponsablePrincipal(token, codigo, dispositivoId, tokenSesion, idGrupo)
function sesionActivaPorIdGrupo_(idGrupo, dispositivoId, tokenSesion)

// Instituciones.gs
function obtenerInstitucionesDelGrupo(idGrupo)

// Grupos.gs
function obtenerGrupos()
function obtenerGrupoPorId(idGrupo)
function generarInformeCompletoGrupo(idGrupo)   // orquesta Informes.gs + Drive.gs + InformesComunal

// Sesion1.gs
function guardarSesion1(idGrupo, tokenSesion, dispositivoId, datos)   // UPSERT + fusión de campos
function obtenerSesion1(idGrupo)
function enviarSesion1Definitiva(idGrupo, tokenSesion, dispositivoId)

// ConectaEduca.gs
function guardarActorConectaEduca(idGrupo, tokenSesion, dispositivoId, registro)
function listarConectaEduca(idGrupo)
function eliminarActorConectaEduca(idGrupo, idFila, tokenSesion, dispositivoId)

// Informes.gs
function generarInformeGrupo(idGrupo)           // DocumentApp.create + gráficos + PDF
function obtenerInformeGrupo(idGrupo)            // URL/DOC_ID/PDF_ID ya generados (para las demás IE)

// Drive.gs
function asegurarEstructuraDriveComunal_()       // 6 carpetas raíz, una vez
function asegurarCarpetaGrupo_(idGrupo)
function hacerPublicoSiEsPosible_(file)

// Correo.gs
function enviarAccesosGrupo(idGrupo)
function enviarInformeGrupo(idGrupo, pdfId)
function registrarEnvioDiferido_(idGrupo)
function reintentarEnviosDiferidos()

// Utils.gs
function tallyOpciones_(personas, campo)
function normalizarTexto_(texto)
function formatearFechaLargaEs_(fecha)

// Tests.gs
function crearAccesoDePrueba()
function probarFlujoCompletoGrupo()
```

---

## Fase 10 en adelante — decisiones confirmadas y arranque de implementación

Los 4 puntos de decisión quedaron resueltos con el usuario:

1. **Sesión 1**: los 6 insumos de socialización se agregan como 6 columnas más en `Sesion1Comunal` (un
   valor por grupo, no por IE) — ver Fase 5.3 y Fase 6.
2. **Cupos de sesión**: `MAX_SESIONES_SIMULTANEAS_GRUPO = 4`, igual que 3.1, configurable desde
   `Config.gs`/`ConfiguracionComunal`.
3. **Informe editable**: mismo criterio de seguridad que 3.1 — `.docx` no público, en la misma carpeta
   `GRUPO N` que el `.pdf` público (permisos por archivo).
4. **Despliegue**: implementación en este repositorio Git primero; `clasp push` al proyecto Apps Script
   **Foro comunal 1.0** (`1tqsSNT-3BiCkQhcfSCVmzMe9IeVIsDuFujf_rzMRygz3pDUbRP7fKy7Y`) cuando cada entrega
   quede validada, usando las credenciales de `clasp login` ya autorizadas por el usuario en esta sesión.

Con esto arrancan las Fases 10-16 (estructura CLASP, implementación de backend/frontend, generación de
informes/PDF, pruebas integrales, auditoría final de dependencias).

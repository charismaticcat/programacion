# FASE 1 — Arquitectura y diseño (v5)
### Formulario público de captura + Super Admin / CRM (Google Apps Script · Sheets · Drive · clasp)

> Estado: **implementado**. Los ajustes hechos durante la implementación están en [FASE2_IMPLEMENTACION.md](FASE2_IMPLEMENTACION.md).
>
> **Objetivo central:** que el cliente ingrese su información, que el administrador la vea y la edite, y que al final se genere **un reporte editable por cliente**.
>
> Cambios de la v2: sin NIT, sin tipo de persona, formulario público sin identificación, propietario en blanco, emails repetidos descartados, actividad completada con el CIIU, teléfonos fijos con 608, solo 4 municipios y cuenta @gmail.com.
>
> Cambios de la v3: **todos los campos son editables**, **el email solo alimenta la base de datos** (el sistema no envía correos) y **reporte editable por cliente** en Google Docs.

---

## 0. Decisiones confirmadas y reglas de importación

### 0.1 Decisiones del cliente

| # | Tema | Decisión | Cómo se aplica |
|---|---|---|---|
| D1 | Nombre propietario | **Se deja en blanco** en los 30.780 registros importados | `NOMBRE_PROPIETARIO` queda vacío. Solo lo llena el cliente en el formulario o un admin. |
| D2 | Fuente de la importación | Importar desde `CCH@E26-9316.xlsx`. `Base alimentada` solo sirve para verificar. | Se enlazan por `ID_REGISTRO` = número de fila (coinciden las 30.780). |
| D3 / D8 | NIT | **No se usa** | La columna `NIT` **no se importa a ninguna hoja**, ni siquiera a la copia de origen, y no se usa para buscar, identificar ni detectar duplicados. |
| D4 | Persona natural o jurídica | **No se usa** | No existe el campo `TIPO_PERSONA`. |
| D5 | Emails repetidos | **No se usan** | Un email que aparece en más de un registro queda en blanco en `CLIENTES` (2.105 filas). También quedan en blanco los 4 mal formados (`@gmailcom`) y el 1 vacío. Resultado: **28.670 con email y 2.110 sin email**. |
| D6 | Actividad vacía | Completarla con el CIIU | Las 7.754 filas sin `ACTIVIDAD` toman la descripción de `CIIU-1` (el texto que va después de `**`). Todas la tienen. `ACTIVIDAD_FUENTE` indica si el dato viene de `BASE` o de `CIIU`. |
| D7 | Teléfonos fijos | Anteponer 608 a todos | Los 1.496 números de 7 dígitos pasan a `608` + número (ej. `8620841` → `6088620841`). Los unos 50 inválidos (valor `0`, 8, 9 u 11 dígitos) quedan en blanco en `TELEFONO` y conservan el valor original en `TELEFONO_ORIGINAL` con una observación. |
| D9 | Municipios | Solo Neiva, Pitalito, Palermo y Rivera | El catálogo `MUNICIPIOS` tiene solo esos 4 (41001, 41551, 41524, 41615). El formulario público los ofrece en una lista cerrada y el servidor rechaza cualquier otro. |
| D10 | Formulario público | **No requiere identificación**: el cliente solo llena los datos | Sin búsqueda en la base, sin código de verificación y sin mostrar datos existentes. La vinculación con la base la hace el admin (§5.6). |
| — | Cuenta | @gmail.com | Aplica los límites de cuota de la cuenta gratuita (§10). |
| D11 | Campos editables | **Todos los campos son editables** | En el formulario público todos los campos son de libre edición: ninguno es de solo lectura ni viene precargado. En el Super Admin, **cada campo de cada cliente se puede editar** (incluidos los importados de la base: razón social, teléfono, dirección, actividad, etc.) y cada cambio queda en la auditoría con el valor anterior. Las listas (municipio, tipo de establecimiento) también se pueden cambiar en cualquier momento. |
| D12 | Uso del email | **Solo alimenta la base de datos** | El email es un dato más del cliente. **El sistema no envía correos**: no hay confirmación, ni enlace para continuar, ni verificación. El borrador vive en el navegador del cliente. |
| D14 | Precios | **Espacio de precios editable** | Ningún precio es fijo. `PRODUCTOS_WEB.PRECIO` es solo una **sugerencia opcional** (puede quedar vacío). En oportunidades, propuestas y ventas cada valor (`VALOR_ESTIMADO`, precio por ítem, descuento, total) **se escribe o modifica a mano**. Los precios del cliente (servicios, productos, domicilios) también son campos libres. |
| D15 | Campos autocompletables vs. libres | **Todo lo que no tenga datos para autocompletar es editable** | Solo hay 2 campos con **sugerencias**: municipio (los 4) y tipo de establecimiento (catálogo). Ambos tienen autocompletado **y** siguen siendo editables: el tipo de establecimiento admite texto libre si no está en la lista (se guarda y el admin puede añadirlo al catálogo). **Todos los demás campos son texto libre editable.** |
| D16 | Software de diseño web | **Aún no definido** | El reporte usa la estructura genérica por secciones web + JSON opcional (§4.1). Cuando se elija el software, solo se ajusta la plantilla del reporte, sin cambiar datos ni código de captura. |
| D13 | Resultado final | **Un reporte editable por cliente** | Google Doc generado con todos los datos e imágenes del cliente, guardado en su carpeta de Drive y editable por el admin (§4.1). |
| — | Importación | Los 30.780 registros | Todos entran al CRM como `PROSPECTO`. |
| — | Catálogo de productos web y precios | Sí | Tabla `PRODUCTOS_WEB` configurable. **Faltan los valores** (§11). |
| — | Responsable del tratamiento de datos | **Equipo Webpaya** | Figura en la política, el aviso de privacidad y los emails. |

### 0.2 Mapeo final: `CCH@E26-9316.xlsx` → `CLIENTES`

| Origen (CCH) | Destino | Transformación |
|---|---|---|
| (número de fila) | `CLIENTE_ID` | `CLI-` + número de fila con 6 dígitos (`CLI-000001`) |
| `RAZON SOCIAL` | `RAZON_SOCIAL` | Recortar espacios, unir espacios dobles, quitar `\xa0` |
| — | `NOMBRE_PROPIETARIO` | **Vacío** (D1) |
| — | `NOMBRE_COMERCIAL` | **Vacío**. Nunca se infiere. |
| `TEL-COM-1` | `TELEFONO` / `TELEFONO_ORIGINAL` | Solo dígitos. 7 dígitos → `608`+número. 10 dígitos → igual. Otro caso → vacío con observación (D7). |
| `EMAIL-COMERCIAL` | `EMAIL` | En minúsculas. Vacío si se repite o está mal formado (D5). |
| `ACTIVIDAD` / `CIIU-1` | `ACTIVIDAD`, `ACTIVIDAD_FUENTE`, `CIIU_CODIGO`, `CIIU_DESCRIPCION` | Si `ACTIVIDAD` está vacía, se usa `CIIU_DESCRIPCION` (D6) |
| `MUN-COMERCIAL` | `MUNICIPIO_COD`, `MUNICIPIO` | Se separa `41001 - NEIVA`. Se valida contra los 4 municipios (D9). |
| `DIR-COMERCIAL` | `DIRECCION` | Texto original, solo recortado |
| `BARRIO-COMERCIAL` | `BARRIO_COD`, `BARRIO` | Se separa `01112 - IPANEMA`. 1.198 filas vacías quedan vacías. |
| `FEC-MATRICULA`, `FEC-RENOVACION`, `ULT-ANO_REN` | `FEC_MATRICULA`, `FEC_RENOVACION`, `ULT_ANO_RENOVACION` | Se convierten a fecha (`yyyymmdd`) |
| `TAM-EMPRESA` | `TAMANO_EMPRESA` | Sin cambio |
| `EST-MATRICULA` | — | Se descarta (siempre `MA`) |
| `NIT` | — | **Se descarta** (D3) |
| — | `OBSERVACION_IMPORTACION` | Lista de ajustes aplicados, ej. `EMAIL_REPETIDO_DESCARTADO; TEL_608; ACTIVIDAD_DESDE_CIIU` |

**Otros hallazgos que se mantienen como contexto:**

- Direcciones con formato libre ("C A L L E 8 100 102 C A S A 75"): se guardan como vienen.
- **Teléfonos compartidos** (1.421 números en 3.118 filas): se importan, porque el cliente no pidió descartarlos. Se marcan como `TELEFONO_COMPARTIDO = true` para que el CRM no los trate como una sola persona.
- Hay 403 códigos CIIU distintos, útiles como filtro comercial.

---

## 1. Arquitectura general

```text
 ┌────────────────────────────┐  importación única (sin NIT)
 │ CCH@E26-9316.xlsx (30.780) │─────────────────────────────┐
 └────────────────────────────┘                             ▼
                                              ┌───────────────────────────┐
                                              │ DB_ORIGEN.BASE_CCH        │ copia de solo lectura (sin NIT)
                                              └────────────┬──────────────┘
                                                           │ normalización D1–D9
                                                           ▼
 ┌─────── FORMULARIO PÚBLICO (Web App A) ───────┐    ┌───────────────────────────┐
 │ Sin identificación. Llenar ─► Revisar ─►     │───►│ DB_CORE                   │
 │ Enviar ─► radicado en pantalla               │    │ CLIENTES (30.780 PROSPECTO│
 │ (borrador guardado en el navegador)          │    │  + nuevos ORIGEN=FORMULARIO)│
 └──────────────┬───────────────────────────────┘    │ CONTACTOS, SEDES,         │
                │ archivos                           │ SERVICIOS, PRODUCTOS,     │
                ▼                                    │ TESTIMONIOS, CONTENIDO_WEB│
 ┌───────────────────────────┐                       │ ARCHIVOS, AUTORIZACIONES, │
 │ DRIVE CRM_WEB/01_CLIENTES │◄──── IDs ─────────────│ SOLICITUDES, REPORTES     │
 │  └ 12_REPORTE (Google Doc │                       └────────────┬──────────────┘
 │     editable por cliente) │
 └───────────────────────────┘
                                                                  │
 ┌─────── SUPER ADMIN / CRM (Web App B) ────────┐    ┌────────────▼──────────────┐
 │ Conciliación: solicitud ⇄ prospecto de base  │◄──►│ DB_CRM                    │
 │ Clientes (todo editable) · Reporte por       │    │ OPORTUNIDADES, ACTIVIDADES│
 │ cliente · Duplicados · Archivos · Pipeline   │    │                           │
 │ Kanban · Actividades · Seguimientos          │    │ SEGUIMIENTOS, PROPUESTAS, │
 │ Propuestas · Ventas · KPIs · Auditoría       │    │ VENTAS, PROYECTOS_WEB,    │
 └──────────────────────────────────────────────┘    │ DUPLICADOS                │
                                                     ├───────────────────────────┤
                                                     │ DB_SISTEMA: USUARIOS_ADMIN│
                                                     │ CONFIGURACION, CATALOGOS, │
                                                     │ SECUENCIAS, FORM_CAMPOS,  │
                                                     │ PRODUCTOS_WEB             │
                                                     ├───────────────────────────┤
                                                     │ DB_AUDITORIA_YYYY         │
                                                     └───────────────────────────┘
```

**Flujo de negocio:**

```text
Base existente ─► (CRM: prospección por teléfono/WhatsApp/email con enlace al formulario)
Formulario público (sin identificación) ─► Sheets (datos) + Drive (archivos)
   ─► REPORTE EDITABLE del cliente (Google Doc)
   ─► Super Admin: ver y editar todo ─► conciliación con la base ─► oportunidad
   ─► Seguimiento ─► Propuesta ─► Venta ─► Proyecto web
```

**Principios:**

1. **La base de origen no se modifica.** Las correcciones quedan en `CLIENTES`, con auditoría.
2. **El formulario público nunca lee la base.** Solo escribe su propia solicitud. Así no se exponen datos de los 30.780 registros y no hay riesgo de que alguien averigüe quién está en la base.
3. **Cada envío crea un cliente `ORIGEN = FORMULARIO`**. El sistema propone con qué prospecto de la base podría coincidir (por email, teléfono o nombre + municipio) y **un admin decide** si se fusionan (§5.6). Nunca se fusiona automáticamente.
4. **Dos proyectos Apps Script** (público y admin) con código compartido mediante un paso de build.
5. **Capa `Repo`** que mapea columnas por el nombre del encabezado, para poder migrar el almacenamiento en el futuro.
6. **El servidor valida todo.** La validación en el navegador solo mejora la experiencia.

---

## 2. Modelo de datos

### 2.1 Convenciones

- `CLIENTE_ID` es secuencial (`CLI-000001`…`CLI-030780` para la base; los nuevos continúan con `SECUENCIAS`). El resto de entidades usa `PREFIJO-<base36 timestamp><4 aleatorios>`.
- **Columnas de control en todas las tablas:** `CREADO_EN`, `CREADO_POR`, `ACTUALIZADO_EN`, `ACTUALIZADO_POR`, `VERSION`, `ELIMINADO`. Nunca se borran filas.
- Fechas ISO 8601 en `America/Bogota`. Dinero en COP entero.
- `CREADO_POR` vale `PUBLICO:<SOLICITUD_ID>`, `ADMIN:<email>` o `SISTEMA:<proceso>`.

### 2.2 Entidades

#### CLIENTES (DB_CORE)
| Columna | Notas |
|---|---|
| CLIENTE_ID | PK |
| ORIGEN | `BASE_CCH`, `FORMULARIO`, `ADMIN`, `REFERIDO` |
| ID_REGISTRO_ORIGEN | Fila en `BASE_CCH` (vacío si no viene de la base) |
| NOMBRE_PROPIETARIO | **Vacío en la importación** (D1) |
| RAZON_SOCIAL | De la base o escrita por el cliente |
| NOMBRE_COMERCIAL | Solo lo escribe el cliente o un admin |
| TIPO_ESTABLECIMIENTO / TIPO_ESTABLECIMIENTO_ALT | Catálogo, opciones 1 y 2 de la plantilla |
| ACTIVIDAD | Actividad de la base o descripción CIIU (D6) |
| ACTIVIDAD_FUENTE | `BASE`, `CIIU`, `CLIENTE` |
| CIIU_CODIGO / CIIU_DESCRIPCION | `G4721` / texto |
| TAMANO_EMPRESA | MICRO, PEQUEÑA, MEDIANA, GRAN |
| TELEFONO | 10 dígitos normalizado (D7) |
| TELEFONO_ORIGINAL | Valor de la base |
| TELEFONO_COMPARTIDO | `true` si otro cliente tiene el mismo número |
| EMAIL | Vacío si estaba repetido o mal formado (D5) |
| MUNICIPIO_COD / MUNICIPIO | Uno de los 4 (D9) |
| DIRECCION | |
| BARRIO_COD / BARRIO | |
| FEC_MATRICULA / FEC_RENOVACION / ULT_ANO_RENOVACION | Solo registros de la base |
| ESTADO_COMERCIAL | §5.1 |
| ESTADO_FORMULARIO | §5.1 |
| RESPONSABLE_EMAIL | Admin asignado |
| PRIORIDAD | ALTA, MEDIA, BAJA |
| ETIQUETAS | Separadas por coma |
| FUSIONADO_EN | CLIENTE_ID superviviente tras una conciliación |
| CLAVE_BUSQUEDA | Razón social + nombre comercial + propietario, normalizados (sin tildes ni signos, en mayúsculas) |
| DRIVE_FOLDER_ID | Se crea con la primera subida |
| NO_CONTACTAR | Oposición del titular (habeas data) |
| OBSERVACION_IMPORTACION | §0.2 |
| + control | |

#### CONTACTOS (DB_CORE)
`CONTACTO_ID, CLIENTE_ID, NOMBRE, CARGO, ROL (PROPIETARIO|ADMINISTRADOR|OTRO), TELEFONO_FIJO, WHATSAPP, EMAIL, ES_PRINCIPAL, PREFERENCIA_CANAL (WHATSAPP|LLAMADA|EMAIL), HORARIO_CONTACTO, + control`

#### SEDES (DB_CORE)
`SEDE_ID, CLIENTE_ID, NOMBRE_SEDE, ES_PRINCIPAL, DIRECCION, BARRIO, MUNICIPIO_COD, MUNICIPIO, REFERENCIA, GOOGLE_MAPS_URL, TELEFONO, WHATSAPP, HORARIO_SEMANA, HORARIO_FIN_SEMANA, HORARIOS_ESPECIALES, ATIENDE_EN_SITIO, ORDEN, + control`

#### SERVICIOS (DB_CORE)
`SERVICIO_ID, CLIENTE_ID, NOMBRE, DESCRIPCION, PRECIO, MOSTRAR_PRECIO, REQUIERE_CITA, IMAGEN_ARCHIVO_ID, ORDEN, ACTIVO, + control`

#### PRODUCTOS (DB_CORE) — catálogo del cliente
`PRODUCTO_ID, CLIENTE_ID, NOMBRE, DESCRIPCION, CATEGORIA, PRECIO (costo, según la plantilla), MOSTRAR_PRECIO, DISPONIBLE, IMAGEN_ARCHIVO_ID, ORDEN, + control`

#### TESTIMONIOS (DB_CORE)
`TESTIMONIO_ID, CLIENTE_ID, AUTOR_NOMBRE, TEXTO (≤30 palabras, validado en el servidor), FOTO_ARCHIVO_ID, PERMISO_USO_FOTO, PERMISO_USO_NOMBRE, APROBADO_ADMIN, + control`

#### CONTENIDO_WEB (DB_CORE) — respuestas de la plantilla, clave-valor
`CONTENIDO_ID, CLIENTE_ID, SOLICITUD_ID, SECCION, CAMPO_CLAVE, VALOR, VALOR_TIPO, + control`

Claves que salen de la hoja `Plantilla original`: `nombre_pagina_opcion_1`, `nombre_pagina_opcion_2`, `mensaje_a_transmitir`, `domicilios_ofrece`, `domicilios_costo`, `agenda_citas_desea`, `agenda_citas_version_pago`, `base_datos_clientes_desea`, `base_datos_clientes_version_pago`. Opcionales: `redes_sociales`, `colores_marca`, `sitios_referencia`.

#### ARCHIVOS (DB_CORE)
`ARCHIVO_ID, CLIENTE_ID, SOLICITUD_ID, ENTIDAD, ENTIDAD_ID, CATEGORIA (LOGO|FOTO_NEGOCIO|FOTO_EQUIPO|FOTO_PRODUCTO|FOTO_SERVICIO|FOTO_TESTIMONIO|PROPUESTA|COMPROBANTE|OTRO), NOMBRE_ORIGINAL, NOMBRE_DRIVE, DRIVE_FILE_ID, MIME_DETECTADO, TAMANO_BYTES, SHA256, ESTADO (CUARENTENA|APROBADO|RECHAZADO|PAPELERA), MOTIVO_RECHAZO, + control`

#### AUTORIZACIONES (DB_CORE)
`AUTORIZACION_ID, CLIENTE_ID, SOLICITUD_ID, TIPO (TRATAMIENTO_DATOS|USO_IMAGENES|PUBLICACION_CONTENIDO|CONTACTO_WHATSAPP|CONTACTO_EMAIL|CONTACTO_LLAMADA|USO_RAZON_SOCIAL_BASE|USO_FOTOS_TESTIMONIOS), OTORGADA, TEXTO_VERSION, TEXTO_HASH, FECHA_HORA, FIRMANTE_NOMBRE, USER_AGENT, REVOCADA_EN, + control`

> Como no hay identificación ni confirmación por email (D12), la prueba del consentimiento es el texto aceptado (versión y hash), la hora y el nombre del firmante. El admin puede corroborarlo por teléfono o WhatsApp y dejarlo como actividad.

#### SOLICITUDES (DB_CORE) — cada formulario diligenciado
`SOLICITUD_ID, CLIENTE_ID, ESTADO (BORRADOR|ENVIADA|EN_REVISION|REQUIERE_CORRECCION|APROBADA|DESCARTADA_SPAM), PASO_ACTUAL, PORCENTAJE, RADICADO (WEB-2026-000123), TOKEN_HASH, ENVIADA_EN, SNAPSHOT_FILE_ID, CONCILIACION_ESTADO (PENDIENTE|VINCULADA|NUEVO_CLIENTE), CLIENTE_BASE_VINCULADO, OBSERVACIONES_ADMIN, REVISADO_POR, REVISADO_EN, + control`

#### OPORTUNIDADES (DB_CRM)
`OPORTUNIDAD_ID, CLIENTE_ID, TITULO, PRODUCTO_WEB_ID, ETAPA, PROBABILIDAD, VALOR_ESTIMADO (editable), FECHA_CIERRE_ESTIMADA, ORIGEN_LEAD, RESPONSABLE_EMAIL, MOTIVO_PERDIDA, FECHA_CIERRE_REAL, ORDEN_KANBAN, + control`

#### ACTIVIDADES (DB_CRM)
`ACTIVIDAD_ID, CLIENTE_ID, OPORTUNIDAD_ID, TIPO (LLAMADA|WHATSAPP|EMAIL|REUNION|NOTA|VISITA|CAMBIO_ETAPA), DIRECCION (ENTRANTE|SALIENTE), RESULTADO (CONTESTO|NO_CONTESTO|NUMERO_ERRADO|INTERESADO|NO_INTERESADO|VOLVER_A_LLAMAR|ENVIADO|RESPONDIDO|REALIZADA|CANCELADA), ASUNTO, DETALLE, DURACION_MIN, FECHA_HORA, CALENDAR_EVENT_ID, REALIZADA_POR, + control`

#### SEGUIMIENTOS (DB_CRM)
`SEGUIMIENTO_ID, CLIENTE_ID, OPORTUNIDAD_ID, ACTIVIDAD_ORIGEN_ID, TIPO (LLAMAR|WHATSAPP|EMAIL|REUNION|ENVIAR_PROPUESTA|REVISAR_SOLICITUD|COBRAR|OTRO), DESCRIPCION, FECHA_VENCIMIENTO, PRIORIDAD, ESTADO (PENDIENTE|COMPLETADO|VENCIDO|CANCELADO), ASIGNADO_A, COMPLETADO_EN, + control`

#### PROPUESTAS (DB_CRM)
`PROPUESTA_ID, OPORTUNIDAD_ID, CLIENTE_ID, NUMERO (PRO-2026-0001), ITEMS_JSON (descripción, cantidad y precio unitario, todos editables), SUBTOTAL, DESCUENTO, TOTAL, VALIDEZ_HASTA, ESTADO (BORRADOR|ENVIADA|ACEPTADA|RECHAZADA|VENCIDA), PDF_ARCHIVO_ID, ENVIADA_EN, + control`

#### VENTAS (DB_CRM)
`VENTA_ID, OPORTUNIDAD_ID, PROPUESTA_ID, CLIENTE_ID, NUMERO (VEN-2026-0001), FECHA_VENTA, VALOR_TOTAL, FORMA_PAGO, ESTADO_PAGO (PENDIENTE|PARCIAL|PAGADA|ANULADA), VALOR_PAGADO, COMPROBANTE_ARCHIVO_ID, VENDEDOR_EMAIL, PROYECTO_ID, + control`

#### PROYECTOS_WEB (DB_CRM)
`PROYECTO_ID, VENTA_ID, CLIENTE_ID, ESTADO (BRIEF|DISENO|DESARROLLO|REVISION_CLIENTE|PUBLICADO|MANTENIMIENTO), BRIEF_ARCHIVO_ID, DOMINIO, URL_PUBLICADA, FECHA_ENTREGA_ESTIMADA, RESPONSABLE_EMAIL, + control`

#### DUPLICADOS (DB_CRM) — también sirve para la conciliación
`DUP_ID, CLIENTE_A, CLIENTE_B, TIPO (DUPLICADO|CONCILIACION_SOLICITUD), REGLA, PUNTAJE, ESTADO (PENDIENTE|ES_EL_MISMO|NO_ES_EL_MISMO|FUSIONADO), RESUELTO_POR, RESUELTO_EN`

#### USUARIOS_ADMIN (DB_SISTEMA)
`USUARIO_ID, EMAIL (cuenta Google), NOMBRE, ROL (SUPER_ADMIN|ADMIN|VENDEDOR|LECTURA), ACTIVO, MUNICIPIOS_ASIGNADOS, ULTIMO_ACCESO, + control`

#### PRODUCTOS_WEB (DB_SISTEMA) — lo que vende Webpaya (precios sugeridos y editables, D14)
`PRODUCTO_WEB_ID, NOMBRE (ej. Página básica, Catálogo, Agenda de citas, Base de datos de clientes), DESCRIPCION, PRECIO, ES_ADICIONAL, ES_VERSION_PAGO, ACTIVO, ORDEN`

`PRECIO` es opcional y editable en cualquier momento desde Configuración. Al crear una propuesta u oportunidad **se precarga como sugerencia** y el admin lo cambia libremente en cada caso. Si está vacío, el campo aparece en blanco para escribirlo.

#### REPORTES (DB_CORE) — el reporte editable de cada cliente (D13)
`REPORTE_ID, CLIENTE_ID, SOLICITUD_ID, VERSION_REPORTE (1, 2, 3…), DOC_ID, DOC_URL, PDF_ARCHIVO_ID (opcional), GENERADO_POR (SISTEMA|ADMIN:email), GENERADO_EN, ULTIMA_EDICION_DOC (fecha de modificación leída de Drive), EDITADO_MANUALMENTE (bool), ESTADO (PENDIENTE|GENERADO|ERROR|REEMPLAZADO), + control`

#### CONFIGURACION, CATALOGOS, SECUENCIAS, FORM_CAMPOS (DB_SISTEMA)
- `CONFIGURACION` (clave-valor): `MAX_UPLOAD_MB` (10), `MAX_TOTAL_MB_SOLICITUD` (100), `TEXTO_AUTORIZACION_VERSION`, `ETAPAS_PIPELINE_JSON`, `RESPONSABLE_TRATAMIENTO` (= "Equipo Webpaya"), `PLANTILLA_REPORTE_DOC_ID`, `MANTENIMIENTO`.
- `CATALOGOS`: `MUNICIPIOS` (solo 4), `TIPOS_ESTABLECIMIENTO`, `ETAPAS`, `MOTIVOS_PERDIDA`, `RESULTADOS_ACTIVIDAD`.
- `SECUENCIAS`: radicados, `CLIENTE_ID` nuevos, propuestas y ventas (bajo `LockService`).
- `FORM_CAMPOS`: definición de pasos y campos del formulario.

#### AUDITORIA (DB_AUDITORIA_YYYY) — solo se agregan filas
`AUDIT_ID, FECHA_HORA, ACTOR, APP (PUBLICA|ADMIN|TRIGGER), ACCION, ENTIDAD, ENTIDAD_ID, CAMPOS_CAMBIADOS, VALOR_ANTERIOR_JSON, VALOR_NUEVO_JSON, RESULTADO (OK|DENEGADO|ERROR), DETALLE`

### 2.3 Relaciones

```text
CLIENTES 1─N CONTACTOS | SEDES | SERVICIOS | PRODUCTOS | TESTIMONIOS | CONTENIDO_WEB | ARCHIVOS | AUTORIZACIONES | SOLICITUDES | REPORTES
CLIENTES 1─N OPORTUNIDADES 1─N ACTIVIDADES / SEGUIMIENTOS / PROPUESTAS
OPORTUNIDADES 1─0..1 VENTAS 1─0..1 PROYECTOS_WEB
PRODUCTOS_WEB 1─N OPORTUNIDADES
CLIENTES N─N CLIENTES (DUPLICADOS / conciliación)
```

### 2.4 Distribución en hojas de cálculo

| Spreadsheet | Contenido |
|---|---|
| `DB_ORIGEN` | `BASE_CCH`: copia de solo lectura, sin NIT |
| `DB_CORE` | Clientes y todo lo que escribe el formulario |
| `DB_CRM` | Lo que escriben los admins |
| `DB_SISTEMA` | Configuración, usuarios, catálogos y productos web |
| `DB_AUDITORIA_YYYY` | Auditoría, rota por año |

Tamaño estimado: `CLIENTES` tiene unas 38 columnas × 31.000 filas ≈ 1,2 M celdas, dentro del límite de 10 M por archivo.

---

## 3. Estructura de Google Drive

```text
CRM_WEB/                                   (cuenta @gmail.com propietaria; nada compartido públicamente)
├── 00_SISTEMA/
│   ├── DB_ORIGEN · DB_CORE · DB_CRM · DB_SISTEMA   (Sheets)
│   ├── AUDITORIA/DB_AUDITORIA_2026 …
│   ├── INDICES/                           (índice de búsqueda generado)
│   └── BACKUPS/YYYY-MM-DD/                (copias nocturnas; se conservan 14 por la cuota de 15 GB)
├── 01_CLIENTES/
│   └── CLI-030xxx/                        ← grupos de 1.000 clientes
│       └── CLI-030801__PANADERIA-EL-TRIGAL/
│           ├── 01_LOGO/
│           ├── 02_FOTOS_NEGOCIO/
│           ├── 03_FOTOS_EQUIPO/
│           ├── 04_PRODUCTOS/
│           ├── 05_SERVICIOS/
│           ├── 06_SEDES/
│           ├── 07_TESTIMONIOS/
│           ├── 08_SOLICITUDES/            (snapshot JSON de cada envío)
│           ├── 09_PROPUESTAS/
│           ├── 10_VENTAS/
│           ├── 11_PROYECTO_WEB/
│           └── 12_REPORTE/                (REPORTE_CLI-030801_v1, _v2… Google Docs editables)
├── 02_CUARENTENA/YYYY-MM-DD/              (entrada de toda subida pública)
├── 03_PLANTILLAS/                         (PLANTILLA_REPORTE_CLIENTE (Google Doc), propuesta, textos legales)
└── 04_EXPORTS/                            (se borran a los 7 días)
```

**Reglas:**

- **Carpetas perezosas:** los 30.780 prospectos no tienen carpeta hasta que la necesiten.
- Cuando el admin concilia una solicitud con un prospecto de la base, los archivos **se mueven** a la carpeta del cliente superviviente.
- Nombre en Drive: `<CATEGORIA>_<ENTIDAD_ID>_<yyyyMMddHHmmss>.<ext>`. El nombre original solo se guarda en `ARCHIVOS`.
- Nada se comparte con enlace público.
- Papelera lógica. Un trigger mensual limpia lo que tenga más de 90 días.

---

## 4. Flujo del formulario público (sin identificación)

Es una SPA en `HtmlService` con pasos, guardado automático y un **token de solicitud** que da acceso únicamente a ese borrador. **Todos los campos son editables** y ninguno viene precargado (D11). El cliente puede volver a cualquier paso y cambiar lo que quiera hasta enviar.

| # | Paso | Contenido | Reglas |
|---|---|---|---|
| 0 | Bienvenida | Qué es, tiempo estimado, aviso de privacidad de **Equipo Webpaya** y enlace a la política | Aceptar el aviso para continuar |
| 1 | Datos del negocio | Razón social o nombre del establecimiento\*, nombre del propietario, nombre comercial (opcional), tipo de establecimiento (opciones 1 y 2), nombre de página deseado (opciones 1 y 2) | Al guardar este paso se crea el borrador y el token |
| 2 | Contacto | Email (solo como dato, D12), teléfono fijo, WhatsApp (\* al menos uno de los dos), canal preferido | Fijo de 7 dígitos → se antepone 608 automáticamente. Celular de 10 dígitos. |
| 3 | Ubicación y horarios | Municipio (**solo los 4**), dirección, barrio. Sede 1 y opcional Sede 2+. Horario entre semana, fin de semana, especiales. Atiende en sitio. | |
| 4 | Mensaje | ¿Qué le gustaría transmitir en su página web? | Texto |
| 5 | Imágenes | Logo, fotos del negocio y del equipo | §6.4 |
| 6 | Servicios | Lista repetible (Servicio 1, 2, …) | |
| 7 | Productos | Catálogo repetible con costo. Domicilios (sí/no + costo). | |
| 8 | Testimonios | Cliente, texto ≤30 palabras, foto, permiso de uso de la foto | La foto exige el permiso |
| 9 | Servicios adicionales | Agenda de citas (sí/no, versión de pago). Base de datos de clientes (sí/no, versión de pago). | Señales de venta para el CRM |
| 10 | Autorizaciones | Casillas separadas y **desmarcadas**: tratamiento de datos (obligatoria), uso de imágenes, publicación del contenido, contacto por WhatsApp, email y llamada, "¿Permite el uso de su razón social para alimentar nuestra base de datos?" | Versión y hash del texto |
| 11 | Revisión | Resumen con enlaces para editar cada paso | |
| 12 | Envío | Radicado `WEB-2026-000123` **en pantalla** (sin email) y snapshot JSON en Drive | Crea el cliente `ORIGEN = FORMULARIO`, la solicitud `ENVIADA`, las propuestas de conciliación, un seguimiento `REVISAR_SOLICITUD` y **pone en cola la generación del reporte** (§4.1) |

\* = obligatorio. Los obligatorios se pueden ajustar en `FORM_CAMPOS`.

**Continuar después:** el token del borrador se guarda en `localStorage` del navegador, así que en el mismo dispositivo y navegador se retoma automáticamente. **No hay enlace por email** (D12). Si el cliente cambia de dispositivo, empieza de nuevo o termina en el primero. El token se invalida al enviar.

### 4.1 Reporte editable por cliente (D13)

**Propósito:** llevar la información del cliente **al software de diseño con el que se construye su página web**. Por eso el reporte está pensado para **copiar y pegar sección por sección**. No es un informe para leer.

**Formato: Google Doc** en la carpeta `12_REPORTE` del cliente. El admin lo edita directamente en Google Docs y puede descargarlo como PDF, Word o texto.

**Reglas de diseño para facilitar el copiado:**

- Está **ordenado como las secciones de una página web**, no como el formulario.
- Cada dato va en un bloque `ETIQUETA` + valor en texto plano, sin tablas anidadas ni formato que se pierda al pegar.
- Los textos largos van completos y sin cortes. Los precios van ya formateados (`$ 25.000`).
- Las imágenes aparecen en miniatura para ubicarlas, **con el enlace a la carpeta de Drive donde está el archivo original en resolución completa** (logo, fotos por sección). Así el diseñador descarga los originales y no una copia reducida.
- Se omiten los campos vacíos, para no pegar huecos. El admin puede añadir texto donde quiera.

**Estructura (sección de página → contenido):**

| # | Sección web | Contenido del reporte |
|---|---|---|
| 0 | Ficha de trabajo | Radicado, fecha, CLIENTE_ID, enlace a la carpeta Drive y nombres de página deseados (opciones 1 y 2) |
| 1 | Encabezado / menú | Logo (enlace al original), nombre comercial o razón social |
| 2 | Portada (hero) | "¿Qué le gustaría transmitir?", tipo de establecimiento y botón de WhatsApp (número listo con formato `wa.me/57…`) |
| 3 | Quiénes somos | Propietario, tipo de establecimiento, actividad y fotos del equipo |
| 4 | Servicios | Un bloque por servicio: nombre, descripción, precio (si se muestra) e imagen |
| 5 | Catálogo / productos | Un bloque por producto: nombre, descripción, costo, categoría e imagen. Domicilios (sí/no + costo). |
| 6 | Testimonios | Solo los que tienen permiso: autor, texto (≤30 palabras) y foto |
| 7 | Galería | Fotos del negocio (enlace a la carpeta) |
| 8 | Ubicación y horarios | Por sede: dirección, barrio, municipio, enlace a Google Maps, horario entre semana, fin de semana y especiales |
| 9 | Contacto / pie de página | Teléfono fijo, WhatsApp, email y redes |
| 10 | Funciones adicionales | Agenda de citas y base de datos de clientes (sí/no, versión de pago), para saber qué módulos activar |
| 11 | Permisos | Autorizaciones otorgadas, sobre todo uso de imágenes y publicación de contenido: **qué se puede publicar y qué no** |
| 12 | Notas de Webpaya | Espacio libre para el diseñador o el admin |

**Exportación opcional:** junto al Doc se puede generar `DATOS_<CLIENTE_ID>.json` con la misma estructura. Sirve si el software de diseño importa datos o si más adelante se automatiza la creación de la página (§9).

**Cuándo se genera:**

- **Automáticamente al enviar el formulario.** Un trigger del proyecto admin revisa cada 5 minutos las solicitudes `ENVIADA` sin reporte y lo genera. Así el cliente no espera.
- **A demanda:** botón "Generar/Regenerar reporte" en la ficha del cliente. Usa los datos **actuales** de las hojas, incluidas las ediciones del admin.

**Regla de versiones:** regenerar **nunca sobrescribe** un Doc existente. Crea `_v2`, `_v3`… y marca el anterior como `REEMPLAZADO`. Si el Doc anterior fue editado a mano (`ULTIMA_EDICION_DOC` > `GENERADO_EN`), se avisa antes de regenerar para no perder ese trabajo.

**Dos niveles de edición:**

- **Datos** (Sheets): se editan en la ficha del cliente del Super Admin y alimentan el CRM y las versiones futuras del reporte.
- **Reporte** (Doc): se edita libremente en Google Docs. Sus cambios **no** vuelven a las hojas; es el documento final de trabajo.

---

## 5. Flujo del CRM

### 5.1 Estados

**ESTADO_FORMULARIO:** `NO_INICIADO` (los 30.780 importados) → `EN_PROGRESO` → `ENVIADO` → `EN_REVISION` → (`REQUIERE_CORRECCION` ↺) → `APROBADO`

**ESTADO_COMERCIAL:** `PROSPECTO → CONTACTADO → INTERESADO → CLIENTE → INACTIVO`, más `NO_INTERESADO`, `NO_CONTACTAR` y `DATOS_ERRADOS`.

**ETAPA de OPORTUNIDAD** (Kanban, configurable):

| Etapa | Prob. | Entrada |
|---|---|---|
| NUEVA | 5 % | Creada por un admin |
| CONTACTO_INICIAL | 10 % | Primera actividad saliente |
| FORMULARIO_RECIBIDO | 30 % | Automática al enviar el formulario |
| BRIEF_VALIDADO | 45 % | El admin aprueba la solicitud |
| PROPUESTA_ENVIADA | 60 % | La propuesta pasa a `ENVIADA` |
| NEGOCIACION | 75 % | Manual |
| GANADA | 100 % | Venta obligatoria |
| PERDIDA | 0 % | Motivo obligatorio |

### 5.2 Reglas automáticas

- Primera actividad saliente sobre un `PROSPECTO` → `CONTACTADO`.
- Envío del formulario → cliente nuevo + oportunidad en `FORMULARIO_RECIBIDO` + seguimiento a 24 h + candidatos de conciliación.
- Conciliación aceptada → los datos del formulario se fusionan en el prospecto de la base (el admin elige campo por campo). Las oportunidades, archivos y solicitudes se reasignan al superviviente y el otro queda con `FUSIONADO_EN`.
- `GANADA` exige una venta; la venta `PAGADA` crea el `PROYECTOS_WEB` en `BRIEF`.
- Resultado `VOLVER_A_LLAMAR` exige un seguimiento con fecha.
- Trigger diario: los seguimientos vencidos pasan a `VENCIDO` y se muestran en el dashboard. No se envía email (D12).
- `NO_CONTACTAR` bloquea las actividades salientes en el servidor.

### 5.3 Kanban

Una columna por etapa. Mover una tarjeta llama a `moverOportunidad(id, etapa, versionEsperada)`, con control de conflicto por `VERSION`. Cada movimiento queda como actividad `CAMBIO_ETAPA` y en la auditoría. `GANADA` y `PERDIDA` abren un modal obligatorio.

### 5.4 Actividades y comunicación

- **Llamada:** enlace `tel:` y registro del resultado.
- **WhatsApp:** enlace `wa.me/57…` con un mensaje-plantilla que incluye **el enlace al formulario público**. Así se invita a los 30.780 prospectos. Es el canal principal para los 2.110 que no tienen email.
- **Email:** el sistema **no envía correos** (D12). Hay un botón `mailto:` que abre el correo del admin y la actividad se registra manualmente.
- **Reunión:** evento en Calendar con `CalendarApp`.

### 5.5 Vistas del Super Admin

1. **Bandeja de entrada de formularios:** lo último que enviaron los clientes, con acceso directo a la ficha y al reporte. **Es la vista principal.**
2. **Dashboard:** embudo (prospectos → contactados → formularios enviados → propuestas → ventas), pipeline ponderado, ventas del mes, seguimientos vencidos, actividad por admin y distribución por los 4 municipios y por CIIU.
3. **Conciliación:** cada solicitud nueva con sus candidatos de la base, comparados lado a lado.
4. **Clientes:** tabla paginada en el servidor. Búsqueda por nombre, email o teléfono. Filtros por municipio, estados, CIIU, tamaño, año de renovación, responsable, "sin email" y "teléfono compartido". Ficha 360° con **edición en línea de todos los campos** (D11): datos importados, datos del formulario, sedes, servicios, productos, testimonios y archivos (reemplazar o eliminar). Cada guardado controla conflictos por `VERSION` y deja en la auditoría el valor anterior y el nuevo. Botones **Generar/Regenerar reporte** y **Abrir reporte**.
5. **Duplicados**, **Archivos** (aprobar o rechazar), **Pipeline/Kanban**, **Seguimientos**, **Propuestas**, **Ventas**.
6. **Auditoría** (solo `SUPER_ADMIN`) y **Configuración** (catálogos, productos web y precios, usuarios, textos legales).

### 5.6 Conciliación y duplicados (sin NIT)

Se normaliza (mayúsculas, sin tildes ni signos, sin sufijos societarios) y se aplican estas reglas:

| Regla | Puntaje | Observación |
|---|---|---|
| Mismo email | 90 | Fiable, porque en la base solo quedaron emails únicos (D5) |
| Mismo teléfono **y** nombre similar | 85 | |
| Nombre similar (trigramas ≥ 0,85) + mismo municipio | 70 | |
| Mismo teléfono y el número no está marcado como compartido | 60 | |
| Mismo teléfono y el número está marcado como compartido | 20 | Solo informativo |

La conciliación se ejecuta al enviar cada solicitud, sobre un índice en caché. La detección de duplicados dentro de la base se ejecuta en lotes nocturnos. **Toda fusión es manual.**

---

## 6. Autenticación y seguridad de acceso

### 6.1 Super Admin (Web App B)

- Se despliega con **Ejecutar como: usuario que accede** y **Acceso: cualquier persona con cuenta de Google**. Con una cuenta @gmail.com es la única forma de conocer el email real de quien entra (`Session.getActiveUser()`).
- En cada llamada, `requireRole(rol)` comprueba el email contra `USUARIOS_ADMIN` y lo audita. A quien no está autorizado se le muestra una página neutra.
- Los spreadsheets y `CRM_WEB/` se comparten **solo** con los admins, nunca con enlace.
- Roles aplicados en el servidor: `SUPER_ADMIN`, `ADMIN`, `VENDEDOR` y `LECTURA`.
- Se recomienda la verificación en dos pasos de Google en toda cuenta admin.
- Limitación: un admin puede abrir el Sheet directamente. Es aceptable con admins de confianza (§9).

### 6.2 Formulario público (Web App A)

- Se despliega con **Ejecutar como: yo** y **Acceso: cualquiera, incluso anónimo**.
- **No hay identificación.** El único control de acceso es el **token de solicitud**: 128 bits aleatorios, guardados en `SOLICITUDES` solo como hash SHA-256.
- Toda función pública recibe el token y deriva de él la `SOLICITUD_ID`. **Nunca acepta IDs como parámetro**, lo que evita IDOR (acceder a datos ajenos cambiando un ID).
- **El formulario público no tiene ninguna función que lea `CLIENTES` ni la base.** Solo puede crear y editar su propia solicitud. Esto elimina la exposición de los 30.780 registros y la enumeración.
- Después del envío, el token deja de servir para editar.

### 6.3 Anti-spam y abuso (al no haber identificación)

- Campo honeypot, tiempo mínimo de llenado y límite de creación de borradores por hora (global, en `CacheService`).
- Límites por solicitud: número de servicios, productos, testimonios y megabytes.
- Estado `DESCARTADA_SPAM` y botón para marcarla desde el admin.
- reCAPTCHA v3 queda como opción si aparece spam real.

### 6.4 Archivos subidos

- Validación **en el servidor**: tamaño ≤10 MB por archivo y ≤100 MB por solicitud. Extensiones `jpg`, `jpeg`, `png`, `webp` y `pdf` (**sin SVG**). Firma mágica de los primeros bytes y SHA-256.
- Redimensionado en el navegador a 1.920 px antes de subir, para ahorrar cuota.
- Todo entra por `02_CUARENTENA`. Se transporta en base64 por fragmentos. Nunca se sirve con enlace público.

---

## 7. Seguridad: riesgos y medidas

| Área | Riesgo | Medida |
|---|---|---|
| Apps Script | El web app anónimo corre con los permisos del propietario | Proyecto separado, scopes mínimos, solo las funciones `api_*` expuestas (las privadas terminan en `_`), sin `eval`. |
| Apps Script | XSS | `textContent` siempre; prohibido `innerHTML` con datos. |
| Sheets | Inyección de fórmulas (`=HYPERLINK…`) | Todo texto que empiece por `= + - @` se guarda con un apóstrofo delante. Las exportaciones CSV también se sanean. |
| Sheets | Escrituras concurrentes | `LockService`, búsqueda por ID y `VERSION`. |
| Sheets / Drive | Compartir por error | Revisión semanal automática de permisos con alerta si aparece acceso público. |
| Drive | Archivos maliciosos o pesados | §6.4. |
| Sesiones | Robo del token del borrador | Solo da acceso a ese borrador, expira a los 7 días, se invalida al enviar y nunca da acceso a la base. |
| Formulario abierto | Spam o datos falsos | §6.3 y revisión del admin antes de cualquier acción comercial. |
| Datos personales (Ley 1581/2012, Decreto 1377/2013) | Uso comercial de 30.780 registros del registro mercantil | Responsable: **Equipo Webpaya**. Política publicada, aviso en el formulario, autorización expresa con prueba, canal de consulta y supresión (`NO_CONTACTAR`) y minimización (**el NIT no se almacena**). Se recomienda revisión legal y evaluar el registro en el RNBD. |
| Edición total (D11) | Un cambio erróneo del admin sobrescribe un dato bueno | Auditoría con el valor anterior, botón "deshacer" desde la auditoría y backups nocturnos. |
| Datos de terceros | Fotos de clientes en los testimonios | Permiso explícito y aprobación del admin. |
| Exportaciones | Fuga masiva | Solo `SUPER_ADMIN`, auditadas, con caducidad de 7 días. |

---

## 8. Arquitectura clasp (estructura de archivos)

Monorepo con **código compartido copiado en el build**, no una *Library* de Apps Script.

```text
crm-web/
├── package.json                 # build, push:public, push:admin, deploy:*, lint, test
├── .eslintrc.json  .prettierrc  .gitignore (dist/, .clasprc.json)
├── scripts/build.mjs            # shared/ + apps/<x>/ → dist/<x>/
├── docs/FASE1_ARQUITECTURA.md
├── shared/
│   ├── 00_Config.js             # IDs desde PropertiesService + CONFIGURACION
│   ├── 01_Utils.js              # ids, fechas, normalización (tildes, teléfono 608, email), anti-fórmulas
│   ├── 02_Repo.js               # encabezados→objetos, findById, TextFinder, insert, update(versión), softDelete
│   ├── 03_Schema.js             # tablas y columnas (fuente única)
│   ├── 04_Audit.js
│   ├── 05_Validators.js         # ≤30 palabras, email, teléfono, municipio ∈ 4
│   ├── 06_DriveStore.js         # carpetas perezosas, cuarentena, firma mágica, hash
│   ├── 07_ReportBuilder.js      # Google Doc desde plantilla: marcadores {{campo}}, tablas, imágenes
│   └── 08_Errors.js             # respuesta {ok, data, error}
├── apps/
│   ├── public/
│   │   ├── .clasp.json  appsscript.json   # USER_DEPLOYING, ANYONE_ANONYMOUS
│   │   ├── server/ Main.js · Api.js (api_crearBorrador, api_guardarPaso, api_subirFragmento,
│   │   │           api_enviar) · TokenService.js ·
│   │   │           AntiSpam.js · FormService.js · UploadService.js
│   │   └── client/ index.html · styles.html · app.js.html · steps/*.html · components/*.html
│   └── admin/
│       ├── .clasp.json  appsscript.json   # USER_ACCESSING, ANYONE (cuenta Google)
│       ├── server/ Main.js · Api.js · AuthAdmin.js · ClientesService.js · ConciliacionService.js ·
│       │           DuplicadosService.js · ArchivosService.js · ReporteService.js · OportunidadesService.js ·
│       │           ActividadesService.js · SeguimientosService.js · PropuestasService.js ·
│       │           VentasService.js · KpiService.js · AuditService.js · SearchIndex.js · Triggers.js
│       ├── setup/ Install.js (idempotente) · ImportBase.js (lotes de 2.000 reanudables, reglas D1–D9)
│       └── client/ index.html · styles.html · app.js.html · views/*.html
├── tests/                       # Jest: normalización 608, emails repetidos, CIIU, municipios, reglas de conciliación
└── dist/                        # generado; rootDir de clasp
```

- **Configuración sensible** (IDs de DB y carpetas) en Script Properties, nunca en el repositorio. `.clasprc.json` nunca se sube.
- **Despliegues** con versión fija (`clasp deploy -i <id>`) para que la URL pública no cambie.
- **Validación previa de la importación:** el mismo algoritmo de `ImportBase` se prueba en Jest y debe reproducir los conteos de §0.1: 28.670 emails, 2.110 vacíos, 1.496 teléfonos con 608 y 7.754 actividades desde el CIIU.

---

## 9. Escalabilidad

| Capacidad futura | Preparación |
|---|---|
| **Múltiples admins** | Roles, auditoría por actor y bloqueos optimistas desde el inicio. |
| **Vendedores** | `RESPONSABLE_EMAIL`, `MUNICIPIOS_ASIGNADOS` y filtros por rol. Para que un vendedor **no** pueda abrir el Sheet completo hace falta Google Workspace o un backend propio. La capa `Repo` hace ese cambio local. |
| **Propuestas** | Plantilla de Google Docs → PDF, con precios de `PRODUCTOS_WEB`. |
| **Pagos** | Tabla `PAGOS` y webhook de pasarela (Wompi, ePayco o Mercado Pago) en un tercer proyecto con firma HMAC. |
| **Gestión de proyectos** | `PROYECTOS_WEB` + `TAREAS_PROYECTO`, reutilizando el Kanban. |
| **IA** | `AiService` con cola `IA_TRABAJOS` procesada por un trigger. La clave de API va en Script Properties. |
| **Brief automático** | A partir del snapshot JSON de la solicitud → Google Doc en `11_PROYECTO_WEB`, al llegar a `BRIEF_VALIDADO`. |
| **Contenido web** | Textos generados desde `CONTENIDO_WEB`, servicios, productos y testimonios, con aprobación del admin. |
| **Volumen** | Rotación de `ACTIVIDADES` y `AUDITORIA`. Migración de tablas calientes detrás del `Repo`. |
| **Dominio propio** | La API pública responde JSON uniforme, así que el formulario podría servirse desde GitHub Pages (el repositorio ya tiene `CNAME`) llamando a `doPost`. |

---

## 10. Riesgos y mitigaciones

| # | Riesgo | Mitigación |
|---|---|---|
| R1 | ~~Cuota de email de @gmail.com~~ | **Eliminado**: el sistema no envía correos (D12). Las invitaciones a prospectos van por WhatsApp o llamada. |
| R1b | El cliente pierde el borrador si cambia de dispositivo o borra el navegador (no hay enlace por email) | Autosave en el servidor con cada paso, formulario corto por pasos y aviso visible de "termine en este mismo dispositivo". El admin ve también los borradores sin enviar en la bandeja (filtro `BORRADOR`) y puede completarlos o contactar al cliente. |
| R1c | Generar reportes consume tiempo de trigger (90 min/día en @gmail.com) | Unos 10–20 s por reporte: caben más de 200 al día. Generación en cola, con reintento en caso de error. |
| R2 | Formulario abierto → spam o datos falsos | §6.3 y revisión del admin. |
| R3 | La vinculación con la base es menos precisa sin NIT ni identificación | Conciliación por email único, teléfono y nombre + municipio, siempre con decisión manual (§5.6). |
| R4 | 2.110 prospectos sin email (D5) y unos 50 sin teléfono válido | Contacto por teléfono o WhatsApp, filtro "sin email" en el CRM. El admin puede completar el email a mano (D11). |
| R5 | Teléfonos compartidos (3.118 filas) | Marca `TELEFONO_COMPARTIDO` y puntaje bajo en la conciliación. |
| R6 | Límite de 6 minutos y lectura lenta de 31.000 filas | `TextFinder` para búsquedas puntuales, índice compacto gzip en `CacheService` y paginación en el servidor. |
| R7 | Unas 30 ejecuciones simultáneas, contando todas las anónimas contra el propietario | Autosave con *debounce*, subidas secuenciales e invitaciones escalonadas. |
| R8 | Escrituras que colisionan | `LockService`, `VERSION` y reintentos con espera creciente. |
| R9 | **15 GB de almacenamiento en @gmail.com** (compartidos con Gmail y Fotos) | Redimensionado de imágenes, límite por solicitud, backups limitados a 14 y alerta al 80 %. Google One o Workspace si se llena. |
| R10 | Cumplimiento de habeas data | §7. Equipo Webpaya como responsable. Revisión legal antes de lanzar. |
| R11 | Una cuenta @gmail.com personal como único propietario | Cuenta dedicada a Webpaya (no personal) con verificación en dos pasos, backups y procedimiento de recuperación documentado. |
| R12 | Triggers de @gmail.com limitados a 90 min/día | Triggers cortos e idempotentes. Importación inicial y detección de duplicados en lotes repartidos en varios días si fuera necesario. |
| R13a | Regenerar un reporte borra las ediciones hechas a mano en el Doc | Nunca se sobrescribe: se crean versiones y se avisa si el Doc fue editado (§4.1). |
| R13 | Cambios en la plantilla del formulario | `FORM_CAMPOS` + `CONTENIDO_WEB` clave-valor. |
| R14 | Autorizaciones inválidas | Casillas separadas y desmarcadas, con texto versionado y hash. |

---

## 11. Pendiente antes de la Fase 2

1. ~~Precios~~ → **Resuelto (D14):** espacio de precios editable. No hace falta definir valores para implementar.
2. ~~Destino del reporte~~ → **Resuelto:** es para pasar la información al software de diseño de la página web (§4.1).
3. ~~Campos obligatorios~~ → **Resuelto:** no todos son obligatorios. Solo lo son la razón social o nombre del establecimiento, **un medio de contacto** (WhatsApp o teléfono fijo) y la autorización de tratamiento de datos (exigida por la Ley 1581). **Todo lo demás es opcional.** Se ajusta en `FORM_CAMPOS` sin tocar código.
4. ~~Software de diseño~~ → **Aún no definido (D16):** se implementa la estructura genérica. Adaptarla después solo requiere cambiar la plantilla del reporte.

**No quedan decisiones bloqueantes para la Fase 2.**
---

*Fin de la FASE 1. En espera de la instrucción para implementar.*

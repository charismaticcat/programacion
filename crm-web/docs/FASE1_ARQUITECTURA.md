# FASE 1 — Arquitectura y diseño
### Aplicativo público de captura + Super Admin / CRM (Google Apps Script · Sheets · Drive · clasp)

> Estado: **diseño**. No contiene código de implementación. Pendiente de aprobación para pasar a FASE 2.

---

## 0. Hallazgos del análisis de los archivos

Se analizaron las 30.780 filas completas de ambos archivos.

### 0.1 `CCH@E26-9316.xlsx` (hoja `Hoja1`, 14 columnas, 30.780 filas)

| Columna | Hallazgo | Consecuencia de diseño |
|---|---|---|
| `EST-MATRICULA` | 100 % `MA` (matrícula activa). | Se conserva solo como dato de origen. |
| `RAZON SOCIAL` | Mezcla personas naturales y jurídicas. Los nombres naturales vienen **en dos órdenes** ("OSCAR JAIR ARCINIEGAS ROMERO" y "TORRENTE TRUJILLO JORGE ENRIQUE"). Hay dobles espacios y `\xa0`. | No se separan nombres/apellidos automáticamente. Se normalizan espacios solo para buscar. |
| `NIT` | Número entero con el **dígito de verificación pegado al final** (ej. `8130054227` = NIT 813005422-7). El DV valida con el algoritmo DIAN en 30.746 de 30.767 filas. 13 filas sin NIT y 2 NIT duplicados. | Se guardan `DOC_NUMERO` (sin DV) y `DOC_DV` por separado. La identificación del cliente busca por `DOC_NUMERO`. |
| `NIT` → tipo de persona | Los NIT de 10 dígitos que empiezan por 8 o 9 (7.761 filas) son personas jurídicas. Esta regla es **más fiable que buscar sufijos** ("SAS", "LTDA"): el sufijo da falsos positivos como "SA**S**TOQUE" y no detecta las "S EN C". | `TIPO_PERSONA` se infiere por el prefijo del NIT, con el sufijo como regla secundaria, y **siempre lo confirma el cliente**. |
| `FEC-MATRICULA`, `FEC-RENOVACION` | Enteros `yyyymmdd`. | Se convierten a fecha al importar. |
| `ULT-ANO_REN` | 2026: 22.905 · 2025: 4.865 · 2024: 3.010. | Sirve para priorizar la prospección (quien renueva está activo). |
| `DIR-COMERCIAL` | Formatos libres ("C A L L E 8 100 102 C A S A 75", "CL 1 G N 6 27"). | Se guarda el texto original y, aparte, una versión normalizada que el cliente confirma. |
| `BARRIO-COMERCIAL` / `MUN-COMERCIAL` | Formato `código - nombre` (`01112 - IPANEMA`, `41001 - NEIVA`). Solo hay 4 municipios: Neiva 21.255, Pitalito 7.289, Palermo 1.208 y Rivera 1.028. | Se separan el código DANE y el nombre del municipio. El municipio se valida contra un catálogo. |
| `TEL-COM-1` | 29.232 números de 10 dígitos, 1.496 fijos de 7 dígitos (formato antiguo), unos 50 inválidos (valor `0`, 9 u 11 dígitos). **1.421 teléfonos se repiten en 3.118 filas.** | Los fijos de 7 dígitos se normalizan con el indicativo `608`. Un teléfono repetido no identifica a una sola persona. |
| `EMAIL-COMERCIAL` | 4 correos mal formados (`@gmailcom`). **942 correos se repiten en 2.105 filas**: son contadores o tramitadores (ej. `contabilidad537@hotmail.com` aparece 15 veces). | Un correo **no identifica a una sola persona**, lo que afecta a la verificación por OTP (§6.2 y §10). |
| `CIIU-1` | Siempre presente, con 403 códigos distintos. | Se usa como respaldo cuando `ACTIVIDAD` está vacía y como filtro del CRM. |
| `ACTIVIDAD` | **Vacía en 7.754 filas (25 %)**. Tiene 17.614 textos distintos y a veces contiene lo que parece un nombre comercial ("EVENTOS VOZ A VOS."). | `TIPO_ESTABLECIMIENTO` lo escoge el cliente de un catálogo. `ACTIVIDAD_BASE` guarda el texto original. **El nombre comercial no se extrae de aquí.** |
| `TAM-EMPRESA` | Micro 29.797 · Pequeña 766 · Mediana 162 · Gran 55. | Sirve para segmentar el CRM. |

### 0.2 `Plantilla para paginas (alimentada).xlsx`

**Hoja `Plantilla original`**: es el cuestionario para construir la página web. De ella salen los campos del formulario: razón social o nombre del establecimiento, tipo de establecimiento (2 opciones), nombre de página deseado (2 opciones), qué quiere transmitir, servicios, teléfonos (fijo y WhatsApp), email, sedes con horarios especiales, horarios (entre semana, fin de semana, sitio comercial), imágenes (logo, equipo), catálogo de productos con costo, domicilios (sí/no y costo), agenda de citas (sí/no, versión de pago), testimonios (texto de hasta 30 palabras, foto y permiso de uso de la foto), base de datos de clientes (sí/no, versión de pago) y autorización para usar la razón social en nuestra base de datos.

**Hoja `Base alimentada`** (30.780 filas): es el mapeo ya aplicado. Tiene estos problemas:

1. **`NOMBRE PROPIETARIO` es igual a `RAZON SOCIAL` en las 30.780 filas**, incluidas las 7.504 que llevan la observación *"no se infirió un propietario persona natural"*. La observación y el dato se contradicen: en las jurídicas se copió el nombre de la empresa como propietario. **Esto incumple la regla de separar propietario y razón social.**
2. `NOMBRE NEGOCIO COMERCIAL` está vacío en el 100 % de las filas. Es correcto, porque no se inventó.
3. **No incluye NIT, CIIU, tamaño ni fechas.** Sin NIT no hay forma segura de identificar al cliente.
4. `ID_REGISTRO` coincide con el número de fila de `CCH` en las 30.780 filas (las diferencias en razón social son solo espacios).

**Decisión:** la importación **se hace desde `CCH@E26-9316.xlsx`**, que es la fuente con NIT. `Base alimentada` se usa solo como verificación del mapeo, enlazada por `ID_REGISTRO` = número de fila. El propietario se rellena así:

| Caso | `NOMBRE_PROPIETARIO` | `RAZON_SOCIAL` | `NOMBRE_COMERCIAL` |
|---|---|---|---|
| Persona natural (NIT que no empieza por 8/9, 10 dígitos) | = RAZON SOCIAL, marcado `SIN_CONFIRMAR` | = RAZON SOCIAL | **vacío** |
| Persona jurídica | **vacío** (lo completa el cliente como representante o propietario) | = RAZON SOCIAL | **vacío** |
| Sin NIT (13 filas) | vacío, `TIPO_PERSONA = DESCONOCIDO` | = RAZON SOCIAL | vacío |

---

## 1. Arquitectura general

```text
 ┌──────────────────────────┐   importación única + re-sync controlado
 │ CCH@E26-9316.xlsx (30.780)│──────────────────────────────┐
 └──────────────────────────┘                               ▼
                                              ┌───────────────────────────┐
                                              │ SHEETS: DB_ORIGEN         │ solo lectura, inmutable
                                              │  BASE_CCH (copia fiel)    │
                                              └────────────┬──────────────┘
                                                           │ normalización (script de importación)
                                                           ▼
 ┌──────────── APP PÚBLICA (Web App A) ────────────┐   ┌───────────────────────────┐
 │ 1 Identificación (doc) ─► 2 Verificación (OTP)  │◄─►│ SHEETS: DB_CORE           │
 │ 3 Confirmar datos ─► 4..11 Formulario por pasos │   │ CLIENTES, CONTACTOS, SEDES│
 │ 12 Autorizaciones ─► 13 Revisión ─► 14 Envío    │   │ SERVICIOS, PRODUCTOS,     │
 └───────────┬─────────────────────────────────────┘   │ TESTIMONIOS, CONTENIDO_WEB│
             │ uploads (base64 por fragmentos)         │ ARCHIVOS, AUTORIZACIONES, │
             ▼                                         │ SOLICITUDES               │
 ┌───────────────────────────┐                         └────────────┬──────────────┘
 │ DRIVE: CRM_WEB/           │◄── IDs de archivo ────────────────────┤
 │ 01_CLIENTES/<bucket>/<id>/│                                      │
 └───────────────────────────┘                                      ▼
 ┌──────────── SUPER ADMIN / CRM (Web App B) ──────┐   ┌───────────────────────────┐
 │ Clientes · Búsqueda · Duplicados · Archivos     │◄─►│ SHEETS: DB_CRM            │
 │ Oportunidades · Pipeline/Kanban · Actividades   │   │ OPORTUNIDADES, ACTIVIDADES│
 │ Seguimientos · Propuestas · Ventas · KPIs       │   │ SEGUIMIENTOS, PROPUESTAS, │
 │ Auditoría · Usuarios · Configuración            │   │ VENTAS, PROYECTOS_WEB,    │
 └─────────────────────────────────────────────────┘   │ DUPLICADOS                │
                                                       ├───────────────────────────┤
                                                       │ SHEETS: DB_SISTEMA        │
                                                       │ USUARIOS_ADMIN, CONFIG,   │
                                                       │ CATALOGOS, SECUENCIAS,    │
                                                       │ FORM_CAMPOS               │
                                                       ├───────────────────────────┤
                                                       │ SHEETS: DB_AUDITORIA_YYYY │ solo se agregan filas; rota por año
                                                       └───────────────────────────┘
```

**Flujo de negocio:**

```text
Base existente ─► Identificación ─► Formulario ─► Sheets (datos) + Drive (archivos)
   ─► Super Admin (revisión / aprobación) ─► CRM (oportunidad + actividades)
   ─► Seguimiento ─► Propuesta ─► Venta ─► Proyecto web (brief ─► construcción ─► publicación)
```

**Principios:**

1. **La base de origen no se toca nunca.** `BASE_CCH` es una copia fiel. Todo lo que el cliente corrige queda en `CLIENTES`, y la auditoría guarda el valor anterior y el nuevo.
2. **Un cliente = un registro en `CLIENTES`.** Los 30.780 se importan como `PROSPECTO`, para que el CRM trabaje sobre un solo universo. Los clientes que llegan sin estar en la base entran como `ORIGEN = AUTOREGISTRO`.
3. **Dos proyectos Apps Script separados** (público y admin) que **comparten código** gracias a un paso de build. Así se aíslan permisos, cuotas y superficie de ataque (§6).
4. **Capa de repositorio.** Ningún módulo lee celdas por posición. Todo pasa por un `Repo` que mapea las columnas por el nombre del encabezado. Esto permite migrar después a Firestore, Cloud SQL o Supabase sin reescribir la lógica.
5. **El servidor es la única fuente de validación.** El cliente HTML valida solo para mejorar la experiencia; el servidor vuelve a validar todo.

---

## 2. Modelo de datos

### 2.1 Convenciones comunes

- **IDs:** `CLIENTE_ID` es secuencial y legible (`CLI-000001`; los importados conservan `ID_REGISTRO`, así que el registro 1 es `CLI-000001`). El resto de entidades usa `PREFIJO-<base36 timestamp><4 aleatorios>` (ej. `OPO-LZ3K9A7F2Q`), que se genera sin bloqueo y no choca.
- **Columnas de control en todas las tablas:** `CREADO_EN`, `CREADO_POR`, `ACTUALIZADO_EN`, `ACTUALIZADO_POR`, `VERSION` (entero para concurrencia optimista) y `ELIMINADO` (borrado lógico: nunca se borran filas).
- **Fechas** en ISO 8601 con zona `America/Bogota`. **Dinero** en COP como entero, sin decimales.
- **Listas cortas** (redes, días) en JSON dentro de una celda **solo** si no se filtran por ellas.
- `CREADO_POR` vale `CLIENTE:<CLIENTE_ID>`, `ADMIN:<email>` o `SISTEMA:<proceso>`.

### 2.2 Entidades

#### CLIENTES (DB_CORE) — una fila por establecimiento o matrícula
| Columna | Tipo | Notas |
|---|---|---|
| CLIENTE_ID | PK | `CLI-000001` |
| ORIGEN | enum | `BASE_CCH`, `AUTOREGISTRO`, `ADMIN`, `REFERIDO` |
| ID_REGISTRO_ORIGEN | int | Fila en `BASE_CCH` (vacío si no viene de la base) |
| TIPO_PERSONA | enum | `NATURAL`, `JURIDICA`, `DESCONOCIDO` |
| TIPO_PERSONA_FUENTE | enum | `INFERIDO_NIT`, `INFERIDO_SUFIJO`, `CONFIRMADO_CLIENTE`, `ADMIN` |
| DOC_TIPO | enum | `CC`, `NIT`, `CE`, `PPT`, `PAS` |
| DOC_NUMERO | string | Sin DV, sin puntos. **Indexado** para buscar |
| DOC_DV | string(1) | Dígito de verificación |
| NOMBRE_PROPIETARIO | string | Persona natural o representante. **Vacío si no se conoce** |
| NOMBRE_PROPIETARIO_ESTADO | enum | `SIN_CONFIRMAR`, `CONFIRMADO`, `VACIO` |
| RAZON_SOCIAL | string | Nombre legal de la matrícula |
| NOMBRE_COMERCIAL | string | **Solo lo escribe el cliente o un admin; nunca se infiere** |
| TIPO_ESTABLECIMIENTO | catálogo | Elegido por el cliente (ej. Restaurante, Peluquería) |
| TIPO_ESTABLECIMIENTO_ALT | catálogo | "Opción 2" de la plantilla |
| ACTIVIDAD_BASE | string | Texto original de `ACTIVIDAD` |
| CIIU_CODIGO / CIIU_DESCRIPCION | string | `G4721` / descripción |
| TAMANO_EMPRESA | enum | MICRO, PEQUEÑA, MEDIANA, GRAN |
| TELEFONO_PRINCIPAL | string | Normalizado E.164 sin `+` (`573208481672`, `576088620841`) |
| TELEFONO_ORIGINAL | string | Tal cual viene de la base |
| EMAIL_PRINCIPAL | string | En minúsculas |
| EMAIL_COMPARTIDO | bool | `true` si el email aparece en más de un cliente (se calcula al importar) |
| MUNICIPIO_COD / MUNICIPIO | string | `41001` / `NEIVA` |
| DEPARTAMENTO | string | `HUILA` |
| DIRECCION / DIRECCION_ORIGINAL | string | Confirmada / original |
| BARRIO_COD / BARRIO | string | |
| FEC_MATRICULA / FEC_RENOVACION / ULT_ANO_RENOVACION | date/int | |
| ESTADO_COMERCIAL | enum | Ver §5.1 |
| ESTADO_FORMULARIO | enum | Ver §5.1 (derivado de SOLICITUDES) |
| DATOS_CONFIRMADOS_EN | datetime | Cuándo el cliente confirmó sus datos |
| RESPONSABLE_EMAIL | string | Admin o vendedor asignado |
| PRIORIDAD | enum | ALTA, MEDIA, BAJA |
| ETIQUETAS | string | Separadas por coma |
| FUSIONADO_EN | FK CLIENTE_ID | Si se fusionó por duplicado |
| CLAVE_BUSQUEDA | string | Razón social, nombre comercial y documento normalizados (sin tildes ni signos, en mayúsculas) |
| DRIVE_FOLDER_ID | string | Se crea la primera vez que hay una subida |
| NO_CONTACTAR | bool | Oposición del titular (habeas data) |
| + columnas de control | | |

> Tamaño: unas 45 columnas × 31.000 filas ≈ 1,4 M celdas. El límite de una hoja de Google es 10 M, así que cabe.

#### CONTACTOS (DB_CORE)
`CONTACTO_ID, CLIENTE_ID, NOMBRE, CARGO, ROL (PROPIETARIO|REPRESENTANTE|ADMINISTRADOR|CONTADOR|OTRO), TELEFONO, WHATSAPP, TELEFONO_FIJO, EMAIL, ES_PRINCIPAL, PREFERENCIA_CANAL (WHATSAPP|LLAMADA|EMAIL), HORARIO_CONTACTO, AUTORIZA_CONTACTO, + control`

#### SEDES (DB_CORE)
`SEDE_ID, CLIENTE_ID, NOMBRE_SEDE, ES_PRINCIPAL, DIRECCION, BARRIO, MUNICIPIO_COD, MUNICIPIO, REFERENCIA, LAT, LNG, GOOGLE_MAPS_URL, TELEFONO, WHATSAPP, HORARIO_SEMANA, HORARIO_FIN_SEMANA, HORARIO_FESTIVOS, HORARIOS_ESPECIALES, ATIENDE_EN_SITIO (bool, "Sitio comercial"), ORDEN, + control`

La sede principal se crea precargada con la dirección de la base.

#### SERVICIOS (DB_CORE)
`SERVICIO_ID, CLIENTE_ID, NOMBRE, DESCRIPCION, PRECIO_DESDE, PRECIO_HASTA, MOSTRAR_PRECIO, DURACION, REQUIERE_CITA, IMAGEN_ARCHIVO_ID, SEDES_IDS, ORDEN, ACTIVO, + control`

#### PRODUCTOS (DB_CORE)
`PRODUCTO_ID, CLIENTE_ID, NOMBRE, DESCRIPCION, CATEGORIA, PRECIO, MOSTRAR_PRECIO, UNIDAD, SKU, DISPONIBLE, IMAGEN_ARCHIVO_ID, ORDEN, + control`

#### TESTIMONIOS (DB_CORE)
`TESTIMONIO_ID, CLIENTE_ID, AUTOR_NOMBRE, AUTOR_CARGO_EMPRESA, TEXTO (≤30 palabras, validado en el servidor), CALIFICACION (1-5), FOTO_ARCHIVO_ID, PERMISO_USO_FOTO (bool), PERMISO_USO_NOMBRE (bool), DECLARA_CONSENTIMIENTO_TERCERO (bool), FECHA, APROBADO_ADMIN, + control`

#### CONTENIDO_WEB (DB_CORE) — respuestas de la plantilla, flexibles
Hay campos de la plantilla que cambiarán con el tiempo (opciones, versiones de pago). Por eso se guardan como **clave-valor versionado** en lugar de columnas fijas:

`CONTENIDO_ID, CLIENTE_ID, SOLICITUD_ID, SECCION, CAMPO_CLAVE, VALOR, VALOR_TIPO (texto|bool|numero|json), + control`

Claves iniciales (definidas en `FORM_CAMPOS`): `nombre_pagina_opcion_1`, `nombre_pagina_opcion_2`, `mensaje_a_transmitir`, `dominio_actual`, `redes_sociales`, `colores_marca`, `estilo_deseado`, `publico_objetivo`, `diferenciales`, `domicilios_ofrece`, `domicilios_costo`, `agenda_citas_desea`, `agenda_citas_version_pago`, `base_datos_clientes_desea`, `base_datos_clientes_version_pago`, `sitios_referencia`.

#### ARCHIVOS (DB_CORE)
`ARCHIVO_ID, CLIENTE_ID, ENTIDAD (CLIENTE|SEDE|SERVICIO|PRODUCTO|TESTIMONIO|PROPUESTA|VENTA|PROYECTO), ENTIDAD_ID, CATEGORIA (LOGO|FOTO_NEGOCIO|FOTO_EQUIPO|FOTO_PRODUCTO|FOTO_SERVICIO|FOTO_TESTIMONIO|DOCUMENTO|PROPUESTA|CONTRATO|COMPROBANTE_PAGO|OTRO), NOMBRE_ORIGINAL, NOMBRE_DRIVE, DRIVE_FILE_ID, DRIVE_FOLDER_ID, MIME_TYPE, MIME_DETECTADO, TAMANO_BYTES, SHA256, ESTADO (CUARENTENA|APROBADO|RECHAZADO|PAPELERA), MOTIVO_RECHAZO, SUBIDO_POR, + control`

#### AUTORIZACIONES (DB_CORE) — prueba del consentimiento
`AUTORIZACION_ID, CLIENTE_ID, SOLICITUD_ID, TIPO (TRATAMIENTO_DATOS|USO_IMAGENES|PUBLICACION_CONTENIDO|CONTACTO_COMERCIAL_WHATSAPP|CONTACTO_COMERCIAL_EMAIL|CONTACTO_COMERCIAL_LLAMADA|USO_RAZON_SOCIAL_BASE|TERMINOS_SERVICIO|USO_FOTOS_TESTIMONIOS), OTORGADA (bool), TEXTO_VERSION (ej. v2026-09), TEXTO_HASH (SHA-256 del texto mostrado), FECHA_HORA, FIRMANTE_NOMBRE, FIRMANTE_DOC, CANAL (WEB_APP), SESION_ID, USER_AGENT, REVOCADA_EN, + control`

> Apps Script **no entrega la IP del cliente**. La prueba del consentimiento se apoya en la sesión verificada por OTP, el hash del texto aceptado, la hora y el user-agent.

#### SOLICITUDES (DB_CORE) — ciclo de vida del formulario *(entidad añadida)*
`SOLICITUD_ID, CLIENTE_ID, ESTADO (BORRADOR|ENVIADA|EN_REVISION|REQUIERE_CORRECCION|APROBADA|ANULADA), PASO_ACTUAL, PORCENTAJE, RADICADO (WEB-2026-000123), ENVIADA_EN, SNAPSHOT_JSON_FILE_ID (JSON congelado en Drive al enviar), OBSERVACIONES_ADMIN, REVISADO_POR, REVISADO_EN, + control`

#### VERIFICACIONES (DB_SISTEMA) — OTP e intentos *(entidad añadida)*
`VERIF_ID, CLIENTE_ID, CANAL (EMAIL|WHATSAPP|MANUAL), DESTINO_ENMASCARADO, CODIGO_HASH, EXPIRA_EN, INTENTOS, ESTADO, CREADO_EN`

El código **nunca se guarda en claro**. Las sesiones activas viven en `CacheService` (§6.2).

#### OPORTUNIDADES (DB_CRM)
`OPORTUNIDAD_ID, CLIENTE_ID, TITULO, PRODUCTO_OFERTADO (WEB_BASICA|WEB_CATALOGO|WEB_AGENDA|TIENDA|ADDON_BD_CLIENTES|...), ETAPA, PROBABILIDAD, VALOR_ESTIMADO, FECHA_CIERRE_ESTIMADA, ORIGEN_LEAD, RESPONSABLE_EMAIL, MOTIVO_PERDIDA, FECHA_CIERRE_REAL, ORDEN_KANBAN, + control`

#### ACTIVIDADES (DB_CRM) — llamadas, WhatsApp, emails y reuniones en una sola tabla
`ACTIVIDAD_ID, CLIENTE_ID, OPORTUNIDAD_ID, TIPO (LLAMADA|WHATSAPP|EMAIL|REUNION|NOTA|VISITA|CAMBIO_ETAPA), DIRECCION (ENTRANTE|SALIENTE), RESULTADO (CONTESTO|NO_CONTESTO|BUZON|NUMERO_ERRADO|INTERESADO|NO_INTERESADO|VOLVER_A_LLAMAR|ENVIADO|RESPONDIDO|REALIZADA|CANCELADA|REPROGRAMADA), ASUNTO, DETALLE, DURACION_MIN, FECHA_HORA, CONTACTO_ID, ADJUNTO_ARCHIVO_ID, MEETING_URL, CALENDAR_EVENT_ID, REALIZADA_POR, + control`

#### SEGUIMIENTOS (DB_CRM) — tareas con vencimiento
`SEGUIMIENTO_ID, CLIENTE_ID, OPORTUNIDAD_ID, ACTIVIDAD_ORIGEN_ID, TIPO (LLAMAR|WHATSAPP|EMAIL|REUNION|ENVIAR_PROPUESTA|REVISAR_FORMULARIO|COBRAR|OTRO), DESCRIPCION, FECHA_VENCIMIENTO, PRIORIDAD, ESTADO (PENDIENTE|COMPLETADO|VENCIDO|CANCELADO), ASIGNADO_A, COMPLETADO_EN, ACTIVIDAD_RESULTADO_ID, + control`

#### PROPUESTAS (DB_CRM) *(entidad añadida, necesaria para "administrar propuestas")*
`PROPUESTA_ID, OPORTUNIDAD_ID, CLIENTE_ID, NUMERO (PRO-2026-0001), VERSION_PROPUESTA, ITEMS_JSON, SUBTOTAL, DESCUENTO, IMPUESTOS, TOTAL, VALIDEZ_HASTA, ESTADO (BORRADOR|ENVIADA|VISTA|ACEPTADA|RECHAZADA|VENCIDA), PDF_ARCHIVO_ID, ENVIADA_EN, RESPUESTA_EN, + control`

#### VENTAS (DB_CRM)
`VENTA_ID, OPORTUNIDAD_ID, PROPUESTA_ID, CLIENTE_ID, NUMERO (VEN-2026-0001), FECHA_VENTA, VALOR_TOTAL, FORMA_PAGO, ESTADO_PAGO (PENDIENTE|PARCIAL|PAGADA|ANULADA), VALOR_PAGADO, FACTURA_NUMERO, COMPROBANTE_ARCHIVO_ID, VENDEDOR_EMAIL, COMISION, PROYECTO_ID, + control`

> Los pagos individuales (`PAGOS`) quedan para la fase de escalabilidad (§9).

#### PROYECTOS_WEB (DB_CRM) *(se deja creada, se usa después)*
`PROYECTO_ID, VENTA_ID, CLIENTE_ID, ESTADO (BRIEF|DISENO|DESARROLLO|REVISION_CLIENTE|PUBLICADO|MANTENIMIENTO), BRIEF_ARCHIVO_ID, DOMINIO, URL_PUBLICADA, FECHA_ENTREGA_ESTIMADA, RESPONSABLE_EMAIL, + control`

#### DUPLICADOS (DB_CRM) *(entidad añadida)*
`DUP_ID, CLIENTE_A, CLIENTE_B, REGLA (MISMO_DOC|MISMO_EMAIL_Y_TEL|MISMO_TEL|NOMBRE_SIMILAR_MISMO_MUN|MISMO_EMAIL), PUNTAJE (0-100), ESTADO (PENDIENTE|ES_DUPLICADO|NO_ES_DUPLICADO|FUSIONADO), RESUELTO_POR, RESUELTO_EN`

#### USUARIOS_ADMIN (DB_SISTEMA)
`USUARIO_ID, EMAIL (cuenta Google), NOMBRE, ROL (SUPER_ADMIN|ADMIN|VENDEDOR|LECTURA), ACTIVO, MUNICIPIOS_ASIGNADOS, ULTIMO_ACCESO, + control`

#### CONFIGURACION (DB_SISTEMA) — clave-valor
`CLAVE, VALOR, TIPO, DESCRIPCION, ACTUALIZADO_EN, ACTUALIZADO_POR`

Claves previstas: `ROOT_FOLDER_ID`, `DB_*_ID`, `MAX_UPLOAD_MB` (10), `MAX_TOTAL_MB_CLIENTE` (150), `OTP_TTL_MIN` (10), `OTP_MAX_INTENTOS` (5), `SESION_TTL_MIN` (60), `TEXTO_AUTORIZACION_VERSION`, `ETAPAS_PIPELINE_JSON`, `EMAIL_REMITENTE_NOMBRE`, `MANTENIMIENTO` (bool).

#### CATALOGOS, SECUENCIAS, FORM_CAMPOS (DB_SISTEMA)
- `CATALOGOS`: `CATALOGO, CODIGO, ETIQUETA, ORDEN, ACTIVO`. Incluye municipios, tipos de establecimiento, etapas, motivos de pérdida y resultados de actividad.
- `SECUENCIAS`: `NOMBRE, ULTIMO_VALOR`. Se usa con `LockService` para radicados y números de propuesta o venta.
- `FORM_CAMPOS`: `PASO, CAMPO_CLAVE, ETIQUETA, TIPO, OBLIGATORIO, OPCIONES_CATALOGO, AYUDA, VALIDACION, ORDEN, ACTIVO`. Permite cambiar el cuestionario sin desplegar código.

#### AUDITORIA (DB_AUDITORIA_YYYY) — solo se agregan filas
`AUDIT_ID, FECHA_HORA, ACTOR (CLIENTE:CLI-…|ADMIN:email|SISTEMA:…), APP (PUBLICA|ADMIN|TRIGGER), ACCION (LOGIN|OTP_ENVIADO|OTP_FALLIDO|CREAR|ACTUALIZAR|ELIMINAR|FUSIONAR|EXPORTAR|DESCARGAR_ARCHIVO|CAMBIO_ETAPA|CAMBIO_PERMISOS|...), ENTIDAD, ENTIDAD_ID, CAMPOS_CAMBIADOS, VALOR_ANTERIOR_JSON, VALOR_NUEVO_JSON, SESION_ID, RESULTADO (OK|DENEGADO|ERROR), DETALLE`

Los datos personales en `VALOR_*` se guardan completos (el propósito es la trazabilidad). Solo pueden leerlos los `SUPER_ADMIN`.

### 2.3 Relaciones

```text
CLIENTES 1─N CONTACTOS | SEDES | SERVICIOS | PRODUCTOS | TESTIMONIOS | CONTENIDO_WEB | ARCHIVOS | AUTORIZACIONES | SOLICITUDES
CLIENTES 1─N OPORTUNIDADES 1─N ACTIVIDADES / SEGUIMIENTOS / PROPUESTAS
OPORTUNIDADES 1─0..1 VENTAS 1─0..1 PROYECTOS_WEB
CLIENTES N─N CLIENTES (DUPLICADOS)
```

### 2.4 Distribución en hojas de cálculo

| Spreadsheet | Por qué está separado |
|---|---|
| `DB_ORIGEN` | Inmutable. Solo la lee la importación y el CRM. Se puede cambiar la base sin tocar nada más. |
| `DB_CORE` | Lo que escribe la app pública. |
| `DB_CRM` | Lo que escriben los admins. Separarlo reduce la contención de bloqueos y el tamaño de cada lectura. |
| `DB_SISTEMA` | Configuración y usuarios. Es pequeño y se cachea. |
| `DB_AUDITORIA_YYYY` | Crece sin límite, así que rota por año y no afecta al rendimiento del resto. |

---

## 3. Estructura de Google Drive

```text
CRM_WEB/                                   (propietario: cuenta dueña del sistema; NO compartido públicamente)
├── 00_SISTEMA/
│   ├── DB_ORIGEN            (Sheets)
│   ├── DB_CORE              (Sheets)
│   ├── DB_CRM               (Sheets)
│   ├── DB_SISTEMA           (Sheets)
│   ├── AUDITORIA/DB_AUDITORIA_2026 …
│   ├── INDICES/             (índice de búsqueda JSON gzip, generado)
│   └── BACKUPS/YYYY-MM-DD/  (copias nocturnas de las DB, se conservan 30)
├── 01_CLIENTES/
│   └── CLI-000xxx/                        ← bucket de 1.000 clientes (evita carpetas con 30k hijos)
│       └── CLI-000123__PAPELERIA-CARTAGENA/   (slug de razón social, solo informativo; la referencia real es el FOLDER_ID)
│           ├── 01_LOGO/
│           ├── 02_FOTOS_NEGOCIO/
│           ├── 03_FOTOS_EQUIPO/
│           ├── 04_PRODUCTOS/
│           ├── 05_SERVICIOS/
│           ├── 06_SEDES/
│           ├── 07_TESTIMONIOS/
│           ├── 08_DOCUMENTOS/
│           ├── 09_SOLICITUDES/            (snapshot JSON/PDF de cada envío)
│           ├── 10_PROPUESTAS/
│           ├── 11_VENTAS/                 (comprobantes, contratos)
│           └── 12_PROYECTO_WEB/           (brief, entregables)
├── 02_CUARENTENA/YYYY-MM-DD/              (lo que se sube entra aquí primero; se valida y se mueve)
├── 03_PLANTILLAS/                         (plantilla de propuesta, textos legales versionados)
└── 04_EXPORTS/                            (exportaciones CSV de admins; se borran a los 7 días)
```

**Reglas:**

- **Las carpetas se crean de forma perezosa**, cuando un cliente sube su primer archivo, y no para los 30.780 clientes al importar.
- El nombre del archivo en Drive es `<CATEGORIA>_<ENTIDAD_ID>_<yyyyMMddHHmmss>.<ext>`. **El nombre original nunca se usa como ruta**; solo se guarda en `ARCHIVOS`.
- No se usa `setSharing(ANYONE…)` en ningún caso. El admin ve las imágenes a través del servidor (miniatura en base64 o `getThumbnailLink` con su propia sesión), no con enlaces públicos.
- La papelera es lógica: `ESTADO = PAPELERA` en `ARCHIVOS`. Un trigger mensual envía a la papelera de Drive los archivos con más de 90 días.

---

## 4. Flujo del formulario (app pública)

Es una SPA en `HtmlService` con pasos, **guardado automático por paso** en el servidor y reanudación desde cualquier dispositivo.

| # | Paso | Contenido | Reglas clave |
|---|---|---|---|
| 0 | Bienvenida | Qué es, cuánto tarda, aviso de privacidad corto y enlace a la política completa | Debe aceptar el aviso para continuar (se registra como `TRATAMIENTO_DATOS` preliminar) |
| 1 | Identificación | Tipo de documento + número (sin DV). Campo *honeypot*. | Límite de intentos por documento y global (§6.3). **La respuesta es siempre genérica**, nunca revela si el documento existe. |
| 2 | Verificación | Envío del OTP de 6 dígitos al email de la base, mostrado enmascarado (`ge*****@hotmail.com`). Alternativa: *"Ya no tengo acceso a ese correo"* → se pide un nuevo contacto y se crea un seguimiento `VERIFICACION_MANUAL` para el admin. | TTL de 10 minutos, 5 intentos. Si hay varias matrículas con el mismo documento, **después** de verificar se muestran enmascaradas para elegir una. |
| 2b | No encontrado | Registro nuevo: datos básicos + email, que se verifica con OTP | `ORIGEN = AUTOREGISTRO`, `ESTADO_FORMULARIO = IDENTIFICADO`, se marca para revisión |
| 3 | Confirmar datos | Muestra razón social, teléfono, email, municipio, dirección y barrio. Para cada uno: *correcto* o *corregir*. | Las correcciones se guardan en `CLIENTES` con auditoría del valor anterior. Se registra `DATOS_CONFIRMADOS_EN`. |
| 4 | Identidad del negocio | Tipo de persona (se confirma), propietario o representante, razón social (solo lectura si viene de la base), **nombre comercial (opcional, vacío por defecto)**, tipo de establecimiento (opciones 1 y 2), nombre de página (opciones 1 y 2) | El nombre comercial nunca se autocompleta |
| 5 | Contacto | Teléfono fijo, WhatsApp, email, redes, canal y horario preferido | Crea o actualiza `CONTACTOS` |
| 6 | Mensaje y marca | Qué quiere transmitir, público objetivo, diferenciales, colores, estilo, sitios de referencia, dominio actual | → `CONTENIDO_WEB` |
| 7 | Archivos | Logo, fotos del negocio, fotos del equipo | Validación en cliente y en servidor, barra de progreso, miniaturas (§6.4) |
| 8 | Servicios | Lista repetible: nombre, descripción, precio, imagen | Máximo configurable (ej. 30) |
| 9 | Productos | Lista repetible: nombre, precio, categoría, imagen. Domicilios (sí/no + costo). | Máximo configurable (ej. 100) |
| 10 | Sedes y horarios | La sede principal viene precargada. Se pueden añadir sedes. Horario entre semana, fin de semana y especiales; indicar si atiende en sitio. | |
| 11 | Testimonios | Lista repetible: autor, texto (≤30 palabras), foto, permiso de uso de foto y nombre | La foto exige marcar el permiso del tercero |
| 12 | Servicios adicionales | Agenda de citas (sí/no, versión de pago), base de datos de clientes (sí/no, versión de pago) | Genera señales de venta para el CRM |
| 13 | Autorizaciones | Casillas independientes, **ninguna marcada por defecto**: tratamiento de datos (obligatoria), uso de imágenes, publicación de contenido, contacto por WhatsApp, email y llamada, uso de la razón social en nuestra base | Se guarda la versión y el hash del texto |
| 14 | Revisión | Resumen de solo lectura con enlaces para editar cada paso y lista de faltantes | |
| 15 | Envío | Se congela el snapshot JSON en Drive, se asigna el radicado `WEB-2026-000123` y se envía email de confirmación | `SOLICITUD = ENVIADA`, `ESTADO_FORMULARIO = ENVIADO`. Se crea automáticamente una `OPORTUNIDAD` en etapa `FORMULARIO_RECIBIDO` y un `SEGUIMIENTO` de revisión para el responsable. |

**Reanudación:** si el cliente cierra la ventana, vuelve a identificarse con OTP y continúa en `PASO_ACTUAL`. Después del envío, las ediciones solo son posibles si el admin pone la solicitud en `REQUIERE_CORRECCION`.

---

## 5. Flujo del CRM

### 5.1 Estados (tres ejes independientes para no mezclar conceptos)

**ESTADO_FORMULARIO** (lo que ha hecho el cliente en la app):

`NO_INICIADO → IDENTIFICADO → EN_PROGRESO → ENVIADO → EN_REVISION → (REQUIERE_CORRECCION ↺) → APROBADO`

**ESTADO_COMERCIAL** (relación con el cliente):

`PROSPECTO → CONTACTADO → INTERESADO → CLIENTE → (INACTIVO)`

Además: `NO_INTERESADO`, `NO_CONTACTAR` (bloquea toda actividad saliente) y `DATOS_ERRADOS`.

**ETAPA de OPORTUNIDAD** (pipeline / Kanban), configurable en `CONFIGURACION.ETAPAS_PIPELINE_JSON`:

| Etapa | Prob. | Entrada típica |
|---|---|---|
| NUEVA | 5 % | Creada por un admin desde un prospecto |
| CONTACTO_INICIAL | 10 % | Primera actividad saliente registrada |
| FORMULARIO_RECIBIDO | 30 % | Automática al enviar el formulario |
| BRIEF_VALIDADO | 45 % | El admin aprueba la solicitud |
| PROPUESTA_ENVIADA | 60 % | La propuesta pasa a `ENVIADA` |
| NEGOCIACION | 75 % | Manual |
| GANADA | 100 % | Se registra la venta (obligatoria para entrar aquí) |
| PERDIDA | 0 % | Motivo de pérdida obligatorio |

### 5.2 Reglas automáticas

- Primera actividad saliente sobre un `PROSPECTO` → `ESTADO_COMERCIAL = CONTACTADO`.
- Envío del formulario → se crea la oportunidad (o se reutiliza si hay una abierta) en `FORMULARIO_RECIBIDO` y un seguimiento para revisar en 24 h.
- Oportunidad `GANADA` → se exige una `VENTA`; el cliente pasa a `CLIENTE`. Cuando la venta queda `PAGADA` (o con el anticipo configurado) se crea `PROYECTOS_WEB` en estado `BRIEF`.
- Actividad con resultado `VOLVER_A_LLAMAR` → exige crear un seguimiento con fecha.
- Trigger diario → los seguimientos vencidos pasan a `VENCIDO` y se envía un resumen por email a cada responsable.
- Si el cliente está en `NO_CONTACTAR`, el servidor **rechaza** actividades salientes y el botón de WhatsApp aparece deshabilitado.

### 5.3 Kanban

- Una columna por etapa, con tarjetas de oportunidad (cliente, valor, responsable, días en la etapa y próximo seguimiento).
- Arrastrar una tarjeta llama a `moverOportunidad(id, etapaDestino, versionEsperada)`. Si `VERSION` no coincide, se responde con conflicto y se recarga la tarjeta. Cada movimiento queda como `ACTIVIDAD` de tipo `CAMBIO_ETAPA` y en `AUDITORIA`.
- Las transiciones a `GANADA` y `PERDIDA` abren un modal obligatorio (venta o motivo).

### 5.4 Actividades y comunicación

- **Llamada:** botón `tel:` y registro del resultado al volver.
- **WhatsApp:** botón `https://wa.me/57XXXXXXXXXX?text=<plantilla>` (se abre fuera de la app) y registro manual de la actividad. Una API de WhatsApp Business queda para §9.
- **Email:** se envía desde la app con `GmailApp` y plantillas, y se registra automáticamente. Hay cuota diaria (§10).
- **Reunión:** se crea un evento en Calendar con `CalendarApp` y se guarda `CALENDAR_EVENT_ID`.

### 5.5 Vistas del Super Admin

1. **Dashboard / KPIs:** prospectos contactados, formularios iniciados o enviados (embudo), tasa de conversión por etapa, valor del pipeline ponderado, ventas del mes, seguimientos vencidos, actividad por vendedor y distribución por municipio y CIIU.
2. **Clientes:** tabla paginada en el servidor con búsqueda (documento, nombre, email, teléfono), filtros (municipio, estados, CIIU, tamaño, año de renovación, responsable, etiquetas) y ficha 360° con pestañas (datos, contactos, sedes, catálogo, archivos, autorizaciones, oportunidades, línea de tiempo y auditoría).
3. **Duplicados:** cola de pares candidatos con comparación lado a lado y acción *fusionar* (elige el registro superviviente campo por campo; el otro queda con `FUSIONADO_EN`).
4. **Archivos:** revisión de lo que está en cuarentena o pendiente (aprobar o rechazar con motivo).
5. **Pipeline / Kanban**, **Seguimientos (mi agenda)**, **Propuestas**, **Ventas**.
6. **Auditoría:** filtros por actor, entidad y fecha (solo `SUPER_ADMIN`).
7. **Configuración:** catálogos, etapas, usuarios, textos legales, límites.

### 5.6 Detección de duplicados

Se normaliza: mayúsculas, sin tildes, sin signos, sin espacios dobles, y se eliminan sufijos societarios para comparar nombres. Luego se aplican las reglas por orden de peso:

| Regla | Puntaje | Observación |
|---|---|---|
| Mismo `DOC_NUMERO` | 100 | En la base hay 2 casos y 13 filas sin NIT |
| Mismo email **y** mismo teléfono | 85 | |
| Nombre similar (trigramas ≥ 0,85) en el mismo municipio | 70 | |
| Mismo teléfono | 40 | **Solo informativo**: 1.421 teléfonos compartidos |
| Mismo email | 30 | **Solo informativo**: 942 emails de contadores |

Se ejecuta en lote (trigger nocturno por bloques para no pasar de 6 minutos) y también en línea cuando se crea o edita un cliente. Nunca se fusiona automáticamente.

---

## 6. Autenticación y seguridad de acceso

### 6.1 Super Admin (Web App B)

**Decisión:** el proyecto admin se despliega con **Ejecutar como: usuario que accede** y **Acceso: cualquier persona con cuenta de Google**.

- Con esa configuración, `Session.getActiveUser().getEmail()` devuelve el email real y verificado por Google. Con "Ejecutar como yo" y cuentas `@gmail.com` **ese email llega vacío**, por lo que no sirve para autenticar.
- En cada llamada al servidor, `requireRole(rol)` comprueba el email contra `USUARIOS_ADMIN` (`ACTIVO = true`), con caché de 5 minutos. Si no está, devuelve `DENEGADO` y lo audita. `doGet` solo sirve la interfaz a emails autorizados; a los demás les muestra una página neutra.
- Los spreadsheets y `CRM_WEB/` se comparten **solo** con los admins (editor) y nunca con enlace. Esto es necesario porque el código corre con la identidad de quien accede.
- Los permisos se aplican por rol en el servidor, no en la interfaz:
  - `SUPER_ADMIN`: todo, incluidos auditoría, usuarios y exportación.
  - `ADMIN`: CRM completo sin usuarios.
  - `VENDEDOR`: solo clientes asignados.
  - `LECTURA`: solo consulta.
- Sin contraseñas propias: la autenticación, incluido el 2FA, la hace Google. Se recomienda **exigir la verificación en dos pasos** a toda cuenta que esté en `USUARIOS_ADMIN`.

**Limitación conocida:** un admin con acceso de editor puede abrir el Sheet directamente y saltarse la app. Para 1 a 5 admins de confianza es aceptable. Para vendedores con restricción por fila, ver §9 (Workspace + "Ejecutar como yo", o un backend propio).

### 6.2 Cliente (Web App A)

Se despliega con **Ejecutar como: yo** y **Acceso: cualquiera, incluso anónimo**. El cliente nunca toca Drive ni Sheets directamente.

- **Sesión:** después de validar el OTP, el servidor genera un token (`Utilities.getUuid()` + `computeDigest` de un aleatorio) y guarda en `CacheService` `{clienteId, solicitudId, expira}` con TTL de 60 minutos que se renueva con cada uso. El cliente lo conserva en `sessionStorage`, nunca en `localStorage`, y lo envía en **cada** `google.script.run`.
- **Todas las funciones públicas** reciben el token y derivan de él el `clienteId`. **Nunca aceptan `clienteId` como parámetro.** Esto evita IDOR (que un cliente acceda a datos de otro cambiando un ID).
- Antes de verificar, **nunca** se devuelven datos personales: solo "si el documento está registrado, enviamos un código a ge*****@hotmail.com". El enmascaramiento revela lo mínimo.
- El OTP se guarda hasheado con sal, es de un solo uso, expira en 10 minutos y admite 5 intentos. Si falla, el documento queda bloqueado 30 minutos.

### 6.3 Límite de intentos y abuso

- Por documento: 3 OTP por hora. Global: N envíos por hora (protege la cuota de email).
- Contadores guardados en `CacheService` y operaciones críticas bajo `LockService`.
- Campo honeypot más un tiempo mínimo por paso para filtrar bots básicos. Se puede añadir reCAPTCHA v3 con `UrlFetchApp`, pero su efectividad dentro del iframe de Apps Script es limitada; queda como opción.

### 6.4 Archivos subidos

- Se valida **en el servidor**: tamaño (≤10 MB por archivo, ≤150 MB por cliente), extensión permitida (`jpg`, `jpeg`, `png`, `webp`, `pdf`; SVG **no**, porque puede contener scripts), **firma mágica** de los primeros bytes (JPEG `FFD8FF`, PNG `89504E47`, PDF `25504446`, WEBP `RIFF…WEBP`) y SHA-256 para detectar duplicados.
- El archivo entra en `02_CUARENTENA`, se valida y se mueve a la carpeta del cliente. Si no pasa la validación se rechaza y se audita.
- El transporte usa base64 por fragmentos de unos 5 MB para no chocar con el límite de payload de `google.script.run`, y se reensambla en el servidor.
- Los archivos nunca se sirven con enlace público. Las previsualizaciones para el cliente las genera el servidor y solo para archivos de su propia sesión.

---

## 7. Seguridad: riesgos y medidas

| Área | Riesgo | Medida |
|---|---|---|
| Apps Script | Un web app anónimo "Ejecutar como yo" actúa con **todos** los permisos del propietario | Proyecto público separado con **scopes mínimos** declarados en `appsscript.json` (`spreadsheets`, `drive` limitado a lo necesario, `script.send_mail`). Nunca `eval` ni ejecución dinámica. Lista blanca de funciones expuestas (las privadas terminan en `_`, porque Apps Script no expone a `google.script.run` las que acaban en guion bajo). |
| Apps Script | XSS en la interfaz con datos de la base o del cliente | Escapar siempre con `textContent` o una plantilla que escape; prohibido `innerHTML` con datos. `HtmlService` con `XFrameOptionsMode.DEFAULT`. |
| Apps Script | Cuotas: 6 min por ejecución, unas 30 ejecuciones simultáneas, 90 min diarios de triggers (cuenta gratuita), emails diarios | Ver §10. |
| Sheets | Inyección de fórmulas: un cliente escribe `=IMPORTXML(...)` o `=HYPERLINK(...)` | **Todo valor de texto que empiece por `= + - @` se guarda con un apóstrofo delante** y las exportaciones CSV también se sanean. |
| Sheets | Escrituras concurrentes que pisan filas | `LockService.getScriptLock()` en inserciones y actualizaciones. Localizar filas por ID, nunca por número de fila en caché. Columna `VERSION` para concurrencia optimista. |
| Sheets | Que alguien comparta el archivo por error | Revisión semanal automática de permisos (`DriveApp.getAccess`) con alerta al super admin si aparece `ANYONE` o `ANYONE_WITH_LINK`. |
| Drive | Archivos maliciosos o pesados | §6.4. |
| Drive | Enlaces públicos filtrados | Prohibido compartir públicamente, y verificado en la auditoría semanal. |
| Sesiones | Robo o reutilización del token | TTL corto, renovación con cada uso, invalidación al enviar el formulario, token de un solo contexto y nunca en la URL. |
| Sesiones | Enumeración de documentos (averiguar quién está en la base) | Respuesta genérica, límites de intentos y registro en auditoría de las búsquedas fallidas repetidas. |
| Permisos | Escalada de rol desde la interfaz | Rol comprobado en el servidor en cada llamada. |
| Datos personales (Ley 1581 de 2012, Decreto 1377 de 2013) | Contactar y almacenar datos de 30.780 titulares tomados de un registro público | Política de tratamiento publicada, autorización previa, expresa e informada con registro probatorio (tabla `AUTORIZACIONES`), mecanismo de consulta, reclamo y supresión (`NO_CONTACTAR`), finalidad limitada y minimización (no pedir cédula completa si basta el NIT). **Consulta legal recomendada** sobre el uso comercial de datos de un registro mercantil y el eventual registro en el RNBD de la SIC. |
| Datos personales | Exportaciones masivas | Solo `SUPER_ADMIN`, auditadas, con caducidad de 7 días en `04_EXPORTS`. |
| Datos de terceros | Fotos y nombres en testimonios | Declaración de consentimiento del tercero y aprobación del admin antes de publicar. |

---

## 8. Arquitectura clasp (estructura de archivos)

**Decisión:** un monorepo con **código compartido copiado en el build**, no una *Library* de Apps Script. Las librerías añaden latencia en cada llamada y complican el versionado.

```text
crm-web/
├── package.json                 # scripts: build, push:public, push:admin, deploy:*, lint, test
├── .eslintrc.json  .prettierrc
├── scripts/
│   ├── build.mjs                # copia shared/ + apps/<x>/ → dist/<x>/ (orden de carga estable)
│   └── import-base.md           # procedimiento de importación inicial
├── docs/
│   └── FASE1_ARQUITECTURA.md
├── shared/                      # lógica pura, sin estado global al cargar
│   ├── 00_Config.js             # IDs de DB/carpetas desde PropertiesService + CONFIGURACION
│   ├── 01_Utils.js              # fechas, ids, normalización (tildes, teléfonos, NIT/DV), sanitización anti-fórmulas
│   ├── 02_Repo.js               # Repo genérico: encabezados→objetos, findById, findBy (TextFinder), insert, update (versión), softDelete
│   ├── 03_Schema.js             # definición de tablas/columnas/validaciones (fuente única)
│   ├── 04_Audit.js
│   ├── 05_Validators.js         # reglas de negocio compartidas (≤30 palabras, email, teléfono, etc.)
│   ├── 06_DriveStore.js         # carpetas perezosas, cuarentena, firma mágica, hash
│   └── 07_Errors.js             # AppError con códigos; respuesta {ok, data, error}
├── apps/
│   ├── public/
│   │   ├── .clasp.json          # scriptId del proyecto PÚBLICO
│   │   ├── appsscript.json      # webapp: executeAs USER_DEPLOYING, access ANYONE_ANONYMOUS; oauthScopes mínimos
│   │   ├── server/
│   │   │   ├── Main.js          # doGet
│   │   │   ├── Api.js           # ÚNICAS funciones expuestas: api_identificar, api_verificarOtp, api_guardarPaso, api_subirFragmento, api_enviar…
│   │   │   ├── AuthCliente.js   # OTP, sesiones, rate limit
│   │   │   ├── FormService.js
│   │   │   └── UploadService.js
│   │   └── client/
│   │       ├── index.html       # shell
│   │       ├── styles.html
│   │       ├── app.js.html      # router de pasos, estado, autosave
│   │       ├── steps/*.html     # un parcial por paso
│   │       └── components/*.html
│   └── admin/
│       ├── .clasp.json          # scriptId del proyecto ADMIN
│       ├── appsscript.json      # webapp: executeAs USER_ACCESSING, access ANYONE (con cuenta Google)
│       ├── server/
│       │   ├── Main.js  Api.js  AuthAdmin.js (requireRole)
│       │   ├── ClientesService.js  DuplicadosService.js  ArchivosService.js
│       │   ├── OportunidadesService.js  ActividadesService.js  SeguimientosService.js
│       │   ├── PropuestasService.js  VentasService.js  KpiService.js  AuditService.js
│       │   ├── SearchIndex.js   # índice comprimido para búsqueda rápida
│       │   └── Triggers.js      # nocturnos: duplicados, índice, vencidos, backups, auditoría de permisos
│       ├── setup/
│       │   ├── Install.js       # crea carpetas, spreadsheets, hojas y encabezados (idempotente)
│       │   └── ImportBase.js    # importa CCH por lotes reanudables
│       └── client/ index.html, styles.html, app.js.html, views/*.html (dashboard, clientes, ficha, kanban, …)
├── tests/                       # Jest sobre shared/ (normalización, DV, validadores, reglas de duplicados)
└── dist/                        # generado (rootDir de clasp); en .gitignore
```

**Configuración sensible:** los IDs de los spreadsheets y carpetas y la sal del OTP van en `PropertiesService` (Script Properties), **nunca en el repositorio**. `.clasp.json` y `.clasprc.json` no contienen secretos, pero `.clasprc.json` (credenciales de clasp) **nunca se sube**.

**Despliegues:** se usan versiones con nombre (`clasp deploy -i <deploymentId> -d "v1.3"`) para que la URL pública no cambie.

**Importación inicial:** `ImportBase` lee `BASE_CCH` en bloques de 2.000 filas, guarda un cursor en `PropertiesService` y se reprograma con un trigger hasta terminar. Esto respeta el límite de 6 minutos. Escribe con `setValues` en lote y no fila a fila.

---

## 9. Escalabilidad

| Capacidad futura | Qué se deja preparado en la Fase 2 |
|---|---|
| **Múltiples administradores** | `USUARIOS_ADMIN` con roles, auditoría por actor y bloqueos optimistas desde el día 1. |
| **Vendedores** | `RESPONSABLE_EMAIL`, `MUNICIPIOS_ASIGNADOS` y filtro por rol en el `Repo`. **Para que un vendedor no pueda abrir el Sheet completo** hay que migrar a Google Workspace (con "Ejecutar como yo" dentro del mismo dominio sí se obtiene el email del usuario, y los datos quedan sin compartir) o mover el backend a Cloud Run/Firebase. La capa `Repo` hace ese cambio local. |
| **Propuestas** | Tabla `PROPUESTAS` + plantilla de Google Docs en `03_PLANTILLAS` → PDF con `DocumentApp` y `getAs('application/pdf')`. |
| **Pagos** | Tabla `PAGOS (PAGO_ID, VENTA_ID, VALOR, METODO, REFERENCIA_PASARELA, ESTADO, FECHA)`. Webhook de pasarela (Wompi, ePayco o Mercado Pago) hacia el `doPost` de un **tercer** proyecto con firma HMAC verificada. |
| **Gestión de proyectos** | `PROYECTOS_WEB` + `TAREAS_PROYECTO`, Kanban reutilizando el mismo componente. |
| **IA** | Módulo `AiService` en el servidor (`UrlFetchApp` hacia la API de Claude con la clave en Script Properties), con cola en `IA_TRABAJOS (JOB_ID, TIPO, ENTRADA_REF, ESTADO, SALIDA_FILE_ID, COSTO_TOKENS)` procesada por un trigger para no bloquear la interfaz. |
| **Brief automático** | Toma el snapshot JSON de la solicitud, genera un brief en Google Docs y lo guarda en `12_PROYECTO_WEB`. Lo dispara el paso a `BRIEF_VALIDADO`. |
| **Contenido web** | Con `CONTENIDO_WEB`, servicios, productos y testimonios genera textos (hero, "quiénes somos", SEO) guardados como versiones que el admin aprueba antes de publicar. |
| **Volumen** | Si `ACTIVIDADES` o `AUDITORIA` superan unas 200.000 filas, se rotan por año. Si la latencia de Sheets deja de ser aceptable, se migran las tablas calientes a Firestore o Cloud SQL detrás del mismo `Repo`. |
| **Front con marca propia** | El repositorio ya tiene `CNAME` (GitHub Pages). La API pública se diseña con respuestas JSON uniformes para que en el futuro el formulario pueda servirse desde un dominio propio y llamar a un `doPost`. |

---

## 10. Riesgos y mitigaciones

| # | Riesgo | Impacto | Mitigación |
|---|---|---|---|
| R1 | **Cuota de email en cuentas @gmail.com: unos 100 destinatarios al día con `MailApp`.** Cada OTP consume uno. | Bloquea la identificación si hay campaña masiva | Usar una cuenta de **Google Workspace** (1.500 al día), campañas escalonadas, contador de cuota con `MailApp.getRemainingDailyQuota()` y alternativa de verificación manual o WhatsApp. **Es el riesgo operativo n.º 1.** |
| R2 | **Emails compartidos (942 emails en 2.105 filas)**: el OTP llega al contador y no al propietario | Suplantación: un contador podría completar datos de varios clientes | Mostrar a qué email enmascarado se envía el código. En emails compartidos, pedir un segundo factor (últimos 4 dígitos del teléfono registrado) y marcar la solicitud para revisión. Permitir que el cliente cambie el email principal tras verificar. |
| R3 | Email desactualizado o inválido | El cliente no puede verificarse | Flujo "ya no tengo ese correo" → verificación manual por llamada. |
| R4 | `Base alimentada` copia la razón social como propietario incluso en jurídicas | Datos erróneos en el CRM y comunicaciones mal dirigidas | Importar desde `CCH`, aplicar la regla de §0.2 y pedir confirmación al cliente. |
| R5 | Orden ambiguo de nombres en personas naturales | No se puede separar nombres y apellidos | Guardar el nombre completo sin partir. Si en el futuro hace falta, el cliente lo separa en el formulario. |
| R6 | Límite de 6 minutos y lecturas lentas de Sheets con 31.000 filas | Búsquedas lentas, timeouts | Búsqueda puntual con `TextFinder` sobre la columna `DOC_NUMERO`. Para listados, un índice compacto (id, nombre, doc, municipio, estados, responsable) gzip + base64 en `CacheService` por fragmentos de 90 KB, reconstruido cada noche y marcado como obsoleto en cada escritura. Paginación en el servidor. |
| R7 | Concurrencia: unas 30 ejecuciones simultáneas por usuario, y todas las anónimas cuentan contra el propietario | Errores "demasiadas ejecuciones" en picos | Autosave con *debounce* (no en cada tecla), subidas secuenciales y campañas escalonadas. |
| R8 | Colisiones al escribir | Pérdida de datos | `LockService` con `tryLock(10000)`, `VERSION` y reintento con retroceso exponencial. |
| R9 | Crecimiento de Drive (fotos) | Cuota de almacenamiento: 15 GB en la cuenta gratuita | Límite por cliente, compresión o redimensionado en el navegador antes de subir (canvas → JPEG/WebP a 1.920 px) y Workspace o almacenamiento adicional. |
| R10 | Cumplimiento de habeas data | Sanciones de la SIC | §7, más política y aviso de privacidad redactados o revisados por un abogado antes de lanzar. |
| R11 | Duplicados de teléfono o email tratados como la misma persona | Fusiones erróneas | Reglas con puntaje y fusión solo manual (§5.6). |
| R12 | 13 filas sin NIT y 2 NIT repetidos | No se pueden identificar por documento | Marcar `DATOS_INCOMPLETOS`. Solo pueden entrar como autoregistro y el admin las reconcilia. |
| R13 | Dependencia de una sola cuenta propietaria | Si la cuenta se pierde, se pierde todo el sistema | Workspace con cuenta de servicio o administrativa dedicada (no personal), backups nocturnos y documentación del procedimiento de recuperación. |
| R14 | Cambios del cuestionario (la plantilla evolucionará) | Redespliegues frecuentes | `FORM_CAMPOS` + `CONTENIDO_WEB` clave-valor. |
| R15 | Autorizaciones marcadas por defecto o agrupadas | Consentimiento inválido | Casillas independientes, desmarcadas y con texto versionado y hash. |

---

## 11. Decisiones que requieren confirmación antes de la Fase 2

1. **Tipo de cuenta:** ¿el sistema vivirá en una cuenta **@gmail.com** o en **Google Workspace**? Esto afecta a R1 (cuota de email), a §6.1 y a §9 (vendedores). **Recomendación:** Workspace.
2. **Verificación del cliente:** ¿se acepta OTP por email como método principal con alternativa manual? ¿O se quiere también WhatsApp, que requiere un proveedor de pago?
3. **Importación de 30.780 prospectos al CRM:** ¿se importan todos (recomendado) o solo quienes se identifiquen en la app?
4. **Precios y catálogo de productos web** (básica, catálogo, agenda, base de datos de clientes, "versión de pago"): hacen falta los valores para propuestas y KPIs.
5. **Responsable legal** de la política de tratamiento de datos y del canal de atención de reclamos.

---

*Fin de la FASE 1. En espera de la instrucción para implementar.*

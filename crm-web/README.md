# Webpaya · Formulario de onboarding + Super Admin / CRM

Sistema en **Google Apps Script** (clasp, HTML5, CSS3, JavaScript, Google Sheets y Google Drive) con dos aplicaciones en un mismo proyecto:

| Aplicación | URL | Para quién |
|---|---|---|
| **Formulario público** | `…/exec` | El cliente llena la información para su página web. No necesita identificarse. |
| **Super Admin / CRM** | `…/exec?page=admin` | El equipo de Webpaya: clientes, CRM, Kanban, ventas, reportes y auditoría. |

Flujo completo:

```text
BASE DE CLIENTES (CCH, 30.780) → IDENTIFICACIÓN (búsqueda opcional) → FORMULARIO → DRIVE + SHEETS
→ SUPER ADMIN → CRM → SEGUIMIENTO → PROPUESTA → VENTA → PROYECTO WEB (+ reporte editable en Google Docs)
```

Documentos relacionados:

- [`docs/FASE1_ARQUITECTURA.md`](docs/FASE1_ARQUITECTURA.md): diseño aprobado.
- [`docs/FASE2_IMPLEMENTACION.md`](docs/FASE2_IMPLEMENTACION.md): ajustes de arquitectura hechos durante la implementación y su porqué.
- [`docs/BASE_DE_DATOS.md`](docs/BASE_DE_DATOS.md): todas las hojas y columnas (generado desde el código).
- [`docs/PRUEBAS.md`](docs/PRUEBAS.md): checklist de pruebas.

---

## 1. Estructura del proyecto

```text
crm-web/
├── .clasp.json.example        Plantilla de configuración de clasp (copie a .clasp.json)
├── .claspignore               Solo se suben a Apps Script los .gs, .html y appsscript.json
├── package.json               Scripts: test, test:ui, dev, push, pull, open, deploy
├── README.md
├── docs/
│   ├── FASE1_ARQUITECTURA.md
│   ├── FASE2_IMPLEMENTACION.md
│   ├── BASE_DE_DATOS.md
│   └── PRUEBAS.md
├── src/                       ← rootDir de clasp (lo que se sube a Apps Script)
│   ├── appsscript.json        Manifiesto: zona horaria, V8, Web App, permisos (scopes)
│   ├── Code.gs                doGet, include_, respuesta estándar, funciones del propietario y activadores
│   ├── Config.gs              CONFIG central, ajustes editables, textos de autorización, Script Properties
│   ├── Utils.gs               Utilidades puras: fechas, IDs, hash, normalización, teléfonos 608, dinero, similitud
│   ├── Schema.gs              Modelo de datos (fuente única): tablas, columnas, listas, etapas del CRM
│   ├── SheetService.gs        Repositorio sobre Sheets, bloqueo (LockService), secuencias, contexto
│   ├── AuditService.gs        Auditoría inmutable
│   ├── AuthService.gs         Login de administradores, sesiones, roles, bloqueo, 2FA, token del formulario
│   ├── ValidationService.gs   Validación y limpieza en el servidor
│   ├── SearchIndex.gs         Índice comprimido de clientes (búsqueda pública y del CRM)
│   ├── DriveService.gs        Carpetas por cliente, validación de archivos (extensión, MIME real, tamaño)
│   ├── ClientService.gs       Clientes, datos públicos, archivos, duplicados y fusiones
│   ├── FormService.gs         Formulario público: borrador, búsqueda, precarga, cargas y envío
│   ├── CRMService.gs          Oportunidades, Kanban, actividades, seguimientos, bandeja de solicitudes
│   ├── SalesService.gs        Propuestas, ventas, pagos y proyectos web
│   ├── ReportService.gs       Reporte editable por cliente (Google Doc + JSON), versionado
│   ├── DashboardService.gs    Indicadores
│   ├── AdminService.gs        Catálogos, productos web, usuarios, ajustes, exportación CSV
│   ├── ImportService.gs       Importación de la base CCH (reglas D1–D9), por lotes reanudables
│   ├── SetupService.gs        Instalación idempotente, datos iniciales, activadores, tareas diarias
│   ├── PublicApi.gs           FUNCIONES PÚBLICAS  (pub_*)
│   ├── AdminApi.gs            FUNCIONES ADMINISTRATIVAS (adm_*) con sesión + permiso
│   ├── Index.html             Formulario público (estructura)
│   ├── Styles.html            Estilos del formulario
│   ├── Scripts.html           Lógica del formulario
│   ├── Admin.html             Super Admin (estructura)
│   ├── AdminStyles.html       Estilos del Super Admin
│   └── AdminScripts.html      Lógica del Super Admin
└── tests/
    ├── harness/gas-mocks.js   Simulación en memoria de SpreadsheetApp, DriveApp, CacheService…
    ├── harness/load-gas.js    Carga src/*.gs como lo hace Apps Script
    ├── fixtures.js            Datos 100 % sintéticos
    ├── run-tests.js           56 pruebas del servidor (público, importación, CRM, seguridad, contrato)
    ├── ui-tests.js            5 recorridos de interfaz con Playwright (escritorio, celular, sin conexión)
    ├── dev-server.js          Servidor local para ver y probar la app sin desplegar
    └── gen-db-doc.js          Genera docs/BASE_DE_DATOS.md desde Schema.gs
```

El código completo de cada archivo está en `src/`. Ningún archivo tiene partes omitidas.

---

## 2. Configuración inicial

### 2.1 Qué crea `setup()` automáticamente

No hay que crear hojas ni carpetas a mano. `setup()` (ejecutado una vez desde el editor) crea:

| Recurso | Nombre | Dónde queda el ID |
|---|---|---|
| Carpeta de sistema | `CRM_WEB_SISTEMA` | Script Property `SYSTEM_FOLDER_ID` |
| Carpeta de clientes | `CLIENTES_WEB` | `ROOT_FOLDER_ID` |
| Base de datos | `CRM_WEB_DB` (24 hojas) | `DB_SPREADSHEET_ID` |
| Auditoría | `CRM_WEB_AUDITORIA` | `AUDIT_SPREADSHEET_ID` |
| Secreto para contraseñas | — | `AUTH_PEPPER` (aleatorio; **no lo cambie** o dejarán de funcionar las contraseñas) |
| Catálogos, productos web y ajustes | Filas iniciales | Hojas `CATALOGOS`, `PRODUCTOS_WEB`, `CONFIGURACION` |
| Activadores | Reportes cada 10 min, tareas diarias 6 a.m. | Proyecto |
| Código para el primer Super Admin | Se muestra en el registro de ejecución | `SETUP_CODE_HASH` (solo el hash) |

Es **idempotente**: si lo ejecuta otra vez, solo crea lo que falte y agrega columnas nuevas al final de las hojas.

### 2.2 Valores que configura el administrador

| Dónde | Clave | Obligatorio | Valor |
|---|---|---|---|
| Script Properties | `IMPORT_SOURCE_SPREADSHEET_ID` | Para importar la base | ID de la hoja de Google con `CCH@E26-9316` (ver §6) |
| Script Properties | `ROOT_FOLDER_ID` | No | Solo si quiere usar una carpeta `CLIENTES_WEB` existente: defínalo **antes** de `setup()` |
| Panel › Configuración | `APP_NAME`, `RESPONSABLE_TRATAMIENTO` (= Equipo Webpaya), `CONTACTO_PRIVACIDAD` | Recomendado | Textos del formulario y del aviso de privacidad |
| Panel › Configuración | `MAX_FILE_SIZE_MB` | No (10) | Máximo 20 |
| Panel › Configuración | `REQUIRED_FIELDS` | No | Por defecto: `razonSocial,medioContacto,TRATAMIENTO_DATOS` |
| Panel › Configuración | `DEFAULT_ASSIGNEE` | Recomendado | Email del usuario que recibe los formularios nuevos |
| Panel › Configuración | `REPORT_EDITORS` | Recomendado | Cuentas Google que podrán **editar** los reportes (Docs) |
| Panel › Configuración | `ADMIN_2FA` | Recomendado | Sí = código por email al iniciar sesión |
| Panel › Configuración | `WHATSAPP_PLANTILLA`, `PUBLIC_FORM_URL` | No | Mensaje de invitación por WhatsApp |
| Panel › Configuración › Productos web | Precios sugeridos | No | Editables; se precargan en propuestas (D14) |

Las constantes técnicas (límites, zona horaria, municipios) están en `CONFIG` de [`src/Config.gs`](src/Config.gs).

### 2.3 Permisos

- La Web App se ejecuta **como el propietario** (`executeAs: USER_DEPLOYING`) y es accesible por **cualquiera** (`ANYONE_ANONYMOUS`), porque el formulario es público.
- Por eso las hojas y carpetas **no se comparten con nadie**: los administradores trabajan solo a través del panel, que controla sesión y rol en cada llamada. La tarea diaria avisa en la auditoría si algo quedó compartido públicamente.
- Para editar los reportes en Google Docs, agregue las cuentas en `REPORT_EDITORS`: se comparte solo cada Doc, no las hojas.
- Scopes del manifiesto: Sheets, Drive, Docs, activadores, envío de email (solo 2FA de administradores) y email del usuario (solo para reconocer al propietario en funciones de mantenimiento).

---

## 3. Instalación con clasp

Requisitos: Node.js 18+ y una cuenta Google (la dueña del sistema, idealmente dedicada a Webpaya).

```bash
# 1. Instalar clasp
npm install -g @google/clasp

# 2. Habilitar la API de Apps Script (una sola vez):
#    https://script.google.com/home/usersettings  →  "Google Apps Script API" = Activado

# 3. Iniciar sesión con la cuenta dueña
clasp login

# 4. Crear el proyecto (elija UNA opción)
cd crm-web

#    Opción A (recomendada): crear un proyecto vacío en https://script.google.com → "Nuevo proyecto",
#    copiar el ID en Configuración del proyecto → "ID de la secuencia de comandos", y:
cp .clasp.json.example .clasp.json        # y pegue el scriptId en .clasp.json

#    Opción B: crearlo desde la terminal
clasp create --type webapp --title "Webpaya CRM" --rootDir src
git checkout -- src/appsscript.json       # clasp create reemplaza el manifiesto: recupere el del proyecto
#    (confirme que .clasp.json tenga "rootDir": "src")

# 5. Subir el código
clasp push -f                             # -f acepta reemplazar el manifiesto remoto

# 6. Abrir el editor
clasp open                                # en clasp 3.x: clasp open-script
```

Otros comandos:

| Comando | Qué hace |
|---|---|
| `clasp push` | Sube `src/` a Apps Script (reemplaza los archivos remotos). |
| `clasp pull` | Descarga la versión remota a `src/` (**sobrescribe** los cambios locales: haga commit antes). |
| `clasp open` | Abre el proyecto en el editor web (`clasp open-script` en clasp 3.x). |
| `clasp deployments` | Lista las implementaciones y sus IDs. |
| `clasp logs` | Muestra los registros de ejecución (requiere proyecto de Google Cloud vinculado). |

---

## 4. Despliegue como Web App

1. En el editor (`clasp open`), elija la función **`setup`** y pulse **Ejecutar**. Autorice los permisos cuando los pida.
2. Abra **Registro de ejecución**: verá las URLs creadas y el **código de configuración** (formato `XXXX-XXXX-XXXX`, válido 24 h).
3. **Implementar › Nueva implementación** → tipo **Aplicación web**:
   - Ejecutar como: **Yo (cuenta dueña)**
   - Quién tiene acceso: **Cualquier usuario**
   - Pulse **Implementar** y copie la URL `…/exec`.

   También puede hacerlo por terminal (usa la configuración del manifiesto): `clasp deploy --description "v1"`.
4. Formulario público: `https://script.google.com/macros/s/…/exec`
   Super Admin: `https://script.google.com/macros/s/…/exec?page=admin`

**Actualizar sin cambiar la URL:** después de `clasp push`, use **Implementar › Gestionar implementaciones › Editar › Nueva versión**, o por terminal:

```bash
clasp deployments                                      # copie el deploymentId de la Web App
clasp deploy -i <deploymentId> -d "v1.1 descripción"
```

---

## 5. Creación del Super Admin

1. Abra `…/exec?page=admin`. Si no hay ningún Super Admin, aparece **"Crear o restablecer Super Admin"**.
2. Escriba el **código de configuración** del registro de `setup()`, su nombre, email y una contraseña (mínimo 10 caracteres, letras y números).
3. Inicie sesión. Desde **Usuarios** cree el resto del equipo. A cada usuario nuevo se le muestra **una sola vez** una contraseña temporal que deberá cambiar al entrar.

**Recuperar el acceso** (olvidó la contraseña del único Super Admin): en el editor ejecute `createSetupCode()`, copie el código del registro y use de nuevo "Crear o restablecer Super Admin" con el mismo email. Se reemplaza la contraseña y se cierran sus sesiones.

Roles:

| Rol | Puede |
|---|---|
| `SUPER_ADMIN` | Todo, incluidos usuarios, configuración, auditoría, importación y exportación. |
| `ADMIN` | CRM completo, clientes, archivos, reportes y duplicados. |
| `VENDEDOR` | Solo los clientes asignados a él: ver, editar, actividades, seguimientos y ventas. |
| `LECTURA` | Solo consultar. |

---

## 6. Importar la base existente (30.780 registros)

1. Suba `CCH@E26-9316.xlsx` a Drive → ábralo → **Archivo › Guardar como Hojas de cálculo de Google**.
2. Copie el ID de la nueva hoja (el texto entre `/d/` y `/edit` en la URL).
3. En el editor: **Configuración del proyecto › Propiedades del script** → agregue `IMPORT_SOURCE_SPREADSHEET_ID` con ese ID.
4. Ejecute **`importBase()`** en el editor o use el botón **Configuración › Iniciar importación** del panel.

Reglas aplicadas (verificadas sobre la base real):

| Regla | Resultado esperado |
|---|---|
| D1 Propietario en blanco | 30.780 vacíos |
| D3/D8 NIT no se usa | La columna NIT no se lee ni se guarda |
| D5 Emails repetidos o mal escritos en blanco | 28.670 con email · 2.110 en blanco |
| D6 Actividad vacía ← CIIU | 7.754 completadas desde el CIIU |
| D7 Fijos de 7 dígitos → 608 | 1.496 teléfonos |
| D9 Municipios | Neiva 21.255 · Pitalito 7.289 · Palermo 1.208 · Rivera 1.028 |

La importación va por lotes de 2.500 filas y, si se acerca al límite de 6 minutos, se reprograma sola. Reimportar no duplica: solo agrega las filas que falten.

---

## 7. Pruebas

```bash
cd crm-web
npm test                 # 56 pruebas del servidor sobre el código real (sin conexión a Google)
npm run test:ui          # 5 recorridos en Chromium (requiere Playwright instalado)
npm run dev              # http://localhost:8787 y /?page=admin (admin@webpaya.test / ClaveSegura2026)
```

El checklist completo, incluidas las pruebas manuales sobre la implementación real, está en [`docs/PRUEBAS.md`](docs/PRUEBAS.md).

---

## 8. Problemas conocidos y límites de Google Apps Script

| Límite | Impacto | Cómo lo maneja el sistema |
|---|---|---|
| **6 minutos por ejecución** | Importación y tareas largas | Lotes reanudables con activador; reportes en cola cada 10 min |
| **~30 ejecuciones simultáneas** por usuario (todas las del formulario cuentan contra el propietario) | Picos de muchos clientes a la vez | Guardado con espera de 2 s, cargas secuenciales, bloqueo con reintento ("intente de nuevo en unos segundos") |
| **Activadores: 90 min/día** en cuentas @gmail.com | Reportes e índice | Cada reporte toma segundos; alcanza para cientos por día |
| **Email: ~100 destinatarios/día** en @gmail.com | Solo afecta al 2FA de administradores | Los clientes nunca reciben email (D12) |
| **Almacenamiento: 15 GB** en @gmail.com (compartidos con Gmail y Fotos) | Fotos de clientes | Las fotos se reducen en el navegador a 2000 px; máx. 10 MB por archivo y 100 MB por solicitud |
| **Tamaño de llamada `google.script.run`** | Archivos grandes | Máximo configurable de 20 MB; por defecto 10 MB |
| **CacheService: 100 KB por valor y máximo 6 h** | Sesiones e índice | Sesiones de 6 h con cierre a los 60 min sin uso; el índice se guarda en trozos y el activador lo reconstruye si expira |
| **La IP del cliente no está disponible** | Prueba de consentimiento | Se guarda el texto aceptado (versión + hash), la fecha y el nombre del firmante |
| **La página corre dentro de un iframe** de `googleusercontent.com` | `localStorage`, ventanas emergentes | Si el navegador bloquea el almacenamiento, el formulario sigue funcionando con guardado en el servidor |
| **No se puede tener dominio propio** directamente | URL larga de Google | Ver recomendaciones |
| **Sin "ejecutar como usuario"** para los administradores | El email de Google del admin no se conoce | Login propio con contraseña con hash, bloqueo por intentos y 2FA opcional |
| **Ordenar o borrar filas a mano rompe referencias** | Integridad | Todo se localiza por ID; el borrado es lógico. **No ordene las hojas** (use filtros de vista) |
| Si el cliente cambia de dispositivo | Pierde el borrador (no hay enlace por email, D12) | El borrador queda en el servidor y el admin lo ve en "Solicitudes › Borradores" |

---

## 9. Recomendaciones futuras

1. **Google Workspace** para la cuenta dueña: más cuota de email, triggers y almacenamiento, y cuentas administradas.
2. **Dominio propio para el formulario**: servir `Index.html` desde GitHub Pages (el repositorio ya tiene `CNAME`) llamando a un `doPost` JSON. La API ya responde en formato uniforme `{ok, data, error}`.
3. **Plantilla del reporte según el software de diseño** cuando se elija (D16): solo cambia `ReportService._render`.
4. **Pagos en línea** (Wompi / ePayco / Mercado Pago) con webhook firmado hacia `PAGOS`.
5. **IA para briefs y textos** de la página a partir del JSON del reporte, con aprobación humana.
6. **Migrar tablas calientes** (ACTIVIDADES, AUDITORIA) a Firestore o Cloud SQL si superan ~200.000 filas; `SheetService` concentra el acceso a datos.
7. **Revisión legal** de la política de tratamiento de datos (Ley 1581) antes de invitar masivamente a la base.

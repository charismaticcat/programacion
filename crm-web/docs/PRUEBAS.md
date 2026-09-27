# Checklist de pruebas

Hay dos niveles:

- **Automáticas** (`npm test`, `npm run test:ui`): ejecutan el código real de `src/` sobre servicios de Google simulados en memoria y un navegador Chromium. Corren sin cuenta de Google.
- **Manuales en la implementación real**: confirman lo que el simulador no puede reproducir, como cuotas reales, permisos de Drive o el iframe de Google.

Leyenda: 🤖 automática · 👤 manual en producción.

## Formulario público

| # | Caso | Cómo | Resultado esperado |
|---|---|---|---|
| P1 | Cliente existente | 🤖 `Público › cliente existente…` · 👤 buscar un teléfono de la base | Encuentra 1 resultado con teléfono enmascarado; al elegirlo precarga razón social, teléfono, email, actividad, municipio, dirección y barrio. El propietario queda vacío. |
| P2 | Búsqueda por nombre | 🤖 | Muestra máximo 5 resultados; teléfono y email **no** se precargan. |
| P3 | Cliente nuevo | 🤖 · 👤 "Llenar desde cero" y enviar | Cliente `ORIGEN=FORMULARIO`, oportunidad "Nuevo", seguimiento de revisión, carpeta `WEB-AAAA-NNNNNN_NOMBRE` con 6 subcarpetas y reporte generado en ≤ 10 min. |
| P4 | Duplicado | 🤖 | Un cliente nuevo con el email de un registro de la base aparece en **Duplicados** (regla MISMO_EMAIL). |
| P5 | Corrección | 🤖 | Elegido un registro de la base y cambiado su teléfono: los vacíos se completan y el cambio queda como corrección pendiente en **Solicitudes**. |
| P6 | Carga de imágenes | 🤖 · 👤 subir logo y fotos desde el celular | Vista previa inmediata; archivo en la subcarpeta correcta; metadatos en `ARCHIVOS`. |
| P7 | Archivo inválido | 🤖 `.exe`, PNG falso, PDF en el logo, base64 dañado | Mensaje "Tipo de archivo no permitido" o "contenido no corresponde"; nada llega a Drive. |
| P8 | Archivo demasiado grande | 🤖 10,5 MB | "El archivo supera el máximo de 10 MB". |
| P9 | Formulario incompleto | 🤖 · 👤 | La revisión lista lo que falta con enlaces al paso; el servidor devuelve los errores por campo. |
| P10 | Doble envío | 🤖 (servidor e interfaz: recargar tras enviar) · 👤 doble clic en Enviar | Un solo cliente; el segundo intento devuelve el mismo radicado. |
| P11 | Pérdida de conexión | 🤖 `ui-tests` con el navegador sin conexión · 👤 modo avión | Indicador "Sin conexión: guardado en este dispositivo", copia local y sincronización al volver. |
| P12 | Celular | 🤖 iPhone 12 · 👤 Android e iPhone reales | Barra de progreso móvil y sin desplazamiento horizontal; botones de 44 px o más. |
| P13 | Testimonio > 30 palabras | 🤖 | Contador en rojo y error en el servidor. |
| P14 | Envío en menos de 20 s (bot) | 🤖 | "Revise la información antes de enviar". |
| P15 | Inyección de fórmulas | 🤖 `=HYPERLINK(...)` | Se guarda como texto. |
| P16 | Municipio fuera de la lista | 🤖 | Rechazado: solo Neiva, Pitalito, Palermo y Rivera. |
| P17 | Accesibilidad | 👤 navegar solo con teclado; lector de pantalla | Foco visible, etiquetas en todos los campos, errores anunciados, foco al título al cambiar de paso. |

## Importación

| # | Caso | Cómo | Esperado |
|---|---|---|---|
| I1 | Reglas D1–D9 | 🤖 fixtures sintéticos · 👤 base real | Ver cifras en el README §6. |
| I2 | Reimportar | 🤖 | No duplica. |
| I3 | Reanudación | 👤 base real en Sheets | Si pasa de 4,5 min, continúa sola con el activador `cronContinueImport`. |

## Super Admin / CRM

| # | Caso | Cómo | Esperado |
|---|---|---|---|
| C1 | Login | 🤖 · 👤 | Credenciales incorrectas → mensaje genérico; 5 fallos → bloqueo de 15 min; contraseña temporal → cambio obligatorio. |
| C2 | Clientes | 🤖 · 👤 | Listado paginado de 30.780 con respuesta en segundos. |
| C3 | Búsqueda | 🤖 | Por nombre, teléfono, email o ID. |
| C4 | Filtros | 🤖 | Municipio, etapa, formulario, origen, responsable, prioridad, tamaño, CIIU, sin email, teléfono compartido, no contactar. |
| C5 | Ficha y edición | 🤖 | Todos los campos editables; conflicto si otro usuario editó antes; auditoría con valores antes y después. |
| C6 | Pipeline | 🤖 | Cambio de etapa; Ganado exige venta; Perdido y No interesado exigen motivo. |
| C7 | Kanban | 🤖 (10 columnas, mover con selector) · 👤 arrastrar tarjetas | Totales por columna; el cambio se refleja en el cliente. |
| C8 | Actividad | 🤖 llamada, WhatsApp, email, reunión y nota | La primera actividad saliente mueve a Contactado; "Volver a llamar" exige seguimiento; NO CONTACTAR bloquea. |
| C9 | Seguimiento | 🤖 | Vencidos, hoy y próximos; completar y reprogramar. |
| C10 | Oportunidad y propuesta | 🤖 | El total lo calcula el servidor (ignora un TOTAL manipulado); marcar enviada mueve a Propuesta enviada. |
| C11 | Venta y pagos | 🤖 | Ganado + proyecto web; pagos → Parcial / Pagada; anular revierte a Negociación. |
| C12 | Reporte | 🤖 · 👤 abrir el Doc | Google Doc en `05_DOCUMENTOS`, secciones de página, sin notas internas; regenerar crea v2 y conserva v1; PDF y JSON descargables. |
| C13 | Duplicados | 🤖 | Fusión manual: A se conserva, los datos de B se mueven y B queda "fusionado". |
| C14 | Auditoría | 🤖 | Creación, edición, cambio de etapa, actividad, seguimiento, venta, pago, login y accesos denegados. |
| C15 | Dashboard | 🤖 · 👤 | Indicadores coherentes con las hojas. |

## Seguridad

| # | Caso | Cómo | Esperado |
|---|---|---|---|
| S1 | Usuario no autenticado | 🤖 | `UNAUTHENTICATED` en toda función `adm_*`; el panel muestra el login. |
| S2 | Usuario sin permisos | 🤖 | LECTURA no escribe; ADMIN no ve auditoría ni exporta; VENDEDOR no ve clientes ajenos. Cada intento queda como DENEGADO en la auditoría. |
| S3 | Manipulación del frontend | 🤖 | Tablas no permitidas, campos no editables, municipios o emails inválidos, referencias inventadas y archivos de otra solicitud → rechazados. |
| S4 | Acceso a funciones administrativas | 🤖 | `setup`, `importBase` y `cron*` solo para el propietario o sus activadores; las funciones terminadas en `_` no son invocables; toda función global es `pub_*`, `adm_*` o de mantenimiento. |
| S5 | Secretos | 🤖 | Ninguna respuesta incluye hashes, sales, tokens ni IDs de Drive (al público). |
| S6 | Sesión | 🤖 | Vence tras 60 min sin uso; desactivar un usuario revoca sus sesiones. |
| S7 | Código de configuración | 🤖 | Es de un solo uso. |
| S8 | Permisos de Drive | 👤 abrir `CLIENTES_WEB` desde otra cuenta | Sin acceso; la tarea diaria alerta si algo se comparte públicamente. |

## Contrato y páginas

| # | Caso | Esperado |
|---|---|---|
| K1 | 🤖 Cada `rpc()`/`api()` del HTML existe en el servidor | Ninguna llamada huérfana. |
| K2 | 🤖 Los includes de plantilla existen | `doGet` sirve ambas páginas completas. |
| K3 | 🤖 Sin `innerHTML`/`insertAdjacentHTML` | Todo el contenido dinámico se inserta como texto. |

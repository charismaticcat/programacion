# Base de datos (Google Sheets)

> Documento generado automáticamente desde `src/Schema.gs` con `node tests/gen-db-doc.js`. No lo edite a mano.

## Archivos

| Spreadsheet | Hojas |
|---|---|
| `CRM_WEB_DB` | `CLIENTES`, `CONTACTOS`, `SEDES`, `SERVICIOS`, `PRODUCTOS`, `TESTIMONIOS`, `CONTENIDO_WEB`, `ARCHIVOS`, `AUTORIZACIONES`, `SOLICITUDES`, `REPORTES`, `OPORTUNIDADES`, `ACTIVIDADES`, `SEGUIMIENTOS`, `PROPUESTAS`, `VENTAS`, `PAGOS`, `PROYECTOS_WEB`, `DUPLICADOS`, `USUARIOS_ADMIN`, `CONFIGURACION`, `CATALOGOS`, `PRODUCTOS_WEB`, `SECUENCIAS` |
| `CRM_WEB_AUDITORIA` | `AUDITORIA` (solo se agregan filas) |

**Reglas comunes**

- Fila 1 = encabezados. El código ubica las columnas **por nombre**, no por posición: se pueden agregar columnas al final, pero **no renombre, ordene ni borre filas** a mano.
- Todas las celdas usan formato de texto plano (`@`): no se pierden ceros (`01112`) ni se convierten fechas.
- Los textos que empiezan por `= + - @` se guardan con apóstrofo para impedir fórmulas.
- Borrado lógico: `ELIMINADO = SI`. Sí/No se guarda como `SI`/`NO`.
- Columnas **editables** = se pueden cambiar desde el Super Admin (decisión D11). Las demás las mantiene el sistema.

**Columnas de control** (en las tablas marcadas con ✓):

| Columna | Tipo | Descripción |
|---|---|---|
| `CREADO_EN` | Fecha y hora | Fecha y hora de creación |
| `CREADO_POR` | Texto | PUBLICO:<radicado> · ADMIN:<email> · SISTEMA:<proceso> |
| `ACTUALIZADO_EN` | Fecha y hora | Última modificación |
| `ACTUALIZADO_POR` | Texto | Autor de la última modificación |
| `VERSION` | Número | Contador para detectar ediciones simultáneas |
| `ELIMINADO` | Sí/No | Borrado lógico: las filas nunca se borran |

## CLIENTES

Un registro por establecimiento/cliente. Los 30.780 de la base entran como ORIGEN=BASE_CCH. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `CLIENTE_ID` | Texto |  | CLI-000001… (número de fila de la base) o siguiente de la secuencia |
| `ORIGEN` | Lista |  | De dónde viene el registro · Valores: `BASE_CCH`, `FORMULARIO`, `ADMIN`, `REFERIDO` |
| `ID_REGISTRO_ORIGEN` | Texto |  | Número de fila en CCH@E26-9316 (vacío si no viene de la base) |
| `NOMBRE_PROPIETARIO` | Texto | Sí | En blanco al importar (D1); lo completa el cliente o un admin |
| `RAZON_SOCIAL` | Texto | Sí | Nombre legal o del establecimiento |
| `NOMBRE_COMERCIAL` | Texto | Sí | Solo lo escribe el cliente o un admin; nunca se infiere |
| `TIPO_ESTABLECIMIENTO` | Texto | Sí | Catálogo con autocompletado y texto libre (D15) |
| `ACTIVIDAD` | Texto largo | Sí | Actividad de la base o descripción CIIU si venía vacía (D6) |
| `ACTIVIDAD_FUENTE` | Lista |  | BASE, CIIU, CLIENTE o ADMIN · Valores: `BASE`, `CIIU`, `CLIENTE`, `ADMIN` |
| `CIIU_CODIGO` | Texto | Sí | Ej. G4721 |
| `CIIU_DESCRIPCION` | Texto largo | Sí | Texto del código CIIU |
| `TAMANO_EMPRESA` | Lista | Sí | Micro, pequeña, mediana o gran empresa · Valores: `MICRO EMPRESA`, `PEQUEÑA EMPRESA`, `MEDIANA EMPRESA`, `GRAN EMPRESA` |
| `TELEFONO` | Texto | Sí | 10 dígitos; los fijos de 7 dígitos llevan 608 (D7) |
| `TELEFONO_ORIGINAL` | Texto |  | Valor tal como venía en la base |
| `TELEFONO_COMPARTIDO` | Sí/No |  | Otro cliente tiene el mismo número |
| `EMAIL` | Texto | Sí | En blanco si en la base estaba repetido o mal escrito (D5) |
| `MUNICIPIO_COD` | Texto |  | Código DANE (41001, 41551, 41524, 41615) |
| `MUNICIPIO` | Lista | Sí | Solo Neiva, Pitalito, Palermo o Rivera (D9) · Valores: `NEIVA`, `PITALITO`, `PALERMO`, `RIVERA` |
| `DIRECCION` | Texto | Sí | Dirección comercial |
| `BARRIO_COD` | Texto |  | Código de barrio de la base |
| `BARRIO` | Texto | Sí | Barrio comercial |
| `FEC_MATRICULA` | Fecha |  | De la base |
| `FEC_RENOVACION` | Fecha |  | De la base |
| `ULT_ANO_RENOVACION` | Texto |  | De la base |
| `ESTADO_CRM` | Lista |  | Copia de la etapa de su oportunidad abierta más reciente (para filtrar rápido) · Valores: `NUEVO`, `POR_CONTACTAR`, `CONTACTADO`, `INTERESADO`, `REUNION`, `PROPUESTA_ENVIADA`, `NEGOCIACION`, `GANADO`, `PERDIDO`, `NO_INTERESADO` |
| `ESTADO_FORMULARIO` | Lista | Sí | NO_INICIADO, ENVIADO, EN_REVISION, APROBADO · Valores: `NO_INICIADO`, `ENVIADO`, `EN_REVISION`, `APROBADO` |
| `RESPONSABLE_EMAIL` | Texto |  | Email del usuario admin asignado |
| `PRIORIDAD` | Lista | Sí | ALTA, MEDIA, BAJA · Valores: `ALTA`, `MEDIA`, `BAJA` |
| `ETIQUETAS` | Texto | Sí | Separadas por coma |
| `FUSIONADO_EN` | Texto |  | CLIENTE_ID que sobrevivió a una fusión |
| `DRIVE_FOLDER_ID` | Texto |  | Carpeta del cliente en CLIENTES_WEB |
| `NO_CONTACTAR` | Sí/No | Sí | Oposición del titular: bloquea actividades salientes |
| `OBSERVACION_IMPORTACION` | Texto |  | Ajustes aplicados al importar |

## CONTACTOS

Personas de contacto del cliente. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `CONTACTO_ID` | Texto |  | ID contacto |
| `CLIENTE_ID` | Texto |  | Cliente |
| `SOLICITUD_ID` | Texto |  | Formulario que lo creó (vacío si lo creó un admin) |
| `NOMBRE` | Texto | Sí | Nombre |
| `CARGO` | Texto | Sí | Cargo |
| `ROL` | Lista | Sí | Rol · Valores: `PROPIETARIO`, `ADMINISTRADOR`, `EMPLEADO`, `OTRO` |
| `TELEFONO_FIJO` | Texto | Sí | Teléfono fijo |
| `WHATSAPP` | Texto | Sí | WhatsApp |
| `EMAIL` | Texto | Sí | Email |
| `ES_PRINCIPAL` | Sí/No | Sí | Principal |
| `PREFERENCIA_CANAL` | Lista | Sí | Canal preferido · Valores: `WHATSAPP`, `LLAMADA`, `EMAIL` |
| `HORARIO_CONTACTO` | Texto | Sí | Horario de contacto |
| `REDES_SOCIALES` | Texto largo | Sí | Enlaces o usuarios |

## SEDES

Sedes, direcciones y horarios. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `SEDE_ID` | Texto |  | ID sede |
| `CLIENTE_ID` | Texto |  | Cliente |
| `SOLICITUD_ID` | Texto |  | Solicitud |
| `NOMBRE_SEDE` | Texto | Sí | Nombre de la sede |
| `ES_PRINCIPAL` | Sí/No | Sí | Principal |
| `DIRECCION` | Texto | Sí | Dirección |
| `BARRIO` | Texto | Sí | Barrio |
| `MUNICIPIO` | Lista | Sí | Municipio · Valores: `NEIVA`, `PITALITO`, `PALERMO`, `RIVERA` |
| `REFERENCIA` | Texto | Sí | Punto de referencia |
| `GOOGLE_MAPS_URL` | Texto | Sí | Enlace de Google Maps |
| `TELEFONO` | Texto | Sí | Teléfono |
| `WHATSAPP` | Texto | Sí | WhatsApp |
| `HORARIO_SEMANA` | Texto | Sí | Horario entre semana |
| `HORARIO_FIN_SEMANA` | Texto | Sí | Horario fin de semana |
| `HORARIOS_ESPECIALES` | Texto largo | Sí | Festivos, temporadas… |
| `ATIENDE_EN_SITIO` | Sí/No | Sí | Sitio comercial abierto al público |
| `ORDEN` | Número | Sí | Orden |

## SERVICIOS

Servicios que ofrece el cliente (se publican en la web). Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `SERVICIO_ID` | Texto |  | ID servicio |
| `CLIENTE_ID` | Texto |  | Cliente |
| `SOLICITUD_ID` | Texto |  | Solicitud |
| `NOMBRE` | Texto | Sí | Nombre |
| `DESCRIPCION` | Texto largo | Sí | Descripción |
| `PRECIO` | Texto | Sí | Texto libre editable (D14): "25000", "Desde $20.000", "A convenir" |
| `MOSTRAR_PRECIO` | Sí/No | Sí | Mostrar precio |
| `REQUIERE_CITA` | Sí/No | Sí | Requiere cita |
| `IMAGEN_ARCHIVO_ID` | Texto |  | ARCHIVO_ID de la imagen |
| `ORDEN` | Número | Sí | Orden |
| `ACTIVO` | Sí/No | Sí | Activo |

## PRODUCTOS

Catálogo de productos del cliente. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `PRODUCTO_ID` | Texto |  | ID producto |
| `CLIENTE_ID` | Texto |  | Cliente |
| `SOLICITUD_ID` | Texto |  | Solicitud |
| `NOMBRE` | Texto | Sí | Nombre |
| `DESCRIPCION` | Texto largo | Sí | Descripción |
| `CATEGORIA` | Texto | Sí | Categoría |
| `PRECIO` | Texto | Sí | Texto libre editable (D14) |
| `MOSTRAR_PRECIO` | Sí/No | Sí | Mostrar precio |
| `DISPONIBLE` | Sí/No | Sí | Disponible |
| `IMAGEN_ARCHIVO_ID` | Texto |  | ARCHIVO_ID de la imagen |
| `ORDEN` | Número | Sí | Orden |

## TESTIMONIOS

Testimonios de clientes del negocio. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `TESTIMONIO_ID` | Texto |  | ID testimonio |
| `CLIENTE_ID` | Texto |  | Cliente |
| `SOLICITUD_ID` | Texto |  | Solicitud |
| `AUTOR_NOMBRE` | Texto | Sí | Autor |
| `TEXTO` | Texto largo | Sí | Máximo 30 palabras |
| `FOTO_ARCHIVO_ID` | Texto |  | ARCHIVO_ID de la foto |
| `PERMISO_USO_FOTO` | Sí/No | Sí | Permiso uso de foto |
| `PERMISO_USO_NOMBRE` | Sí/No | Sí | Permiso uso de nombre |
| `APROBADO_ADMIN` | Sí/No | Sí | Solo los aprobados se marcan como publicables |

## CONTENIDO_WEB

Respuestas de la plantilla de página (clave-valor). Claves en CONTENT_KEYS. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `CONTENIDO_ID` | Texto |  | ID |
| `CLIENTE_ID` | Texto |  | Cliente |
| `SOLICITUD_ID` | Texto |  | Solicitud |
| `CAMPO_CLAVE` | Texto |  | Clave de CONTENT_KEYS |
| `VALOR` | Texto largo | Sí | Valor |

## ARCHIVOS

Metadatos de archivos. El archivo vive en Drive, nunca en Sheets. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `ARCHIVO_ID` | Texto |  | ID archivo |
| `CLIENTE_ID` | Texto |  | Se completa al enviar el formulario |
| `SOLICITUD_ID` | Texto |  | Solicitud |
| `ENTIDAD` | Lista |  | A qué pertenece · Valores: `CLIENTE`, `SERVICIO`, `PRODUCTO`, `TESTIMONIO`, `VENTA`, `PAGO`, `SOLICITUD` |
| `ENTIDAD_ID` | Texto |  | SERVICIO_ID, PRODUCTO_ID… |
| `TEMP_REF` | Texto |  | Referencia del ítem mientras el formulario es borrador |
| `CATEGORIA` | Lista | Sí | Categoría · Valores: `LOGO`, `FOTO_NEGOCIO`, `FOTO_EQUIPO`, `FOTO_PRODUCTO`, `FOTO_SERVICIO`, `FOTO_TESTIMONIO`, `DOCUMENTO`, `COMPROBANTE`, `PROPUESTA`, `REPORTE`, `SNAPSHOT`, `OTRO` |
| `NOMBRE_ORIGINAL` | Texto |  | Nombre con que se subió (solo informativo) |
| `NOMBRE_DRIVE` | Texto |  | <CATEGORIA>_<ID>_<fecha>.<ext> |
| `DRIVE_FILE_ID` | Texto |  | Privado: nunca se envía al formulario público |
| `DRIVE_FOLDER_ID` | Texto |  | Carpeta Drive |
| `MIME_TYPE` | Texto |  | Detectado por firma de bytes |
| `TAMANO_BYTES` | Número |  | Tamaño (bytes) |
| `SHA256` | Texto |  | Detecta archivos repetidos |
| `ESTADO` | Lista | Sí | Estado · Valores: `PENDIENTE`, `APROBADO`, `RECHAZADO`, `PAPELERA` |
| `MOTIVO_RECHAZO` | Texto | Sí | Motivo de rechazo |

## AUTORIZACIONES

Prueba de cada autorización otorgada o negada. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `AUTORIZACION_ID` | Texto |  | ID |
| `CLIENTE_ID` | Texto |  | Cliente |
| `SOLICITUD_ID` | Texto |  | Solicitud |
| `TIPO` | Lista |  | Tipo · Valores: `TRATAMIENTO_DATOS`, `USO_IMAGENES`, `PUBLICACION_CONTENIDO`, `USO_FOTOS_TESTIMONIOS`, `CONTACTO_WHATSAPP`, `CONTACTO_LLAMADA`, `CONTACTO_EMAIL`, `USO_RAZON_SOCIAL_BASE` |
| `OTORGADA` | Sí/No |  | Otorgada |
| `TEXTO_VERSION` | Texto |  | Versión del texto |
| `TEXTO_HASH` | Texto |  | SHA-256 del texto mostrado |
| `FECHA_HORA` | Fecha y hora |  | Fecha y hora |
| `FIRMANTE_NOMBRE` | Texto |  | Firmante |
| `USER_AGENT` | Texto |  | Declarado por el navegador |
| `REVOCADA_EN` | Fecha y hora | Sí | Si el titular revoca |

## SOLICITUDES

Cada formulario diligenciado (borrador o enviado). Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `SOLICITUD_ID` | Texto |  | ID solicitud |
| `RADICADO` | Texto |  | WEB-2026-000001 (se asigna al iniciar) |
| `CLIENTE_ID` | Texto |  | Cliente al que quedó asociada al enviar |
| `CLIENTE_BASE_ID` | Texto |  | Registro que el usuario eligió en la búsqueda |
| `TIPO_VINCULACION` | Lista |  | Cómo encontró su registro · Valores: `NINGUNA`, `BUSQUEDA_TELEFONO`, `BUSQUEDA_EMAIL`, `BUSQUEDA_NOMBRE` |
| `ESTADO` | Lista | Sí | Estado · Valores: `BORRADOR`, `ENVIADA`, `EN_REVISION`, `APROBADA`, `DESCARTADA` |
| `PASO_ACTUAL` | Número |  | Paso actual |
| `PORCENTAJE` | Número |  | Porcentaje |
| `TOKEN_HASH` | Secreto |  | SHA-256 del token del navegador (el token no se guarda) |
| `DATOS_JSON` | JSON |  | Estado completo del formulario |
| `DATOS_FILE_ID` | Texto |  | Si el borrador supera el tamaño de celda |
| `CAMBIOS_PROPUESTOS_JSON` | JSON |  | Diferencias frente al registro de la base |
| `CAMBIOS_ESTADO` | Lista |  | Estado de correcciones · Valores: `NINGUNO`, `PENDIENTE`, `APLICADO`, `DESCARTADO` |
| `DRIVE_FOLDER_ID` | Texto |  | WEB-2026-000001_NOMBRE |
| `SNAPSHOT_FILE_ID` | Texto |  | JSON congelado al enviar |
| `REPORTE_ESTADO` | Lista |  | Estado del reporte · Valores: `NO_APLICA`, `PENDIENTE`, `GENERADO`, `ERROR`, `REEMPLAZADO` |
| `INICIADA_EN` | Fecha y hora |  | Iniciada en |
| `ENVIADA_EN` | Fecha y hora |  | Enviada en |
| `USER_AGENT` | Texto |  | Navegador |
| `OBSERVACIONES_ADMIN` | Texto largo | Sí | Nota interna (no sale en el reporte) |
| `REVISADO_POR` | Texto |  | Revisado por |
| `REVISADO_EN` | Fecha y hora |  | Revisado en |

## REPORTES

Reportes editables (Google Docs) por cliente, versionados. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `REPORTE_ID` | Texto |  | ID reporte |
| `CLIENTE_ID` | Texto |  | Cliente |
| `SOLICITUD_ID` | Texto |  | Solicitud |
| `VERSION_REPORTE` | Número |  | 1, 2, 3… |
| `DOC_ID` | Texto |  | ID del Doc |
| `DOC_URL` | Texto |  | URL del Doc |
| `JSON_FILE_ID` | Texto |  | Mismos datos en JSON |
| `GENERADO_POR` | Texto |  | Generado por |
| `GENERADO_EN` | Fecha y hora |  | Generado en |
| `ESTADO` | Lista |  | Estado · Valores: `NO_APLICA`, `PENDIENTE`, `GENERADO`, `ERROR`, `REEMPLAZADO` |
| `ERROR` | Texto |  | Error |

## OPORTUNIDADES

Negocios en el pipeline (Kanban). Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `OPORTUNIDAD_ID` | Texto |  | ID oportunidad |
| `CLIENTE_ID` | Texto |  | Cliente |
| `TITULO` | Texto | Sí | Título |
| `PRODUCTO_WEB_ID` | Lista | Sí | De PRODUCTOS_WEB |
| `ETAPA` | Lista |  | Etapa · Valores: `NUEVO`, `POR_CONTACTAR`, `CONTACTADO`, `INTERESADO`, `REUNION`, `PROPUESTA_ENVIADA`, `NEGOCIACION`, `GANADO`, `PERDIDO`, `NO_INTERESADO` |
| `PROBABILIDAD` | Número | Sí | Sugerida por la etapa, editable |
| `VALOR_ESTIMADO` | Dinero (COP) | Sí | Editable (D14) |
| `FECHA_CIERRE_ESTIMADA` | Fecha | Sí | Cierre estimado |
| `ORIGEN_LEAD` | Texto | Sí | FORMULARIO, BASE, REFERIDO… |
| `RESPONSABLE_EMAIL` | Texto |  | Responsable |
| `MOTIVO_CIERRE` | Texto largo | Sí | Obligatorio en PERDIDO y NO_INTERESADO |
| `FECHA_CIERRE_REAL` | Fecha |  | Fecha de cierre |
| `FECHA_CAMBIO_ETAPA` | Fecha y hora |  | Último cambio de etapa |
| `ABIERTA` | Sí/No |  | Abierta |

## ACTIVIDADES

Llamadas, WhatsApp, emails, reuniones y notas internas. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `ACTIVIDAD_ID` | Texto |  | ID actividad |
| `CLIENTE_ID` | Texto |  | Cliente |
| `OPORTUNIDAD_ID` | Texto |  | Oportunidad |
| `TIPO` | Lista | Sí | Tipo · Valores: `LLAMADA`, `WHATSAPP`, `EMAIL`, `REUNION`, `NOTA`, `VISITA`, `CAMBIO_ETAPA`, `SISTEMA` |
| `DIRECCION` | Lista | Sí | Dirección · Valores: `SALIENTE`, `ENTRANTE`, `INTERNA` |
| `RESULTADO` | Lista | Sí | Resultado · Valores: `CONTESTO`, `NO_CONTESTO`, `NUMERO_ERRADO`, `INTERESADO`, `NO_INTERESADO`, `VOLVER_A_LLAMAR`, `ENVIADO`, `RESPONDIDO`, `REALIZADA`, `CANCELADA`, `REPROGRAMADA` |
| `ASUNTO` | Texto | Sí | Asunto |
| `DETALLE` | Texto largo | Sí | Nota interna del CRM |
| `DURACION_MIN` | Número | Sí | Duración (min) |
| `FECHA_HORA` | Fecha y hora | Sí | Fecha y hora |
| `REALIZADA_POR` | Texto |  | Realizada por |

## SEGUIMIENTOS

Tareas con fecha de vencimiento. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `SEGUIMIENTO_ID` | Texto |  | ID seguimiento |
| `CLIENTE_ID` | Texto |  | Cliente |
| `OPORTUNIDAD_ID` | Texto |  | Oportunidad |
| `ACTIVIDAD_ORIGEN_ID` | Texto |  | Actividad de origen |
| `TIPO` | Lista | Sí | Tipo · Valores: `LLAMAR`, `WHATSAPP`, `EMAIL`, `REUNION`, `ENVIAR_PROPUESTA`, `REVISAR_SOLICITUD`, `COBRAR`, `OTRO` |
| `DESCRIPCION` | Texto largo | Sí | Descripción |
| `FECHA_VENCIMIENTO` | Fecha | Sí | Vence |
| `PRIORIDAD` | Lista | Sí | Prioridad · Valores: `ALTA`, `MEDIA`, `BAJA` |
| `ESTADO` | Lista |  | Vencido se calcula: PENDIENTE con fecha pasada · Valores: `PENDIENTE`, `COMPLETADO`, `CANCELADO` |
| `ASIGNADO_A` | Texto | Sí | Email de usuario admin |
| `COMPLETADO_EN` | Fecha y hora |  | Completado en |
| `RESULTADO` | Texto largo | Sí | Resultado |

## PROPUESTAS

Propuestas comerciales con ítems y valores editables. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `PROPUESTA_ID` | Texto |  | ID propuesta |
| `OPORTUNIDAD_ID` | Texto |  | Oportunidad |
| `CLIENTE_ID` | Texto |  | Cliente |
| `NUMERO` | Texto |  | PRO-2026-0001 |
| `ITEMS_JSON` | JSON |  | [{descripcion, cantidad, precioUnitario}] |
| `SUBTOTAL` | Dinero (COP) |  | Calculado en el servidor |
| `DESCUENTO` | Dinero (COP) | Sí | Descuento |
| `TOTAL` | Dinero (COP) |  | Calculado en el servidor |
| `VALIDEZ_HASTA` | Fecha | Sí | Válida hasta |
| `ESTADO` | Lista |  | Estado · Valores: `BORRADOR`, `ENVIADA`, `ACEPTADA`, `RECHAZADA`, `VENCIDA` |
| `NOTAS` | Texto largo | Sí | Notas |
| `ENVIADA_EN` | Fecha y hora |  | Enviada en |
| `RESPUESTA_EN` | Fecha y hora |  | Respuesta en |

## VENTAS

Cierres de venta. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `VENTA_ID` | Texto |  | ID venta |
| `OPORTUNIDAD_ID` | Texto |  | Oportunidad |
| `PROPUESTA_ID` | Texto |  | Propuesta |
| `CLIENTE_ID` | Texto |  | Cliente |
| `NUMERO` | Texto |  | VEN-2026-0001 |
| `FECHA_VENTA` | Fecha | Sí | Fecha de venta |
| `VALOR_TOTAL` | Dinero (COP) | Sí | Editable (D14) |
| `FORMA_PAGO` | Lista | Sí | Forma de pago · Valores: `CONTADO`, `ANTICIPO_Y_SALDO`, `CUOTAS`, `OTRO` |
| `ESTADO_PAGO` | Lista |  | Calculado con los pagos · Valores: `PENDIENTE`, `PARCIAL`, `PAGADA`, `ANULADA` |
| `VALOR_PAGADO` | Dinero (COP) |  | Suma de PAGOS |
| `VENDEDOR_EMAIL` | Texto |  | Vendedor |
| `NOTAS` | Texto largo | Sí | Notas |

## PAGOS

Pagos recibidos por venta. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `PAGO_ID` | Texto |  | ID pago |
| `VENTA_ID` | Texto |  | Venta |
| `CLIENTE_ID` | Texto |  | Cliente |
| `FECHA_PAGO` | Fecha | Sí | Fecha de pago |
| `VALOR` | Dinero (COP) | Sí | Valor |
| `METODO` | Lista | Sí | Método · Valores: `TRANSFERENCIA`, `EFECTIVO`, `NEQUI`, `DAVIPLATA`, `TARJETA`, `OTRO` |
| `REFERENCIA` | Texto | Sí | N.º de transacción |
| `COMPROBANTE_ARCHIVO_ID` | Texto |  | ARCHIVO_ID |
| `NOTAS` | Texto largo | Sí | Notas |

## PROYECTOS_WEB

Proyecto de construcción de la página (se crea al registrar la venta). Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `PROYECTO_ID` | Texto |  | ID proyecto |
| `VENTA_ID` | Texto |  | Venta |
| `CLIENTE_ID` | Texto |  | Cliente |
| `ESTADO` | Lista | Sí | Estado · Valores: `BRIEF`, `DISENO`, `DESARROLLO`, `REVISION_CLIENTE`, `PUBLICADO`, `MANTENIMIENTO`, `CANCELADO` |
| `REPORTE_ID` | Texto |  | Reporte usado para construir la página |
| `DOMINIO` | Texto | Sí | Dominio |
| `URL_PUBLICADA` | Texto | Sí | URL publicada |
| `FECHA_ENTREGA_ESTIMADA` | Fecha | Sí | Entrega estimada |
| `RESPONSABLE_EMAIL` | Texto | Sí | Responsable |
| `NOTAS` | Texto largo | Sí | Notas |

## DUPLICADOS

Posibles duplicados y conciliación formulario ↔ base. Toda fusión es manual. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `DUP_ID` | Texto |  | ID |
| `CLIENTE_A` | Texto |  | Normalmente el registro de la base |
| `CLIENTE_B` | Texto |  | Normalmente el creado por el formulario |
| `REGLA` | Texto |  | MISMO_EMAIL, MISMO_TELEFONO_Y_NOMBRE… |
| `PUNTAJE` | Número |  | 0-100 |
| `ESTADO` | Lista |  | Estado · Valores: `PENDIENTE`, `ES_EL_MISMO`, `NO_ES_EL_MISMO` |
| `RESUELTO_POR` | Texto |  | Resuelto por |
| `RESUELTO_EN` | Fecha y hora |  | Resuelto en |

## USUARIOS_ADMIN

Usuarios del Super Admin. La contraseña se guarda como hash con sal. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `USUARIO_ID` | Texto |  | ID usuario |
| `EMAIL` | Texto |  | Usuario de acceso |
| `NOMBRE` | Texto | Sí | Nombre |
| `ROL` | Lista |  | Rol · Valores: `SUPER_ADMIN`, `ADMIN`, `VENDEDOR`, `LECTURA` |
| `ACTIVO` | Sí/No |  | Activo |
| `PASSWORD_HASH` | Secreto |  | HMAC-SHA256 iterado |
| `PASSWORD_SALT` | Secreto |  | Sal |
| `PASSWORD_ITER` | Secreto |  | Iteraciones |
| `DEBE_CAMBIAR_PASSWORD` | Sí/No |  | Debe cambiar contraseña |
| `INTENTOS_FALLIDOS` | Número |  | Intentos fallidos |
| `BLOQUEADO_HASTA` | Fecha y hora |  | Bloqueado hasta |
| `ULTIMO_ACCESO` | Fecha y hora |  | Último acceso |

## CONFIGURACION

Ajustes editables desde el panel (EDITABLE_SETTINGS).

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `CLAVE` | Texto |  | Clave |
| `VALOR` | Texto largo | Sí | Valor |
| `DESCRIPCION` | Texto |  | Descripción |
| `ACTUALIZADO_EN` | Fecha y hora |  | Actualizado en |
| `ACTUALIZADO_POR` | Texto |  | Actualizado por |

## CATALOGOS

Listas con autocompletado: MUNICIPIOS, TIPOS_ESTABLECIMIENTO, MOTIVOS_CIERRE. Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `ITEM_ID` | Texto |  | ID |
| `CATALOGO` | Texto |  | Catálogo |
| `CODIGO` | Texto | Sí | Código |
| `ETIQUETA` | Texto | Sí | Etiqueta |
| `ORDEN` | Número | Sí | Orden |
| `ACTIVO` | Sí/No | Sí | Activo |

## PRODUCTOS_WEB

Lo que vende Webpaya. Precio sugerido opcional y editable (D14). Columnas de control: ✓.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `PRODUCTO_WEB_ID` | Texto |  | ID |
| `NOMBRE` | Texto | Sí | Nombre |
| `DESCRIPCION` | Texto largo | Sí | Descripción |
| `PRECIO_SUGERIDO` | Dinero (COP) | Sí | Opcional; se precarga y se edita en cada propuesta |
| `ES_ADICIONAL` | Sí/No | Sí | Agenda de citas, base de datos de clientes… |
| `ACTIVO` | Sí/No | Sí | Activo |
| `ORDEN` | Número | Sí | Orden |

## SECUENCIAS

Contadores para radicados y números consecutivos.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `NOMBRE` | Texto |  | Nombre |
| `ULTIMO_VALOR` | Número |  | Último valor |

## AUDITORIA

Historial inmutable (spreadsheet separado). Solo se agregan filas.

| Columna | Tipo | Editable | Descripción |
|---|---|---|---|
| `AUDIT_ID` | Texto |  | ID |
| `FECHA_HORA` | Fecha y hora |  | Fecha y hora |
| `ACTOR` | Texto |  | PUBLICO:<radicado> · ADMIN:<email> · SISTEMA:<proceso> |
| `ROL` | Texto |  | Rol |
| `APP` | Texto |  | PUBLICA, ADMIN, SISTEMA |
| `ACCION` | Texto |  | CREAR, EDITAR, CAMBIO_ETAPA, ACTIVIDAD, SEGUIMIENTO, VENTA, LOGIN… |
| `ENTIDAD` | Texto |  | Tabla afectada |
| `ENTIDAD_ID` | Texto |  | ID |
| `CLIENTE_ID` | Texto |  | Cliente |
| `CAMPOS` | Texto |  | Campos modificados |
| `ANTES_JSON` | JSON |  | Valores anteriores |
| `DESPUES_JSON` | JSON |  | Valores nuevos |
| `RESULTADO` | Texto |  | OK, DENEGADO, ERROR |
| `DETALLE` | Texto largo |  | Detalle |

## Claves de CONTENIDO_WEB

| Clave | Pregunta | Sección del reporte |
|---|---|---|
| `nombre_pagina_1` | Nombre de página deseado (opción 1) | FICHA |
| `nombre_pagina_2` | Nombre de página deseado (opción 2) | FICHA |
| `tipo_establecimiento_alt` | Tipo de establecimiento (opción 2) | QUIENES |
| `mensaje` | ¿Qué le gustaría transmitir en su página web? | PORTADA |
| `publico_objetivo` | Público objetivo | QUIENES |
| `diferenciales` | Qué lo hace diferente | QUIENES |
| `colores` | Colores de marca | MARCA |
| `referencias` | Páginas de referencia que le gustan | MARCA |
| `dominio_actual` | Dominio o página actual | MARCA |
| `domicilios_ofrece` | Ofrece domicilios | CATALOGO |
| `domicilios_costo` | Costo del domicilio | CATALOGO |
| `agenda_citas` | Desea agenda de citas | ADICIONALES |
| `agenda_version_pago` | Agenda de citas: versión de pago | ADICIONALES |
| `bd_clientes` | Desea base de datos de clientes | ADICIONALES |
| `bd_version_pago` | Base de datos de clientes: versión de pago | ADICIONALES |

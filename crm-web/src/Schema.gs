/**
 * Schema.gs — Fuente única del modelo de datos.
 *
 * Cada tabla = una hoja. Cada columna: [NOMBRE, TIPO, ETIQUETA, DESCRIPCIÓN, EDITABLE_ADMIN, ENUM]
 * Tipos: s texto corto · t texto largo · n número · m dinero (entero COP) · b sí/no
 *        d fecha yyyy-MM-dd · dt fecha-hora · e lista (ENUM) · j JSON · x secreto (nunca sale del servidor)
 * Todas las celdas se guardan con formato de texto plano (@) para no perder ceros ni convertir fechas.
 */

const ENUMS = Object.freeze({
  ORIGEN: ['BASE_CCH', 'FORMULARIO', 'ADMIN', 'REFERIDO'],
  // Parte G: etapas del CRM (pipeline / Kanban)
  ETAPA: ['NUEVO', 'POR_CONTACTAR', 'CONTACTADO', 'INTERESADO', 'REUNION', 'PROPUESTA_ENVIADA', 'NEGOCIACION', 'GANADO', 'PERDIDO', 'NO_INTERESADO'],
  ESTADO_FORMULARIO: ['NO_INICIADO', 'ENVIADO', 'EN_REVISION', 'APROBADO'],
  PRIORIDAD: ['ALTA', 'MEDIA', 'BAJA'],
  TAMANO: ['MICRO EMPRESA', 'PEQUEÑA EMPRESA', 'MEDIANA EMPRESA', 'GRAN EMPRESA'],
  ACTIVIDAD_FUENTE: ['BASE', 'CIIU', 'CLIENTE', 'ADMIN'],
  ROL_CONTACTO: ['PROPIETARIO', 'ADMINISTRADOR', 'EMPLEADO', 'OTRO'],
  CANAL: ['WHATSAPP', 'LLAMADA', 'EMAIL'],
  ARCHIVO_CATEGORIA: ['LOGO', 'FOTO_NEGOCIO', 'FOTO_EQUIPO', 'FOTO_PRODUCTO', 'FOTO_SERVICIO', 'FOTO_TESTIMONIO', 'DOCUMENTO', 'COMPROBANTE', 'PROPUESTA', 'REPORTE', 'SNAPSHOT', 'OTRO'],
  ARCHIVO_ESTADO: ['PENDIENTE', 'APROBADO', 'RECHAZADO', 'PAPELERA'],
  ENTIDAD_ARCHIVO: ['CLIENTE', 'SERVICIO', 'PRODUCTO', 'TESTIMONIO', 'VENTA', 'PAGO', 'SOLICITUD'],
  AUT_TIPO: ['TRATAMIENTO_DATOS', 'USO_IMAGENES', 'PUBLICACION_CONTENIDO', 'USO_FOTOS_TESTIMONIOS', 'CONTACTO_WHATSAPP', 'CONTACTO_LLAMADA', 'CONTACTO_EMAIL', 'USO_RAZON_SOCIAL_BASE'],
  SOL_ESTADO: ['BORRADOR', 'ENVIADA', 'EN_REVISION', 'APROBADA', 'DESCARTADA'],
  VINCULACION: ['NINGUNA', 'BUSQUEDA_TELEFONO', 'BUSQUEDA_EMAIL', 'BUSQUEDA_NOMBRE'],
  CAMBIOS_ESTADO: ['NINGUNO', 'PENDIENTE', 'APLICADO', 'DESCARTADO'],
  REPORTE_ESTADO: ['NO_APLICA', 'PENDIENTE', 'GENERADO', 'ERROR', 'REEMPLAZADO'],
  ACT_TIPO: ['LLAMADA', 'WHATSAPP', 'EMAIL', 'REUNION', 'NOTA', 'VISITA', 'CAMBIO_ETAPA', 'SISTEMA'],
  ACT_DIRECCION: ['SALIENTE', 'ENTRANTE', 'INTERNA'],
  ACT_RESULTADO: ['', 'CONTESTO', 'NO_CONTESTO', 'NUMERO_ERRADO', 'INTERESADO', 'NO_INTERESADO', 'VOLVER_A_LLAMAR', 'ENVIADO', 'RESPONDIDO', 'REALIZADA', 'CANCELADA', 'REPROGRAMADA'],
  SEG_TIPO: ['LLAMAR', 'WHATSAPP', 'EMAIL', 'REUNION', 'ENVIAR_PROPUESTA', 'REVISAR_SOLICITUD', 'COBRAR', 'OTRO'],
  SEG_ESTADO: ['PENDIENTE', 'COMPLETADO', 'CANCELADO'],
  PROP_ESTADO: ['BORRADOR', 'ENVIADA', 'ACEPTADA', 'RECHAZADA', 'VENCIDA'],
  ESTADO_PAGO: ['PENDIENTE', 'PARCIAL', 'PAGADA', 'ANULADA'],
  FORMA_PAGO: ['CONTADO', 'ANTICIPO_Y_SALDO', 'CUOTAS', 'OTRO'],
  METODO_PAGO: ['TRANSFERENCIA', 'EFECTIVO', 'NEQUI', 'DAVIPLATA', 'TARJETA', 'OTRO'],
  PROY_ESTADO: ['BRIEF', 'DISENO', 'DESARROLLO', 'REVISION_CLIENTE', 'PUBLICADO', 'MANTENIMIENTO', 'CANCELADO'],
  DUP_ESTADO: ['PENDIENTE', 'ES_EL_MISMO', 'NO_ES_EL_MISMO'],
  ROL: ['SUPER_ADMIN', 'ADMIN', 'VENDEDOR', 'LECTURA'],
  SI_NO: ['', 'SI', 'NO']
});

/** Etapas del CRM: etiqueta, probabilidad y si cierra la oportunidad. */
const STAGES = Object.freeze({
  NUEVO: { label: 'Nuevo', prob: 5, closed: false },
  POR_CONTACTAR: { label: 'Por contactar', prob: 10, closed: false },
  CONTACTADO: { label: 'Contactado', prob: 20, closed: false },
  INTERESADO: { label: 'Interesado', prob: 35, closed: false },
  REUNION: { label: 'Reunión', prob: 50, closed: false },
  PROPUESTA_ENVIADA: { label: 'Propuesta enviada', prob: 60, closed: false },
  NEGOCIACION: { label: 'Negociación', prob: 75, closed: false },
  GANADO: { label: 'Ganado', prob: 100, closed: true },
  PERDIDO: { label: 'Perdido', prob: 0, closed: true },
  NO_INTERESADO: { label: 'No interesado', prob: 0, closed: true }
});

const CONTROL_COLUMNS = [
  ['CREADO_EN', 'dt', 'Creado en', 'Fecha y hora de creación', 0],
  ['CREADO_POR', 's', 'Creado por', 'PUBLICO:<radicado> · ADMIN:<email> · SISTEMA:<proceso>', 0],
  ['ACTUALIZADO_EN', 'dt', 'Actualizado en', 'Última modificación', 0],
  ['ACTUALIZADO_POR', 's', 'Actualizado por', 'Autor de la última modificación', 0],
  ['VERSION', 'n', 'Versión', 'Contador para detectar ediciones simultáneas', 0],
  ['ELIMINADO', 'b', 'Eliminado', 'Borrado lógico: las filas nunca se borran', 0]
];

/**
 * Definición de tablas.
 * key: columna ID · prefix: prefijo de IDs · control: agrega columnas de control
 * clientScoped: tiene CLIENTE_ID (se reasigna al fusionar clientes)
 */
const SCHEMA = {
  CLIENTES: {
    key: 'CLIENTE_ID', prefix: 'CLI', control: true, clientScoped: false,
    desc: 'Un registro por establecimiento/cliente. Los 30.780 de la base entran como ORIGEN=BASE_CCH.',
    columns: [
      ['CLIENTE_ID', 's', 'ID cliente', 'CLI-000001… (número de fila de la base) o siguiente de la secuencia', 0],
      ['ORIGEN', 'e', 'Origen', 'De dónde viene el registro', 0, 'ORIGEN'],
      ['ID_REGISTRO_ORIGEN', 's', 'Fila en la base', 'Número de fila en CCH@E26-9316 (vacío si no viene de la base)', 0],
      ['NOMBRE_PROPIETARIO', 's', 'Nombre del propietario', 'En blanco al importar (D1); lo completa el cliente o un admin', 1],
      ['RAZON_SOCIAL', 's', 'Razón social', 'Nombre legal o del establecimiento', 1],
      ['NOMBRE_COMERCIAL', 's', 'Nombre comercial', 'Solo lo escribe el cliente o un admin; nunca se infiere', 1],
      ['TIPO_ESTABLECIMIENTO', 's', 'Tipo de establecimiento', 'Catálogo con autocompletado y texto libre (D15)', 1],
      ['ACTIVIDAD', 't', 'Actividad', 'Actividad de la base o descripción CIIU si venía vacía (D6)', 1],
      ['ACTIVIDAD_FUENTE', 'e', 'Fuente de la actividad', 'BASE, CIIU, CLIENTE o ADMIN', 0, 'ACTIVIDAD_FUENTE'],
      ['CIIU_CODIGO', 's', 'Código CIIU', 'Ej. G4721', 1],
      ['CIIU_DESCRIPCION', 't', 'Descripción CIIU', 'Texto del código CIIU', 1],
      ['TAMANO_EMPRESA', 'e', 'Tamaño de empresa', 'Micro, pequeña, mediana o gran empresa', 1, 'TAMANO'],
      ['TELEFONO', 's', 'Teléfono', '10 dígitos; los fijos de 7 dígitos llevan 608 (D7)', 1],
      ['TELEFONO_ORIGINAL', 's', 'Teléfono original', 'Valor tal como venía en la base', 0],
      ['TELEFONO_COMPARTIDO', 'b', 'Teléfono compartido', 'Otro cliente tiene el mismo número', 0],
      ['EMAIL', 's', 'Email', 'En blanco si en la base estaba repetido o mal escrito (D5)', 1],
      ['MUNICIPIO_COD', 's', 'Código municipio', 'Código DANE (41001, 41551, 41524, 41615)', 0],
      ['MUNICIPIO', 'e', 'Municipio', 'Solo Neiva, Pitalito, Palermo o Rivera (D9)', 1, 'cat:MUNICIPIOS'],
      ['DIRECCION', 's', 'Dirección', 'Dirección comercial', 1],
      ['BARRIO_COD', 's', 'Código barrio', 'Código de barrio de la base', 0],
      ['BARRIO', 's', 'Barrio', 'Barrio comercial', 1],
      ['FEC_MATRICULA', 'd', 'Fecha de matrícula', 'De la base', 0],
      ['FEC_RENOVACION', 'd', 'Fecha de renovación', 'De la base', 0],
      ['ULT_ANO_RENOVACION', 's', 'Último año renovado', 'De la base', 0],
      ['ESTADO_CRM', 'e', 'Etapa CRM', 'Copia de la etapa de su oportunidad abierta más reciente (para filtrar rápido)', 0, 'ETAPA'],
      ['ESTADO_FORMULARIO', 'e', 'Estado del formulario', 'NO_INICIADO, ENVIADO, EN_REVISION, APROBADO', 1, 'ESTADO_FORMULARIO'],
      ['RESPONSABLE_EMAIL', 's', 'Responsable', 'Email del usuario admin asignado', 0],
      ['PRIORIDAD', 'e', 'Prioridad', 'ALTA, MEDIA, BAJA', 1, 'PRIORIDAD'],
      ['ETIQUETAS', 's', 'Etiquetas', 'Separadas por coma', 1],
      ['FUSIONADO_EN', 's', 'Fusionado en', 'CLIENTE_ID que sobrevivió a una fusión', 0],
      ['DRIVE_FOLDER_ID', 's', 'Carpeta Drive', 'Carpeta del cliente en CLIENTES_WEB', 0],
      ['NO_CONTACTAR', 'b', 'No contactar', 'Oposición del titular: bloquea actividades salientes', 1],
      ['OBSERVACION_IMPORTACION', 's', 'Observación de importación', 'Ajustes aplicados al importar', 0]
    ]
  },

  CONTACTOS: {
    key: 'CONTACTO_ID', prefix: 'CON', control: true, clientScoped: true,
    desc: 'Personas de contacto del cliente.',
    columns: [
      ['CONTACTO_ID', 's', 'ID contacto', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', 'Formulario que lo creó (vacío si lo creó un admin)', 0],
      ['NOMBRE', 's', 'Nombre', '', 1],
      ['CARGO', 's', 'Cargo', '', 1],
      ['ROL', 'e', 'Rol', '', 1, 'ROL_CONTACTO'],
      ['TELEFONO_FIJO', 's', 'Teléfono fijo', '', 1],
      ['WHATSAPP', 's', 'WhatsApp', '', 1],
      ['EMAIL', 's', 'Email', '', 1],
      ['ES_PRINCIPAL', 'b', 'Principal', '', 1],
      ['PREFERENCIA_CANAL', 'e', 'Canal preferido', '', 1, 'CANAL'],
      ['HORARIO_CONTACTO', 's', 'Horario de contacto', '', 1],
      ['REDES_SOCIALES', 't', 'Redes sociales', 'Enlaces o usuarios', 1]
    ]
  },

  SEDES: {
    key: 'SEDE_ID', prefix: 'SED', control: true, clientScoped: true,
    desc: 'Sedes, direcciones y horarios.',
    columns: [
      ['SEDE_ID', 's', 'ID sede', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', '', 0],
      ['NOMBRE_SEDE', 's', 'Nombre de la sede', '', 1],
      ['ES_PRINCIPAL', 'b', 'Principal', '', 1],
      ['DIRECCION', 's', 'Dirección', '', 1],
      ['BARRIO', 's', 'Barrio', '', 1],
      ['MUNICIPIO', 'e', 'Municipio', '', 1, 'cat:MUNICIPIOS'],
      ['REFERENCIA', 's', 'Punto de referencia', '', 1],
      ['GOOGLE_MAPS_URL', 's', 'Enlace de Google Maps', '', 1],
      ['TELEFONO', 's', 'Teléfono', '', 1],
      ['WHATSAPP', 's', 'WhatsApp', '', 1],
      ['HORARIO_SEMANA', 's', 'Horario entre semana', '', 1],
      ['HORARIO_FIN_SEMANA', 's', 'Horario fin de semana', '', 1],
      ['HORARIOS_ESPECIALES', 't', 'Horarios especiales', 'Festivos, temporadas…', 1],
      ['ATIENDE_EN_SITIO', 'b', 'Atiende en sitio', 'Sitio comercial abierto al público', 1],
      ['ORDEN', 'n', 'Orden', '', 1]
    ]
  },

  SERVICIOS: {
    key: 'SERVICIO_ID', prefix: 'SRV', control: true, clientScoped: true,
    desc: 'Servicios que ofrece el cliente (se publican en la web).',
    columns: [
      ['SERVICIO_ID', 's', 'ID servicio', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', '', 0],
      ['NOMBRE', 's', 'Nombre', '', 1],
      ['DESCRIPCION', 't', 'Descripción', '', 1],
      ['PRECIO', 's', 'Precio', 'Texto libre editable (D14): "25000", "Desde $20.000", "A convenir"', 1],
      ['MOSTRAR_PRECIO', 'b', 'Mostrar precio', '', 1],
      ['REQUIERE_CITA', 'b', 'Requiere cita', '', 1],
      ['IMAGEN_ARCHIVO_ID', 's', 'Imagen', 'ARCHIVO_ID de la imagen', 0],
      ['ORDEN', 'n', 'Orden', '', 1],
      ['ACTIVO', 'b', 'Activo', '', 1]
    ]
  },

  PRODUCTOS: {
    key: 'PRODUCTO_ID', prefix: 'PRD', control: true, clientScoped: true,
    desc: 'Catálogo de productos del cliente.',
    columns: [
      ['PRODUCTO_ID', 's', 'ID producto', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', '', 0],
      ['NOMBRE', 's', 'Nombre', '', 1],
      ['DESCRIPCION', 't', 'Descripción', '', 1],
      ['CATEGORIA', 's', 'Categoría', '', 1],
      ['PRECIO', 's', 'Costo / precio', 'Texto libre editable (D14)', 1],
      ['MOSTRAR_PRECIO', 'b', 'Mostrar precio', '', 1],
      ['DISPONIBLE', 'b', 'Disponible', '', 1],
      ['IMAGEN_ARCHIVO_ID', 's', 'Imagen', 'ARCHIVO_ID de la imagen', 0],
      ['ORDEN', 'n', 'Orden', '', 1]
    ]
  },

  TESTIMONIOS: {
    key: 'TESTIMONIO_ID', prefix: 'TES', control: true, clientScoped: true,
    desc: 'Testimonios de clientes del negocio.',
    columns: [
      ['TESTIMONIO_ID', 's', 'ID testimonio', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', '', 0],
      ['AUTOR_NOMBRE', 's', 'Autor', '', 1],
      ['TEXTO', 't', 'Texto', 'Máximo 30 palabras', 1],
      ['FOTO_ARCHIVO_ID', 's', 'Foto', 'ARCHIVO_ID de la foto', 0],
      ['PERMISO_USO_FOTO', 'b', 'Permiso uso de foto', '', 1],
      ['PERMISO_USO_NOMBRE', 'b', 'Permiso uso de nombre', '', 1],
      ['APROBADO_ADMIN', 'b', 'Aprobado por admin', 'Solo los aprobados se marcan como publicables', 1]
    ]
  },

  CONTENIDO_WEB: {
    key: 'CONTENIDO_ID', prefix: 'CNT', control: true, clientScoped: true,
    desc: 'Respuestas de la plantilla de página (clave-valor). Claves en CONTENT_KEYS.',
    columns: [
      ['CONTENIDO_ID', 's', 'ID', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', '', 0],
      ['CAMPO_CLAVE', 's', 'Campo', 'Clave de CONTENT_KEYS', 0],
      ['VALOR', 't', 'Valor', '', 1]
    ]
  },

  ARCHIVOS: {
    key: 'ARCHIVO_ID', prefix: 'ARC', control: true, clientScoped: true,
    desc: 'Metadatos de archivos. El archivo vive en Drive, nunca en Sheets.',
    columns: [
      ['ARCHIVO_ID', 's', 'ID archivo', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', 'Se completa al enviar el formulario', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', '', 0],
      ['ENTIDAD', 'e', 'Entidad', 'A qué pertenece', 0, 'ENTIDAD_ARCHIVO'],
      ['ENTIDAD_ID', 's', 'ID entidad', 'SERVICIO_ID, PRODUCTO_ID…', 0],
      ['TEMP_REF', 's', 'Referencia temporal', 'Referencia del ítem mientras el formulario es borrador', 0],
      ['CATEGORIA', 'e', 'Categoría', '', 1, 'ARCHIVO_CATEGORIA'],
      ['NOMBRE_ORIGINAL', 's', 'Nombre original', 'Nombre con que se subió (solo informativo)', 0],
      ['NOMBRE_DRIVE', 's', 'Nombre en Drive', '<CATEGORIA>_<ID>_<fecha>.<ext>', 0],
      ['DRIVE_FILE_ID', 's', 'ID en Drive', 'Privado: nunca se envía al formulario público', 0],
      ['DRIVE_FOLDER_ID', 's', 'Carpeta Drive', '', 0],
      ['MIME_TYPE', 's', 'Tipo MIME', 'Detectado por firma de bytes', 0],
      ['TAMANO_BYTES', 'n', 'Tamaño (bytes)', '', 0],
      ['SHA256', 's', 'Huella SHA-256', 'Detecta archivos repetidos', 0],
      ['ESTADO', 'e', 'Estado', '', 1, 'ARCHIVO_ESTADO'],
      ['MOTIVO_RECHAZO', 's', 'Motivo de rechazo', '', 1]
    ]
  },

  AUTORIZACIONES: {
    key: 'AUTORIZACION_ID', prefix: 'AUT', control: true, clientScoped: true,
    desc: 'Prueba de cada autorización otorgada o negada.',
    columns: [
      ['AUTORIZACION_ID', 's', 'ID', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', '', 0],
      ['TIPO', 'e', 'Tipo', '', 0, 'AUT_TIPO'],
      ['OTORGADA', 'b', 'Otorgada', '', 0],
      ['TEXTO_VERSION', 's', 'Versión del texto', '', 0],
      ['TEXTO_HASH', 's', 'Hash del texto', 'SHA-256 del texto mostrado', 0],
      ['FECHA_HORA', 'dt', 'Fecha y hora', '', 0],
      ['FIRMANTE_NOMBRE', 's', 'Firmante', '', 0],
      ['USER_AGENT', 's', 'Navegador', 'Declarado por el navegador', 0],
      ['REVOCADA_EN', 'dt', 'Revocada en', 'Si el titular revoca', 1]
    ]
  },

  SOLICITUDES: {
    key: 'SOLICITUD_ID', prefix: 'SOL', control: true, clientScoped: true,
    desc: 'Cada formulario diligenciado (borrador o enviado).',
    columns: [
      ['SOLICITUD_ID', 's', 'ID solicitud', '', 0],
      ['RADICADO', 's', 'Radicado', 'WEB-2026-000001 (se asigna al iniciar)', 0],
      ['CLIENTE_ID', 's', 'Cliente', 'Cliente al que quedó asociada al enviar', 0],
      ['CLIENTE_BASE_ID', 's', 'Cliente de la base elegido', 'Registro que el usuario eligió en la búsqueda', 0],
      ['TIPO_VINCULACION', 'e', 'Vinculación', 'Cómo encontró su registro', 0, 'VINCULACION'],
      ['ESTADO', 'e', 'Estado', '', 1, 'SOL_ESTADO'],
      ['PASO_ACTUAL', 'n', 'Paso actual', '', 0],
      ['PORCENTAJE', 'n', 'Porcentaje', '', 0],
      ['TOKEN_HASH', 'x', 'Hash del token', 'SHA-256 del token del navegador (el token no se guarda)', 0],
      ['DATOS_JSON', 'j', 'Datos (borrador)', 'Estado completo del formulario', 0],
      ['DATOS_FILE_ID', 's', 'Datos en Drive', 'Si el borrador supera el tamaño de celda', 0],
      ['CAMBIOS_PROPUESTOS_JSON', 'j', 'Correcciones propuestas', 'Diferencias frente al registro de la base', 0],
      ['CAMBIOS_ESTADO', 'e', 'Estado de correcciones', '', 0, 'CAMBIOS_ESTADO'],
      ['DRIVE_FOLDER_ID', 's', 'Carpeta Drive', 'WEB-2026-000001_NOMBRE', 0],
      ['SNAPSHOT_FILE_ID', 's', 'Snapshot', 'JSON congelado al enviar', 0],
      ['REPORTE_ESTADO', 'e', 'Estado del reporte', '', 0, 'REPORTE_ESTADO'],
      ['INICIADA_EN', 'dt', 'Iniciada en', '', 0],
      ['ENVIADA_EN', 'dt', 'Enviada en', '', 0],
      ['USER_AGENT', 's', 'Navegador', '', 0],
      ['OBSERVACIONES_ADMIN', 't', 'Observaciones internas', 'Nota interna (no sale en el reporte)', 1],
      ['REVISADO_POR', 's', 'Revisado por', '', 0],
      ['REVISADO_EN', 'dt', 'Revisado en', '', 0]
    ]
  },

  REPORTES: {
    key: 'REPORTE_ID', prefix: 'REP', control: true, clientScoped: true,
    desc: 'Reportes editables (Google Docs) por cliente, versionados.',
    columns: [
      ['REPORTE_ID', 's', 'ID reporte', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['SOLICITUD_ID', 's', 'Solicitud', '', 0],
      ['VERSION_REPORTE', 'n', 'Versión', '1, 2, 3…', 0],
      ['DOC_ID', 's', 'ID del Doc', '', 0],
      ['DOC_URL', 's', 'URL del Doc', '', 0],
      ['JSON_FILE_ID', 's', 'JSON de datos', 'Mismos datos en JSON', 0],
      ['GENERADO_POR', 's', 'Generado por', '', 0],
      ['GENERADO_EN', 'dt', 'Generado en', '', 0],
      ['ESTADO', 'e', 'Estado', '', 0, 'REPORTE_ESTADO'],
      ['ERROR', 's', 'Error', '', 0]
    ]
  },

  OPORTUNIDADES: {
    key: 'OPORTUNIDAD_ID', prefix: 'OPO', control: true, clientScoped: true,
    desc: 'Negocios en el pipeline (Kanban).',
    columns: [
      ['OPORTUNIDAD_ID', 's', 'ID oportunidad', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['TITULO', 's', 'Título', '', 1],
      ['PRODUCTO_WEB_ID', 'e', 'Producto web', 'De PRODUCTOS_WEB', 1, 'ref:PRODUCTOS_WEB'],
      ['ETAPA', 'e', 'Etapa', '', 0, 'ETAPA'],
      ['PROBABILIDAD', 'n', 'Probabilidad %', 'Sugerida por la etapa, editable', 1],
      ['VALOR_ESTIMADO', 'm', 'Valor estimado', 'Editable (D14)', 1],
      ['FECHA_CIERRE_ESTIMADA', 'd', 'Cierre estimado', '', 1],
      ['ORIGEN_LEAD', 's', 'Origen del lead', 'FORMULARIO, BASE, REFERIDO…', 1],
      ['RESPONSABLE_EMAIL', 's', 'Responsable', '', 0],
      ['MOTIVO_CIERRE', 't', 'Motivo de cierre', 'Obligatorio en PERDIDO y NO_INTERESADO', 1],
      ['FECHA_CIERRE_REAL', 'd', 'Fecha de cierre', '', 0],
      ['FECHA_CAMBIO_ETAPA', 'dt', 'Último cambio de etapa', '', 0],
      ['ABIERTA', 'b', 'Abierta', '', 0]
    ]
  },

  ACTIVIDADES: {
    key: 'ACTIVIDAD_ID', prefix: 'ACT', control: true, clientScoped: true,
    desc: 'Llamadas, WhatsApp, emails, reuniones y notas internas.',
    columns: [
      ['ACTIVIDAD_ID', 's', 'ID actividad', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['OPORTUNIDAD_ID', 's', 'Oportunidad', '', 0],
      ['TIPO', 'e', 'Tipo', '', 1, 'ACT_TIPO'],
      ['DIRECCION', 'e', 'Dirección', '', 1, 'ACT_DIRECCION'],
      ['RESULTADO', 'e', 'Resultado', '', 1, 'ACT_RESULTADO'],
      ['ASUNTO', 's', 'Asunto', '', 1],
      ['DETALLE', 't', 'Detalle', 'Nota interna del CRM', 1],
      ['DURACION_MIN', 'n', 'Duración (min)', '', 1],
      ['FECHA_HORA', 'dt', 'Fecha y hora', '', 1],
      ['REALIZADA_POR', 's', 'Realizada por', '', 0]
    ]
  },

  SEGUIMIENTOS: {
    key: 'SEGUIMIENTO_ID', prefix: 'SEG', control: true, clientScoped: true,
    desc: 'Tareas con fecha de vencimiento.',
    columns: [
      ['SEGUIMIENTO_ID', 's', 'ID seguimiento', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['OPORTUNIDAD_ID', 's', 'Oportunidad', '', 0],
      ['ACTIVIDAD_ORIGEN_ID', 's', 'Actividad de origen', '', 0],
      ['TIPO', 'e', 'Tipo', '', 1, 'SEG_TIPO'],
      ['DESCRIPCION', 't', 'Descripción', '', 1],
      ['FECHA_VENCIMIENTO', 'd', 'Vence', '', 1],
      ['PRIORIDAD', 'e', 'Prioridad', '', 1, 'PRIORIDAD'],
      ['ESTADO', 'e', 'Estado', 'Vencido se calcula: PENDIENTE con fecha pasada', 0, 'SEG_ESTADO'],
      ['ASIGNADO_A', 's', 'Asignado a', 'Email de usuario admin', 1],
      ['COMPLETADO_EN', 'dt', 'Completado en', '', 0],
      ['RESULTADO', 't', 'Resultado', '', 1]
    ]
  },

  PROPUESTAS: {
    key: 'PROPUESTA_ID', prefix: 'PRO', control: true, clientScoped: true,
    desc: 'Propuestas comerciales con ítems y valores editables.',
    columns: [
      ['PROPUESTA_ID', 's', 'ID propuesta', '', 0],
      ['OPORTUNIDAD_ID', 's', 'Oportunidad', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['NUMERO', 's', 'Número', 'PRO-2026-0001', 0],
      ['ITEMS_JSON', 'j', 'Ítems', '[{descripcion, cantidad, precioUnitario}]', 0],
      ['SUBTOTAL', 'm', 'Subtotal', 'Calculado en el servidor', 0],
      ['DESCUENTO', 'm', 'Descuento', '', 1],
      ['TOTAL', 'm', 'Total', 'Calculado en el servidor', 0],
      ['VALIDEZ_HASTA', 'd', 'Válida hasta', '', 1],
      ['ESTADO', 'e', 'Estado', '', 0, 'PROP_ESTADO'],
      ['NOTAS', 't', 'Notas', '', 1],
      ['ENVIADA_EN', 'dt', 'Enviada en', '', 0],
      ['RESPUESTA_EN', 'dt', 'Respuesta en', '', 0]
    ]
  },

  VENTAS: {
    key: 'VENTA_ID', prefix: 'VEN', control: true, clientScoped: true,
    desc: 'Cierres de venta.',
    columns: [
      ['VENTA_ID', 's', 'ID venta', '', 0],
      ['OPORTUNIDAD_ID', 's', 'Oportunidad', '', 0],
      ['PROPUESTA_ID', 's', 'Propuesta', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['NUMERO', 's', 'Número', 'VEN-2026-0001', 0],
      ['FECHA_VENTA', 'd', 'Fecha de venta', '', 1],
      ['VALOR_TOTAL', 'm', 'Valor total', 'Editable (D14)', 1],
      ['FORMA_PAGO', 'e', 'Forma de pago', '', 1, 'FORMA_PAGO'],
      ['ESTADO_PAGO', 'e', 'Estado de pago', 'Calculado con los pagos', 0, 'ESTADO_PAGO'],
      ['VALOR_PAGADO', 'm', 'Valor pagado', 'Suma de PAGOS', 0],
      ['VENDEDOR_EMAIL', 's', 'Vendedor', '', 0],
      ['NOTAS', 't', 'Notas', '', 1]
    ]
  },

  PAGOS: {
    key: 'PAGO_ID', prefix: 'PAG', control: true, clientScoped: true,
    desc: 'Pagos recibidos por venta.',
    columns: [
      ['PAGO_ID', 's', 'ID pago', '', 0],
      ['VENTA_ID', 's', 'Venta', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['FECHA_PAGO', 'd', 'Fecha de pago', '', 1],
      ['VALOR', 'm', 'Valor', '', 1],
      ['METODO', 'e', 'Método', '', 1, 'METODO_PAGO'],
      ['REFERENCIA', 's', 'Referencia', 'N.º de transacción', 1],
      ['COMPROBANTE_ARCHIVO_ID', 's', 'Comprobante', 'ARCHIVO_ID', 0],
      ['NOTAS', 't', 'Notas', '', 1]
    ]
  },

  PROYECTOS_WEB: {
    key: 'PROYECTO_ID', prefix: 'PRY', control: true, clientScoped: true,
    desc: 'Proyecto de construcción de la página (se crea al registrar la venta).',
    columns: [
      ['PROYECTO_ID', 's', 'ID proyecto', '', 0],
      ['VENTA_ID', 's', 'Venta', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['ESTADO', 'e', 'Estado', '', 1, 'PROY_ESTADO'],
      ['REPORTE_ID', 's', 'Reporte base', 'Reporte usado para construir la página', 0],
      ['DOMINIO', 's', 'Dominio', '', 1],
      ['URL_PUBLICADA', 's', 'URL publicada', '', 1],
      ['FECHA_ENTREGA_ESTIMADA', 'd', 'Entrega estimada', '', 1],
      ['RESPONSABLE_EMAIL', 's', 'Responsable', '', 1],
      ['NOTAS', 't', 'Notas', '', 1]
    ]
  },

  DUPLICADOS: {
    key: 'DUP_ID', prefix: 'DUP', control: true, clientScoped: false,
    desc: 'Posibles duplicados y conciliación formulario ↔ base. Toda fusión es manual.',
    columns: [
      ['DUP_ID', 's', 'ID', '', 0],
      ['CLIENTE_A', 's', 'Cliente A (se conserva)', 'Normalmente el registro de la base', 0],
      ['CLIENTE_B', 's', 'Cliente B', 'Normalmente el creado por el formulario', 0],
      ['REGLA', 's', 'Regla', 'MISMO_EMAIL, MISMO_TELEFONO_Y_NOMBRE…', 0],
      ['PUNTAJE', 'n', 'Puntaje', '0-100', 0],
      ['ESTADO', 'e', 'Estado', '', 0, 'DUP_ESTADO'],
      ['RESUELTO_POR', 's', 'Resuelto por', '', 0],
      ['RESUELTO_EN', 'dt', 'Resuelto en', '', 0]
    ]
  },

  USUARIOS_ADMIN: {
    key: 'USUARIO_ID', prefix: 'USR', control: true, clientScoped: false,
    desc: 'Usuarios del Super Admin. La contraseña se guarda como hash con sal.',
    columns: [
      ['USUARIO_ID', 's', 'ID usuario', '', 0],
      ['EMAIL', 's', 'Email', 'Usuario de acceso', 0],
      ['NOMBRE', 's', 'Nombre', '', 1],
      ['ROL', 'e', 'Rol', '', 0, 'ROL'],
      ['ACTIVO', 'b', 'Activo', '', 0],
      ['PASSWORD_HASH', 'x', 'Hash de contraseña', 'HMAC-SHA256 iterado', 0],
      ['PASSWORD_SALT', 'x', 'Sal', '', 0],
      ['PASSWORD_ITER', 'x', 'Iteraciones', '', 0],
      ['DEBE_CAMBIAR_PASSWORD', 'b', 'Debe cambiar contraseña', '', 0],
      ['INTENTOS_FALLIDOS', 'n', 'Intentos fallidos', '', 0],
      ['BLOQUEADO_HASTA', 'dt', 'Bloqueado hasta', '', 0],
      ['ULTIMO_ACCESO', 'dt', 'Último acceso', '', 0]
    ]
  },

  CONFIGURACION: {
    key: 'CLAVE', prefix: '', control: false, clientScoped: false,
    desc: 'Ajustes editables desde el panel (EDITABLE_SETTINGS).',
    columns: [
      ['CLAVE', 's', 'Clave', '', 0],
      ['VALOR', 't', 'Valor', '', 1],
      ['DESCRIPCION', 's', 'Descripción', '', 0],
      ['ACTUALIZADO_EN', 'dt', 'Actualizado en', '', 0],
      ['ACTUALIZADO_POR', 's', 'Actualizado por', '', 0]
    ]
  },

  CATALOGOS: {
    key: 'ITEM_ID', prefix: 'CAT', control: true, clientScoped: false,
    desc: 'Listas con autocompletado: MUNICIPIOS, TIPOS_ESTABLECIMIENTO, MOTIVOS_CIERRE.',
    columns: [
      ['ITEM_ID', 's', 'ID', '', 0],
      ['CATALOGO', 's', 'Catálogo', '', 0],
      ['CODIGO', 's', 'Código', '', 1],
      ['ETIQUETA', 's', 'Etiqueta', '', 1],
      ['ORDEN', 'n', 'Orden', '', 1],
      ['ACTIVO', 'b', 'Activo', '', 1]
    ]
  },

  PRODUCTOS_WEB: {
    key: 'PRODUCTO_WEB_ID', prefix: 'PWB', control: true, clientScoped: false,
    desc: 'Lo que vende Webpaya. Precio sugerido opcional y editable (D14).',
    columns: [
      ['PRODUCTO_WEB_ID', 's', 'ID', '', 0],
      ['NOMBRE', 's', 'Nombre', '', 1],
      ['DESCRIPCION', 't', 'Descripción', '', 1],
      ['PRECIO_SUGERIDO', 'm', 'Precio sugerido', 'Opcional; se precarga y se edita en cada propuesta', 1],
      ['ES_ADICIONAL', 'b', 'Es adicional', 'Agenda de citas, base de datos de clientes…', 1],
      ['ACTIVO', 'b', 'Activo', '', 1],
      ['ORDEN', 'n', 'Orden', '', 1]
    ]
  },

  SECUENCIAS: {
    key: 'NOMBRE', prefix: '', control: false, clientScoped: false,
    desc: 'Contadores para radicados y números consecutivos.',
    columns: [
      ['NOMBRE', 's', 'Nombre', '', 0],
      ['ULTIMO_VALOR', 'n', 'Último valor', '', 0]
    ]
  },

  AUDITORIA: {
    key: 'AUDIT_ID', prefix: 'AUD', control: false, clientScoped: false, auditDb: true,
    desc: 'Historial inmutable (spreadsheet separado). Solo se agregan filas.',
    columns: [
      ['AUDIT_ID', 's', 'ID', '', 0],
      ['FECHA_HORA', 'dt', 'Fecha y hora', '', 0],
      ['ACTOR', 's', 'Actor', 'PUBLICO:<radicado> · ADMIN:<email> · SISTEMA:<proceso>', 0],
      ['ROL', 's', 'Rol', '', 0],
      ['APP', 's', 'Aplicación', 'PUBLICA, ADMIN, SISTEMA', 0],
      ['ACCION', 's', 'Acción', 'CREAR, EDITAR, CAMBIO_ETAPA, ACTIVIDAD, SEGUIMIENTO, VENTA, LOGIN…', 0],
      ['ENTIDAD', 's', 'Entidad', 'Tabla afectada', 0],
      ['ENTIDAD_ID', 's', 'ID', '', 0],
      ['CLIENTE_ID', 's', 'Cliente', '', 0],
      ['CAMPOS', 's', 'Campos', 'Campos modificados', 0],
      ['ANTES_JSON', 'j', 'Antes', 'Valores anteriores', 0],
      ['DESPUES_JSON', 'j', 'Después', 'Valores nuevos', 0],
      ['RESULTADO', 's', 'Resultado', 'OK, DENEGADO, ERROR', 0],
      ['DETALLE', 't', 'Detalle', '', 0]
    ]
  }
};

/** Claves del contenido web (CONTENIDO_WEB), agrupadas por sección de página. */
const CONTENT_KEYS = Object.freeze({
  nombre_pagina_1: { label: 'Nombre de página deseado (opción 1)', section: 'FICHA' },
  nombre_pagina_2: { label: 'Nombre de página deseado (opción 2)', section: 'FICHA' },
  tipo_establecimiento_alt: { label: 'Tipo de establecimiento (opción 2)', section: 'QUIENES' },
  mensaje: { label: '¿Qué le gustaría transmitir en su página web?', section: 'PORTADA' },
  publico_objetivo: { label: 'Público objetivo', section: 'QUIENES' },
  diferenciales: { label: 'Qué lo hace diferente', section: 'QUIENES' },
  colores: { label: 'Colores de marca', section: 'MARCA' },
  referencias: { label: 'Páginas de referencia que le gustan', section: 'MARCA' },
  dominio_actual: { label: 'Dominio o página actual', section: 'MARCA' },
  domicilios_ofrece: { label: 'Ofrece domicilios', section: 'CATALOGO' },
  domicilios_costo: { label: 'Costo del domicilio', section: 'CATALOGO' },
  agenda_citas: { label: 'Desea agenda de citas', section: 'ADICIONALES' },
  agenda_version_pago: { label: 'Agenda de citas: versión de pago', section: 'ADICIONALES' },
  bd_clientes: { label: 'Desea base de datos de clientes', section: 'ADICIONALES' },
  bd_version_pago: { label: 'Base de datos de clientes: versión de pago', section: 'ADICIONALES' }
});

/** Utilidades sobre el esquema. */
const SchemaUtil = {
  table(name) {
    const t = SCHEMA[name];
    if (!t) throw new AppError('BAD_TABLE', 'Tabla desconocida: ' + name);
    return t;
  },
  /** Columnas completas (incluye control) como objetos. */
  columns(name) {
    const t = SchemaUtil.table(name);
    const cols = t.control ? t.columns.concat(CONTROL_COLUMNS) : t.columns;
    return cols.map(c => ({ name: c[0], type: c[1], label: c[2], desc: c[3], editable: !!c[4], enumKey: c[5] || '' }));
  },
  columnNames(name) {
    return SchemaUtil.columns(name).map(c => c.name);
  },
  column(name, col) {
    return SchemaUtil.columns(name).find(c => c.name === col) || null;
  },
  editableColumns(name) {
    return SchemaUtil.columns(name).filter(c => c.editable).map(c => c.name);
  },
  secretColumns(name) {
    return SchemaUtil.columns(name).filter(c => c.type === 'x').map(c => c.name);
  },
  tables() {
    return Object.keys(SCHEMA);
  },
  clientScopedTables() {
    return Object.keys(SCHEMA).filter(t => SCHEMA[t].clientScoped);
  },
  /** Metadatos para formularios genéricos del admin (sin columnas secretas). */
  uiMeta(name) {
    return SchemaUtil.columns(name).filter(c => c.type !== 'x').map(c => ({
      name: c.name, type: c.type, label: c.label, editable: c.editable, enumKey: c.enumKey
    }));
  }
};

/**
 * Config.gs — Configuración central.
 *
 * Tres niveles:
 *  1. CONFIG (este archivo): valores por defecto y constantes del sistema.
 *  2. Script Properties (Proyecto > Configuración > Propiedades del script): IDs
 *     privados (spreadsheets, carpetas) y secretos. Los crea setup() automáticamente.
 *     NUNCA se envían al navegador.
 *  3. Hoja CONFIGURACION: ajustes editables por el Super Admin desde el panel
 *     (solo las claves listadas en EDITABLE_SETTINGS).
 */

const CONFIG = Object.freeze({
  APPLICATION_NAME: 'Webpaya · Información para su página web',
  ADMIN_APPLICATION_NAME: 'Webpaya CRM',
  RESPONSABLE_TRATAMIENTO: 'Equipo Webpaya',
  TIMEZONE: 'America/Bogota',

  // Nombres de los recursos que crea setup() si no existen.
  ROOT_FOLDER_NAME: 'CLIENTES_WEB',
  SYSTEM_FOLDER_NAME: 'CRM_WEB_SISTEMA',
  DB_SPREADSHEET_NAME: 'CRM_WEB_DB',
  AUDIT_SPREADSHEET_NAME: 'CRM_WEB_AUDITORIA',

  // Archivos
  MAX_FILE_SIZE_MB: 10,
  MAX_FILES_PER_SOLICITUD: 40,
  MAX_TOTAL_MB_SOLICITUD: 100,
  PREVIEW_MAX_BYTES: 1500000,

  // Límites del formulario
  MAX_SEDES: 10,
  MAX_SERVICIOS: 30,
  MAX_PRODUCTOS: 100,
  MAX_TESTIMONIOS: 20,
  TESTIMONIO_MAX_PALABRAS: 30,
  TEXT_SHORT_MAX: 200,
  TEXT_LONG_MAX: 3000,
  DRAFT_JSON_MAX_CHARS: 45000,

  // Sesiones y autenticación de administradores
  SESSION_TTL_SEC: 21600,          // 6 h (máximo de CacheService)
  SESSION_IDLE_SEC: 3600,          // 60 min sin actividad cierra sesión
  LOGIN_MAX_ATTEMPTS: 5,
  LOGIN_LOCK_MINUTES: 15,
  PASSWORD_MIN_LENGTH: 10,
  PASSWORD_ITERATIONS: 3000,
  TWO_FA_TTL_SEC: 600,
  SETUP_CODE_TTL_HOURS: 24,

  // Límites anti-abuso del formulario público (CacheService)
  RATE_DRAFTS_PER_10MIN: 60,
  RATE_SEARCH_PER_DRAFT: 25,
  RATE_SEARCH_PER_MIN_GLOBAL: 120,
  RATE_UPLOADS_PER_DRAFT: 60,
  MIN_SECONDS_BEFORE_SUBMIT: 20,

  // Búsqueda
  SEARCH_MAX_RESULTS: 5,
  SEARCH_MIN_NAME_CHARS: 4,
  INDEX_CHUNK_CHARS: 90000,
  INDEX_TTL_SEC: 21600,

  // Importación de la base existente
  IMPORT_SOURCE_SHEET: 'Hoja1',
  IMPORT_BATCH_SIZE: 2500,
  IMPORT_MAX_MS: 270000,

  // Tipos de archivo permitidos: extensión -> MIME aceptados
  ALLOWED_UPLOADS: {
    jpg: ['image/jpeg'],
    jpeg: ['image/jpeg'],
    png: ['image/png'],
    webp: ['image/webp'],
    pdf: ['application/pdf']
  },
  // Categorías que aceptan PDF (el resto solo imágenes)
  PDF_CATEGORIES: ['DOCUMENTO', 'OTRO', 'COMPROBANTE', 'PROPUESTA'],

  // Estructura de carpetas por cliente (Parte E)
  CLIENT_SUBFOLDERS: [
    '01_IDENTIDAD_VISUAL',
    '02_EQUIPO',
    '03_PRODUCTOS',
    '04_TESTIMONIOS',
    '05_DOCUMENTOS',
    '06_OTROS'
  ],
  CATEGORY_FOLDER: {
    LOGO: '01_IDENTIDAD_VISUAL',
    FOTO_NEGOCIO: '01_IDENTIDAD_VISUAL',
    FOTO_EQUIPO: '02_EQUIPO',
    FOTO_PRODUCTO: '03_PRODUCTOS',
    FOTO_SERVICIO: '03_PRODUCTOS',
    FOTO_TESTIMONIO: '04_TESTIMONIOS',
    DOCUMENTO: '05_DOCUMENTOS',
    REPORTE: '05_DOCUMENTOS',
    SNAPSHOT: '05_DOCUMENTOS',
    COMPROBANTE: '05_DOCUMENTOS',
    PROPUESTA: '05_DOCUMENTOS',
    OTRO: '06_OTROS'
  },

  // Municipios permitidos (decisión D9)
  MUNICIPIOS: [
    { code: '41001', name: 'NEIVA' },
    { code: '41551', name: 'PITALITO' },
    { code: '41524', name: 'PALERMO' },
    { code: '41615', name: 'RIVERA' }
  ],

  AUTORIZACION_VERSION: 'v2026-09'
});

/**
 * Ajustes que el Super Admin puede cambiar desde Configuración.
 * type: s=texto, n=número, b=sí/no, list=lista separada por comas.
 */
const EDITABLE_SETTINGS = Object.freeze({
  APP_NAME: { type: 's', def: CONFIG.APPLICATION_NAME, label: 'Nombre del formulario público' },
  RESPONSABLE_TRATAMIENTO: { type: 's', def: CONFIG.RESPONSABLE_TRATAMIENTO, label: 'Responsable del tratamiento de datos' },
  CONTACTO_PRIVACIDAD: { type: 's', def: '', label: 'Canal para consultas y reclamos de datos (WhatsApp o email)' },
  MAX_FILE_SIZE_MB: { type: 'n', def: CONFIG.MAX_FILE_SIZE_MB, label: 'Tamaño máximo por archivo (MB, máx. 20)' },
  REQUIRED_FIELDS: {
    type: 'list',
    def: 'razonSocial,medioContacto,TRATAMIENTO_DATOS',
    label: 'Campos obligatorios del formulario (razonSocial, medioContacto, municipio, direccion, email, nombrePropietario, TRATAMIENTO_DATOS)'
  },
  ADMIN_2FA: { type: 'b', def: false, label: 'Exigir código por email al iniciar sesión (2FA de administradores)' },
  REPORT_EDITORS: { type: 'list', def: '', label: 'Cuentas Google con permiso de edición sobre los reportes (emails separados por coma)' },
  DEFAULT_ASSIGNEE: { type: 's', def: '', label: 'Responsable por defecto de nuevas solicitudes (email de usuario admin)' },
  WHATSAPP_PLANTILLA: {
    type: 's',
    def: 'Hola {NOMBRE}, somos Webpaya. Queremos ayudarle a tener su página web. Puede enviarnos su información aquí: {URL}',
    label: 'Mensaje de WhatsApp para invitar prospectos ({NOMBRE}, {URL})'
  },
  PUBLIC_FORM_URL: { type: 's', def: '', label: 'URL pública del formulario (se detecta sola si se deja vacía)' }
});

/** Textos de autorización (versionados; se guarda el hash del texto aceptado). */
const AUTH_TEXTS = Object.freeze([
  {
    tipo: 'TRATAMIENTO_DATOS',
    titulo: 'Tratamiento de datos personales',
    obligatoria: true,
    texto: 'Autorizo a {RESPONSABLE} para recolectar, almacenar, usar y actualizar los datos que suministro en este formulario, ' +
      'con la finalidad de elaborar y gestionar mi página web y contactarme en relación con ese servicio, conforme a la ' +
      'Ley 1581 de 2012 y el Decreto 1377 de 2013. Conozco mi derecho a conocer, actualizar, rectificar y suprimir mis datos ' +
      'y a revocar esta autorización.'
  },
  {
    tipo: 'USO_IMAGENES',
    titulo: 'Uso de imágenes',
    obligatoria: false,
    texto: 'Autorizo el uso del logo, fotografías y demás imágenes que cargo, para publicarlas en mi página web.'
  },
  {
    tipo: 'PUBLICACION_CONTENIDO',
    titulo: 'Publicación del contenido',
    obligatoria: false,
    texto: 'Autorizo publicar en mi página web los textos, servicios, productos, precios, horarios y datos de contacto que suministro.'
  },
  {
    tipo: 'USO_FOTOS_TESTIMONIOS',
    titulo: 'Testimonios de clientes',
    obligatoria: false,
    texto: 'Declaro que cuento con el permiso de las personas cuyos testimonios, nombres o fotografías registro, para publicarlos.'
  },
  {
    tipo: 'CONTACTO_WHATSAPP',
    titulo: 'Contacto por WhatsApp',
    obligatoria: false,
    texto: 'Acepto ser contactado por WhatsApp en relación con mi página web.'
  },
  {
    tipo: 'CONTACTO_LLAMADA',
    titulo: 'Contacto telefónico',
    obligatoria: false,
    texto: 'Acepto ser contactado por llamada telefónica en relación con mi página web.'
  },
  {
    tipo: 'CONTACTO_EMAIL',
    titulo: 'Contacto por email',
    obligatoria: false,
    texto: 'Acepto ser contactado por correo electrónico en relación con mi página web.'
  },
  {
    tipo: 'USO_RAZON_SOCIAL_BASE',
    titulo: 'Uso de la razón social',
    obligatoria: false,
    texto: 'Permito el uso de mi razón social para alimentar la base de datos de {RESPONSABLE}.'
  }
]);

/** Acceso a Script Properties (valores privados). */
const Props = {
  get(key) {
    return PropertiesService.getScriptProperties().getProperty(key) || '';
  },
  set(key, value) {
    PropertiesService.getScriptProperties().setProperty(key, String(value));
  },
  remove(key) {
    PropertiesService.getScriptProperties().deleteProperty(key);
  },
  require(key) {
    const v = Props.get(key);
    if (!v) {
      throw new AppError('NOT_CONFIGURED', 'El sistema no está configurado. El propietario debe ejecutar setup().');
    }
    return v;
  }
};

/** Ajustes editables (hoja CONFIGURACION) con caché de 5 minutos. */
const Settings = {
  _mem: null,
  _load() {
    if (Settings._mem) return Settings._mem;
    const cache = CacheService.getScriptCache();
    const cached = cache.get('settings:v1');
    if (cached) {
      Settings._mem = JSON.parse(cached);
      return Settings._mem;
    }
    const map = {};
    try {
      SheetService.readAll('CONFIGURACION', { includeDeleted: true }).forEach(r => {
        map[r.CLAVE] = r.VALOR;
      });
    } catch (e) {
      // Antes de setup() la hoja no existe: se usan los valores por defecto.
    }
    cache.put('settings:v1', JSON.stringify(map), 300);
    Settings._mem = map;
    return map;
  },
  get(key) {
    const def = EDITABLE_SETTINGS[key];
    if (!def) throw new AppError('BAD_SETTING', 'Ajuste desconocido: ' + key);
    const raw = Settings._load()[key];
    const value = (raw === undefined || raw === null || raw === '') ? def.def : raw;
    if (def.type === 'n') return Number(value) || Number(def.def);
    if (def.type === 'b') return Utils.toBool(value);
    if (def.type === 'list') {
      return String(value || '').split(',').map(s => s.trim()).filter(Boolean);
    }
    return String(value);
  },
  all() {
    const out = {};
    Object.keys(EDITABLE_SETTINGS).forEach(k => {
      const v = Settings.get(k);
      out[k] = { value: Array.isArray(v) ? v.join(',') : v, label: EDITABLE_SETTINGS[k].label, type: EDITABLE_SETTINGS[k].type };
    });
    return out;
  },
  invalidate() {
    Settings._mem = null;
    CacheService.getScriptCache().remove('settings:v1');
  },
  maxFileBytes() {
    const mb = Math.min(Math.max(Settings.get('MAX_FILE_SIZE_MB'), 1), 20);
    return mb * 1024 * 1024;
  }
};

/** Texto de una autorización con el responsable ya reemplazado. */
function authTextFor_(tipo) {
  const def = AUTH_TEXTS.find(a => a.tipo === tipo);
  if (!def) return '';
  return def.texto.replace(/\{RESPONSABLE\}/g, Settings.get('RESPONSABLE_TRATAMIENTO'));
}

/**
 * DriveService.gs — Carpetas y archivos de clientes en Google Drive (Partes E e I).
 *
 * CLIENTES_WEB/
 *   WEB-2026-000001_NOMBRE/          (o CLI-000123_NOMBRE si lo crea un admin)
 *     01_IDENTIDAD_VISUAL/ 02_EQUIPO/ 03_PRODUCTOS/ 04_TESTIMONIOS/ 05_DOCUMENTOS/ 06_OTROS/
 *
 * Nada se comparte públicamente. Los archivos quedan asociados al cliente por
 * ARCHIVOS.CLIENTE_ID; el nombre de la carpeta es solo informativo.
 */
const DriveService = {
  root() {
    return DriveApp.getFolderById(Props.require('ROOT_FOLDER_ID'));
  },

  /** Crea la carpeta del cliente/solicitud con sus 6 subcarpetas. */
  createClientFolder(name) {
    const folder = DriveService.root().createFolder(Utils.truncate(name, 120));
    CONFIG.CLIENT_SUBFOLDERS.forEach(sub => folder.createFolder(sub));
    return folder;
  },

  /** Devuelve la carpeta existente o crea una nueva. */
  ensureFolder(existingId, name) {
    if (existingId) {
      try {
        const f = DriveApp.getFolderById(existingId);
        if (!f.isTrashed()) return f;
      } catch (e) {
        console.warn('Carpeta no disponible, se crea una nueva', existingId);
      }
    }
    return DriveService.createClientFolder(name);
  },

  rename(folderId, name) {
    try {
      DriveApp.getFolderById(folderId).setName(Utils.truncate(name, 120));
    } catch (e) {
      console.warn('No se pudo renombrar la carpeta', folderId, e && e.message);
    }
  },

  /** Subcarpeta por nombre (se crea si falta). Caché de 6 h. */
  subfolder(parentId, name) {
    const cache = CacheService.getScriptCache();
    const key = 'sf:' + parentId + ':' + name;
    const cached = cache.get(key);
    if (cached) {
      try { return DriveApp.getFolderById(cached); } catch (e) { cache.remove(key); }
    }
    const parent = DriveApp.getFolderById(parentId);
    const it = parent.getFoldersByName(name);
    const folder = it.hasNext() ? it.next() : parent.createFolder(name);
    cache.put(key, folder.getId(), 21600);
    return folder;
  },

  folderForCategory(parentId, category) {
    return DriveService.subfolder(parentId, CONFIG.CATEGORY_FOLDER[category] || '06_OTROS');
  },

  // ---------- Validación de archivos ----------
  /** Tipo real por "firma mágica" de los primeros bytes. */
  detectMime(bytes) {
    const b = (i) => (bytes[i] === undefined ? -1 : (bytes[i] & 0xff));
    if (b(0) === 0xFF && b(1) === 0xD8 && b(2) === 0xFF) return 'image/jpeg';
    if (b(0) === 0x89 && b(1) === 0x50 && b(2) === 0x4E && b(3) === 0x47 && b(4) === 0x0D && b(5) === 0x0A) return 'image/png';
    if (b(0) === 0x52 && b(1) === 0x49 && b(2) === 0x46 && b(3) === 0x46 &&
        b(8) === 0x57 && b(9) === 0x45 && b(10) === 0x42 && b(11) === 0x50) return 'image/webp';
    if (b(0) === 0x25 && b(1) === 0x50 && b(2) === 0x44 && b(3) === 0x46) return 'application/pdf';
    return '';
  },

  /**
   * Valida extensión, MIME declarado, MIME real y tamaño.
   * @return {{ext:string, mime:string}}
   */
  validateUpload(fileName, declaredMime, category, bytes) {
    const name = String(fileName || '');
    const m = name.toLowerCase().match(/\.([a-z0-9]{1,5})$/);
    const ext = m ? m[1] : '';
    const allowed = CONFIG.ALLOWED_UPLOADS[ext];
    if (!allowed) {
      throw new AppError('BAD_FILE', 'Tipo de archivo no permitido. Use JPG, PNG, WEBP' +
        (CONFIG.PDF_CATEGORIES.indexOf(category) >= 0 ? ' o PDF.' : '.'));
    }
    if (ext === 'pdf' && CONFIG.PDF_CATEGORIES.indexOf(category) < 0) {
      throw new AppError('BAD_FILE', 'En esta sección solo se aceptan imágenes (JPG, PNG o WEBP).');
    }
    if (declaredMime && allowed.indexOf(String(declaredMime).toLowerCase()) < 0) {
      throw new AppError('BAD_FILE', 'El tipo del archivo no coincide con su extensión.');
    }
    if (!bytes || !bytes.length) throw new AppError('BAD_FILE', 'El archivo está vacío.');
    const max = Settings.maxFileBytes();
    if (bytes.length > max) {
      throw new AppError('FILE_TOO_LARGE', 'El archivo supera el máximo de ' + Math.round(max / 1048576) + ' MB.');
    }
    const real = DriveService.detectMime(bytes);
    if (!real || allowed.indexOf(real) < 0) {
      throw new AppError('BAD_FILE', 'El contenido del archivo no corresponde a una imagen o PDF válido.');
    }
    return { ext: ext === 'jpeg' ? 'jpg' : ext, mime: real };
  },

  /** Guarda bytes en la subcarpeta de la categoría. */
  saveFile(parentFolderId, category, driveName, bytes, mime) {
    const folder = DriveService.folderForCategory(parentFolderId, category);
    const blob = Utilities.newBlob(bytes, mime, driveName);
    return folder.createFile(blob);
  },

  /** Vista previa como data URL (solo imágenes; PDF usa miniatura de Drive). */
  preview(fileId, mime, size) {
    try {
      const file = DriveApp.getFileById(fileId);
      if (mime && mime.indexOf('image/') === 0 && Number(size) <= CONFIG.PREVIEW_MAX_BYTES) {
        const blob = file.getBlob();
        return 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
      }
      const thumb = file.getThumbnail();
      if (thumb) return 'data:image/png;base64,' + Utilities.base64Encode(thumb.getBytes());
    } catch (e) {
      console.warn('Sin vista previa', fileId, e && e.message);
    }
    return '';
  },

  download(fileId) {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return { name: file.getName(), mime: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) };
  },

  trash(fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (e) {
      console.warn('No se pudo enviar a la papelera', fileId, e && e.message);
    }
  },

  // ---------- JSON en Drive ----------
  writeJson(parentFolderId, name, obj) {
    const folder = DriveService.folderForCategory(parentFolderId, 'SNAPSHOT');
    const file = folder.createFile(Utilities.newBlob(JSON.stringify(obj, null, 2), 'application/json', name));
    return file.getId();
  },
  updateJson(fileId, obj) {
    DriveApp.getFileById(fileId).setContent(JSON.stringify(obj));
  },
  readJson(fileId) {
    return Utils.safeJsonParse(DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8'), null);
  },

  folderUrl(id) {
    return id ? 'https://drive.google.com/drive/folders/' + id : '';
  },
  fileUrl(id) {
    return id ? 'https://drive.google.com/file/d/' + id + '/view' : '';
  },

  /** Revisa que CLIENTES_WEB y las hojas no estén compartidas públicamente. */
  auditSharing() {
    const problems = [];
    const check = (label, getter) => {
      try {
        const item = getter();
        const access = item.getSharingAccess();
        if (access === DriveApp.Access.ANYONE || access === DriveApp.Access.ANYONE_WITH_LINK) problems.push(label);
      } catch (e) {
        problems.push(label + ' (no accesible)');
      }
    };
    check('Carpeta CLIENTES_WEB', () => DriveApp.getFolderById(Props.require('ROOT_FOLDER_ID')));
    check('Hoja CRM_WEB_DB', () => DriveApp.getFileById(Props.require('DB_SPREADSHEET_ID')));
    check('Hoja CRM_WEB_AUDITORIA', () => DriveApp.getFileById(Props.require('AUDIT_SPREADSHEET_ID')));
    return problems;
  }
};

/**
 * AdminApi.gs — FUNCIONES ADMINISTRATIVAS.
 * Cada una valida en el servidor: token de sesión + permiso del rol (+ acceso al
 * cliente dentro del servicio). El navegador nunca decide permisos.
 */

/** Envoltura: autentica, autoriza y ejecuta. */
function admin_(token, permission, fn) {
  return respond_(() => {
    const user = Auth.requireAdmin(token, permission);
    return fn(user);
  });
}

// ---------------- Sesión ----------------
function adm_setupStatus() {
  return respond_(() => Auth.setupStatus());
}
function adm_createSuperAdmin(code, email, nombre, password) {
  return respond_(() => Auth.createSuperAdminWithCode(code, email, nombre, password));
}
function adm_login(email, password, userAgent) {
  return respond_(() => Auth.login(email, password, userAgent));
}
function adm_verify2fa(challenge, code, userAgent) {
  return respond_(() => Auth.verifyTwoFactor(challenge, code, userAgent));
}
function adm_logout(token) {
  return respond_(() => { Auth.logout(token); return true; });
}
function adm_me(token) {
  return admin_(token, 'self', user => ({
    user: user,
    meta: {
      etapas: ENUMS.ETAPA.map(e => ({ key: e, label: STAGES[e].label, prob: STAGES[e].prob, closed: STAGES[e].closed })),
      enums: ENUMS,
      schema: {
        CLIENTES: SchemaUtil.uiMeta('CLIENTES'), CONTACTOS: SchemaUtil.uiMeta('CONTACTOS'), SEDES: SchemaUtil.uiMeta('SEDES'),
        SERVICIOS: SchemaUtil.uiMeta('SERVICIOS'), PRODUCTOS: SchemaUtil.uiMeta('PRODUCTOS'), TESTIMONIOS: SchemaUtil.uiMeta('TESTIMONIOS'),
        CONTENIDO_WEB: SchemaUtil.uiMeta('CONTENIDO_WEB'), OPORTUNIDADES: SchemaUtil.uiMeta('OPORTUNIDADES'),
        ACTIVIDADES: SchemaUtil.uiMeta('ACTIVIDADES'), SEGUIMIENTOS: SchemaUtil.uiMeta('SEGUIMIENTOS'),
        VENTAS: SchemaUtil.uiMeta('VENTAS'), PAGOS: SchemaUtil.uiMeta('PAGOS'), PROYECTOS_WEB: SchemaUtil.uiMeta('PROYECTOS_WEB'),
        PRODUCTOS_WEB: SchemaUtil.uiMeta('PRODUCTOS_WEB')
      },
      contentKeys: CONTENT_KEYS,
      municipios: CONFIG.MUNICIPIOS.map(m => m.name),
      tiposEstablecimiento: CatalogService.list('TIPOS_ESTABLECIMIENTO').map(t => t.ETIQUETA),
      motivosCierre: CatalogService.list('MOTIVOS_CIERRE').map(t => t.ETIQUETA),
      productosWeb: ProductoWebService.list(false).map(p => ({ id: p.PRODUCTO_WEB_ID, nombre: p.NOMBRE, precio: p.PRECIO_SUGERIDO })),
      usuarios: UserService.assignable(),
      today: Utils.today()
    }
  }));
}
function adm_changePassword(token, current, next) {
  return admin_(token, 'self', user => Auth.changePassword(user, current, next));
}

// ---------------- Dashboard ----------------
function adm_dashboard(token) {
  return admin_(token, 'dashboard', user => DashboardService.summary(user));
}

// ---------------- Clientes ----------------
function adm_searchClients(token, params) {
  return admin_(token, 'clients.read', user => {
    const r = SearchIndex.adminQuery(user, params || {});
    delete r.all;
    return r;
  });
}
function adm_getClient(token, id) {
  return admin_(token, 'clients.read', user => ClientService.getFull(user, String(id || '')));
}
function adm_createClient(token, data) {
  return admin_(token, 'clients.create', user => ClientService.create(user, data));
}
function adm_updateClient(token, id, patch, version) {
  return admin_(token, 'clients.write', user => ClientService.update(user, String(id || ''), patch, version));
}
function adm_assignClient(token, id, email) {
  return admin_(token, 'clients.write', user => ClientService.assign(user, String(id || ''), email));
}
function adm_saveChild(token, table, record) {
  return admin_(token, 'clients.write', user => ClientService.saveChild(user, String(table || ''), record));
}
function adm_deleteChild(token, table, id) {
  return admin_(token, 'clients.write', user => ClientService.deleteChild(user, String(table || ''), String(id || '')));
}
function adm_exportClients(token, params) {
  return admin_(token, 'export', user => ExportService.clientsCsv(user, params || {}));
}
function adm_whatsappLink(token, clienteId) {
  return admin_(token, 'crm.write', user => CRMService.whatsappLink(user, String(clienteId || '')));
}

// ---------------- Archivos ----------------
function adm_uploadFile(token, clienteId, meta, base64) {
  return admin_(token, 'files.write', user => ClientService.uploadFile(user, String(clienteId || ''), meta, base64));
}
function adm_filePreview(token, archivoId) {
  return admin_(token, 'files.read', user => ClientService.filePreview(user, String(archivoId || '')));
}
function adm_fileDownload(token, archivoId) {
  return admin_(token, 'files.read', user => ClientService.fileDownload(user, String(archivoId || '')));
}
function adm_setFileStatus(token, archivoId, estado, motivo) {
  return admin_(token, 'files.write', user => ClientService.setFileStatus(user, String(archivoId || ''), estado, motivo));
}
function adm_pendingFiles(token) {
  return admin_(token, 'files.read', user => ClientService.pendingFiles(user));
}

// ---------------- Solicitudes y reportes ----------------
function adm_listSolicitudes(token, filters) {
  return admin_(token, 'solicitudes.read', user => CRMService.listSolicitudes(user, filters));
}
function adm_getSolicitud(token, id) {
  return admin_(token, 'solicitudes.read', user => CRMService.getSolicitud(user, String(id || '')));
}
function adm_applyCorrections(token, id, campos, descartar) {
  return admin_(token, 'clients.write', user => CRMService.applyCorrections(user, String(id || ''), campos, !!descartar));
}
function adm_setSolicitudEstado(token, id, estado, observaciones) {
  return admin_(token, 'clients.write', user => CRMService.setSolicitudEstado(user, String(id || ''), estado, observaciones));
}
function adm_generateReport(token, clienteId, force) {
  return admin_(token, 'reports', user => ReportService.generate(user, String(clienteId || ''), '', !!force));
}
function adm_reportPdf(token, reporteId) {
  return admin_(token, 'reports', user => ReportService.pdf(user, String(reporteId || '')));
}
function adm_reportJson(token, reporteId) {
  return admin_(token, 'reports', user => ReportService.json(user, String(reporteId || '')));
}

// ---------------- Duplicados ----------------
function adm_listDuplicates(token, estado) {
  return admin_(token, 'duplicates', user => ClientService.listDuplicates(user, estado));
}
function adm_compareDuplicate(token, dupId) {
  return admin_(token, 'duplicates', user => ClientService.compareDuplicate(user, String(dupId || '')));
}
function adm_resolveDuplicate(token, dupId, decision, choices) {
  return admin_(token, 'duplicates', user => ClientService.resolveDuplicate(user, String(dupId || ''), decision, choices));
}

// ---------------- CRM ----------------
function adm_listOpportunities(token, filters) {
  return admin_(token, 'crm.read', user => CRMService.listOpportunities(user, filters));
}
function adm_createOpportunity(token, data) {
  return admin_(token, 'crm.write', user => CRMService.createOpportunity(user, data));
}
function adm_updateOpportunity(token, id, patch, version) {
  return admin_(token, 'crm.write', user => CRMService.updateOpportunity(user, String(id || ''), patch, version));
}
function adm_moveOpportunity(token, id, etapa, version, extra) {
  return admin_(token, 'crm.write', user => CRMService.moveOpportunity(user, String(id || ''), etapa, version, extra));
}
function adm_setClientStage(token, clienteId, etapa, extra) {
  return admin_(token, 'crm.write', user => CRMService.setClientStage(user, String(clienteId || ''), etapa, extra));
}
function adm_logActivity(token, data) {
  return admin_(token, 'crm.write', user => CRMService.logActivity(user, data));
}
function adm_listActivities(token, filters) {
  return admin_(token, 'crm.read', user => CRMService.listActivities(user, filters));
}
function adm_saveFollowup(token, data) {
  return admin_(token, 'crm.write', user => CRMService.saveFollowup(user, data));
}
function adm_completeFollowup(token, id, resultado, estado) {
  return admin_(token, 'crm.write', user => CRMService.completeFollowup(user, String(id || ''), resultado, estado));
}
function adm_listFollowups(token, bucket, asignado) {
  return admin_(token, 'crm.read', user => CRMService.listFollowups(user, bucket, asignado));
}

// ---------------- Ventas ----------------
function adm_saveProposal(token, data) {
  return admin_(token, 'sales.write', user => SalesService.saveProposal(user, data));
}
function adm_setProposalStatus(token, id, estado) {
  return admin_(token, 'sales.write', user => SalesService.setProposalStatus(user, String(id || ''), estado));
}
function adm_listProposals(token, filters) {
  return admin_(token, 'sales.read', user => SalesService.listProposals(user, filters));
}
function adm_registerSale(token, data) {
  return admin_(token, 'sales.write', user => SalesService.registerSale(user, data));
}
function adm_updateSale(token, id, patch, version) {
  return admin_(token, 'sales.write', user => SalesService.updateSale(user, String(id || ''), patch, version));
}
function adm_cancelSale(token, id, motivo) {
  return admin_(token, 'sales.write', user => SalesService.cancelSale(user, String(id || ''), motivo));
}
function adm_registerPayment(token, data) {
  return admin_(token, 'sales.write', user => SalesService.registerPayment(user, data));
}
function adm_deletePayment(token, id) {
  return admin_(token, 'sales.write', user => SalesService.deletePayment(user, String(id || '')));
}
function adm_listSales(token, filters) {
  return admin_(token, 'sales.read', user => SalesService.listSales(user, filters));
}
function adm_listPayments(token, filters) {
  return admin_(token, 'sales.read', user => SalesService.listPayments(user, filters));
}
function adm_listProjects(token, filters) {
  return admin_(token, 'sales.read', user => SalesService.listProjects(user, filters));
}
function adm_updateProject(token, id, patch, version) {
  return admin_(token, 'sales.write', user => SalesService.updateProject(user, String(id || ''), patch, version));
}

// ---------------- Auditoría (solo SUPER_ADMIN) ----------------
function adm_listAudit(token, filters, page) {
  return admin_(token, 'audit.read', () => Audit.query(filters, page, 50));
}

// ---------------- Usuarios y configuración (solo SUPER_ADMIN) ----------------
function adm_listUsers(token) {
  return admin_(token, 'users.manage', () => UserService.list());
}
function adm_createUser(token, data) {
  return admin_(token, 'users.manage', user => UserService.create(user, data));
}
function adm_updateUser(token, id, data) {
  return admin_(token, 'users.manage', user => UserService.update(user, String(id || ''), data));
}
function adm_resetUserPassword(token, id) {
  return admin_(token, 'users.manage', user => UserService.resetPassword(user, String(id || '')));
}
function adm_getConfig(token) {
  return admin_(token, 'config.write', () => ConfigService.get());
}
function adm_saveSetting(token, key, value) {
  return admin_(token, 'config.write', user => ConfigService.save(user, String(key || ''), value));
}
function adm_saveCatalogItem(token, data) {
  return admin_(token, 'catalog.write', user => CatalogService.save(user, data));
}
function adm_saveProductoWeb(token, data) {
  return admin_(token, 'config.write', user => ProductoWebService.save(user, data));
}
function adm_importStatus(token) {
  return admin_(token, 'import', () => ImportService.status());
}
function adm_importStart(token) {
  return admin_(token, 'import', () => ImportService.start());
}
function adm_systemInfo(token) {
  return admin_(token, 'config.write', () => ({
    publicUrl: (function () { try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; } })(),
    dbUrl: 'https://docs.google.com/spreadsheets/d/' + Props.get('DB_SPREADSHEET_ID'),
    auditUrl: 'https://docs.google.com/spreadsheets/d/' + Props.get('AUDIT_SPREADSHEET_ID'),
    rootFolderUrl: DriveService.folderUrl(Props.get('ROOT_FOLDER_ID')),
    permisos: DriveService.auditSharing()
  }));
}

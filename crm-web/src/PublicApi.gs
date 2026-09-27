/**
 * PublicApi.gs — FUNCIONES PÚBLICAS (formulario). Llamables por cualquiera.
 * Nunca reciben IDs de clientes: todo se deriva del token del borrador.
 */

function pub_bootstrap() {
  return respond_(() => FormService.bootstrap());
}

function pub_start(meta) {
  return respond_(() => FormService.startDraft(meta));
}

function pub_resume(token) {
  return respond_(() => FormService.resume(token));
}

function pub_search(token, type, value, municipio) {
  return respond_(() => FormService.search(token, type, value, municipio));
}

function pub_select(token, ref) {
  return respond_(() => FormService.select(token, ref));
}

function pub_unlink(token) {
  return respond_(() => FormService.unlink(token));
}

function pub_save(token, data, step) {
  return respond_(() => FormService.saveDraft(token, data, step));
}

function pub_upload(token, meta, base64) {
  return respond_(() => FormService.upload(token, meta, base64));
}

function pub_removeFile(token, fileId) {
  return respond_(() => FormService.removeFile(token, fileId));
}

function pub_filePreview(token, fileId) {
  return respond_(() => FormService.filePreview(token, fileId));
}

function pub_submit(token, data, meta) {
  return respond_(() => FormService.submit(token, data, meta));
}

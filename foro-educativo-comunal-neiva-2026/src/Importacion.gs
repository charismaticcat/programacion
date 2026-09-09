/**
 * Importacion.gs — Foro Educativo Comunal Neiva 2026
 *
 * Importa a GruposComunal la relación GRUPO → IE real, extraída por
 * lectura (NUNCA escritura) del spreadsheet de FEI 3.1
 * (1OiBPO8BEsa0TpmYGRfEu2I2tMpxIMKAJdr9WtTRd14Y) el 2026-09-09:
 *
 *   - hoja "Participacion", columna "Grupo": asigna cada institución a
 *     G1-G6 (37 IE oficiales, de las cuales 36 tienen grupo asignado).
 *   - hoja "Oficiales": CODIGO DANE y Comuna de cada institución,
 *     cruzados por nombre normalizado contra la lista anterior.
 *
 * SAN MIGUEL ARCANGEL apareció como "Sin grupo asignado" en el origen —
 * se deja fuera de esta importación a propósito (regla de no invención,
 * spec sección 23: no se le asigna un grupo que nadie ha decidido). Debe
 * agregarse manualmente a GruposComunal si se define su grupo.
 *
 * Esta lectura fue de solo consulta: en ningún momento se escribió ni se
 * modificó el spreadsheet de FEI 3.1.
 */

/**
 * Ejecutar UNA vez desde el editor. Idempotente: si una fila
 * (ID_GRUPO+ID_IE) ya existe en GruposComunal, no la duplica.
 */
function importarGruposRealesDesdeFEI31() {
  var filas = [
    ["G1", "Grupo 1", "241001000711", "AIPECITO", "RURAL", "SI"],
    ["G1", "Grupo 1", "241001001890", "CHAPINERO", "RURAL", "SI"],
    ["G1", "Grupo 1", "141001060441", "I.E. CLARETIANO GUSTAVO TORRES PARRA", "2", "SI"],
    ["G1", "Grupo 1", "141001003341", "INEM JULIAM MOTTA SALAS", "1", "SI"],
    ["G1", "Grupo 1", "141001000066", "LICEO DE SANTA LIBRADA", "1", "SI"],
    ["G1", "Grupo 1", "141001000040", "PROMOCION SOCIAL", "1", "SI"],
    ["G2", "Grupo 2", "141001004720", "EDUARDO SANTOS", "9", "SI"],
    ["G2", "Grupo 2", "341001004559", "GABRIEL GARCIA MARQUEZ", "9", "SI"],
    ["G2", "Grupo 2", "241001000486", "JAIRO MOSQUERA MORENO", "RURAL", "SI"],
    ["G2", "Grupo 2", "141001003171", "LUIS IGNACIO ANDRADE", "2", "SI"],
    ["G2", "Grupo 2", "441001004839", "MARIA AUXILIADORA FORTALECILLAS", "RURAL", "SI"],
    ["G2", "Grupo 2", "141001001038", "MARIA CRISTINA ARANGO DE PASTRANA.", "2", "SI"],
    ["G3", "Grupo 3", "141001002557", "ANGEL MARIA PAREDES", "4", "SI"],
    ["G3", "Grupo 3", "141001004061", "CEINAR", "4", "SI"],
    ["G3", "Grupo 3", "141001000058", "DEPARTAMENTAL TIERRA DE PROMISIÓN", "3", "SI"],
    ["G3", "Grupo 3", "141001001321", "RICARDO BORRERO ALVAREZ", "4", "SI"],
    ["G3", "Grupo 3", "141001000023", "SANTA LIBRADA", "3", "SI"],
    ["G3", "Grupo 3", "141001000031", "TECNICO SUPERIOR", "3", "SI"],
    ["G4", "Grupo 4", "141001005866", "ATANASIO GIRARDOT", "5", "SI"],
    ["G4", "Grupo 4", "141001005301", "ENRIQUE OLAYA HERRERA", "10", "SI"],
    ["G4", "Grupo 4", "141001004312", "HUMBERTO TAFUR CHARRY", "10", "SI"],
    ["G4", "Grupo 4", "141001004398", "JOSE EUSTASIO RIVERA", "5", "SI"],
    ["G4", "Grupo 4", "141001003481", "MISAEL PASTRANA BORRERO", "10", "SI"],
    ["G4", "Grupo 4", "241001000664", "ROBERTO DURAN ALVIRA", "RURAL", "SI"],
    ["G5", "Grupo 5", "141001002247", "ESCUELA NORMAL SUPERIOR", "7", "SI"],
    ["G5", "Grupo 5", "141001003855", "INSTITUTO TECNICO IPC ANDRES ROSA", "8", "SI"],
    ["G5", "Grupo 5", "141001005181", "JAIRO MORERA LIZCANO", "8", "SI"],
    ["G5", "Grupo 5", "141001001259", "JUAN DE CABRERA", "8", "SI"],
    ["G5", "Grupo 5", "241001000435", "SAN ANTONIO DE ANACONIA", "RURAL", "SI"],
    ["G5", "Grupo 5", "141001000899", "SANTA TERESA", "7", "SI"],
    ["G6", "Grupo 6", "141001001763", "AGUSTIN CODAZZI", "6", "SI"],
    ["G6", "Grupo 6", "441001002747", "EL CAGUAN", "RURAL", "SI"],
    ["G6", "Grupo 6", "141001004452", "EL LIMONAR", "6", "SI"],
    ["G6", "Grupo 6", "141001000082", "OLIVERIO LARA BORRERO", "6", "SI"],
    ["G6", "Grupo 6", "141001060336", "RODRIGO LARA BONILLA", "6", "SI"],
    ["G6", "Grupo 6", "441001003433", "SAN LUIS BELTRAN", "RURAL", "SI"]
  ];

  var hoja = obtenerHoja_(HOJA_GRUPOS_COMUNAL_, cabecerasGruposComunal_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var existentes = leerFilasComoObjetos_(hoja);
  var clavesExistentes = {};
  existentes.forEach(function (f) {
    clavesExistentes[String(f.ID_GRUPO).trim() + "|" + String(f.ID_IE).trim()] = true;
  });

  var insertadas = 0;
  var omitidas = 0;
  filas.forEach(function (fila) {
    var clave = fila[0] + "|" + fila[2];
    if (clavesExistentes[clave]) {
      omitidas++;
      return;
    }
    hoja.appendRow(fila);
    insertadas++;
  });

  Logger.log("Importación completa: " + insertadas + " filas nuevas, " + omitidas + " ya existían.");
  Logger.log("6 grupos, 36 IE. SAN MIGUEL ARCANGEL (DANE 141001001593) quedó 'Sin grupo asignado' en el " +
    "origen y NO se importó — agrégala manualmente si se le define un grupo.");
  return { ok: true, insertadas: insertadas, omitidas: omitidas };
}

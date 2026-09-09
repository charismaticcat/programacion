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

/**
 * Importa a CaracterizacionIE los datos oficiales de directorio (dirección,
 * sector, comuna, zona, email institucional, rector, sedes) de las 37 IE
 * oficiales de Neiva, extraídos por lectura (nunca escritura) de la hoja
 * "Oficiales" del spreadsheet de FEI 3.1 el 2026-09-09. Una vez importados,
 * quedan disponibles automáticamente para cualquier parte del sistema que
 * llame a obtenerInstitucionesDelGrupo()/obtenerCaracterizacionIE()
 * (Instituciones.gs) — no hace falta volver a pedirlos ni leer 3.1 de
 * nuevo. Ejecutar UNA vez; es idempotente (UPSERT por ID_IE).
 */
function importarCaracterizacionRealDesdeFEI31() {
  var filas = [
    ["141001001763", "AGUSTIN CODAZZI", "Carrera 18 No 12-11 sur", "Oficial", "6", "Urbana", "ieagustincodazzi@alcaldianeiva.gov.co", "RENE RODRIGUEZ ECHEVERRY", "El  Rosario; Emaya"],
    ["241001000711", "AIPECITO", "CORREGIMIENTO AIPECITO", "Oficial", "RURAL", "Rural", "ieaipecito@alcaldianeiva.gov.co", "ARBEY LUQUE DIAZ", "La Cristalina; El Nogal; La Florida; La Primavera; La Unión; La Pradera; El Triunfo"],
    ["141001002557", "ANGEL MARIA PAREDES", "Calle 9 No 14-18", "Oficial", "4", "Urbana", "ieangelmaria@alcaldianeiva.gov.co", "YANETH GARCIA SANCHEZ", "Luis Calixto Neiva"],
    ["141001005866", "ATANASIO GIRARDOT", "Carrera 32 No 18-90", "Oficial", "5", "Urbana", "ieatanasiog@alcaldianeiva.gov.co", "ALAYAM ABIV VALDERRAMA RODRIGUEZ", "Guillermo Montenegro; Liceo Batallon Tenerife; Loma De La Cruz"],
    ["141001004061", "CEINAR", "Calle 14 No 1-50", "Oficial", "4", "Urbana", "ieceinar@alcaldianeiva.gov.co", "MARIA CRISTINA BORRERO HERMIDA", "Renaciendo"],
    ["241001001890", "CHAPINERO", "CORREGIMIENTO CHAPINERO", "Oficial", "RURAL", "Rural", "iechapinero@alcaldianeiva.gov.co", "JOHN ALEXANDER CASTILLO FORERO", "Altamira; Bajo Horizonte; Cachichi; Diamante; Horizonte; Jardin; La Cabaña; Libano; Omega"],
    ["141001060441", "I.E. CLARETIANO GUSTAVO TORRES PARRA", "Calle 55 No. 31 151", "Oficial", "2", "Urbana", "ieclaretiano@alcaldianeiva.gov.co", "ALVARO CAMACHO TORRES", ""],
    ["141001000058", "DEPARTAMENTAL TIERRA DE PROMISIÓN", "Calle 21 No 1E bis 40", "Oficial", "3", "Urbana", "iedepartamental@alcaldianeiva.gov.co", "PABLO EMILIO POLO COLLAZOS", "Efrain Rojas Trujillo; Enriqueta Solano Duran; El Lago"],
    ["141001004720", "EDUARDO SANTOS", "Calle 80C No 4-61", "Oficial", "9", "Urbana", "ieeduardosantos@alcaldianeiva.gov.co", "RICAROD SANCHEZ QUINTERO", "Alberto Rosero Concha; Luis Carlos Galán"],
    ["441001002747", "EL CAGUAN", "Calle 2 No 4-63", "Oficial", "RURAL", "Rural", "iecaguan@alcaldianeiva.gov.co", "PABLO EMILIO CASTILLO QUIROGA", "Barro Negro; El Chapuro; El Triunfo; La Gabriela; La Lindosa; San Bartolo"],
    ["141001004452", "EL LIMONAR", "Carrera 37 sur No 19-21", "Oficial", "6", "Urbana", "ielimonar@alcaldianeiva.gov.co", "HENRY WILLIAM UNI YUGCHA", "Buenos Aires; Garabaticos; Lomalinda"],
    ["141001005301", "ENRIQUE OLAYA HERRERA", "Calle 27A No 51A-35", "Oficial", "10", "Urbana", "ieenriqueolaya@alcaldianeiva.gov.co", "ISLENIA ROBAYO GUZMAN", "Las  Camelias; San Bernardo"],
    ["141001002247", "ESCUELA NORMAL SUPERIOR", "Calle 8 No 36-20", "Oficial", "7", "Urbana", "ienormalsuperior@alcaldianeiva.gov.co", "LIBARDO PERDOMO CEBALLOS", "Las Brisas; Escuela Popular Claretiana; Platanillal; El Centro; El Vergel; Tuquila; Motilon; Floragaita; Las Nubes; Los Cauchos; Pueblo Nuevo; Santa Helena; Santa Barbara"],
    ["341001004559", "GABRIEL GARCIA MARQUEZ", "Calle 86 No 7-28", "Oficial", "9", "Urbana", "iegabrielgarcia@alcaldianeiva.gov.co", "NELSON VIVAS CUPITRE (Asignación de funciones)", "Humberto Tafur Charry; Alberto Galindo; Jose Maria Carbonell; El Venado"],
    ["141001004312", "HUMBERTO TAFUR CHARRY", "Carrera 50C No 18-38", "Oficial", "10", "Urbana", "iehumbertotafur@alcaldianeiva.gov.co", "LUZ MARINA ALDANA GARCIA", "Las Palmas; Las Palmitas"],
    ["141001003341", "INEM JULIAM MOTTA SALAS", "Carrera 1 No 26-215", "Oficial", "1", "Urbana", "ieinem@alcaldianeiva.gov.co", "LUIS ALFREDO CERQUERA AYALA", "Candido Leguizamo; Mauricio Sanchez Garcia"],
    ["141001005181", "JAIRO MORERA LIZCANO", "Calle 1A No 28-82", "Oficial", "8", "Urbana", "iejairomorera@alcaldianeiva.gov.co", "MIGUEL ANTONIO PEREZ SUAREZ", "Panorama; Guillermo Lievano"],
    ["241001000486", "JAIRO MOSQUERA MORENO", "Corregimiento Guacirco", "Oficial", "RURAL", "Rural", "ieguacirco@alcaldianeiva.gov.co", "CARLOS ARTURO BOHORQUEZ SANCHEZ", "Nina Andrade de Lievano (Altares); Peñas Blancas; San Francisco; San Jorge; Tamarindo"],
    ["141001004398", "JOSE EUSTASIO RIVERA", "Carrera 32 No 15-50", "Oficial", "5", "Urbana", "iejoseeustasio@alcaldianeiva.gov.co", "CECILIA LOSADA DE FIERRO", "Ciudad Jardin; Centro Docente Monserrate; Eliseo Cabrera"],
    ["141001001259", "JUAN DE CABRERA", "Carrera 21 No 1C-80", "Oficial", "8", "Urbana", "iejuandecabrera@alcaldianeiva.gov.co", "ABRAHAM GOMEZ GALINDO", "Sur Oriental; Alfonso Lopez; Ventilador"],
    ["141001000066", "LICEO DE SANTA LIBRADA", "Carrera 1 No 26-345", "Oficial", "1", "Urbana", "ieliceosantal@alcaldianeiva.gov.co", "GLORIA GONZALEZ PERDOMO", "El Triangulo"],
    ["141001003171", "LUIS IGNACIO ANDRADE", "Calle 35 No 6-62", "Oficial", "2", "Urbana", "ieluisignacio@alcaldianeiva.gov.co", "AMANDA BERMEO CARVAJAL", "Reinaldo Matiz  Trujillo; Eugenio Salas Trujillo"],
    ["441001004839", "MARIA AUXILIADORA FORTALECILLAS", "Plaza Principal Fortalecillas", "Oficial", "RURAL", "Rural", "iefortalecillas@alcaldianeiva.gov.co", "JUAN ALONSO ESPINOSA HERRERA", "La Jagua; La Mojarra"],
    ["141001001038", "MARIA CRISTINA ARANGO DE PASTRANA.", "Carrera 8 bis No 33-25", "Oficial", "2", "Urbana", "iemariacristina@alcaldianeiva.gov.co", "TOBIAS RENGIFO RENGIFO", "Los Pinos; Mi Pequeño Mundo"],
    ["141001003481", "MISAEL PASTRANA BORRERO", "Carrera 49 A No 20B-30", "Oficial", "10", "Urbana", "iemisaelpastrana@alcaldianeiva.gov.co", "JAIRO RAMIREZ CEDEÑO", "La Rioja"],
    ["141001000082", "OLIVERIO LARA BORRERO", "Calle 16 sur No 21A-17", "Oficial", "6", "Urbana", "ieoliveriolara@alcaldianeiva.gov.co", "LUIS ALFONSO BURBANO CLEVES", "Timanco; Santa Isabel; Manuela Beltran"],
    ["141001000040", "PROMOCION SOCIAL", "Calle 48 No 1B-55", "Oficial", "1", "Urbana", "iepromocion@alcaldianeiva.gov.co", "TEOFILO ORTIZ TOVAR", "Las Mercedes; Contraloría General de la Republica; Colombo Andino"],
    ["141001001321", "RICARDO BORRERO ALVAREZ", "Calle 5 No 5A-61", "Oficial", "4", "Urbana", "iericardoborrero@alcaldianeiva.gov.co", "LUZ MARY VARGAS PLAZAS (E)", "Oriente; Jardin Infantil Nacional"],
    ["241001000664", "ROBERTO DURAN ALVIRA", "CORREG VEGALARGA", "Oficial", "RURAL", "Rural", "ierobertoduran@alcaldianeiva.gov.co", "MAGNA BADELEY LOPEZ NAVARRO", "Ahuyamales; Jorge Villamil Cordovez; Piedra Marcada; El Colegio"],
    ["141001060336", "RODRIGO LARA BONILLA", "Carre 33 No. 30 02 Sur", "Oficial", "6", "Urbana", "ierodrigolara@alcaldianeiva.gov.co", "AURA MARIA LOSADA LEMUS", "VI centenario"],
    ["241001000435", "SAN ANTONIO DE ANACONIA", "VDA. SAN ANTONIO DE ANACONIA", "Oficial", "RURAL", "Rural", "iesanantonio@alcaldianeiva.gov.co", "OSCAR LEONARDO SOTO CASTRO", "Canoas; La Espiga; Alpes; Palacio; Palestina; Primavera; Roblal; San Jose; San Miguel; Santa Librada; Santa Lucia; Pedro Calderon Castro"],
    ["441001003433", "SAN LUIS BELTRAN", "CORREGIMIENTO SAN LUIS", "Oficial", "RURAL", "Rural", "iesanluisbeltran@alcaldianeiva.gov.co", "EDUARDO GUILOMBO ROJAS", "Alpes; Alta Libertad; Alto Cocal; Corozal; Avila; Centro; Cocal; Piñuelo; Quebradon; Julia; Libertad; Merceditas; Organos; Palmar"],
    ["141001001593", "SAN MIGUEL ARCANGEL", "Carrera 1 No 39-23", "Oficial", "1", "Urbana", "menev.cosma-rec@policia.gov.co", "LIZETH LOZANO VACA", ""],
    ["141001000023", "SANTA LIBRADA", "Carrera 12 No 16-12", "Oficial", "3", "Urbana", "ienacionalsl@alcaldianeiva.gov.co", "CARLOS FERNANDO MANCHOLA REYES", "Gabino Charry; Martha Tello"],
    ["141001000899", "SANTA TERESA", "Carrera 23 No 4B-18 BRIO OBRERO", "Oficial", "7", "Urbana", "iesantateresa@alcaldianeiva.gov.co", "JUAN PABLO YAGUARA GALVIS", "Blanca Motta Salas; Francisca Borrero De Perdomo; Jorge Villamil Cordovez; Oliverio Lara Borrero"],
    ["141001003855", "INSTITUTO TECNICO IPC ANDRES ROSA", "Diagonal 2B No 28B-32", "Oficial", "8", "Urbana", "ieipc@alcaldianeiva.gov.co", "CARLOS ALBERTO VARGAS VARGAS (E)", "La Gaitana; Jardin Picardias; Rafael Azuero; La Paz; Picardias Primaria"],
    ["141001000031", "TECNICO SUPERIOR", "Calle 21 No 2-72-albergue cra 3 #21-15", "Oficial", "3", "Urbana", "ietecnicos@alcaldianeiva.gov.co", "STEVENSON MALDONA MEDINA", "Elena Lara; Floresmiro Azuero; Los Martires"],
  ];

  var hoja = obtenerHoja_(HOJA_CARACTERIZACION_IE_, cabecerasCaracterizacionIE_());
  var mapa = obtenerMapaCabeceras_(hoja);
  var cabeceras = cabecerasCaracterizacionIE_();

  var insertadas = 0;
  var actualizadas = 0;
  filas.forEach(function (fila) {
    var valores = {};
    cabeceras.forEach(function (c, i) {
      valores[c] = fila[i];
    });
    var idIE = valores["ID_IE"];
    var yaExistia = buscarFilaPorColumna_(hoja, mapa, "ID_IE", idIE) !== -1;
    upsertFila_(HOJA_CARACTERIZACION_IE_, cabeceras, "ID_IE", idIE, valores);
    if (yaExistia) actualizadas++;
    else insertadas++;
  });

  Logger.log("Caracterización importada: " + insertadas + " nuevas, " + actualizadas + " actualizadas (37 IE en total).");
  return { ok: true, insertadas: insertadas, actualizadas: actualizadas };
}

/**
 * Importa a CaracterizacionIE.LOGO_ID el logo real de cada una de las 36
 * IE que sí tienen grupo asignado, extraído por lectura (NUNCA escritura)
 * de la hoja "AccesosIE" del spreadsheet de FEI 3.1
 * (1OiBPO8BEsa0TpmYGRfEu2I2tMpxIMKAJdr9WtTRd14Y) el 2026-09-10 — mismos
 * IDs de archivo de Drive que ya usa 3.1 para mostrar el logo de cada IE
 * (ver auditoría §1.4/§3.1, columna LOGO_ID de AccesosIE). Usados aquí
 * por el carrusel de bienvenida (spec: "aparece uno a uno cada logo de
 * las IE") en vez del ícono genérico 🏫.
 *
 * SAN MIGUEL ARCANGEL también tiene LOGO_ID en el origen, pero se deja
 * fuera de esta importación por la misma razón que en
 * importarGruposRealesDesdeFEI31(): no tiene grupo asignado en este
 * proyecto.
 *
 * Ejecutar UNA vez desde el editor. Idempotente: solo actualiza la
 * columna LOGO_ID de una fila que ya existe en CaracterizacionIE (creada
 * por importarCaracterizacionRealDesdeFEI31()); nunca toca las demás
 * columnas de esa fila.
 */
function importarLogosIEDesdeFEI31() {
  var logosPorIE = {
    "141001001763": "1CypzNY5rHr0hT5jJaJDKKLme6gXqaKab", // AGUSTIN CODAZZI
    "241001000711": "1AUEaynnesN6BDKfQwCPCG_260e1Wp_fu", // AIPECITO
    "141001002557": "1Wsj0n0XpNagX6Dnm2btyBM2_G0nqGwZx", // ANGEL MARIA PAREDES
    "141001005866": "1_5Bgo-l65Mn1mzx-k9mpXs47n7i0x5FY", // ATANASIO GIRARDOT
    "141001004061": "1CgJolG0Kaltfk-JydOyVmiuKR5qptZdQ", // CEINAR
    "241001001890": "1RwTJKVT0Lz0d1aqXaS_k6Rr80U1lcNY-", // CHAPINERO
    "141001000058": "1fgJ2IGktyPH8lCjIsTtG6aHkZP8kmJvK", // DEPARTAMENTAL TIERRA DE PROMISIÓN
    "141001004720": "1khSyS3PadHCbg9g2pIoUtagJdnlKmDBR", // EDUARDO SANTOS
    "441001002747": "1-EOOOwrBTv6v8fNFvJ1bxdcueA10GXk0", // EL CAGUAN
    "141001004452": "1YPUzYhXcHpCP5yW0bxEuxzAXY-lKNcK1", // EL LIMONAR
    "141001005301": "1f2EfP1reThB3cZtnSk9a21z0k3gHuTMw", // ENRIQUE OLAYA HERRERA
    "141001002247": "1A4NVFeqJRzdk8Hf4vhOlvp0RDSD1FbM7", // ESCUELA NORMAL SUPERIOR
    "341001004559": "1yOm29DTzmlSZXuebjcuaIRrjxS_EArWv", // GABRIEL GARCIA MARQUEZ
    "141001004312": "12IMJnIMuXw-Y9tOlRoFvoNQVtIzgaNfU", // HUMBERTO TAFUR CHARRY
    "141001060441": "16YEwkaUpOpFkLD5DuLBxxdJ4bZFQWx_5", // I.E. CLARETIANO GUSTAVO TORRES PARRA
    "141001003341": "1AResGxiz_RYP7hcWpQzjnVaC0lQ1ee0Z", // INEM JULIAM MOTTA SALAS
    "141001003855": "13VipiW9jLCgFKXfvlzjw9oZaNWMerKgT", // INSTITUTO TECNICO IPC ANDRES ROSA
    "141001005181": "1GSpglPbQKsuS8knlPSIApexwmiFqvfw7", // JAIRO MORERA LIZCANO
    "241001000486": "1n66NFNmaQUx4eZ_kPdBETLujzTOSeP2w", // JAIRO MOSQUERA MORENO
    "141001004398": "1rF7GgwfAgGE2NiWQVsWz9-N7eE79tOMj", // JOSE EUSTASIO RIVERA
    "141001001259": "1LNVlw94GlDUXcDBgnrHBNUznO0_vbQIJ", // JUAN DE CABRERA
    "141001000066": "1wod9Lo49mXaQZocZEGrnuBrPMrHkj1CL", // LICEO DE SANTA LIBRADA
    "141001003171": "1A0Npz9kbYP5q-rzad5Tv96eTQLepnh08", // LUIS IGNACIO ANDRADE
    "441001004839": "1YpVP_7QvXtaNKNMw4iBMLAJRHqoun0SX", // MARIA AUXILIADORA FORTALECILLAS
    "141001001038": "19bPnYVAM0m9zxBlTKOAC2Uku1EDLBkeI", // MARIA CRISTINA ARANGO DE PASTRANA.
    "141001003481": "1rawhFbEdsY46Smi6J78F9T2c0tD8H4_D", // MISAEL PASTRANA BORRERO
    "141001000082": "1ajZPYR-UgahGVZ3dyEbC0wrOmmAX94T4", // OLIVERIO LARA BORRERO
    "141001000040": "1ikm7V_R5vUUfNonZ48VDvu77KuSeB0PS", // PROMOCION SOCIAL
    "141001001321": "11STHwjLyXVZRHYqCIHOZA4TI5sFG-cet", // RICARDO BORRERO ALVAREZ
    "241001000664": "14oGMdR0iBuy8SzQOT8sydEVFNaM6NRZI", // ROBERTO DURAN ALVIRA
    "141001060336": "1nNLed3VZN91Yk8dgxlcs2-d9aj-AyVba", // RODRIGO LARA BONILLA
    "241001000435": "1ika98rWcmBP39OYZkxcLz2hxonBDrvDJ", // SAN ANTONIO DE ANACONIA
    "441001003433": "1KgQjp8xg3Q4GeVxnxnoI_g3560y57FzJ", // SAN LUIS BELTRAN
    "141001000023": "1M2q_Pe0JLkCBdc78rdOsAfHvghcNMvYm", // SANTA LIBRADA
    "141001000899": "1hvIKi4A2FcGAv_xjFyoptgl3vblL_w-g", // SANTA TERESA
    "141001000031": "1jtFYXfvM8WUFxIr2Tp-9CotrCQYK1CTz" // TECNICO SUPERIOR
  };

  var cabeceras = cabecerasCaracterizacionIE_();
  var actualizadas = 0;
  Object.keys(logosPorIE).forEach(function (idIE) {
    upsertFila_(HOJA_CARACTERIZACION_IE_, cabeceras, "ID_IE", idIE, { LOGO_ID: logosPorIE[idIE] });
    actualizadas++;
  });

  Logger.log("Logos de IE importados: " + actualizadas + " de 36.");
  return { ok: true, actualizadas: actualizadas };
}

# Guía de puesta en marcha — Foro Educativo Comunal Neiva 2026

Este documento cubre las Fases 10-16 de la especificación: cómo pasar de este código a un Foro
funcionando. El proyecto de Apps Script objetivo es **Foro comunal 1.0**
(`1tqsSNT-3BiCkQhcfSCVmzMe9IeVIsDuFujf_rzMRygz3pDUbRP7fKy7Y`), independiente de **Foro Educativo 3.1**
(solo lectura, nunca se modifica ni se le escribe).

## 1. Publicar el código (`clasp push`)

```bash
cd foro-educativo-comunal-neiva-2026
clasp push
```

Requiere haber hecho `clasp login` una vez (ya autorizado en esta sesión con
`jhonefrainsanchez@gmail.com`). El `.clasp.json` del proyecto ya apunta al scriptId de **Foro comunal
1.0**.

## 2. Primera ejecución — autoprovisión de Sheets y Drive

Desde el editor de Apps Script (script.google.com), abrir `Tests.gs` y ejecutar, en orden:

1. `testInicializarProyecto()` — crea el spreadsheet nuevo del proyecto (independiente de 3.1) con
   todas las hojas base (`GruposComunal`, `AccesosGrupo`, `ParticipacionComunal`, `Sesion1Comunal`,
   `ConectaEduca`, `InformesComunal`, `ConfiguracionComunal`, `EnviosDiferidosComunal`) y la estructura
   de carpetas de Drive (`FORO EDUCATIVO COMUNAL NEIVA 2026/01_DATOS ... 06_EVIDENCIAS`). La primera vez
   pedirá autorizar los permisos del proyecto.
2. Revisar el log: ahí queda el enlace directo al spreadsheet nuevo.

**Regla de no invención** (spec sección 23): este proyecto nunca genera datos reales de grupos,
instituciones, entidades invitadas ni estadísticas. Todo lo que sigue debe cargarse a mano con
información real suministrada por la SEM/organización del Foro.

## 3. Cargar datos reales

En el spreadsheet nuevo (no en el de 3.1):

- **`GruposComunal`**: una fila por cada IE de cada grupo (`ID_GRUPO, GRUPO, ID_IE, INSTITUCION, COMUNA,
  ACTIVO`). Esta hoja es la única fuente de la relación Grupo → IE.
- **`ConfiguracionComunal`**: completar al menos `NOMBRE_FORO`, `SUBTITULO`, `FECHA`,
  `LOGO_ENCABEZADO_ID`, `LOGO_PIE_ID` (subir los logos a Drive y pegar su ID de archivo),
  `CORREO_REMITENTE` (debe ser una cuenta o alias que la persona que ejecuta el script pueda usar en
  `GmailApp`), `COPIAS_CORREO`. `CARPETA_DRIVE_ID` y `ID_FORO_COMUNAL` ya quedan poblados por el paso 2.
- **`AccesosGrupo`**: completar manualmente `EMAIL_RESPONSABLE_GRUPO` de cada grupo — es solo el correo
  de arranque para enviarle el TOKEN+código inicial (paso 4). El informe final ya no se envía a este
  correo por defecto: se envía al **responsable de envío** que el propio grupo registra dentro de la
  app (hasta 4: 1 principal + 3 asistentes, ver `ResponsablesComunal`) — si el grupo no llega a
  registrar ninguno, `enviarInformeGrupo` cae de vuelta a este correo de `AccesosGrupo`.
- **`CaracterizacionIE`**: ya se pobló con `importarCaracterizacionRealDesdeFEI31()` (ver mensaje
  anterior de esta guía) — no requiere carga manual.

## 3.1 Consentimiento informado y flujo nuevo

- La primera vez que un grupo entra, ve una pantalla de **consentimiento informado** (contenido de la
  sección 3.2 del Documento Orientador FEM2026, sin fechas/horarios) que debe aceptar antes de
  continuar — queda registrado en `AccesosGrupo.CONSENTIMIENTO_GRUPO` y no se vuelve a mostrar a otros
  dispositivos del mismo grupo.
- En Participación, cada grupo registra su **responsable de envío** (principal) y hasta **3 asistentes
  de envío**, cada uno con nombre, IE, **rol en el foro** (catálogo oficial: Líder, Dinamizador
  Pedagógico, Dinamizador de Mesas de Trabajo, Relator(a), Dinamizador del Tiempo, Dinamizador de la
  Sistematización, Participante, Otro) y correo.
- La misma pantalla muestra en vivo la **matriz de participación por estamento e institución**
  (equivalente a la hoja `Participacion` de 3.1), calculada a partir de las firmas ya registradas.
- Antes de pasar a Sesión 1, el grupo ve una pantalla de **confirmación de caracterización** con la
  ficha completa (instituciones, matriz, responsables) para verificar antes de continuar.
- Si el grupo elige el método QR, puede abrir una **vista de pantalla completa** (botón 👁) para
  proyectar el código QR y el contador de firmantes en vivo durante el evento.
- Cada sesión (1 y 2/ConectaEduca) tiene un desplegable opcional **"➕ Aporte propio del grupo"** para
  hallazgos que no encajen en las preguntas orientadoras.

## 4. Generar accesos y publicar la web app

1. Ejecutar `generarAccesosGrupo()` desde el editor — genera TOKEN + código + 3 códigos de contingencia
   por cada grupo de `GruposComunal` que todavía no tenga fila en `AccesosGrupo` (idempotente: nunca
   regenera credenciales ya emitidas).
2. **Implementar → Nueva implementación → Aplicación web** (executeAs: yo, acceso: cualquier usuario,
   igual que 3.1).
3. Ejecutar `actualizarUrlsAcceso()` para que `URL_ACCESO` en `AccesosGrupo` apunte al deployment
   publicado (no al `/dev` del editor).
4. Ejecutar `enviarAccesosGrupo(idGrupo)` por cada grupo (o adaptar un bucle sobre `obtenerGrupos()`)
   para enviar el correo de acceso a cada responsable.

## 4.1 Logo por grupo y método de asistencia (Asistencia.gs)

Igual que 3.1 resolvía el logo de cada IE dinámicamente (columna `LOGO_ID` en `AccesosIE`, ver
auditoría §1.4), cada grupo puede tener su propio logo:

1. Sube la imagen del logo a Drive (a mano) y copia el ID del archivo.
2. Ejecuta `asignarLogoGrupo(idGrupo, logoFileId)` desde el editor, una vez por grupo.

El método de asistencia (QR/enlace, o listado físico + fotografía) lo elige cada grupo desde la
pantalla de **Participación** de la aplicación (no hace falta configurarlo desde el editor) — queda
guardado en `AccesosGrupo.METODO_ASISTENCIA` y es visible para todos los dispositivos conectados al
grupo. La página pública de firma rápida (QR/enlace, sin token ni código, pensada para escanear en el
evento) vive en `?asistencia=<ID_GRUPO>` — su URL se construye automáticamente y se muestra junto con
un código QR generado por [api.qrserver.com](https://api.qrserver.com) cuando el grupo elige ese
método. El listado físico sube a `02_ASISTENCIA/GRUPO N` y la fotografía de evidencia a
`06_EVIDENCIAS/GRUPO N`, ambas dentro de la estructura de Drive del proyecto.

**Diferencia deliberada frente a 3.1**: cambiar de método NO borra los datos del método anterior (3.1
sí lo hacía — `eliminarAsistenciaQRPorCambioMetodo`/`PDF`, ver auditoría §7). Aquí se prefiere no
arriesgar pérdida de datos durante un evento en vivo.

## 4.2 Transición inicial y carrusel de instituciones

Al abrir la aplicación, antes de la pantalla de bienvenida, aparece una transición (menos de 5
segundos, o toque para saltar) con el logo del Foro (`ConfiguracionComunal.LOGO_ENCABEZADO_ID`) y el
logo de la Alcaldía (`ConfiguracionComunal.LOGO_PIE_ID`) — los mismos IDs de Drive que ya se usan para
encabezar el informe (Informes.gs), reutilizados aquí. Si algún logo no está configurado, esa imagen
simplemente no se muestra (nunca se inventa un logo).

Al tocar "Comenzar", antes de pasar a la Presentación, se reproduce un carrusel que muestra una a una
(1 segundo cada una) todas las I.E. activas de `GruposComunal` (de todos los grupos, porque en este
punto el visitante todavía no ha ingresado el código de su grupo), con su nombre y su logo si ya fue
cargado. Para asignar el logo de una IE:

```js
asignarLogoIE(idIE, logoFileId); // una vez por IE, desde el editor de Apps Script
```

Si una IE no tiene logo asignado, el carrusel muestra un ícono genérico (🏫) en su lugar — nunca un
logo inventado. Hay un botón "Saltar" en todo momento para omitir el carrusel.

## 4.3 Archivos de apoyo antes de Sesión 1

Justo después de confirmar la caracterización del grupo, y también como botones pequeños dentro de la
propia pantalla de Sesión 1, la aplicación ofrece 4 archivos descargables del Foro Educativo
Institucional 2026: las Respuestas Compiladas y el Informe de Síntesis del grupo activo (resueltos
dinámicamente según `ID_GRUPO`), y el informe de síntesis municipal y las FAQs (fijos para todos los
grupos). Los IDs de Drive están en `Recursos.gs` (`RECURSOS_POR_GRUPO_`, `RECURSOS_FIJOS_SESION1_`),
verificados uno a uno contra la carpeta pública
["Informes por grupos"](https://drive.google.com/drive/folders/1SAjGsKFNudF94ag_xCWmPNuqE0nj9T85) del
FEI — si se agregan grupos nuevos o se reemplazan esos documentos, hay que actualizar ese mapa a mano
(nunca se generan IDs automáticamente).

**Importante — permisos de Drive**: estos 4 recursos siguen siendo propiedad de la cuenta que ejecuta
el script; el enlace funciona para el dueño, pero cualquier otra persona recibe "no se puede abrir el
archivo en estos momentos" si el archivo/carpeta no está compartido con ella. Antes de enviar accesos a
usuarios reales, comparte las 6 carpetas de grupo y los 2 archivos fijos con "Cualquier persona con el
enlace" (o con cada destinatario específico) desde Google Drive.

## 4.4 Contador de palabras, contador de participantes y foto general del grupo

- **Contador de palabras**: cada pregunta abierta de Sesión 1 y Sesión 2 (incluidos los campos de
  "Aporte propio") muestra en vivo, debajo del cuadro de texto, cuántas palabras lleva escritas
  (`inicializarContadoresPalabras()` en `JS.html`, genérico sobre cualquier `textarea` de esas dos
  pantallas — no hace falta tocarlo si se agregan preguntas nuevas).
- **Contador de participantes**: además del contador que ya aparece fijo arriba de toda la pantalla
  (barra sticky) desde que se entra a Participación, el mismo total ahora también se muestra dentro de
  la propia tarjeta "Participación y asistencia".
- **Fotografía general del grupo**: en la tarjeta de Participación, cualquier dispositivo puede subir
  una foto representativa del encuentro — independiente del método de asistencia elegido, a diferencia
  de la fotografía de evidencia (que solo aplica cuando el método es "Listado físico"). Se guarda en
  `06_EVIDENCIAS/GRUPO N` (columna `FOTO_GRUPO_ID` de `AccesosGrupo`, distinta de `ID_FOTO_EVIDENCIA`).

## 4.5 Valoración del Foro como condición para el informe

A diferencia de FEI 3.1 (donde la valoración "no bloquea nada" — es una encuesta de cierre opcional
después de enviar el informe), en esta entrega la valoración SÍ es condición: en "Revisión y cierre", el
grupo debe responder primero las 4 preguntas de corazones (1 a 5) + la pregunta abierta final antes de
que el botón "Generar informe del grupo" se habilite (`Valoracion.gs`, gatea
`generarInformeCompletoGrupo` en `Grupos.gs`). El instrumento (preguntas, mecánica de corazones,
umbrales de la nota, mensaje final condicionado al promedio) es el mismo de FEI 3.1
(`inicializarValoracionFEM_`), adaptado de "Foro Educativo Institucional"/IE a "Foro Educativo
Comunal"/grupo.

Además, el botón "Enviar informe por correo" (pantalla de informe generado) queda deshabilitado hasta
que el grupo haga clic en "Descargar informe" (se registra en `InformesComunal.DESCARGADO`) — antes, el
correo se enviaba automáticamente al generar el informe; ahora es una acción explícita y separada,
gateada por valoración + descarga (`enviarInformeSiCorresponde` en `Correo.gs`).

## 4.6 Transición inicial, recorrido de logos y otros ajustes de recorrido

- **Transición inicial**: primero el logo del Foro (`LOGO_ENCABEZADO_ID`), luego el de la SEM Neiva
  (`LOGO_PIE_ID`), uno a la vez con crossfade — ya vienen con los IDs reales de Drive por defecto en
  `CONFIG_POR_DEFECTO_`; se pueden cambiar editando `ConfiguracionComunal` sin tocar código.
- **Recorrido de logos de las IE**: ahora se muestra justo después de validar el código de acceso (ya
  no al tocar "Comenzar", que ahora navega de inmediato — antes esperaba una llamada al servidor y se
  sentía lento).
- **Transiciones entre pantallas**: se alargó y suavizó la animación (`aparecer`, en `CSS.html`).
- Sesión 1 ahora tiene un botón "Atrás" hacia Confirmación de caracterización.
- La fotografía de evidencia dentro del método "Listado físico" se eliminó (quedaba duplicada con la
  fotografía general del grupo, independiente del método, agregada en Participación).

## 4.7 Participación por estamento e institución (formato de FEI 3.1) y ficha editable

- La tarjeta "Cantidad de asistentes por estamento e institución" ya no se calcula solo a partir de las
  firmas QR: ahora es un conteo manual, con el mismo formato y los mismos 11 estamentos que la
  caracterización de FEI 3.1 (incluye "Tutor PTA PFI/3.0"), repetido una vez por cada IE del grupo, con
  autosuma (`ParticipacionEstamento.gs`, hoja `ParticipacionEstamentoIE`). Si el grupo eligió el método
  QR, el resumen muestra "firmantes / participantes declarados" para comparar ambos números.
- En "Confirmación de caracterización", la ficha de cada IE ahora incluye su código DANE y el nombre del
  rector(a) como campo editable (con botón 💾), igual que en FEI 3.1 — se guarda en
  `CaracterizacionIE.RECTOR` y solo se puede editar la de una IE que pertenezca al mismo grupo.

## 5. Pruebas antes de producción (Fase 15 de la spec)

Usar `GRUPO-PRUEBA` (nunca datos reales) para validar el flujo sin afectar la carga real:

```js
testCrearGrupoDePrueba();        // crea GRUPO-PRUEBA con 2 IE ficticias
testGenerarAccesos();
testMostrarAccesoDePrueba();     // TOKEN/CODIGO/URL para probar desde el navegador
testFlujoCompletoGrupoPrueba();  // simula todo el recorrido de punta a punta
testLimpiarDatosDePrueba();      // borra los datos de prueba al terminar
```

Verificar además, manualmente desde el navegador con el acceso de prueba: panel de firmantes en tiempo
real, guardado concurrente de Sesión 1 desde dos pestañas/dispositivos distintos (probar la fusión de
campos), y que el PDF del informe quede compartido públicamente mientras el `.docx` no.

Durante las pruebas, **no ejecutar `enviarAccesosGrupo`/`enviarInformeGrupo` de forma masiva** (spec
sección 26) — probar con un solo grupo real o con `GRUPO-PRUEBA`.

## 6. Auditoría final de dependencias (Fase 16 de la spec)

Antes de dar el proyecto por terminado, confirmar cada punto de la "Condición de finalización" (spec,
al inicio del documento):

- [ ] El aplicativo funciona independientemente de FEI 3.1 (ningún archivo de este proyecto abre
      `SPREADSHEET_ID`, carpetas ni propiedades de 3.1 — verificado: `Data.gs`/`Config.gs`/`Drive.gs` solo
      usan `getConfig().SPREADSHEET_ID`/`CARPETA_DRIVE_ID`, ambos autoprovisionados por este proyecto).
- [ ] No escribe datos en las hojas de 3.1.
- [ ] Usa sus propias hojas (`*Comunal`, `AccesosGrupo`, `ConectaEduca`).
- [ ] Usa sus propias carpetas (`FORO EDUCATIVO COMUNAL NEIVA 2026/...`).
- [ ] Usa sus propias plantillas (el informe se genera con `DocumentApp.create`, no copia nada de 3.1).
- [ ] Usa la lógica GRUPO → IE (`GruposComunal`, sin catálogo hardcodeado).
- [ ] La asistencia funciona dentro de Participación.
- [ ] Los firmantes son visibles durante todo el recorrido (panel `sticky`).
- [ ] Sesión 1 es colectiva por grupo (UPSERT con fusión, un registro por `ID_GRUPO`).
- [ ] ConectaEduca es colectivo por grupo.
- [ ] Existe un único informe por grupo (`InformesComunal`, un DOC_ID/PDF_ID por `ID_GRUPO`).
- [ ] Todas las IE del grupo tienen acceso al mismo informe (`obtenerInformeGrupo`, mismo resultado para
      cualquier IE del grupo).
- [ ] El sistema puede desplegarse como aplicación independiente.
- [ ] Todo el código se gestiona mediante CLASP (`.clasp.json` propio, scriptId de **Foro comunal 1.0**).

## 7. Pendiente fuera del alcance de este código

- **Dashboard administrativo** (spec sección 25): no implementado en esta entrega — `obtenerEstadoGrupo`
  ya expone los datos base (participantes, Sesión 1/2 enviada, informe generado) para construirlo como
  siguiente iteración, posiblemente como una pantalla adicional o un spreadsheet con `IMPORTRANGE`/Apps
  Script consultando `obtenerEstadoGrupo` para cada grupo.
- **Consolidado municipal** (carpeta `05_CONSOLIDADO_MUNICIPAL`, spec sección 24): la carpeta se crea,
  pero el informe consolidado municipal (agregando los informes de todos los grupos) no está
  implementado todavía — es la contraparte, a nivel de todo el Foro Comunal, de
  `generarInformeSintesisMunicipalFEM_` en 3.1 (ver auditoría §1.6), y conviene diseñarlo una vez haya
  datos reales de al menos un ciclo completo de grupos.
- **Trigger de reintento de correo diferido**: `reintentarEnviosDiferidos()` existe pero no tiene un
  trigger de tiempo instalado — crear uno diario (Apps Script → Activadores) apuntando a esa función.

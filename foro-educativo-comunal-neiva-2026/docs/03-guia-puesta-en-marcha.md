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

## 4.8 Sesión de preparación (pre-socialización) — `Preparacion.gs`

- Nueva etapa entre "Confirmación de caracterización" y "Sesión 1": todas las IE de un grupo ingresan con
  el mismo código de acceso (ya no hay límite de 4 responsables). En "Sesión de preparación" el grupo
  elige, IE por IE, quién va a diligenciar sus respuestas; al confirmar el nombre se hace una transición
  con el logo de esa institución y se abre su pantalla de preparación.
- En la pantalla de preparación de cada IE aparece, si existe, un enlace para ver/descargar su propio
  "Informe Ejecutivo (IE) FEM 2026" real (Drive), y 6 preguntas totalmente editables cuyo texto de partida
  es un **resumen sugerido** extraído por lectura (nunca escritura) de ese mismo informe:
  1. Avances en el logro de retos y propósitos del SEM 2025.
  2. Implementación de niveles de preescolar (jardín, prejardín).
  3. Pertinencia curricular con las realidades de la comunidad.
  4. Acciones pedagógicas para articular el currículo con la comunidad.
  5. Equipos de trabajo para articular con la comunidad.
  6. Democracia institucional.
- `RESUMENES_PREPARACION_IE_` (en `Preparacion.gs`) contiene el resumen sugerido de las 36 IE reales,
  extraído el 2026-09-09 de los archivos "Informe Ejecutivo - `<IE>` FEM 2026.pdf" de Drive (propiedad de
  `jhonefrainsanchez@gmail.com`). La única excepción es MARIA CRISTINA ARANGO DE PASTRANA (DANE
  141001001038): no existe ese archivo en Drive (confirmado también por el Informe de Síntesis Municipal
  FEM 2026), así que sus 6 campos llevan una nota honesta en vez de contenido inventado, y no se muestra
  enlace al informe ejecutivo para esa IE.
- Cada respuesta se autoguarda (mismo patrón de debounce que Sesión 1) en la hoja `PreparacionIE`
  (clave compuesta `ID_GRUPO|ID_IE`, igual que `ParticipacionEstamentoIE`). El botón "Enviar aportes de
  esta IE" exige al menos una respuesta no vacía y marca `ENVIADO=SI`.
- Lo que cada IE envía aparece en una tarjeta al inicio de la Sesión 1 ("Aportes de preparación por
  institución"): un `<details>` colapsado por IE, con un párrafo con título por cada pregunta diligenciada
  y un botón de pantalla completa (overlay genérico reutilizable), más un botón para expandir/colapsar
  todas a la vez.
- Este paso es opcional: en cualquier momento el grupo puede pasar directamente a Sesión 1 sin preparar
  aportes por escrito.

## 4.9 Asistencia: solo PDF, un solo método obligatorio; matriz de participación horizontal

- El método "Listado físico" ahora solo admite **PDF** (ya no fotos/imágenes) — validado en el cliente
  (`accept="application/pdf"` + chequeo de `file.type`) y en el servidor (`subirListadoAsistencia` en
  `Asistencia.gs` rechaza cualquier `mimeType` distinto de `application/pdf`).
- Elegir un método es **obligatorio** para continuar desde Participación: el botón "Continuar" valida que
  `estado.metodoAsistencia` sea `QR` o `LISTADO` antes de avanzar; el botón del método elegido queda
  resaltado (`.btn-metodo-activo`).
- Los dos métodos son mutuamente excluyentes en pantalla: si se elige **PDF**, desaparece el conteo de
  firmas en vivo (barra superior de firmantes y el renglón "Total de participantes registrados" de la
  tarjeta) porque no aplica; si se elige **QR**, desaparece el panel de subir PDF.
- La "Cantidad de asistentes por estamento e institución" pasó de una tarjeta por IE a **una sola tabla
  horizontal**: filas = estamento, columnas = IE (encabezado corto "Comuna (N)", con el nombre completo
  de la IE como tooltip), con fila y columna de Total — mismo formato que ya usa el cuadro del informe
  generado (`obtenerMatrizParticipacionGrupo` en Grupos.gs). Sigue siendo editable y con autoguardado
  (ahora por columna/IE, con debounce de 1.5 s), y el backend (`ParticipacionEstamento.gs`) no cambió.
- Esta misma tabla horizontal ahora también aparece, editable, dentro de **Confirmación de
  caracterización** (no solo en Participación) — mismos datos, mismo endpoint, dos contenedores en
  pantalla (`renderParticipacionEstamento(datos, idContenedor)` acepta el id del contenedor).
- Dondequiera que aparece el número de comuna junto a una IE (lista de instituciones en Participación,
  encabezados de la matriz horizontal), ahora se muestra como "Comuna (N)" (`formatearComuna()` en
  Components.html); los valores no numéricos (p. ej. "RURAL") se dejan tal cual, solo capitalizados.

## 4.10 Confirmación de caracterización, preparación obligatoria y Consolidado de Socialización

- **Confirmación de caracterización**: el botón "Confirmo que la información es correcta" ahora permanece
  oculto hasta que TODA la información de la pantalla haya terminado de cargar (ficha del grupo + matriz
  de participación por estamento) — se revela recién cuando ambas cargas terminan.
- **Sesión de preparación — ya NO es opcional**: es condición para continuar a Sesión 1. Los botones
  "Continuar a Sesión 1" (en "¿Qué institución va a preparar sus aportes?" y en la pantalla de
  preparación de una IE) verifican primero, contra el servidor, que TODAS las IE del grupo hayan enviado
  sus aportes; si falta alguna, muestran cuáles y no dejan avanzar.
  - Se corrigió una condición de carrera: al hacer clic en "Enviar aportes" se guarda primero y se espera
    la confirmación del servidor antes de marcar el envío — antes, guardar y enviar se disparaban en
    paralelo y a veces mostraba por error "Responda al menos una pregunta antes de enviar" aunque la
    pantalla ya mostrara el resumen sugerido.
  - Cuando una IE ya envió su preparación (desde este dispositivo u otro, cualquier día — todo queda
    guardado en la nube por `PreparacionIE`), al volver a entrar se muestra un aviso "✅ Ya realizó la
    preparación de esta institución. Puede continuar con Sesión 1" sin bloquear la edición.
  - Las 6 preguntas ahora son colapsables (`<details>`, cerradas por defecto, con un ✓ en el título si ya
    tienen contenido) y su casilla de texto es más grande al desplegarlas — para no llenar tanta pantalla
    con las seis a la vez.
- **Sesión 1 → "Consolidado de Socialización"**: se redujo a solo 4 preguntas —Reflexiones, Desafíos
  comunes, Apuestas compartidas y Conclusiones—, cada una con mínimo 50 y máximo 400 palabras (contador
  en vivo que se pone rojo fuera de rango; el envío definitivo también lo valida en el servidor), más el
  aporte propio del grupo (opcional). "Construcción colectiva del grupo" se mantiene con Prioridades,
  Propuestas colectivas, Acuerdos y Ruta de trabajo (se quitaron Convergencias e Identidad territorial;
  Desafíos y Apuestas se trasladaron a la primera tarjeta). Los campos retirados de pantalla
  (`PROPUESTAS_IE`, `EXPERIENCIAS`, `RETOS`, `APORTES_TERRITORIALES`, `CONVERGENCIAS`, `IDENTIDAD`) se
  conservan como columnas heredadas en `Sesion1Comunal` por si algún grupo ya las había diligenciado, pero
  ya no se muestran ni se piden.
- **Sesión 2 → ConectaEduca reestructurado**: al entrar se explica primero, en una ventana emergente, qué
  es ConectaEduca; solo al cerrarla se despliega el resto de la pantalla. Justo después aparecen dos
  preguntas de GRUPO (no por actor): "Necesidades de articulación" y "Oportunidades identificadas"
  (`NECESIDADES_ARTICULACION_GRUPO`/`OPORTUNIDADES_GRUPO`, guardadas igual que el resto en
  `Sesion1Comunal`). El formulario de cada actor se redujo a Actor o entidad, Tipo de actor, Área de
  interés (ahora una lista de opciones) e Instituciones del grupo interesadas (ahora casillas de
  selección múltiple con las IE del grupo, en vez de texto libre); se quitaron los campos de necesidades,
  oportunidad, alianzas, conexiones y observaciones por actor (columnas heredadas en `ConectaEduca`, ya
  no se piden). Después de la tabla de actores hay una segunda parte de construcción colectiva —
  "Prioridades de las instituciones del grupo (#)", "Acuerdos entre instituciones de las comunas (comunas
  # y rural)", "Propuestas de las instituciones del grupo (#)" y "Ruta de trabajo sugerida por las IE del
  grupo (#)" (`PRIORIDADES_CE`/`ACUERDOS_CE`/`PROPUESTAS_CE`/`RUTA_CE`, también en `Sesion1Comunal`, mismo
  rango de 50-400 palabras, validado también en el servidor) — seguida del aporte propio del grupo
  (opcional) que ya existía.

## 4.11 Código antes que cualquier pantalla, preparación no bloqueante, foto obligatoria, invitados

- **Orden de pantallas**: el código de acceso (`pantallaAcceso`) ahora es la PRIMERA pantalla que se ve —
  antes se mostraban primero Bienvenida/Presentación/Metodología. `ORDEN_PANTALLAS` empieza en
  `pantallaAcceso`; tras validar el código, el recorrido de logos de las IE lleva a Bienvenida →
  Presentación → Metodología (la introducción, que ahora va DESPUÉS de validar) y de ahí a
  Consentimiento/Participación, igual que antes.
- **Comuna sin paréntesis**: `formatearComuna()` ahora muestra "Comuna 7" en vez de "Comuna (7)".
- **Sesión de preparación — ya NO bloquea el paso a Sesión 1**: la IE que termine su preparación puede
  continuar a la siguiente etapa sin esperar a que las demás terminen (antes se bloqueaba a todo el grupo
  hasta que TODAS enviaran, lo cual afectaba también a la que ya había cumplido). El aviso de "faltan por
  diligenciar" sigue existiendo, pero es solo informativo, nunca bloquea el botón "Continuar a Sesión 1".
- **Fotografía general del grupo, obligatoria para finalizar**:
  - Al subirse, se muestra la foto y el texto cambia a "✅ La foto del Grupo X ha sido subida" (se oculta
    el formulario de subida, ya no hace falta volver a subirla).
  - En "Confirmación de caracterización" aparece un pie de foto con la cantidad de participantes
    declarados (`renderPieFotoCaracterizacion`, usa el total de la matriz de participación por estamento).
  - En Participación hay un botón "Subir más tarde" que solo tranquiliza (mensaje: "puedes continuar con
    el instrumento, se te volverá a pedir al finalizar") sin subir nada.
  - En "Informe generado" vuelve a aparecer el mismo formulario de subida (si aún no se subió), y el botón
    final "Continuar" (ahora `btnFinalizarForo`) NO deja pasar a Despedida sin la foto — spec: "No permitir
    finalizar el foro si no se sube la foto".
- **Acceso de invitado (estudiante/acudiente) — sin código**: en la pantalla de Acceso hay un botón "🎓 Soy
  invitado". Elige tipo (Estudiante/Acudiente), elige su institución (de todas las IE reales, vía
  `rpcTodasLasInstituciones`) y entra DIRECTO a la preparación de esa IE en un overlay aparte (fuera de
  `ORDEN_PANTALLAS`/`cambiarPantalla`). Al enviar sus aportes queda en una pantalla final sin ningún botón
  de regreso — "no tiene más acceso" a la aplicación.
  - `Invitados.gs` (nuevo): hoja `InvitadosPreparacion`, `iniciarAccesoInvitado()` (resuelve el grupo de la
    IE con `obtenerGrupoDeInstitucion_`, sin pedir código), `sesionInvitadoValida_()` (token de un solo
    uso, ligado a esa IE y ese dispositivo — nunca pasa por `sesionActivaPorIdGrupo_`, así que un invitado
    JAMÁS puede escribir en Participación, Sesión 1, Sesión 2 ni nada del resto del grupo).
  - `Preparacion.gs` se refactorizó: `_guardarPreparacionIEInterno_`/`_marcarPreparacionEnviadaInterno_`
    contienen la escritura real sin verificación de sesión; `guardarPreparacionIE`/`marcarPreparacionEnviada`
    (sesión de grupo) y las funciones de `Invitados.gs` (sesión de invitado) llaman a lo mismo tras cada
    una validar a su manera.
- **Mensajes y carga**: no fue posible obtener el CSS/gif exacto de FEI 3.1 en este entorno (sin acceso al
  contenido de ese proyecto de Apps Script desde aquí). Se rediseñaron `.mensaje.error/.exito/.info` con
  icono, sombra y animación de entrada, y se agregó `.spinner-cargando` (spinner CSS) como indicador de
  carga en vez de un gif.

## 4.12 Correcciones reportadas en uso real (logo, responsable de envío, invitados, orden de Sesión 1, etiquetas de ConectaEduca, modal de confirmación y sesiones de grupo)

- **Logo de Foro no cargaba**: el archivo de Drive de `LOGO_ENCABEZADO_ID`/`LOGO_PIE_ID` solo era visible
  para quien lo cargó, así que la miniatura (`urlImagenDrive`) fallaba para cualquier otro visitante.
  `asegurarLogosSplashPublicos_()` (`Drive.gs`) llama `setSharing(ANYONE_WITH_LINK, VIEW)` sobre ambos
  archivos una sola vez por instalación (bandera `LOGOS_SPLASH_PUBLICOS`), invocada desde `doGet`.
- **Responsable de envío — botón "Cambiar responsable de envío"**: cuando ya hay un principal guardado, el
  formulario (`formularioResponsablePrincipal`) se colapsa y solo se ve la ficha + el botón
  "✏️ Cambiar responsable de envío" (`filaBotonCambiarResponsable`), que reabre el formulario precargado
  con los datos actuales para editarlos y volver a guardar (sigue siendo un UPSERT en
  `guardarResponsableEnvio`, solo puede haber un principal).
- **Aportes de invitados (estudiantes/acudientes) — ya NO comparten fila con la IE, ni se prellenan del
  informe**: antes, `guardarPreparacionIEInvitado` escribía en la MISMA fila de `PreparacionIE` que la
  respuesta oficial de la IE (riesgo de que se sobrescribieran entre sí) y `obtenerPreparacionIEInvitado`
  heredaba el prellenado con el resumen sugerido del Informe Ejecutivo. Ahora:
  - Nueva hoja `AportesInvitadosPreparacion` (una fila por sesión de invitado, clave `TOKEN_INVITADO`),
    completamente independiente de `PreparacionIE`.
  - `PREGUNTAS_PREPARACION_INVITADO_` (`Invitados.gs`) usa el MISMO título de cada pregunta (no se
    modifica la redacción) pero con una `ayuda` propia — explicación en lenguaje sencillo + un ejemplo de
    respuesta — en vez de remitir al Informe Ejecutivo; las preguntas del invitado arrancan siempre en
    blanco (construcción libre).
  - `obtenerAportesInvitadosIE(idGrupo, idIE)` expone los aportes YA ENVIADOS por estudiantes/acudientes
    de esa IE; se muestran en la propia pantalla de Preparación de la IE (`listaAportesInvitadosPreparacionIE`,
    tarjeta "🧑‍🎓👨‍👩‍👧 Aportes de estudiantes y acudientes invitados"), como insumo de referencia — no se
    copian automáticamente en la respuesta oficial.
- **Sesión 1 — orden de tarjetas**: "📎 Archivos de apoyo" ahora aparece ANTES de "🗣️ Aportes de
  preparación por institución" (antes era al revés).
- **ConectaEduca — etiquetas dinámicas sin "(#)"**: los 4 campos de la segunda parte ("Prioridades...",
  "Acuerdos...", "Propuestas...", "Ruta de trabajo...") mostraban literalmente "grupo (#)" / "(comunas #
  y rural)". Ahora `actualizarEtiquetasConectaEducaGrupo()` (`JS.html`) reemplaza esos `<span>` con el
  número real del grupo (extraído de `estado.grupo`, p.ej. "Grupo 3") y la lista real de comunas de las
  IE del grupo (`textoComunasGrupoActual_()`, p.ej. "comunas 1, 2 y zona rural"), sin paréntesis.
- **Modal de confirmación propio, en vez de `window.confirm()`**: dentro del iframe con sandbox de
  HtmlService, `confirm()` puede verse feo (diálogo nativo del navegador) o directamente no mostrarse/no
  hacer nada según el navegador (bloqueado en silencio) — causa más probable de "clic en Enviar Sesión 2 y
  no pasa nada". Se agregó `#modalConfirmar` (`Modal.html`, estilo del propio aplicativo) y el helper
  `confirmarAccion(mensaje, callback, tituloOpcional)` (`JS.html`), que reemplaza los 5 `confirm()` que
  quedaban (enviar aportes de invitado, Sesión 1, Sesión 2/ConectaEduca, preparación de una IE, generar
  informe).
- **"Esta sesión ya no está activa en este dispositivo" espurio**: el cupo de sesión por grupo
  (`Session.gs`) tenía un tope fijo de 4 dispositivos SIN expiración por inactividad — insuficiente para
  un GRUPO de hasta 6 IE, cada una conectada desde su propio dispositivo durante toda la jornada, así que
  dispositivos activos pero más antiguos terminaban siendo expulsados por dispositivos nuevos. Ahora:
  - `podarSesionesInactivas_()` quita del arreglo los cupos NO PRINCIPALES sin actividad en los últimos
    `TIEMPO_SESION` minutos (antes un valor muerto, ahora usado) antes de evaluar el tope — se llama en
    cada `reclamarSesionGrupo_`/`mantenerSesionGrupo` (heartbeat cada 30s), así que el tope casi nunca se
    llega a activar en uso normal.
  - `MAX_SESIONES_SIMULTANEAS_GRUPO` sube de 4 a 10 por defecto; `asegurarLimiteSesionesGrupoRazonable_()`
    (llamada desde `doGet`) sube automáticamente, una sola vez, las instalaciones que ya tenían el valor
    antiguo (4) guardado en `ConfiguracionComunal`.
  - El heartbeat (`iniciarHeartbeat`, `JS.html`) ya NO interrumpe con un mensaje de error si un tick falla
    — si de verdad se perdió el cupo, la siguiente acción de escritura (guardar/enviar) lo reporta con un
    mensaje concreto en su propio contexto, en vez de una alerta genérica en cualquier pantalla.
- **"Ocurrió un error de comunicación con el servidor" al enviar**: cualquier excepción no controlada del
  servidor (timeout, bug real) llegaba al cliente como ese mensaje genérico, sin pista de la causa. Se
  agregó `ejecutarRpcSeguro_(fn)` (`Utils.gs`) y se envolvió CADA función `rpcXxx` de `Code.gs` con ella:
  ahora una excepción no prevista se registra en los logs de Apps Script (Ejecuciones) y se devuelve al
  cliente como `{ok:false, mensaje:"Ocurrió un error inesperado en el servidor..."}`, un mensaje que cada
  pantalla ya sabe mostrar en su propio contenedor de mensaje.

## 4.13 Segundo lote de correcciones en uso real: invitados, matriz de participación, asistencia QR, informe, pantalla completa editable y guardado inline

- **Invitados — mensaje del Informe Ejecutivo, perfiles y ayuda con ejemplos de contraste**: el aviso "No se
  encontró el Informe Ejecutivo..." se reformuló en tono propositivo (solo queda en la pantalla de
  Preparación oficial de la IE, ya no en la de invitado). Los perfiles pasan a llamarse "Estudiantes y
  egresados" y "Adulto responsable de un(a) estudiante" en toda la interfaz. `AYUDA_PREPARACION_INVITADO_ESTUDIANTE_`/
  `AYUDA_PREPARACION_INVITADO_ADULTO_` (`Invitados.gs`) dan, para cada pregunta y cada perfil, un ejemplo de
  respuesta afirmativa Y uno negativo (para no insinuar una única "respuesta correcta"), y piden escribir
  "no sé" en vez de dejar la pregunta en blanco.
- **Modal de confirmación e invitado finalizado**: `.modal-fondo` subía de z-index (300, por encima de
  `.overlay-invitado`, 260) — antes el modal de confirmación quedaba invisible/inaccesible dentro del flujo
  de invitado, causa más probable de que "no hiciera nada". La pantalla final de invitado agrega un botón
  "✕" grande (`window.close()`, mejor esfuerzo — límite del navegador si la pestaña no fue abierta por
  script) y cambia el texto a "Espera a que se inicie la socialización para que puedas leer tus aportes a
  este foro."
- **Sección "Invitados" en Selección de IE**: nueva hoja consultada vía `obtenerResumenInvitadosGrupo`
  (`Invitados.gs`) — resumen (conteo) de aportes ya enviados por estudiantes/egresados y adultos
  responsables, por IE, puramente informativo. Título de la pantalla actualizado a "¿Qué institución o
  participante va a preparar sus aportes?".
- **Matriz de participación por estamento**: cada columna ahora muestra el nombre de la IE (no solo la
  comuna) — `.th-ie`/`.nombre-ie-columna`/`.comuna-ie-columna` (`Components.html`/`CSS.html`). Al hacer clic
  en "Continuar" desde Participación, si alguna IE quedó con la columna en cero, se pregunta con el modal de
  confirmación (`confirmarAccion`, ahora con etiquetas de botón personalizables) si la información es
  correcta o si hay que corregirla.
- **Rector(a) sin ícono de guardado**: el botón 💾 se quitó; el campo se guarda solo al salir de él (evento
  `change`, delegado en `document`), igual patrón de autoguardado que el resto de la app.
- **Asistencia QR**: arriba del panel aparece "Firmantes / participantes declarados" (comparación rápida,
  `actualizarComparativoFirmantesQR_`); el botón "Ver QR en pantalla completa" queda en el centro, y tanto
  el panel normal como la vista de pantalla completa ahora incluyen el enlace y un botón "Copiar enlace". La
  firma por QR se cierra en cuanto el informe del grupo se genera (`ESTADO = "INFORME_GENERADO"`,
  verificado en `registrarAsistenciaPublica` y en `paginaAsistenciaGrupo_`, que oculta el formulario público
  y muestra un aviso de cierre).
- **Informe generado**: se agregó un listado nominal de asistentes que firmaron (nombre, institución,
  estamento) además del conteo/gráfico agregado que ya existía.
- **Pantalla completa editable (👁️) y guardado inline (💾)**: `agregarAccionesTextareas_` (`JS.html`)
  agrega, después de cada pregunta de texto largo (Sesión 1, Sesión 2, Preparación IE, invitado), un botón
  👁️ que abre un editor de pantalla completa (mismo campo, sincronizado en vivo; se cierra con "✕" o con
  ESC) y, solo para Sesión 1/2, un botón 💾 que guarda esa pregunta de inmediato con un texto auxiliar de
  confirmación (verde si fue exitoso). El botón "Guardar avance" al final de Sesión 1 se eliminó — el
  guardado sigue siendo automático (debounce de 2s) y ahora también inmediato por pregunta.
- **Texto de apoyo por pregunta**: las 8 preguntas del Consolidado de Socialización/Construcción colectiva
  (Sesión 1) ahora tienen una explicación breve debajo del título, igual que ya tenían las de ConectaEduca.

## 4.14 Página pública de asistencia QR: estamentos en singular, consentimiento grande, correo con autocompletado/validación, y pantalla de agradecimiento

- **Estamentos en singular**: el `<select>` de estamento de `AsistenciaPublica.html` pasa de "Docentes",
  "Estudiantes", "Padres/madres/acudientes", "Egresados", "Otros" a "Docente", "Estudiante",
  "Padre/madre/acudiente", "Egresado(a)", "Otro" — cada persona firma por sí misma, no por su categoría (la
  hoja `ESTAMENTOS_PARTICIPACION_` de `ParticipacionEstamento.gs`, que sí es un conteo por categoría, se
  deja igual).
- **Consentimiento como botón grande**: el checkbox nativo (diminuto en celular) se reemplazó por un botón
  ancho `.boton-consentimiento` con un ícono grande que cambia a ✅ (y fondo/borde verdes) al aceptar.
- **Correo — autocompletado y validación**: al escribir `@g` se completa a `@gmail.com`, al escribir `@h` a
  `@hotmail.com` (según el patrón exacto `usuario@g`/`usuario@h`, para no interferir con otros dominios).
  Si el correo no tiene formato válido (regex simple `usuario@dominio.algo`), se avisa "Corrige el formato
  del correo electrónico" tanto al salir del campo como al intentar confirmar — el correo sigue siendo
  opcional, la validación solo aplica si se escribió algo.
- **Pantalla de agradecimiento y "registrar otra asistencia"**: al confirmar, el formulario se reemplaza por
  una tarjeta "¡Gracias por su asistencia!" con "Puede continuar en la preparación del Foro Comunal 2026."
  y un enlace a la app principal (`URL_APP_BASE`, agregado por `paginaAsistenciaGrupo_` en `Asistencia.gs` —
  ahí cualquiera puede entrar como invitado sin código), más un botón "➕ Registrar otra asistencia" que
  reinicia el formulario para la siguiente persona en el mismo dispositivo (uso típico: alguien pasa el
  celular de mano en mano para firmar).

## 4.15 Tercer lote de correcciones en uso real: firmantes visibles, correo en responsable de envío, matriz más ancha, gate de preparación reintroducido, rol del responsable de preparación, y guardado inline en Preparación IE/invitado

- **Firmantes visibles tras confirmar (AsistenciaPublica.html)**: la tarjeta de agradecimiento ahora muestra
  "👥 N personas han firmado en este grupo hasta el momento" (actualizado por el mismo `actualizarContador()`
  que ya refrescaba el contador de arriba) — antes solo se actualizaba fuera de la vista, sin que quien
  acaba de firmar lo notara.
- **Responsable de envío — correo con autocompletado y validación**: `activarAutocompletadoCorreo_`/
  `correoValido_` (`JS.html`) reutilizan el mismo patrón de `AsistenciaPublica.html` (`@g` → `@gmail.com`,
  `@h` → `@hotmail.com`, validación de formato) en `campoCorreoResponsable` y `campoCorreoAsistente`.
- **Matriz de participación — estructura más ancha**: `main { max-width: 1100px; }` (antes 780px) para que la
  tabla de participación por estamento e institución (hasta 6 IE + Estamento + Total) quepa sin
  desplazamiento horizontal en pantallas medianas/grandes. La advertencia de columna en cero
  (`confirmarAccion`) se agregó también al botón "Confirmo que la información es correcta" de Confirmación
  de caracterización (antes solo estaba en el botón Continuar de Participación, pero la tabla es editable en
  ambas pantallas).
- **Gate de preparación reintroducido**: "Continuar a Sesión 1" (desde Selección de IE y desde Preparación de
  una IE) vuelve a estar oculto hasta que TODAS las IE del grupo enviaron sus aportes de preparación
  (`todasLasIEPrepararon_`/`actualizarBotonesContinuarSesion1_`, `JS.html`) — revierte la decisión "ya NO
  bloquea" de la sección 4.11, a pedido explícito del usuario.
- **Rol de quien prepara los aportes**: el modal "¿Quién va a diligenciar los aportes?" (`modalResponsablePreparacion`,
  `Modal.html`) agrega un select de rol (mismo catálogo `ROLES_FORO_`) debajo del nombre; el campo
  `RESPONSABLE` guardado combina "Nombre — Rol".
- **Aportes de invitados — texto más directo**: el aviso de la tarjeta de aportes de invitados en
  Preparación IE cambia a "Los aportes de los invitados no se han incluido en los aportes institucionales.
  Estos serán leídos de manera independiente."
- **Preparación IE e invitado — guardado inline**: se eliminaron los botones "Guardar avance" de ambas
  pantallas (igual que ya se había hecho en Sesión 1); `agregarAccionesTextareas_`/`configGuardadoParaTextarea_`
  (`JS.html`) generalizan el botón 💾 + texto de confirmación para las tres clases de pregunta
  (`data-campo-sesion1`, `data-prep-pregunta`, `data-inv-prep-pregunta`), cada una usando su propia función
  de guardado y su propio temporizador de autoguardado.
- **Editor de pantalla completa más amplio**: `#pantallaCompletaEdicion` pasa a ocupar el 80% del ancho de
  pantalla (antes 720px fijos) y el `<textarea>` el 80% del alto, para escribir más cómodo — el visor de
  solo lectura de la socialización (`pantallaCompletaTexto`) se deja igual, más angosto para lectura.

## 4.16 Cuarto lote: foto obligatoria con aviso, aportes de invitados agregados, ayuda con el nombre real de la IE, y "Continuar a Sesión 1" por institución

- **Foto general del grupo — aviso explícito si no se decide**: al hacer clic en "Continuar" desde
  Participación sin haber subido la foto NI haber confirmado "Subir más tarde", aparece un
  `confirmarAccion` preguntando "¿Deseas subirla ahora o continuar y subirla más adelante?" — clic en
  "Subir ahora" desplaza la pantalla al formulario de foto; clic en "Continuar, subirla más adelante"
  marca la decisión (`estado.fotoGrupoMasTardeConfirmado`) y reintenta el mismo botón.
- **Aportes de invitados — agregados por grupo, mismo formato que las IE, ya no se repiten por
  institución**: `obtenerAportesInvitadosGrupo` (`Invitados.gs`, reemplaza a `obtenerResumenInvitadosGrupo`
  + `obtenerAportesInvitadosIE`) devuelve el contenido completo (no solo conteos) agrupado por IE;
  `renderAportesInvitadosGrupo` (`Components.html`) lo pinta con el mismo `<details class="socializacion-ie">`
  que los aportes institucionales de Sesión 1, sin el encabezado "Invitados" (queda solo el texto
  explicativo). La tarjeta que antes se repetía dentro de la pantalla de cada IE individual
  (`tarjetaAportesInvitadosPreparacionIE`) se eliminó — los aportes de invitados ahora solo aparecen en
  este listado agregado de Selección de IE.
- **Ayuda de las preguntas — nombre real de la institución**: las ayudas de `PREGUNTAS_PREPARACION_`
  (`Preparacion.gs`) que mencionan el Informe Ejecutivo llevan el marcador `{{IE}}`, sustituido por
  `preguntasPreparacionConNombreIE_()` con el nombre real de la institución en `obtenerPreparacionIE` — p.
  ej. "(Informe Ejecutivo (IE) 2026)" pasa a "(Informe Ejecutivo de la SANTA LIBRADA 2026)".
- **"Continuar a Sesión 1" — por institución, sin esperar a las demás**: revierte la sección 4.15 (aclarado
  explícitamente por el usuario: "aparece por IE individual, sin esperar a las demás"). Desde la pantalla
  de una IE, el botón aparece en cuanto ESA IE envía sus aportes; desde Selección de IE, aparece en cuanto
  al menos una IE del grupo ya envió (`algunaIEPreparo_`, `JS.html`) — ya no exige que todas terminen.

## 4.17 Quinto lote: transición de logos invertida, overlay de carga con imagen propia en 3 momentos, y blindaje del botón "Continuar a Sesión 1"/mensaje de error en la pantalla de aportes

- **Splash inicial — orden invertido (alcaldía primero, #Foro después)**: `iniciarSplashInicial()` (`JS.html`)
  mostraba `[LOGO_ENCABEZADO_ID, LOGO_PIE_ID]` (Foro primero); pasa a `[LOGO_PIE_ID, LOGO_ENCABEZADO_ID]` —
  el usuario compartió el enlace de Drive del logo "#Foro" (mismo `fileId` que `LOGO_ENCABEZADO_ID` en
  `Config.gs`) pidiendo que se muestre "luego de logo de la alcaldía".
- **Overlay de carga con imagen propia en 3 momentos concretos** (`overlayCargaAccion`, `Index.html`;
  `mostrarCargaAccion_`/`ocultarCargaAccion_`, `JS.html`; `.overlay-carga-accion`, `CSS.html`):
  1. Al validar el código de acceso (`intentarValidarAcceso`).
  2. Al subir la foto del grupo, en los dos momentos disponibles — Participación y pantalla de Informe
     generado (`subirFotoGrupoDesde_`, función compartida por ambos botones).
  3. Al guardar el responsable de envío (`rpcGuardarResponsable` con tipo `PRINCIPAL`).

  Para el tercer momento se usa el sticker real pedido por el usuario ("Cargando Sticker by dipielbella",
  `https://media.giphy.com/media/SYIu9YMtvUc6tBzDH6/giphy.gif` — localizado por búsqueda web; el entorno
  de código no pudo verificar la carga del CDN de Giphy por una política de red propia del sandbox, no
  del navegador de quien usa la app; si no carga, `img.onerror` hace caer el overlay de vuelta al spinner
  genérico ya existente). Para el primer y segundo momento el usuario adjuntó dos imágenes directamente en
  el chat (un gif de una carrera y un ícono de obturador de cámara) que un entorno de código no puede leer
  como archivo — solo enlaces o archivos reales son accesibles — así que se usan animaciones propias con
  emoji (🏃 y 📷) + una barra de progreso indeterminada, mismo overlay y estilo; quedan listas para
  sustituirse por las imágenes reales en cuanto el usuario comparta un enlace de Drive para cada una.
- **Diagnóstico de "Ocurrió un error de comunicación con el servidor" y "Continuar" ausente en Selección de
  IE (Preparación)**: se auditó exhaustivamente cada nombre de RPC llamado desde el cliente contra los
  definidos en `Code.gs` (sin discrepancias — se descarta un nombre de función obsoleto) y se revisó la
  serialización de todas las RPC involucradas (todas envueltas en `ejecutarRpcSeguro_`). La causa más
  probable identificada: `rpcPreparacionesEnviadasGrupo` devuelve un arreglo en el camino normal, pero
  `{ok:false, mensaje}` si el servidor atrapó una excepción — sin protección, el `.forEach` posterior
  lanzaba un `TypeError` silencioso (no capturado por `withFailureHandler`, que solo actúa en fallos de
  transporte) que dejaba a medias el pintado de la pantalla, y con él, `actualizarBotonesContinuarSesion1_()`
  nunca se ejecutaba — explicando por qué "Continuar" seguía sin aparecer pese a la lógica por IE ya
  implementada. `cargarListaIEPreparacion()` (`JS.html`) ahora valida `Array.isArray`, envuelve el pintado
  en `try/catch`, garantiza que `actualizarBotonesContinuarSesion1_()` se ejecute siempre (incluso en el
  `onError` de transporte) y muestra el mensaje de error real en `mensajePreparacionPendiente` cuando
  `ok === false`. `cargarPreparacionIE()` recibió la misma guarda para `rpcObtenerPreparacionIE`.

## 4.18 Sexto lote: cambiar fotografía en caracterización, gifs reales adicionales, acceso invitado por grupo, caracterización de invitados + informe consolidado, gates de Sesión 2, valoración de 4 preguntas también pública, y rediseño de Revisión/Informe generado/última pantalla

- **Confirmación de caracterización — cambiar/agregar fotografía**: nuevo botón "✏️ Cambiar fotografía" /
  "📷 Agregar fotografía" (según haya o no una ya subida) con su propio formulario (`formularioFotoGrupoCaracterizacion`),
  reutiliza `subirFotoGrupoDesde_` y `refrescarVistasFotoGrupo_` (`JS.html`).
- **"Verifique la participación por institución" — solo en la confirmación, no antes**: se quitó la
  comprobación de `btnContinuarAParticipacion` (Participación); queda únicamente en
  `btnConfirmarCaracterizacion` (Confirmación de caracterización).
- **6 gifs reales de Drive con texto de apoyo**, reemplazando/ampliando el overlay de carga del lote
  anterior (`GIFS_CARGA_ACCION_`, `JS.html`; todos con permiso público "cualquiera con el enlace"):
  acceso (tras el código), foto (subir/cambiar), aportes (enviar preparación IE/invitado), sesionDefinitiva
  (enviar Sesión 1/2), valoracionAsistencia (valoración y listado de asistencia PDF), informe (generar
  informe). El de responsable de envío sigue usando el sticker de Giphy del lote anterior.
- **Acceso invitado — Grupo primero, IE después**: nuevo paso `overlayInvitadoGrupo` (spec: "no debe
  aparecer IE en el listado... si no grupos 1 a 6") entre elegir tipo de invitado y elegir institución; la
  IE se filtra por el grupo elegido.
- **Caracterización de invitados + informe consolidado**: nuevo paso `overlayInvitadoCaracterizacion`
  (nombre, edad, sexo y, según el perfil, rol/años estudiando o vínculo/rol en la IE) — nunca se muestra en
  pantalla, solo al final en el informe del grupo (`guardarCaracterizacionInvitado`,
  `obtenerCaracterizacionInvitadosGrupo_`, `Invitados.gs`; nueva sección "Informe consolidado de invitados"
  en `Informes.gs` con conteos de estudiantes/graduados/adultos y la tabla nominal).
- **Documentos de apoyo — solo en Preparación, no en Socialización**: se quitó el `abrirModal("modalRecursosSesion1")`
  automático de `pantallaSesion1`; ahora se abre al entrar a `pantallaPreparacionIE`. En Sesión 1 queda
  solo la sección "Archivos de apoyo para esta sesión" (sin ventana emergente).
- **Aportes de preparación — aviso de cambios + instrucción de copiar/pegar**: el texto de
  "🗣️ Aportes de preparación por institución" añade que los cambios surgidos en la socialización se
  discuten en la siguiente sesión, y que seleccionar un fragmento de cualquier IE lo deja listo para pegar
  en las preguntas de "Construcción colectiva del grupo".
- **"Construcción colectiva del grupo" — título único de Reflexiones a Ruta de trabajo**: se fusionaron las
  dos tarjetas de preguntas de Sesión 1 en una sola, con un único `<h2>` que cubre las 8 preguntas.
- **Gates de Sesión 2**: "Continuar a Sesión 2" (`btnContinuarASesion2`) permanece deshabilitado hasta que
  `rpcEstadoGrupo` confirme `sesion1Enviada`; el envío de Sesión 2 valida en cliente
  (`camposSesion2Faltantes_`) los 4 campos que el servidor ya exigía (`CAMPOS_SESION2_OBLIGATORIOS_`,
  `ConectaEduca.gs`), evitando un viaje al servidor solo para el mismo rechazo.
- **Valoración — 4 preguntas de satisfacción sobre el desarrollo del Foro y la comunicación entre pares**:
  se reescribió el enunciado de las 4 preguntas de corazones (antes copiadas de FEI 3.1); misma valoración,
  anónima y por asistente, disponible ahora también en la página pública de asistencia QR
  (`ValoracionAsistentesPublica`, `guardarValoracionAsistentePublica`, `Valoracion.gs`;
  `rpcGuardarValoracionAsistentePublica`, `Code.gs`; formulario en `AsistenciaPublica.html`).
- **Método de asistencia — bloqueo de QR a PDF**: una vez elegido QR ya no se puede cambiar a listado en
  PDF (validado en cliente, `actualizarBloqueoMetodoAsistencia_`, y en servidor, `guardarMetodoAsistencia`,
  `Asistencia.gs`); de PDF a QR sigue permitido en cualquier momento.
- **Revisión y cierre — verificación de firmantes**: nueva sección "Verifique que todos los participantes
  hayan firmado la asistencia" con conteos y, si el método es QR, el visor del código y el enlace de nuevo
  (`cargarVerificacionFirmantesCierre_`, `JS.html`).
- **Asistencia cerrada al generar el informe**: el encabezado permanente de firmantes (`panelFirmantes`)
  pasa a mostrar "🔒 La asistencia se ha cerrado" y el botón cambia a "Ver listado de asistencia"
  (`actualizarPanelFirmantesCerrado_`, disparado al generar el informe y restaurado con `rpcEstadoGrupo`
  al reingresar); la página pública de asistencia ya mostraba este mismo mensaje cuando el estado es
  `INFORME_GENERADO`.
- **Fotografía — ya no se pide en "Informe generado", último punto de exigencia es "Generar informe"**: se
  quitó la tarjeta de foto de `pantallaInformeGenerado`; el gate pasó a `generarInformeCompletoGrupo`
  (`Grupos.gs`, nuevo `obtenerFotoGrupoId_`, `Asistencia.gs`) y a un aviso en cliente antes de la
  confirmación de generar informe. La fotografía se incluye ahora dentro del documento del informe, junto
  al listado de asistentes (`Informes.gs`).
- **Última pantalla — 4 pasos y "Finalizar Foro Comunal Educativo 2026"**: `pantallaInformeGenerado`
  reemplaza el botón genérico "Continuar" por una lista de 4 pasos (verificar asistentes, descargar
  informe, enviarlo por correo, finalizar) y el botón "Finalizar Foro Comunal Educativo 2026"
  (`btnFinalizarForo`), que se deshabilita tras el primer clic.

## 4.19 Séptimo lote: gifs como base64 (fin del hotlinking de Drive), envíos abiertos a todo el grupo con un solo envío total, cambiar foto, QR↔PDF flexible, copiar al seleccionar aportes, y nuevos títulos de Bienvenida/Ruta del Foro

- **Gifs de carga — base64 embebido, ya no hotlink a Drive**: `drive.google.com/uc?export=view&id=...`
  dejó de ser confiable (Google fue restringiendo el hotlinking de archivos de Drive por la desactivación
  de cookies de terceros) y por eso solo se veía el círculo de carga genérico, nunca el gif. Los 6 gifs
  (acceso, foto, aportes, sesionDefinitiva, valoracionAsistencia, informe) se descargaron una sola vez y
  se embebieron como `data:image/gif;base64,...` en un archivo nuevo, `GifsData.html`
  (`var GIFS_DATAURI_`), incluido en `Index.html`; `mostrarCargaAccion_` (`JS.html`) ya no construye
  ninguna URL de Drive. Además, la pantalla "¿Qué institución o participante va a preparar sus aportes?"
  ahora muestra el gif "acceso" centrado con texto de apoyo mientras carga la lista de instituciones y el
  resumen de aportes (`cargarListaIEPreparacion`).
- **Sesión 1/Sesión 2/Informe — cualquiera puede enviar, pero un solo envío total**: se quitó la
  restricción de "solo el responsable principal" en `enviarSesion1Definitiva` (`Sesion1.gs`),
  `enviarSesion2Definitiva` (`ConectaEduca.gs`) y `generarInformeCompletoGrupo` (`Grupos.gs`) — ahora
  basta con tener sesión activa en el grupo. Cada uno de los tres solo se concreta una vez: Sesión 1/2
  devuelven `{codigo:"YA_ENVIADO"}` si ya se enviaron, y el informe es idempotente (si ya existe, se
  devuelve el mismo, sin regenerarlo, para que quien llegue después pueda seguir descargándolo y
  enviándolo por correo con normalidad). En cliente, `actualizarEstadoEnvioSesion1_`/`_2_` muestran un
  aviso "✅ ya fue enviada" y deshabilitan el botón correspondiente, sincronizado con `rpcEstadoGrupo` al
  entrar a cada pantalla y justo después de un envío exitoso; `actualizarBotonGenerarInforme` ya no exige
  ser el responsable principal, solo la valoración enviada y que el informe no exista todavía.
- **Cambiar fotografía del grupo**: la vista de "foto ya subida" (`renderFotoGrupo`, `Components.html`)
  ahora incluye un botón "Cambiar fotografía" (`data-cambiar-foto`) que reabre el formulario de subida
  correspondiente (nuevo manejador delegado en `JS.html`).
- **Método de asistencia — QR→PDF permitido mientras no haya informe**: se relajó el bloqueo del lote
  anterior; `guardarMetodoAsistencia` (`Asistencia.gs`) y `actualizarBloqueoMetodoAsistencia_` (`JS.html`)
  ahora solo impiden el cambio de QR a PDF cuando el estado del grupo ya es `INFORME_GENERADO`.
  Documentos de apoyo en Preparación: se revisó todo el flujo (`abrirModalResponsablePreparacion` →
  `cargarPreparacionIE` → `abrirModal("modalRecursosSesion1")`) sin encontrar una causa reproducible; se
  añadió de todas formas manejo defensivo de errores en `rpcObtenerRecursosSesion1`.
- **Copiar automático al seleccionar texto en los aportes de preparación**: al soltar una selección
  dentro de "🗣️ Aportes de preparación por institución" se copia al portapapeles
  (`navigator.clipboard.writeText`, con `document.execCommand("copy")` de respaldo) y aparece un aviso
  centrado y sutil "Texto copiado" (`avisoTextoCopiado`, `.aviso-texto-copiado`) que se desvanece solo.
- **"Aporte propio" al final de Sesión 1**: el desplegable opcional se movió de después de "Conclusiones"
  a justo antes del botón "Enviar Sesión 1 de forma definitiva", después de "Ruta de trabajo".
- **ConectaEduca — cantidad de actores agregados**: `renderTablaConectaEduca` (`Components.html`) ahora
  también actualiza `contadorActoresCE` con el total de actores/entidades ya registrados por el grupo,
  visible justo debajo del botón "Agregar actor".
- **Nuevos títulos — Bienvenida y Metodología**: "Bienvenida" pasó a "Querida Comunidad Educativa del
  [grupo]" (el nombre del grupo se rellena tras validar el código de acceso, `nombreGrupoBienvenida`);
  "Metodología" pasó a "Ruta del Foro Educativo Comunal", con la lista de pasos corregida para reflejar
  el recorrido real de pantallas de esta entrega (consentimiento y método de asistencia, participación,
  confirmación de caracterización, sesión de preparación, Sesión 1, Sesión 2/ConectaEduca, revisión y
  cierre, y generación del informe).

## 4.20 Octavo lote: valoración obligatoria antes de firmar en la asistencia pública, mejora condicional, cierre/verificación de listado, y retomar el recorrido tras perder la señal o cambiar de dispositivo

- **Asistencia pública — la valoración va antes de poder firmar**: en `AsistenciaPublica.html`, el
  formulario de "Registrar mi asistencia" queda oculto tras un aviso ("Para poder registrar su
  asistencia, complete primero la valoración del Foro") hasta que la valoración se envíe
  (`actualizarBloqueoAsistenciaPorValoracion_`) — se exige aunque ya se haya marcado el consentimiento,
  porque el consentimiento está dentro del bloque que permanece oculto.
- **Valoración pública — textarea de mejora condicional**: igual que en la app principal
  (`inicializarValoracion`, `Index.html`), cada una de las 4 preguntas de corazones despliega un textarea
  opcional ("¿Nos dice brevemente cómo podría mejorar esto?") solo cuando la respuesta es de 1 o 2
  corazones (`data-mejora-pub-de`). Las 4 respuestas de mejora se guardan en `ValoracionAsistentesPublica`
  (columnas nuevas `P1_MEJORA`.._MEJORA, `Valoracion.gs`).
- **Pantalla final de asistencia — nuevo mensaje y botones**: "Puede continuar en la preparación del
  Foro Comunal 2026" (con un enlace a la app) se reemplazó por "Gracias por sus aportes. Puede continuar
  en el desarrollo del Foro Educativo Comunal, Neiva, 2026." y dos botones nuevos, "Cerrar asistencia" (
  muestra un aviso de despedida y detiene el sondeo) y "Verificar listado" (despliega el listado de
  firmantes en vivo del grupo, actualizado cada 4 s vía `rpcEstadoFirmantes`, reutilizado tal cual de la
  app principal).
- **"Registrar otra asistencia" también borra la valoración**: `reiniciarFormularioAsistencia_` ahora
  llama a `reiniciarValoracionPublica_` (limpia corazones, textareas de mejora y el estado enviado), así
  que la siguiente persona que use el mismo dispositivo debe volver a valorar antes de poder firmar.
- **Retomar el recorrido tras perder la señal o cambiar de dispositivo**: nueva columna `ULTIMA_PANTALLA`
  en `AccesosGrupo` (compartida por todo el grupo, no por dispositivo) que se actualiza en cada
  `cambiarPantalla()` real (`rpcGuardarUltimaPantalla`, `Access.gs`/`Code.gs`) — la introducción
  (Bienvenida/Presentación/Metodología) nunca se marca ni se retoma directamente en ella. Al validar el
  código de acceso, `destinoResumenPantalla_` decide si saltar el carrusel de bienvenida y la
  introducción para ir directo a la última pantalla registrada; las pantallas que dependen de una
  selección que solo vive en el dispositivo (ej. cuál IE se estaba preparando) se remapean a un paso
  anterior seguro (`MAPA_RESUMEN_PANTALLA_`: `pantallaPreparacionIE` → `pantallaSeleccionIEPreparacion`).
  Así, si la señal se cae o alguien continúa desde otro dispositivo, basta con volver a ingresar el
  código de acceso del grupo para retomar donde se quedó, sin repetir pantallas ya superadas.
- **Verificado (sin cambios necesarios)**: el conteo de "personas han firmado" que ve todo el grupo
  (`rpcEstadoFirmantes`/`contarParticipantesGrupo`) ya se actualiza automáticamente cuando alguien firma
  desde la asistencia pública, porque ambas rutas escriben en la misma hoja `ParticipacionComunal`
  (`registrarParticipante`, `Data.gs`) y el panel de firmantes de la app principal ya la sondea cada 15 s.

## 4.21 Noveno lote: restaurar tildes/ñ en los resúmenes de preparación, cuadros de texto configurados para español, vista de solo lectura tras enviar Sesión 1/2, y advertencia + confirmación antes de generar el informe

- **Tildes y ñ restauradas en `RESUMENES_PREPARACION_IE_`**: 17 de las 36 instituciones (más el texto de
  "no se encontró el archivo") habían perdido todas las tildes y la ñ en sus 6 párrafos (p1..p6) — un
  subagente de los que compiló ese contenido en el lote de preparación entregó el texto sin
  diacríticos. Se restauraron con un script de sustitución por diccionario (sustantivos en -ción/-sión,
  esdrújulas, hiatos en -ía, ñ, y un puñado de verbos en pretérito y casos ambiguos —"más", "aún", "sí"—
  verificados uno por uno por contexto antes de corregirlos) sobre las 17 instituciones afectadas; las
  19 restantes, ya correctas, quedaron intactas porque el diccionario solo sustituye ortografías
  incorrectas literales.
- **Cuadros de texto configurados para español**: todos los `<textarea>` y los `<input type="text">`
  de texto libre (nombre, aporte propio, actor de ConectaEduca, nombre de invitado, nombre en la
  asistencia pública) llevan ahora `lang="es" spellcheck="true"` explícitos — se excluyó a propósito el
  código de acceso (`campoCodigoAcceso`), que es alfanumérico, no prosa.
- **Sesión 1/2 ya enviada — vista de solo lectura, no cuadros en blanco**: al entrar a `pantallaSesion1`/
  `pantallaSesion2` (o justo después de un envío exitoso) con la sesión ya enviada, `camposEditablesSesion1`/
  `camposEditablesSesion2Previos`+`camposEditablesSesion2CE` (todos los `<textarea>`/`<input>` y el botón
  de envío) se ocultan por completo y en su lugar aparece un mensaje grande y centrado
  (`.mensaje-envio-definitivo`) seguido de las respuestas ya registradas, renderizadas en solo lectura
  (`renderRespuestasSoloLectura_`, `JS.html`) — nunca vuelve a mostrarse el cuadro en blanco para
  reescribir. `cargarSesion1()` ahora también se llama al entrar a `pantallaSesion2` (antes solo se
  llamaba desde `pantallaSesion1`), necesario para que esos campos tengan su valor real si se llega
  directo a Sesión 2 (p. ej. al retomar el recorrido tras perder la señal, lote anterior).
- **Advertencia + confirmación antes de generar el informe**: en "Revisión y cierre" aparece un aviso
  distinto según el método de asistencia del grupo — si es QR: "Verifique que todos los participantes
  hayan firmado la asistencia. Al generar el informe, se cerrará la asistencia..."; si es listado en
  PDF: "Verifique nuevamente el listado de asistencia..." — y un botón "✅ Asistencia verificada"
  (`btnAsistenciaVerificada`) que hay que pulsar antes de que "Generar informe" se habilite
  (`actualizarBotonGenerarInforme`, `JS.html`); se reinicia cada vez que se entra a la pantalla
  (`prepararCierre`), para que la verificación sea siempre fresca.

## 4.22 Décimo lote: estamento "Otro" en asistencia, doble emoji corregido, firmantes en vivo con `flush()`, guardado local + en la nube, gate de "Finalizar Foro" al informe por correo, y pantalla de despedida con el CSS de FEI 3.1

- **Estamento "Otro" en la asistencia pública**: al elegir "Otro" en `campoEstamento` (asistencia por QR),
  aparece un campo "¿Cuál?" (`bloqueEstamentoOtro`/`campoEstamentoOtro`) que se vuelve obligatorio para
  poder firmar — el valor final enviado al servidor es `"Otro: " + lo que escribió la persona`, así el
  estamento real queda registrado en vez de un simple "Otro" genérico.
- **Doble emoji corregido**: `.mensaje.exito::before` ya añade automáticamente el ✅ vía CSS; tres
  mensajes tenían además el emoji escrito literalmente en el texto ("¡Gracias! Tu valoración fue
  enviada.", el de foto de grupo subida y el de preparación completa), duplicándolo. Se quitó el emoji
  literal de los tres — el `::before` sigue mostrando uno solo.
- **"👥 Firmantes en vivo — 0" atascado en 0**: una revisión exhaustiva de toda la ruta de lectura
  (`rpcEstadoFirmantes`, `contarParticipantesGrupo`, `listarFirmantesGrupo`, lectura por cabeceras,
  `LockService`) no encontró ningún error de lógica; la causa más probable es que Apps Script no
  garantiza que un `appendRow()` sea visible de inmediato a una lectura muy próxima en el tiempo sin un
  `SpreadsheetApp.flush()` explícito — y este proyecto nunca lo llamaba en ningún punto
  (`grep -rn "SpreadsheetApp.flush"` no arrojaba resultados). Se agregó `SpreadsheetApp.flush()`
  inmediatamente después del `appendRow` en `registrarParticipante` (`Data.gs`), que es el punto único
  por el que pasan tanto la firma de asistencia pública como el registro de participantes del grupo, así
  que corrige el conteo tanto para el responsable de grupo como para quienes firman asistencia. **Nota
  de transparencia**: este es el diagnóstico mejor sustentado tras la revisión estática del código, pero
  no se pudo reproducir el error en vivo desde este entorno para confirmar que sea la única causa — se
  recomienda verificarlo tras el despliegue.
- **Guardado local (además del guardado en la nube)**: los campos de texto libre de Sesión 1/2/
  ConectaEduca (`CAMPOS_SESION1`) ahora también se guardan en `localStorage` en cada tecleo (sin
  debounce, es instantáneo), como respaldo del guardado en la nube existente (2s de debounce + 30s
  periódico). Si se cierra la pestaña de golpe antes de que el guardado en la nube alcance a dispararse,
  al reabrir en el **mismo dispositivo** se recupera el borrador local; si se cambia de **dispositivo**,
  no hay borrador local ahí y se retoma desde el último guardado en la nube, porque
  `poblarDesdeBorradorLocalSesion1_()` solo rellena los campos que la nube todavía no haya llenado.
  También se corrigió que `cargarSesion1()` no se llamaba al entrar directo a `pantallaSesion2` sin haber
  pasado por `pantallaSesion1` en ese dispositivo, dejando esos campos compartidos sin cargar. **Nota de
  transparencia**: la lógica se trazó a mano con cuidado, pero no se pudo ejecutar una prueba real de
  cierre súbito de pestaña y reapertura desde este entorno — se recomienda probarlo en vivo tras el
  despliegue.
- **"Finalizar Foro" ahora exige haber enviado el informe por correo**: el botón `btnFinalizarForo`
  queda deshabilitado (con un aviso `avisoFaltaEnvioCorreo` visible) hasta que se haga clic en "Enviar
  informe por correo" al menos una vez (`estado.informeEnviadoPorCorreo`); también se agregó un texto de
  apoyo (`avisoDestinatariosInforme`) que nombra, con los datos reales del grupo ya cargados en
  `estado.instituciones`, todas las instituciones educativas del grupo a cuyos correos llegará el
  informe del Foro Educativo Comunal 2026. Ese texto ahora es fiel a lo que realmente ocurre: en
  `Correo.gs`, `enviarInformeGrupo` agrega los correos de `obtenerInstitucionesDelGrupo` (columna EMAIL
  de `CaracterizacionIE`) como copia del envío — antes el informe solo llegaba al responsable y a los
  asistentes que hubieran dejado su correo, nunca a las IE como institución.
- **Pantalla de despedida con el CSS de FEI 3.1**: `pantallaDespedida` se rediseñó con el mismo lenguaje
  visual que la pantalla de despedida de FEI 3.1 (tarjeta centrada, ícono grande, texto con interlineado
  amplio, cierre en verde) mediante las nuevas clases `.despedida-tarjeta/-icono/-texto/-cierre`
  (`CSS.html`), omitiendo cualquier fecha o año concreto y usando solo un texto general de agradecimiento
  para esta entrega, según lo pedido.
- **Verificación de permisos de envío de correo**: se agregó `testVerificarPermisosCorreo()` (`Tests.gs`)
  para ejecutar manualmente desde el editor de Apps Script — reporta si la cuenta que corre el proyecto
  puede enviar como `calidadeducacion@alcaldianeiva.gov.co` (el remitente que usaba FEI 3.1,
  `REMITENTE_FEM` en su `Código.js`), la cuota diaria restante, y una recomendación. **A propósito no se
  cambió** el valor por defecto de `CORREO_REMITENTE` (sigue vacío, usa la cuenta que ejecuta el script):
  cambiarlo a ciegas podría romper el envío de correo ya funcionando en producción si la cuenta real no
  tiene ese alias de Gmail, y no fue posible verificarlo en vivo desde este entorno. Se documentó en
  `Config.gs` que se debe correr `testVerificarPermisosCorreo()` primero para decidir con certeza.

## 4.23 Undécimo lote: pantalla de invitados especiales (estudiantes/egresados y padres/madres/acudientes) tras la Confirmación de caracterización, con enlace directo por grupo y envío por correo

- **Nueva pantalla `pantallaInvitadosEspeciales`**: se inserta entre "Confirmación de caracterización" y
  "Sesión de preparación" en `ORDEN_PANTALLAS`, pero **solo aparece si el grupo declaró estudiantes,
  egresados o padres/madres/acudientes** en la matriz de "Participación por estamento e institución"
  (`estado.participacionEstamento.totalesPorEstamento` — `ESTUDIANTES`+`EGRESADOS` y `PADRES`); si el
  grupo no declaró ninguno, `btnConfirmarCaracterizacion` sigue saltando directo a la sesión de
  preparación, igual que antes (`irTrasConfirmarCaracterizacion_`, `JS.html`). El texto se arma
  dinámicamente con los conteos reales: "Los/Las N estudiantes y egresados(as) registrados(as) y los N
  padres/madres/acudientes registrados(as) tienen un espacio especial como invitados. Sus aportes deben
  ser enviados desde el perfil de invitados." — cada uno de los dos bloques (estudiantes/egresados y
  padres/madres/acudientes) solo se muestra si su conteo es mayor que cero, tal como pidió el usuario.
- **Enlace directo por grupo y por tipo de invitado**: cada bloque tiene "🔗 Dar enlace" (copia al
  portapapeles) y "✉️ Enviar enlace" (despliega un textarea para escribir el/los correo(s) del
  responsable de envío de ese grupo — estudiante/egresado(a) o padre/madre/acudiente — y un botón "Enviar
  por correo"). El enlace reutiliza el mismo TOKEN de acceso del grupo, con `&invitado=ESTUDIANTE` o
  `&invitado=ACUDIENTE` agregado (`construirUrlInvitado_`/`obtenerEnlacesInvitadosGrupo`, `Access.gs`) —
  al abrirlo, `doGet` (`Code.gs`) resuelve tanto el grupo (`ID_GRUPO_ACCESO`, ya existente) como el tipo de
  invitado (`INVITADO_TIPO_ACCESO`, nuevo) y el cliente salta directo a "Tu institución educativa"
  (`overlayInvitadoIE`), sin pasar por "¿Quién eres?" ni "Tu grupo" — ambos ya vienen resueltos por el
  enlace. El correo se envía desde el mismo remitente configurado (`remitenteValido_`) con instrucciones
  cortas y claras (`enviarEnlaceInvitados`, `Correo.gs`).
- Se generan los enlaces bajo demanda (`rpcObtenerEnlacesInvitadosGrupo`) y se cachean en
  `estado.enlacesInvitados` para no repetir la llamada al servidor entre "Dar enlace" y "Enviar enlace" de
  un mismo bloque.

## 4.24 Duodécimo lote: doble emoji corregido en el mensaje de aportes de preparación enviados, e instrucción del Consolidado de Socialización reescrita

- **Doble emoji corregido**: el mensaje "Aportes de la IE … enviados para socializar oralmente con el
  grupo … y la comunidad asistente." (`pantallaPreparacionIE`, tras enviar la preparación) llevaba un "✅ "
  literal además del que ya agrega `.mensaje.exito::before` por CSS — se quitó el emoji literal, mismo
  patrón de corrección de lotes anteriores.
- **Instrucción del Consolidado de Socialización reescrita** (`pantallaSesion1`, encabezado): de "Registre
  lo que cada institución compartió durante la socialización (un consolidado por grupo, no formulario por
  IE). Cada pregunta debe tener entre 50 y 400 palabras." a "Luego de leer los aportes de cada IE, y de
  los invitados, registre lo más relevante de la socialización (un consolidado por grupo, no por IE).
  Opcionalmente, en caso de existir hallazgos, reflexiones o propuestas propias de la comunidad que no
  encajen en las preguntas anteriores pueden escribirlos al final. Todos los campos son obligatorios." —
  el rango de palabras por pregunta sigue mostrándose igual junto a cada `<textarea>` (contador en vivo,
  `inicializarContadoresPalabras`, `JS.html`), que no se tocó; solo cambió el texto de apoyo general.

## 4.25 Décimo tercer lote: texto de apoyo bajo "Construcción colectiva del grupo" (Sesión 1)

- Justo debajo del título "Construcción colectiva del grupo" (`pantallaSesion1`) se agregó el texto de
  apoyo: "Se sugiere que, mientras se proyectan y van leyendo los aportes, en otro dispositivo se vayan
  registrando lo más relevante de esta sesión." — recomendación práctica para dividir el trabajo entre
  quien proyecta/lee los aportes de preparación y quien va diligenciando las respuestas del consolidado.

## 4.26 Décimo cuarto lote: cantidad declarada del listado en PDF, ver el PDF y cambiar a QR, todo desde Revisión y cierre

- **Nueva columna `CANTIDAD_LISTADO_ASISTENCIA`** (`AccesosGrupo`, `Access.gs`): con el método "Listado en
  PDF" no existe un conteo de firmas en vivo (nadie firma por celular), así que en "Revisión y cierre" el
  grupo ahora **declara cuántas personas quedaron registradas/firmadas en el listado físico**
  (`campoCantidadListadoCierre`, autoguardado con `rpcGuardarCantidadListadoAsistencia`) — ese número
  reemplaza al "0 han firmado" que quedaba fijo en ese método.
- **Mensaje unificado**: la línea de conteos ahora dice, para el método PDF, "N participantes registrados
  — M **han firmado asistencia en formato PDF**" (antes solo "han firmado", igual que QR, dando a entender
  un conteo en vivo que en realidad nunca ocurría) — `textoFirmantesCierre` cambia según el método
  (`cargarVerificacionFirmantesCierre_`, `JS.html`). El aviso adicional debajo también se simplificó,
  quitando la frase redundante con el título de la sección.
- **Ver el PDF del listado** ("👁 Ver PDF del listado"): enlace directo al PDF ya subido
  (`obtenerInfoListadoAsistencia`, `Asistencia.gs`, vía `https://drive.google.com/file/d/<id>/view`) —
  visible solo si ya se subió alguno.
- **Cambiar método a QR desde Revisión y cierre** ("🔄 Cambiar método a QR / enlace"): antes solo se podía
  cambiar de método desde la pantalla de Participación; ahora también se puede hacer aquí mismo, al final,
  sin tener que retroceder — reutiliza `elegirMetodoAsistencia` (mismo guardado que ya existía y que ya
  permitía pasar de Listado a QR en cualquier momento) y refresca los paneles propios de esta pantalla.

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

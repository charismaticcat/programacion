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

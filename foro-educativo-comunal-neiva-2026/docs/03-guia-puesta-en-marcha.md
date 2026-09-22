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

## 4.27 Décimo quinto lote: agregar/cambiar la fotografía del grupo directamente desde Revisión y cierre

- **Fotografía general del grupo, también en Revisión y cierre**: antes, si faltaba subirla, el mensaje de
  error remitía a "Atrás — en Participación" o al botón de Confirmación de caracterización, obligando a
  salir de la pantalla. Ahora "Revisión y cierre" tiene su propia vista previa/botón "📷 Agregar
  fotografía" (pasa a "✏️ Cambiar fotografía" una vez subida, igual que en Confirmación de
  caracterización) con su propio formulario de subida — se puede resolver ahí mismo, sin retroceder.
  `renderPieFotoCaracterizacion` (`Components.html`) ahora acepta un contenedor opcional para reutilizarse
  en el nuevo bloque (`pieFotoCierre`); `refrescarVistasFotoGrupo_` (`JS.html`) actualiza las tres vistas
  (Participación, Confirmación de caracterización y Revisión y cierre) a la vez.
- **Mensaje de error simplificado**: "Falta subir la fotografía general del grupo antes de generar el
  informe. Puede hacerlo desde el botón 'Agregar fotografía' (arriba)." — con scroll automático hasta ese
  botón al hacer clic en "Generar informe" sin foto.
- Sin cambios en el resto de gates de "Generar informe" (valoración completada, asistencia verificada):
  siguen deshabilitando el botón con su propio aviso, como ya funcionaba.

## 4.28 Décimo sexto lote: "Revisar todo antes de enviar", CSS de la cantidad del listado en PDF, correo obligatorio en asistencia, y estamento cruzado con rol

- **"🔍 Revisar todo antes de enviar"** (nuevo botón en "Revisión y cierre"): abre una pantalla completa con
  toda la información ya registrada por el grupo, sección por sección — fotografía general del grupo,
  caracterización, participación por estamento e institución, responsable de envío y asistentes,
  Consolidado de Socialización (Sesión 1) y ConectaEduca (Sesión 2). La fotografía y la matriz de
  participación son **editables directamente ahí** (reutilizan los mismos `render*`/RPC de sus pantallas
  originales, así que un cambio aquí es el mismo dato real). El Consolidado de Socialización y ConectaEduca
  también son editables ahí, **incluso después de haberse enviado de forma definitiva** — el servidor
  (`guardarSesion1`, `Sesion1.gs`) nunca bloqueó esos campos tras el envío definitivo, solo la vista de
  solo lectura del cliente lo hacía, así que no fue necesario tocar nada del backend para permitirlo, según
  decisión explícita del usuario. Caracterización y Responsables se muestran en solo lectura con un botón
  "✏️ Editar en…" que lleva directo a su pantalla original — esos dos formularios usan ids únicos por
  campo que no se pueden repetir sin arriesgar guardar el valor equivocado (p. ej. el rector editable), así
  que se prefirió un salto directo en vez de duplicar esos campos.
- **CSS de "Cantidad de personas registradas en el listado de asistencia"**: el `input[type="number"]` no
  tenía ningún estilo propio en toda la aplicación (se veía como un campo nativo del navegador, fuera de
  lugar) — se agregó a la regla compartida de inputs, y además se le dio su propia caja
  (`#panelListadoCierre`) con el campo compacto, centrado y en un tono más marcado, en vez de ocupar todo
  el ancho de la tarjeta.
- **Correo obligatorio en la asistencia pública** (`AsistenciaPublica.html`): la etiqueta pasó de "Correo
  (opcional)" a "Correo electrónico", y ahora se exige antes de poder firmar — validado tanto en el
  formulario como en el servidor (`registrarAsistenciaPublica`, `Asistencia.gs`).
- **Estamento cruzado con rol en el reporte de asistencia**: `registrarAsistenciaPublica` guardaba
  `ROL_FORO` vacío para toda firma por QR (la página pública no pide un rol aparte del estamento) — ahora
  guarda el mismo estamento declarado como `ROL_FORO`, para que ningún reporte que lea esa columna la vea
  en blanco. `listarFirmantesGrupo` (`Data.gs`) también cruza en el momento de leer (usa `ROL_FORO` si
  existe, si no cae de vuelta al `ESTAMENTO`), como red de seguridad para cualquier registro. Para los
  registros que ya existían con `ROL_FORO` vacío antes de este cambio, se agregó
  `corregirRolForoVacioConEstamento()` (`Tests.gs`) — ejecutar una sola vez, manualmente, desde el editor
  de Apps Script.

## 4.29 Décimo séptimo lote: barra de progreso del informe, y rediseño de la asistencia pública ("No aplica", Sede/Comuna/Jornada/Grado condicionales, y Género + rango de edad para todos)

- **Barra de progreso al generar el informe**: en el overlay de carga (`#overlayCargaAccion`), debajo del
  emoji/imagen, aparece una barra de progreso propia solo cuando `mostrarCargaAccion_("informe", …)` está
  activo. Apps Script no reporta avance real de una sola llamada síncrona (no hay forma de saber "va en el
  40%" mientras se arma el Doc/PDF), así que es **simulada**: avanza rápido al principio y se frena cerca
  del 90% mientras se espera la respuesta del servidor (`iniciarBarraProgresoInforme_`, `JS.html`), y solo
  llega a 100% cuando el informe ya está listo (`ocultarCargaAccion_` → `detenerBarraProgresoInforme_`).
  Para el resto de acciones (que ya tienen la barra indeterminada existente) permanece oculta.
- **"No aplica" en la selección de institución** (`AsistenciaPublica.html`): se agrega como una opción más
  al final del listado de IE (`idIE = "NO_APLICA"`), tratada como "sin institución" en todas partes:
  `institucionSeleccionada_()` la excluye de Sede/Comuna, y `listarFirmantesGrupo` (`Data.gs`) reporta la
  institución como "No aplica" en vez de dejarla en blanco.
- **Jornada / Sede / Comuna / Grado condicionales según estamento**, aplicados en cascada tras elegir
  estamento (y, para Sede/Comuna, tras elegir una IE real distinta de "No aplica"):
  - **Coordinador(a), Docente, Tutor PTA PFI/3.0, Orientador(a), Estudiante, Padre/madre/acudiente,
    Egresado(a)**: habilitan Jornada (Mañana/Tarde/Única) **y** Sede + Comuna.
  - **Rector(a), Personal administrativo, Sector productivo**: habilitan Sede + Comuna, pero **no**
    Jornada (un rector o el sector productivo no pertenece a una jornada del estudiantado).
  - **Otro**: ninguna de las tres.
  - Sede se puebla cruzando con las sedes reales de la IE elegida (columna `sedes` de
    `CaracterizacionIE`, separadas por `;`), con **"Sede central"** agregada siempre como primera opción
    sintética (la columna solo lista sedes adicionales). Comuna se autocompleta (solo lectura) según la
    IE/sede elegida.
  - Si el estamento es **Padre/madre/acudiente**, la etiqueta de Jornada cambia a referirse al
    estudiante ("Jornada del/de la estudiante") y se habilita además **Grado que cursa el/la estudiante**.
    Si el estamento es **Estudiante**, se habilita esa misma pregunta de Grado.
- **Género y rango de edad para todos los estamentos**: dos preguntas nuevas, siempre visibles — Género
  (Femenino/Masculino/Otro/Prefiero no decirlo) y Rango de edad en tramos de 10 años (Menor de 10, 10 a 19,
  20 a 29, … 70 o más).
- **Correo obligatorio** (lote anterior) y estos campos nuevos se guardan en `ParticipacionComunal` en
  columnas nuevas (`JORNADA`, `SEDE`, `COMUNA`, `GRADO`, `GENERO`, `RANGO_EDAD`), agregadas al final de
  `cabecerasParticipacionComunal_()` — se automigran solas en la hoja existente (`obtenerHoja_`). Ninguno
  de estos campos nuevos tiene validación obligatoria en el cliente ni el servidor: son preguntas simples
  con valores por defecto razonables, tratadas igual que "Estamento" (que tampoco se valida como
  obligatorio), para mantener el mismo criterio que ya existía en este formulario.
- **Nota de transparencia**: esta pantalla se verificó de forma estática (`node --check` sobre el JS
  extraído, e IDs/etiquetas balanceadas) pero no se probó en vivo desde un navegador en este entorno —
  conviene una pasada manual antes de usarla con un grupo real.

## 4.30 Décimo octavo lote: registro de invitados en plural, estimado de estudiantes/egresados(as) por institución, rango de edad, y perfil de egresado(a) separado

- **Todo el flujo de invitados en plural ("ustedes")**: las pantallas "¿Quiénes son?", "Su grupo",
  "Instituciones que representan", "Su institución educativa", "Antes de empezar", "Preparación" y la
  pantalla final, además de las 18 ayudas de pregunta (P1-P6 × 3 perfiles) y el correo que envía el enlace
  a invitados (`enviarEnlaceInvitados`, `Correo.gs`), se redactaron en plural — un mismo dispositivo suele
  pasar de mano en mano entre varios estudiantes/egresados(as) o adultos responsables que entran uno tras
  otro. Los roles/vínculos con la IE también se pluralizaron ("Padres de estudiantes", "Acudientes",
  "Representantes de los comités institucionales").
- **Perfil "egresado(a)" separado de "estudiante actual"** desde el primer paso: la pantalla "¿Quiénes
  son?" ahora tiene 3 botones — Estudiantes actuales, Egresados(as) y Adultos responsables — en vez de
  agrupar estudiante/egresado(a) en un solo botón con una pregunta de seguimiento ("¿Eres estudiante actual
  o egresado(a)?", eliminada por pedido del usuario). Por eso, en "Antes de empezar", los estudiantes
  actuales siguen viendo "Años que llevan estudiando en la institución", mientras que los egresados(as) ven
  en su lugar **Año de graduación** y **Profesión u ocupación actual** (con opciones lógicas: estudiando
  educación superior, ejerciendo su profesión, trabajando en algo distinto, independiente, buscando
  empleo, labores del hogar, servicio militar, otra situación). Las 6 preguntas de preparación para
  egresados(as) tienen la misma ayuda que las de estudiante actual, más una nota aclarando que, si ya no
  estudian en la institución pero recuerdan algo relacionado, también pueden responder así: "Recuerdo
  que…". *Nota: el enlace directo de la pantalla de "invitados especiales" (que ya venía con el tipo
  resuelto de antes) todavía no distingue egresado(a) de estudiante actual — por defecto entra como
  estudiante actual; quien llegue por ese enlace puede simplemente usar "Somos invitados" desde el inicio
  si quiere marcarse como egresado(a).*
- **Rango de edad (de 10 en 10) en vez de edad puntual**: "Antes de empezar" ahora pregunta un rango
  (Menor de 10, 10 a 19, … 70 o más) en vez de un número — se guarda en la misma columna EDAD de
  `AportesInvitadosPreparacion` (ahora como texto), sin necesidad de migrar la hoja.
- **"Instituciones que representan"** (pantalla nueva, entre "Su grupo" y "Su institución educativa"):
  antes de que cada quien se registre individualmente, se marca (checkbox) cada institución del grupo de
  la que traen estudiantes y/o egresados(as), con un campo de cantidad para cada una — un estimado
  agregado, puramente informativo (nunca bloquea el flujo, incluso si no se marca ninguna), guardado en una
  hoja nueva `DeclaracionInvitadosIE` (`guardarInstitucionesRepresentadasInvitado`/
  `obtenerDeclaracionInvitadosGrupo_`, `Invitados.gs`) y mostrado en el informe consolidado final si hay
  algún valor declarado. El registro individual (nombre, rango de edad, respuestas P1-P6) sigue exactamente
  igual después de este paso.
- **Todas las preguntas de preparación son obligatorias para invitados**: `marcarPreparacionEnviadaInvitado`
  ahora exige las 6 respondidas (antes bastaba con una sola) — mensaje "Respondan todas las preguntas antes
  de enviar."
- **Nota de transparencia**: verificado de forma estática (`node --check`, IDs sin duplicar, etiquetas
  balanceadas) pero no probado en vivo desde un navegador en este entorno.

## 4.31 Décimo noveno lote: "Verificar listado" en modo lectura (método QR) y fix de la pantalla de invitados especiales que no aparecía

- **"🔍 Verificar listado" en Revisión y cierre (método QR)**: nuevo botón junto al QR/enlace que muestra, en
  modo **solo lectura**, quiénes ya firmaron la asistencia — mismo dato (`rpcEstadoFirmantes`) y mismo
  formato de fila que la pantalla pública de asistencia, con sondeo cada 4 segundos mientras el listado
  está abierto. Nunca permite editar ni eliminar una firma desde aquí, solo consultar antes de generar el
  informe.
- **Fix: la opción de enviar el enlace de invitados a veces no aparecía antes de la Sesión de
  preparación.** Causa raíz: al hacer click en "Confirmar caracterización", la app decide si mostrar la
  pantalla "Invitados especiales" (y por lo tanto la opción de enviar su enlace) leyendo
  `estado.participacionEstamento` en memoria — pero el autoguardado de la matriz de participación por
  estamento (que se edita en esa misma pantalla) pone ese valor en `null` justo al guardar y lo vuelve a
  poblar en segundo plano sin avisar; si el grupo terminaba de escribir la cantidad de
  estudiantes/egresados/padres y le daba click a "Confirmar" casi de inmediato, la decisión se tomaba con
  ese `null` (o con un valor desactualizado) y la pantalla de invitados especiales se saltaba por completo,
  aunque sí se hubiera declarado gente. Ahora, al hacer click en "Confirmar caracterización": (1) se
  vuelcan de inmediato los guardados de la matriz que todavía estuvieran esperando su debounce de 1.5s
  (`flushGuardadosPendientesParticipacionEstamento_`), (2) se refresca `estado.participacionEstamento`
  desde el servidor, y solo entonces se decide si hay que mostrar "Invitados especiales".
- **Nota de transparencia**: verificado de forma estática (`node --check`, IDs sin duplicar, etiquetas
  balanceadas) pero no probado en vivo desde un navegador en este entorno.

## 4.32 Vigésimo lote: mensajes de error junto al campo, fix de "Revisar todo" vacío, QR en pantalla completa en Cierre, correo opcional para acudientes, aviso de tiempo del informe, barra de progreso navegable y botón de inicio

- **Fix: el mensaje de error de Sesión 2 aparecía muy arriba.** "Estos campos deben tener entre 50 y 400
  palabras: PRIORIDADES_CE, ACUERDOS_CE, PROPUESTAS_CE, RUTA_CE" (y "Complete todos los campos…") se
  mostraban en `mensajeConectaEduca`, el mismo mensaje de "Agregar actor" que está arriba del todo en la
  pantalla — muy lejos de esos campos y del botón "Enviar Sesión 2 definitiva". Se agregó un mensaje propio
  (`mensajeEnviarSesion2`) justo junto a ese botón; "Agregar actor" sigue usando el suyo, sin tocarlo.
- **Fix: en "Revisar todo antes de enviar", los textareas de Consolidado de Socialización (Sesión 1) y
  ConectaEduca (Sesión 2) aparecían vacíos** aunque las respuestas sí existieran (y si el grupo iba a la
  pantalla original de Sesión 1/2, ahí sí se veían). Causa: el autoguardado de esas pantallas tiene un
  debounce de hasta 2 segundos antes de llegar al servidor; si el grupo escribía algo y entraba a "Revisar
  todo" casi de inmediato, la foto que trae `rpcObtenerSesion1` en ese momento podía no incluir todavía lo
  último escrito, y "Revisar todo" mostraba esa foto vacía en vez de lo que había en pantalla. Ahora, para
  cada campo, "Revisar todo" prioriza el valor que ya esté en su campo original (`campo_<CLAVE>`) sobre el
  del servidor, y al editar dentro de "Revisar todo" también refleja el cambio en ese campo original —
  ambas vistas del mismo dato quedan siempre sincronizadas entre sí.
- **"👁 Ver QR en pantalla completa" en Revisión y cierre (método QR)**: mismo visor de pantalla completa
  que ya existía en Participación (para proyectar en un monitor durante el evento), ahora también
  disponible junto al QR de Revisión y cierre — se extrajo a una función compartida
  (`abrirPantallaCompletaQR_`).
- **Correo opcional para padres/madres/acudientes**: en `AsistenciaPublica.html`, el correo sigue siendo
  obligatorio para todos los estamentos excepto Padre/madre/acudiente, cuya etiqueta cambia a "Correo
  electrónico (opcional)" — validado tanto en el cliente como en `registrarAsistenciaPublica` (`Asistencia.gs`).
  En el flujo de invitados (perfil "adultos responsables"), se agregó un campo de correo, también opcional,
  nuevo (columna `CORREO` en `AportesInvitadosPreparacion`), que aparece en el informe consolidado junto al
  resto de datos del invitado.
- **Aviso de tiempo en la generación del informe**: el overlay de carga ahora dice "Generando el informe…
  Esto puede tomar hasta dos minutos, no cierre esta ventana."
- **Barra de progreso navegable, con el nombre de cada sesión**: la barra superior (antes solo un relleno
  de porcentaje) ahora es una fila de botones — Participación, Caracterización, Preparación, Sesión 1,
  Sesión 2, Cierre — que llevan directo a esa pantalla. Nunca permite saltar hacia adelante más allá de la
  pantalla más lejana ya alcanzada en el recorrido (no se salta ningún candado, p. ej. no se puede entrar a
  Sesión 2 sin haber enviado Sesión 1 de forma definitiva); ir hacia atrás siempre está permitido, igual que
  ya permitían los botones "✏️ Editar en…" de "Revisar todo".
- **Botón 🏠 "Volver al inicio"**: en el encabezado, siempre visible una vez el grupo validó su código de
  acceso — lleva directo a la pantalla de bienvenida ("Querida Comunidad Educativa del grupo…").
- **Nota de transparencia**: verificado de forma estática (`node --check`, IDs sin duplicar, etiquetas
  balanceadas) pero no probado en vivo desde un navegador en este entorno.

## 4.33 Vigésimo primer lote: la barra de progreso navegable se habilita apenas se valida el código

- **Fix**: la barra de progreso navegable (lote 20) solo se habilitaba a medida que el grupo iba avanzando
  pantalla por pantalla — recién validado el código de acceso, todos sus pasos (Participación,
  Caracterización, Preparación, Sesión 1, Sesión 2, Cierre) aparecían deshabilitados hasta pasar
  manualmente por Inicio/Presentación/Metodología/Consentimiento (spec del usuario: "Navegación debe estar
  disponible luego de ingresar el código"). Esas pantallas de introducción no son un candado real, así que
  ahora, apenas `rpcValidarAcceso` confirma el código, se habilita de una vez hasta "Participación" (o más
  allá, si el grupo ya venía más adelante en un dispositivo distinto).

## 4.34 Vigésimo segundo lote: logos SEM/FEM en el encabezado, crédito del desarrollador en el pie, y utilidad para limpiar todos los datos ingresados

- **Logos institucionales en las esquinas superiores** del encabezado principal (spec del usuario): SEM/Alcaldía
  a la izquierda, Foro/FEM a la derecha — reutilizan los mismos IDs de Drive que ya usaba el splash inicial y
  el encabezado del informe (`LOGO_PIE_ID`/`LOGO_ENCABEZADO_ID`, `ConfiguracionComunal`), poblados en JS.html;
  si algún ID no está configurado, ese logo simplemente no aparece. El botón 🏠 "Inicio" (lote 20) se movió del
  encabezado a un botón normal justo encima de la barra de progreso navegable, para no competir por el mismo
  espacio con los logos.
- **Crédito del desarrollador en el pie de página**: debajo de "Foro Educativo Comunal Neiva 2026 —
  Secretaría de Educación de Neiva" ahora dice "Desarrollado por Jhon Sanchez". *Pendiente*: el logo que el
  usuario compartió (HelpRofe) no llega a este entorno como un archivo accesible cuando se pega directo en el
  chat — se dejó el `<img>` reservado (`#logoDesarrolladorPie`, oculto) listo para completar su `src` en
  cuanto se suba como archivo adjunto.
- **`limpiarTodosLosDatosIngresadosPorUsuarios()`** (Tests.gs, IRREVERSIBLE, ejecutar manualmente desde el
  editor de Apps Script): vacía por completo todo lo diligenciado por cualquier grupo (firmantes, Sesión 1/
  ConectaEduca y sus actores, responsables, valoraciones, preparación por IE, invitados individuales y su
  estimado agregado, envíos diferidos pendientes, e informes generados) — pero conserva intactos el catálogo
  de grupos/IE y las filas de `AccesosGrupo` (los códigos/enlaces ya generados siguen funcionando), solo
  reiniciando sus campos de progreso (estado, envíos, método/evidencias de asistencia, foto, consentimiento,
  última pantalla) a su valor inicial. Alcance confirmado explícitamente con el usuario antes de escribirla.
- **Intento de habilitar `clasp run`**: se agregó `executionApi: {"access": "MYSELF"}` a `appsscript.json`
  (cambio de manifiesto inocuo, no afecta el acceso público del webapp) y se reintentó `clasp run
  testCrearGrupoDePrueba` — el error cambió de un `NOT_FOUND` genérico a "Unable to run script function.
  Please make sure you have permission to run the script function.", que es exactamente el mensaje que
  Google devuelve cuando la cuenta todavía no tiene activado el interruptor "Google Apps Script API" en
  https://script.google.com/home/usersettings. Ese interruptor solo se puede activar entrando con el
  navegador a esa página (no hay forma de hacerlo por línea de comandos) — una vez activado, se puede
  reintentar `clasp run testCrearGrupoDePrueba` y `clasp run testGenerarAccesos` para crear GRUPO-PRUEBA
  automáticamente; mientras tanto, ejecutar esas dos funciones manualmente desde el editor de Apps Script
  (Extensiones → Apps Script → seleccionar función → Ejecutar) sigue siendo la vía garantizada.
- **Nota de transparencia**: verificado de forma estática (`node --check`, IDs sin duplicar, etiquetas
  balanceadas) pero no probado en vivo desde un navegador en este entorno.

## 4.35 Vigésimo tercer lote: logo del desarrollador en el pie de página

- **Logo HelpRofe en el pie de página**: el usuario compartió el archivo real (Google Drive,
  `1BXkKDuSH_XhlLbdPtyYlXpJypbFH9f38`, `image/png`, de su propia cuenta) — se agregó como
  `LOGO_DESARROLLADOR_ID` en `ConfiguracionComunal` (`Config.gs`), se asegura su visibilidad pública una sola
  vez por instalación (`asegurarLogoDesarrolladorPublico_`, `Drive.gs` — bandera propia
  `LOGO_DESARROLLADOR_PUBLICO`, independiente de `LOGOS_SPLASH_PUBLICOS` porque esa ya pudo estar en "SI" en
  instalaciones existentes) y se muestra debajo de "Desarrollado por Jhon Sanchez", con el mismo patrón que
  los logos SEM/FEM del encabezado (`urlImagenDrive`, JS.html).

## 4.36 Vigésimo cuarto lote: logos del encabezado más grandes, y fix de la transición inicial (el logo de la Alcaldía/SEM no alcanzaba a cargar)

- **Logos SEM/FEM más grandes** en las esquinas del encabezado (de 52px a 80px de alto máximo; en móvil, de
  38px a 54px), con el padding del encabezado ajustado para que sigan sin encimarse con el título.
- **Fix: en la transición inicial, solo se veía el logo del Foro/FEM** — el logo de la Alcaldía/SEM (el
  primero de la secuencia, sin caché del navegador todavía) tardaba en descargar más de lo que duraba su
  temporizador, que arrancaba apenas se PEDÍA la imagen (no cuando terminaba de cargar); para cuando el
  navegador terminaba de pintarlo, el temporizador ya había avanzado al siguiente logo. Ahora el conteo para
  pasar al siguiente logo empieza únicamente cuando la imagen actual ya cargó (evento `onload`, con
  `onerror` y un respaldo de 5s por si ninguno de los dos llegara a disparar) — nunca antes.
- **Transición más larga** (spec del usuario): cada logo permanece 3.2s en pantalla (antes 2.2s, y ese conteo
  arrancaba antes de que la imagen siquiera hubiera empezado a cargar), y el cierre final del splash pasó de
  0.9s a 1.2s.
- **Nota de transparencia**: verificado de forma estática (`node --check`); no probado en vivo desde un
  navegador en este entorno — conviene confirmar con conexión lenta que ambos logos ya se alcanzan a ver.

## 4.37 Vigésimo quinto lote: invitados — no repetir la selección de IE, sin pedir datos personales, y estimado por hombres/mujeres

- **"Luego de seleccionar las IE no debe devolverse a seleccionar las IE"**: marcar institución(es) en
  "Instituciones que representan" ahora ES la elección de institución del invitado — se usa la primera
  marcada para entrar directo a su preparación, sin volver a preguntar en una pantalla aparte ("Su
  institución educativa" ya no aparece en el flujo normal, solo sigue existiendo para el enlace directo de
  "invitados especiales", que nunca pasó por esa pantalla de marcar instituciones). Se exige marcar al menos
  una institución para poder continuar.
- **"Omite la parte de pedir datos personales"**: se eliminó del flujo la pantalla "Antes de empezar"
  (nombre, rango de edad, sexo, años estudiando/año de graduación/profesión, vínculo con la IE, correo) —
  al marcar la institución se entra directo a la Preparación (P1-P6). El estudiante actual/egresado(a) sigue
  distinguiéndose correctamente (ayuda de cada pregunta, "Recuerdo que…" para egresados(as)) porque ese dato
  ya se eligió en el primer paso ("¿Quiénes son?") y ahora viaja directo del cliente en cada llamada
  (`rpcObtenerPreparacionIEInvitado`/`rpcGuardarPreparacionIEInvitado`), en vez de leerse de una
  caracterización que ya no se pide. El informe consolidado de invitados sigue funcionando (los conteos de
  estudiantes/egresados(as)/adultos responsables se mantienen), solo que ahora todos aparecen sin nombre
  registrado, ya que no se vuelve a pedir.
- **"Solo deja cantidad de hombres y mujeres"**: en "Instituciones que representan", los campos por
  institución cambiaron de "Estudiantes"/"Egresados(as)" a "Hombres"/"Mujeres" — columnas
  `CANTIDAD_HOMBRES`/`CANTIDAD_MUJERES` en `DeclaracionInvitadosIE` (antes `CANTIDAD_ESTUDIANTES`/
  `CANTIDAD_EGRESADOS`; renombradas directamente, sin mantener las viejas, porque la función se agregó hace
  muy poco y no hay datos reales todavía) y en la tabla del informe.
- **Nota de transparencia**: las pantallas "Su institución educativa" (caso enlace directo) y "Antes de
  empezar" siguen en el HTML/CSS por si se necesitan más adelante, pero ya no se usan desde el flujo normal
  — código muerto sin riesgo, no se tocó más para minimizar el alcance del cambio. Verificado de forma
  estática (`node --check`, IDs sin duplicar); no probado en vivo desde un navegador en este entorno.

## 4.38 Vigésimo sexto lote: foto junto al método de asistencia, cantidad del listado en Participación, gris hasta valorar, y cambio de método siempre posible

- **"Fotografía general del grupo debe aparecer en la misma casilla de subir PDF y QR"**: el bloque completo
  se movió de la tarjeta "Participación y asistencia" a la tarjeta "Método de asistencia del grupo", justo
  después de los paneles de QR y de Listado en PDF (visible siempre, sin importar el método elegido).
- **"Cuando se sube PDF debe decir PDF subido correctamente"**: el mensaje de éxito de "Subir listado"
  cambió de "Listado subido correctamente." a "PDF subido correctamente."
- **"Enseguida desplegarse casilla de cantidad de firmantes... y debe registrarse enseguida de 👥 Total de
  participantes declarados: / Total de participantes en Listado de asistencia:"**: al subir el PDF con
  éxito se despliega de inmediato, dentro del mismo panel de "Listado en PDF", el campo "Cantidad de
  personas registradas en el listado de asistencia" (mismo mecanismo que ya existía en Revisión y cierre:
  `rpcGuardarCantidadListadoAsistencia`/`rpcObtenerInfoListadoAsistencia`, autoguardado al cambiar el
  valor). El valor declarado aparece en una segunda línea, justo debajo de "👥 Total de participantes
  declarados:", con el texto "Total de participantes en Listado de asistencia: N" —
  `renderResumenParticipacionEstamento` (Components.html) ya es una función compartida por Participación,
  Confirmación de caracterización y "Revisar todo", así que agregar esa segunda línea ahí la deja visible
  automáticamente **en las tres pantallas**, incluida "la confirmación de la caracterización" tal como se
  pidió, sin duplicar HTML/lógica. El campo de Participación y el de Revisión y cierre quedan sincronizados
  entre sí (cada guardado exitoso actualiza también el valor mostrado en el otro).
- **Sticker de carga para la subida del listado**: el gif "Cargando Sticker by dipielbella" que el usuario
  compartió de nuevo (esta vez por Drive, `1MqumC6-CeumTVkMtlkHILfgTz7sHQarb` — confirmado por metadatos
  como el mismo gif ya usado para "Guardar responsable de envío") ahora también se muestra como indicador de
  carga al subir el PDF del listado (`GIFS_CARGA_ACCION_.listado`, misma URL de Giphy ya cableada para
  "responsable").
- **"Revisión y cierre y todo lo posterior deben aparecer en gris hasta que no se haya enviado la
  valoración"**: la tarjeta "Revisión y cierre" (verificar firmantes, fotografía, "Generar informe") se ve
  en gris y no reacciona a clics (`opacity`, `grayscale`, `pointer-events:none`) mientras la valoración del
  Foro (misma pantalla, tarjeta de arriba) no se haya enviado, con un aviso "🔒 Complete primero la
  valoración..." — se recalcula cada vez que se entra a la pantalla y cada vez que cambia el estado de la
  valoración (enviarla la desbloquea al instante, sin recargar). El botón "Generar informe" ya exigía la
  valoración por separado desde un lote anterior; esto añade el bloqueo visual de toda la sección.
- **"Cambiar de QR a PDF y PDF a QR debe ser posible"**: se quitó la restricción que impedía volver de QR a
  Listado en PDF una vez generado el informe del grupo (tanto la validación del servidor en
  `guardarMetodoAsistencia`, Asistencia.gs, como el bloqueo del botón en el cliente,
  `actualizarBloqueoMetodoAsistencia_`, JS.html) — cambiar de método es ahora siempre posible en cualquiera
  de los dos sentidos, sin borrar lo ya capturado con el método anterior (la decisión de no borrar nada al
  cambiar de método ya venía de un lote anterior y se mantiene).
- **"Envíe el informe al correo de las Instituciones Educativas y al responsable..."**: se corrigió el texto
  del paso 2 de "Para terminar, complete estos pasos" (antes decía "al correo del responsable" únicamente) y
  los avisos junto al botón "Enviar informe por correo" — el envío en sí (`enviarInformeGrupo`, Correo.gs)
  ya enviaba al responsable de envío Y en copia a los correos institucionales de todas las IE del grupo
  desde un lote anterior; solo el texto en pantalla no lo reflejaba con precisión. Se verificó leyendo
  `enviarInformeGrupo` de punta a punta: destinatario principal = responsable de envío (o el correo de
  acceso inicial si no hay responsable registrado), copia = asistentes de envío + correos institucionales de
  cada IE del grupo + `COPIAS_CORREO` de la configuración.
- **Nota de transparencia**: verificado de forma estática (`node --check` en los bloques `<script>`
  extraídos de JS.html/Index.html/Components.html, con el falso positivo ya conocido y documentado en lotes
  anteriores al extraer el bloque de Index.html; conteo de IDs duplicados y balance de
  div/section/header/footer en Index.html). No probado en vivo desde un navegador en este entorno.

## 4.39 Vigésimo séptimo lote: la barra de progreso navegable se habilita también con el progreso real en la nube

- **"Si ya hay avances locales y en la nube puedes retomar en cualquier punto ya realizado y habilitar la
  barra"**: hasta este lote, al validar el código de acceso la barra de progreso navegable solo se
  desbloqueaba (a) hasta "Participación" de entrada, y (b) hasta la pantalla exacta guardada en
  `ULTIMA_PANTALLA` (si existía) al saltar ahí con `destinoResumenPantalla_`. Ahora, además, la respuesta de
  `rpcEstadoGrupo` (que ya se pedía en ese mismo momento) también desbloquea la barra según el progreso real
  del grupo en la nube — si `sesion1Enviada` es verdadero se desbloquea al menos hasta "Sesión 2", y si
  `sesion2Enviada` o `informeGenerado` es verdadero se desbloquea hasta "Cierre" — de forma que el avance
  real del grupo (no solo la última pantalla puntual) siempre deja la barra navegable hasta ahí, sin importar
  desde qué dispositivo se ingrese. Se agregó `pantallaActual_` (qué pantalla está visible en este momento)
  para poder refrescar correctamente qué botón de la barra queda "activo" cuando esta respuesta llega de
  forma asíncrona.
- **Nota de transparencia**: verificado de forma estática (`node --check` sobre el bloque `<script>` de
  JS.html). No se identificó una causa concreta de que la retoma NO ocurriera (la cadena
  `cambiarPantalla`→`rpcGuardarUltimaPantalla`→`ULTIMA_PANTALLA` en `AccesosGrupo`, y su lectura en
  `validarAccesoGrupo`→`destinoResumenPantalla_`→`cambiarPantalla`, se revisó de punta a punta sin encontrar
  un defecto); este lote agrega una vía adicional e independiente para desbloquear la barra (progreso real
  en la nube: Sesión 1/2 enviada, informe generado), que no depende de que `ULTIMA_PANTALLA` esté siempre
  perfectamente al día. No probado en vivo desde un navegador en este entorno — si el problema persiste tras
  este cambio, se necesita el detalle exacto de la reproducción (pantalla exacta antes de salir, pantalla a
  la que llega al reingresar, mismo dispositivo o no) para seguir investigando.

## 4.40 Vigésimo octavo lote: mismo gif "Cargando" en todas las cargas y cambios de pantalla del flujo de invitados

- **"Usa este gif para todas las cargas y cambios de pantalla en los perfiles de estudiantes, egresados y
  acudientes"**: el mismo sticker "Cargando" (Giphy `SYIu9YMtvUc6tBzDH6`, el mismo gif que el usuario ya
  había compartido antes por Drive como `1MqumC6-CeumTVkMtlkHILfgTz7sHQarb`, confirmado como el mismo
  archivo — ya usado en "Guardar responsable de envío" y, desde el lote 26, en "Subir listado en PDF") ahora
  se muestra en **todos** los pasos del flujo de invitados (nuevo tipo `GIFS_CARGA_ACCION_.invitado`),
  reemplazando las transiciones instantáneas o sin indicador que había antes:
  - Elegir tipo (Estudiante/Egresado(a)/Acudiente) → cargar la lista de grupos.
  - Elegir grupo → abrir "Instituciones que representan" (sin llamada al servidor; se muestra el gif un
    instante fijo de todas formas, para que la transición nunca se sienta abrupta).
  - Marcar institución(es) → entrar a preparación (`rpcIniciarAccesoInvitado`).
  - Enlace directo de invitado especial (`?t=TOKEN&invitado=...`) → cargar las instituciones del grupo
    (`rpcInstitucionesDelGrupo`).
  - Entrar a la pantalla de Preparación → cargar las preguntas/respuestas (`rpcObtenerPreparacionIEInvitado`)
    — de paso corrige un pequeño parpadeo que había antes (la pantalla se mostraba vacía un instante, antes
    de que llegaran las preguntas; ahora no se muestra hasta tenerlas listas).
  - Enviar los aportes definitivos (antes usaba el gif genérico de "aportes", compartido con el flujo
    principal de moderadores — ahora tiene su propio tipo `invitado`, así un cambio futuro al gif de
    "aportes" del flujo principal no afecta sin querer al de invitados, y viceversa).
  - Se agregaron manejadores de error (antes ausentes en varias de estas llamadas) que ocultan el gif y
    dejan la pantalla siguiente visible igual, en vez de quedarse con el gif trabado si el servidor falla.
- **Nota de transparencia**: verificado de forma estática (`node --check`). No probado en vivo desde un
  navegador en este entorno.

## 4.41 Vigésimo noveno lote: foto oculta en caracterización si se aplazó, método obligatorio reforzado, y recordatorio de colaboradores

- **"Si se pide subir fotografía más adelante, no debe aparecer opción de cambiar fotografía ni de subir
  fotografía en pantalla de caracterización"**: la fila con el botón "Cambiar/Agregar fotografía" (y su
  formulario) en Confirmación de caracterización ahora se oculta por completo mientras no haya una foto ya
  subida (`estado.fotoGrupoId` vacío) — como el gate obligatorio de Participación ya exige subir la foto o
  aplazarla explícitamente antes de continuar, llegar a Caracterización sin foto siempre significa que se
  aplazó, así que no tiene sentido ofrecer la opción ahí: se vuelve a pedir en Revisión y cierre (donde de
  verdad se exige, antes de generar el informe), tal como ya decía el propio mensaje de "más tarde". Se usa
  el dato del servidor (`fotoGrupoId`), no el flag local del dispositivo, para que se comporte igual sin
  importar desde qué dispositivo se entre.
- **"Seleccionar un tipo de asistencia es obligatorio para continuar a pantalla de caracterización"**: ya lo
  exigía el botón "Continuar" de Participación, pero solo ahí — se agregó la misma validación directamente
  dentro de `cambiarPantalla` (la función central de navegación), así que ahora también queda bloqueado
  intentar llegar a Caracterización (o a cualquier pantalla posterior) sin método elegido por otras vías: la
  barra de progreso navegable, los enlaces "Editar en…" de "Revisar todo", o cualquier otra llamada directa
  — en cualquiera de esos casos, redirige de vuelta a Participación con el mismo aviso.
- **"Al dar continuar, recordar que es posible tener hasta 4 colaboradores"**: al pulsar "Continuar" en
  Participación (una vez superados los dos gates anteriores — método y fotografía), aparece un aviso
  recordando que se pueden registrar hasta 4 colaboradores (responsable de envío + hasta 3 asistentes) para
  recibir el informe, con un botón para continuar. Se muestra una sola vez por sesión, no en cada clic.
- **Nota de transparencia**: verificado de forma estática (`node --check`, IDs sin duplicar, balance de
  div/section/header/footer en Index.html). No probado en vivo desde un navegador en este entorno.

## 4.42 Trigésimo lote: conteo de estudiantes/egresados en aportes de invitados, texto inicial por IE representada, y dos preguntas nuevas en ConectaEduca

- **"Cambiar Un(a) estudiante o egresado(a) por (cantidad) estudiantes y (cantidad) egresados/as opinaron
  sobre:"**: en "¿Qué institución o participante va a preparar sus aportes?" → "Aportes ya enviados por
  estudiantes, egresados y adultos responsables invitados", los aportes de estudiantes/egresados(as) de cada
  IE ya no repiten la etiqueta genérica "🧑‍🎓 Un(a) estudiante o egresado(a)" antes de cada aporte
  individual — se agrupan bajo un solo encabezado con el conteo real de cada uno: "🧑‍🎓 N estudiantes y M
  egresados/as opinaron sobre:", seguido de todas sus respuestas. Los aportes de adultos responsables no
  cambiaron. Se agregó `ROL_ESTUDIANTE` a cada aporte que devuelve `obtenerAportesInvitadosGrupo` (antes solo
  se enviaba `tipoInvitado`) para poder distinguir estudiante actual de egresado(a).
- **"En el formato de textarea, se debe iniciar con el texto que diga: En la Institución Educativa (nombre);
  y mencionar cada IE que haya sido seleccionada"**: cuando un(a) estudiante/egresado(a) invitado(a) abre
  Preparación y todavía no tiene ningún borrador guardado, las 6 respuestas empiezan con "En la Institución
  Educativa \<nombre>; " (o "En las Instituciones Educativas \<nombre1>, \<nombre2> y \<nombre3>; " si marcó
  más de una en "Instituciones que representan" — las respuestas solo quedan asociadas a la primera marcada,
  pero el texto deja constancia de todas las que representa). Nunca se sobrescribe un borrador ya existente.
  El indicador "✓ ya respondida" sigue mirando la respuesta real guardada, no este texto inicial, así que no
  se marca como respondida solo por tener el texto de apertura.
  - Nota de transparencia: como este texto queda dentro del valor real del textarea, es técnicamente posible
    enviar una respuesta que sea solo ese texto inicial sin escribir nada más — no se agregó una validación
    de servidor que descuente ese texto para exigir contenido real más allá de él (la app ya confiaba en el
    criterio del respondiente para escribir algo genuino en las demás preguntas, así que se mantiene el
    mismo criterio aquí).
- **"Aumentar una pregunta en estudiantes... apreciación objetiva sobre los programas de articulación del
  SENA" + "Aumentar pregunta sobre intensificaciones... dejar estas dos nuevas preguntas para sección de
  ConectaEduca"**: se agregaron dos preguntas nuevas de grupo en Sesión 2 — ConectaEduca, tarjeta "Antes de
  registrar actores" (mismo lugar y mismo tratamiento — sin rango de palabras, no obligatorias para el envío
  definitivo — que las dos preguntas de grupo que ya existían ahí: "Necesidades de articulación" y
  "Oportunidades identificadas"):
  - "Apreciación sobre los programas de articulación del SENA" (columna `APRECIACION_SENA_GRUPO`).
  - "Intensificaciones dentro de la institución educativa" (columna `INTENSIFICACIONES_GRUPO`).
  Ambas quedan incluidas también en la vista de solo lectura (tras el envío definitivo), en "Revisar todo
  antes de enviar", y en el informe final del grupo (Informes.gs), igual que las dos preguntas existentes.
- **Nota de transparencia**: verificado de forma estática (`node --check`, IDs sin duplicar, balance de
  div/section/header/footer/label/textarea en Index.html). No probado en vivo desde un navegador en este
  entorno.

## 4.43 Trigésimo primer lote: aportes agrupados por rol, barra de progreso para invitados, y texto inicial general/específico con nombres capitalizados

- **"Las respuestas no deben aparecer como si fueran de una sola IE sino: Aportes de (rol)"**: en "Aportes ya
  enviados por estudiantes, egresados y adultos responsables invitados" (pantalla de Selección de IE), el
  encabezado de cada grupo pasó de "N estudiantes y M egresados/as opinaron sobre:" a "Aportes de estudiantes
  (N) y egresados/as (M):", y los adultos responsables (antes con una etiqueta repetida por cada aporte)
  ahora también se agrupan bajo un único "Aportes de acudientes (N):" — en ambos casos la etiqueta encabeza
  por ROL, no da la impresión de pertenecer en exclusiva a la IE bajo la que está ese `<details>` (el propio
  texto de cada respuesta ya menciona, desde el lote anterior, todas las instituciones que esa persona
  representa).
- **"Debe cargar más rápido o mostrar una barra de progreso mientras se muestra el gif"**: el flujo de
  invitados encadena varias llamadas al servidor seguidas (instituciones → ingresar a la IE → cargar
  preparación); ahora, mientras se muestra el gif "Cargando" de ese flujo, también aparece la barra de
  progreso simulada que ya se usaba para "Generando el informe…" (avanza hasta 90% y se completa al 100% en
  cuanto el servidor de verdad responde), para que la espera se sienta con avance real en vez de un gif
  estático fijo.
- **"En textarea dar botón, responder de forma general o específicando una IE"**: cuando el invitado marcó
  más de una institución en "Instituciones que representan", aparecen dos botones arriba de las 6 preguntas
  de Preparación — "📋 Responder de forma general" (por defecto) y "🏫 Responder mencionando cada
  institución" — que cambian el texto con el que empieza cada respuesta:
  - **General**: mantiene la estructura ya existente, "En las Instituciones Educativas Eduardo Santos,
    Gabriel García Márquez y Jairo Mosquera Moreno; " — con una diferencia: los nombres de las IE ahora se
    capitalizan (solo la inicial de cada palabra en mayúscula) en vez de mostrarse tal cual están guardados
    en mayúsculas sostenidas.
  - **Específico**: una frase por institución, una tras otra — "en la Institución Educativa Eduardo Santos en
    la Institución Educativa Gabriel García Márquez en la Institución Educativa Jairo Mosquera Moreno " —
    para que quien responde pueda intercalar su comentario propio de cada una.
  Cambiar de modo solo reemplaza el texto en los campos que siguen vacíos o que todavía tienen exactamente el
  texto inicial anterior sin tocar — nunca borra algo que ya se haya escrito encima. Con una sola institución
  representada, ambos formatos son idénticos, así que los botones no se muestran.
- **Nota de transparencia**: verificado de forma estática (`node --check` con el falso positivo ya conocido
  al extraer el bloque de Index.html; IDs sin duplicar; balance de div/section/header/footer/label/textarea/
  button en Index.html). No probado en vivo desde un navegador en este entorno — la interpretación de "las
  respuestas no deben aparecer como si fueran de una sola IE" se resolvió reformulando el encabezado como
  "Aportes de (rol)"; si la intención era otra (p. ej. cambiar cómo se agrupan/almacenan los aportes en el
  origen de datos), avisar para ajustarlo.

## 4.44 Trigésimo segundo lote: aportes de invitados agrupados por rol, no por IE

- **"Los aportes se deben reunir como Estudiantes, egresados y acudientes y no como aportes de (IE)"**:
  corrige la interpretación del lote anterior — "Aportes ya enviados por estudiantes, egresados y adultos
  responsables invitados" (Selección de IE) ahora agrupa en primer lugar por ROL, con tres secciones fijas
  "🧑‍🎓 Estudiantes (N)", "🎓 Egresados(as) (N)" y "👨‍👩‍👧 Acudientes (N)" (cada una solo aparece si hay al
  menos un aporte de ese rol) — ya no hay un `<details>` por institución. Dentro de cada sección, cada aporte
  individual sigue mostrando la institución de la que viene (queda archivado bajo una sola en la base de
  datos) como una pequeña etiqueta antes de sus respuestas, seguida de sus secciones/preguntas diligenciadas
  — igual que antes, solo que reordenado: rol primero, institución después, en vez de institución primero.
  La función se reescribió por completo en `Components.html`
  (`renderAportesInvitadosGrupo`/`_seccionAportesInvitadosPorRol_`): aplana los datos que ya llegan agrupados
  por IE desde el servidor (`obtenerAportesInvitadosGrupo`, sin cambios) y los reagrupa por rol en el
  cliente, sin tocar el modelo de datos del servidor.
- **Nota de transparencia**: verificado de forma estática (`node --check`). No probado en vivo desde un
  navegador en este entorno.

## 4.45 Trigésimo tercer lote: botón general/específico en cada pregunta, validación antes del gif, confirmación por rol, y renombre a "Profundizaciones"

- **"'Como marcaron más de una institución...' debe aparecer en cada pregunta"**: el botón general/específico
  (lote anterior) ya no vive en un solo bloque arriba de las 6 preguntas — ahora se repite DENTRO de cada una
  de las 6 (`renderPreguntasInvitado` recibe un nuevo parámetro `mostrarBotonModoTexto`). El modo elegido
  sigue siendo un solo estado compartido (`estadoInvitado.modoTextoInicial_`): un clic en cualquiera de las
  12 copias del botón (2 por pregunta × 6 preguntas) actualiza las 6 respuestas y dejan las 12 copias con el
  mismo estado activo/inactivo — un único listener delegado en `document` (`[data-modo-texto-btn]`) en vez de
  12 listeners individuales, porque estos botones se repintan cada vez que se abre la pantalla.
- **"No aparece pregunta para articulación SENA ni para Profundizaciones"**: verificado en el repositorio —
  ambas preguntas SÍ están en el HTML desplegado desde el lote anterior (@43); lo más probable es una página
  ya abierta desde antes de ese despliegue (las apps de Google Apps Script no se actualizan solas en una
  pestaña ya cargada — hace falta recargarla por completo). De paso, se renombró la segunda pregunta de
  "Intensificaciones" a **"Profundizaciones dentro de la institución educativa"** (columna
  `PROFUNDIZACIONES_GRUPO`, antes `INTENSIFICACIONES_GRUPO` — el cambio es seguro porque la pregunta se
  agregó apenas en el lote anterior, sin datos reales todavía) — si el problema persiste tras recargar la
  página por completo, avisar para seguir investigando.
- **"¿Enviar sus aportes para EDUARDO SANTOS? no debe aparecer como x IE sino como rol"**: el mensaje de
  confirmación al enviar los aportes de un invitado cambió de "¿Enviar sus aportes para \<institución>?" a
  "¿Enviar los aportes de \<rol>?" (estudiantes actuales / egresados(as) / adultos responsables de un(a)
  estudiante), igual que ya se había corregido para el listado agrupado de "Aportes ya enviados".
- **"'Respondan todas las preguntas antes de enviar' debe aparecer antes de iniciar carga de gif"**: antes,
  al pulsar "Enviar nuestros aportes" con preguntas sin responder, se mostraba el gif y se hacía el viaje al
  servidor ANTES de que apareciera este mensaje (el servidor ya lo rechazaba, pero solo después). Ahora se
  valida primero en el propio dispositivo (que las 6 respuestas no estén vacías) y, si falta alguna, el
  mensaje aparece de inmediato sin mostrar el gif ni llamar al servidor.
- **"Si ya hay barra de progreso, quitar gif"**: para los tipos de carga que muestran la barra de progreso
  simulada ("informe" e "invitado"), ya no se muestra ningún gif/emoji encima — solo el texto y la barra, sin
  duplicar la señal de "algo está avanzando".
- **Nota de transparencia**: verificado de forma estática (`node --check` con el falso positivo ya conocido
  al extraer el bloque de Index.html; IDs sin duplicar; balance de
  div/section/header/footer/label/textarea/button en Index.html). No probado en vivo desde un navegador en
  este entorno.

## 4.46 Trigésimo cuarto lote: perfil de prueba con 2 IE y correos reales de prueba

- **"Crear un perfil de prueba... crea un grupo con 2 IE de prueba con los correos..."**: `GRUPO-PRUEBA` (2
  IE ficticias, `IE-PRUEBA-1`/`IE-PRUEBA-2`, ya existía en `Tests.gs` desde un lote anterior) ahora también
  les asigna un correo real de prueba en `CaracterizacionIE` — `jhonefrainsanchez@gmail.com` (IE-PRUEBA-1) y
  `hablaconhelprofe@gmail.com` (IE-PRUEBA-2; se corrigió el dominio, escrito como "gmaill.com" en el pedido) —
  para poder recorrer TODO el flujo, incluido el envío del informe por correo (que copia a los correos
  institucionales de las IE del grupo), sin tocar el correo de ninguna IE real. `testCrearGrupoDePrueba()` es
  idempotente: si `GRUPO-PRUEBA` ya existía, no duplica sus filas, pero siempre revisa/actualiza los correos.
  Se agregó `testCrearPerfilPruebaCompleto()` que encadena crear el grupo + generar accesos + mostrar el
  código, en una sola ejecución.
- **Nota de transparencia**: verificado de forma estática (`node --check`). No probado en vivo desde un
  navegador en este entorno; se intentó ejecutar `testCrearPerfilPruebaCompleto()` de forma remota
  (`clasp run`) — ver el resultado exacto en la respuesta de este lote.

## 4.47 Trigésimo quinto lote (parcial): pantalla de acceso — invitados ocultos, número de grupo, modo de visualización

Primer lote de un backlog mucho más grande (Documento Orientador FEM2026, sección "Encuentro voces que
construyen territorio" — ver más abajo "Backlog pendiente" para el resto). Este lote cubre solo la parte
segura e independiente de la sección B del pedido del usuario:

- **"Ocultar el perfil de invitados"**: el botón "🎓 Somos invitados..." (y su separador "o") ya no se
  muestran en la pantalla de Acceso — quedan en el HTML con `class="oculto"`, sin borrar el flujo completo de
  invitados (por si se necesita reactivar más adelante).
- **"Mostrar el número de grupo en tamaño grande... debajo de 'Acceso del grupo' y encima de 'Código de
  acceso'"**: se agregó `#numeroGrupoAcceso`, que muestra el mismo dato ya disponible (`NOMBRE_GRUPO_ACCESO`,
  p. ej. "Grupo 3") con tipografía grande, justo en esa posición.
- **"Modo de visualización... 'No podrá ingresar información hasta que ingrese el código del grupo'"**: se
  agregó como aviso permanente dentro de la propia pantalla de Acceso (en vez de una pantalla nueva aparte,
  ya que esa pantalla YA es, por definición, el único punto donde no se puede ingresar nada hasta escribir el
  código — una pantalla separada habría sido redundante con la misma).
- **"Investigar y corregir el error intermitente 'Ocurrió un error de comunicación'"**: revisado el código —
  el manejo de errores ya está correcto en ambos lados: cada RPC de servidor pasa por `ejecutarRpcSeguro_`
  (Utils.gs, desde un lote anterior), que atrapa cualquier excepción y devuelve `{ok:false, mensaje:...}` en
  vez de dejarla sin controlar; y el cliente (`llamarServidor`, JS.html) solo muestra "Ocurrió un error de
  comunicación..." cuando `google.script.run` falla a nivel de TRANSPORTE (no cuando el servidor responde con
  un error de negocio, que ya llega con un mensaje propio). No se encontró un error de código que lo cause —
  es coherente con una falla intermitente de red/conexión, no con un bug de lógica. Se evaluó agregar un
  reintento automático, pero se descartó: varias operaciones (`appendRow` en `iniciarAccesoInvitado`,
  `registrarParticipante`, `guardarActorConectaEduca`, etc.) NO son idempotentes — si el servidor sí llegó a
  ejecutar la escritura pero la respuesta se perdió en el camino, un reintento automático podría duplicar esa
  fila. Diagnosticar la causa real (¿tiempos de espera de `LockService`? ¿llamadas lentas bajo carga?
  ¿problemas de red del dispositivo?) requiere ver el Registro de ejecuciones de Apps Script en el momento
  real en que ocurre, algo a lo que no tengo acceso en este entorno (`clasp run` sigue bloqueado por el mismo
  permiso de cuenta reportado en un lote anterior). Si vuelve a pasar, lo más útil sería anotar la hora exacta
  y qué acción se estaba haciendo, para buscarlo en ese registro.
- **Nota de transparencia**: verificado de forma estática (`node --check` con el falso positivo ya conocido;
  IDs sin duplicar; balance de etiquetas en Index.html). No probado en vivo desde un navegador.

### Backlog pendiente (Documento Orientador FEM2026 — resto del pedido)

El resto del pedido (división en dos secciones independientes "Encuentro de voces..." / "ConectaEduca" con
navegación, consentimientos, asistencia y valoración propios de cada una, nueva "Sesión de socialización" con
temporizador, reemplazo completo de las preguntas de Sesión 1, y el nuevo campo de estamento) es una
reestructuración grande y profundamente interdependiente — no es seguro intentarla de golpe sin ir verificando
cada pieza. El usuario confirmó el orden sugerido para los próximos lotes: **C (división estructural en dos
secciones) → D (Encuentro de voces...) → E (ConectaEduca) → F (valoraciones) → G (estamento)**. El listado
real de IE por grupo NO hace falta para la "Sesión de socialización" (item 10): una sola persona ingresa los
datos, así que basta con reutilizar el listado de IE ya disponible por grupo (`estado.instituciones`); el
usuario aclaró que ese listado real de instituciones solo hace falta más adelante para la caracterización, no
para esto.

## 4.48 Trigésimo sexto lote: renombrado global "Foro Comunal" → "Encuentro de voces que construyen territorio"

Cubre el item 1 del Documento Orientador FEM2026 (Sección A), hecho como lote independiente (spec del
usuario: "Do it separately", antes de encarar la división en dos secciones de la Sección C).

- **`Config.gs`**: se intercambiaron los valores de `NOMBRE_FORO` y `SUBTITULO` en `CONFIG_POR_DEFECTO_` —
  ahora `NOMBRE_FORO = "Encuentro de voces que construyen territorio"` (el nombre grande, en el encabezado,
  el título de pestaña, los asuntos de correo, la portada del informe, etc. — todo lo que ya usaba
  `config.NOMBRE_FORO`/`NOMBRE_FORO` como variable de plantilla se actualizó automáticamente, sin tocar cada
  archivo) y `SUBTITULO = "Foro Educativo Comunal Neiva 2026"` (pasa a ser el subtítulo entre comillas debajo
  del título, y aparece igual en la portada del informe vía `config.SUBTITULO`). **Importante**: esto solo
  cambia el valor por defecto (`CONFIG_POR_DEFECTO_`, usado al sembrar la hoja `ConfiguracionComunal` la
  primera vez o como fallback); si la hoja `ConfiguracionComunal` de la instalación real ya tiene una fila
  `NOMBRE_FORO`/`SUBTITULO` con el valor viejo, hay que actualizarla ahí manualmente (o borrar esas dos filas
  para que vuelva a tomar el default) — el código no sobrescribe valores ya guardados en la hoja.
- Se dejó sin tocar todo identificador interno no visible para el usuario, por riesgo de romper continuidad
  de datos ya existentes: nombres de hoja (`GruposComunal`, `ConfiguracionComunal`, `ParticipacionComunal`,
  `Sesion1Comunal`, `InformesComunal`, `EnviosDiferidosComunal`, `HOJA_RESPONSABLES_COMUNAL_`,
  `HOJA_VALORACION_COMUNAL_`), la clave de `PropertiesService` (`SPREADSHEET_ID_COMUNAL`), el valor
  `ID_FORO_COMUNAL: "FEC-NEIVA-2026"` (prefijo real de los códigos de acceso ya generados y distribuidos,
  visible en el placeholder "FEC-XXXXX" del campo de código), y los comentarios de encabezado de cada
  archivo (`* X.gs — Foro Educativo Comunal Neiva 2026`, puramente documentales). Renombrar cualquiera de
  estos habría hecho que el código dejara de encontrar hojas/propiedades/códigos ya existentes.
- **Textos de UI reemplazados uno por uno** (no un reemplazo ciego de todo el archivo, por convivir en los
  mismos archivos con los identificadores internos de arriba que no debían tocarse):
  - `Index.html`: "Presentación del Foro" → "Presentación del Encuentro"; "Ruta del Foro Educativo Comunal" →
    "Ruta del Encuentro de voces que construyen territorio"; el párrafo de bienvenida, el aviso de
    sistematización en el consentimiento, las 4 preguntas de valoración + la abierta condicional, el paso a
    paso final, el botón "Finalizar..." y el texto de despedida.
  - `AsistenciaPublica.html`: la misma sección de valoración (texto idéntico al de `Index.html`) y el mensaje
    de agradecimiento final.
  - `JS.html`: las 3 variantes de la pregunta abierta de valoración (según nota 😭/😬/🎉) y el aviso de a qué
    correos institucionales llega el informe.
  - `Access.gs`: el mensaje de "se habilitará a las [hora]" cuando el acceso está bloqueado por horario.
  - `Correo.gs`: asunto, cuerpo de texto y cuerpo HTML del correo de "recorrido de prueba" (`_construirCorreoRecorridoPrueba_`) — aquí también se intercambiaron nombre/subtítulo, igual que en `Config.gs`, para que el encabezado grande del correo diga "Encuentro de voces que construyen territorio" y el pie de página entre comillas diga "Foro Educativo Comunal Neiva 2026".
  - `Informes.gs`: el párrafo final "Insumos para el Foro Educativo Municipal FEM 2026" que menciona de qué
    evento vienen los resultados consolidados.
- Verificado: `node --check` sobre cada `.gs` (copiados a `.js` en un directorio temporal, ya que `node
  --check` no reconoce la extensión `.gs` directamente) y sobre los bloques `<script>` de `Index.html`,
  `AsistenciaPublica.html` y `JS.html` (con las etiquetas de plantilla de Apps Script recortadas) — todos
  limpios salvo los dos falsos positivos ya conocidos y reconfirmados en cada lote (`Index.html` y
  `AsistenciaPublica.html` fallan solo por `var TOKEN_ACCESO = ;` / `var ID_GRUPO = ;`, variables de
  plantilla que quedan vacías al recortar `<?!= ... ?>` fuera de una petición real de Apps Script — no es un
  bug). IDs sin duplicar y balance de etiquetas OK en los dos archivos HTML tocados. No probado en vivo desde
  un navegador.

## 4.49 Trigésimo séptimo lote: pantalla de elección de sección (Documento Orientador FEM2026, sección C, item 7)

Cubre el item 7 (división de la pantalla de inicio en dos secciones independientes, cada una con su logo y su
propia entrada por código) — el primer paso de la Sección C confirmada con el usuario ("C → D → E → F → G").

- **Nueva pantalla `pantallaEleccionSeccion`** (Index.html), ahora la primera que se ve (antes de
  `pantallaAcceso`): dos tarjetas clicables, una por sección —
  "Encuentro de voces que construyen territorio" (logo Drive `1FyoUu8gT0SZHfJq8i3OUalBQY5ALzghb`, horario
  7:00 a. m.–12:00 m.) y "Conecta Educa" (logo Drive `1IKhtm9U_Yuri06VkCsC7W9-ktt_lIltj`, horario 2:00–5:00
  p. m.). Elegir una guarda la elección (`estado.seccion`) y pasa a `pantallaAcceso`, que ahora muestra el
  logo y el título de la sección elegida arriba del campo de código, con un botón "← Elegir otra sección"
  para volver atrás.
- **Decisión de arquitectura (comunicada al usuario, sin objeción)**: el código de acceso sigue siendo **el
  mismo para ambas secciones** — no se crea un segundo código por grupo. Un GRUPO conserva una sola fila en
  `AccesosGrupo`/una sola sesión; "entrada por código independiente" se interpretó como independencia de
  **presentación** (logo, marca, punto de entrada propio), no como un segundo sistema de credenciales. Esto
  evita tocar el modelo de datos y es coherente con que ambas jornadas (mañana y tarde) son del mismo grupo,
  el mismo día.
- **Item 3 ("clic en Inicio debe ir directo a Ingresar código") reinterpretado en este nuevo esquema**: la
  elección de sección se recuerda en `localStorage` del dispositivo (clave `fec_seccion_elegida`) — en la
  próxima visita desde el mismo dispositivo/navegador, `pantallaEleccionSeccion` se salta automáticamente y
  se entra directo a `pantallaAcceso` con la sección ya aplicada, sin perder la posibilidad de cambiar de
  sección ahí mismo.
- **Alcance deliberadamente acotado a solo el item 7**: por ahora, elegir cualquiera de las dos secciones
  lleva exactamente al mismo recorrido de siempre (Metodología → Consentimiento → Participación → … →
  Sesión 1 → Sesión 2/ConectaEduca → Cierre) — todavía no hay una ruta separada para ConectaEduca (eso es el
  item 8, "trasladar toda la sesión actual de ConectaEduca a su nueva pantalla propia", que depende de tocar
  el motor de navegación central — `ORDEN_PANTALLAS`, `cambiarPantalla`, la barra de progreso navegable — y
  se hace mejor como su propio lote, verificado aparte, dado lo sensible que es ese código para un evento en
  vivo). Queda para el próximo lote.
- **`Config.gs`**: nuevas claves `LOGO_ENCUENTRO_ID`/`LOGO_CONECTAEDUCA_ID` (con los IDs de Drive reales
  dados por el usuario) y bandera `LOGOS_SECCION_PUBLICOS`. **`Drive.gs`**: nueva `asegurarLogosSeccionPublicos_()`
  (mismo patrón que `asegurarLogoDesarrolladorPublico_`) para que ambos logos sean visibles públicamente,
  llamada desde `doGet` (Code.gs) junto a las demás.
- Verificado: `node --check` sobre `Config.gs`/`Drive.gs`/`Code.gs` y sobre los bloques `<script>` de
  `Index.html`/`JS.html`/`AsistenciaPublica.html` (mismos dos falsos positivos ya conocidos, reconfirmados);
  IDs sin duplicar y balance de etiquetas OK en `Index.html`. Se detectó y corrigió en el propio desarrollo
  un bug de orden de inicialización (la lógica de "recordar sección elegida" intentaba fijar `pantallaActual_`
  antes de que esa variable se inicializara más abajo en el archivo, lo que la habría dejado sin efecto) —
  se resolvió moviendo ese bloque a después de la definición de `cambiarPantalla()`. No probado en vivo desde
  un navegador (en particular, no se pudo verificar visualmente que los dos logos de Drive carguen).

## 4.50 Trigésimo octavo lote: ConectaEduca como rama independiente del recorrido (item 8)

Cubre el item 8 (D. Sección "Encuentro de voces que construyen territorio"): "trasladar toda la sesión
actual de ConectaEduca a su nueva pantalla propia" — hasta este lote, ConectaEduca era un paso intermedio
("Sesión 2") del mismo recorrido lineal de Encuentro; ahora es una rama propia, alcanzable directo desde la
puerta de ConectaEduca (Lote 37) sin pasar por Metodología, Preparación ni Sesión 1.

- **Verificación previa (antes de tocar nada)**: se confirmó leyendo `enviarSesion1Definitiva` (Sesion1.gs) y
  `enviarSesion2Definitiva` (ConectaEduca.gs) que el servidor **nunca exigió Sesión 1 enviada para poder
  enviar Sesión 2** — esa restricción ("Continuar a Sesión 2" deshabilitado hasta enviar Sesión 1") era
  puramente un botón deshabilitado en el cliente. Esto significa que darle a ConectaEduca su propia puerta,
  sin pasar por Sesión 1, es seguro: el backend ya trataba ambos envíos como independientes.
- **Nuevas listas de pertenencia por sección** (JS.html): `PANTALLAS_COMUNES_SECCION_` (código, consentimiento,
  participación, caracterización, cierre — todavía compartidas, ver Sección E/F pendiente),
  `PANTALLAS_SOLO_ENCUENTRO_` (Inicio, Presentación, Metodología, Invitados especiales, Preparación completa,
  Sesión 1) y `PANTALLAS_SOLO_CONECTAEDUCA_` (solo `pantallaSesion2`) — con `pantallaPerteneceASeccionActual_()`
  para consultarlas.
- **Ruta de ConectaEduca tras el código**: `intentarValidarAcceso` ahora, para `estado.seccion ===
  "CONECTAEDUCA"`, salta directo a Consentimiento/Participación (sin carrusel de IE, sin Inicio/Presentación/
  Metodología). Y `irTrasConfirmarCaracterizacion_` (que antes SIEMPRE mandaba a "Invitados especiales" o
  "Preparación") ahora, para ConectaEduca, va directo a `pantallaSesion2` — este último fue el punto de fuga
  más importante detectado durante el propio desarrollo: sin este cambio, cualquiera que entrara por la
  puerta de ConectaEduca habría terminado de todos modos en el flujo de Preparación de Encuentro al confirmar
  la caracterización.
- **Ruta de Encuentro tras Sesión 1**: el botón "Continuar a Sesión 2" de `pantallaSesion1` (`data-ir`) ahora
  apunta directo a `pantallaRevisionCierre` — renombrado a `btnContinuarACierreEncuentro`, con su texto
  actualizado a "Continuar a Revisión y cierre" — porque Encuentro ya no visita ConectaEduca en su propio
  recorrido. La función de gate se renombró de `actualizarGateSesion2_` a `actualizarGateContinuarCierre_`
  (mismo comportamiento: deshabilitado hasta que Sesión 1 se envíe de forma definitiva).
- **"Retomar donde se quedó" ahora es consciente de la sección**: `destinoResumenPantalla_` descarta la
  última pantalla guardada (`ULTIMA_PANTALLA`, una sola columna compartida por todo el grupo) si esa
  pantalla no pertenece a la sección con la que se está entrando ahora — evita que alguien que entre por
  ConectaEduca en la tarde sea enviado a una pantalla de Encuentro que el grupo dejó a medias en la mañana
  (o viceversa). El cálculo de "hasta dónde desbloquear la barra de progreso según lo avanzado en la nube"
  también se separó por sección (antes usaba `sesion1Enviada`/`sesion2Enviada` mezclados sin distinguir).
- **Barra de progreso navegable con pasos distintos por sección**: `PASOS_PROGRESO_ENCUENTRO_` (Participación,
  Caracterización, Preparación, Sesión 1, Cierre) y `PASOS_PROGRESO_CONECTAEDUCA_` (Participación,
  Caracterización, Conecta Educa, Cierre). `inicializarBarraProgresoNav_` (antes una IIFE de una sola vez) se
  convirtió en `reconstruirBarraProgresoNav_()`, invocable varias veces — se llama al cargar la página (lista
  por defecto de Encuentro, antes de elegir sección) y de nuevo cada vez que se elige/cambia de sección desde
  `aplicarSeccionElegida_` (Lote 37).
- **"🏠 Inicio" (encabezado) también depende de la sección**: para Encuentro sigue yendo a la pantalla de
  bienvenida; para ConectaEduca (que no tiene bienvenida propia) va a Participación, su primer paso real.
- **Deliberadamente NO tocado en este lote** (para no ensanchar el cambio más allá de la navegación): la
  pantalla "Revisión y cierre" sigue siendo una sola para ambas secciones (un único informe, una sola
  valoración) — su encabezado todavía dice "Valoración del Encuentro de voces..." incluso para quien solo
  pasó por ConectaEduca, y "Revisar todo antes de enviar" sigue mostrando las secciones de Sesión 1 Y
  ConectaEduca aunque el grupo solo haya usado una de las dos ramas (se ven en blanco las que no aplican).
  Arreglar esto de raíz es exactamente el trabajo de los items 19-22 (valoraciones separadas) — items 3 y 12
  (texto de rutas/horario en la tarjeta de Metodología, que sigue describiendo el recorrido lineal antiguo)
  también quedan pendientes, ya notados en el lote anterior.
- Verificado: `node --check` sobre los bloques `<script>` de `Index.html`/`JS.html`/`AsistenciaPublica.html`
  (los mismos dos falsos positivos ya conocidos, reconfirmados) y revisión manual completa del grafo de
  navegación (todos los `data-ir` y llamadas a `cambiarPantalla` de Index.html/JS.html) para confirmar que
  ninguna pantalla compartida deja una ruta hacia una pantalla de la sección contraria. IDs sin duplicar y
  balance de etiquetas OK en `Index.html`. No probado en vivo desde un navegador — en particular, no se
  probó de punta a punta el recorrido de ConectaEduca con datos reales de `GRUPO-PRUEBA`.

## 4.51 Trigésimo noveno lote: Sesión de socialización reemplaza a Preparación (items 9 + 10)

Cubre juntos los items 9 ("ocultar Preparación y todas sus subpáginas") y 10 ("añadir una nueva Sesión de
socialización") del Documento Orientador FEM2026, sección D — se hicieron en el mismo lote porque ocultar el
uno sin tener listo el otro habría dejado un hueco en el recorrido de Encuentro (nada llevaría de
Confirmación de caracterización a Sesión 1). También cubre el item 11 ("trasladar la página pop-up actual a
la Sesión 1").

- **Qué se ocultó (item 9)**: `pantallaInvitadosEspeciales`, `pantallaPreSocializacion`,
  `pantallaSeleccionIEPreparacion` y `pantallaPreparacionIE` — la "Sesión de preparación" completa, donde
  cada IE elegía un(a) responsable y editaba un resumen sugerido (extraído de su Informe Ejecutivo real)
  antes de enviarlo. Mismo patrón de todo el proyecto: se dejan en el HTML/backend sin borrar (útiles si se
  necesita reactivar esa lógica), simplemente ya no son alcanzables — se sacaron de
  `PANTALLAS_SOLO_ENCUENTRO_` (Lote 38), así que `pantallaPerteneceASeccionActual_` las rechaza y ningún
  botón ni llamada a `cambiarPantalla` restante en el código apunta hacia ellas (verificado con una revisión
  completa de todos los `data-ir`/`cambiarPantalla` de Index.html y JS.html).
- **Por qué se puede ocultar sin perder valor**: la Sesión de preparación producía contenido que luego se
  socializaba oralmente en Sesión 1 (aportes leídos por cada IE). El item 16 (todavía pendiente) va a
  reemplazar las preguntas actuales de Sesión 1 por las preguntas reales del Foro Educativo Institucional —
  el contenido ya no se prepara por escrito de antemano por cada IE, sino que se construye en vivo a partir
  de lo que cada institución socialice oralmente ese día, con el nuevo mecanismo de este lote llevando el
  registro de quién ya pasó.
- **Qué se agregó (item 10) — `Socializacion.gs` (nuevo) + `pantallaSesionSocializacion` (Index.html)**:
  una sola persona (spec del usuario: "una persona va a ingresar los datos") lleva en vivo el checklist de
  qué institución ya socializó, con un temporizador de apoyo de 10 minutos por IE (`+2 minutos`,
  `Finalizar`). Al dar "Finalizar" (o en cualquier momento, sin depender del temporizador) se abre el modal
  "Datos relevantes de [IE]" (`modalDatosRelevantesIE`, Modal.html) para anotar lo compartido; guardar ahí
  marca esa IE como socializada Y guarda la nota, en una sola llamada (`guardarSocializacionIE`).
  - Nueva hoja `SocializacionIE` (mismo patrón de clave compuesta `ID_GRUPO|ID_IE` que `PreparacionIE`/
    `ParticipacionEstamento`, columna `CLAVE`): `obtenerSocializacionGrupo(idGrupo)` arma el checklist
    completo (todas las IE del grupo, vía `obtenerInstitucionesDelGrupo` ya existente) cruzando lo guardado;
    `guardarSocializacionIE(idGrupo, tokenSesion, dispositivoId, idIE, socializo, datosRelevantes)` hace el
    UPSERT (con el mismo chequeo de sesión activa que el resto de escrituras del proyecto). RPCs nuevas:
    `rpcObtenerSocializacionGrupo`/`rpcGuardarSocializacionIE` (Code.gs).
  - El temporizador es puramente de cliente (no se guarda en el servidor ni se sincroniza entre
    dispositivos) — es un apoyo visual para quien modera en vivo, no un dato del registro; lo único que se
    persiste es el resultado (marcado + nota).
  - Ruta: Confirmación de caracterización → Sesión de socialización → Sesión 1 (reemplaza a "→ invitados
    especiales/Sesión de preparación → Sesión 1"). Nuevo paso "Socialización" en la barra de progreso
    navegable de Encuentro (`PASOS_PROGRESO_ENCUENTRO_`), reemplazando al antiguo "Preparación".
- **Item 11 — pop-up de recursos trasladado a Sesión 1**: la ventana emergente de documentos de apoyo
  (Informe Ejecutivo, etc., `modalRecursosSesion1`) aparecía automáticamente solo en Preparación de la IE;
  como esa pantalla ya no es alcanzable, ahora se abre automáticamente (una sola vez por sesión, mismo flag
  `estado.recursosSesion1Mostrado`) al entrar a Sesión 1 — la lógica vive ahora en `cargarRecursosSesion1()`
  en vez de en `cargarPreparacionIE()` (que queda intacta pero ya no se ejecuta).
- Verificado: `node --check` sobre `Socializacion.gs` y el resto de `.gs` tocados (`Code.gs`, `Config.gs`), y
  sobre los bloques `<script>` de `Index.html`, `JS.html` (limpio), `Components.html` y `Modal.html`; IDs sin
  duplicar y balance de etiquetas OK en `Index.html`/`Modal.html`. Revisión manual del grafo completo de
  navegación (como en el lote anterior) para confirmar que ninguna pantalla compartida sigue apuntando a la
  vieja Preparación. No probado en vivo desde un navegador — en particular, no se probó el temporizador ni
  el guardado real del checklist con `GRUPO-PRUEBA`.

## 4.52 Cuadragésimo lote: ruta/horario de Metodología (item 12) y solo asistencia en PDF (item 13)

- **Item 12**: la tarjeta "Ruta del Encuentro..." (pantallaMetodologia, Encuentro-only) todavía describía el
  recorrido lineal de antes de la Sección C/D (incluía "Sesión de preparación" e incluso "Sesión 2:
  ConectaEduca" como si fueran pasos de Encuentro). Se reescribió la lista de 8 a 7 pasos (sin ConectaEduca,
  con "Sesión de socialización" en vez de "Sesión de preparación") y se agregó un aviso de horario arriba de
  la lista: "Este Encuentro se realiza de 7:00 a. m. a 12:00 m. Conecta Educa es una jornada aparte, de
  2:00 p. m. a 5:00 p. m., con su propia puerta de acceso (mismo código del grupo) — no hace falta terminar
  aquí para poder entrar a Conecta Educa."
- **Item 13**: se ocultaron `btnMetodoQR`/`panelMetodoQR` (envueltos en `<div class="fila-botones oculto">`,
  sin borrar — mismo patrón de todo el proyecto) y `panelMetodoListado` quedó siempre visible, sin necesidad
  de elegirlo con un clic. En `intentarValidarAcceso` (JS.html), si el método guardado del grupo no es ya
  `"LISTADO"` (grupo nuevo, o un valor `"QR"` de antes de este cambio) se fuerza a `"LISTADO"` automáticamente
  llamando a `elegirMetodoAsistencia("LISTADO")`, que ya se encarga de guardarlo — no se tocó `Asistencia.gs`:
  el backend siempre aceptó cualquiera de los dos métodos, el cambio es puramente de qué se ofrece en la UI.
  Se agregó un botón "📥 Descargar formato de asistencia" que enlaza al PDF oficial del Encuentro (Drive
  `1rB1diE0iMthTDclJSxfazqJzld76nRx8`, dado por el usuario) — nueva clave `FORMATO_ASISTENCIA_ENCUENTRO_ID`
  en `Config.gs`, pasada a la plantilla desde `Code.gs` e insertada directo en el `href` con
  `<?!= FORMATO_ASISTENCIA_ENCUENTRO_ID ?>` (no hizo falta pasarlo también a JS.html: es un enlace estático,
  no una imagen que necesite `urlImagenDrive`/miniatura).
- **Alcance del item 13 acotado a Encuentro por ahora**: `pantallaParticipacion` sigue siendo la misma
  pantalla compartida con ConectaEduca (Lote 38) — por ahora ambas secciones ven el mismo formato de
  asistencia del Encuentro. El formato propio de ConectaEduca (Drive `1sO3ddWU9PcL3CMnygYgSleTxs17qb_lS`,
  item 18) requiere separar la pantalla de asistencia por sección, que es justamente el trabajo de la
  Sección E — se hace ahí, no aquí, para no tocar dos veces la misma pantalla.
- Verificado: `node --check` sobre `Config.gs`/`Code.gs` y los bloques `<script>` de `Index.html`/`JS.html`
  (limpio, mismos dos falsos positivos ya conocidos en Index.html/AsistenciaPublica.html); IDs sin duplicar y
  balance de etiquetas OK en `Index.html`. No probado en vivo desde un navegador.

## 4.53 Cuadragésimo primer lote: consentimiento en dos partes + listado de asistencia obligatorio (items 14+15)

- **Item 15 — consentimiento dividido en dos partes**: la tarjeta de `pantallaConsentimientoGrupo` ahora
  tiene "Parte 1 — Encuentro de voces que construyen territorio (7:00 a. m. a 12:00 m.)" con el desglose
  exacto del Documento Orientador (sesión de socialización 7:00–9:30, receso 9:30–10:00, trabajo colectivo
  10:00–12:00) y "Parte 2 — Conecta Educa (2:00 p. m. a 5:00 p. m.)", aclarando que se ingresa por su propia
  puerta con el mismo código del grupo. El resto del texto (estructura sugerida no fija, espacio de aporte
  propio, uso de la información) se mantuvo igual.
- **Item 14 — listado de asistencia obligatorio antes de continuar**: nueva sección "Listado de asistencia"
  en la misma pantalla, con el mismo formato PDF ya usado en Participación (item 13,
  `FORMATO_ASISTENCIA_ENCUENTRO_ID`) y tres formas de gestionarlo: **Descargar** (enlace directo),
  **Copiar enlace** (portapapeles, mismo patrón que el enlace de asistencia QR) y **Enviar por correo**
  (nueva función `enviarFormatoAsistenciaPorCorreo` en Correo.gs, mismo patrón que
  `enviarEnlaceInvitados` — nueva RPC `rpcEnviarFormatoAsistenciaPorCorreo`). El botón "Aceptar y continuar"
  ahora exige DOS condiciones (`actualizarBotonAceptarConsentimiento_`, JS.html): el checkbox marcado Y
  `estado.listadoAsistenciaGestionado` en `true` (se pone en `true` al usar cualquiera de las tres
  opciones) — con un aviso visible ("Antes de continuar, descarguen, copien el enlace o envíen por correo el
  listado...") cuando el checkbox ya está marcado pero falta gestionar el listado. También se agregó la
  instrucción pedida: "Quien reciba este listado debe encargarse de imprimirlo y, al finalizar, reportar la
  cantidad total de asistentes y los cargos de los asistentes por institución educativa."
- **"Mover el campo 'cantidad de asistentes por estamento e institución' al final del formulario"**: se
  revisó `pantallaParticipacion` y esa tarjeta (`resumenParticipacionEstamento`/`contenedorParticipacionEstamento`)
  ya es la última tarjeta de contenido antes del botón "Continuar" (después de Método de asistencia,
  Participación y asistencia, y Responsable de envío) — no hizo falta moverla, ya estaba al final.
- Verificado: `node --check` sobre `Correo.gs`/`Code.gs`/`Config.gs` y los bloques `<script>` de
  `Index.html`/`JS.html` (limpio, mismos dos falsos positivos ya conocidos); IDs sin duplicar y balance de
  etiquetas OK en `Index.html`. No probado en vivo desde un navegador — en particular, no se probó el envío
  real de `enviarFormatoAsistenciaPorCorreo` ni el flujo completo de gate en el checkbox.

## 4.54 Cuadragésimo segundo lote: las 11 preguntas reales del FEI reemplazan a Sesión 1 (item 16)

El más grande de la Sección D — reemplaza las 8 preguntas genéricas de Sesión 1 (REFLEXIONES, DESAFIOS,
APUESTAS, CONCLUSIONES, PRIORIDADES, PROPUESTAS_COLECTIVAS, ACUERDOS, RUTA) por las 11 preguntas reales del
punto 3.1 del Documento Orientador FEM2026, agrupadas en sus 3 temas originales, con nuevas claves de campo
(`FEM2025_P1/P2`, `CURRICULO_P1..P5`, `GOBIERNO_P1..P4`):

- **1. Avances FEM2025 y políticas públicas** (2 preguntas): avances en los retos/propósitos del FEM2025;
  avances en la implementación de preescolar (jardín, prejardín).
- **2. Currículos que nazcan del territorio** (5 preguntas): pertinencia de los currículos con la realidad
  territorial; acciones para transformarlos (mín. 3, máx. 5 acciones); equipos de trabajo institucional;
  articulación de esos equipos; mecanismos de seguimiento.
- **3. Gobierno en instituciones educativas donde participar signifique decidir** (4 preguntas): participación
  y democracia en la toma de decisiones; acciones para una participación más incidente (mín. 3, máx. 5
  acciones); equipos de trabajo institucional; mecanismos de seguimiento.
- **"Mismos criterios de contador de palabras y tipo de selección que ya usa la app"**: se interpretó como
  aplicar a las 11 preguntas el mismo criterio que ya tenían las 4 preguntas más exigentes de la Sesión 1
  original (50-400 palabras, contador en vivo) y hacerlas todas obligatorias para el envío definitivo — antes
  la app tenía dos niveles distintos (4 preguntas con rango de palabras + 4 sin rango); como las 11 nuevas
  son todas preguntas orientadoras igual de centrales del FEI, no había una forma natural de mantener esa
  distinción de dos niveles, así que se unificó en el nivel más exigente que ya existía.
- **Sin cambio de tipo de campo**: las preguntas "enunciar mínimo 3, máximo 5 acciones concretas" siguen
  siendo textareas de texto libre (igual que el resto) — no se construyó un input de lista/repetible nuevo,
  ya que el pedido no especificaba una interfaz distinta y el criterio de "mismo tipo de selección que ya usa
  la app" no señala ningún campo de opción múltiple existente del que copiar un patrón.
- **`Sesion1.gs`**: `cabecerasSesion1Comunal_()` con las 11 columnas nuevas; REFLEXIONES..RUTA (las 8 viejas)
  se movieron al bloque de "columnas heredadas, ya no se muestran en pantalla" (mismo patrón que
  PROPUESTAS_IE/EXPERIENCIAS/etc de una versión anterior) — se conservan por si algún grupo ya las hubiera
  diligenciado, pero no se pierden ni se leen más. `CAMPOS_SESION1_CON_RANGO_PALABRAS_` y
  `CAMPOS_SESION1_OBLIGATORIOS_` apuntan ahora a las 11 nuevas.
- **Punto de fuga encontrado y corregido durante el propio desarrollo**: `CAMPOS_SESION1` en JS.html (el
  arreglo que controla tanto `recolectarCamposSesion1()`/autoguardado como `cargarSesion1()`/carga desde la
  nube) todavía tenía los nombres de campo viejos — sin este cambio, los 11 textareas nuevos habrían quedado
  completamente desconectados del guardado y la carga (visualmente presentes pero inertes). Se corrigió junto
  con `PREGUNTAS_SOLO_LECTURA_SESION1_` (vista de solo lectura tras el envío) y `CAMPOS_REVISAR_SESION1_`
  ("Revisar todo antes de enviar", Revisión y cierre).
- **`Informes.gs`**: el informe generado ahora presenta las 11 preguntas agrupadas por sus 3 temas, en vez
  de los dos bloques "Consolidado de Socialización"/"Construcción colectiva del grupo" (fusionados en uno).
- **`Tests.gs`**: `testFlujoCompletoGrupoPrueba` ahora llena las 11 claves nuevas (con las 4 originales ya
  no habría alcanzado el mínimo de campos obligatorios que exige `enviarSesion1Definitiva`).
- **"Aportes de preparación por institución" (`tarjetaSocializacionPreparacion`, dentro de Sesión 1) se
  ocultó**: dependía enteramente de la vieja "Preparación" (oculta desde el item 9) — sin ningún productor de
  datos, siempre se habría visto vacía junto a las preguntas nuevas. Se dejó con `class="oculto"` sin
  borrarla, y se quitó la llamada a `cargarSocializacionPreparacion()` en `cambiarPantalla` (la función en sí
  queda intacta). El botón "Atrás" de Sesión 1 apuntaba a `pantallaSeleccionIEPreparacion` (también oculta,
  del mismo item 9) — se corrigió a `pantallaSesionSocializacion`, su verdadero paso anterior desde el Lote 39.
- Verificado: `node --check` sobre `Sesion1.gs`/`Informes.gs`/`Tests.gs` y los bloques `<script>` de
  `Index.html`/`JS.html` (limpio, mismos dos falsos positivos ya conocidos); búsqueda exhaustiva en todo el
  proyecto de los 8 nombres de campo viejos para confirmar que solo quedan en el bloque de columnas heredadas
  de `Sesion1.gs` (intencional); IDs sin duplicar y balance de etiquetas OK en `Index.html`. No probado en
  vivo desde un navegador — en particular, no se verificó visualmente el guardado/autoguardado real de los
  11 campos nuevos ni la generación del informe con datos reales.

## 4.55 Cuadragésimo tercer lote: consentimiento y asistencia propios de Conecta Educa (Sección E, items 17+18)

Cierra la Sección E completa: Conecta Educa deja de compartir el consentimiento informado y el formato de
asistencia del Encuentro.

- **Nueva pantalla `pantallaConsentimientoConectaEduca`** (Index.html), paralela a `pantallaConsentimientoGrupo`
  pero con texto propio: describe solo Conecta Educa (2:00-5:00 p. m., sin las Partes 1/2 del Encuentro),
  explica la **metodología de la caracterización** (registrar cada actor/entidad del sector productivo o de
  educación superior —nombre, tipo, área de interés, IE interesadas— a medida que se dialogue con ellos,
  reutilizando la sección "Registrar actor o entidad" que ya existe en Sesión 2) y trae su propio bloque de
  listado de asistencia (descargar/copiar enlace/enviar por correo), con el mismo candado del item 14: no se
  puede aceptar sin haber marcado el checkbox Y gestionado el listado.
- **`PANTALLAS_SOLO_CONECTAEDUCA_`/`PANTALLAS_SOLO_ENCUENTRO_` (JS.html)**: `pantallaConsentimientoGrupo` pasó
  de "común a ambas secciones" a exclusiva del Encuentro; `pantallaConsentimientoConectaEduca` se agregó como
  exclusiva de Conecta Educa. `pantallaParticipacion` sigue siendo común a ambas (ver siguiente punto) — dividir
  esa pantalla también no lo pedía el item 17/18 y habría significado duplicar foto de grupo, matriz de
  estamento y responsable de envío sin ninguna diferencia real entre secciones.
- **Columna de consentimiento independiente**: `CONSENTIMIENTO_CONECTAEDUCA`/`FECHA_CONSENTIMIENTO_CONECTAEDUCA`
  (`Access.gs`, `cabecerasAccesosGrupo_`) + `guardarConsentimientoConectaEduca()`, paralela a
  `guardarConsentimientoGrupo()` — un grupo puede aceptar el consentimiento del Encuentro sin haber pasado
  todavía por la puerta de Conecta Educa (y viceversa), así que no podían compartir una sola columna/bandera.
- **Formato de asistencia propio (item 18)**: `FORMATO_ASISTENCIA_CONECTAEDUCA_ID` (Drive
  `1sO3ddWU9PcL3CMnygYgSleTxs17qb_lS`) en `Config.gs`, pasado al template en `doGet` igual que el del Encuentro.
  `enviarFormatoAsistenciaPorCorreo()` (Correo.gs) y `rpcEnviarFormatoAsistenciaPorCorreo` (Code.gs) ahora
  reciben un 5º parámetro `seccion` ("ENCUENTRO" por defecto o "CONECTAEDUCA") que decide qué ID de Drive y qué
  nombre de sesión usar en el enlace/asunto/cuerpo del correo, en vez de duplicar toda la función.
- **`pantallaParticipacion` (compartida) también usa el formato correcto según sección**: el botón de descarga
  ganó `id="btnDescargarFormatoAsistenciaParticipacion"` y `cambiarPantalla` le reescribe el `href` según
  `estado.seccion` cada vez que se entra a esa pantalla — la pantalla en sí no se duplicó, solo el enlace.
- **Gap preexistente corregido de paso**: `FORMATO_ASISTENCIA_ENCUENTRO_ID` (agregado en el Lote 40) nunca se
  había puesto público en Drive — se agregó `asegurarFormatosAsistenciaPublicos_()` (Drive.gs, mismo patrón que
  `asegurarLogosSeccionPublicos_`, bandera propia `FORMATOS_ASISTENCIA_PUBLICOS`) que cubre ambos IDs
  (Encuentro y Conecta Educa) y se llama desde `doGet`.
- **JS.html**: `estado.consentimientoConectaEduca` (paralelo a `estado.consentimientoGrupo`), cargado desde
  `r.consentimientoConectaEduca` en `intentarValidarAcceso`; la rama `estado.seccion === "CONECTAEDUCA"` (que
  antes, desde el Lote 38, mandaba siempre a `pantallaConsentimientoGrupo` — el consentimiento del Encuentro,
  por no existir todavía uno propio) ahora manda a `pantallaConsentimientoConectaEduca`. Bloque completo de
  manejadores (`actualizarBotonAceptarConsentimientoCE_`, descargar/copiar/enviar/aceptar) calcado del
  consentimiento del Encuentro pero con sus propios IDs de elemento — se optó por duplicar en vez de
  generalizar con un solo helper parametrizado porque los IDs de Encuentro no siguen un sufijo consistente
  (`avisoListadoAsistenciaPendiente`, `mensajeConsentimientoGrupo` no tienen el mismo patrón que
  `btnDescargarFormatoAsistenciaConsentimiento`), así que forzar un mapeo genérico habría sido más riesgoso
  que el costo de un bloque paralelo, acotado a una sola pantalla.
- Verificado: `node --check` sobre `Config.gs`/`Drive.gs`/`Code.gs`/`Access.gs`/`Correo.gs`/`Tests.gs` y los
  bloques `<script>` de `Index.html` (mismo falso positivo ya conocido)/`JS.html` (limpio); IDs sin duplicar y
  balance de etiquetas OK en `Index.html`. `Tests.gs` ganó una línea que ejercita
  `guardarConsentimientoConectaEduca` junto a la ya existente de `guardarConsentimientoGrupo`. No probado en
  vivo desde un navegador.

## 4.56 Cuadragésimo cuarto lote: valoraciones separadas y escala numérica (Sección F) + estamento nuevo (Sección G)

Cierra las dos últimas secciones del Documento Orientador FEM2026 — items 19 a 23.

- **Item 19 — Quitar el registro de asistencia de Revisión y cierre**: el bloque "Verifique que todos los
  participantes hayan firmado la asistencia" (panel QR/listado, cantidad declarada, botón "Asistencia
  verificada") se ocultó (`class="oculto"` en un contenedor nuevo, sin borrar nada — mismo criterio de "ocultar,
  no eliminar" del resto del proyecto) porque duplicaba lo que ya se registra y verifica en Participación, y
  ahora esta pantalla se visita una vez por sección. `actualizarBotonGenerarInforme()` dejó de exigir
  `estado.asistenciaVerificada`; `cargarVerificacionFirmantesCierre_()`/`btnAsistenciaVerificada` quedaron sin
  usar (funciones intactas, sin llamar).
- **Item 20 — Valoraciones separadas para Encuentro y Conecta Educa**: `ValoracionComunal` (Valoracion.gs) pasó
  de una fila por grupo a una fila por `CLAVE = ID_GRUPO + "|" + SECCION` (mismo patrón que
  `_claveSocializacion_`/`_claveValoracion_`) — `guardarValoracionGrupo`/`obtenerValoracionGrupo` ganaron un
  parámetro `seccion`. El informe sigue siendo uno solo por grupo (spec sección 18, sin cambios); lo que se
  separó es la valoración, no el documento: `generarInformeCompletoGrupo` ahora exige la valoración de la
  sección que intenta generar (quien llegue primero, Encuentro o Conecta Educa, necesita la suya). El correo de
  envío del informe (`enviarInformeSiCorresponde`) sigue llamando a `obtenerValoracionGrupo(idGrupo)` sin
  sección — a esa altura el informe ya existe sin importar cuál sección lo generó, así que basta con que
  cualquiera de las dos valoraciones exista. `rpcGuardarValoracion`/`rpcObtenerValoracion`/`rpcGenerarInforme`
  (Code.gs) y sus llamadas en JS.html pasan `estado.seccion`.
- **Item 21 — Escala numérica en vez de corazones**: los 5 botones de corazón (relleno acumulado 🤍/❤️) se
  reemplazaron por 5 botones nombrados de selección única — MALO=1, DEFICIENTE=2, REGULAR=3, BUENO=4,
  EXCELENTE=5 — con la clase `.opcion-valoracion`/`.opcion-valoracion-seleccionada` (antes
  `.corazon-valoracion`, renombrada en Index.html/JS.html/CSS.html **y también en AsistenciaPublica.html**, la
  valoración pública/anónima que ya existía desde el Lote 51, para no dejar dos escalas distintas conviviendo
  en la misma app). El cálculo de nota (promedio de 1 a 5) no cambió, solo cómo se recoge cada respuesta.
- **Item 22 — No nombrar "el foro"/"Conecta Educa" en las preguntas**: reformuladas con "la sesión de hoy" en
  las 4 preguntas, los 4 textos de "mejora", la pregunta abierta final y los 3 mensajes de resultado
  (`actualizarResultadoValoracion`, según la nota promedio) — en Index.html y, por consistencia, también en
  AsistenciaPublica.html. **Decisión clave**: como el texto ya no distingue qué sección es, no hizo falta
  ninguna variante de contenido por `estado.seccion` — la misma tarjeta de valoración sirve para ambas
  secciones sin condicionales, y el ítem 20 (separación) queda resuelto enteramente por la capa de datos
  (`CLAVE` con `SECCION`), no por la interfaz.
- **Item 23 — "Funcionario Secretaría de Educación" en el campo de estamento**: se agregó a
  `ESTAMENTOS_PARTICIPACION_` (ParticipacionEstamento.gs), la matriz de conteo por estamento e IE que alimenta
  el informe real — mismo criterio ya usado para "Sector productivo" (no afiliado a una IE en particular, mismo
  código; se registra bajo la columna de la IE que corresponda al contexto). También se agregó al `<select
  id="campoEstamento">` de AsistenciaPublica.html, por consistencia con el mismo catálogo. **Deliberadamente
  NO se tocó** `ROLES_FORO_` (Responsables.gs, catálogo de roles operativos del Foro — Líder, Dinamizador,
  Relator, etc. — tomado textualmente del Documento Orientador, no es un "estamento"), porque el pedido dice
  "estamento/rol" refiriéndose a la categoría de asistente (Rector, Docente, Sector productivo…), no al rol de
  organización del evento.
- Verificado: `node --check` sobre `Valoracion.gs`/`Grupos.gs`/`Code.gs`/`ParticipacionEstamento.gs` y los
  bloques `<script>` de `Index.html`/`AsistenciaPublica.html` (mismos dos falsos positivos ya conocidos, uno
  por archivo)/`JS.html` (limpio); IDs sin duplicar y balance de etiquetas OK en ambos HTML. No probado en vivo
  desde un navegador.

### Documento Orientador FEM2026: las 23 preguntas quedan implementadas

Con este lote se completan los 23 puntos del documento (Secciones A a G). Queda pendiente, como en cualquier
entrega de este tamaño, una prueba end-to-end real en un navegador (crear un grupo de prueba, recorrer las dos
puertas —Encuentro y Conecta Educa— de principio a fin, generar el informe) antes de considerar esto listo
para producción; ningún lote de esta serie se probó fuera de `node --check` y las verificaciones estáticas
descritas en cada sección.

## 4.57 Cuadragésimo quinto lote: "Subir más tarde" en el listado de asistencia

Pedido suelto del usuario, fuera del Documento Orientador: en Participación, "Método de asistencia del grupo"
solo tenía "Subir listado" — quien no tuviera el PDF firmado a mano en ese momento no tenía forma explícita de
aplazarlo, a diferencia de la fotografía del grupo (que sí tiene "Subir más tarde" desde el Lote de la Fase
25/26). Se agregó el mismo patrón: botón `btnListadoMasTarde` junto a "Subir listado" que solo marca
`estado.listadoAsistenciaMasTardeConfirmado = true` y muestra un mensaje tranquilizador — no sube nada, y a
diferencia de la fotografía (que sí es obligatoria para generar el informe, `FOTO_REQUERIDA` en
`generarInformeCompletoGrupo`), el listado nunca tuvo ni tiene un candado equivalente, así que no hizo falta
replicar el segundo aviso de confirmación en "Continuar" ni un punto de re-exigencia en Revisión y cierre
(ítem 19 del lote anterior, además, ya quitó de ahí todo lo relacionado con asistencia) — el listado se puede
subir en cualquier momento después, desde el mismo panel de Participación.

Verificado: bloques `<script>` de `Index.html` (mismo falso positivo ya conocido)/`JS.html` (limpio); IDs sin
duplicar y balance de etiquetas OK. No probado en vivo desde un navegador.

## 4.58 Cuadragésimo sexto lote: advertencia de formato único, confirmación de carpetas por grupo, y Caracterización movida al final

Tres pedidos sueltos del usuario.

- **"Solo se debe subir el formato que se descarga, se imprime y se llena — no se aceptarán formatos
  diferentes"**: como no es técnicamente viable validar en el servidor que un PDF subido sea realmente el
  formato oficial (no hay forma de verificar la plantilla de un PDF sin OCR/procesamiento complejo, fuera de
  alcance de este pedido), se resolvió con advertencias destacadas (`class="mensaje info"`, con ⚠️ y texto en
  negrita) en los tres puntos donde el grupo interactúa con el listado: los dos consentimientos informados
  (Encuentro y Conecta Educa, junto al botón de descarga) y el panel de subida en Participación (justo antes
  del `<input type="file">`). Es una regla explicada y resaltada para las personas, no una validación técnica.
- **"Generar carpeta por cada grupo con las asistencias en la carpeta de asistencias del proyecto"**: ya
  estaba implementado desde antes (Fase 17) — `asegurarCarpetaAsistenciaGrupo_()` (Drive.gs) crea/reutiliza una
  carpeta `GRUPO N` dentro de `02_ASISTENCIA` (una de las 6 subcarpetas raíz del proyecto en Drive) y
  `subirListadoAsistencia()` (Asistencia.gs) ya sube ahí cada PDF. No se tocó nada — se verificó el código
  existente para confirmarlo.
- **Confirmación de caracterización movida al final del recorrido** (antes iba justo después de Participación,
  como paso intermedio hacia las sesiones de trabajo): el usuario confirmó, ante una pregunta de aclaración,
  que debía ir justo antes de Revisión y cierre — como último chequeo de que la información del grupo está
  correcta, en vez de un paso intermedio. Cambios:
  - `ORDEN_PANTALLAS` (JS.html) y las dos `PASOS_PROGRESO_*_` (barra de progreso navegable) reordenados:
    `pantallaConfirmacionCaracterizacion` pasó de estar entre `pantallaParticipacion` y las sesiones de
    trabajo, a estar entre `pantallaSesion1`/`pantallaSesion2` y `pantallaRevisionCierre`.
  - **Nueva función `irTrasParticipacion_()`** (antes esta lógica de ramificar por `estado.seccion` vivía en
    `irTrasConfirmarCaracterizacion_`): "Continuar" en Participación ahora va directo a
    `pantallaSesionSocializacion` (Encuentro) o `pantallaSesion2` (Conecta Educa), sin pasar por
    Caracterización.
  - **`irTrasConfirmarCaracterizacion_()` simplificada**: ya no ramifica por sección — confirmar la
    caracterización ahora siempre lleva a `pantallaRevisionCierre` (ambas secciones convergen ahí).
  - **Nuevo botón `btnAtrasCaracterizacion`** en la pantalla de Caracterización (reemplaza el `data-ir`
    estático que antes apuntaba a Participación): como la pantalla anterior ahora depende de la sección
    (Sesión 1 o Sesión 2), el "Atrás, corregir" necesita lógica, no un destino fijo.
  - **Botones "Continuar" de Sesión 1 y Sesión 2** (`btnContinuarACierreEncuentro` y el de Conecta Educa):
    apuntaban a `pantallaRevisionCierre`, ahora apuntan a `pantallaConfirmacionCaracterizacion` (con el texto
    actualizado a "Continuar a Confirmación de caracterización").
  - **"Atrás" de Sesión de socialización**: apuntaba a `pantallaConfirmacionCaracterizacion` (su vecino
    anterior en el recorrido viejo), ahora apunta a `pantallaParticipacion` (su vecino anterior real ahora).
  - Las dos referencias a `pantallaConfirmacionCaracterizacion` que quedaron en `pantallaInvitadosEspeciales` y
    `pantallaPreSocializacion` (subpáginas de la vieja "Preparación", ocultas desde el item 9) **no se
    tocaron** — son código muerto inalcanzable, no vale la pena su mantenimiento.
  - La ficha de caracterización, la matriz de participación por estamento y la fotografía siguen funcionando
    igual (se recargan cada vez que se entra a la pantalla, sin importar su posición en el recorrido); el
    único texto ajustado fue la frase introductoria ("...antes de continuar a la valoración y el cierre" en
    vez de "...a las sesiones de trabajo").
- Verificado: bloques `<script>` de `Index.html`/`JS.html` (mismo falso positivo ya conocido en Index, limpio
  en JS); IDs sin duplicar y balance de etiquetas (`div`/`section`/`p`/`button`/`label`/`textarea`/`a`) OK. No
  probado en vivo desde un navegador — en particular, no se verificó visualmente el recorrido completo con las
  nuevas rutas.

## 4.59 Cuadragésimo séptimo lote: pantalla partida y alarmas sutiles en Sesión de socialización

Pedido suelto del usuario sobre el temporizador de la Sesión de socialización (item 10).

- **Pantalla partida**: el panel del temporizador (antes solo el reloj + "+2 minutos"/"Finalizar") ahora se
  divide en dos mitades (`.panel-partido-socializacion`, flex 50/50 que colapsa a una columna en móvil): la
  izquierda con el reloj y los mismos botones de siempre; la derecha con un textarea "Datos relevantes de esta
  IE" y un botón "💾 Guardar" — para consolidar en vivo lo que comparte la institución mientras el temporizador
  corre, en vez de escribirlo después en una ventana emergente.
  - **"Finalizar" ya no abre el cuadro emergente**: ahora guarda directamente lo que haya en el textarea en
    vivo (`guardarDatosRelevantesEnVivo_()`) y cierra el panel — el flujo antiguo (temporizador → modal
    "Datos relevantes" al finalizar) se reemplazó por el flujo en vivo.
  - **El cuadro emergente `modalDatosRelevantesIE` se conserva**, pero solo para el botón "✏️" de una IE que
    YA terminó (no está en temporizador) — permite editar sus notas sin tener que reiniciar el reloj.
  - Al iniciar el temporizador de una IE, el textarea en vivo se prellena con lo que ya tuviera guardado esa
    IE (por si se reanuda una socialización interrumpida).
- **Alarmas sutiles (sonido + visual) en 4 momentos**, todas dentro del mismo `setInterval` que ya llevaba la
  cuenta regresiva (comparación exacta contra el segundo restante, sin riesgo de "saltarse" un umbral porque
  el reloj siempre decrementa de a uno):
  - **Inicio** (arranca en 10:00): al llamar `iniciarTemporizadorSocializacionIE_`, antes de arrancar el
    intervalo.
  - **Mitad alcanzada** (quedan 5:00): `segundosRestantes === 300`.
  - **Cierre** (queda 1:00): `segundosRestantes === 60`.
  - **Fin** (llega a 0:00): `segundosRestantes === 0`.
  - **Sonido**: `reproducirSonidoSocializacion_(tipo)` genera tonos cortos con Web Audio API (osciladores
    seno, volumen bajo ~0.06, sin archivos de audio que alojar en Drive) — un pitido para inicio, dos para
    mitad, tres para cierre, uno largo y grave para fin. Envuelto en `try/catch` silencioso: si el navegador
    bloquea audio sin interacción previa del usuario (política común de autoplay), el temporizador sigue
    funcionando con normalidad, solo sin sonido.
  - **Visual**: `mostrarAvisoMomentoSocializacion_(texto)` muestra un mensaje (`#avisoMomentoSocializacion`,
    arriba del reloj) con un destello CSS (`@keyframes destelloSocializacion`, un halo verde que se desvanece
    en ~1.1s) — se reinicia la animación en cada llamada (forzando reflow) para que dispare de nuevo aunque el
    mensaje sea igual al anterior.
  - **"+2 minutos" también avisa**: adaptado del pedido "Añadir dos minutos a la presentación de (IE
    nombre)" — el botón ya existía desde el Lote 39; se le agregó el mismo aviso visual nombrando la IE, para
    que quede claro a quién se le extendió el tiempo.
- Verificado: bloques `<script>` de `Index.html`/`JS.html` (mismo falso positivo ya conocido en Index, limpio
  en JS); IDs sin duplicar, balance de etiquetas OK, balance de llaves `{`/`}` de `CSS.html` OK (299/299). No
  probado en vivo desde un navegador — en particular, no se verificó el sonido real en ningún navegador (la
  política de autoplay varía por navegador y por si hubo interacción previa del usuario en la página).

## 4.60 Cuadragésimo octavo lote: PDF de aportes relevantes + reactivar el consolidado de socialización en Sesión 1

Dos pedidos sueltos del usuario sobre Consolidado de Socialización (Sesión 1), ambos aprovechando los datos
reales que ya captura la Sesión de socialización (pantalla anterior) desde el Lote 39.

- **"Construir un PDF llamado 'Aportes relevantes del Grupo (número)' con todos los aportes por IE guardados
  en la sesión anterior"**: nueva función `generarPdfAportesRelevantesSocializacion()` (Socializacion.gs),
  mismo patrón de generación de documentos que `generarInformeGrupo()` (Informes.gs: crea un
  `DocumentApp`, un título por IE (`titulo1_`) con su texto (`parrafo_`), lo exporta a PDF y lo guarda en la
  misma carpeta de Drive del grupo (`03_INFORMES_GRUPALES/GRUPO N`, vía `asegurarCarpetaGrupo_()`) — es un
  documento aparte del informe oficial, no lo reemplaza ni lo toca. Se registra en una hoja nueva
  (`AportesRelevantesSocializacion`, `ID_GRUPO`→`DOC_ID`/`PDF_ID`/`URL`/`FECHA`) para que **regenerar no
  acumule copias**: si ya existía una versión, se manda a la papelera antes de crear la nueva. Si el grupo
  todavía no anotó ningún dato relevante, el botón avisa en vez de generar un PDF vacío. En Index.html
  (dentro de la tarjeta reactivada, ver abajo): botón "📄 Generar PDF de aportes relevantes" + enlace "👁 Ver
  PDF" que aparece solo (sin regenerar) si el grupo entra a Sesión 1 y ya había uno generado antes.
- **"En las preguntas de Consolidado de Socialización, relacionar los aportes más relevantes de la
  socialización que se puedan vincular al tipo de pregunta"**: se reactivó `tarjetaSocializacionPreparacion`
  (la tarjeta "🗣️ Aportes de preparación por institución" que se había ocultado en el Lote 42 porque su fuente
  de datos original —la vieja "Preparación", item 9— ya no existía) y se recableó a
  `rpcObtenerSocializacionGrupo`, la fuente real de datos de hoy. Queda como "🗣️ Aportes relevantes de la
  socialización": un `<details>` colapsable por IE con lo que el grupo anotó durante su temporizador,
  colocado justo antes de las 3 preguntas temáticas — para que el grupo lea/relacione cada aporte con el tema
  que corresponda mientras responde. **Decisión de alcance**: clasificar automáticamente cada aporte por tema
  (FEM2025/Currículo/Gobierno) no es viable de forma confiable sin NLP real — el texto libre que se anota en
  vivo durante una presentación oral no trae una etiqueta de tema; en cambio, mostrar los aportes completos
  justo antes de las preguntas, con el mecanismo de "seleccionar texto → copiar automáticamente" que ya
  existía (reutilizado sin cambios, sigue funcionando porque el contenedor `#listaSocializacionPreparacion` es
  el mismo), es lo que permite al grupo "vincularlos al tipo de pregunta" — de forma manual y deliberada, no
  automática.
  - `cargarSocializacionPreparacion()` (JS.html) y `renderSocializacionPreparacion()` (Components.html)
    reescritas para el nuevo shape de datos (`{idIE, institucion, datosRelevantes}`, un solo texto libre por
    IE) en vez del shape viejo (`{idIE, institucion, responsable, secciones:[{titulo,texto,clave}]}` de
    Preparación, con varias secciones tituladas por IE).
  - `abrirPantallaCompletaTexto()` simplificada a recibir solo `idIE` (antes `idIE|claveSeccion`, ya no hace
    falta distinguir secciones).
  - Vuelve a llamarse `cargarSocializacionPreparacion()` al entrar a `pantallaSesion1` (se había quitado la
    llamada en el Lote 42 por la misma razón por la que se ocultó la tarjeta).
- Verificado: `node --check` sobre `Socializacion.gs`/`Code.gs` y los bloques `<script>` de
  `Index.html`/`JS.html`/`Components.html` (mismo falso positivo ya conocido en Index, limpio en JS y
  Components); IDs sin duplicar y balance de etiquetas OK en los tres HTML. No probado en vivo desde un
  navegador — en particular, no se generó un PDF real ni se verificó su contenido/formato en Drive.

## 4.61 Fix: "Generar PDF de aportes relevantes" se quedaba trabado

El usuario reportó que el botón nuevo del Lote 48 se quedaba en "Generando PDF…" sin resolver. Sin acceso a
los logs de ejecución reales (este proyecto no tiene un proyecto de GCP enlazado para `clasp logs`), el
diagnóstico fue por revisión de código, comparando contra el único patrón equivalente que sí lleva tiempo en
producción (`generarInformeGrupo`, Informes.gs) para encontrar qué era distinto en el código nuevo:

- **Se reescribió `generarPdfAportesRelevantesSocializacion()`** (Socializacion.gs) para sacar la parte lenta
  (crear el `DocumentApp`, escribir el cuerpo, exportarlo a PDF — llamadas reales a Docs/Drive que pueden
  tardar varios segundos) de dentro de `conLock_()`. Antes, todo el proceso —incluida esa parte lenta— corría
  bajo `LockService.getScriptLock()`, que es un lock de **todo el proyecto**, compartido por cualquier grupo
  que esté guardando algo al mismo tiempo (autoguardado de Sesión 1/2, Socialización, valoración, etc.) —
  mantenerlo tomado varios segundos podía encadenar esperas y sentirse como que la app entera se trababa, no
  solo el botón. Ahora el lock solo envuelve la escritura final (rápida) en la hoja de seguimiento
  `AportesRelevantesSocializacion` y el borrado del PDF/Doc anterior — timeout bajado de 30s a 15s, acorde a
  lo que realmente protege. `generarInformeGrupo`/`generarInformeCompletoGrupo` tienen el mismo patrón de raíz
  (todo el proceso lento bajo el mismo lock global) — no se tocó en este lote por no ser lo reportado, pero
  queda como riesgo conocido si vuelve a pasar algo parecido con "Generar informe".
- **`mostrarCargaAccion_` (JS.html)**: el nuevo botón usaba el ⏳ estático genérico en vez de la barra de
  progreso simulada que ya existía para "informe"/"invitado" (acciones igual de lentas) — se agregó
  `"aportesRelevantesSocializacion"` a esa lista, para que se vea que algo avanza en vez de una espera con un
  ícono inmóvil, que es fácil de leer como "se quedó pegado" aunque el servidor sí estuviera trabajando.
- Verificado: `node --check` sobre `Socializacion.gs` y el bloque `<script>` de `JS.html` (limpio). No se pudo
  confirmar en vivo si esto resuelve el reporte original del usuario — no hay logs de ejecución disponibles
  para este proyecto (`clasp logs` requiere un proyecto de GCP vinculado, no configurado aquí); si sigue
  pasando, hace falta vincular uno para poder ver la traza real del error.

## 4.62 Cuadragésimo noveno lote (revertido en el lote 50, ver 4.63): Conecta Educa pasa a operarlo un único administrador de la SEM

**Este lote fue completamente revertido en el lote 50 (4.63)** — se deja este apartado solo como registro
histórico de por qué existió, no como referencia de código vigente. Nada de lo descrito aquí abajo sigue en
el proyecto: `CODIGO_SUPERADMIN_CONECTAEDUCA`, `generarCodigoSuperadminConectaEduca`,
`validarSuperadminConectaEduca`, `iniciarSesionConectaEducaComoSuperadmin`, `rpcValidarSuperadminConectaEduca`,
`rpcIniciarSesionConectaEducaComoSuperadmin`, `pantallaAccesoSuperadminConectaEduca`,
`pantallaSeleccionGrupoConectaEducaSuperadmin` y `renderListaGruposSuperadminConectaEduca` ya no existen.

Pedido del usuario en ese momento: "para conecta Educa no pidas el mismo codigo, Genera 1 codigo de
superadmin que solo manejen un admin dela SEM" — un único administrador de la Secretaría de Educación de
Neiva entraba con un código maestro, elegía uno de los 6 grupos comunales de una lista y diligenciaba
Conecta Educa en su nombre. Un pedido posterior del usuario ("Elimina el sistema de desarrollo de conecta
Educa por grupo... Genera 1 codigo de superadmin...") reemplazó esta idea por completo: ver 4.63.

## 4.63 Quincuagésimo lote: Conversatorio — Conecta Educa deja de depender de los 6 grupos comunales

Pedido del usuario, en dos mensajes: primero compartió un documento de Google Sheets con la agrupación de
las Instituciones Educativas de Neiva en 7 grupos por técnica/articulación SENA (no 6 grupos comunales), y
pidió "Elimina el sistema de desarrollo de coencta educa por grupo... llames a la primera parte
'conversatorio'... Caracteriza la sesión con pantalla donde se vea la información extraída desde Hoja
grupos conecta educa... selección de grupo y direccionar a grupo que pertenezca... Segunda pantalla
muestra datos desde Resolución Vs IE editable... preguntas debajo de cada técnica... agrupa a las IE en
[los grupos] desde hoja Matriz IE-Técnica". Se aclaró con 3 rondas de preguntas al usuario (número real de
grupos, alcance del reemplazo, y quién entra) — decisiones finales: **usar los 7 grupos reales del
documento** (no 6); **Conversatorio es la nueva primera parte, independiente**, sin conectarse con el resto
de Conecta Educa (consentimiento, Sesión 2, cierre, que quedan intactos y sin uso por ahora); y **cada
Institución Educativa entra por su cuenta**, con su propio código — no un administrador único ni el código
de un grupo comunal.

**Paso 1 — revertir el lote 49 (4.62) por completo.** `git checkout <commit-previo-al-49> -- Access.gs
Code.gs Components.html Config.gs Index.html` (esos 4 archivos eran puramente aditivos, sin mezclar con
otros cambios — se pudo revertir de forma limpia); `JS.html` se revirtió también al estado previo al 49
completo (el refactor de `aplicarSesionGrupoValidada_`/`estado.accionForzarSesion_` de ese lote solo tenía
sentido para compartir lógica con el superadmin que ya no existe, así que no valía la pena conservarlo
aparte).

**Paso 2 — Conversatorio, mini-herramienta nueva e independiente.** Fuente de datos: hoja de cálculo de
Google externa (`ID_HOJA_FUENTE_CONVERSATORIO_` en Conversatorio.gs) con 3 pestañas relevantes: "Grupos
ConectaEduca" (7 grupos: nombre, N.º de IE, técnicas incluidas, criterio de agrupación), "Matriz
IE-Técnica" (qué IE pertenece a qué grupo(s) — varias IE pertenecen a más de uno) y "Resolución Vs IE"
(técnica SENA, programa aprobado, resolución/decreto, por IE — incluye instituciones que NO son parte de
los 7 grupos de Conecta Educa, que se descartan al importar).

- **Conversatorio.gs** (archivo nuevo):
  - `importarDatosConversatorio()` — se ejecuta a mano desde el editor cada vez que la Secretaría actualice
    el documento fuente (mismo patrón que `generarAccesosGrupo`). Abre el documento fuente directamente con
    `SpreadsheetApp.openById` (misma cuenta de Google que el proyecto — evita transcribir a mano ~150 filas
    y el riesgo de errores de transcripción) y puebla 3 hojas propias del proyecto:
    `ConversatorioGrupos`, `ConversatorioMatrizIETecnica` (se reescriben por completo en cada import, sin
    datos de usuario) y `ConversatorioResolucionIE` (solo AGREGA filas institución+técnica que no existan
    todavía — nunca pisa ediciones o respuestas ya guardadas por una IE). También genera (idempotente,
    nunca regenera) un código de acceso `CE-XXXXX` por cada institución que aparece en la Matriz
    IE-Técnica, en la hoja `AccesosIEConversatorio`.
  - Nombres de institución inconsistentes entre pestañas del documento fuente (p. ej. "NACIONAL SANTA
    LIBRADA" en Resolución Vs IE vs. "IE SANTA LIBRADA" en Matriz IE-Técnica) se resuelven con
    `normalizarNombreIEConversatorio_` (mayúsculas, sin tildes, sin prefijo "IE") + un pequeño mapa de
    alias (`ALIAS_IE_CONVERSATORIO_`) para los 2-3 casos que ni así calzan.
  - `validarAccesoIEConversatorio(codigo)` — valida el código y devuelve el catálogo completo de los 7
    grupos, marcando (`perteneceIE`) a cuáles pertenece esa institución.
  - `elegirGrupoConversatorio(codigo, idGrupo)` — guarda qué grupo eligió trabajar la IE
    (`GRUPO_ELEGIDO` en `AccesosIEConversatorio`); valida que la IE de verdad pertenezca a ese grupo.
  - `obtenerTecnicasConversatorio(codigo)` / `guardarCampoTecnicaConversatorio(codigo, idFila, campo,
    valor)` — lectura y autoguardado por campo (técnica/programa/resolución o una de las 2 preguntas de
    proyección 2027) de las filas de `ConversatorioResolucionIE` de esa institución.
  - `finalizarConversatorio(codigo)` — marca `ESTADO = COMPLETADO`.
  - Deliberadamente **sin** el mecanismo de cupos-por-dispositivo/heartbeat/"sesión ya abierta" que sí
    tiene el resto de la app (Session.gs) — el Conversatorio lo suele llenar una sola persona de contacto
    por institución, así que se mantuvo simple a propósito.
- **Code.gs**: `rpcValidarAccesoIEConversatorio`, `rpcElegirGrupoConversatorio`,
  `rpcObtenerTecnicasConversatorio`, `rpcGuardarCampoTecnicaConversatorio`, `rpcFinalizarConversatorio`.
- **Index.html**: 4 pantallas nuevas insertadas después de `pantallaAcceso`: `pantallaAccesoConversatorio`
  (código de la institución), `pantallaCaracterizacionConversatorio` (catálogo de los 7 grupos, resalta a
  cuáles pertenece la IE, botón "Elegir este grupo" solo en los propios), `pantallaResolucionConversatorio`
  (técnicas editables + las 2 preguntas debajo de cada una) y `pantallaCierreConversatorio`. La tarjeta
  "Conecta Educa" de `pantallaEleccionSeccion` ya no usa `data-elegir-seccion` (ese mecanismo sigue
  existiendo solo para "ENCUENTRO"): tiene `id="btnIrConversatorio"` y lleva directo a
  `pantallaAccesoConversatorio`. Texto de consentimiento de grupo actualizado para explicar que Conecta
  Educa ahora lo diligencia cada IE por su cuenta.
- **Components.html**: `renderGruposCaracterizacionConversatorio(grupos, grupoElegido)` y
  `renderTecnicasResolucionConversatorio(tecnicas)`.
- **CSS.html**: `.tarjeta-grupo-ie`/`.tarjeta-grupo-ajena`/`.tarjeta-grupo-elegida` (variantes de
  `.tarjeta-seccion` reutilizada como tarjeta no-interactiva, `div` en vez de `button`).
- **JS.html**: `estado.codigoConversatorio`/`institucionConversatorio`/`grupoElegidoConversatorio`;
  `pantallaAccesoConversatorio`/`pantallaCaracterizacionConversatorio`/`pantallaResolucionConversatorio`/
  `pantallaCierreConversatorio` agregadas a `ORDEN_PANTALLAS` (justo después de `pantallaAcceso`, antes de
  `pantallaParticipacion` — así el guard de "método de asistencia obligatorio" en `cambiarPantalla` nunca
  las alcanza) y a `PANTALLAS_NO_RESUMIBLES_`; autoguardado por campo con el mismo patrón "change" que el
  resto de la app (p. ej. el rector en Confirmación de caracterización).

**Decisión deliberada de alcance**: el resto de Conecta Educa (`pantallaConsentimientoConectaEduca`,
`pantallaSesion2`, `pantallaConfirmacionCaracterizacion`/`pantallaRevisionCierre` cuando
`estado.seccion === "CONECTAEDUCA"`) **no se tocó ni se conectó con el Conversatorio** — sigue en el código
tal cual estaba, pero ahora inalcanzable desde la interfaz (nada pone `estado.seccion = "CONECTAEDUCA"` ya
que la tarjeta de la sección ya no usa `data-elegir-seccion`). Se deja así a propósito (decisión explícita
del usuario) en vez de borrarlo, por si se vuelve a necesitar.

**Importante para la Secretaría**: ejecutar `importarDatosConversatorio()` una vez desde el editor de Apps
Script (con el documento fuente ya accesible a la misma cuenta de Google del proyecto) antes de que
cualquier institución pueda usar el Conversatorio — después de eso, cada institución aparece en el listado
de acceso. Si el documento fuente cambia más adelante, se puede volver a ejecutar sin perder respuestas ya
guardadas. **Nota (ver 4.64): el acceso por código de este apartado fue reemplazado por un listado — cada
institución se elige a sí misma, sin código.**

Verificado: `node --check` sobre Conversatorio.gs y Code.gs; extracción y `node --check` de los bloques
`<script>` de Index.html, JS.html y Components.html (limpios, aparte de los dos falsos positivos
permanentes ya conocidos); sin IDs duplicados ni etiquetas sin cerrar en Index.html/Components.html. No fue
posible probar en vivo el `SpreadsheetApp.openById` contra el documento fuente real (sin acceso a
`clasp run` ni `clasp logs` en este entorno) — si al ejecutar `importarDatosConversatorio()` la cuenta del
proyecto no tuviera acceso a ese documento, fallaría con un error de permisos de Google, no silenciosamente.

## 4.64 Quincuagésimo primer lote: Conversatorio sin códigos por IE + pregunta de apertura 2027 con Sí/No

Pedido del usuario, dos ajustes puntuales al lote anterior (4.63): "omite que se generen codigos por IE y
mas bien que ellos mismos escojan la IE de un listado" y "opciones si o no desplegable en ¿Se aperturará
esta articulación para grado 10° en el 2027?".

- **Conversatorio.gs**: se quitó `CODIGO` de `cabecerasConversatorioAccesos_()` — `AccesosIEConversatorio`
  ya no genera ni guarda un código por institución, solo una fila (`INSTITUCION_EDUCATIVA`, `ESTADO`,
  `GRUPO_ELEGIDO`, `ULTIMA_ACTIVIDAD`) por cada una, creada en `importarDatosConversatorio()`. Nueva función
  `obtenerInstitucionesConversatorio()` — devuelve los nombres (ordenados alfabéticamente, sin las
  bloqueadas) para el listado de elección. `buscarAccesoConversatorioPorCodigo_` se reemplazó por
  `buscarAccesoConversatorioPorInstitucion_` (mismo `normalizarNombreIEConversatorio_` para la
  comparación); `validarAccesoIEConversatorio(codigo)` se renombró a
  `seleccionarInstitucionConversatorio(institucion)`; `elegirGrupoConversatorio`,
  `obtenerTecnicasConversatorio`, `guardarCampoTecnicaConversatorio` y `finalizarConversatorio` reciben
  ahora `institucion` en vez de `codigo` como primer parámetro.
- **Code.gs**: `rpcObtenerInstitucionesConversatorio` (nuevo) y `rpcSeleccionarInstitucionConversatorio`
  (reemplaza a `rpcValidarAccesoIEConversatorio`); el resto de RPCs de Conversatorio actualizados al nuevo
  nombre de parámetro.
- **Index.html**: `pantallaAccesoConversatorio` ya no tiene el campo de código ni el botón "Continuar" —
  ahora es directamente el contenedor `listaInstitucionesAccesoConversatorio`, poblado al entrar a la
  pantalla. Textos de la tarjeta "Conecta Educa" (elección de sección) y del consentimiento de grupo
  actualizados para decir "se elige de una lista" en vez de "con su propio código".
- **Components.html**: `renderListaInstitucionesAccesoConversatorio(instituciones)` (mismo patrón de fila +
  botón que `renderGruposCaracterizacionConversatorio`). En `renderTecnicasResolucionConversatorio`, el
  campo `preguntaApertura2027` pasó de `<input type="text">` a `<select>` con 3 opciones ("— Sin
  responder —", "Sí", "No") — el autoguardado en "change" (JS.html) no necesitó cambios, ya funciona igual
  para `<select>` que para `<input>`. La segunda pregunta (`preguntaNuevaArticulacion2027`) sigue como
  texto libre — el pedido del usuario solo mencionó la primera.
- **JS.html**: `estado.codigoConversatorio` eliminado — `estado.institucionConversatorio` (el nombre, no un
  código) es ahora la única clave que se manda a los RPC. `btnIrConversatorio` ahora carga el listado
  (`rpcObtenerInstitucionesConversatorio`) antes de mostrar `pantallaAccesoConversatorio`, en vez de solo
  cambiar de pantalla. Nuevo delegado `[data-elegir-institucion-conversatorio]` junto al de
  `[data-elegir-grupo-conversatorio]` ya existente.

Verificado: `node --check` sobre Conversatorio.gs y Code.gs; extracción y `node --check` de los bloques
`<script>` (limpios, mismos 2 falsos positivos permanentes); sin IDs duplicados ni etiquetas (incluido
`<select>`) sin cerrar en Index.html/Components.html.

## 4.65 Quincuagésimo segundo lote: panel de administrador, reordenar el recorrido, solo lectura + lápiz, registro de cambios

Pedido del usuario, varios ajustes en un solo mensaje: "haz un login de superadmin con el codigo que ya me
habias dado antes. ¿Se planea incluir una ueva articulación para el año 2027? igual con desplegable. Si
selecciona si. Deplegar listado de tecnicas actuales. y poner arriba casilla otro. la selección de grupo al
que desea tener el conversatorio se dará despues de repsonder las preguntas de cada [técnica] mostrar
unicamente las IE que pertenecen a cada grupo, una vez se haya seleccionado la IE. Agregar poner todo en
solo lectura, y si hay necesidad de cambiar algo, dar click en el lápíz, luego boton guardar. Esos cambios
deben mencionarse explicitamente en el informe...".

- **Panel de superadministrador** (revive el código `CODIGO_SUPERADMIN_CONECTAEDUCA` de un lote anterior —
  Config.gs, Conversatorio.gs: `generarCodigoSuperadminConversatorio()`/`validarSuperadminConversatorio()`):
  desde la tarjeta "Conversatorio" hay un enlace "Entrar con código de administrador"
  (`pantallaAccesoSuperadminConversatorio`) que lleva a `pantallaPanelSuperadminConversatorio` — lista de
  todas las instituciones con su estado, botón "Entrar" por cada una (reutiliza exactamente el mismo camino
  que el autoselección de una IE — `rpcSeleccionarInstitucionConversatorio` — el administrador solo llega
  ahí por una puerta distinta) y el botón "Generar informe consolidado".
- **Reordenar el recorrido** (spec: "la selección de grupo... se dará despues de repsonder las preguntas"):
  ahora es institución → técnicas y las 2 preguntas (`pantallaResolucionConversatorio`) → grupo de
  articulación técnica (`pantallaCaracterizacionConversatorio`, ahora AL FINAL) → cierre. Elegir un grupo ya
  finaliza el Conversatorio de una vez (antes eran dos pasos separados).
- **Filtrar a solo los grupos propios** (spec: "mostrar unicamente las IE que pertenecen a cada grupo, una
  vez se haya seleccionado la IE"): `renderGruposCaracterizacionConversatorio` ya no muestra los 7 grupos
  con los ajenos atenuados — solo lista los que `perteneceIE`.
- **Pregunta 2 con desplegable Sí/No + catálogo de técnicas**: "¿Se planea incluir una nueva articulación
  para el año 2027?" (antes texto libre) ahora es un `<select>` Sí/No, igual que la primera pregunta. Si
  responde "Sí", se despliega un subformulario (`subform-nueva-articulacion-conversatorio`) con un campo
  "Otro" arriba y, debajo, un checklist de las técnicas únicas de los 7 grupos
  (`obtenerCatalogoTecnicasConversatorio()`, Conversatorio.gs — RPC `rpcObtenerCatalogoTecnicasConversatorio`,
  se pide una sola vez y se cachea en `estado.catalogoTecnicasConversatorio`). Al guardar, "Otro" + las
  técnicas marcadas se combinan en un solo texto (`NUEVA_ARTICULACION_DETALLE`, columna nueva en
  `ConversatorioResolucionIE`) con el formato `Otro: <texto>; <técnica>; <técnica>...`. Limitación conocida:
  al reabrir el modo edición no se reconstruyen las casillas marcadas a partir de ese texto guardado (el
  checklist vuelve a empezar sin marcar) — el valor de solo-lectura sí se sigue mostrando correctamente;
  solo el estado de las casillas en el propio formulario de edición no persiste entre visitas.
- **Solo lectura + lápiz** (spec: "poner todo en solo lectura, y si hay necesidad de cambiar algo, dar
  click en el lápiz, luego boton guardar"): cada tarjeta de técnica en `renderTecnicasResolucionConversatorio`
  arranca en `.modo-lectura-tecnica-conversatorio` (texto plano) con un botón "✏️ Editar"; al hacer clic se
  oculta la lectura y aparece `.modo-edicion-tecnica-conversatorio` (los campos editables de siempre) con un
  botón "💾 Guardar" que manda los 6 campos de la tarjeta en secuencia
  (`guardarCamposTecnicaConversatorioEnSecuencia_`, JS.html) y vuelve a cargar la tarjeta en modo lectura.
- **Registro de cambios para el informe** (spec: "la IE X cambió la respuesta X por Y"): nueva hoja
  `ConversatorioCambios` (Conversatorio.gs) — `guardarCampoTecnicaConversatorio` ahora compara el valor
  anterior contra el nuevo antes de sobrescribir; si el anterior tenía contenido y es distinto del nuevo,
  agrega una fila (institución, técnica, campo, valor anterior, valor nuevo, fecha). Nunca se registra el
  llenado inicial de un campo vacío, solo ediciones reales sobre algo ya guardado.
- **Informe consolidado**: `generarInformeConversatorio(codigoSuperadmin)` (Conversatorio.gs) — mismo patrón
  que `generarPdfAportesRelevantesSocializacion` (§4.61: la parte lenta de Docs/Drive va fuera de
  `conLock_`, solo la escritura final de seguimiento en `ConversatorioInforme` queda adentro). Un Doc/PDF
  con, por institución, cada técnica y sus 2 respuestas, seguido de un apartado final "Cambios presentados
  por las Instituciones Educativas en las respuestas" con una línea por cada fila de `ConversatorioCambios`,
  con exactamente la redacción pedida: `La IE <institución> cambió la respuesta de <campo> (técnica
  "<técnica>") de "<anterior>" a "<nueva>".` Se guarda en la carpeta `04_CONECTAEDUCA` de Drive (spec del
  proyecto) y se regenera cada vez (nunca acumula copias, mismo criterio que el resto del proyecto).
- **Code.gs**: `rpcObtenerCatalogoTecnicasConversatorio`, `rpcValidarSuperadminConversatorio`,
  `rpcGenerarInformeConversatorio`.

**Importante para la Secretaría**: si `CODIGO_SUPERADMIN_CONECTAEDUCA` ya tenía un valor guardado de un
lote anterior, ese mismo código sigue sirviendo para el panel del Conversatorio — no hace falta generar uno
nuevo. Si está vacío, ejecutar `generarCodigoSuperadminConversatorio()` desde el editor de Apps Script.

Verificado: `node --check` sobre Conversatorio.gs, Code.gs y Config.gs; extracción y `node --check` de los
bloques `<script>` de Index.html, JS.html y Components.html (limpios, aparte de los dos falsos positivos
permanentes ya conocidos); sin IDs duplicados ni etiquetas (incluidos `<select>`) sin cerrar en
Index.html/Components.html. No fue posible probar en vivo la generación del informe ni el flujo completo
del panel de administrador (sin `clasp run`/`clasp logs` en este entorno).

## 4.66 Quincuagésimo tercer lote: fix pantalla inicial, pantalla de responsable + banner, preguntas obligatorias, botón "es correcta"

Pedido del usuario: "la pantalla inicial debe ser la de escoger entre las dos sesiones, actualmente esta
iniciando en conecta educa. despues de seleccionar IE, hacer pantalla de responsable de llenar sesion de
conecta educa. en pantalla donde se edita las tecnicas debe editarse obligatoriamente las preguntas de si y
de no y guardar Boton de es correcto pra informacion de tecnica o deseamos editar esta información. En la
pantalla final de selección inidicar que puede participar del conversatorio con el grupo seleccionado. haz
un banner que diga, respetado (a), (rol seleccionado) de la IE (nombre) y (acompañante si registró)".

- **Fix: la app abría directo en la pantalla vieja de acceso de Conecta Educa en vez de
  `pantallaEleccionSeccion`.** Causa: `retomarSeccionElegida_` (JS.html) recuerda en `localStorage` la
  última sección elegida y salta directo a su pantalla de acceso — cualquier dispositivo que hubiera
  elegido "Conecta Educa" ANTES del rediseño a Conversatorio (lotes 4.63-4.65) seguía teniendo
  `"CONECTAEDUCA"` guardado, y `TEXTOS_SECCION_` (que decide qué valores son válidos para retomar) todavía
  la reconocía, aunque la tarjeta ya no usara ese mecanismo (usa `btnIrConversatorio` directo). Se quitó
  `CONECTAEDUCA` de `TEXTOS_SECCION_` — ahora solo `ENCUENTRO` es "recordable"; un valor viejo de
  `"CONECTAEDUCA"` en `localStorage` ya no coincide con nada y la app cae de vuelta a
  `pantallaEleccionSeccion`, como debe ser.
- **Nueva pantalla `pantallaResponsableConversatorio`** (spec: "despues de seleccionar IE, hacer pantalla
  de responsable de llenar sesion de conecta educa"): aparece justo después de elegir la institución (tanto
  autoselección como entrada de administrador) y antes de las técnicas — pide nombre completo, rol (select:
  Rector(a), Coordinador(a) académico(a), Docente, Enlace/articulador(a) SENA, Orientador(a), Otro) y un
  acompañante opcional. Se guarda en 3 columnas nuevas de `AccesosIEConversatorio`
  (`RESPONSABLE_NOMBRE`/`RESPONSABLE_ROL`/`ACOMPANANTE`) vía `guardarResponsableConversatorio` (Conversatorio.gs) /
  `rpcGuardarResponsableConversatorio` (Code.gs) — se pre-llena si ya había datos guardados (reingreso o
  administrador revisando).
- **Banner de bienvenida** (spec, literal: "respetado (a), (rol seleccionado) de la IE (nombre) y
  (acompañante si registró)"): `textoBannerResponsableConversatorio_()` (JS.html) arma el texto exacto
  "Respetado(a) `<rol>` de la IE `<institución>` y `<acompañante>`." (el "y `<acompañante>`" solo aparece si
  se registró uno) y se pinta en `#bannerResponsableConversatorio` (pantalla de técnicas) y
  `#bannerCaracterizacionConversatorio` (pantalla de grupo).
- **Preguntas Sí/No obligatorias antes de guardar** (spec: "debe editarse obligatoriamente las preguntas de
  si y de no y guardar"): el clic en "💾 Guardar" de una tarjeta de técnica ahora valida primero que ambos
  `<select>` (apertura 2027 / nueva articulación 2027) tengan una respuesta — si falta alguna, muestra un
  error y no guarda nada.
- **Botón "✅ Es correcta esta información"** (spec: "Boton de es correcto pra informacion de tecnica o
  deseamos editar esta información"): junto al lápiz ✏️, en modo lectura, un botón para confirmar los datos
  tal como están sin necesidad de abrir edición — nuevas columnas `CONFIRMADO_POR_IE`/`FECHA_CONFIRMACION`
  en `ConversatorioResolucionIE` (`confirmarTecnicaConversatorio`/`rpcConfirmarTecnicaConversatorio`); una
  vez confirmada, el botón se reemplaza por la insignia "✅ Información confirmada".
- **Aviso de participación con el grupo elegido** (spec: "indicar que puede participar del conversatorio
  con el grupo seleccionado"): el texto de introducción de `pantallaCaracterizacionConversatorio` ahora dice
  explícitamente que la institución podrá participar del Conversatorio con el grupo que elija, y la
  pantalla de cierre repite el nombre concreto del grupo elegido ("Su institución puede participar del
  Conversatorio con el grupo seleccionado: G0X — Nombre del grupo.").

Verificado: `node --check` sobre Conversatorio.gs y Code.gs; extracción y `node --check` de los bloques
`<script>` de Index.html, JS.html y Components.html (limpios, mismos 2 falsos positivos permanentes); sin
IDs duplicados ni etiquetas sin cerrar en Index.html/Components.html.

## 4.67 Quincuagésimo cuarto lote: limpiar mención cruzada a Conecta Educa, rol de la SEM en los catálogos, matriz de estamento al final en Confirmación de caracterización

Pedido del usuario: "elimina esto de pantalla de bienvenida de pantalla de foro no hace falta terminar aquí
para poder entrar a Conecta Educa. agregar funcionario Secretaría de Educación Municipal en el listado de
roles de responsables de foro y conecta educa. en responsable de envio, pon también Secretaría de educación
en foro y en conecta educa. Cantidad de asistentes por estamento e institución debe aparecer al final y no
al principio".

- **Index.html, `pantallaMetodologia` ("Ruta del Encuentro")**: se quitó la mención a Conecta Educa
  ("...con su propia puerta de acceso (mismo código del grupo) — no hace falta terminar aquí para poder
  entrar a Conecta Educa"), que además de ser lo pedido por el usuario ya estaba desactualizada — Conecta
  Educa (Conversatorio) ya no se entra "con el mismo código del grupo" desde los lotes 4.63-4.66.
- **Responsables.gs, `ROLES_FORO_`**: se agregó `"Funcionario Secretaría de Educación Municipal"` al
  catálogo — este mismo arreglo alimenta tanto el "Rol en el foro" de Responsable de envío/asistentes de
  envío (Participación, pantalla común a Encuentro y a la vieja Conecta Educa por grupo) como el rol en
  Preparación IE, así que un solo cambio cubre "responsables de foro y conecta educa" y "responsable de
  envío... en foro y en conecta educa" a la vez.
- **Index.html, `pantallaResponsableConversatorio`** (el "Conecta Educa" vigente hoy, el Conversatorio):
  se agregó la misma opción `"Funcionario Secretaría de Educación Municipal"` al select "Rol en la
  institución" — con el panel de superadministrador (4.65), quien diligencia en nombre de una IE puede ser
  justamente un funcionario de la Secretaría.
- **Index.html, `pantallaConfirmacionCaracterizacion`**: el bloque "Participación por estamento e
  institución" (matriz + resumen) estaba justo después de la ficha de caracterización, casi al principio de
  la pantalla — se movió después de la fotografía del grupo, justo antes del aviso de carga y los botones
  finales (spec del usuario: "debe aparecer al final y no al principio"). En `pantallaParticipacion` esa
  misma matriz ya estaba cerca del final (justo antes de "Continuar"), así que no hizo falta tocarla ahí.

Verificado: `node --check` sobre Responsables.gs; extracción y `node --check` del bloque `<script>` de
Index.html (limpio, mismo falso positivo permanente ya conocido); sin IDs duplicados ni etiquetas sin
cerrar.

## 4.68 Quincuagésimo quinto lote: separar las preguntas Sí/No en su propia casilla, botón Atrás en Grupo de articulación técnica

Pedido del usuario, sobre la tarjeta de técnica del Conversatorio: "¿Se aperturará esta articulación para
grado 10° en el 2027?: Sin responder. ¿Se planea incluir una nueva articulación para el año 2027?: Sin
responder. debe aoarecer aparte para seleccionar si o no no debe aparecer dentro de la casilla" — y, sobre
la pantalla de elección de grupo: "en pantalla de Grupo de articulación técnica debe haber botón atrás".

- **Components.html, `renderTecnicasResolucionConversatorio`**: la tarjeta de cada técnica ahora tiene DOS
  casillas separadas visualmente (antes todo estaba mezclado en una sola): `.caja-datos-tecnica-conversatorio`
  (técnica/programa/resolución) y, aparte, `.caja-preguntas-tecnica-conversatorio` (las 2 preguntas Sí/No +
  el detalle de nueva articulación) — con borde y fondo propios (verde claro) para que se note que es una
  sección distinta. Esto aplica tanto en modo lectura como en modo edición.
- **Modo lectura de las preguntas Sí/No**: nueva función `_filaPreguntaSiNoLecturaConversatorio_` — en vez
  de texto plano "pregunta: Sin responder." pegado en un párrafo, la respuesta se muestra como una insignia
  de color (`.badge-si` verde, `.badge-no` rojo, `.badge-sin-responder` amarillo), separada visualmente de
  la pregunta.
- **Index.html, `pantallaCaracterizacionConversatorio`**: el botón que ya llevaba de vuelta a las técnicas
  ("← Volver a las técnicas") se renombró a "← Atrás (volver a las técnicas)" para que se reconozca
  claramente como el botón Atrás de la pantalla.
- **CSS.html**: estilos nuevos `.caja-datos-tecnica-conversatorio`, `.caja-preguntas-tecnica-conversatorio`,
  `.fila-pregunta-si-no`, `.badge-si`/`.badge-no`/`.badge-sin-responder`.

Verificado: extracción y `node --check` de los bloques `<script>` de Index.html (limpio, mismo falso
positivo permanente) y Components.html (limpio); sin IDs duplicados ni etiquetas sin cerrar.

## 4.69 Quincuagésimo sexto lote: pantalla de bienvenida siempre con las dos sesiones, varios acompañantes con rol

Dos pedidos del usuario: "pantalla de bienvenida debe mostar las dos sesiones para seleccionar" y
"Acompañantes (opcional) debe permitir también agregar más acompañantes y seleccionar rol".

- **Fix del salto automático a "Encuentro" (JS.html)**: `retomarSeccionElegida_`, una función que existía
  desde antes de la pivote a Conversatorio, saltaba automáticamente `pantallaEleccionSeccion` e iba directo
  a `pantallaAcceso` si el dispositivo ya tenía una sección guardada en `localStorage`
  (`fec_seccion_elegida`) — cualquier dispositivo que ya hubiera entrado antes al Encuentro dejaba de ver
  la pantalla de bienvenida con las dos tarjetas. Se eliminó esa función por completo: la pantalla de
  bienvenida ahora **siempre** se muestra primero, con las dos sesiones para elegir. El valor guardado en
  `localStorage` se sigue usando (solo para rotular el título/logo de `pantallaAcceso` una vez elegida la
  sección), pero ya no decide saltarse la pantalla de bienvenida.
- **Acompañantes múltiples con rol (Index.html, Components.html, JS.html, CSS.html)**: el campo único de
  texto libre "Acompañante (opcional)" de la pantalla de responsable del Conversatorio se reemplazó por una
  lista dinámica de filas — cada fila tiene un campo de nombre y un desplegable de rol (mismo catálogo que
  el rol del responsable: Rector(a)/Coordinador(a) académico(a)/Docente/Enlace-articulador(a)
  SENA/Orientador(a)/Funcionario Secretaría de Educación Municipal/Otro), con un botón "+ Agregar
  acompañante" y un "✕ Quitar" por fila.
  - `Components.html`: nueva `renderAcompanantesResponsableConversatorio(filas)` — repinta las filas desde
    cero cada vez (JS.html lee los valores actuales del DOM antes de agregar/quitar una fila, para no
    perder lo ya escrito).
  - `JS.html`: `leerFilasAcompanantesConversatorioDelDom_()` lee las filas actuales; el esquema de la hoja
    **no cambió** — la columna `ACOMPANANTE` de `ConversatorioAccesos` sigue siendo un solo texto, ahora
    compuesto como "Nombre (Rol), Nombre (Rol)" cuando hay varios (`formatearAcompanantesConversatorio_`
    antes de guardar). Al reabrir la pantalla, `parsearAcompanantesConversatorio_` reconstruye las filas a
    partir de ese texto guardado (un acompañante guardado antes de este lote, sin rol entre paréntesis, se
    recupera igual como una fila con el nombre y el rol vacío). El banner de bienvenida
    (`textoBannerResponsableConversatorio_`) no cambió — sigue mostrando el mismo texto combinado tal cual
    quedó guardado.
  - `CSS.html`: `.fila-acompanante-conversatorio` (layout en fila, responsive).

Verificado: `node --check` de los 4 `.gs` tocados (ninguno en este lote) y extracción + `node --check` de
los bloques `<script>` de Index.html, Components.html, JS.html y CSS.html (limpio, mismo falso positivo
permanente de `TOKEN_ACCESO`/`ID_GRUPO`); sin IDs duplicados nuevos ni etiquetas sin cerrar en Index.html.

## 4.70 Quincuagésimo séptimo lote (parte 1 de un pedido grande): rename FEI, agenda Jornada Tarde, asistencia solo al final, Participación al final del recorrido

Primeros 4 de 8 cambios pedidos por el usuario en un solo mensaje grande (el resto — respuestas reales en
Socialización, acceso de grupo sin código, IE presentes, fix de valoración — sigue en lotes siguientes,
ver `## 4.71` en adelante). Antes de tocar nada se investigó el código actual con un subagente de solo
lectura y se confirmaron con el usuario los puntos ambiguos/riesgosos (ver resumen de la conversación).

- **Rename "Foro Educativo Institucional" → "VOCES QUE CONSTRUYEN TERRITORIO"** (Index.html, 7
  ocurrencias): confirmado con el doc real de Drive "Informe de Síntesis Grupal" que el lema oficial de
  FEM2026 es "Escuela Viva: Voces que construyen territorio" — se aplicó el reemplazo literal en
  mayúsculas en todo el texto orientado al usuario (no se tocaron comentarios de código ni títulos de
  recursos de Drive en `Recursos.gs`, que describen documentos externos reales con ese nombre).
- **Consentimiento de grupo → Agenda de la Jornada Tarde** (Index.html, JS.html): se reemplazó por
  completo `pantallaConsentimientoGrupo` — el checkbox de consentimiento obligatorio y la
  descarga/copia/envío del listado de asistencia se eliminaron de esa pantalla (decisión explícita del
  usuario) — por una tabla `table.tabla-simple` con la agenda 2:00–5:00 p. m. de ConectaEduca (Apertura,
  Cámara de Comercio, Educación Superior, un rector, SEM, intensificación media académica, diálogo libre)
  que el usuario dio literal. El botón pasó de "Aceptar y continuar" a "Continuar" (ya no hay nada que
  aceptar). `rpcGuardarConsentimientoGrupo` se sigue llamando al continuar (sin condición), solo para que
  "retomar donde se quedó" siga funcionando igual.
- **Asistencia en PDF: solo al final** (Index.html, JS.html): se quitó de Participación el bloque completo
  de "Método de asistencia del grupo" (descarga del formato, subir PDF, declarar cantidad) — spec del
  usuario: "Remove asistencia PDF at the beginning, just keep it at the end". Ese mismo bloque, con los
  MISMOS IDs (`campoArchivoListado`, `btnSubirListado`, `btnListadoMasTarde`, `panelCantidadListado`,
  `campoCantidadListado`, `btnDescargarFormatoAsistenciaParticipacion`), se reinsertó dentro de
  `pantallaRevisionCierre`, justo antes del bloque de verificación de firmantes que ya estaba oculto desde
  antes (ese bloque de verificación NO se reactivó — es una funcionalidad distinta y más grande, no pedida
  esta vez). Mantener los mismos IDs evitó tener que reescribir la lógica de subida/guardado, que ya
  estaba probada. `prepararCierre()` ahora también fija el `href` de descarga y llama a
  `cargarCantidadListado_()` (antes eso pasaba al entrar a Participación).
  - **Limpieza de referencias colgantes**: al borrar el panel del método QR de Participación (ya estaba
    oculto/deshabilitado desde el lote de "solo PDF", pero el HTML seguía ahí) varias funciones de JS.html
    tocaban esos elementos sin verificar que existieran — `mostrarPanelMetodoAsistencia`,
    `elegirMetodoAsistencia`, `cargarEnlaceAsistencia`, y dos `addEventListener` de nivel superior
    (`btnCopiarEnlaceAsistencia`, `btnPantallaCompletaQR`) — se les agregaron guardas nulas
    (`var x = document.getElementById(...); if (x) ...`) para que no lancen error al cargar la página. Se
    verificó con un script que compara cada `document.getElementById("X")` de JS.html/Components.html
    contra los `id=` reales de Index.html/Modal.html — 8 IDs ya no existen (todos del panel QR muerto),
    los 8 confirmados con guarda nula.
- **Participación se mueve al final del recorrido** (JS.html, Index.html): spec del usuario, "Screen 2:
  Remove participacion and set it to the end of the form". Antes iba justo después de la
  agenda/Consentimiento; ahora va justo después de Confirmación de caracterización y antes de Revisión y
  cierre (mismo lugar donde ya estaba Confirmación de caracterización desde un lote anterior). Cambios:
  - `ORDEN_PANTALLAS`, `PASOS_PROGRESO_ENCUENTRO_`, `PASOS_PROGRESO_CONECTAEDUCA_`: se movió la entrada de
    Participación a su nueva posición.
  - `irTrasConfirmarCaracterizacion_()` ahora va a `pantallaParticipacion` (antes iba directo a
    `pantallaRevisionCierre`); `irTrasParticipacion_()` ahora va a `pantallaRevisionCierre` (antes iba a
    `pantallaSesionSocializacion`, simplificada porque la rama de ConectaEduca ya era código muerto).
  - El botón "Continuar" de la nueva Agenda ahora va directo a `pantallaSesionSocializacion` (ya no pasa
    por Participación).
  - `pantallaSesionSocializacion` ahora tiene "Atrás" → `pantallaConsentimientoGrupo` (antes → Participación).
  - Se agregó un botón "Atrás" → `pantallaConfirmacionCaracterizacion` en Participación (antes no tenía
    ninguno, porque estaba casi al principio del recorrido).
  - `indiceMaximoOrdenPantallas_` (se fija apenas se valida el código, para habilitar la barra de progreso
    de una vez) ahora apunta a `pantallaSesionSocializacion` en vez de `pantallaParticipacion`, que ya no
    es el primer paso real.

Verificado: extracción y `node --check` de los bloques `<script>` de Index.html, JS.html, Components.html,
Modal.html y CSS.html (limpio, mismo falso positivo permanente); sin IDs duplicados; comparación
automatizada de cada referencia `getElementById` contra los IDs reales del HTML (0 referencias colgantes
sin guarda nula).

## 4.71 Quincuagésimo séptimo lote (parte 2): fix del error de valoración + casilla de IE presentes

### Fix: "Debe completar la valoración de esta sesión antes de generar el informe" apareciendo sin razón

Diagnóstico (sin logs de ejecución disponibles — `clasp logs` no funciona en este entorno, se investigó
solo por lectura de código): `guardarValoracionGrupo` (Valoracion.gs) guarda la fila con `upsertFila_`
dentro de `conLock_` pero **nunca llamaba a `SpreadsheetApp.flush()`** antes de devolver éxito al cliente.
Apps Script no garantiza que un `appendRow`/`setValue` sea visible de inmediato para OTRA ejecución del
script — y "Generar informe" es casi siempre una ejecución completamente separada de `google.script.run`,
disparada por el cliente apenas `estado.valoracionCompletada` pasa a `true`. Si el clic en "Generar
informe" llega antes de que el guardado de la valoración se haya "asentado" en la hoja, `Grupos.gs` lee
`obtenerValoracionGrupo` y no encuentra la fila todavía, aunque sí se guardó. **Este es el mismo patrón de
bug ya encontrado y corregido una vez antes** en este proyecto: `registrarParticipante` (Data.gs:261) tiene
el comentario "Sin flush(), una lectura casi inmediata... puede no ver todavía esta fila" — exactamente el
mismo mecanismo, aplicado ahora a la valoración.
- **Fix**: se agregó `SpreadsheetApp.flush()` al final de `guardarValoracionGrupo`, antes de `return { ok:
  true, ... }`.
- **Endurecimiento adicional** (no confirmado como la causa, pero cerraba una asimetría real):
  `guardarValoracionGrupo` normaliza `seccion` a mayúsculas antes de construir la CLAVE, pero
  `obtenerValoracionGrupo` no lo hacía — se agregó la misma normalización en la lectura, por si acaso
  `estado.seccion` llegara alguna vez en minúscula desde el cliente.

### Casilla "Instituciones educativas presentes"

Pedido del usuario: "Responsable de envio debe seleccionar qué instituciones educativas están presentes"
(confirmado por AskUserQuestion: casilla nueva junto a Responsable de envío, que filtra la matriz de
estamento — no existía ningún otro mecanismo previo de "selección de escuelas" que remover).

- **`ParticipacionEstamento.gs`**: nueva columna `PRESENTE` (SI/NO) en la hoja `ParticipacionEstamentoIE`
  (se agrega sola al final vía la auto-migración de `obtenerHoja_`, sin tocar filas existentes).
  `obtenerParticipacionEstamentoGrupo` ahora incluye `presente` por cada IE, y `totalesPorEstamento`/
  `totalGeneral` se calculan **solo sobre las IE presentes**. Nueva función `guardarPresenciaIE` (UPSERT,
  no pisa los conteos por estamento ya guardados de esa IE).
- **`Code.gs`**: nuevo RPC `rpcGuardarPresenciaIE`.
- **`Index.html`**: nueva tarjeta "Instituciones educativas presentes" en Participación, justo después de
  Responsable de envío (antes de "Cantidad de asistentes por estamento e institución").
- **`Components.html`**: nueva `renderListaIEPresentes(instituciones)` (checklist); `renderParticipacionEstamento`
  ahora filtra `datos.instituciones` a solo `.presente` antes de construir la tabla horizontal — si ninguna
  IE está marcada, muestra un aviso en vez de una tabla vacía.
- **`JS.html`**: `cargarParticipacionEstamento` pinta el checklist; nuevo listener delegado en `change` para
  `[data-ie-presente]` que llama a `rpcGuardarPresenciaIE` y refresca la matriz. `institucionesSinParticipacionMarcada_`
  (la advertencia de "verificar participación" en Confirmación de caracterización) ahora solo considera IE
  presentes.
- **Nota**: el informe generado (Informes.gs) no incluye hoy una tabla de la matriz de estamento — no había
  nada que filtrar ahí; si se agrega esa tabla al informe en el futuro, debe construirse ya filtrada por
  `presente`, igual que la pantalla.

Verificado: `node --check` de ParticipacionEstamento.gs, Code.gs y Valoracion.gs; extracción y `node
--check` de los bloques `<script>` de Index.html/JS.html/Components.html/Modal.html/CSS.html (limpio,
mismo falso positivo permanente); sin IDs duplicados; comparación automatizada `getElementById` vs. IDs
reales (mismas 8 referencias colgantes ya confirmadas con guarda nula del lote anterior, ninguna nueva).

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

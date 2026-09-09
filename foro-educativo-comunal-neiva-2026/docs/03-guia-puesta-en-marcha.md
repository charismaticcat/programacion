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
- **`AccesosGrupo`**: completar manualmente `EMAIL_RESPONSABLE_GRUPO` de cada grupo (para poder enviarle
  el acceso y luego el informe) — las demás columnas las genera el paso 4.

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

# FASE 2 — Implementación: ajustes a la arquitectura

La implementación sigue la arquitectura de la [Fase 1](FASE1_ARQUITECTURA.md). Durante la construcción aparecieron inconsistencias reales entre ese diseño, las decisiones del cliente y los requisitos de la Fase 2. Este documento registra cada ajuste y su motivo.

## 1. Un solo proyecto Apps Script con login propio para administradores

**Fase 1:** dos proyectos (público "Ejecutar como yo" y admin "Ejecutar como usuario que accede"), con código compartido por build.

**Problema:**

- La Fase 2 pide un único conjunto de archivos (`Code.gs`, `Index.html`, `Admin.html`…).
- Con dos proyectos, `LockService` **no coordina entre proyectos**: los consecutivos (radicados, IDs de cliente) y las escrituras simultáneas del formulario y del admin podían chocar.
- "Ejecutar como usuario que accede" obliga a **compartir las hojas** con cada administrador. Así cualquier vendedor podría abrir la base completa, que era la limitación anotada en §9 de la Fase 1.

**Solución:**

- Un proyecto y una implementación ("Ejecutar como yo", acceso público).
- El Super Admin usa **email + contraseña** con hash HMAC-SHA256 iterado, sal por usuario y un secreto del servidor.
- Sesión en el servidor, cierre por inactividad, bloqueo tras 5 intentos y 2FA por email opcional.
- Las hojas **no se comparten con nadie** y el rol se valida en cada llamada. Un vendedor ya no puede ver clientes que no tiene asignados.

## 2. Identificación (Parte F) sin exigir identificarse (D10)

**Conflicto:** la Fase 2 pide buscar en la base y precargar datos, pero la decisión D10 dice que el cliente no necesita identificarse.

**Solución:** la búsqueda es un **paso opcional** ("Buscar mis datos" / "Llenar desde cero"), con estas protecciones para no exponer los 30.780 registros:

- **Teléfono o email:** solo coincidencia exacta, y entonces se precargan todos los datos.
- **Nombre o razón social:** se muestran como máximo 5 resultados con el teléfono y el email **enmascarados**. Al elegir, **no** se precargan teléfono ni email; solo se muestran como pista.
- **Resultados:** llevan referencias opacas de un solo uso, nunca el `CLIENTE_ID`. El servidor solo acepta referencias que salieron de la búsqueda de ese mismo borrador.
- **Límites:** 25 búsquedas por formulario y 120 por minuto en total. Todas quedan en la auditoría.
- **Correcciones:** si el cliente elige un registro de la base, los campos que estaban vacíos se completan solos. Los que cambian quedan como **correcciones propuestas** que un admin aplica o descarta.

## 3. Estructura de Drive (Parte E)

Se adopta exactamente la estructura pedida:

```text
CLIENTES_WEB/
  WEB-2026-000001_NOMBRE/
    01_IDENTIDAD_VISUAL/ 02_EQUIPO/ 03_PRODUCTOS/ 04_TESTIMONIOS/ 05_DOCUMENTOS/ 06_OTROS/
```

- El radicado se asigna al iniciar el formulario, para poder nombrar la carpeta desde la primera carga.
- Los archivos quedan asociados al cliente por `ARCHIVOS.CLIENTE_ID`.
- Los reportes y el JSON van en `05_DOCUMENTOS`.
- Los clientes creados por un admin usan `CLI-000123_NOMBRE`.
- Se eliminan los "buckets" de 1.000 carpetas y la carpeta de cuarentena. La validación (extensión, MIME real por firma de bytes y tamaño) ocurre **en memoria, antes** de escribir en Drive, así que ningún archivo inválido llega a Drive.

## 4. Etapas del CRM (Parte G)

Se usan las 10 etapas pedidas: Nuevo, Por contactar, Contactado, Interesado, Reunión, Propuesta enviada, Negociación, Ganado, Perdido y No interesado.

- La etapa vive en la **oportunidad**, y el Kanban es de oportunidades.
- `CLIENTES.ESTADO_CRM` guarda una copia de la etapa de su oportunidad abierta más reciente. Es necesaria para filtrar 30.780 clientes sin leer todas las oportunidades, y solo la actualiza el servidor.

## 5. Hojas de cálculo

Se usan **2** hojas de cálculo en lugar de 5: `CRM_WEB_DB` (24 hojas) y `CRM_WEB_AUDITORIA`.

- Así hay un solo ID principal (como pide la Parte L) y la auditoría, que crece rápido, queda aparte.
- No se crea copia `DB_ORIGEN`: el archivo de la base se lee directamente en la importación, y la columna NIT nunca se lee (D3/D8).

## 6. Otras simplificaciones justificadas

| Fase 1 | Fase 2 | Motivo |
|---|---|---|
| Hoja `FORM_CAMPOS` para definir el formulario | Ajuste `REQUIRED_FIELDS` | El formulario tiene pasos fijos y de diseño cuidado. Lo que cambia en la práctica es qué campos son obligatorios. |
| Estado `REQUIERE_CORRECCION` | Eliminado | Sin email (D12) no hay forma de pedirle al cliente que vuelva. El admin corrige directamente (D11). |
| Carga por fragmentos | Una llamada ≤ 10 MB con compresión previa en el navegador | Más simple y suficiente para fotos reducidas a 2000 px. |
| `COLA_EMAILS` | Eliminada | El sistema no envía emails a clientes (D12). Solo el 2FA opcional de administradores. |
| — | Nueva tabla `PAGOS` | La Parte C pide gestionar pagos. |
| — | Nueva tabla `PROPUESTAS` | Ya estaba en el diseño; se implementa con ítems y totales calculados en el servidor. |

## 7. Verificación

- **Importación:** ejecutada sobre la base real `CCH@E26-9316.xlsx` en el simulador. Reproduce exactamente las cifras de la Fase 1: 30.780 registros, 28.670 con email, 2.110 en blanco, 1.496 teléfonos con 608, 7.754 actividades desde el CIIU y ningún NIT almacenado.
- **Servidor:** 56 pruebas automáticas.
- **Interfaz:** 5 recorridos en Chromium: escritorio, celular, sin conexión y Super Admin completo.

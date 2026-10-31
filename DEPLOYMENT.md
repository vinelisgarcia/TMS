# Sincronización de sesiones

La versión `20261031-session-sync` usa `tms_get_snapshot` para leer una copia
consistente y `tms_save_snapshot` para guardar todas las tablas modificadas en una
transacción con revisión esperada. Una revisión antigua produce `40001` sin
escribir datos. El cliente conserva los cambios pendientes por cuenta y permite
descargarlos antes de cargar la versión compartida. No fuerza reintentos sobre
una revisión nueva.

Las sesiones reciben invalidaciones de la fila `tms_sync_state` por Realtime.
También comprueban la revisión cada ocho segundos y al recuperar conexión o
volver a la ventana. La carga no dispara un guardado. Los cambios hechos mientras
otro guardado está en curso se envían después, en orden.

## Orden de publicación

1. Ejecutar las pruebas locales de JavaScript, adaptadores y PostgreSQL indicadas
   abajo. Conservar los cambios ajenos a la sincronización.
2. Aplicar `supabase/migrations/0008_session_sync.sql` en el proyecto Supabase TMS
   `vbyrlgaygjqnzjwjaczh` mediante herramientas administrativas. La migración no
   reemplaza ni siembra datos operativos y no aplica las migraciones 0005–0007
   pendientes ni cambia los permisos de los roles.
3. Ejecutar `tests/session-sync-rollback.sql` con acceso administrativo. Cada
   comprobación usa cuentas activas existentes dentro de una transacción que
   termina en `ROLLBACK`; no crea cuentas ni cambia contraseñas.
4. Publicar el frontend que incluye `tms-sync.js` antes de `tms-v2.js` en
   `index.html`. GitHub `main` dispara el despliegue de producción en Vercel.
5. Verificar la publicación y la coincidencia de los archivos con el commit,
   además de la publicación Realtime y los permisos de las nuevas funciones.

Desde el paso 2, una pestaña antigua ya no puede guardar sin control de versión.
Debe recargar para utilizar el nuevo cliente. Si falla la carga de la nueva RPC,
el frontend muestra el error y conserva el acceso a recuperación; no recurre al
reemplazo antiguo de tablas. No revertir únicamente el frontend: el cliente
antiguo es incompatible con las protecciones nuevas.

## Pruebas

```sh
node --check tms-sync.js
node --check tms-v2.js
node --test tests/session-sync.test.cjs tests/session-sync-adapter.test.cjs tests/audit-regression.cjs tests/warehouse-regression.cjs
git diff --check
```

La validación de PostgreSQL usa una base efímera con usuarios ficticios y no
contacta producción. Instalar la dependencia de prueba fuera del sitio:

```sh
npm install --prefix /tmp/tms-sync-validation @electric-sql/pglite
TMS_PGLITE_RUNTIME=/tmp/tms-sync-validation node tests/session-sync-db.cjs
```

Incluye revisiones obsoletas, rollback de fallos parciales, permisos de lectura,
sesiones de la misma cuenta, reconexiones, ediciones durante un guardado,
recuperación local separada por usuario y más de mil filas.

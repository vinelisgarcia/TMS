# Calendario y sincronización entre sesiones

Esta revisión requiere desplegar juntos el frontend y la migración
`supabase/migrations/0008_versioned_cloud_sync.sql`, después de las migraciones
0001–0007. No se ha aplicado a la base de datos alojada ni publicado el frontend.

## Despliegue

1. Conserva un respaldo de producción y de los cambios locales pendientes de las
   pestañas abiertas. La migración añade una revisión global y funciones; no
   sustituye las líneas, materiales ni configuraciones existentes.
2. Aplica 0008 en el proyecto Supabase al que apunta el frontend. Despliega
   `index.html`, `tms-v2.js`, `tms-sync.js`, `tms-calendar-details.js`,
   `tms-audit-ui.js`, `tms-domain.js`, `tms-glass.css` y los recursos existentes.
3. Recarga las pestañas antiguas. La migración rechaza sus escrituras directas
   para evitar que sobrescriban datos sin comprobar la revisión. Si el frontend
   nuevo llega antes que la migración, permite consultar las tablas paginadas,
   pero bloquea el guardado hasta que se aplique 0008.
4. Comprueba con dos sesiones autenticadas: guarda una modificación de prueba
   autorizada en A y verifica que B la recibe al recuperar el foco o en la próxima
   consulta (cada 15 segundos mientras la pestaña está visible). No se exige
   activar Supabase Realtime. El inicio de sesión y la consulta no guardan datos.
5. Si ambas sesiones editan, la segunda conserva su versión local y muestra el
   conflicto. Descarga el respaldo antes de descartar esa versión y recargar la
   nube; reconcilia los cambios necesarios. No hay fusión automática de conflictos.

El estado de nube muestra consultas, guardados, pendientes, errores y conflictos.
Los guardados se serializan dentro de la pestaña y usan una única transacción de
base de datos con una revisión esperada. Los respaldos pendientes se separan por
cuenta y pestaña; Web Locks impide que las pestañas duplicadas compartan el mismo
identificador. Los respaldos de versiones anteriores sin cuenta identificada se
archivan antes de actualizar la copia principal. Un administrador puede usar
«Descargar respaldo anterior» para revisarlos y reconciliar los cambios; nunca
se importan automáticamente en una cuenta. Si el almacenamiento del navegador falla, el guardado remoto
continúa y el respaldo descargable se genera desde la memoria actual.
Los cambios de contraseña obligatorios de usuarios no administradores guardan
solo sus propios metadatos de cuenta en esa misma transacción, sin cambiar roles,
identidad ni permisos de módulos.

## Validación local

Desde la raíz del repositorio:

```sh
python3 -m http.server 8765 --bind 127.0.0.1 --directory . >/tmp/tms-http.log 2>&1
```

En otra terminal, con Node, Playwright y Chromium instalados:

```sh
node --test tests/audit-regression.cjs tests/sync-regression.cjs
node tests/warehouse-regression.cjs
TMS_BROWSER=/usr/bin/chromium node tests/ui-smoke.cjs
TMS_BROWSER=/usr/bin/chromium node tests/calendar-sync-smoke.cjs
```

Las dos pruebas de navegador bloquean o simulan las peticiones a Supabase; no
escriben producción. Chromium necesita acceso a su almacén NSS del certificado
del proxy; no desactives TLS. Si cdnjs está bloqueado, el modo opcional
`TMS_CDN_CACHE=/ruta/cache` permite servir tres archivos locales de las mismas
versiones, obtenidos con TLS verificado:

| Archivo local | Origen alternativo |
| --- | --- |
| `xlsx.full.min.js` | `https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js` |
| `html2pdf.bundle.min.js` | `https://cdn.jsdelivr.net/npm/html2pdf.js@0.10.1/dist/html2pdf.bundle.min.js` |
| `chart.umd.min.js` | `https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js` |

Ese modo valida las bibliotecas y la aplicación; no valida la disponibilidad del
CDN original de producción.

Para probar la migración, utiliza un PostgreSQL local desechable y el módulo
Node `pg` (se verificó con PostgreSQL 18.6 y pg 8.16.3):

```sh
TMS_PG_TEST_URL=postgres://usuario@127.0.0.1:55432/postgres \
TMS_PG_MODULE=/ruta/a/node_modules/pg node tests/cloud-sync-postgres.cjs
```

La prueba solo acepta destinos loopback, crea y elimina su propia base de datos
y utiliza usuarios sintéticos. Crea los roles locales `authenticated` y `anon`
si faltan. Verifica todas las migraciones, repetición de 0008, concurrencia,
reversión completa ante errores, permisos y una lectura de 1206 líneas.

## Límites de la comprobación

Los detalles del calendario reflejan los saldos del plan/control importados.
Las coincidencias de tránsito no confirman reservas. Los importes facturados
solo se muestran si hay líneas coincidentes del reporte SAP de facturación
efectiva; no se calculan a partir de cantidades ni de materiales en tránsito.

La validación local no confirma que la URL publicada use esta revisión, que 0008
esté aplicada en el proyecto alojado ni que una sesión real pueda escribir allí.
Esas comprobaciones necesitan la URL de producción y acceso autorizado.

# Sincronización de sesiones

El nuevo módulo independiente de showroom requiere aplicar la migración aditiva
`0010_showroom_inventory.sql` antes de publicar su frontend. Ver
[SHOWROOM.md](SHOWROOM.md) para el alcance, los permisos y las pruebas. Preparar
el cambio en una rama no aplica esa migración ni publica el sitio.

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

## Fichas del calendario

Formato compartido del frontend: pedidos distintos, confirmado no entregado,
entregado pendiente de facturar y valor del cliente en esa ruta (suma de ambos
importes). Las líneas se limitan al grupo cliente/fecha/camión/ruta. Los avisos
agrupan facturas en tránsito con sus ETA; el detalle de una factura contiene
únicamente las líneas afectadas del despacho. Las facturas recibidas se excluyen.
El detalle abierto se recalcula al recibir una nueva versión compartida.

No añade datos ni cambia el esquema. Publicar juntos `tms-domain.js`, `tms-v2.js`
e `index.html`, que versiona ambos scripts. Las sesiones que ya estaban abiertas
antes de publicar deben recargar para obtener el nuevo frontend.

Validar con `node --test tests/calendar-dispatch.test.cjs` y
`node tests/calendar-dispatch-ui.cjs`. La prueba de navegador usa dos sesiones
con las RPC reales del cliente contra un servidor simulado; bloquea Supabase y
no modifica producción. También ejecutar `tests/ui-smoke.cjs` para verificar
movimientos de rutas, exportaciones y módulos existentes.

## Uso desde teléfono y tablet

`tms-responsive.css` y `tms-responsive.js` completan el frontend común: encabezado
compacto, accesos rápidos y selector de módulos mediante diálogo. Las opciones
se derivan de los módulos autorizados existentes; no amplían permisos. En móvil
el calendario ocupa una columna y muestra inicialmente los días con entregas o
notas; «Todos los días» permite ver y programar el resto. Los detalles de despacho
se presentan por campos en lugar de una tabla ancha. El escritorio conserva su
barra lateral y calendario semanal.

Importar SAP utiliza tres fichas de documento con los mismos inputs y parsers:
programación, control diario y solicitudes/control de almacén. El progreso y la
confirmación se muestran dentro de cada ficha; el control diario requiere una
programación. Se conservan las validaciones, los datos y la sincronización.

Publicar `index.html`, `tms-v2.js` y ambos archivos responsive juntos. Recargar las
sesiones abiertas para obtener el nuevo frontend. Validación adicional:
`node tests/mobile-workspace-ui.cjs` usa archivos ficticios, bloquea Supabase y
prueba cargas correctas/incorrectas, permisos, navegación y controles visibles
en 10 módulos a 320/390/430/768/1024/1440 píxeles, además de móvil en horizontal.

La misma prueba se ejecutó también con WebKit mediante
`TMS_TEST_ENGINE=webkit node tests/mobile-workspace-ui.cjs`.

El calendario móvil utiliza fichas por día y tarjetas compactas de cliente. El
menú de tres puntos conserva Mover, Completar y Nota; en escritorio las acciones
siguen visibles. Los controles de lote aparecen al seleccionar despachos y se
omite la cola vacía en móvil. En tablet se distribuyen los días en dos columnas.

## Alertas de tránsito y saldo pendiente

El cruce requiere una falta sin stock de la línea del pedido. El saldo explícito
pendiente de servir limita esa falta; cero la elimina. La cantidad ya facturada
contra el pedido original también limita el faltante. Un artículo disponible
completo no genera aviso por compartir referencia con una importación, y otro
pedido del cliente no hereda esa alerta. Las líneas totalmente sin stock siguen
avisando aunque tengan cero unidades disponibles para el despacho.

El control diario actualiza el saldo pendiente de servir de la línea programada.
Los adaptadores conservan ceros explícitos al guardar/recargar; un saldo ausente
en datos antiguos se reconstruye con unidades confirmadas y sin stock. Las
pruebas cubren pedido cubierto, parcial, falta total y actualización de un detalle
abierto en otra sesión, incluida su recarga. No requiere migración de esquema.

## Packing lists en Materiales en tránsito

El mismo importador admite facturas y packing lists de Excel/CSV. Detecta la fila
de encabezados después de títulos o filas vacías y procesa las hojas visibles
que contengan columnas de material y cantidad. La planilla de entrada reconoce
`Codigo cliente del producto`, `Descripción de Producto`, `Referencia proveedor`,
`Packslip`, `Cantidad solicitada`, `PEDIDO DOM`, `Proveedor`, `Factura No.` y
`Referencia embarque`. El código interno tiene prioridad sobre la referencia del
proveedor; PEDIDO DOM no se interpreta como cliente.

Se conserva la factura cuando existe; en su ausencia se identifica el documento
por Packslip. La ETA y la asignación opcional de cliente se indican en el
formulario cuando no vienen en el archivo. Antes de modificar el estado se
confirman el número de líneas y la cantidad total. Las líneas repetidas del
mismo documento/material/ETA/cliente dentro del archivo se suman. Recargar ese
mismo grupo reemplaza su cantidad, conservando el identificador y la recepción.
Las cantidades inválidas bloquean la carga completa. Los datos de origen se
conservan en el estado compartido existente y en la exportación Excel; no se
necesita migración de base de datos.

Publicar juntos `index.html`, `tms-v2.js` y `tms-audit-ui.js`. Validar con
`node tests/transit-import.test.cjs` y `node tests/transit-import-ui.cjs`, además
de las regresiones de interfaz, calendario y móvil. La prueba de navegador usa
datos sintéticos y bloquea Supabase. Puede validarse un archivo privado con
`TMS_PACKING_FILE` y, opcionalmente, `TMS_PACKING_EXPECTED_LINES` y
`TMS_PACKING_EXPECTED_QUANTITY`, sin copiarlo al repositorio ni cargarlo en
producción. Usa `TMS_PLAYWRIGHT` y `TMS_TEST_URL` como las demás pruebas de UI.

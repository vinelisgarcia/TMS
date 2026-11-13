# Sincronización de sesiones

El nuevo módulo independiente de showroom requiere aplicar la migración aditiva
`0010_showroom_inventory.sql` antes de publicar su frontend. Ver
[SHOWROOM.md](SHOWROOM.md) para el alcance, los permisos y las pruebas. Preparar
el cambio en una rama no aplica esa migración ni publica el sitio. Una publicación
de la interfaz sin esa migración deja el showroom en «Activación pendiente», sin
cargas ni conteos, hasta que el administrador active su almacenamiento.

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

Materiales en tránsito admite tres etapas por factura: En tránsito, Recibido y
Contabilizado. Recibido confirma la entrada física; Contabilizado registra la
confirmación manual de que la factura ya fue contabilizada en SAP en un proceso
posterior. Solo se permite contabilizar una factura cuando todas sus líneas
están recibidas. No ejecuta una contabilización ni modifica SAP.

Se conservan `receivedAt/receivedBy` y `accountedAt/accountedBy` por separado,
con un historial de cambios. Quitar la marca contable deja la mercancía recibida;
para volver a tránsito hay que retirar primero esa marca. Recibido y Contabilizado
se excluyen de alertas y cobertura futura del calendario. La reimportación de la
misma identidad conserva ambas confirmaciones. El Excel incluye estado, fechas
y responsables. Los campos viajan en `app_state.transitMaterials` mediante la
sincronización existente; no requiere migración ni activación adicional.

Validar con `node tests/transit-accounted-ui.cjs`,
`node --test tests/audit-regression.cjs tests/transit-entry-import.test.cjs` y la
prueba PostgreSQL `tests/session-sync-db.cjs`. Todas usan datos ficticios.

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

## Planilla de entrada de mercancía en tránsito

El importador admite, además de las facturas y formatos anteriores, la planilla
con `Código cliente del producto`, `Cantidad solicitada` y `Factura No.`. Detecta
los encabezados después de títulos/filas vacías y busca hojas compatibles. El
código del producto se conserva como referencia AD; no se utiliza Referencia
proveedor como SKU ni Pedido DOM como cliente o pedido de venta.

Las líneas se consolidan por factura, artículo, ETA y cliente asignado. Facturas
con puntuación distinta siguen separadas. Una recarga sustituye únicamente las
coincidencias de esa identidad y conserva el resto, sus identificadores y su
recepción registrada. Se guarda la trazabilidad de las filas de la planilla,
incluidos pedido DOM, proveedor, packslip y referencia de embarque.

Factura y ETA de cada fila prevalecen sobre los campos del formulario. Si falta
la factura, puede completarse con el formulario. Si falta ETA en ambos sitios,
queda por confirmar y se indica al terminar; nunca se deduce del packslip. Un
error de código, cantidad, factura o fecha bloquea toda la carga antes de cambiar
los materiales existentes. No cambia el esquema de persistencia ni los permisos.

Validación: `node --test tests/transit-entry-import.test.cjs` y
`node tests/transit-entry-import-ui.cjs`. Los datos de prueba son ficticios y las
peticiones a Supabase están bloqueadas. Un archivo privado adicional puede
validarse localmente con `TMS_TRANSIT_ENTRY_XLSX` y, opcionalmente, un objeto JSON
de totales esperados en `TMS_TRANSIT_ENTRY_EXPECTED`; no se copia al repositorio.

## Tamaño de lectura

El botón Aa abre Normal, Grande (inicial) y Muy grande. La preferencia se guarda
por usuario en el almacenamiento de este navegador y se actualiza entre pestañas
del mismo navegador. No modifica datos operativos ni solicita permisos de edición.
El tamaño se aplica al texto de módulos y detalles, con etiquetas de al menos
12 px en los modos ampliados. Se recalculan los tamaños originales al cambiar de
módulo, renderizar datos o redimensionar; no se acumula el aumento en texto anidado.
El encabezado compacto y los controles de iconos conservan su disposición, y el
escalado no se aplica a las exportaciones ni a la impresión.

Validación: `node tests/reading-size-ui.cjs`, también con `TMS_TEST_ENGINE=webkit`.
Comprueba tamaños, cierre del selector, persistencia por cuenta, actualización
entre pestañas, texto nuevo, ausencia de cambios/guardados operativos y diez
módulos a 320/390/768/1440 px con el tamaño mayor.

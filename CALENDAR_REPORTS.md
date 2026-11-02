# Informes ejecutivos del calendario

El botón **Informe ejecutivo** abre una vista previa y permite descargar Excel
o PDF desde el detalle de un cliente, una factura en tránsito, nota/evidencia,
mover un cliente, mover una selección, asignar pedidos de la cola y notas públicas
del día. Las confirmaciones de acciones y mensajes de error no generan informes.

Cada apertura resuelve los datos actuales de ese despacho y conserva una
instantánea con fecha de generación. Una factura exporta únicamente las líneas
afectadas de ese cliente en ese despacho; los indicadores se recalculan sobre
ese alcance. Una factura recibida, retirada o que ya no afecta al pedido no
puede exportarse desde una referencia obsoleta.

El valor es confirmado no entregado más entregado pendiente de facturar, en DOP.
Las cantidades no se suman entre artículos distintos. El tránsito se presenta
como coincidencia por material, sin considerarlo reserva ni sumar repetidamente
la misma cantidad de factura como disponibilidad asignada.

Excel incluye Resumen, Pedidos, Tránsito y Seguimiento, con formatos numéricos,
identificadores de texto, filtros, encabezados fijos y total DOP. PDF genera texto
seleccionable y tablas paginadas con encabezados repetidos y número de página.
Las notas, evidencias y propuestas sin guardar se identifican como borradores.
Las fotografías adjuntas se referencian por su nombre; no se incrustan.

Exportar no guarda, mueve, reserva ni modifica datos operativos. Los usuarios
con acceso de lectura al calendario pueden exportar; los permisos se comprueban
antes de abrir y antes de descargar. Los generadores se cargan bajo demanda
desde `vendor/`, con versiones fijas y licencias incluidas. Ningún dato del
informe se envía a otro servicio para convertirlo.

Pruebas sintéticas, sin peticiones al backend de producción:

```sh
node --test tests/calendar-reports.test.cjs
node tests/calendar-reports-ui.cjs
TMS_TEST_ENGINE=webkit node tests/calendar-reports-ui.cjs
node tests/calendar-dispatch-ui.cjs
node tests/reading-size-ui.cjs
```

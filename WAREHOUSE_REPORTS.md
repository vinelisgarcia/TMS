# Informes ejecutivos de solicitudes a almacén

La ficha de Solicitudes a almacén en Importar SAP ofrece **Informe ejecutivo**
para el conjunto cargado. El módulo Solicitudes ofrece **Informe ejecutivo · PDF /
Excel**, que respeta la pestaña y la búsqueda actuales. Los botones requieren el
permiso Ver de Solicitudes a almacén; no requieren edición. No se amplían permisos
de ningún rol.

La vista previa y las descargas incluyen indicadores de solicitudes únicas,
clientes, pedidos y proporción de líneas preparadas, barras de preparación,
agenda por fecha/camión/ruta y detalle de pedidos, materiales, cantidades,
estados y motivos pendientes. El Excel tiene Resumen ejecutivo y Solicitudes,
con filtros, paneles inmovilizados, fechas nativas para entregas de un solo día,
identificadores de texto y barras de datos. El PDF usa tablas con paginación.
No se suman cantidades de materiales distintos ni se presentan importes estimados
como valoración de la preparación.

Las fechas de entrega proceden exclusivamente de `APP.lineItems`, la fuente que
forma el calendario. El cruce utiliza cliente y pedido, y exige posición y material
cuando la solicitud los identifica. Se excluyen filas archivadas. No se usa la fecha
del archivo, la fecha actual ni otro pedido del mismo cliente como sustituto.
Las líneas sin programación figuran como **Sin programar en calendario**. Si hay
varias fechas/camiones se muestran todos; si parte del pedido no está programada,
se indica programación parcial. No se atribuye toda la cantidad a cada fecha de
un despacho dividido. La agenda cuenta referencias por despacho y puede repetir
una solicitud entre fechas; el resumen cuenta solicitudes únicas.

Planificación y control se cruzan por cliente, solicitud, pedido, posición y
material. El control aporta la preparación más reciente; se conserva la cantidad
solicitada de planificación, incluido el cero explícito. Así se evita duplicar
líneas por haber importado ambos documentos. Con planificación cargada, las vistas
programadas y pendientes incluyen sus líneas y no agregan solicitudes ajenas de
un control anterior. Completadas conserva la consulta de líneas del control.
La misma base alimenta las fichas
del módulo y los reportes. Los archivos originales no se modifican.

Cada descarga vuelve a resolver los datos y fechas vigentes, aunque se haya movido
un pedido desde que se abrió la vista previa. El reporte es una instantánea con
fecha de generación y fuentes. Se comprueba el permiso y la cuenta al finalizar la
exportación; cerrar la vista o perder acceso invalida la descarga pendiente. La
generación de informes no escribe datos operativos ni cambia el calendario.

Pruebas sintéticas, con las solicitudes al backend bloqueadas:

```sh
node --test tests/warehouse-reports.test.cjs
node tests/warehouse-regression.cjs
node tests/warehouse-reports-ui.cjs
TMS_TEST_ENGINE=webkit node tests/warehouse-reports-ui.cjs
node tests/calendar-reports-ui.cjs
```

La prueba de navegador importa archivos Excel, comprueba un archivo con fechas
ajenas al calendario, reprograma antes de descargar y verifica contenido, tipos
Excel, filtros, permisos, dispositivos de 320/390/768 px y las 120 líneas de un PDF
largo. Los PDF y Excel de prueba son sintéticos y permanecen fuera del repositorio.

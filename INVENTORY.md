# Inventario CT / AD

El módulo compara mercancía de Álvarez Dominicana depositada en Caribetrans.
Se abre en **Inventario CT / AD**. Admin tiene acceso; los demás roles necesitan
permisos explícitos `inventario.ver`, `inventario.editar` o `inventario.importar`
en la configuración de roles. La migración no modifica las matrices existentes.

## Carga directa

1. Seleccionar CT o AD y cargar Excel (`.xlsx`, `.xls`) o CSV.
2. Revisar hoja, fila de encabezados, fecha de corte, separador decimal y columnas.
   AD requiere cantidad y valor **total** en DOP; no se toma costo promedio como
   si fuera valor total. Si falta la unidad, se muestra la suposición UND.
3. Confirmar el alcance, calcular la vista previa y revisar errores, repeticiones
   y totales. Las filas inválidas bloquean la confirmación.
4. Confirmar. La carga sustituye esa fuente; no acumula reportes ni versiones del
   mismo corte. Se compara una huella SHA-256 del contenido normalizado para
   reconocer CSV equivalentes con distinto nombre o codificación.

Los originales no se suben a una ubicación pública ni se incluyen en el código.
La base privada almacena los datos normalizados, trazabilidad por fila, nombre,
fecha, usuario y huellas del reporte. No se precargan archivos del usuario en el
sitio publicado. Cada pareja de reportes mantiene sus observaciones; el historial
por material permanece disponible después de reemplazar los archivos.

## Cálculo

Se suman las cantidades por material y unidad compatible, conservando las filas,
almacenes y lotes originales. Las repeticiones exactas se suman y se señalan para
revisión; no se eliminan sin evidencia de que sean una duplicación de stock.
No se suman unidades incompatibles y no se borran ceros iniciales de códigos.

Después de consolidar AD:

```text
Costo ponderado AD = SUM(valor total DOP) / SUM(cantidad AD)
Diferencia = cantidad AD − cantidad CT
Impacto absoluto = ABS(diferencia) × costo ponderado AD
```

El costo no se redondea antes de multiplicar. El importe de cada material se
redondea a centavos; los totales suman esos importes. Se presentan impacto
absoluto y neto por separado para que sobrantes y faltantes no se compensen en
la priorización. El orden principal es impacto absoluto descendente.

Un material sin costo AD suficiente, con cantidad AD no positiva o con unidades
incompatibles queda sin valoración; no se inventa costo cero. Un valor AD cero
explícito sí se conserva como cero. CT FOB y sus monedas no se usan para valorar
las diferencias DOP. Las fechas distintas generan un aviso de corte.

El seguimiento permite responsable, notas, estado y soporte. Registrar una
rectificación o justificación exige una nota de revisión; la referencia de ajuste
o soporte es opcional. Se permite registrar materiales solo en CT y sin
valoración AD. Los avisos de validación y errores aparecen dentro del formulario.
El diálogo muestra Guardando y se cierra tras recibir la confirmación de la RPC.
Si falla el guardado o existe un conflicto entre sesiones, conserva los campos y
la copia local para reintentar o revisar la versión compartida. Los datos de los
reportes no cambian: la confirmación económica se verifica en una nueva carga.

## Carpeta en la nube

La carga directa es la opción disponible: permite revisar el corte y el mapeo
antes de reemplazar una fuente. Una carpeta fija de Drive o OneDrive requiere
definir proveedor, ubicación privada, acceso, emparejamiento CT/AD y un proceso
de ingestión por eventos o calendario. El conector de la conversación no es una
credencial de ejecución del TMS. No se anuncia una sincronización de carpeta que
no esté conectada. El mismo parser y control de identidad permiten incorporarla
posteriormente sin cambiar el cálculo.

## Persistencia y publicación

La migración `supabase/migrations/0009_inventory_reconciliation.sql` añade estado
y revisión aislados, sin modificar inventarios operativos anteriores. Las RPC
`tms_inventory_get_snapshot` y `tms_inventory_save_snapshot` usan revisión
esperada y transacción atómica. Realtime publica solo la revisión; no el payload
económico. El controlador tiene namespace, RPC y respaldos propios para evitar
interferir con la sincronización operativa.

Aplicar la migración antes de publicar los archivos nuevos. Ejecutar:

```sh
node --test tests/inventory-domain.test.cjs tests/inventory-sync.test.cjs
node tests/inventory-db.cjs
node tests/inventory-ui-smoke.cjs
node tests/inventory-case-ui.cjs
TMS_TEST_ENGINE=webkit node tests/inventory-case-ui.cjs
```

Los tests de base usan PGlite instalado fuera del repositorio como se documenta
en `DEPLOYMENT.md`. La prueba de interfaz usa RPC y usuarios ficticios y bloquea
las llamadas reales a Supabase. Los CSV privados pueden validarse con las
variables `TMS_CT_CSV` y `TMS_CT_CSV_ALT`, sin copiarlos al repositorio.

## Formato AD SAP confirmado

El reporte «Resumen de stock valorado» reconoce `ID de producto`, `Cantidad`,
`Cantidad (Unidad)`, `Descripción de producto`, `Valor calculado` y
`Valor calculado (Moneda)`. La fecha se toma de la celda numérica de Excel junto a
«Última actualización». `UN` se normaliza a `UND`.

Si existe `Tipo de stock`, la vista previa propone `Almacén externo` para el
cruce físico y permite cambiar la selección. `En tránsito` permanece fuera de
ese alcance propuesto; la pantalla indica cuántas filas se excluyeron. Una
selección vacía incluye todos los tipos y muestra un aviso. Cuando existe la
columna de moneda, las filas distintas de DOP bloquean la confirmación.

## Inventory Data y conteos físicos

Los tops son independientes: 10 sobrantes con mayor importe positivo y 20
faltantes con mayor impacto económico (importe más negativo). Al iniciar un
ciclo guiado o libre se incorporan como artículos obligatorios. Los sobrantes
sin costo AD se presentan aparte, sin asignarles valor cero, y pueden añadirse
manualmente al guiado, incluidos los artículos que existen solo en CT.

Inventory Data integra el maestro y las ubicaciones del modelo independiente:
plantillas Articulos y Ubicaciones, Excel/CSV, códigos de barras, proveedor y
subfamilia. Las filas repetidas por artículo/ubicación suman su stock. Las
ubicaciones importadas sustituyen únicamente las de los artículos incluidos.
Los reportes originales entregados traen almacén, pero no posición física;
1015 no se utiliza como pasillo o estante. Puede cargarse una columna Ubicación
en CT/AD o el maestro de ubicaciones por separado.

En ambos conteos se selecciona el artículo y aparece su ubicación. El operador
confirma expresamente o escribe la correcta, sin escanear ubicaciones. Si hay
varias posiciones deben contarse todas, incluso con cero. Las cajas adicionales
requieren confirmación y se suman. Los errores de cantidad se corrigen o anulan
con motivo y bitácora; cualquier cambio vuelve a requerir procesamiento.

El ciclo congela sus referencias CT/AD y ubicaciones; nuevas cargas no alteran
el ciclo anterior. No se permite cerrarlo hasta completar y procesar los
obligatorios y los artículos iniciados. Crear otro ciclo archiva el anterior
con sus pendientes visibles, sin declararlos completados.

Los conteos usan la revisión atómica y los permisos del inventario existente,
sin nuevas tablas ni cambios de roles. Un conflicto conserva el estado local
y permite respaldarlo antes de cargar la versión compartida. También se
recupera el formulario de conteo tras recargar la misma sesión.

La exportación incluye Resumen, Detalle, Correcciones, Guiado, Ciclo y Altas AD.
Esta última contiene existencias positivas, contadas y procesadas de artículos
sin registro AD. Es una propuesta para registrar en AD: no crea artículos en
el ERP ni autoriza automáticamente despachos.

Pruebas adicionales:

```sh
node --test tests/inventory-counts.test.cjs
node tests/inventory-counts-ui.cjs
node tests/inventory-db.cjs
```

La convención confirmada es AD − CT: positivo (verde) cuando AD tiene más,
negativo (rojo) cuando CT tiene más. El top 10 usa positivos y el top 20 los
negativos más extremos. Los artículos solo CT siguen visibles sin valoración.
Los ciclos anteriores conservan referencias; iniciar un nuevo ciclo aplica
los tops de esta convención sin eliminar conteos anteriores.

## Conteo con ubicaciones sin confirmar y apuntes de diferencias

El guiado mantiene los tops obligatorios y ordena los códigos de material antes
de las ubicaciones de referencia. Una ubicación no informada por CT puede quedar
en blanco: se confirma el artículo y la cantidad, sin inventar un pasillo ni usar
el almacén 1015 como ubicación. Las posiciones observadas se distinguen de una
confirmación oficial de CT. Si existe una referencia cargada, el operador también
puede seleccionar Sin ubicación confirmada.

Los motivos disponibles son En revisión de CT, Pedido Interno AD pendiente
revertir, Duplicidad en CT, Duplicidad en AD, Ajuste inventario positivo CT, Ajuste
inventario negativo CT y Mudanza CT. Se incluye Rotura, solicitado previamente.
El responsable CT/AD se selecciona aparte; no se infiere culpabilidad del motivo.

Puede añadirse un apunte al guardar el conteo o desde Resultados y bitácora. El
conteo puede guardarse sin clasificar cuando la investigación siga pendiente.
Cada apunte guarda responsable, motivo, observaciones, usuario, fecha y cantidad
contada en ese momento. Se conservan todos los apuntes; no se sustituyen por el
último. Si cambia el conteo, se señala que el apunte previo requiere revisión.
Los apuntes no modifican cantidades CT/AD ni realizan ajustes en sus sistemas.

El Excel de conteos añade la hoja Apuntes, los campos de clasificación en Resumen
y una columna explícita de confirmación de ubicación por CT en Detalle. Las
sesiones comparten los apuntes mediante la revisión atómica existente. Un
formulario de apunte sin guardar se conserva ante cambios simultáneos y al
recargar. No requiere migración: se guarda dentro del payload privado existente.


## Movimientos AD y trazabilidad entre cortes

Desde Revisar en un material se puede marcar **Registrar una entrada o salida ya
realizada en AD**. Se indican tipo, cantidad positiva y fecha efectiva; el
responsable, el motivo y el soporte opcional se toman de los campos del seguimiento.
La fecha del movimiento es distinta de la fecha/usuario que registran la anotación.
El movimiento no altera las cantidades ni la valoración de los reportes cargados.

**Movimientos y trazabilidad AD** muestra entradas, salidas y neto por material y
unidad, estado de contraste y un historial de observaciones, movimientos, nuevos
cortes y revisiones. Cada cambio de una observación agrega una versión; las
observaciones anteriores disponibles se conservan como tales. No se extraen
cantidades de su texto para inventar movimientos históricos. Se congelan las
referencias de los archivos/cortes y sus cantidades disponibles antes de reemplazar
una fuente. Los saldos históricos que ya no están disponibles permanecen desconocidos.

Al cargar otro AD se compara su cantidad con el saldo AD base más las entradas
menos las salidas registradas después de ese corte. Los movimientos del mismo
material/corte se consideran juntos para calcular ese saldo, sin duplicar el
saldo base por cada entrada. Un movimiento de la misma fecha del corte requiere
confirmar que el archivo aún no lo incluía; los anteriores al corte se mantienen
como históricos para revisión documental. Un reporte cuyo corte preceda al
movimiento no puede verificarlo. Las unidades incompatibles bloquean la proyección.

El nuevo archivo genera evidencia de contraste, incluso si el saldo es distinto.
**Saldo compatible** es un indicador; la confirmación la realiza un usuario con
permiso de edición y una nota de revisión. Las ventas, despachos y otros movimientos
pueden explicar cambios adicionales. Cada revisión se aplica únicamente a la
anotación seleccionada, conservando los movimientos que formaron la base del saldo
esperado. Las revisiones anteriores y los archivos usados permanecen en el historial.

Las anotaciones guardadas se corrigen mediante anulación con motivo y un registro
nuevo, conservando el original. La anulación solo afecta el registro del TMS.
Los reintentos conservan el identificador de un movimiento pendiente para evitar
sumarlo dos veces. Los formularios de movimiento y revisión se recuperan al recargar;
los conflictos entre sesiones conservan la copia local y no sobrescriben la compartida.

Se añaden Movimientos AD e Historial inventario al Excel de conciliación, respetando
los materiales filtrados. El libro del visor de trazabilidad permite exportar el
historial completo o el material buscado, incluidos cortes anteriores.

La persistencia utiliza `tracking.version = 1` y `tracking.events` dentro del payload
privado existente. La misma RPC/revisión atómica y permisos de inventario protegen
el guardado. No requiere nuevas tablas ni migración SQL: el contrato vigente admite
el estado adicional y se valida el viaje de ida/vuelta con PGlite. El historial es
la trazabilidad operativa de la aplicación, con anulaciones visibles en su interfaz.

Pruebas adicionales:

```sh
node --test tests/inventory-trace.test.cjs
node tests/inventory-trace-ui.cjs
TMS_TEST_ENGINE=webkit node tests/inventory-trace-ui.cjs
node tests/inventory-db.cjs
```

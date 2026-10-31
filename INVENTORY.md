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
sitio publicado. Una nueva pareja de reportes tiene su propio seguimiento.

## Cálculo

Se suman las cantidades por material y unidad compatible, conservando las filas,
almacenes y lotes originales. Las repeticiones exactas se suman y se señalan para
revisión; no se eliminan sin evidencia de que sean una duplicación de stock.
No se suman unidades incompatibles y no se borran ceros iniciales de códigos.

Después de consolidar AD:

```text
Costo ponderado AD = SUM(valor total DOP) / SUM(cantidad AD)
Diferencia = cantidad CT − cantidad AD
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
rectificación o justificación exige nota y referencia. No realiza ajustes en los
sistemas de origen: la confirmación económica se verifica en una nueva carga.

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

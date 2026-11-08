# Inventario del showroom

Módulo independiente de Inventario CT/AD. Incluye conciliación, resumen económico,
top 10 sobrantes/top 20 faltantes, seguimiento por artículo, Inventory Data,
ubicaciones, conteo guiado y libre, ciclos, correcciones, bitácora y exportaciones.
Comparte los motores de cálculo y conteo, pero mantiene sus propios controles,
permisos, respaldos locales, estado, revisión y sincronización.

## Fuentes

- **Virtual Álvarez:** reporte «Resumen de stock valorado» del showroom. Usa ID de
  producto, descripción, cantidad/unidad, valor calculado total y moneda DOP.
  La propuesta de alcance incluye el tipo «Almacén». Verifica que corresponde
  al showroom antes de confirmar.
- **Físico:** del archivo de inventario del showroom se selecciona automáticamente
  la hoja **Conteo**. Reconoce `ID`, `Cantidad Contada` y `Ubicacion Showroom`.
  Conserva responsable, sector de la columna contigua sin título y la nota
  original de DIFERENCIAS como trazabilidad. Esa nota no se usa para calcular
  diferencias. Las hojas «Inventario SR SAP» y «Sheet3» no se cargan junto con
  Conteo. También admite reportes individuales con columnas equivalentes y
  mapeo revisable.

El corte se propone desde el reporte virtual y desde el nombre del archivo de
conteo; puede corregirse antes de cargar. Las fechas diferentes generan un aviso.
La unidad ausente se propone como UND y se muestra en la vista previa. Las filas
repetidas se suman conservando sus posiciones. Una cantidad negativa, vacía o no
válida en el conteo bloquea la confirmación completa. Reimportar contenido
equivalente no duplica cantidades; cambios de responsable o notas se detectan.

## Diferencias y valoración

`Diferencia = cantidad física contada − cantidad virtual`.
Positivo es sobrante físico, negativo es faltante y cero coincide. Un artículo
virtual sin registro físico figura **Sin conteo físico**, con cantidad física,
diferencia e impacto vacíos. No se considera faltante ni se incorpora a los tops.
Un cero explícitamente contado sí permite calcular la falta. No se interpreta la
ausencia de un conteo como un cero.

El costo se obtiene del valor total virtual consolidado en DOP dividido entre
su cantidad consolidada. Se muestran impacto absoluto y neto. Un artículo físico
sin registro virtual queda sin valoración; puede añadirse al guiado y, una vez
contado y procesado, exportarse como propuesta de alta virtual. El TMS no ajusta
SAP ni inventa costos para completar el informe.

La exportación incluye conciliación filtrada, fuentes/método, detalle físico y
virtual, con ubicaciones, responsables y notas. Los conteos conservan las hojas
Resumen, Detalle, Correcciones, Guiado, Ciclo y Altas virtuales. Un nuevo ciclo
archiva el anterior y conserva las referencias con las que se inició.

## Activación y permisos

Antes de publicar el frontend, aplicar administrativamente
`supabase/migrations/0010_showroom_inventory.sql`. La migración es aditiva e
idempotente, crea un estado vacío y sus RPC; no importa los archivos del usuario,
no altera CT/AD, no toca la revisión operativa ni cambia matrices de roles.
Admin tiene acceso; los otros roles requieren permisos explícitos del módulo
`showroom` (`ver`, `editar`, `importar`). Tener permisos de `inventario` no concede
acceso al showroom. Solo se publica la revisión mediante Realtime, no los datos.

Publicar juntos index.html, tms-v2.js, tms-responsive.js, tms-inventory.css,
inventory-counts-domain.js y los cuatro archivos showroom-*.js. Recargar las
sesiones abiertas. La preparación del código y las pruebas no aplican la
migración ni publican el sitio.

## Validación local

```sh
node tests/showroom-domain.test.cjs
TMS_PGLITE_RUNTIME=/workspace/.tms-runtime node tests/showroom-db.cjs
TMS_PLAYWRIGHT=/workspace/.tms-runtime/playwright.cjs node tests/showroom-ui-smoke.cjs
TMS_PLAYWRIGHT=/workspace/.tms-runtime/playwright.cjs node tests/showroom-counts-ui.cjs
```

El servidor de pruebas sirve el checkout en el puerto 8766; puede cambiarse con
TMS_TEST_URL. Las pruebas bloquean Supabase y emplean usuarios ficticios. La
migración se comprueba en PostgreSQL efímero, incluyendo permisos, conflictos,
persistencia de ciclos y aislamiento de CT/AD.

Para validar archivos privados en la prueba de UI, indicar
TMS_SHOWROOM_PHYSICAL y TMS_SHOWROOM_VIRTUAL. Los adjuntos y las exportaciones de
datos privados permanecen fuera del repositorio. Ejecutar también las pruebas
existentes de inventario, conteos, interfaz y móvil tras cambiar piezas comunes.

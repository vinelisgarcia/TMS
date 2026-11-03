# Packing list Miami

`packingMiami` integra un generador Excel en el TMS. Admin lo puede usar por
defecto; los demás roles empiezan sin acceso y se configuran desde Roles y permisos.
No se cargan documentos ni datos privados en el repositorio. Los documentos de
trabajo se mantienen en la memoria de la sesión del navegador.

La estructura reproduce las 23 columnas del ejemplo: cabecera en fila 6,
referencia del consolidado, contenedor, precinto, bultos, kg y m³. La plantilla
puede cargarse para verificar el formato y conservar su nombre de hoja. Su
mercancía de ejemplo nunca se incorpora al consolidado. Se crea un libro nuevo
para no reutilizar datos ni metadatos del embarque usado como ejemplo.

La planilla de entrada identifica artículo por código de cliente del producto,
descripción, cantidad solicitada y albarán por Packslip. Pedido DOM requiere
seleccionar explícitamente NPI, PEDIDO_CLIENTE o PEDIDO. Proveedor, referencia del
proveedor, factura, embarque, hoja, fila y cantidad original se conservan en Datos
de origen. Los identificadores se exportan como texto. Las repeticiones de un
material se mantienen por documento, sin eliminar parciales de distintos pedidos.
Un archivo con el mismo SHA-256 no vuelve a sumarse aunque se cambie su nombre.
Una selección que contiene errores se rechaza sin incorporar parcialmente filas.

Los datos físicos y aduaneros ausentes se completan en el editor de cada línea.
También se pueden agregar documentos PDF de cualquier tipo desde **Documentos
PDF** o desde la entrada general: facturas, packing lists, transporte y otros
soportes, con texto, escaneados o mixtos. PDF.js extrae el texto y Tesseract.js 7
aplica OCR en español e inglés a las páginas con imágenes o sin texto legible.
Los lectores, fuentes, motores y modelos se sirven desde el TMS y la lectura se
realiza en el navegador, sin enviar los documentos a servicios OCR externos.
La lectura conserva las páginas, identifica etiquetas explícitas y cruza la referencia
AD o del proveedor y el número de factura con la planilla cargada. Si hay varias
líneas posibles, la revisión exige elegir destino. Un número de factura que
contradice el de la planilla no produce una asignación automática.

**Revisar lectura** muestra el texto por página, los valores detectados, el valor
actual y la evidencia. **Seleccionar datos nuevos** solo marca campos vacíos con
una única línea de destino y sin valores contradictorios. Los valores existentes
se sustituyen únicamente al seleccionarlos expresamente. No se crean líneas de
mercancía ni se suman cantidades desde las facturas; las cantidades leídas se
muestran para comparar y se señalan diferencias. Los totales de peso de factura
no se asignan a bultos: solo se proponen pesos de bulto identificados explícitamente
como tales y con unidades. Las unidades explícitas se convierten a kg y metros.

El Excel añade **Datos de facturas PDF**, con documento, página, fila, referencia
y registro de valores anteriores. El texto completo y la evidencia original se
consultan dentro del TMS; no se copian al packing list. Retirar un PDF revierte sus
valores vigentes, respetando los cambios manuales posteriores y las revisiones
de otros PDF. Los PDF duplicados por contenido no se vuelven a cargar. Cada página
indica si se leyó por texto u OCR; los escaneos tienen vista previa y confianza
OCR. Las páginas en blanco se identifican sin marcarlas como errores. La lectura
incompleta o de baja confianza requiere revisión visual expresa antes de la
descarga final; no inventa campos. Un documento sin referencias reconocidas
conserva su texto para consulta y permite completar la planilla manualmente.
Los encabezados de tablas OCR se reconstruyen a partir de las palabras y posiciones.

Los PDF protegidos solicitan contraseña, con reintento y cancelación; la contraseña
no se conserva en documentos, registros ni Excel. El usuario puede cancelar la
lectura y la revocación de acceso detiene el proceso. Se mantienen los límites
de 50 MB por documento y 300 páginas. Un archivo corrupto se rechaza de forma
explícita; un escaneo ilegible o un idioma OCR no admitido puede necesitar una
copia más nítida o revisión manual. La extracción de texto no implica que cualquier
diseño arbitrario pueda convertirse automáticamente en las 23 columnas.

Las dimensiones se expresan en metros y los pesos en kg. El peso bruto y el volumen
se cuentan una sola vez por CAJAS_ID. Se rechazan totales físicos completos si hay
valores incompatibles para el mismo bulto. El peso bruto por línea se prorratea
según peso neto, y no según cantidades de artículos diferentes. Los totales
desconocidos quedan vacíos; nunca se sustituyen por cero.

Mientras existan faltantes solo puede descargarse un borrador, identificado en la
cabecera y con hoja Pendientes. La descarga final exige referencia, contenedor,
precinto, cliente y los datos de bulto/origen/clasificación/peso necesarios. EAN,
EU y referencias comerciales que no figuren en las fuentes pueden quedar vacíos.
No se infieren países, partidas arancelarias ni Incoterms.

Operaciones se retira del catálogo y no se vuelve a sembrar al recargar. La rutina
de mantenimiento aborta si hay perfiles con ese rol. Auditoría Jamaica se registra
en la matriz con permiso Ver y empieza habilitada solo para Admin. La página
independiente verifica sesión con Auth, perfil activo y permiso del rol desde la
base de datos; revalida al recuperar foco y cada 15 segundos. Oculta y limpia los
documentos locales al perder acceso. No modifica el análisis documental existente.

Pruebas sintéticas, sin escrituras de datos operativos:

```sh
node --test tests/packing-miami.test.cjs
node tests/packing-miami-ui.cjs
node --test tests/packing-pdf.test.cjs
node tests/packing-pdf-ui.cjs
node tests/packing-pdf-ocr-ui.cjs
TMS_TEST_ENGINE=webkit node tests/packing-pdf-ocr-ui.cjs
node tests/jamaica-access-ui.cjs
node tests/access-settings-ui.cjs
```

Las variables opcionales TMS_PACKING_REAL_SOURCE y TMS_PACKING_REAL_TEMPLATE
permiten probar archivos privados locales sin incorporarlos al repositorio.

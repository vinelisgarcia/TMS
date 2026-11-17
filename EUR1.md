# Verificador EUR.1

**Módulos → Verificador EUR.1 → Nuevo expediente.** Adjuntar el certificado PDF y
un ZIP o documentos PDF/XLSX individuales. Se conservan los originales privados,
incluido el ZIP, con huella SHA-256 e integridad CRC de cada entrada. Los documentos
se comparten entre usuarios con permiso de visualizar EUR.1. Admin tiene acceso
por defecto; Configuración permite conceder Ver, Editar e Importar a otros roles.

La identificación usa títulos y columnas del contenido. Los libros de factura y
datos EUR.1 pueden tener encabezados idénticos: en ese caso el nombre aporta una
pista y el tipo queda pendiente de confirmación. El usuario puede corregirlo y
revisar valores contra una página o celda del original. Se conserva la autoría,
motivo y fuente de las revisiones. Una lectura fallida o un campo obligatorio
ausente mantiene la conclusión incompleta, incluso si hay diferencias confirmadas.

## Comparación

- Los datos EUR.1 delimitan las referencias propuestas. Se agrupan referencias
  repetidas, conservando las líneas originales. PDF y Excel de la misma familia
  documental se contrastan; solo una representación participa en las sumas.
- Se conservan los HS completos como texto. La cobertura del certificado se
  compara por HS4; longitud y subpartida se distinguen de referencias ausentes.
  Las contradicciones explícitas de materiales se señalan para aclaración,
  sin proponer clasificaciones arancelarias.
- Los bultos se cuentan por identificador y por su contenido: embarque,
  íntegramente incluidos, mixtos y con alguna mercancía incluida. Los subtotales
  no son mercancías. Una coincidencia con los íntegros que deja fuera un mixto
  exige aclarar el criterio, sin imponer automáticamente otro número.
- El bruto del certificado se compara con las líneas propuestas. El peso físico
  de los bultos se suma una sola vez y se muestra por separado. Se recalculan los
  totales del propio documento y el resumen HS frente a la factura completa.
- Tolerancias explícitas: cantidades 0,000001; importes 0,01; pesos 0,001 kg.
  El redondeo del certificado usa ±0,5 × 10 elevado a menos sus decimales, más
  tolerancia técnica. Las representaciones PDF también consideran la precisión
  de sus líneas, conservando la diferencia exacta.

El resultado empieza por su conclusión y ofrece comprobaciones, diferencias,
dudas, mercancías, HS y bultos. El Excel incluye conciliación y evidencia por línea.
Los países y marcas CEE son evidencia documental; el país por sí solo no acredita
origen preferencial. El resultado no es certificación legal ni autorización de despacho.

Referencias de método: [modelo oficial del certificado](https://taxation-customs.ec.europa.eu/document/download/965a3317-e36f-43ca-b9d2-0fc2d9d41d9a_en)
y [países de la UE](https://european-union.europa.eu/principles-countries-history/eu-countries_en).

## Procesamiento y límites

La carga va directamente al almacenamiento privado de Supabase; los documentos
se procesan por etapas en el navegador, con PDF.js, OCR local de Tesseract y un
Web Worker para ZIP, XLSX y conciliación. No depende de una función Vercel con
límite de cuerpo o tiempo. La pestaña debe permanecer abierta mientras procesa.
Al volver se pueden reanudar las lecturas desde los originales conservados.
Los errores se muestran por archivo, con reintento y detención.

Límites: 50 MB por original/ZIP; ZIP de 100 entradas, 250 MB extraídos, relación
de expansión máxima 200 y profundidad 12. Se rechazan rutas inseguras, cifrado,
enlaces simbólicos, ZIP64, cabeceras inconsistentes y CRC incorrecto. No se escribe
el contenido del ZIP en el sistema de archivos. Los ZIP anidados se cargan por
separado. XLSX valida también la expansión real interna: hasta 100 MB, 150.000
celdas, 50.000 filas y 256 columnas; no ejecuta macros ni fórmulas. Solo lee los
resultados guardados de fórmulas y señala resultados ausentes.

Cada expediente admite 250 documentos, 40 MB de extracción y 10.000 líneas para
analizar. Cada extracción tiene un máximo de 12 MB y el análisis guardado 25 MB.
PDF admite hasta 300 páginas y OCR secuencial sin acumular imágenes de vista previa.
Un documento que exceda los límites debe dividirse; el error permanece visible.

Las revisiones se guardan con versión esperada. Una edición obsoleta no sobrescribe
otra sesión. Un cambio documental invalida el informe guardado. La sincronización
publica solo una revisión pequeña y comprueba cambios cada ocho segundos. Los
originales no se sobrescriben ni eliminan desde el cliente. El registro global de
actividad identifica cambios, visitas, consultas y generación del informe.

## Activación y validación

Aplicar primero `20261116195535_system_activity_audit.sql` y después
`20261116200546_eur1_verifier.sql`. La segunda crea tablas aisladas, permisos y
el bucket privado `tms-eur1-originals`; no cambia el stock ni la planificación.
Publicar frontend y archivos vendor juntos. `tests/eur1-rollback.sql` valida
la activación con una transacción que termina en ROLLBACK.

Pruebas locales:

```sh
node --test tests/eur1-domain.test.cjs tests/eur1-archive.test.cjs
node tests/eur1-db.cjs
node tests/eur1-ui.cjs
TMS_TEST_ENGINE=webkit node tests/eur1-ui.cjs
node tests/packing-pdf-ocr-ui.cjs
```

Las pruebas de navegador cubren PDF + ZIP con nombres acentuados, XLSX con
resultado guardado de fórmula, clasificación, análisis, Excel, evidencia, revisión
simultánea obsoleta, recuperación, lectura fallida, permisos y teléfono. Las de
PostgreSQL usan identidades ficticias y una base efímera. Las dependencias de
prueba PGlite 0.5.16 y fflate 0.8.2 se instalan fuera del sitio.

Se revisó además, sin subirlo a producción, el formato del embarque H260958
previamente facilitado: datos EUR.1 de 35 líneas, factura de 105, packing PDF y
XLSX de 107 líneas y resumen HS. Ambas representaciones del packing coincidieron
sin duplicar mercancía; se identificaron 25 bultos. Al no incluir un certificado
en esa prueba, la conclusión permaneció incompleta.

**Validación de aceptación pendiente:** no están disponibles `EUR1_2615641.pdf`
ni el ZIP H261096 / factura ESF3-1000652-2026. No se afirma haber obtenido sus
resultados. El código no contiene sus referencias, importes ni cantidades de
aceptación. Los formatos ambiguos o las lecturas OCR dudosas requieren revisión
contra el original antes de completar la verificación.

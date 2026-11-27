# Corrector XML fiscal SAP

Facturación contiene Herramientas fiscales, con secciones separadas de crédito y débito e historial. El permiso `fiscalXml` se asigna desde Configuración; inicialmente solo el administrador tiene acceso. Los usuarios deben tener permiso de lectura para consultar el módulo y de edición o importación para generar una corrección. El servidor vuelve a verificar la sesión y los permisos y guarda los documentos en tablas independientes con RLS; no utiliza servicios de inteligencia artificial.

Cada tipo comienza **deshabilitado para generación**. Un administrador puede registrar un par de XML (original y corregido) únicamente tras comprobar que el corregido fue aceptado por la DGII. La activación es independiente para cada tipo. El par se revisa estructuralmente para impedir que cambie otros campos. Si una plantilla demuestra la incorporación de un campo ausente, solo se permite insertarlo cuando la ubicación y secuencia de elementos coinciden con el original de referencia. Si falta una plantilla o existe cualquier ambigüedad, la generación se bloquea.

El procesamiento admite XML UTF-8 de hasta 2 MB, rechaza DTD, entidades, firmas digitales, XML mal formado y campos fiscales repetidos. Solo se modifican NCF, fecha y, si se autoriza expresamente, referencia SAP. Se compara la estructura completa y el historial conserva original, resultado, autor, fecha, cambios y validaciones. La validación es estructural; no representa aceptación de la DGII. Si no hay XSD configurado, se indica que no se ha validado el esquema.

## Orden de despliegue

1. Ejecutar las pruebas locales abajo y revisar los permisos y la migración aditiva `supabase/migrations/20261127162536_fiscal_xml.sql`. No incorpora plantillas ni datos ficticios.
2. Aplicar únicamente esa migración al proyecto Supabase del TMS; comprobar las funciones, concesiones y RLS.
3. Publicar `supabase/functions/fiscal-xml` con verificación JWT, incluyendo `fiscal-xml-domain.js`. La copia debe ser idéntica al archivo raíz.
4. Publicar el frontend y verificar que los dos tipos aparecen pendientes y que usuarios sin permiso no acceden.
5. Activar cada tipo solo con su par real: original emitido por SAP y corregido aceptado por la DGII; comprobar el resultado real antes de su uso operativo.

## Pruebas

```sh
npm install --prefix /tmp/tms-fiscal-validation --no-save @electric-sql/pglite @xmldom/xmldom@0.9.12
node tests/fiscal-xml-domain-ui.cjs
node tests/fiscal-xml-ui.cjs
node tests/fiscal-xml-edge.cjs
node tests/fiscal-xml-db.cjs
npx --yes deno check --node-modules-dir=none supabase/functions/fiscal-xml/index.ts
cmp fiscal-xml-domain.js supabase/functions/fiscal-xml/fiscal-xml-domain.js
git diff --check
```

Las pruebas de navegador utilizan `TMS_PLAYWRIGHT` o la instalación local habitual del proyecto. Las pruebas de servidor y base de datos utilizan datos sintéticos, una API simulada y PostgreSQL efímero; no escriben en producción. `TMS_FISCAL_RUNTIME` permite cambiar la carpeta de dependencias temporales.

La migración `20261127163834_fiscal_dgii_acceptance.sql` corrige el criterio de aceptación. Conserva los ejemplos existentes pero desactiva las plantillas anteriores hasta que se confirme su aceptación por la DGII.

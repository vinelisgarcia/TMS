# Actividad del sistema

El administrador abre **Módulos → Actividad del sistema**. Puede filtrar por
usuario, período, cambios o accesos, y seleccionar una visita para consultar su
actividad. La paginación conserva el historial anterior. Las visitas se agrupan
por la sesión autenticada; las pestañas simultáneas de una misma sesión comparten
visita. Después de 30 minutos sin actividad se inicia otra visita. El cierre
explícito se registra; cerrar el navegador conserva la última actividad observada,
sin inventar una hora exacta de salida.

Los cambios se capturan en PostgreSQL dentro de la transacción que los guarda.
El buffer privado se vacía antes del commit; cancela sustituciones completas de
snapshots y modificaciones intermedias que terminan sin diferencia. Un rollback
no genera historial ni aviso. Se guardan usuario, módulo, fecha, operación,
secciones, identificadores y cantidades de registros afectados. Los textos de
notas, credenciales y borradores no se copian al registro de auditoría. Los
seguimientos y movimientos de inventario conservan además su detalle propio.

Las consultas de módulos, aperturas de detalles y generación de informes son
acciones del navegador, separadas de los cambios confirmados por el servidor.
No representan grabación de pantalla, pulsaciones ni una garantía de descarga
guardada en el dispositivo. Jamaica registra su acceso e informes también.

Los avisos de cambios son breves, no bloquean el trabajo y se agrupan por
transacción y módulo. Cada usuario solo recibe cambios de módulos que puede ver.
El historial de visitas y acciones es exclusivo de administradores activos. Los
clientes no pueden insertar, modificar ni eliminar registros directamente.

Aplicar `supabase/migrations/20261116195535_system_activity_audit.sql` antes de
publicar el frontend. Es aditiva: no cambia inventarios, rutas ni roles existentes.
El registro empieza al activarse; no reconstruye actividades anteriores.

Validación: `node tests/activity-audit-db.cjs`,
`node tests/activity-audit-ui.cjs` y
`TMS_TEST_ENGINE=webkit node tests/activity-audit-ui.cjs`. La prueba SQL de
producción `tests/activity-audit-rollback.sql` termina en ROLLBACK.

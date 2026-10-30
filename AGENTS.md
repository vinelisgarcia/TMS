# Instrucciones para trabajar en TMS

## Preferencia permanente del usuario

Al terminar cualquier modificación acordada con el usuario, realiza automáticamente
el commit y el push a la rama de trabajo, sin pedir una confirmación adicional.
Esta regla también se aplica a cambios de documentación e instrucciones y sigue
vigente hasta que el usuario indique expresamente lo contrario.

Antes del commit, ejecuta las verificaciones pertinentes y conserva los cambios
ajenos a la tarea. Informa la rama y el identificador del commit después de
comprobar el push. No incluyas credenciales ni utilices force push.

El commit y push no implican por sí solos fusionar en `main`, ejecutar migraciones
en producción ni desplegar. Para esas acciones, sigue la autorización y los
prerrequisitos de la conversación actual. Consulta `DEPLOYMENT.md` para el orden
de aplicación de la migración de sincronización y del frontend.

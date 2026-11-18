# Instrucciones para trabajar en TMS

## Preferencia permanente del usuario

Al terminar cualquier modificación acordada con el usuario, realiza automáticamente
el commit y el push a la rama de trabajo, sin pedir una confirmación adicional.
Esta regla también se aplica a cambios de documentación e instrucciones y sigue
vigente hasta que el usuario indique expresamente lo contrario.

Antes del commit, ejecuta las verificaciones pertinentes y conserva los cambios
ajenos a la tarea. Informa la rama y el identificador del commit después de
comprobar el push. No incluyas credenciales ni utilices force push.

El usuario autoriza publicar automáticamente en producción cada cambio terminado,
sin esperar otro visto bueno. Tras ejecutar las verificaciones pertinentes, haz
commit, push y despliega por el procedimiento del proyecto; verifica el resultado e
informa cualquier fallo. Conserva los cambios ajenos y nunca uses force push.
Antes de aplicar migraciones, revisa su alcance y los prerrequisitos de
`DEPLOYMENT.md`; no ejecutes migraciones ajenas al cambio.

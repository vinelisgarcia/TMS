-- Clamp operational role permissions to the intended module boundaries.
-- Apply after 0006_strict_auth_role_sync.sql.

begin;

insert into public.tms_settings (clave, valor)
values (
  'role_permissions',
  '{
    "Admin": {
      "importar": {"ver": true, "editar": true, "importar": true},
      "dashboard": {"ver": true, "editar": true, "importar": false},
      "calendario": {"ver": true, "editar": true, "importar": false},
      "reportes": {"ver": true, "editar": true, "importar": false},
      "comercial": {"ver": true, "editar": true, "importar": false},
      "rutas": {"ver": true, "editar": true, "importar": false},
      "importaciones": {"ver": true, "editar": true, "importar": true},
      "materialesTransito": {"ver": true, "editar": true, "importar": true},
      "prioridades": {"ver": true, "editar": true, "importar": true},
      "incidenciasDespacho": {"ver": true, "editar": true, "importar": false},
      "configuracion": {"ver": true, "editar": true, "importar": true},
      "solicitudesAlmacen": {"ver": true, "editar": true, "importar": false},
      "almacen": {"ver": true, "editar": true, "importar": false}
    },
    "Operaciones": {
      "importar": {"ver": true, "editar": true, "importar": true},
      "dashboard": {"ver": true, "editar": false, "importar": false},
      "calendario": {"ver": true, "editar": true, "importar": false},
      "reportes": {"ver": true, "editar": true, "importar": false},
      "comercial": {"ver": true, "editar": false, "importar": false},
      "rutas": {"ver": true, "editar": true, "importar": false},
      "importaciones": {"ver": true, "editar": true, "importar": true},
      "materialesTransito": {"ver": true, "editar": true, "importar": true},
      "prioridades": {"ver": false, "editar": false, "importar": false},
      "incidenciasDespacho": {"ver": false, "editar": false, "importar": false},
      "configuracion": {"ver": false, "editar": false, "importar": false},
      "solicitudesAlmacen": {"ver": true, "editar": true, "importar": false},
      "almacen": {"ver": true, "editar": true, "importar": false}
    },
    "Comercial": {
      "importar": {"ver": false, "editar": false, "importar": false},
      "dashboard": {"ver": false, "editar": false, "importar": false},
      "calendario": {"ver": false, "editar": false, "importar": false},
      "reportes": {"ver": false, "editar": false, "importar": false},
      "comercial": {"ver": true, "editar": false, "importar": false},
      "rutas": {"ver": false, "editar": false, "importar": false},
      "importaciones": {"ver": false, "editar": false, "importar": false},
      "materialesTransito": {"ver": false, "editar": false, "importar": false},
      "prioridades": {"ver": false, "editar": false, "importar": false},
      "incidenciasDespacho": {"ver": false, "editar": false, "importar": false},
      "configuracion": {"ver": false, "editar": false, "importar": false},
      "solicitudesAlmacen": {"ver": false, "editar": false, "importar": false},
      "almacen": {"ver": false, "editar": false, "importar": false}
    },
    "Almacén": {
      "importar": {"ver": false, "editar": false, "importar": false},
      "dashboard": {"ver": true, "editar": false, "importar": false},
      "calendario": {"ver": true, "editar": false, "importar": false},
      "reportes": {"ver": true, "editar": false, "importar": false},
      "comercial": {"ver": false, "editar": false, "importar": false},
      "rutas": {"ver": false, "editar": false, "importar": false},
      "importaciones": {"ver": false, "editar": false, "importar": false},
      "materialesTransito": {"ver": false, "editar": false, "importar": false},
      "prioridades": {"ver": false, "editar": false, "importar": false},
      "incidenciasDespacho": {"ver": false, "editar": false, "importar": false},
      "configuracion": {"ver": false, "editar": false, "importar": false},
      "solicitudesAlmacen": {"ver": true, "editar": true, "importar": false},
      "almacen": {"ver": true, "editar": true, "importar": false}
    },
    "Solo lectura": {
      "importar": {"ver": false, "editar": false, "importar": false},
      "dashboard": {"ver": true, "editar": false, "importar": false},
      "calendario": {"ver": true, "editar": false, "importar": false},
      "reportes": {"ver": true, "editar": false, "importar": false},
      "comercial": {"ver": true, "editar": false, "importar": false},
      "rutas": {"ver": false, "editar": false, "importar": false},
      "importaciones": {"ver": false, "editar": false, "importar": false},
      "materialesTransito": {"ver": false, "editar": false, "importar": false},
      "prioridades": {"ver": false, "editar": false, "importar": false},
      "incidenciasDespacho": {"ver": false, "editar": false, "importar": false},
      "configuracion": {"ver": false, "editar": false, "importar": false},
      "solicitudesAlmacen": {"ver": false, "editar": false, "importar": false},
      "almacen": {"ver": false, "editar": false, "importar": false}
    }
  }'::jsonb
)
on conflict (clave) do update
set valor = excluded.valor,
    updated_at = now();

commit;

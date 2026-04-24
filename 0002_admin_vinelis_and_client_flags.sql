begin;

alter table public.tms_client_configs
  add column if not exists no_miercoles boolean not null default false,
  add column if not exists no_ultima_semana boolean not null default false;

update public.tms_user_profiles
set nombre = 'Vinelis Garcia',
    email = 'vinelis.garcia@tms-alvarez.local',
    updated_at = now()
where rol = 'Admin'
  and (
    nombre = 'Manuel Oñate'
    or email = 'manuel.onate@tms-alvarez.local'
  );

insert into public.tms_user_profiles (nombre, email, rol, permisos_por_modulo)
values (
  'Vinelis Garcia',
  'vinelis.garcia@tms-alvarez.local',
  'Admin',
  jsonb_build_object(
    'importar', jsonb_build_object('ver', true, 'editar', true, 'importar', true),
    'dashboard', jsonb_build_object('ver', true, 'editar', true),
    'calendario', jsonb_build_object('ver', true, 'editar', true),
    'rutas', jsonb_build_object('ver', true, 'editar', true),
    'reportes', jsonb_build_object('ver', true, 'editar', true),
    'comercial', jsonb_build_object('ver', true, 'editar', true),
    'configuracion', jsonb_build_object('ver', true, 'editar', true, 'importar', true),
    'solicitudesAlmacen', jsonb_build_object('ver', true, 'editar', true),
    'almacen', jsonb_build_object('ver', true, 'editar', true)
  )
)
on conflict (email) do update
set nombre = excluded.nombre,
    rol = excluded.rol,
    permisos_por_modulo = excluded.permisos_por_modulo,
    updated_at = now();

commit;

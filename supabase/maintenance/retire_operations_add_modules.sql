-- Requested role retirement and defaults for two newly registered modules.
-- DML only. Does not change accounts, sessions, passwords or existing grants.
begin;
do $$ begin
  if exists (select 1 from public.tms_user_profiles where rol = 'Operaciones') then
    raise exception 'Reasigna primero los usuarios de Operaciones. No se cambió la configuración.';
  end if;
end $$;
update public.tms_settings s
set valor = (
  select jsonb_object_agg(role.key, role.value || jsonb_build_object(
    'auditoriaJamaica', coalesce(role.value -> 'auditoriaJamaica', jsonb_build_object('ver',role.key='Admin','editar',false,'importar',false)),
    'packingMiami', coalesce(role.value -> 'packingMiami', jsonb_build_object('ver',role.key='Admin','editar',role.key='Admin','importar',role.key='Admin'))
  ))
  from jsonb_each(s.valor - 'Operaciones') role
)
where s.clave = 'role_permissions';
commit;

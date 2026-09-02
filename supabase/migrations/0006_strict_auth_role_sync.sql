-- Make Supabase role checks match real Auth users and prevent inactive/unprofiled users
-- from being treated as elevated roles.
-- Apply after 0005_app_state_new_modules_policy.sql.

begin;

update public.tms_user_profiles p
set auth_user_id = u.id,
    updated_at = now()
from auth.users u
where p.auth_user_id is null
  and lower(p.email) = lower(u.email);

update public.tms_user_profiles p
set auth_user_id = u.id,
    updated_at = now()
from auth.users u
where lower(p.email) = lower(u.email)
  and p.auth_user_id is distinct from u.id;

create or replace function public.tms_current_role()
returns text
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(
    (
      select rol
      from public.tms_user_profiles
      where activo is not false
        and (
          auth_user_id = auth.uid()
          or lower(email) = lower(coalesce(auth.jwt() ->> 'email', ''))
        )
      order by case when auth_user_id = auth.uid() then 0 else 1 end, created_at asc
      limit 1
    ),
    'Solo lectura'
  );
$$;

create or replace function public.tms_can_edit(module_name text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.tms_current_role() = 'Admin'
    or coalesce(
      (
        select (valor -> public.tms_current_role() -> module_name ->> 'editar')::boolean
        from public.tms_settings
        where clave = 'role_permissions'
        limit 1
      ),
      false
    )
    or coalesce(
      (
        select (valor -> public.tms_current_role() -> module_name ->> 'importar')::boolean
        from public.tms_settings
        where clave = 'role_permissions'
        limit 1
      ),
      false
    );
$$;

grant execute on function public.tms_current_role() to authenticated;
grant execute on function public.tms_can_edit(text) to authenticated;

update public.tms_settings
set valor = jsonb_set(
    jsonb_set(
      jsonb_set(
        jsonb_set(
          jsonb_set(
            valor,
            '{Admin,importaciones}',
            coalesce(valor #> '{Admin,importaciones}', '{"ver":true,"editar":true,"importar":true}'::jsonb),
            true
          ),
          '{Admin,materialesTransito}',
          coalesce(valor #> '{Admin,materialesTransito}', '{"ver":true,"editar":true,"importar":true}'::jsonb),
          true
        ),
        '{Admin,prioridades}',
        coalesce(valor #> '{Admin,prioridades}', '{"ver":true,"editar":true,"importar":false}'::jsonb),
        true
      ),
      '{Admin,incidenciasDespacho}',
      coalesce(valor #> '{Admin,incidenciasDespacho}', '{"ver":true,"editar":true,"importar":false}'::jsonb),
      true
    ),
    '{Operaciones,materialesTransito}',
    coalesce(valor #> '{Operaciones,materialesTransito}', '{"ver":true,"editar":true,"importar":true}'::jsonb),
    true
  ),
  updated_at = now()
where clave = 'role_permissions';

commit;

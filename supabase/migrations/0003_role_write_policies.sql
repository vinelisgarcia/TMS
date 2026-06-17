-- Restrict browser writes by the operational role stored in tms_user_profiles.
-- Run this in Supabase SQL Editor after deploying the matching frontend changes.

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
      where auth_user_id = auth.uid()
         or lower(email) = lower(coalesce(auth.jwt() ->> 'email', ''))
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

create or replace function public.tms_guard_user_profile_self_write()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  if public.tms_current_role() = 'Admin' then
    return new;
  end if;

  if tg_op = 'INSERT' then
    if new.rol <> 'Solo lectura' or new.activo is false then
      raise exception 'Solo administradores pueden crear perfiles con rol elevado';
    end if;
    return new;
  end if;

  if tg_op = 'UPDATE' then
    if new.rol is distinct from old.rol
      or new.activo is distinct from old.activo
      or new.nombre is distinct from old.nombre
      or new.email is distinct from old.email
      or new.auth_user_id is distinct from old.auth_user_id
      or coalesce(new.permisos_por_modulo - '__account', '{}'::jsonb) is distinct from coalesce(old.permisos_por_modulo - '__account', '{}'::jsonb)
    then
      raise exception 'Solo administradores pueden cambiar roles, permisos o identidad del usuario';
    end if;
    return new;
  end if;

  return new;
end;
$$;

drop trigger if exists tms_guard_user_profile_self_write on public.tms_user_profiles;
create trigger tms_guard_user_profile_self_write
before insert or update on public.tms_user_profiles
for each row execute function public.tms_guard_user_profile_self_write();

drop policy if exists "authenticated users can write tms_user_profiles" on public.tms_user_profiles;
create policy "admins manage profiles and users update own access state"
on public.tms_user_profiles for all
to authenticated
using (
  public.tms_current_role() = 'Admin'
  or auth_user_id = auth.uid()
  or lower(email) = lower(coalesce(auth.jwt() ->> 'email', ''))
)
with check (
  public.tms_current_role() = 'Admin'
  or auth_user_id = auth.uid()
  or lower(email) = lower(coalesce(auth.jwt() ->> 'email', ''))
);

drop policy if exists "authenticated users can write tms_settings" on public.tms_settings;
create policy "admins and allowed roles can write settings"
on public.tms_settings for all
to authenticated
using (
  public.tms_current_role() = 'Admin'
  or (clave = 'warehouse' and public.tms_can_edit('almacen'))
  or (clave = 'app_state' and (public.tms_can_edit('importar') or public.tms_can_edit('solicitudesAlmacen') or public.tms_can_edit('almacen') or public.tms_can_edit('configuracion') or public.tms_can_edit('rutas')))
)
with check (
  public.tms_current_role() = 'Admin'
  or (clave = 'warehouse' and public.tms_can_edit('almacen'))
  or (clave = 'app_state' and (public.tms_can_edit('importar') or public.tms_can_edit('solicitudesAlmacen') or public.tms_can_edit('almacen') or public.tms_can_edit('configuracion') or public.tms_can_edit('rutas')))
);

drop policy if exists "authenticated users can write tms_route_configs" on public.tms_route_configs;
create policy "roles can write route configs"
on public.tms_route_configs for all
to authenticated
using (public.tms_current_role() = 'Admin' or public.tms_can_edit('configuracion') or public.tms_can_edit('rutas'))
with check (public.tms_current_role() = 'Admin' or public.tms_can_edit('configuracion') or public.tms_can_edit('rutas'));

drop policy if exists "authenticated users can write tms_client_configs" on public.tms_client_configs;
create policy "roles can write client configs"
on public.tms_client_configs for all
to authenticated
using (public.tms_current_role() = 'Admin' or public.tms_can_edit('configuracion') or public.tms_can_edit('rutas'))
with check (public.tms_current_role() = 'Admin' or public.tms_can_edit('configuracion') or public.tms_can_edit('rutas'));

drop policy if exists "authenticated users can write tms_order_lines" on public.tms_order_lines;
create policy "roles can write order lines"
on public.tms_order_lines for all
to authenticated
using (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('calendario') or public.tms_can_edit('rutas'))
with check (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('calendario') or public.tms_can_edit('rutas'));

drop policy if exists "authenticated users can write tms_route_cost_history" on public.tms_route_cost_history;
create policy "roles can write route cost history"
on public.tms_route_cost_history for all
to authenticated
using (public.tms_current_role() = 'Admin' or public.tms_can_edit('rutas'))
with check (public.tms_current_role() = 'Admin' or public.tms_can_edit('rutas'));

drop policy if exists "authenticated users can write tms_warehouse_plan" on public.tms_warehouse_plan;
create policy "roles can write warehouse plan"
on public.tms_warehouse_plan for all
to authenticated
using (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('solicitudesAlmacen'))
with check (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('solicitudesAlmacen'));

drop policy if exists "authenticated users can write tms_warehouse_control" on public.tms_warehouse_control;
create policy "roles can write warehouse control"
on public.tms_warehouse_control for all
to authenticated
using (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('almacen'))
with check (public.tms_current_role() = 'Admin' or public.tms_can_edit('importar') or public.tms_can_edit('almacen'));

drop policy if exists "authenticated users can write tms_warehouse_history" on public.tms_warehouse_history;
create policy "roles can write warehouse history"
on public.tms_warehouse_history for all
to authenticated
using (public.tms_current_role() = 'Admin' or public.tms_can_edit('solicitudesAlmacen') or public.tms_can_edit('almacen'))
with check (public.tms_current_role() = 'Admin' or public.tms_can_edit('solicitudesAlmacen') or public.tms_can_edit('almacen'));

-- Inventory reconciliation has its own snapshot, revision and invalidation row.
-- This migration seeds only an empty operational default, never user inventory.
-- Existing TMS role matrices and operational data remain unchanged.
begin;

create table if not exists public.tms_inventory_state (
  id boolean primary key default true check (id),
  payload jsonb not null default '{"version":1,"sources":{"ct":null,"ad":null},"cases":{},"settings":{"sameScope":true}}'::jsonb
    check (jsonb_typeof(payload) = 'object'),
  updated_at timestamptz not null default now(),
  updated_by uuid null references auth.users(id) on delete set null
);
create table if not exists public.tms_inventory_sync_state (
  id boolean primary key default true check (id),
  revision bigint not null default 0 check (revision >= 0),
  updated_at timestamptz not null default now()
);
do $$
begin
  if not exists (select 1 from public.tms_inventory_sync_state where id) then
    insert into public.tms_inventory_sync_state(id) values (true);
  end if;
  if not exists (select 1 from public.tms_inventory_state where id) then
    insert into public.tms_inventory_state(id) values (true);
  end if;
end;
$$;

alter table public.tms_inventory_state enable row level security;
alter table public.tms_inventory_sync_state enable row level security;
revoke all on public.tms_inventory_state from public, anon, authenticated;
revoke all on public.tms_inventory_sync_state from public, anon, authenticated;
grant select on public.tms_inventory_state, public.tms_inventory_sync_state to authenticated;
-- Administrative tools retain a maintenance path; statement triggers invalidate
-- inventory clients without changing the TMS operational revision.
grant select, insert, update, delete on public.tms_inventory_state to service_role;
grant select, update on public.tms_inventory_sync_state to service_role;

create or replace function public.tms_inventory_allowed(p_action text)
returns boolean
language sql stable security definer set search_path = public
as $$
  select auth.uid() is not null and coalesce((
    select p.rol = 'Admin' or coalesce((
      select s.valor #>> array[p.rol, 'inventario', p_action] = 'true'
      from public.tms_settings s where s.clave = 'role_permissions'
    ), false)
    from public.tms_user_profiles p
    where p.activo is true and (
      p.auth_user_id = auth.uid()
      or lower(p.email) = lower(coalesce(auth.jwt() ->> 'email', ''))
    )
    order by case when p.auth_user_id = auth.uid() then 0 else 1 end, p.created_at, p.id
    limit 1
  ), false)
  and p_action in ('ver', 'editar', 'importar');
$$;
revoke all on function public.tms_inventory_allowed(text) from public, anon;
grant execute on function public.tms_inventory_allowed(text) to authenticated;

drop policy if exists "inventory readers access state" on public.tms_inventory_state;
create policy "inventory readers access state" on public.tms_inventory_state
for select to authenticated using (public.tms_inventory_allowed('ver'));
drop policy if exists "inventory readers access revision" on public.tms_inventory_sync_state;
create policy "inventory readers access revision" on public.tms_inventory_sync_state
for select to authenticated using (public.tms_inventory_allowed('ver'));

create or replace function public.tms_inventory_guard_write()
returns trigger
language plpgsql set search_path = public
as $$
begin
  if coalesce(current_setting('tms.inventory_write', true), '') <> 'on'
    and (
      coalesce(auth.jwt() ->> 'role', '') = 'authenticated'
      or current_user = 'authenticated'
      or (auth.uid() is not null and coalesce(auth.jwt() ->> 'role', '') <> 'service_role')
    )
  then
    raise exception using errcode = '42501', message = 'Guarda el inventario mediante sincronización con control de versión.';
  end if;
  return null;
end;
$$;
revoke all on function public.tms_inventory_guard_write() from public, anon, authenticated;

create or replace function public.tms_inventory_stamp()
returns trigger
language plpgsql set search_path = public
as $$
begin
  new.updated_at := clock_timestamp();
  new.updated_by := auth.uid();
  return new;
end;
$$;
revoke all on function public.tms_inventory_stamp() from public, anon, authenticated;

create or replace function public.tms_inventory_maintenance_revision()
returns trigger
language plpgsql security definer set search_path = public
as $$
begin
  if coalesce(current_setting('tms.inventory_write', true), '') <> 'on' then
    update public.tms_inventory_sync_state
    set revision = revision + 1, updated_at = clock_timestamp() where id;
  end if;
  return null;
end;
$$;
revoke all on function public.tms_inventory_maintenance_revision() from public, anon, authenticated;

drop trigger if exists tms_inventory_guard on public.tms_inventory_state;
create trigger tms_inventory_guard before insert or update or delete or truncate
on public.tms_inventory_state for each statement execute function public.tms_inventory_guard_write();
drop trigger if exists tms_inventory_stamp on public.tms_inventory_state;
create trigger tms_inventory_stamp before insert or update
on public.tms_inventory_state for each row execute function public.tms_inventory_stamp();
drop trigger if exists tms_inventory_maintenance_revision on public.tms_inventory_state;
create trigger tms_inventory_maintenance_revision after insert or update or delete or truncate
on public.tms_inventory_state for each statement execute function public.tms_inventory_maintenance_revision();

create or replace function public.tms_inventory_get_snapshot(p_known_revision bigint default null)
returns jsonb
language plpgsql stable security definer set search_path = public
as $$
begin
  if not public.tms_inventory_allowed('ver') then
    raise exception using errcode = '42501', message = 'Tu perfil no tiene acceso al inventario.';
  end if;
  -- One statement provides a consistent revision and full JSON state. Nested
  -- material arrays are not truncated by the Data API's default row-page size.
  return (
    select case when s.revision = p_known_revision then
      jsonb_build_object('revision', s.revision, 'unchanged', true)
    else jsonb_build_object('revision', s.revision, 'tables', jsonb_build_object(
      'tms_inventory_state', coalesce((
        select jsonb_agg(jsonb_build_object('id', t.id, 'payload', t.payload))
        from public.tms_inventory_state t
      ), '[]'::jsonb)
    )) end
    from public.tms_inventory_sync_state s where s.id
  );
end;
$$;
revoke all on function public.tms_inventory_get_snapshot(bigint) from public, anon;
grant execute on function public.tms_inventory_get_snapshot(bigint) to authenticated;

create or replace function public.tms_inventory_save_snapshot(p_expected_revision bigint, p_payload jsonb)
returns jsonb
language plpgsql security definer set search_path = public
as $$
declare
  v_revision bigint;
  v_rows jsonb;
  v_row jsonb;
  v_payload jsonb;
  v_existing jsonb;
  v_previous_write text := coalesce(current_setting('tms.inventory_write', true), '');
begin
  if not (public.tms_inventory_allowed('editar') or public.tms_inventory_allowed('importar')) then
    raise exception using errcode = '42501', message = 'Tu perfil no puede modificar el inventario.';
  end if;
  if p_expected_revision is null or p_expected_revision < 0
    or jsonb_typeof(p_payload) is distinct from 'object'
    or exists (select 1 from jsonb_object_keys(p_payload) as keys(k) where k <> 'tms_inventory_state')
  then
    raise exception using errcode = '22023', message = 'Falta la revisión o el contenido del inventario no es válido.';
  end if;

  select revision into v_revision from public.tms_inventory_sync_state where id for update;
  if v_revision is distinct from p_expected_revision then
    raise exception using errcode = '40001',
      message = 'Otra sesión actualizó el inventario. Conserva tus cambios y revisa la versión compartida.',
      detail = jsonb_build_object('expected_revision', p_expected_revision, 'current_revision', v_revision)::text;
  end if;
  if p_payload = '{}'::jsonb then
    return jsonb_build_object('revision', v_revision, 'changed', false);
  end if;
  v_rows := p_payload -> 'tms_inventory_state';
  if jsonb_typeof(v_rows) is distinct from 'array' or jsonb_array_length(v_rows) <> 1 then
    raise exception using errcode = '22023', message = 'El inventario debe contener una única copia de estado.';
  end if;
  v_row := v_rows -> 0;
  if jsonb_typeof(v_row) is distinct from 'object'
    or v_row -> 'id' is distinct from 'true'::jsonb
    or exists (select 1 from jsonb_object_keys(v_row) as keys(k) where k not in ('id', 'payload'))
  then
    raise exception using errcode = '22023', message = 'La copia de inventario no es válida.';
  end if;
  v_payload := v_row -> 'payload';
  if jsonb_typeof(v_payload) is distinct from 'object'
    or v_payload -> 'version' is distinct from '1'::jsonb
    or jsonb_typeof(v_payload -> 'sources') is distinct from 'object'
    or jsonb_typeof(v_payload -> 'cases') is distinct from 'object'
    or jsonb_typeof(v_payload -> 'settings') is distinct from 'object'
    or coalesce(jsonb_typeof(v_payload #> '{sources,ct}'), '') not in ('object', 'null')
    or coalesce(jsonb_typeof(v_payload #> '{sources,ad}'), '') not in ('object', 'null')
  then
    raise exception using errcode = '22023', message = 'La estructura de conciliación de inventario no es válida.';
  end if;

  select payload into v_existing from public.tms_inventory_state where id;
  if not found then
    raise exception using errcode = '23503', message = 'Falta la copia compartida de inventario.';
  end if;
  if v_existing = v_payload then
    return jsonb_build_object('revision', v_revision, 'changed', false);
  end if;

  perform set_config('tms.inventory_write', 'on', true);
  update public.tms_inventory_state set payload = v_payload where id;
  update public.tms_inventory_sync_state
  set revision = revision + 1, updated_at = clock_timestamp()
  where id returning revision into v_revision;
  perform set_config('tms.inventory_write', v_previous_write, true);
  return jsonb_build_object('revision', v_revision, 'changed', true);
end;
$$;
revoke all on function public.tms_inventory_save_snapshot(bigint, jsonb) from public, anon;
grant execute on function public.tms_inventory_save_snapshot(bigint, jsonb) to authenticated;

-- Invalidation contains only a revision; the confidential inventory payload is
-- fetched separately after its own access check and is never broadcast.
do $$
begin
  if exists (select 1 from pg_publication where pubname = 'supabase_realtime')
    and not exists (
      select 1 from pg_publication_tables
      where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'tms_inventory_sync_state'
    )
  then
    alter publication supabase_realtime add table public.tms_inventory_sync_state;
  end if;
end;
$$;

commit;

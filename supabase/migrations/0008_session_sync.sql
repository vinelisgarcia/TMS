-- Versioned, transactional cloud synchronization. No operational rows are seeded
-- or replaced by this migration. Deploy together with the session-aware client.
-- The existing numbered migration convention is retained (CLI unavailable here).
begin;

create table if not exists public.tms_sync_state (
  id boolean primary key default true check (id),
  revision bigint not null default 0 check (revision >= 0),
  updated_at timestamptz not null default now()
);
insert into public.tms_sync_state (id) values (true) on conflict (id) do nothing;
alter table public.tms_sync_state enable row level security;
revoke all on public.tms_sync_state from public, anon, authenticated;
grant select on public.tms_sync_state to authenticated;
drop policy if exists "active TMS users read sync revision" on public.tms_sync_state;
create policy "active TMS users read sync revision"
on public.tms_sync_state for select to authenticated
using (
  auth.uid() is not null and exists (
    select 1 from public.tms_user_profiles p
    where p.activo is true and (
      p.auth_user_id = auth.uid()
      or lower(p.email) = lower(coalesce(auth.jwt() ->> 'email', ''))
    )
  )
);

-- RPC calls have one statement snapshot, including their revision and rows.
-- JSON aggregation deliberately avoids the Data API's default row-page limit.
create or replace function public.tms_get_snapshot(p_known_revision bigint default null)
returns jsonb
language plpgsql stable security definer set search_path = public
as $$
begin
  if auth.uid() is null or not exists (
    select 1 from public.tms_user_profiles p
    where p.activo is true and (
      p.auth_user_id = auth.uid()
      or lower(p.email) = lower(coalesce(auth.jwt() ->> 'email', ''))
    )
  ) then
    raise exception using errcode = '42501', message = 'Necesitas un perfil TMS activo para sincronizar.';
  end if;

  return (
    select case when s.revision = p_known_revision then
      jsonb_build_object('revision', s.revision, 'unchanged', true)
    else jsonb_build_object('revision', s.revision, 'tables', jsonb_build_object(
      'tms_order_lines', coalesce((select jsonb_agg(to_jsonb(t) order by t.created_at, t.id) from public.tms_order_lines t), '[]'::jsonb),
      'tms_client_configs', coalesce((select jsonb_agg(to_jsonb(t) order by t.codigo_cliente, t.id) from public.tms_client_configs t), '[]'::jsonb),
      'tms_route_configs', coalesce((select jsonb_agg(to_jsonb(t) order by t.nombre, t.id) from public.tms_route_configs t), '[]'::jsonb),
      'tms_route_cost_history', coalesce((select jsonb_agg(to_jsonb(t) order by t.fecha, t.id) from public.tms_route_cost_history t), '[]'::jsonb),
      'tms_warehouse_plan', coalesce((select jsonb_agg(to_jsonb(t) order by t.fecha, t.id) from public.tms_warehouse_plan t), '[]'::jsonb),
      'tms_warehouse_control', coalesce((select jsonb_agg(to_jsonb(t) order by t.fecha, t.id) from public.tms_warehouse_control t), '[]'::jsonb),
      'tms_warehouse_history', coalesce((select jsonb_agg(to_jsonb(t) order by t.fecha, t.id) from public.tms_warehouse_history t), '[]'::jsonb),
      'tms_user_profiles', coalesce((select jsonb_agg(to_jsonb(t) order by t.created_at, t.id) from public.tms_user_profiles t), '[]'::jsonb),
      'tms_settings', coalesce((select jsonb_agg(to_jsonb(t) order by t.clave, t.id) from public.tms_settings t), '[]'::jsonb)
    )) end
    from public.tms_sync_state s where s.id
  );
end;
$$;
revoke all on function public.tms_get_snapshot(bigint) from public, anon;
grant execute on function public.tms_get_snapshot(bigint) to authenticated;

-- Only the owner-called transactional RPC may invoke the old replacement helper.
-- This also makes already-open clients fail safely instead of overwriting data.
revoke all on function public.tms_replace_table(text, jsonb) from public, anon, authenticated;

-- Compare semantic rows, including duplicate multiplicity, ignoring generated
-- IDs/timestamps. Route IDs are preserved because other tables reference them.
-- Project the stored rows onto supplied fields so DB defaults do not cause
-- unnecessary full-table rewrites. This helper is not an exposed client API.
create or replace function public.tms_sync_rows_match(p_table_name text, p_rows jsonb)
returns boolean
language plpgsql stable security definer set search_path = public
as $$
declare
  v_expected jsonb;
  v_actual jsonb;
  v_keys text[];
begin
  if p_table_name not in (
    'tms_user_profiles', 'tms_route_configs', 'tms_client_configs',
    'tms_order_lines', 'tms_route_cost_history', 'tms_warehouse_plan',
    'tms_warehouse_control', 'tms_warehouse_history'
  ) or jsonb_typeof(p_rows) is distinct from 'array' then
    raise exception using errcode = '22023', message = 'Tabla o filas no válidas para sincronizar.';
  end if;

  select coalesce(jsonb_agg(row_data order by row_data::text), '[]'::jsonb)
  into v_expected
  from (
    select case when p_table_name = 'tms_route_configs'
      then row_data - 'created_at' - 'updated_at'
      else row_data - 'id' - 'created_at' - 'updated_at' end as row_data
    from jsonb_array_elements(p_rows) as rows(row_data)
  ) normalized;
  select coalesce(array_agg(distinct field), '{}'::text[]) into v_keys
  from jsonb_array_elements(v_expected) as rows(row_data)
  cross join lateral jsonb_object_keys(row_data) as fields(field);

  execute format(
    'select coalesce(jsonb_agg(projected order by projected::text), ''[]''::jsonb) '
    || 'from (select coalesce((select jsonb_object_agg(k, to_jsonb(t) -> k) '
    || 'from unnest($1::text[]) as keys(k)), ''{}''::jsonb) as projected '
    || 'from public.%I t) rows', p_table_name
  ) into v_actual using v_keys;
  return v_actual = v_expected;
end;
$$;
revoke all on function public.tms_sync_rows_match(text, jsonb) from public, anon, authenticated;

-- Guard before each statement, including an UPDATE that would match zero rows.
-- Direct SQL/service maintenance remains possible; its changes invalidate clients
-- using the AFTER STATEMENT revision trigger below.
create or replace function public.tms_sync_guard_write()
returns trigger
language plpgsql set search_path = public
as $$
begin
  if coalesce(current_setting('tms.sync_write', true), '') <> 'on'
    and (
      coalesce(auth.role(), '') = 'authenticated'
      or current_user = 'authenticated'
      or (auth.uid() is not null and coalesce(auth.role(), '') <> 'service_role')
    )
  then
    raise exception using errcode = '42501', message = 'Actualiza el TMS: guarda mediante sincronización con control de versión.';
  end if;
  return null;
end;
$$;
revoke all on function public.tms_sync_guard_write() from public, anon, authenticated;

create or replace function public.tms_sync_maintenance_revision()
returns trigger
language plpgsql security definer set search_path = public
as $$
begin
  if coalesce(current_setting('tms.sync_write', true), '') <> 'on' then
    update public.tms_sync_state
    set revision = revision + 1, updated_at = clock_timestamp()
    where id;
  end if;
  return null;
end;
$$;
revoke all on function public.tms_sync_maintenance_revision() from public, anon, authenticated;

do $$
declare v_table text;
begin
  foreach v_table in array array[
    'tms_user_profiles', 'tms_route_configs', 'tms_client_configs',
    'tms_settings', 'tms_import_batches', 'tms_order_lines',
    'tms_route_cost_history', 'tms_warehouse_plan',
    'tms_warehouse_control', 'tms_warehouse_history'
  ] loop
    execute format('drop trigger if exists tms_sync_guard on public.%I', v_table);
    execute format('create trigger tms_sync_guard before insert or update or delete or truncate on public.%I for each statement execute function public.tms_sync_guard_write()', v_table);
    execute format('drop trigger if exists tms_sync_maintenance_revision on public.%I', v_table);
    execute format('create trigger tms_sync_maintenance_revision after insert or update or delete or truncate on public.%I for each statement execute function public.tms_sync_maintenance_revision()', v_table);
  end loop;
end;
$$;

create or replace function public.tms_save_snapshot(p_expected_revision bigint, p_payload jsonb)
returns jsonb
language plpgsql security definer set search_path = public
as $$
declare
  v_profile public.tms_user_profiles%rowtype;
  v_role text;
  v_permissions jsonb;
  v_revision bigint;
  v_table text;
  v_rows jsonb;
  v_row jsonb;
  v_modules text[];
  v_allowed boolean;
  v_changed boolean := false;
  v_changed_tables jsonb := '{}'::jsonb;
  v_previous_write text := coalesce(current_setting('tms.sync_write', true), '');
begin
  if auth.uid() is null then
    raise exception using errcode = '42501', message = 'Necesitas iniciar sesión para guardar.';
  end if;
  select p.* into v_profile from public.tms_user_profiles p
  where p.activo is true and (
    p.auth_user_id = auth.uid()
    or lower(p.email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  )
  order by case when p.auth_user_id = auth.uid() then 0 else 1 end, p.created_at, p.id
  limit 1;
  if not found then
    raise exception using errcode = '42501', message = 'Tu perfil TMS no está activo.';
  end if;
  v_role := v_profile.rol;
  select valor into v_permissions from public.tms_settings where clave = 'role_permissions';
  v_permissions := coalesce(v_permissions, '{}'::jsonb);

  if p_expected_revision is null or p_expected_revision < 0
    or jsonb_typeof(p_payload) is distinct from 'object'
  then
    raise exception using errcode = '22023', message = 'Falta la revisión o el contenido del guardado.';
  end if;

  -- All writers lock the same row before inspecting or changing operational data.
  select revision into v_revision from public.tms_sync_state where id for update;
  if v_revision is distinct from p_expected_revision then
    raise exception using errcode = '40001',
      message = 'Otra sesión guardó cambios. Tus cambios locales están conservados; revisa la versión actual antes de guardar.',
      detail = jsonb_build_object('expected_revision', p_expected_revision, 'current_revision', v_revision)::text;
  end if;

  -- Validate the complete request before the first write. The role map remains
  -- unchanged; these module boundaries mirror the existing operational policies.
  for v_table, v_rows in select key, value from jsonb_each(p_payload) loop
    if jsonb_typeof(v_rows) is distinct from 'array'
      or exists (select 1 from jsonb_array_elements(v_rows) r where jsonb_typeof(r) is distinct from 'object')
    then
      raise exception using errcode = '22023', message = 'Cada tabla debe contener un arreglo de filas.';
    end if;
    v_modules := case v_table
      when 'tms_route_configs' then array['configuracion', 'rutas']
      when 'tms_client_configs' then array['configuracion', 'rutas']
      when 'tms_order_lines' then array['importar', 'calendario', 'rutas']
      when 'tms_route_cost_history' then array['rutas']
      when 'tms_warehouse_plan' then array['importar', 'solicitudesAlmacen']
      when 'tms_warehouse_control' then array['importar', 'almacen']
      when 'tms_warehouse_history' then array['solicitudesAlmacen', 'almacen']
      else null end;

    if v_table = 'tms_settings' then
      if (select count(*) from jsonb_array_elements(v_rows)) <>
        (select count(distinct r ->> 'clave') from jsonb_array_elements(v_rows) r)
      then
        raise exception using errcode = '22023', message = 'Los ajustes contienen claves duplicadas o vacías.';
      end if;
      for v_row in select value from jsonb_array_elements(v_rows) loop
        if coalesce(v_row ->> 'clave', '') = '' or jsonb_typeof(v_row -> 'valor') is distinct from 'object' then
          raise exception using errcode = '22023', message = 'Ajuste no válido.';
        end if;
        v_modules := case v_row ->> 'clave'
          when 'warehouse' then array['almacen']
          when 'app_state' then array['importar', 'solicitudesAlmacen', 'almacen', 'configuracion', 'rutas', 'calendario', 'materialesTransito', 'prioridades', 'incidenciasDespacho']
          else '{}'::text[] end;
        select v_role = 'Admin' or exists (
          select 1 from unnest(v_modules) m
          where coalesce((v_permissions -> v_role -> m ->> 'editar')::boolean, false)
            or coalesce((v_permissions -> v_role -> m ->> 'importar')::boolean, false)
        ) into v_allowed;
        if not v_allowed then
          raise exception using errcode = '42501', message = 'Tu rol no puede guardar este ajuste.';
        end if;
        if not exists (select 1 from public.tms_settings s where s.clave = v_row ->> 'clave' and s.valor = v_row -> 'valor') then
          v_changed_tables := jsonb_set(v_changed_tables, array[v_table], coalesce(v_changed_tables -> v_table, '[]'::jsonb) || jsonb_build_array(v_row));
        end if;
      end loop;
    elsif v_table = 'tms_user_profiles' and v_role <> 'Admin' then
      if jsonb_array_length(v_rows) <> 1 then
        raise exception using errcode = '42501', message = 'Solo puedes actualizar el estado de tu propia cuenta.';
      end if;
      v_row := v_rows -> 0;
      if lower(coalesce(v_row ->> 'email', '')) <> lower(v_profile.email)
        or coalesce(v_row ->> 'nombre', '') <> v_profile.nombre
        or coalesce(v_row ->> 'rol', '') <> v_profile.rol
        or (v_row ->> 'activo')::boolean is distinct from v_profile.activo
        or (nullif(v_row ->> 'auth_user_id', '')::uuid is not null
          and nullif(v_row ->> 'auth_user_id', '')::uuid is distinct from coalesce(v_profile.auth_user_id, auth.uid()))
        or coalesce((v_row -> 'permisos_por_modulo') - '__account', '{}'::jsonb)
          is distinct from coalesce(v_profile.permisos_por_modulo - '__account', '{}'::jsonb)
        or jsonb_typeof(v_row #> '{permisos_por_modulo,__account}') is distinct from 'object'
      then
        raise exception using errcode = '42501', message = 'No puedes modificar roles, permisos ni identidad de usuarios.';
      end if;
      if (v_row #> '{permisos_por_modulo,__account}') is distinct from (v_profile.permisos_por_modulo -> '__account') then
        v_changed_tables := jsonb_set(v_changed_tables, array[v_table], v_rows);
      end if;
    else
      if v_table <> 'tms_user_profiles' and v_modules is null then
        raise exception using errcode = '22023', message = 'Tabla no permitida para sincronizar: ' || v_table;
      end if;
      select v_role = 'Admin' or exists (
        select 1 from unnest(coalesce(v_modules, '{}'::text[])) m
        where coalesce((v_permissions -> v_role -> m ->> 'editar')::boolean, false)
          or coalesce((v_permissions -> v_role -> m ->> 'importar')::boolean, false)
      ) into v_allowed;
      if not v_allowed then
        raise exception using errcode = '42501', message = 'Tu rol no puede guardar esta tabla.';
      end if;
      if not public.tms_sync_rows_match(v_table, v_rows) then
        v_changed_tables := jsonb_set(v_changed_tables, array[v_table], v_rows);
      end if;
    end if;
  end loop;

  if v_changed_tables = '{}'::jsonb then
    return jsonb_build_object('revision', v_revision, 'changed', false);
  end if;
  perform set_config('tms.sync_write', 'on', true);

  -- Replace routes before dependent rows, and profiles last so changing one's
  -- own role cannot change authorization partway through a transaction.
  for v_table in select unnest(array[
    'tms_settings', 'tms_route_configs', 'tms_client_configs', 'tms_order_lines',
    'tms_route_cost_history', 'tms_warehouse_plan', 'tms_warehouse_control',
    'tms_warehouse_history', 'tms_user_profiles'
  ]) loop
    if not (v_changed_tables ? v_table) then continue; end if;
    v_rows := v_changed_tables -> v_table;
    if v_table = 'tms_settings' then
      insert into public.tms_settings (clave, valor)
      select r ->> 'clave', r -> 'valor' from jsonb_array_elements(v_rows) r
      on conflict (clave) do update set valor = excluded.valor, updated_at = clock_timestamp();
    elsif v_table = 'tms_route_configs' then
      -- A blanket DELETE would null the route FKs of unchanged order/history
      -- tables. Preserve existing IDs and delete only routes actually removed.
      select coalesce(jsonb_agg(r || jsonb_build_object('id', coalesce(
        nullif(r ->> 'id', '')::uuid,
        (select c.id from public.tms_route_configs c where c.nombre = r ->> 'nombre'),
        gen_random_uuid()
      ))), '[]'::jsonb)
      into v_rows from jsonb_array_elements(v_rows) r;
      delete from public.tms_route_configs c
      where not exists (select 1 from jsonb_array_elements(v_rows) r where (r ->> 'id')::uuid = c.id);
      insert into public.tms_route_configs (id, nombre, zona, dias_operacion, coste_ruta, activa, observaciones)
      select id, nombre, coalesce(zona, ''), coalesce(dias_operacion, '{}'), coalesce(coste_ruta, 0), coalesce(activa, true), coalesce(observaciones, '')
      from jsonb_to_recordset(v_rows) as r(id uuid, nombre text, zona text, dias_operacion smallint[], coste_ruta numeric, activa boolean, observaciones text)
      on conflict (id) do update set nombre = excluded.nombre, zona = excluded.zona,
        dias_operacion = excluded.dias_operacion, coste_ruta = excluded.coste_ruta,
        activa = excluded.activa, observaciones = excluded.observaciones;
    elsif v_table = 'tms_user_profiles' and v_role <> 'Admin' then
      update public.tms_user_profiles
      set permisos_por_modulo = (permisos_por_modulo - '__account') ||
        jsonb_build_object('__account', v_rows #> '{0,permisos_por_modulo,__account}')
      where id = v_profile.id;
    else
      perform public.tms_replace_table(v_table, v_rows);
    end if;
    v_changed := true;
  end loop;

  update public.tms_sync_state
  set revision = revision + 1, updated_at = clock_timestamp()
  where id returning revision into v_revision;
  perform set_config('tms.sync_write', v_previous_write, true);
  return jsonb_build_object('revision', v_revision, 'changed', v_changed);
end;
$$;
revoke all on function public.tms_save_snapshot(bigint, jsonb) from public, anon;
grant execute on function public.tms_save_snapshot(bigint, jsonb) to authenticated;

-- Publish one small invalidation row rather than broadcasting operational data.
do $$
begin
  if exists (select 1 from pg_publication where pubname = 'supabase_realtime')
    and not exists (
      select 1 from pg_publication_tables
      where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'tms_sync_state'
    )
  then
    alter publication supabase_realtime add table public.tms_sync_state;
  end if;
end;
$$;

commit;

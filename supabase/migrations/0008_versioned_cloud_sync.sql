-- Apply after 0007. One atomic save, revision conflicts and a consistent unpaginated read.
-- Does not rewrite existing operational data. Deploy before the matching frontend.
begin;

create table if not exists public.tms_sync_state (
  id integer primary key check (id = 1),
  revision bigint not null default 0,
  updated_at timestamptz not null default now()
);
insert into public.tms_sync_state(id) values (1) on conflict do nothing;
alter table public.tms_sync_state enable row level security;
drop policy if exists "authenticated read revision" on public.tms_sync_state;
create policy "authenticated read revision" on public.tms_sync_state for select to authenticated using (auth.uid() is not null);
grant select on public.tms_sync_state to authenticated;
revoke insert, update, delete on public.tms_sync_state from anon, authenticated;

-- Legacy clients and direct table edits also invalidate every cached snapshot.
create or replace function public.tms_bump_sync_revision()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  if coalesce(current_setting('tms.snapshot_write', true), '') <> 'on' then
    update public.tms_sync_state set revision = revision + 1, updated_at = clock_timestamp() where id = 1;
  end if;
  return null;
end;
$$;
revoke all on function public.tms_bump_sync_revision() from public;
-- Old tabs must refresh rather than bypass the revision check with direct writes.
create or replace function public.tms_require_snapshot_write()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  if auth.role()='authenticated' and coalesce(current_setting('tms.snapshot_write',true),'') <> 'on' then
    raise exception 'TMS_CLIENT_OUTDATED: Actualiza TMS para guardar con sincronización segura.' using errcode='42501';
  end if;
  return null;
end;
$$;
revoke all on function public.tms_require_snapshot_write() from public;
do $$
declare t text;
begin
  foreach t in array array['tms_order_lines','tms_client_configs','tms_route_configs','tms_route_cost_history','tms_warehouse_plan','tms_warehouse_control','tms_warehouse_history','tms_user_profiles','tms_settings'] loop
    execute format('drop trigger if exists tms_sync_revision on public.%I', t);
    execute format('create trigger tms_sync_revision after insert or update or delete on public.%I for each statement execute function public.tms_bump_sync_revision()', t);
    execute format('drop trigger if exists tms_require_snapshot on public.%I', t);
    execute format('create trigger tms_require_snapshot before insert or update or delete on public.%I for each statement execute function public.tms_require_snapshot_write()', t);
  end loop;
end;
$$;

-- One statement uses one MVCC snapshot. Existing row-level read policies still apply.
create or replace function public.tms_read_snapshot()
returns jsonb language sql stable security invoker set search_path = public as $$
  select jsonb_build_object('revision', s.revision, 'tables', jsonb_build_object(
    'tms_order_lines', (select coalesce(jsonb_agg(t order by t.created_at,t.id),'[]') from public.tms_order_lines t),
    'tms_client_configs', (select coalesce(jsonb_agg(t order by t.codigo_cliente),'[]') from public.tms_client_configs t),
    'tms_route_configs', (select coalesce(jsonb_agg(t order by t.nombre),'[]') from public.tms_route_configs t),
    'tms_route_cost_history', (select coalesce(jsonb_agg(t order by t.fecha,t.id),'[]') from public.tms_route_cost_history t),
    'tms_warehouse_plan', (select coalesce(jsonb_agg(t order by t.fecha,t.id),'[]') from public.tms_warehouse_plan t),
    'tms_warehouse_control', (select coalesce(jsonb_agg(t order by t.fecha,t.id),'[]') from public.tms_warehouse_control t),
    'tms_warehouse_history', (select coalesce(jsonb_agg(t order by t.fecha,t.id),'[]') from public.tms_warehouse_history t),
    'tms_user_profiles', (select coalesce(jsonb_agg(t order by t.created_at,t.id),'[]') from public.tms_user_profiles t),
    'tms_settings', (select coalesce(jsonb_agg(t order by t.clave),'[]') from public.tms_settings t)
  )) from public.tms_sync_state s where s.id=1 and auth.uid() is not null;
$$;
revoke all on function public.tms_read_snapshot() from public;
grant execute on function public.tms_read_snapshot() to authenticated;

drop function if exists public.tms_save_snapshot(bigint,jsonb,jsonb);
create or replace function public.tms_save_snapshot(p_expected_revision bigint, p_tables jsonb, p_settings jsonb, p_account jsonb default null)
returns bigint language plpgsql security definer set search_path = public as $$
declare
  current_revision bigint;
  entry record;
  field record;
  setting jsonb;
  allowed_value jsonb;
  module_name text;
  admin_user boolean;
  account_field text;
begin
  if auth.uid() is null or not exists (
    select 1 from public.tms_user_profiles where activo and (auth_user_id=auth.uid() or lower(email)=lower(auth.jwt()->>'email'))
  ) then raise exception 'Sesión sin perfil activo' using errcode='42501'; end if;
  if jsonb_typeof(p_tables) is distinct from 'object' or jsonb_typeof(p_settings) is distinct from 'array' then
    raise exception 'Payload de sincronización inválido';
  end if;
  if p_account is not null then
    if jsonb_typeof(p_account) is distinct from 'object'
      or p_account - array['passwordChangeRequired','authNeedsConfirmation','passwordConfigured','accountStatus']::text[] <> '{}'::jsonb then
      raise exception 'Metadatos de cuenta inválidos: solo puedes actualizar el estado de tu propia cuenta' using errcode='42501';
    end if;
    foreach account_field in array array['passwordChangeRequired','authNeedsConfirmation','passwordConfigured'] loop
      if p_account ? account_field and jsonb_typeof(p_account->account_field) is distinct from 'boolean' then
        raise exception 'El estado % debe ser booleano',account_field;
      end if;
    end loop;
    if p_account ? 'accountStatus' and jsonb_typeof(p_account->'accountStatus') is distinct from 'string' then
      raise exception 'El estado de cuenta debe ser texto';
    end if;
  end if;
  select revision into current_revision from public.tms_sync_state where id=1 for update;
  if p_expected_revision is distinct from current_revision then
    raise exception 'TMS_CONFLICT: Otra sesión guardó cambios. Conserva tu respaldo y recarga la nube.' using errcode='40001';
  end if;
  admin_user := public.tms_current_role()='Admin';
  perform set_config('tms.snapshot_write','on',true);

  -- Reuse existing table-level role checks. No table can commit if any later write fails.
  for entry in select key,value from jsonb_each(p_tables) order by array_position(
    array['tms_route_configs','tms_client_configs','tms_order_lines','tms_route_cost_history','tms_warehouse_plan','tms_warehouse_control','tms_warehouse_history','tms_user_profiles'],key
  ) loop
    perform public.tms_replace_table(entry.key,entry.value);
  end loop;

  for setting in select value from jsonb_array_elements(p_settings) loop
    if setting->>'clave'='app_state' then
      if jsonb_typeof(setting->'valor') is distinct from 'object' then raise exception 'app_state inválido'; end if;
      allowed_value := '{}'::jsonb;
      for field in select key,value from jsonb_each(setting->'valor') loop
        -- Operational modules may change their own settings without replacing tariffs
        -- or import filenames owned by another module.
        if field.key='costSettings' and not admin_user and not public.tms_can_edit('configuracion') then
          if (public.tms_can_edit('calendario') or public.tms_can_edit('rutas')) and field.value ? 'tripHistory' then
            allowed_value := allowed_value || jsonb_build_object('costSettings',
              coalesce((select valor->'costSettings' from public.tms_settings where clave='app_state'),'{}'::jsonb)
              || jsonb_build_object('tripHistory',field.value->'tripHistory'));
          end if;
          continue;
        end if;
        if field.key='importFiles' and not admin_user and not public.tms_can_edit('importar') then
          if public.tms_can_edit('solicitudesAlmacen') or public.tms_can_edit('almacen') then
            allowed_value := allowed_value || jsonb_build_object('importFiles',
              coalesce((select valor->'importFiles' from public.tms_settings where clave='app_state'),'{}'::jsonb)
              || (field.value - 'plan' - 'control'));
          end if;
          continue;
        end if;
        module_name := case
          when field.key in ('planLoaded','planFecha','controlLoaded','controlFecha','importFiles','importShipments','controlHistory') then 'importar'
          when field.key in ('solicitudesLoaded','solicitudesFileName') then 'solicitudesAlmacen'
          when field.key in ('camionExtraEnabled','feriadosRD') then 'rutas'
          when field.key='transitMaterials' then 'materialesTransito'
          when field.key in ('dailyPriorities','prioritySpainTopics') then 'prioridades'
          when field.key='calendarNotes' then 'calendario'
          when field.key='dispatchIncidents' then 'incidenciasDespacho'
          when field.key in ('costSettings','visualTheme') then 'configuracion'
          else null end;
        if module_name is null then raise exception 'Campo de estado desconocido: %',field.key; end if;
        if admin_user or public.tms_can_edit(module_name) then
          allowed_value := allowed_value || jsonb_build_object(field.key,field.value);
        end if;
      end loop;
      if allowed_value <> '{}'::jsonb then
        insert into public.tms_settings(clave,valor) values ('app_state',allowed_value)
        on conflict(clave) do update set valor=public.tms_settings.valor || excluded.valor;
      end if;
    elsif setting->>'clave'='role_permissions' and admin_user then
      insert into public.tms_settings(clave,valor) values ('role_permissions',setting->'valor')
      on conflict(clave) do update set valor=excluded.valor;
    elsif setting->>'clave'='warehouse' and (admin_user or public.tms_can_edit('almacen') or public.tms_can_edit('solicitudesAlmacen')) then
      insert into public.tms_settings(clave,valor) values ('warehouse',setting->'valor')
      on conflict(clave) do update set valor=excluded.valor;
    else raise exception 'Tu rol no puede guardar el ajuste %',setting->>'clave' using errcode='42501';
    end if;
  end loop;
  -- Non-admin users may persist their password/account flags, never an identity,
  -- role, active state or module permission. The existing self-write guard remains
  -- enabled and validates this update; no admin-replacement bypass is set here.
  if p_account is not null then
    update public.tms_user_profiles p set permisos_por_modulo =
      coalesce(p.permisos_por_modulo,'{}'::jsonb) || jsonb_build_object('__account',
        coalesce(p.permisos_por_modulo->'__account','{}'::jsonb) || p_account)
    where p.id = (
      select own.id from public.tms_user_profiles own where own.activo
        and (own.auth_user_id=auth.uid() or lower(own.email)=lower(auth.jwt()->>'email'))
      order by case when own.auth_user_id=auth.uid() then 0 else 1 end,own.created_at,own.id
      limit 1
    );
  end if;
  update public.tms_sync_state set revision=revision+1,updated_at=clock_timestamp() where id=1 returning revision into current_revision;
  perform set_config('tms.snapshot_write','off',true);
  return current_revision;
end;
$$;
revoke all on function public.tms_save_snapshot(bigint,jsonb,jsonb,jsonb) from public;
grant execute on function public.tms_save_snapshot(bigint,jsonb,jsonb,jsonb) to authenticated;

commit;

-- Committed changes, authenticated visits and curated actions. No operational seeds.
begin;
create schema if not exists tms_private;
revoke all on schema tms_private from public, anon;
grant usage on schema tms_private to authenticated;

create table if not exists public.tms_audit_visits (
 id uuid primary key default gen_random_uuid(), user_id uuid not null,
 auth_session_id text not null, actor_name text not null, actor_email text not null,
 started_at timestamptz not null default clock_timestamp(), last_seen_at timestamptz not null default clock_timestamp(),
 ended_at timestamptz, end_reason text, device text not null default 'desktop'
);
create index if not exists tms_audit_visits_user_time on public.tms_audit_visits(user_id,started_at desc);
create index if not exists tms_audit_visits_session on public.tms_audit_visits(user_id,auth_session_id,last_seen_at desc);
create table if not exists public.tms_audit_events (
 id bigint generated always as identity primary key, visit_id uuid references public.tms_audit_visits(id),
 user_id uuid, actor_name text not null, actor_email text not null,
 created_at timestamptz not null default clock_timestamp(), module text not null,
 action text not null, summary text not null, details jsonb not null default '{}',
 is_change boolean not null default false, transaction_id bigint not null default txid_current()
);
create index if not exists tms_audit_events_time on public.tms_audit_events(created_at desc,id desc);
create index if not exists tms_audit_events_visit on public.tms_audit_events(visit_id,id);
create index if not exists tms_audit_events_user_time on public.tms_audit_events(user_id,created_at desc);
alter table public.tms_audit_visits enable row level security;
alter table public.tms_audit_events enable row level security;
revoke all on public.tms_audit_visits,public.tms_audit_events from public,anon,authenticated;
grant select on public.tms_audit_visits,public.tms_audit_events to authenticated;

create or replace function tms_private.profile()
returns jsonb language sql stable security definer set search_path = '' as $$
 select to_jsonb(p) from public.tms_user_profiles p
 where auth.uid() is not null and (p.auth_user_id=auth.uid() or lower(p.email)=lower(coalesce(auth.jwt()->>'email','')))
 order by (p.auth_user_id=auth.uid()) desc nulls last,p.created_at,p.id limit 1
$$;
create or replace function tms_private.allowed_modules()
returns text[] language plpgsql stable security definer set search_path = '' as $$
declare p jsonb:=tms_private.profile(); result text[];
begin
 if coalesce((p->>'activo')::boolean,false) is false then return '{}'; end if;
 if p->>'rol'='Admin' then return array['*']; end if;
 select coalesce(array_agg(e.key),'{}') into result from public.tms_settings s,
 lateral jsonb_each(coalesce(s.valor->(p->>'rol'),'{}')) e where s.clave='role_permissions' and e.value->>'ver'='true';
 return result;
end $$;
create or replace function tms_private.audit_admin()
returns boolean language sql stable security definer set search_path = '' as $$
 select coalesce((tms_private.profile()->>'activo')::boolean,false) and tms_private.profile()->>'rol'='Admin'
$$;
drop policy if exists "audit administrators read visits" on public.tms_audit_visits;
create policy "audit administrators read visits" on public.tms_audit_visits for select to authenticated using ((select tms_private.audit_admin()));
drop policy if exists "audit visible changes and own actions" on public.tms_audit_events;
create policy "audit visible changes and own actions" on public.tms_audit_events for select to authenticated using (
 (select tms_private.audit_admin()) or
 (is_change and (module=any((select tms_private.allowed_modules())::text[])))
);

create or replace function tms_private.visit(p_device text default 'desktop')
returns uuid language plpgsql security definer set search_path = '' as $$
declare p jsonb:=tms_private.profile(); v public.tms_audit_visits; sid text:=auth.jwt()->>'session_id';
begin
 if auth.uid() is null or not coalesce((p->>'activo')::boolean,false) then raise exception 'Acceso no autorizado.' using errcode='42501'; end if;
 if coalesce(sid,'')='' then raise exception 'La sesión no tiene identificador verificable.' using errcode='42501'; end if;
 perform pg_catalog.pg_advisory_xact_lock(pg_catalog.hashtextextended(auth.uid()::text||sid,0));
 select * into v from public.tms_audit_visits where user_id=auth.uid() and auth_session_id=sid and ended_at is null order by last_seen_at desc limit 1 for update;
 if v.id is not null and v.last_seen_at < clock_timestamp()-interval '30 minutes' then
  update public.tms_audit_visits set ended_at=last_seen_at,end_reason='inactividad' where id=v.id; v.id:=null;
 end if;
 if v.id is null then
  insert into public.tms_audit_visits(user_id,auth_session_id,actor_name,actor_email,device)
  values(auth.uid(),sid,coalesce(p->>'nombre','Usuario'),coalesce(p->>'email',''),case when p_device in ('mobile','tablet','desktop') then p_device else 'desktop' end) returning * into v;
  insert into public.tms_audit_events(visit_id,user_id,actor_name,actor_email,module,action,summary)
  values(v.id,auth.uid(),v.actor_name,v.actor_email,'auditoriaSistema','acceso','Accedió al TMS');
 else update public.tms_audit_visits set last_seen_at=clock_timestamp() where id=v.id; end if;
 return v.id;
end $$;
create or replace function tms_private.record_action(p_module text,p_action text,p_device text)
returns uuid language plpgsql security definer set search_path = '' as $$
declare v uuid; p jsonb; label text;
begin
 if p_action not in ('visita','ver_modulo','exportar_excel','exportar_pdf','consultar_detalle','cerrar_sesion') then raise exception 'Acción no admitida.' using errcode='22023'; end if;
 if p_module not in ('auditoriaSistema','importar','dashboard','calendario','reportes','comercial','rutas','importaciones','materialesTransito','prioridades','incidenciasDespacho','configuracion','solicitudesAlmacen','almacen','inventario','showroom','packingMiami','auditoriaJamaica','eur1') then raise exception 'Módulo no admitido.' using errcode='22023'; end if;
 if not ('*'=any(tms_private.allowed_modules()) or p_module=any(tms_private.allowed_modules()) or (p_module='auditoriaSistema' and p_action in ('visita','cerrar_sesion'))) then raise exception 'Sin permiso para este módulo.' using errcode='42501'; end if;
 v:=tms_private.visit(p_device); p:=tms_private.profile();
 if p_action='visita' then return v; end if;
 label:=case p_action when 'ver_modulo' then 'Abrió el módulo' when 'exportar_excel' then 'Descargó Excel' when 'exportar_pdf' then 'Descargó PDF' when 'consultar_detalle' then 'Consultó un detalle' else 'Cerró la sesión' end;
 -- Suppress duplicate navigation from wrappers or multiple render cycles.
 if p_action<>'ver_modulo' or not exists(select 1 from public.tms_audit_events where visit_id=v and action=p_action and module=p_module and created_at>clock_timestamp()-interval '3 seconds') then
 insert into public.tms_audit_events(visit_id,user_id,actor_name,actor_email,module,action,summary)
 values(v,auth.uid(),coalesce(p->>'nombre','Usuario'),coalesce(p->>'email',''),p_module,p_action,label); end if;
 if p_action='cerrar_sesion' then update public.tms_audit_visits set ended_at=clock_timestamp(),end_reason='cierre de sesión' where id=v; end if;
 return v;
end $$;
create or replace function public.tms_audit_action(p_module text,p_action text,p_device text default 'desktop')
returns uuid language sql security invoker set search_path = '' as $$select tms_private.record_action(p_module,p_action,p_device)$$;

-- A private transaction buffer cancels DELETE/INSERT snapshot replacements and
-- intermediate updates. It is emptied by a deferred trigger before commit.
-- Original payloads never become public audit records or persist after commit.
create unlogged table if not exists tms_private.audit_buffer (
 transaction_id bigint not null, table_name text not null, old_rows jsonb not null default '[]',new_rows jsonb not null default '[]',
 primary key(transaction_id,table_name)
);
revoke all on tms_private.audit_buffer from public,anon,authenticated;
create or replace function tms_private.audit_capture()
returns trigger language plpgsql security definer set search_path = '' as $$
declare olds jsonb:='[]'; news jsonb:='[]';
begin
 if tg_op='TRUNCATE' then execute format('select coalesce(jsonb_agg(to_jsonb(o)-array[''id'',''created_at'',''updated_at'',''updated_by'',''source_batch_id'']),''[]''::jsonb) from public.%I o',tg_table_name) into olds; end if;
 if tg_op in ('UPDATE','DELETE') then select coalesce(jsonb_agg(to_jsonb(o)-array['id','created_at','updated_at','updated_by','source_batch_id']),'[]') into olds from old_rows o; end if;
 if tg_op in ('UPDATE','INSERT') then select coalesce(jsonb_agg(to_jsonb(n)-array['id','created_at','updated_at','updated_by','source_batch_id']),'[]') into news from new_rows n; end if;
 if olds='[]' and news='[]' then return null; end if;
 insert into tms_private.audit_buffer(transaction_id,table_name,old_rows,new_rows) values(txid_current(),tg_table_name,olds,news)
 on conflict(transaction_id,table_name) do update set old_rows=tms_private.audit_buffer.old_rows||excluded.old_rows,new_rows=tms_private.audit_buffer.new_rows||excluded.new_rows;
 return null;
end $$;
create or replace function tms_private.emit_change(p_module text,p_summary text,p_details jsonb)
returns void language plpgsql security definer set search_path = '' as $$
declare p jsonb:=tms_private.profile(); v uuid;
begin
 select id into v from public.tms_audit_visits where user_id=auth.uid() and auth_session_id=auth.jwt()->>'session_id' and ended_at is null order by last_seen_at desc limit 1;
 -- A committed write can arrive before the browser has registered its visit.
 if v is null and auth.uid() is not null and coalesce((p->>'activo')::boolean,false) and coalesce(auth.jwt()->>'session_id','')<>'' then v:=tms_private.visit(); end if;
 insert into public.tms_audit_events(visit_id,user_id,actor_name,actor_email,module,action,summary,details,is_change)
 values(v,auth.uid(),coalesce(p->>'nombre','Sistema'),coalesce(p->>'email',''),p_module,'cambio_guardado',p_summary,p_details,true);
 if v is not null then update public.tms_audit_visits set last_seen_at=clock_timestamp() where id=v; end if;
end $$;
create or replace function tms_private.audit_flush()
returns trigger language plpgsql security definer set search_path = '' as $$
declare b tms_private.audit_buffer; olds jsonb; news jsonb; mod text; label text; keys jsonb; r jsonb; k text; before_value jsonb; after_value jsonb; changed_keys jsonb; identifiers jsonb;
begin
 select * into b from tms_private.audit_buffer where transaction_id=new.transaction_id and table_name=new.table_name;
 if not found then return null; end if;
 select coalesce(jsonb_agg(x),'[]') into olds from (select value x from jsonb_array_elements(b.old_rows) except all select value from jsonb_array_elements(b.new_rows)) s;
 select coalesce(jsonb_agg(x),'[]') into news from (select value x from jsonb_array_elements(b.new_rows) except all select value from jsonb_array_elements(b.old_rows)) s;
 delete from tms_private.audit_buffer where transaction_id=b.transaction_id and table_name=b.table_name;
 if olds='[]' and news='[]' then return null; end if;
 if b.table_name='tms_settings' then
  for k in select distinct value->>'clave' from jsonb_array_elements(olds||news) loop
   select value->'valor' into before_value from jsonb_array_elements(olds) where value->>'clave'=k limit 1;
   select value->'valor' into after_value from jsonb_array_elements(news) where value->>'clave'=k limit 1;
   if before_value is not distinct from after_value then continue; end if;
   if k='app_state' then
    for r in select to_jsonb(j) from (select distinct key from jsonb_object_keys(coalesce(before_value,'{}')||coalesce(after_value,'{}')) key) j loop
     label:=r->>'key';
     if before_value->label is not distinct from after_value->label then continue; end if;
     mod:=case label when 'transitMaterials' then 'materialesTransito' when 'dailyPriorities' then 'prioridades' when 'calendarNotes' then 'calendario' when 'dispatchIncidents' then 'incidenciasDespacho' when 'warehouseHistory' then 'almacen' when 'clientFlags' then 'comercial' else 'configuracion' end;
     perform tms_private.emit_change(mod,'Actualizó datos compartidos',jsonb_build_object('tabla',b.table_name,'seccion',label));
    end loop;
   else perform tms_private.emit_change('configuracion','Actualizó la configuración',jsonb_build_object('tabla',b.table_name,'seccion',k)); end if;
  end loop;
 elsif b.table_name in ('tms_inventory_state','tms_showroom_state') then
  before_value:=olds->0->'payload'; after_value:=news->0->'payload';
  select coalesce(jsonb_agg(key),'[]') into changed_keys from jsonb_object_keys(coalesce(before_value,'{}')||coalesce(after_value,'{}')) key where before_value->key is distinct from after_value->key;
  select coalesce(jsonb_agg(code),'[]') into identifiers from (
   select code from jsonb_object_keys(coalesce(before_value->'cases','{}')||coalesce(after_value->'cases','{}')) code
   where before_value->'cases'->code is distinct from after_value->'cases'->code limit 100
  ) s;
  mod:=case when b.table_name='tms_inventory_state' then 'inventario' else 'showroom' end;
  perform tms_private.emit_change(mod,'Actualizó inventario, conteos o seguimiento',jsonb_build_object('tabla',b.table_name,'secciones',changed_keys,'referencias',identifiers));
 else
  mod:=case b.table_name when 'tms_order_lines' then 'calendario' when 'tms_route_configs' then 'rutas' when 'tms_client_configs' then 'configuracion' when 'tms_user_profiles' then 'configuracion' when 'tms_route_cost_history' then 'rutas' when 'tms_import_batches' then 'importar' when 'tms_warehouse_plan' then 'solicitudesAlmacen' when 'tms_warehouse_control' then 'almacen' when 'tms_warehouse_history' then 'almacen' when 'tms_eur1_cases' then 'eur1' when 'tms_eur1_documents' then 'eur1' else 'configuracion' end;
  select coalesce(jsonb_agg(x),'[]') into identifiers from (select distinct coalesce(value->>'clave_unica',value->>'codigo_cliente',value->>'email',value->>'nombre',value->>'name',value->>'title') x from jsonb_array_elements(news||olds) limit 100) s where x is not null;
  perform tms_private.emit_change(mod,'Guardó cambios',jsonb_build_object('tabla',b.table_name,'registros_anteriores',jsonb_array_length(olds),'registros_nuevos',jsonb_array_length(news),'identificadores',identifiers));
 end if;
 return null;
end $$;
drop trigger if exists audit_flush_at_commit on tms_private.audit_buffer;
create constraint trigger audit_flush_at_commit after insert on tms_private.audit_buffer deferrable initially deferred for each row execute function tms_private.audit_flush();

create or replace function tms_private.attach_audit(p_table text)
returns void language plpgsql security definer set search_path = '' as $$
begin
 if p_table not in ('tms_settings','tms_order_lines','tms_route_configs','tms_client_configs','tms_user_profiles','tms_import_batches','tms_route_cost_history','tms_warehouse_plan','tms_warehouse_control','tms_warehouse_history','tms_inventory_state','tms_showroom_state','tms_eur1_cases','tms_eur1_documents') then raise exception 'Tabla no autorizada'; end if;
 if to_regclass('public.'||p_table) is null then return; end if;
 execute format('drop trigger if exists audit_capture_insert on public.%I',p_table);
 execute format('drop trigger if exists audit_capture_update on public.%I',p_table);
 execute format('drop trigger if exists audit_capture_delete on public.%I',p_table);
 execute format('drop trigger if exists audit_capture_truncate on public.%I',p_table);
 execute format('create trigger audit_capture_insert after insert on public.%I referencing new table as new_rows for each statement execute function tms_private.audit_capture()',p_table);
 execute format('create trigger audit_capture_update after update on public.%I referencing old table as old_rows new table as new_rows for each statement execute function tms_private.audit_capture()',p_table);
 execute format('create trigger audit_capture_delete after delete on public.%I referencing old table as old_rows for each statement execute function tms_private.audit_capture()',p_table);
 execute format('create trigger audit_capture_truncate before truncate on public.%I for each statement execute function tms_private.audit_capture()',p_table);
end $$;
do $$declare t text; begin foreach t in array array['tms_settings','tms_order_lines','tms_route_configs','tms_client_configs','tms_user_profiles','tms_import_batches','tms_route_cost_history','tms_warehouse_plan','tms_warehouse_control','tms_warehouse_history','tms_inventory_state','tms_showroom_state'] loop perform tms_private.attach_audit(t); end loop; end $$;

revoke all on all functions in schema tms_private from public,anon,authenticated;
grant execute on function tms_private.allowed_modules(),tms_private.audit_admin(),tms_private.record_action(text,text,text) to authenticated;
revoke all on function public.tms_audit_action(text,text,text) from public,anon;
grant execute on function public.tms_audit_action(text,text,text) to authenticated;
do $$begin
 if exists(select 1 from pg_publication where pubname='supabase_realtime') and not exists(select 1 from pg_publication_tables where pubname='supabase_realtime' and schemaname='public' and tablename='tms_audit_events') then alter publication supabase_realtime add table public.tms_audit_events; end if;
end $$;
commit;

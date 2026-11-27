begin;
-- Fiscal documents are isolated from the shared operational snapshot. Only the
-- verified server function may create records or activate reference templates.
create or replace function tms_private.fiscal_xml_allowed(p_action text)
returns boolean language plpgsql stable security definer set search_path='' as $$
declare p jsonb := tms_private.profile();
begin
  if auth.uid() is null or not coalesce((p->>'activo')::boolean,false)
     or p_action not in ('ver','editar','importar') then return false; end if;
  return p->>'rol'='Admin' or coalesce((
    select s.valor #>> array[p->>'rol','fiscalXml',p_action]='true'
    from public.tms_settings s where s.clave='role_permissions'
  ),false);
end $$;
revoke all on function tms_private.fiscal_xml_allowed(text) from public,anon;
grant execute on function tms_private.fiscal_xml_allowed(text) to authenticated;
create or replace function public.tms_fiscal_xml_access()
returns jsonb language sql stable security invoker set search_path='' as $$
  select jsonb_build_object('view',tms_private.fiscal_xml_allowed('ver'),
    'edit',tms_private.fiscal_xml_allowed('editar'),
    'import',tms_private.fiscal_xml_allowed('importar'),
    'admin',tms_private.audit_admin());
$$;
revoke all on function public.tms_fiscal_xml_access() from public,anon;
grant execute on function public.tms_fiscal_xml_access() to authenticated;

create table if not exists public.tms_fiscal_xml_templates (
  id uuid primary key default gen_random_uuid(),
  document_type text not null check (document_type in ('credit','debit')),
  original_xml text not null check (octet_length(original_xml) between 1 and 2097152),
  corrected_xml text not null check (octet_length(corrected_xml) between 1 and 2097152),
  rules jsonb not null,
  sap_accepted boolean not null default false,
  active boolean not null default false,
  created_by uuid not null references auth.users(id),
  created_at timestamptz not null default now(),
  check (not active or sap_accepted)
);
create unique index if not exists fiscal_xml_active_template on public.tms_fiscal_xml_templates(document_type) where active;

create or replace function public.tms_fiscal_xml_rules()
returns jsonb language plpgsql stable security definer set search_path='' as $$
declare result jsonb;
begin
  if not tms_private.fiscal_xml_allowed('ver') then raise exception 'Acceso denegado' using errcode='42501'; end if;
  select coalesce(jsonb_object_agg(document_type,rules),'{}'::jsonb) into result
  from public.tms_fiscal_xml_templates where active and sap_accepted;
  return result;
end $$;
revoke all on function public.tms_fiscal_xml_rules() from public,anon;
grant execute on function public.tms_fiscal_xml_rules() to authenticated;

create table if not exists public.tms_fiscal_xml_history (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  created_by uuid not null references auth.users(id),
  created_by_label text not null default '',
  original_name text not null check (length(original_name) between 1 and 255),
  document_type text not null check (document_type in ('credit','debit')),
  template_id uuid not null references public.tms_fiscal_xml_templates(id),
  document_number text not null default '',
  note_ncf text not null default '',
  affected_ncf text not null,
  affected_date date not null,
  changes jsonb not null,
  validations jsonb not null,
  observations text not null default '' check (length(observations) <= 2000),
  original_xml text not null check (octet_length(original_xml) between 1 and 2097152),
  corrected_xml text not null check (octet_length(corrected_xml) between 1 and 2097152)
);
create index if not exists fiscal_xml_history_created on public.tms_fiscal_xml_history(created_at desc);
create index if not exists fiscal_xml_history_ncf on public.tms_fiscal_xml_history(affected_ncf);
create table if not exists public.tms_fiscal_xml_events (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  created_by uuid not null references auth.users(id),
  created_by_label text not null default '',
  operation text not null check(operation in ('analyze','correct','template','download')),
  document_type text not null check(document_type in ('credit','debit')),
  original_name text not null default '',
  outcome text not null check(outcome in ('success','error')),
  detail text not null default '' check(length(detail)<=1000)
);
create index if not exists fiscal_xml_events_created on public.tms_fiscal_xml_events(created_at desc);
alter table public.tms_fiscal_xml_templates enable row level security;
alter table public.tms_fiscal_xml_history enable row level security;
alter table public.tms_fiscal_xml_events enable row level security;
revoke all on public.tms_fiscal_xml_templates,public.tms_fiscal_xml_history,public.tms_fiscal_xml_events from public,anon,authenticated;
grant select on public.tms_fiscal_xml_templates,public.tms_fiscal_xml_events to authenticated;
grant select (id,created_at,created_by,created_by_label,original_name,document_type,template_id,document_number,note_ncf,affected_ncf,affected_date,changes,validations,observations) on public.tms_fiscal_xml_history to authenticated;
grant select,insert on public.tms_fiscal_xml_history,public.tms_fiscal_xml_events to service_role;
grant select on public.tms_fiscal_xml_templates to service_role;
drop policy if exists fiscal_xml_template_admin_read on public.tms_fiscal_xml_templates;
create policy fiscal_xml_template_admin_read on public.tms_fiscal_xml_templates for select to authenticated using (
  (select tms_private.fiscal_xml_allowed('ver')) and
  (select tms_private.profile()->>'rol')='Admin'
);
drop policy if exists fiscal_xml_history_read on public.tms_fiscal_xml_history;
create policy fiscal_xml_history_read on public.tms_fiscal_xml_history for select to authenticated using ((select tms_private.fiscal_xml_allowed('ver')));
drop policy if exists fiscal_xml_events_read on public.tms_fiscal_xml_events;
create policy fiscal_xml_events_read on public.tms_fiscal_xml_events for select to authenticated using ((select tms_private.fiscal_xml_allowed('ver')));

create or replace function public.tms_fiscal_activate_template(p_type text,p_original text,p_corrected text,p_rules jsonb,p_actor uuid)
returns uuid language plpgsql security definer set search_path='' as $$
declare template_id uuid;
begin
  -- The edge function validates both XML documents and the administrator's
  -- attestation before invoking this service-only, atomic operation.
  perform pg_advisory_xact_lock(hashtext('tms_fiscal_xml_templates'));
  update public.tms_fiscal_xml_templates set active=false where document_type=p_type and active;
  insert into public.tms_fiscal_xml_templates(document_type,original_xml,corrected_xml,rules,sap_accepted,active,created_by)
  values(p_type,p_original,p_corrected,p_rules,true,true,p_actor) returning id into template_id;
  insert into public.tms_fiscal_xml_events(created_by,created_by_label,operation,document_type,outcome,detail)
  values(p_actor,p_actor::text,'template',p_type,'success','Plantilla declarada aceptada por SAP');
  return template_id;
end $$;
revoke all on function public.tms_fiscal_activate_template(text,text,text,jsonb,uuid) from public,anon,authenticated;
grant execute on function public.tms_fiscal_activate_template(text,text,text,jsonb,uuid) to service_role;
create or replace function public.tms_fiscal_save_correction(p_record jsonb)
returns jsonb language plpgsql security definer set search_path='' as $$
declare r public.tms_fiscal_xml_history;
begin
  insert into public.tms_fiscal_xml_history(created_by,created_by_label,original_name,document_type,template_id,document_number,note_ncf,affected_ncf,affected_date,changes,validations,observations,original_xml,corrected_xml)
  values ((p_record->>'created_by')::uuid,p_record->>'created_by_label',p_record->>'original_name',p_record->>'document_type',(p_record->>'template_id')::uuid,p_record->>'document_number',p_record->>'note_ncf',p_record->>'affected_ncf',(p_record->>'affected_date')::date,p_record->'changes',p_record->'validations',p_record->>'observations',p_record->>'original_xml',p_record->>'corrected_xml') returning * into r;
  insert into public.tms_fiscal_xml_events(created_by,created_by_label,operation,document_type,original_name,outcome,detail)
  values(r.created_by,r.created_by_label,'correct',r.document_type,r.original_name,'success','Validación estructural correcta');
  return jsonb_build_object('id',r.id,'created_at',r.created_at);
end $$;
revoke all on function public.tms_fiscal_save_correction(jsonb) from public,anon,authenticated;
grant execute on function public.tms_fiscal_save_correction(jsonb) to service_role;
commit;

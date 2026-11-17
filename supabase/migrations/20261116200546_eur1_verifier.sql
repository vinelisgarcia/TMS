begin;
create or replace function tms_private.eur1_allowed(p_action text)
returns boolean language plpgsql stable security definer set search_path='' as $$
declare p jsonb:=tms_private.profile();
begin
 if auth.uid() is null or not coalesce((p->>'activo')::boolean,false) or p_action not in ('ver','editar','importar') then return false; end if;
 return p->>'rol'='Admin' or coalesce((select s.valor #>> array[p->>'rol','eur1',p_action]='true' from public.tms_settings s where s.clave='role_permissions'),false);
end $$;
revoke all on function tms_private.eur1_allowed(text) from public,anon;
grant execute on function tms_private.eur1_allowed(text) to authenticated;

create table if not exists public.tms_eur1_cases (
 id uuid primary key default gen_random_uuid(),title text not null check(length(title) between 1 and 180),
 revision bigint not null default 0 check(revision>=0), result jsonb check(result is null or pg_column_size(result)<=26214400),
 created_at timestamptz not null default now(),updated_at timestamptz not null default now(),
 created_by uuid default auth.uid(),updated_by uuid default auth.uid()
);
create table if not exists public.tms_eur1_documents (
 id uuid primary key,case_id uuid not null references public.tms_eur1_cases(id),
 parent_id uuid references public.tms_eur1_documents(id),original_name text not null check(length(original_name) between 1 and 1024),
 storage_path text not null unique,sha256 text not null check(sha256 ~ '^[0-9a-f]{64}$'),
 size_bytes bigint not null check(size_bytes>0 and size_bytes<=52428800),media_type text not null,
 revision bigint not null default 0 check(revision>=0),kind text not null default 'unknown',
 status text not null default 'pending' check(status in ('pending','stored','ready','error')),
 extraction jsonb,error text not null default '',classification_confirmed boolean not null default false,
 extraction_bytes integer generated always as (coalesce(octet_length(extraction::text),0)) stored,
 line_count integer generated always as (coalesce(jsonb_array_length(extraction->'lines'),0)) stored,
 created_at timestamptz not null default now(),updated_at timestamptz not null default now(),
 created_by uuid default auth.uid(),updated_by uuid default auth.uid(),
 unique(case_id,sha256),check(storage_path=case_id::text||'/'||id::text||'/original'),
 check(kind in ('unknown','certificate','proposal','invoice','packing','hs','complement','archive')),
 check(extraction is null or (jsonb_typeof(extraction)='object' and octet_length(extraction::text)<=12582912))
);
create index if not exists tms_eur1_documents_case on public.tms_eur1_documents(case_id,created_at);
create index if not exists tms_eur1_cases_updated on public.tms_eur1_cases(updated_at desc);
create table if not exists public.tms_eur1_sync_state(id boolean primary key default true check(id),revision bigint not null default 0);
insert into public.tms_eur1_sync_state(id) values(true) on conflict(id) do nothing;
alter table public.tms_eur1_sync_state enable row level security;
revoke all on public.tms_eur1_sync_state from public,anon,authenticated;
grant select on public.tms_eur1_sync_state to authenticated;
drop policy if exists "eur1 revision readers" on public.tms_eur1_sync_state;
create policy "eur1 revision readers" on public.tms_eur1_sync_state for select to authenticated using ((select tms_private.eur1_allowed('ver')));
alter table public.tms_eur1_cases enable row level security;
alter table public.tms_eur1_documents enable row level security;
revoke all on public.tms_eur1_cases,public.tms_eur1_documents from public,anon,authenticated;
grant select,insert,update on public.tms_eur1_cases,public.tms_eur1_documents to authenticated;
drop policy if exists "eur1 case readers" on public.tms_eur1_cases;
create policy "eur1 case readers" on public.tms_eur1_cases for select to authenticated using ((select tms_private.eur1_allowed('ver')));
drop policy if exists "eur1 case creators" on public.tms_eur1_cases;
create policy "eur1 case creators" on public.tms_eur1_cases for insert to authenticated with check ((select tms_private.eur1_allowed('ver')) and (select tms_private.eur1_allowed('importar')));
drop policy if exists "eur1 case editors" on public.tms_eur1_cases;
create policy "eur1 case editors" on public.tms_eur1_cases for update to authenticated using ((select tms_private.eur1_allowed('ver')) and ((select tms_private.eur1_allowed('editar')) or (select tms_private.eur1_allowed('importar')))) with check ((select tms_private.eur1_allowed('ver')));
drop policy if exists "eur1 document readers" on public.tms_eur1_documents;
create policy "eur1 document readers" on public.tms_eur1_documents for select to authenticated using ((select tms_private.eur1_allowed('ver')));
drop policy if exists "eur1 document creators" on public.tms_eur1_documents;
create policy "eur1 document creators" on public.tms_eur1_documents for insert to authenticated with check ((select tms_private.eur1_allowed('ver')) and (select tms_private.eur1_allowed('importar')));
drop policy if exists "eur1 document editors" on public.tms_eur1_documents;
create policy "eur1 document editors" on public.tms_eur1_documents for update to authenticated using ((select tms_private.eur1_allowed('ver')) and ((select tms_private.eur1_allowed('editar')) or (select tms_private.eur1_allowed('importar')))) with check ((select tms_private.eur1_allowed('ver')));

create or replace function tms_private.eur1_stamp()
returns trigger language plpgsql set search_path='' as $$
begin
 if tg_op='INSERT' then
  new.revision:=0;new.created_by:=auth.uid();new.created_at:=clock_timestamp();
  if tg_table_name='tms_eur1_documents' then
   perform 1 from public.tms_eur1_cases where id=new.case_id for update;
   if (select count(*) from public.tms_eur1_documents where case_id=new.case_id)>=250 then raise exception 'El expediente admite hasta 250 documentos. Divide el embarque.' using errcode='22023'; end if;
   if new.parent_id is not null and not exists(select 1 from public.tms_eur1_documents where id=new.parent_id and case_id=new.case_id) then raise exception 'El ZIP de origen pertenece a otro expediente.' using errcode='22023'; end if;
  end if;
 else
  if new.id is distinct from old.id then raise exception 'Identificador inmutable.' using errcode='42501'; end if;
  if new.revision<>old.revision+1 then raise exception 'La versión cambió. Recarga el expediente para conservar los cambios de la otra sesión.' using errcode='40001'; end if;
  new.created_by:=old.created_by;new.created_at:=old.created_at;
  if tg_table_name='tms_eur1_documents' then
   if (new.id,new.case_id,new.parent_id,new.original_name,new.storage_path,new.sha256,new.size_bytes,new.media_type) is distinct from (old.id,old.case_id,old.parent_id,old.original_name,old.storage_path,old.sha256,old.size_bytes,old.media_type) then raise exception 'Los documentos originales son inmutables.' using errcode='42501'; end if;
  end if;
 end if;
 new.updated_at:=clock_timestamp();new.updated_by:=auth.uid();return new;
end $$;
create or replace function tms_private.eur1_invalidate()
returns trigger language plpgsql security definer set search_path='' as $$
begin
 update public.tms_eur1_cases set revision=revision+1,result=null where id=new.case_id;
 return null;
end $$;
create or replace function tms_private.eur1_revision()
returns trigger language plpgsql security definer set search_path='' as $$begin update public.tms_eur1_sync_state set revision=revision+1 where id;return null;end $$;
revoke all on function tms_private.eur1_stamp(),tms_private.eur1_invalidate(),tms_private.eur1_revision() from public,anon,authenticated;
drop trigger if exists eur1_stamp on public.tms_eur1_cases;
create trigger eur1_stamp before insert or update on public.tms_eur1_cases for each row execute function tms_private.eur1_stamp();
drop trigger if exists eur1_revision on public.tms_eur1_cases;
create trigger eur1_revision after insert or update on public.tms_eur1_cases for each row execute function tms_private.eur1_revision();
drop trigger if exists eur1_stamp on public.tms_eur1_documents;
create trigger eur1_stamp before insert or update on public.tms_eur1_documents for each row execute function tms_private.eur1_stamp();
drop trigger if exists eur1_invalidate on public.tms_eur1_documents;
create trigger eur1_invalidate after insert or update on public.tms_eur1_documents for each row execute function tms_private.eur1_invalidate();
select tms_private.attach_audit('tms_eur1_cases');
select tms_private.attach_audit('tms_eur1_documents');

insert into storage.buckets(id,name,public,file_size_limit) values('tms-eur1-originals','tms-eur1-originals',false,52428800) on conflict(id) do nothing;
drop policy if exists "eur1 private originals read" on storage.objects;
create policy "eur1 private originals read" on storage.objects for select to authenticated using (
 bucket_id='tms-eur1-originals' and (select tms_private.eur1_allowed('ver'))
 and exists(select 1 from public.tms_eur1_documents d where d.storage_path=name)
);
drop policy if exists "eur1 originals append only" on storage.objects;
create policy "eur1 originals append only" on storage.objects for insert to authenticated with check (
 bucket_id='tms-eur1-originals' and (select tms_private.eur1_allowed('ver')) and (select tms_private.eur1_allowed('importar'))
 and exists(select 1 from public.tms_eur1_documents d where d.storage_path=name and d.created_by=auth.uid())
);
-- No UPDATE or DELETE policies: a revised file is a new immutable original.
do $$declare t text;begin foreach t in array array['tms_eur1_sync_state'] loop
 if exists(select 1 from pg_publication where pubname='supabase_realtime') and not exists(select 1 from pg_publication_tables where pubname='supabase_realtime' and schemaname='public' and tablename=t) then execute format('alter publication supabase_realtime add table public.%I',t); end if;
end loop;end $$;
commit;

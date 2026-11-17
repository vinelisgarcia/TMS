-- Activation smoke test. Every generated case, document and audit event rolls back.
begin;
do $$declare claims text;begin
 select jsonb_build_object('sub',auth_user_id,'email',email,'role','authenticated','session_id',gen_random_uuid())::text into claims
 from public.tms_user_profiles where activo and rol='Admin' and auth_user_id is not null order by created_at limit 1;
 perform set_config('request.jwt.claims',claims,true);
end $$;
set local role authenticated;
do $$declare c public.tms_eur1_cases;d uuid:=gen_random_uuid();n integer;begin
 if not tms_private.eur1_allowed('ver') or not tms_private.eur1_allowed('importar') then raise exception 'Administrator activation failed'; end if;
 insert into public.tms_eur1_cases(title) values('Validación temporal con rollback') returning * into c;
 insert into public.tms_eur1_documents(id,case_id,original_name,storage_path,sha256,size_bytes,media_type)
 values(d,c.id,'Validación temporal.pdf',c.id::text||'/'||d::text||'/original',repeat('0',64),1,'application/pdf');
 update public.tms_eur1_documents set revision=1,status='ready',kind='certificate',extraction='{"lines":[]}' where id=d and revision=0;
 update public.tms_eur1_cases set revision=3,result='{"conclusion":"Validación temporal"}' where id=c.id and revision=2;
 get diagnostics n=row_count;if n<>1 then raise exception 'Versioned analysis save failed'; end if;
 update public.tms_eur1_cases set revision=3,result=null where id=c.id and revision=2;
 get diagnostics n=row_count;if n<>0 then raise exception 'Stale writer was not rejected'; end if;
 set constraints all immediate;
 if not exists(select 1 from public.tms_audit_events where module='eur1' and is_change and details->>'tabla'='tms_eur1_documents') then raise exception 'Document audit missing'; end if;
end $$;
rollback;

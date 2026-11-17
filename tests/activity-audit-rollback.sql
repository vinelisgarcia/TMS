-- Administrative smoke test: signed-claim simulation only inside ROLLBACK.
-- No accounts, credentials or durable operational/audit records are changed.
begin;
do $$
declare p public.tms_user_profiles; v uuid; n bigint; row_id uuid;
begin
 select * into p from public.tms_user_profiles where activo and rol='Admin' and auth_user_id is not null order by created_at limit 1;
 if p.id is null then raise exception 'No active linked administrator available'; end if;
 perform set_config('request.jwt.claims',jsonb_build_object('sub',p.auth_user_id,'email',p.email,'role','authenticated','session_id',gen_random_uuid())::text,true);
 v:=public.tms_audit_action('auditoriaSistema','visita','desktop');
 if v is null then raise exception 'Visit was not recorded'; end if;
 if v is distinct from public.tms_audit_action('auditoriaSistema','visita','desktop') then raise exception 'Heartbeat duplicated a visit'; end if;
 perform public.tms_audit_action('calendario','ver_modulo','desktop');
 select count(*) into n from public.tms_audit_events where visit_id=v and is_change;
 select id into row_id from public.tms_order_lines limit 1;
 perform set_config('tms.sync_write','on',true);
 if row_id is not null then
  update public.tms_order_lines set cantidad_pendiente=cantidad_pendiente+1 where id=row_id;
  update public.tms_order_lines set cantidad_pendiente=cantidad_pendiente-1 where id=row_id;
 end if;
 set constraints all immediate;
 if n<>(select count(*) from public.tms_audit_events where visit_id=v and is_change) then raise exception 'Net no-op created change notifications'; end if;
 if exists(select 1 from tms_private.audit_buffer where transaction_id=txid_current()) then raise exception 'Private transaction buffer was not cleared'; end if;
end $$;
rollback;

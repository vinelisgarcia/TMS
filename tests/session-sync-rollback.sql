-- Run only after 0008_session_sync.sql. This test uses existing active profiles
-- and rolls back every change, including its revision increments.
-- It performs no Auth administration and creates no real user accounts.
begin;

do $$
declare v_user record;
begin
  select u.id, u.email into v_user
  from public.tms_user_profiles p
  join auth.users u on u.id = p.auth_user_id or lower(u.email) = lower(p.email)
  where p.activo is true and p.rol = 'Admin'
  order by p.created_at limit 1;
  if v_user.id is null then raise exception 'Test requires an existing active Admin Auth account.'; end if;
  perform set_config('request.jwt.claims', jsonb_build_object('sub', v_user.id, 'role', 'authenticated', 'email', v_user.email)::text, true);
  perform set_config('request.jwt.claim.sub', v_user.id::text, true);
  perform set_config('request.jwt.claim.role', 'authenticated', true);
end;
$$;
set local role authenticated;

do $$
declare
  v_before jsonb;
  v_after jsonb;
  v_result jsonb;
  v_revision bigint;
  v_state jsonb;
  v_payload jsonb;
  v_failed boolean;
begin
  v_before := public.tms_get_snapshot();
  v_revision := (v_before ->> 'revision')::bigint;
  if not (v_before ? 'tables') then raise exception 'FAIL: full snapshot missing.'; end if;
  if jsonb_array_length(v_before #> '{tables,tms_order_lines}') <> (select count(*) from public.tms_order_lines) then
    raise exception 'FAIL: snapshot row count differs from database.';
  end if;
  if (public.tms_get_snapshot(v_revision) ->> 'unchanged')::boolean is not true then
    raise exception 'FAIL: unchanged poll should return no table payload.';
  end if;

  v_result := public.tms_save_snapshot(v_revision, '{}'::jsonb);
  if (v_result ->> 'revision')::bigint <> v_revision or (v_result ->> 'changed')::boolean is not false then
    raise exception 'FAIL: empty save changed revision.';
  end if;

  select valor into v_state from public.tms_settings where clave = 'app_state';
  v_state := coalesce(v_state, '{}'::jsonb);
  v_result := public.tms_save_snapshot(v_revision, jsonb_build_object('tms_settings', jsonb_build_array(jsonb_build_object('clave', 'app_state', 'valor', v_state))));
  if (v_result ->> 'revision')::bigint <> v_revision or (v_result ->> 'changed')::boolean is not false then
    raise exception 'FAIL: identical settings save changed revision.';
  end if;

  -- A real change advances the revision once.
  v_payload := jsonb_build_object('tms_settings', jsonb_build_array(jsonb_build_object('clave', 'app_state', 'valor', v_state || '{"__session_sync_rollback_test":true}'::jsonb)));
  v_result := public.tms_save_snapshot(v_revision, v_payload);
  if (v_result ->> 'revision')::bigint <> v_revision + 1 or (v_result ->> 'changed')::boolean is not true then
    raise exception 'FAIL: successful save did not advance exactly one revision.';
  end if;
  v_after := public.tms_get_snapshot();

  -- Session B still holds the old revision: its save must fail atomically.
  v_failed := false;
  begin
    perform public.tms_save_snapshot(v_revision, jsonb_build_object('tms_settings', jsonb_build_array(jsonb_build_object('clave', 'app_state', 'valor', v_state))));
  exception when serialization_failure then v_failed := true;
  end;
  if not v_failed then raise exception 'FAIL: stale session save succeeded.'; end if;
  if public.tms_get_snapshot() is distinct from v_after then raise exception 'FAIL: stale save mutated data.'; end if;

  -- A legacy client must fail before updating even one setting.
  v_failed := false;
  begin
    update public.tms_settings set valor = '{}'::jsonb where clave = 'app_state';
  exception when insufficient_privilege then v_failed := true;
  end;
  if not v_failed then raise exception 'FAIL: legacy direct write succeeded.'; end if;
  if has_function_privilege(current_user, 'public.tms_replace_table(text,jsonb)', 'EXECUTE') then
    raise exception 'FAIL: legacy replacement RPC remains executable.';
  end if;
  if public.tms_get_snapshot() is distinct from v_after then raise exception 'FAIL: legacy write mutated data.'; end if;

  -- A late table error rolls back an earlier settings change and revision.
  v_failed := false;
  begin
    perform public.tms_save_snapshot(v_revision + 1, jsonb_build_object(
      'tms_settings', jsonb_build_array(jsonb_build_object('clave', 'app_state', 'valor', v_state || '{"__must_rollback":true}'::jsonb)),
      'tms_order_lines', jsonb_build_array(jsonb_build_object('cliente_id', 'TEST-ROLLBACK', 'clave_unica', null))
    ));
  exception when not_null_violation then v_failed := true;
  end;
  if not v_failed then raise exception 'FAIL: invalid late table unexpectedly succeeded.'; end if;
  if public.tms_get_snapshot() is distinct from v_after then raise exception 'FAIL: failed transaction left partial changes.'; end if;
  raise notice 'PASS: snapshot completeness, unchanged poll, no-op saves, revision CAS, legacy guards, atomic rollback.';
end;
$$;

reset role;
do $$
declare v_user record;
begin
  select u.id, u.email into v_user
  from public.tms_user_profiles p
  join auth.users u on u.id = p.auth_user_id or lower(u.email) = lower(p.email)
  where p.activo is true and p.rol in ('Solo lectura', 'Comercial')
  order by p.created_at limit 1;
  if v_user.id is null then raise exception 'Test requires an existing active read-only or Comercial Auth account.'; end if;
  perform set_config('request.jwt.claims', jsonb_build_object('sub', v_user.id, 'role', 'authenticated', 'email', v_user.email)::text, true);
  perform set_config('request.jwt.claim.sub', v_user.id::text, true);
  perform set_config('request.jwt.claim.role', 'authenticated', true);
end;
$$;
set local role authenticated;
do $$
declare v_before jsonb; v_failed boolean := false;
begin
  v_before := public.tms_get_snapshot();
  begin
    perform public.tms_save_snapshot((v_before ->> 'revision')::bigint,
      jsonb_build_object('tms_settings', jsonb_build_array(jsonb_build_object('clave', 'app_state', 'valor', '{}'::jsonb))));
  exception when insufficient_privilege then v_failed := true;
  end;
  if not v_failed then raise exception 'FAIL: read-only account wrote operational data.'; end if;
  if public.tms_get_snapshot() is distinct from v_before then raise exception 'FAIL: rejected read-only save mutated data.'; end if;
  raise notice 'PASS: read-only account can synchronize but cannot write operational data.';
end;
$$;

rollback;

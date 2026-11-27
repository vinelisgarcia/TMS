begin;
-- An assertion of SAP acceptance is not evidence of DGII acceptance.
-- Preserve any stored examples but require fresh DGII validation.
update public.tms_fiscal_xml_templates set active=false,sap_accepted=false where active or sap_accepted;
alter table public.tms_fiscal_xml_templates rename column sap_accepted to dgii_accepted;
create or replace function public.tms_fiscal_xml_rules()
returns jsonb language plpgsql stable security definer set search_path='' as $$
declare result jsonb;
begin
 if not tms_private.fiscal_xml_allowed('ver') then raise exception 'Acceso denegado' using errcode='42501'; end if;
 select coalesce(jsonb_object_agg(document_type,rules),'{}'::jsonb) into result
 from public.tms_fiscal_xml_templates where active and dgii_accepted;
 return result;
end $$;
create or replace function public.tms_fiscal_activate_template(p_type text,p_original text,p_corrected text,p_rules jsonb,p_actor uuid)
returns uuid language plpgsql security definer set search_path='' as $$
declare template_id uuid;
begin
 perform pg_advisory_xact_lock(hashtext('tms_fiscal_xml_templates'));
 update public.tms_fiscal_xml_templates set active=false where document_type=p_type and active;
 insert into public.tms_fiscal_xml_templates(document_type,original_xml,corrected_xml,rules,dgii_accepted,active,created_by)
 values(p_type,p_original,p_corrected,p_rules,true,true,p_actor) returning id into template_id;
 insert into public.tms_fiscal_xml_events(created_by,created_by_label,operation,document_type,outcome,detail)
 values(p_actor,p_actor::text,'template',p_type,'success','Plantilla declarada aceptada por la DGII');
 return template_id;
end $$;
revoke all on function public.tms_fiscal_xml_rules() from public,anon;
grant execute on function public.tms_fiscal_xml_rules() to authenticated;
revoke all on function public.tms_fiscal_activate_template(text,text,text,jsonb,uuid) from public,anon,authenticated;
grant execute on function public.tms_fiscal_activate_template(text,text,text,jsonb,uuid) to service_role;
commit;

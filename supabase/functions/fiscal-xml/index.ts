import { createClient } from 'npm:@supabase/supabase-js@2.95.3';
import { DOMParser, XMLSerializer } from 'npm:@xmldom/xmldom@0.9.12';
// The parser and engine are bundled with this function. Fiscal XML is never
// sent to a service other than the TMS's own Supabase project.
Object.assign(globalThis, { DOMParser, XMLSerializer });
await import('./fiscal-xml-domain.js');
const xml = (globalThis as any).TMSFiscalXML;
const cors = {'Access-Control-Allow-Origin':'*','Access-Control-Allow-Headers':'authorization, x-client-info, apikey, content-type','Access-Control-Allow-Methods':'POST, OPTIONS'};
const response = (body: unknown, status = 200) => new Response(JSON.stringify(body), {status, headers:{...cors,'Content-Type':'application/json','Cache-Control':'no-store'}});
const LIMIT = 6 * 1024 * 1024;
async function readBody(req: Request) {
  if (Number(req.headers.get('content-length') || 0) > LIMIT) throw Error('La solicitud es demasiado grande.');
  const reader = req.body?.getReader(); if (!reader) throw Error('Falta el contenido.');
  const parts: Uint8Array[] = []; let size = 0;
  while (true) { const {value, done} = await reader.read(); if (done) break; size += value.length; if (size > LIMIT) { await reader.cancel(); throw Error('La solicitud es demasiado grande.'); } parts.push(value); }
  const bytes = new Uint8Array(size); let offset = 0; for (const p of parts) {bytes.set(p,offset);offset+=p.length;}
  return JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(bytes));
}
Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') return new Response('ok',{headers:cors});
  if (req.method !== 'POST') return response({error:'Método no permitido.'},405);
  const url = Deno.env.get('SUPABASE_URL') || '', anon = Deno.env.get('SUPABASE_ANON_KEY') || '', secret = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') || '';
  if (!url || !anon || !secret) return response({error:'Servicio no configurado.'},503);
  const caller = createClient(url,anon,{global:{headers:{Authorization:req.headers.get('Authorization') || ''}},auth:{persistSession:false}});
  const {data:identity,error:authError} = await caller.auth.getUser();
  if (authError || !identity.user) return response({error:'Sesión inválida.'},401);
  const {data:access,error:accessError} = await caller.rpc('tms_fiscal_xml_access');
  if (accessError || !access?.view) return response({error:'No tienes acceso a herramientas fiscales.'},403);
  const admin = createClient(url,secret,{auth:{persistSession:false}});
  let body: any, name = '';
  async function event(outcome: 'success'|'error', detail: string) {
    if (!body || !['credit','debit'].includes(body.type) || !['analyze','correct','template','download'].includes(body.action)) return;
    const {error} = await admin.from('tms_fiscal_xml_events').insert({created_by:identity.user!.id,created_by_label:identity.user!.email || identity.user!.id,operation:body.action,document_type:body.type,original_name:name,outcome,detail:detail.slice(0,1000)});
    if (error) throw Error('No se pudo registrar la operación fiscal.');
  }
  try {
    body = await readBody(req);
    if (!['credit','debit'].includes(body.type)) throw Error('Selecciona el tipo de documento.');
    if (body.action === 'download') {
      if (!['original','corrected'].includes(body.part) || !/^[0-9a-f-]{36}$/i.test(String(body.id || ''))) throw Error('Solicitud de descarga inválida.');
      const {data:record,error} = await admin.from('tms_fiscal_xml_history').select('original_name,document_type,original_xml,corrected_xml').eq('id',body.id).maybeSingle();
      if (error || !record || record.document_type !== body.type) throw Error('No se encontró el documento solicitado.');
      name=record.original_name;
      const fresh=await caller.rpc('tms_fiscal_xml_access');
      if (fresh.error || !fresh.data?.view) return response({error:'Tu acceso cambió.'},403);
      await event('success',body.part==='original'?'Descarga del XML original':'Descarga del XML corregido');
      return response({name,xml:body.part==='original'?record.original_xml:record.corrected_xml});
    }
    if (body.action === 'template') {
      if (!access.admin || body.dgiiAccepted !== true) return response({error:'Solo un administrador puede registrar plantillas aceptadas por la DGII.'},403);
      if (typeof body.original !== 'string' || typeof body.corrected !== 'string') throw Error('Faltan los dos XML de referencia.');
      const rules = xml.template(body.original,body.corrected,body.type);
      const {data,error} = await admin.rpc('tms_fiscal_activate_template',{p_type:body.type,p_original:body.original,p_corrected:body.corrected,p_rules:rules,p_actor:identity.user.id});
      if (error) throw Error('No se pudo guardar la plantilla validada.');
      return response({id:data});
    }
    if (!['analyze','correct'].includes(body.action)) throw Error('Acción no admitida.');
    if (body.action === 'correct' && !access.edit && !access.import) return response({error:'Tu rol no permite generar correcciones.'},403);
    if (typeof body.original !== 'string') throw Error('Falta el XML original.');
    name = String(body.filename || '').replace(/^.*[\\/]/,'');
    if (!name || name.length > 255 || !/\.xml$/i.test(name) || /[\x00-\x1f]/.test(name)) throw Error('El nombre debe corresponder a un archivo XML.');
    const {data:templates,error:templateError} = await admin.from('tms_fiscal_xml_templates').select('id,rules').eq('document_type',body.type).eq('active',true).eq('dgii_accepted',true).limit(1);
    if (templateError) throw Error('No se pudo consultar la plantilla fiscal.');
    if (body.action === 'analyze') {
      const document = xml.parse(body.original), information = xml.information(document);
      if (information.detectedType && information.detectedType !== body.type) throw Error('El tipo identificado no coincide con la sección elegida.');
      let analysis = null;
      if (templates?.length) { const a = xml.analyze(body.original,templates[0].rules); analysis = {fields:a.fields,paths:a.paths}; }
      await event('success',analysis?'XML validado estructuralmente':'XML leído; plantilla aceptada por la DGII pendiente');
      return response({information,analysis});
    }
    if (!templates?.length) throw Error('Este tipo aún no tiene una plantilla validada y aceptada por la DGII.');
    if (body.confirmedType !== true) throw Error('Confirma el tipo de documento antes de generar.');
    const template = templates[0], result = xml.correct(body.original,template.rules,body.input || {});
    const observations = String(body.observations || '').trim(); if (observations.length > 2000) throw Error('Las observaciones superan 2.000 caracteres.');
    // Recheck permissions immediately before persisting, including active status.
    const fresh = await caller.rpc('tms_fiscal_xml_access');
    if (fresh.error || !fresh.data?.view || (!fresh.data?.edit && !fresh.data?.import)) return response({error:'Tu acceso cambió; la corrección no se guardó.'},403);
    const {data,error} = await admin.rpc('tms_fiscal_save_correction',{p_record:{created_by:identity.user.id,created_by_label:identity.user.email || identity.user.id,original_name:name,document_type:body.type,template_id:template.id,document_number:result.information.number,note_ncf:result.information.ownNcf,affected_ncf:result.after.ncf,affected_date:result.after.date,changes:{before:result.before,after:result.after,information:result.information},validations:{structural:'correct',xsd:'unavailable',warnings:result.warnings},observations,original_xml:body.original,corrected_xml:result.xml}});
    if (error) throw Error('No se pudo guardar la corrección; no se generó el archivo.');
    return response({id:data.id,createdAt:data.created_at,...result});
  } catch (error) {
    const detail = error instanceof Error ? error.message : 'No se pudo procesar el XML.';
    try {await event('error',detail);} catch (_) {return response({error:'No se pudo registrar el error de validación.'},503);}
    return response({error:detail},400);
  }
});

(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.TMSEURTransit=api;})(typeof globalThis!=='undefined'?globalThis:this,function(){
 'use strict';
 const value=field=>field?.value;
 const text=value=>String(value??'').trim();
 const key=value=>text(value).toUpperCase();
 function prepare(documents, existing, caseId, createId, timestamp){
  const groups=new Map(), blocked=new Set(), messages=[];
  for(const doc of documents){
   if(doc.status!=='ready'||doc.kind!=='packing')continue;
   const d=doc.extraction;
   if(!d||!Array.isArray(d.lines)||!d.lines.length){messages.push(`${doc.original_name}: no se detectaron líneas de packing.`);continue;}
   const invoice=text(value(d.meta?.invoice));
   if(!doc.classification_confirmed && (d.detected?.kind!=='packing'||d.detected?.review)){if(invoice)blocked.add(key(invoice));messages.push(`${doc.original_name}: confirma el tipo de documento.`);continue;}
   if(!invoice){messages.push(`${doc.original_name}: revisa el número de factura en el documento.`);continue;}
   if((d.issues||[]).length){blocked.add(key(invoice));messages.push(`${doc.original_name}: revisa los problemas de lectura antes de enviar a tránsito.`);continue;}
   const materials=new Map();let invalid=false;
   for(const line of d.lines){
    const code=text(value(line.reference)), qty=value(line.quantity);
    if(!code||typeof qty!=='number'||!Number.isFinite(qty)||qty<=0||line.reference?.source?.method==='ocr'||line.quantity?.source?.method==='ocr'){invalid=true;break;}
    const k=key(code), prior=materials.get(k), source=line.reference?.source||line.source||{};
    if(prior){prior.cantidad=Math.round((prior.cantidad+qty)*1e10)/1e10;prior.eur1.lineas.push(source);}
    else materials.set(k,{factura:invoice,eta:'',codigoMaterial:code,descripcion:text(value(line.description)),cantidad:qty,cliente:'',eur1:{expedienteId:caseId,documentoId:doc.id,archivo:doc.original_name,lineas:[source]}});
   }
   if(invalid){blocked.add(key(invoice));messages.push(`${doc.original_name}: hay referencias o cantidades pendientes de revisión.`);continue;}
   const invoiceKey=key(invoice), representation={doc,invoice,materials};
   if(!groups.has(invoiceKey))groups.set(invoiceKey,[]);groups.get(invoiceKey).push(representation);
  }
  const rows=[];
  for(const [invoice,representations] of groups){
   if(blocked.has(invoice))continue;
   if(existing.some(row=>key(row.factura)===invoice)){messages.push(`La factura ${representations[0].invoice} ya está en tránsito; se conservan sus datos y no se duplica.`);continue;}
   const signature=items=>JSON.stringify([...items.entries()].map(([ref,row])=>[ref,row.cantidad]).sort((a,b)=>a[0].localeCompare(b[0])));
   if(representations.some(rep=>signature(rep.materials)!==signature(representations[0].materials))){messages.push(`Los packing de ${representations[0].invoice} no coinciden; revisa los originales antes de importar.`);continue;}
   for(const row of representations[0].materials.values())rows.push({...row,id:createId(),createdAt:timestamp,updatedAt:timestamp});
  }
  return {rows,messages,invoices:groups.size};
 }
 return {prepare};
});

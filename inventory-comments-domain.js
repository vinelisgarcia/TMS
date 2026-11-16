(function(root,factory){const api=factory(typeof module==='object'&&module.exports?require('./inventory-domain.js'):root.InventoryDomain,typeof module==='object'&&module.exports?require('./inventory-counts-domain.js'):root.InventoryCountDomain,typeof module==='object'&&module.exports?require('./inventory-trace-domain.js'):root.InventoryTraceDomain);if(typeof module==='object'&&module.exports)module.exports=api;else root.InventoryCommentsDomain=api;})(typeof globalThis==='object'?globalThis:this,function(D,C,T){
 'use strict';
 const clean=value=>String(value??'').replace(/\r\n?/g,'\n').trim(),norm=value=>clean(value).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,''),copy=value=>JSON.parse(JSON.stringify(value));
 const codeHeader=value=>/^(codigo|codigomaterial|codigodematerial|codigodearticulo|codigoarticulo|material|materialcode|itemcode|referencia|ref|sku|id|articulo|codigosap)(ct|ad)?$/.test(norm(value));
 const commentHeader=value=>/^(coment|observ|nota|analisis|motivo)/.test(norm(value))||norm(value)==='obs';
 function columns(headers){return {code:headers.findIndex(codeHeader),comments:headers.map((value,index)=>commentHeader(value)?index:-1).filter(index=>index>=0)};}
 function layout(matrix){let best={headerRow:0,code:-1,comments:[],score:-1};matrix.slice(0,40).forEach((headers,headerRow)=>{const found=columns(headers),score=(found.code>=0?5:0)+found.comments.length*3;if(score>best.score)best={...found,headerRow,score};});return best;}
 function owner(value){const text=clean(value).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,' '),plain=norm(value),ct=/\bct\b|caribe\s*trans/.test(text)||/^(comentarios?|observaciones?|notas?)ct$/.test(plain),ad=/\bad\b|alvarez|dominicana/.test(text)||/^(comentarios?|observaciones?|notas?)ad$/.test(plain);return ct&&ad?'CT / AD':ct?'CT':ad?'AD':'';}
 function parse(sheets){
  const rows=[],errors=[];
  for(const sheet of sheets.filter(value=>value.enabled)){
   if(!Number.isInteger(sheet.headerRow)||sheet.headerRow<0||sheet.headerRow>=sheet.matrix.length)throw Error('La fila de encabezados no es válida en '+sheet.name+'.');
   if(sheet.code<0||!sheet.comments.length)throw Error('Selecciona referencia y al menos una columna de comentarios en '+sheet.name+'.');
   const headers=sheet.matrix[sheet.headerRow]||[];
   sheet.matrix.slice(sheet.headerRow+1).forEach((values,index)=>{
    const row=index+sheet.headerRow+2,notes=sheet.comments.map(column=>({column,text:clean(values[column])})).filter(note=>note.text);if(!notes.length)return;
    try{const code=D.normalizeCode(values[sheet.code]);if(!code)throw Error('Falta la referencia.');if(/^(total|totales|grandtotal)$/.test(norm(code)))return;
     for(const note of notes)rows.push({code,text:note.text,owner:owner(headers[note.column])||owner(sheet.name)||'Análisis',field:clean(headers[note.column])||'Comentarios',sheet:sheet.name,row});
    }catch(error){errors.push({sheet:sheet.name,row,message:error.message});}
   });
  }
  if(errors.length){const error=Error('Revisa las referencias antes de importar.');error.invalidRows=errors;throw error;}
  if(!rows.length)throw Error('No se encontraron comentarios en las hojas y columnas seleccionadas.');return rows;
 }
 const signature=row=>JSON.stringify([row.code,row.owner||'Análisis',clean(row.text)]);
 function append(data,rows,{fileName='',user='',now=new Date().toISOString()}={}){
  if(data.articleComments&&data.articleComments.version!==1)throw Error('Actualiza el TMS para importar esta versión de comentarios.');
  const next=copy(data.articleComments||{version:1,byCode:{},imports:[]});next.byCode||={};next.imports||=[];const seen=new Set(Object.values(next.byCode).flat().map(signature)),known=new Set(C.catalog(data).map(row=>row.code));let added=0,duplicates=0;const unmatched=new Set();
  for(const row of rows){const code=D.normalizeCode(row.code),text=clean(row.text);if(!code||!text)throw Error('Cada comentario debe tener referencia y texto.');const entry={...row,code,text},key=signature(entry);if(!known.has(code))unmatched.add(code);if(seen.has(key)){duplicates++;continue;}
   (next.byCode[code]||=[]).push({...entry,id:typeof crypto!=='undefined'&&crypto.randomUUID?crypto.randomUUID():Date.now().toString(36)+'-'+Math.random().toString(36).slice(2),fileName,at:now,by:user});seen.add(key);added++;
  }
  if(added){next.imports.push({fileName,at:now,by:user,added,duplicates,unmatched:[...unmatched]});data.articleComments=next;}
  return {added,duplicates,unmatched:[...unmatched],references:new Set(rows.map(row=>row.code)).size};
 }
 function index(data){
  const result={},seen=new Set(),add=entry=>{const text=clean(entry.text);if(!text)return;const key=JSON.stringify([entry.code,entry.group,entry.owner||'',text]);if(seen.has(key))return;seen.add(key);(result[entry.code]||=[]).push({...entry,text});};
  for(const event of T.events(data))if(event.kind==='followup')add({id:event.id,code:event.code,text:event.fields?.note,owner:event.fields?.owner||'',at:event.at,by:event.by,group:'Seguimiento CT / AD'});
  for(const [code,value]of Object.entries(data.cases||{}))if(!code.includes(':'))add({code,text:value?.note,owner:value?.owner||'',at:value?.updatedAt,by:value?.updatedBy,group:'Seguimiento CT / AD'});
  for(const cycle of data.counting?.cycles||[]){
   for(const [code,notes]of Object.entries(cycle.followups||{}))for(const note of notes)add({id:note.id,code,text:note.notes,owner:note.responsible||'',at:note.at,by:note.by,group:'Apunte de conteo'});
   for(const record of cycle.records||[])if(!(cycle.followups?.[record.code]||[]).some(note=>note.id===record.differenceNoteId&&clean(note.notes)===clean(record.note)))add({id:record.id,code:record.code,text:record.note,at:record.at,by:record.countedBy||record.by,countedOn:C.recordDay(record),group:record.voided?'Conteo anulado':'Conteo físico'});
  }
  for(const [code,notes]of Object.entries(data.articleComments?.byCode||{}))for(const note of notes)add({...note,code,group:'Análisis importado'});
  for(const entries of Object.values(result))entries.sort((a,b)=>String(b.at||'').localeCompare(String(a.at||'')));return result;
 }
 return {layout,columns,parse,append,index};
});

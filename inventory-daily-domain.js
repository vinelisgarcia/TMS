(function(root,factory){const api=factory(typeof module==='object'&&module.exports?require('./inventory-counts-domain.js'):root.InventoryCountDomain,typeof module==='object'&&module.exports?require('./showroom-counts-domain.js'):root.ShowroomCountDomain,typeof module==='object'&&module.exports?require('./inventory-domain.js'):root.InventoryDomain);if(typeof module==='object'&&module.exports)module.exports=api;else root.InventoryDailyDomain=api;})(typeof globalThis==='object'?globalThis:this,function(C,S,D){
 'use strict';
 const clone=value=>JSON.parse(JSON.stringify(value)),clean=value=>String(value??'').trim(),id=()=>crypto.randomUUID();
 const domain=kind=>{if(!['ctad','showroom'].includes(kind))throw Error('Inventario no reconocido.');return kind==='showroom'?S:C;};
 function plan(data,report,{kind,day,scope='all',user='',templateId=id(),now=new Date().toISOString()}={}){
  const api=domain(kind);C.validateCountDay(day);
  const next=clone(data),catalog=api.catalog(next);if(!catalog.length)throw Error('Carga el inventario o el maestro antes de generar la plantilla.');
  const cycle=api.createCycle(next,report,{user,now}),selected=scope==='guided'?catalog.filter(item=>cycle.items[item.code]?.required):catalog;
  if(!selected.length)throw Error('No hay artículos obligatorios para esta plantilla.');
  const rows=[];
  for(const item of selected){
   if(!item.unitConflict)api.addItem(next,item.code,{catalogItem:item});
   const frozen=cycle.items[item.code],bins=frozen?.expectedLocations||(item.locations.length?item.locations:[{location:''}]);
   for(const bin of bins)rows.push({id:templateId+'-'+(rows.length+1),code:item.code,description:frozen?.description||item.description||'',expectedLocation:bin.location||'',unit:frozen?.unit||item.unit||'UND',unitConflict:!!item.unitConflict});
  }
  return {data:next,template:{version:1,kind,templateId,cycleId:cycle.id,sourcePair:cycle.sourcePair,sourceDates:cycle.sourceDates,day,generatedAt:now,rows}};
 }
 function excelDay(value,date1904=false){
  if(value instanceof Date)return C.validateCountDay(value.toISOString().slice(0,10));
  if(typeof value==='number'&&Number.isFinite(value)){const serial=Math.floor(value)+(date1904?1462:0);if(serial===60)throw Error('Fecha de Excel no válida.');const date=new Date(Date.UTC(1899,11,30)+(serial<60?serial+1:serial)*86400000);return C.validateCountDay(date.toISOString().slice(0,10));}
  const text=clean(value),match=text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);return C.validateCountDay(match?match[3]+'-'+match[2].padStart(2,'0')+'-'+match[1].padStart(2,'0'):text);
 }
 function parse(template,matrix,{date1904=false,formulaRows=[]}={}){
  domain(template.kind);if(template.version!==1||!template.templateId||!template.cycleId||!Array.isArray(template.rows))throw Error('La plantilla no es válida. Descárgala desde el TMS.');
  const lookup=new Map(template.rows.map(row=>[row.id,row])),seen=new Set(),rows=[],errors=[];let blank=0;
  matrix.slice(7).forEach((values,index)=>{
   const row=index+8;if(clean(values[6])===''){if(formulaRows.includes(row))errors.push({row,message:'Escribe la cantidad física como número, sin fórmula.'});else blank++;return;}
   try{
    if(formulaRows.includes(row))throw Error('Escribe la cantidad física como número, sin fórmula.');
    const original=lookup.get(clean(values[9]));if(!original)throw Error('Fila sin identificador de la plantilla. No agregues ni sustituyas referencias.');
    if(clean(values[1])!==original.code||C.location(values[3])!==C.location(original.expectedLocation)||clean(values[5])!==original.unit)throw Error('La referencia, unidad o ubicación prevista fue modificada. Corrige solo la ubicación contada.');
    const quantity=D.parseNumber(values[6],'auto');if(quantity===null||quantity<0)throw Error('La cantidad física debe ser mayor o igual a cero.');
    const countedOn=excelDay(values[0],date1904),countedBy=clean(values[7]);if(!countedBy)throw Error('Indica quién realizó el conteo.');
    const importKey=JSON.stringify([template.templateId,original.id,countedOn]);if(seen.has(importKey))throw Error('La misma fila de conteo está repetida para este día.');seen.add(importKey);
    const actualLocation=C.location(values[4]);if(original.expectedLocation&&!actualLocation)throw Error('Confirma o corrige la ubicación contada.');
    rows.push({row,code:original.code,expectedLocation:original.expectedLocation,actualLocation,quantity,countedOn,countedBy,note:clean(values[8]),importKey,rowId:original.id});
   }catch(error){errors.push({row,message:error.message});}
  });
  if(errors.length){const error=Error('Revisa '+errors.length+' filas antes de cargar el conteo.');error.invalidRows=errors;throw error;}
  if(!rows.length)throw Error('No hay cantidades físicas registradas. Las celdas vacías no cuentan como cero.');
  return {rows,blank,days:[...new Set(rows.map(row=>row.countedOn))].sort()};
 }
 const same=(record,row)=>['code','expectedLocation','actualLocation','quantity','countedOn','countedBy','note'].every(key=>(record[key]??'')===(row[key]??''));
 function apply(data,template,rows,{kind,user='',fileName='',now=new Date().toISOString()}={}){
  const api=domain(kind);if(template.kind!==kind)throw Error('La plantilla corresponde a otro inventario.');
  const next=clone(data),cycle=api.active(next);
  if(!cycle||cycle.id!==template.cycleId||cycle.sourcePair!==template.sourcePair||cycle.status!=='open')throw Error('La plantilla pertenece a un ciclo cerrado, archivado o distinto. Revisa el ciclo antes de cargar.');
  let inserted=0,skipped=0;
  for(const row of rows){
   const old=cycle.records.find(record=>record.importKey===row.importKey);
   if(old){if(!old.voided&&same(old,row)){skipped++;continue;}throw Error('Fila '+row.row+': este registro ya fue cargado y cambió. Corrígelo desde Resultados y bitácora.');}
   if(!cycle.items[row.code])throw Error('Fila '+row.row+': el artículo no tiene una referencia de conteo válida. Revisa sus unidades.');
   try{const record=api.recordCount(next,{...row,confirmed:true,unconfirmed:!row.actualLocation},{user,now});Object.assign(record,{importKey:row.importKey,templateId:template.templateId,templateRowId:row.rowId,sourceFile:fileName});inserted++;}
   catch(error){throw Error('Fila '+row.row+': '+error.message);}
  }
  if(inserted){(cycle.imports||=[]).push({id:id(),templateId:template.templateId,fileName,inserted,skipped,days:[...new Set(rows.map(row=>row.countedOn))].sort(),at:now,by:user});Object.assign(data,next);}
  return {inserted,skipped,days:[...new Set(rows.map(row=>row.countedOn))].sort()};
 }
 function history(data,day){
  return C.state(data).cycles.flatMap(cycle=>cycle.records.filter(row=>!day||C.recordDay(row)===day).map(record=>({...clone(record),countedOn:C.recordDay(record),description:cycle.items[record.code]?.description||'',unit:cycle.items[record.code]?.unit||'',cycleId:cycle.id,cycleStatus:cycle.status,current:C.itemReport(cycle,record.code).records.some(row=>row.id===record.id)}))).sort((a,b)=>b.countedOn.localeCompare(a.countedOn)||a.code.localeCompare(b.code));
 }
 return {plan,excelDay,parse,apply,history};
});

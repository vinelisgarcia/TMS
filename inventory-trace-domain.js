(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory();else root.InventoryTraceDomain=factory();})(typeof globalThis!=='undefined'?globalThis:this,function(){
 'use strict';
 const copy=value=>JSON.parse(JSON.stringify(value)),text=value=>String(value??'').trim(),round=value=>Number(value.toFixed(10));
 const uid=()=>typeof crypto!=='undefined'&&crypto.randomUUID?crypto.randomUUID():Date.now().toString(36)+'-'+Math.random().toString(36).slice(2);
 const now=options=>options?.now||new Date().toISOString(),today=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'America/Santo_Domingo',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
 const signature=(key,value)=>JSON.stringify([key,value.status||'',value.owner||'',value.document||'',value.note||'',value.updatedAt||'',value.updatedBy||'']);
 const pair=data=>[data.sources?.ct?.semanticHash||'ct-pendiente',data.sources?.ad?.semanticHash||'ad-pendiente'].join(':');
 function date(value){const s=text(value);if(!/^\d{4}-\d{2}-\d{2}$/.test(s))return '';const d=new Date(s+'T12:00:00Z');return Number.isFinite(d.getTime())&&d.toISOString().slice(0,10)===s?s:'';}
 function source(snapshot,code,lookup){const material=lookup?lookup.get(code):snapshot?.materials?.find(row=>row.code===code);return {hash:snapshot?.semanticHash||'',file:snapshot?.fileName||'',asOf:snapshot?.asOf||'',importedAt:snapshot?.importedAt||'',present:!!material,quantity:snapshot?(material?.quantity??0):null,unit:material?.unit||'',unitConflict:!!material?.unitConflict};}
 function point(data,code,cache){const c=cache?cache.ct.get(code):data.sources?.ct?.materials?.find(row=>row.code===code),a=cache?cache.ad.get(code):data.sources?.ad?.materials?.find(row=>row.code===code),ct=source(data.sources?.ct,code,cache?.ct),ad=source(data.sources?.ad,code,cache?.ad),conflict=c?.unitConflict||a?.unitConflict||c&&a&&c.unit!==a.unit;return {code,description:a?.description||c?.description||code,unit:conflict?'':c?.unit||a?.unit||'',difference:data.sources?.ct&&data.sources?.ad&&(c||a)&&!conflict?round(ad.quantity-ct.quantity):null,ct,ad};}
 function legacyPoint(data,parts){
  const code=parts.slice(2).join(':'),current=point(data,code),read=(name,hash)=>{if(data.sources?.[name]?.semanticHash===hash)return source(data.sources[name],code);const log=(data.settings?.importLog||[]).find(row=>row.source===name&&row.semanticHash===hash);return {hash,file:log?.fileName||'',asOf:log?.asOf||'',importedAt:log?.at||'',present:null,quantity:null,unit:'',unitConflict:false};};
  return {code,description:current.description,unit:'',difference:null,ct:read('ct',parts[0]),ad:read('ad',parts[1])};
 }
 function events(data){
  const stored=data.tracking?.events||[],known=new Set(stored.filter(event=>event.kind==='followup').map(event=>event.origin));
  const legacy=Object.entries(data.cases||{}).flatMap(([key,value])=>{const parts=key.split(':');if(parts.length<3||!value||known.has(signature(key,value)))return [];return [{id:'legacy:'+signature(key,value),kind:'followup',code:parts.slice(2).join(':'),origin:signature(key,value),sourcePair:parts.slice(0,2).join(':'),point:legacyPoint(data,parts),fields:copy(value),at:value.updatedAt||'',by:value.updatedBy||'',legacy:true}];});
  return [...stored,...legacy];
 }
 function ensure(data){if(data.tracking&&data.tracking.version!==1)throw new Error('Actualiza el TMS para editar esta versión del historial.');const all=events(data);if(!data.tracking)data.tracking={version:1,events:[]};data.tracking.events=all;return data.tracking.events;}
 function append(data,event,options={}){
  const all=events(data),previous=all.find(row=>row.id===event.id);
  if(previous){
   const comparable=row=>JSON.stringify({...row,at:'',by:''});if(comparable(previous)===comparable(event))return previous;
   if(options.replacePending?.includes(event.id)){const list=ensure(data),index=list.findIndex(row=>row.id===event.id);list[index]={...event,at:previous.at,by:previous.by};return list[index];}
   if(JSON.stringify(previous)===JSON.stringify(event))return previous;
   throw new Error('Este registro ya fue guardado. Para corregirlo, anula la anotación y registra el movimiento correcto.');
  }
  ensure(data).push(event);return event;
 }
 function recordCase(data,code,fields,options={}){
  const values=Object.fromEntries(['status','owner','document','note'].map(key=>[key,text(fields[key])])),key=pair(data)+':'+code,previous=data.cases?.[key];
  if(!['pendiente','en_revision','rectificado','justificado'].includes(values.status))throw new Error('Estado no válido.');
  if(['rectificado','justificado'].includes(values.status)&&!values.note)throw new Error('Añade una nota de revisión para documentar el seguimiento.');
  if(previous&&Object.keys(values).every(key=>text(previous[key])===values[key]))return null;
  const at=now(options),value={...values,updatedAt:at,updatedBy:options.user||''},event={id:options.id||uid(),kind:'followup',code,sourcePair:pair(data),point:point(data,code),fields:value,origin:signature(key,value),at,by:options.user||''};
  append(data,event);data.cases=data.cases||{};data.cases[key]=copy(value);return event;
 }
 function register(data,code,input,options={}){
  const p=point(data,code),quantity=Number(input.quantity),effectiveDate=date(input.effectiveDate),type=text(input.type),owner=text(input.owner),reason=text(input.reason);
  if(!p.ad.hash||!p.ct.hash||data.sources.ad.valid===false||data.sources.ct.valid===false)throw new Error('Carga ambos inventarios válidos antes de registrar un movimiento.');
  if(!Number.isFinite(p.ad.quantity)||!Number.isFinite(p.ct.quantity))throw new Error('Revisa las cantidades de referencia del material antes de registrar el movimiento.');
  if(!['entrada','salida'].includes(type))throw new Error('Selecciona entrada o salida en AD.');
  if(!Number.isFinite(quantity)||quantity<=0||quantity>Number.MAX_SAFE_INTEGER)throw new Error('Indica una cantidad mayor que cero.');
  if(!effectiveDate||effectiveDate>(options.today||today()))throw new Error('Indica la fecha real del movimiento en AD, sin fechas futuras.');
  if(!owner)throw new Error('Indica el responsable del movimiento.');if(!reason)throw new Error('Explica el motivo del movimiento en las notas de revisión.');
  if(!p.unit||p.difference===null||p.ad.unitConflict||p.ct.unitConflict)throw new Error('Corrige las unidades del material antes de registrar el movimiento.');
  const event={id:input.id||uid(),kind:'movement',system:'AD',code,description:p.description,unit:p.unit,type,quantity:round(quantity),signedQuantity:round((type==='entrada'?1:-1)*quantity),effectiveDate,afterBaseCut:effectiveDate>p.ad.asOf||effectiveDate===p.ad.asOf&&input.afterBaseCut===true,owner,reason,document:text(input.document),point:p,at:now(options),by:options.user||''};
  const previous=events(data).find(row=>row.id===event.id);
  if(previous){const stable=row=>JSON.stringify({...row,at:'',by:''});if(stable(previous)===stable(event))return previous;}
  return append(data,event,options);
 }
 function movements(data,code){return events(data).filter(event=>event.kind==='movement'&&(!code||event.code===code));}
 function isVoided(data,id){return events(data).some(event=>event.kind==='void'&&event.movementId===id);}
 function index(data){
  const all=events(data),records=all.filter(event=>event.kind==='movement'),voids=new Set(all.filter(event=>event.kind==='void').map(event=>event.movementId)),reviews=new Map(),groups=new Map();
  for(const event of all)if(event.kind==='review')for(const id of event.movementIds)reviews.set(id,event);
  for(const movement of records)if(!voids.has(movement.id)&&movement.afterBaseCut){const key=JSON.stringify([movement.code,movement.point.ad.hash]);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(movement);}
  return {all,records,voids,reviews,groups,ct:new Map((data.sources?.ct?.materials||[]).map(row=>[row.code,row])),ad:new Map((data.sources?.ad?.materials||[]).map(row=>[row.code,row]))};
 }
 function assessment(data,movement,cache=index(data)){
  const ad=source(data.sources?.ad,movement.code,cache.ad),base=movement.point.ad,review=cache.reviews.get(movement.id),result={movementIds:[movement.id],ad,base,expected:null,delta:null,status:'awaiting_report',review};
  if(cache.voids.has(movement.id))return {...result,status:'voided'};
  if(review?.result==='verified')return {...result,status:'verified'};
  if(!ad.hash||ad.hash===base.hash||!date(ad.asOf)||ad.asOf<base.asOf||ad.asOf<movement.effectiveDate)return result;
  if(!movement.afterBaseCut)return {...result,status:review?.result==='review'?'review':'documentary'};
  if(ad.unitConflict||ad.present&&ad.unit!==movement.unit)return {...result,status:'unit_conflict'};
  const group=(cache.groups.get(JSON.stringify([movement.code,base.hash]))||[]).filter(row=>row.effectiveDate<=ad.asOf);
  const expected=round(base.quantity+group.reduce((sum,row)=>sum+row.signedQuantity,0)),delta=round(ad.quantity-expected);
  return {...result,movementIds:group.map(row=>row.id),expected,delta,status:review?.result==='review'?'review':Math.abs(delta)<1e-8?'compatible':'different'};
 }

 function checks(data,options={}){
  const cache=index(data),added=[],seen=new Set(cache.all.map(event=>event.id));
  for(const movement of cache.records){
   const result=assessment(data,movement,cache);if(!['compatible','different','documentary','unit_conflict','review'].includes(result.status))continue;
   const id='check:'+result.ad.hash+':'+result.movementIds.slice().sort().join('|');if(seen.has(id))continue;seen.add(id);
   const event={id,kind:'check',code:movement.code,movementIds:result.movementIds,point:point(data,movement.code,cache),base:result.base,expected:result.expected,delta:result.delta,result:result.status,at:now(options),by:options.user||''};append(data,event);added.push(event);
  }
  return added;
 }
 function review(data,movementId,input,options={}){
  const movement=movements(data).find(row=>row.id===movementId);if(!movement||isVoided(data,movementId))throw new Error('El movimiento ya no está disponible para revisar.');
  const result=assessment(data,movement),note=text(input.note);if(!['verified','review'].includes(input.result)||!note)throw new Error('Selecciona el resultado y explica la revisión.');
  if(!result.ad.hash||result.ad.hash===result.base.hash||result.ad.asOf<movement.effectiveDate||result.ad.asOf<result.base.asOf)throw new Error('Carga un nuevo reporte AD cuyo corte incluya la fecha del movimiento antes de revisarlo.');
  const event={id:input.id||uid(),kind:'review',code:movement.code,movementIds:[movementId],basisMovementIds:result.movementIds,result:input.result,note,point:point(data,movement.code),base:result.base,expected:result.expected,delta:result.delta,at:now(options),by:options.user||''};
  return append(data,event,options);
 }
 function voidMovement(data,movementId,input,options={}){
  const movement=movements(data).find(row=>row.id===movementId);if(!movement)throw new Error('Movimiento no encontrado.');const note=text(input.note);if(!note)throw new Error('Indica por qué anulas esta anotación.');
  const existing=events(data).find(event=>event.kind==='void'&&event.movementId===movementId);if(existing)return existing;
  return append(data,{id:input.id||uid(),kind:'void',code:movement.code,movementId,note,at:now(options),by:options.user||''},options);
 }
 function summary(data,cache=index(data)){
  const records=cache.records,active=records.filter(row=>!cache.voids.has(row.id)),states=active.map(row=>assessment(data,row,cache).status),materials=new Map();
  active.forEach(row=>{const key=JSON.stringify([row.code,row.unit]);if(!materials.has(key))materials.set(key,{code:row.code,description:row.description,unit:row.unit,entries:0,exits:0,net:0,movements:0});const item=materials.get(key);item.entries=round(item.entries+(row.type==='entrada'?row.quantity:0));item.exits=round(item.exits+(row.type==='salida'?row.quantity:0));item.net=round(item.net+row.signedQuantity);item.movements++;});
  return {records:records.length,active:active.length,voided:records.length-active.length,verified:states.filter(status=>status==='verified').length,compatible:states.filter(status=>status==='compatible').length,pending:states.filter(status=>status!=='verified').length,materials:[...materials.values()].sort((a,b)=>a.code.localeCompare(b.code))};
 }
 function analyze(data){const cache=index(data);return {events:cache.all,movements:cache.records,summary:summary(data,cache),assessments:Object.fromEntries(cache.records.map(row=>[row.id,assessment(data,row,cache)]))};}
 return {today,date,point,events,analyze,freezeLegacy:data=>ensure(data),recordCase,register,movements,assessment,checks,review,voidMovement,isVoided,summary};
});

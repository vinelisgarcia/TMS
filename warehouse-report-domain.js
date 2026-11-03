(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.TMSWarehouseReportDomain=api;})(typeof globalThis!=='undefined'?globalThis:this,function(){
 'use strict';
 const text=value=>String(value??'').trim(),key=value=>text(value).toUpperCase(),num=value=>Number.isFinite(Number(value))?Number(value):0,unique=values=>[...new Set(values.filter(Boolean))];
 const requestKey=item=>JSON.stringify([item.codigo,item.solicitud,item.pedidoCliente,item.lineaPedidoCliente,item.articulo].map(key));
 const identity=item=>JSON.stringify([key(item.codigo),key(item.solicitud)]);
 function iso(value){
  const valueText=text(value);let year,month,day,match;
  if((match=valueText.match(/^(\d{2})\/(\d{2})\/(\d{4})$/))){[,day,month,year]=match;}
  else if((match=valueText.match(/^(\d{4})-(\d{2})-(\d{2})$/))){[,year,month,day]=match;}
  else return '';
  const result=year+'-'+month+'-'+day,date=new Date(result+'T12:00:00Z');
  return Number.isFinite(date.getTime())&&date.toISOString().slice(0,10)===result?result:'';
 }

 const dateLabel=value=>{const s=iso(value);return s?s.slice(8)+'/'+s.slice(5,7)+'/'+s.slice(0,4):'';};
 function resolveSchedule(item,lines){
  const matches=(lines||[]).filter(line=>!line.archivedAt&&key(line.clienteId)===key(item.codigo)&&key(line.pedidoCliente)===key(item.pedidoCliente)&&(!text(item.lineaPedidoCliente)||key(line.lineaPedidoCliente)===key(item.lineaPedidoCliente))&&(!text(item.articulo)||key(line.articulo)===key(item.articulo)));
  const schedules=new Map();let unplanned=false;
  for(const line of matches){const date=dateLabel(line.fechaPlanificada);if(!date){unplanned=true;continue;}const schedule={date,iso:iso(date),truck:text(line.camionAsignado)||'Sin camión',route:text(line.rutaNombre||line.zona)||'Sin ruta'};schedules.set(JSON.stringify(schedule),schedule);}
  const entries=[...schedules.values()].sort((a,b)=>a.iso.localeCompare(b.iso)||a.truck.localeCompare(b.truck)||a.route.localeCompare(b.route)),dates=unique(entries.map(row=>row.date));
  return {entries,dates,unplanned:unplanned||!entries.length,status:!entries.length?'Sin programar en calendario':unplanned?'Programación parcial':entries.length>1?'Varias entregas':'Programada',label:dates.length?dates.join(' · ')+(unplanned?' · Parte sin programar':''):'Sin programar en calendario'};
 }
 function rows(context){
  const plan=new Map((context.plan||[]).map(item=>[requestKey(item),item])),control=new Map((context.control||[]).map(item=>[requestKey(item),item])),mode=context.mode||'programadas',source=mode==='completadas'&&control.size?control:plan.size?plan:control,keys=new Set(source.keys());
  if(!keys.size)(context.legacy||[]).forEach(item=>{plan.set(requestKey(item),item);keys.add(requestKey(item));});
  const search=text(context.search).toLocaleLowerCase();
  return [...keys].map(id=>{
   const p=plan.get(id),c=control.get(id),item={...p,...c},requested=num(p?.cantidadSolicitada??c?.cantidadSolicitada),processed=num(c?.cantidadCompletada??p?.cantidadCompletada),pending=Math.max(requested-processed,0),schedule=resolveSchedule(item,context.lineItems);
   const client=(context.lineItems||[]).find(line=>key(line.clienteId)===key(item.codigo));
   return {...item,reportKey:id,cantidadSolicitada:requested,cantidadCompletada:processed,cantidadPendiente:pending,estadoVista:pending<=0?'Completada':processed>0?'Parcial':'Pendiente',schedule,fechaProgramada:schedule.label,clienteNombreVista:text(item.cliente||client?.clienteNombre||item.codigo)||'Sin cliente',motivoNoPreparada:text(c?.motivoNoPreparada||p?.motivoNoPreparada)};
  }).filter(item=>(mode!=='pendientes'||item.estadoVista!=='Completada')&&(mode!=='completadas'||item.estadoVista==='Completada')).filter(item=>!search||[item.codigo,item.clienteNombreVista,item.solicitud,item.pedidoCliente,item.articulo,item.descripcionArticulo].some(value=>text(value).toLocaleLowerCase().includes(search))).sort((a,b)=>(a.schedule.entries[0]?.iso||'9999').localeCompare(b.schedule.entries[0]?.iso||'9999')||a.clienteNombreVista.localeCompare(b.clienteNombreVista)||text(a.solicitud).localeCompare(text(b.solicitud)));
 }
 function build(context,at){
  const lines=rows(context),requests=new Map(),agenda=new Map();
  for(const line of lines){const id=identity(line);if(!requests.has(id))requests.set(id,[]);requests.get(id).push(line);
   const entries=[...line.schedule.entries,...(line.schedule.unplanned?[{date:'Sin programar',iso:'',truck:'Sin asignar',route:'Sin programación confirmada'}]:[])];
   for(const entry of entries){const id=JSON.stringify(entry);if(!agenda.has(id))agenda.set(id,{...entry,requests:new Set(),lines:0,completed:0,pending:0});const group=agenda.get(id);group.requests.add(identity(line));group.lines++;if(line.estadoVista==='Completada')group.completed++;else group.pending++;}
  }
  const metrics={requests:requests.size,customers:unique(lines.map(line=>key(line.codigo))).length,orders:unique(lines.map(line=>JSON.stringify([key(line.codigo),key(line.pedidoCliente)]))).length,lines:lines.length,completed:lines.filter(line=>line.estadoVista==='Completada').length,partial:lines.filter(line=>line.estadoVista==='Parcial').length,pending:lines.filter(line=>line.estadoVista==='Pendiente').length,unscheduled:lines.filter(line=>line.schedule.unplanned).length,multiple:lines.filter(line=>line.schedule.entries.length>1).length,withoutReason:lines.filter(line=>line.estadoVista!=='Completada'&&!line.motivoNoPreparada).length,completedRequests:[...requests.values()].filter(items=>items.every(line=>line.estadoVista==='Completada')).length};metrics.completion=metrics.lines?metrics.completed/metrics.lines:0;
  return {title:'Solicitudes a almacén',subtitle:({programadas:'Todas las solicitudes',pendientes:'Pendientes de preparar',completadas:'Líneas completadas'})[context.mode]||'Todas las solicitudes',search:text(context.search),author:text(context.author)||'Usuario',at:at||new Date().toISOString(),files:unique(context.files||[]),lines,metrics,agenda:[...agenda.values()].sort((a,b)=>(a.iso||'9999').localeCompare(b.iso||'9999')||a.truck.localeCompare(b.truck)||a.route.localeCompare(b.route)).map(group=>({...group,requests:group.requests.size})),criteria:['Fechas de entrega tomadas del calendario vigente al generar el informe. No se utilizan las fechas del archivo importado ni se asigna una fecha a pedidos sin programación.','Cruce por cliente, pedido, línea y material. Las entregas parciales muestran todas sus programaciones; no se elige arbitrariamente una fecha.','La planificación y el control se cruzan sin duplicar líneas. El control más reciente aporta la preparación y la planificación conserva la cantidad solicitada.','El porcentaje de preparación cuenta líneas completas sobre el total de líneas del filtro. Las cantidades se muestran por material. La agenda cuenta referencias a despachos: una solicitud con varias entregas puede figurar en más de una fecha; los indicadores cuentan solicitudes únicas.']};
 }
 return {resolveSchedule,rows,build,iso,dateLabel};
});

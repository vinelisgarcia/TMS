(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.InventoryCountDomain=api;})(typeof globalThis==='object'?globalThis:this,function(){
  'use strict';
  const round=value=>Number(value.toFixed(10));
  const copy=value=>JSON.parse(JSON.stringify(value));
  const clean=value=>String(value??'').trim();
  const location=value=>clean(value).toUpperCase();
  const key=(...parts)=>JSON.stringify(parts);
  const stamp=options=>options?.now||new Date().toISOString();
  const uid=()=>typeof crypto!=='undefined'&&crypto.randomUUID?crypto.randomUUID():Date.now().toString(36)+'-'+Math.random().toString(36).slice(2);
  const state=data=>data.counting||{version:1,catalog:{},positions:{},cycles:[],activeCycleId:null};
  function ensure(data){if(!data.counting)data.counting=state(data);const s=data.counting;s.catalog=s.catalog||{};s.positions=s.positions||{};s.cycles=s.cycles||[];return s;}
  const pair=data=>[data.sources?.ct?.semanticHash||'ct-pendiente',data.sources?.ad?.semanticHash||'ad-pendiente'].join(':');
  function priorities(report){
    const rows=report?.ready?report.rows:[];
    const positive=rows.filter(row=>row.difference>0&&row.signedImpactDop!==null).sort((a,b)=>b.signedImpactDop-a.signedImpactDop||b.difference-a.difference||a.code.localeCompare(b.code)).slice(0,10);
    const negative=rows.filter(row=>row.difference<0&&row.signedImpactDop!==null).sort((a,b)=>a.signedImpactDop-b.signedImpactDop||a.difference-b.difference||a.code.localeCompare(b.code)).slice(0,20);
    const unvalued=rows.filter(row=>row.difference>0&&row.impactDop===null).sort((a,b)=>b.difference-a.difference||a.code.localeCompare(b.code));
    return {positive,negative,unvalued,mandatory:[...positive.map(row=>({...row,reason:'Top 10 sobrantes'})),...negative.map(row=>({...row,reason:'Top 20 faltantes'}))]};
  }
  function catalog(data){
    const s=state(data),map=new Map();
    for(const source of ['ct','ad'])for(const row of data.sources?.[source]?.materials||[]){
      if(!map.has(row.code))map.set(row.code,{code:row.code,description:row.description||'',unit:row.unit||'UND',stockCT:null,stockAD:null,adRegistered:false,costAD:null,barcodes:[],supplier:'',subfamily:'',locations:[]});
      const item=map.get(row.code);item.unitConflict=!!(item.unitConflict||row.unitConflict||(item.unit&&row.unit&&item.unit!==row.unit));if(source==='ct')item.stockCT=row.quantity;
      else {item.stockAD=row.quantity;item.adRegistered=true;item.costAD=row.quantity>0&&row.totalValueDop!==null?row.totalValueDop/row.quantity:null;if(row.description)item.description=row.description;}
      const bins=[...(row.locations||[]).map(bin=>typeof bin==='string'?bin:bin.location),...(row.details||[]).map(detail=>detail.location)].filter(Boolean).map(location);
      if(source==='ct'||!item.locations.length)item.locations=[...new Set([...item.locations.map(bin=>bin.location),...bins])].map(bin=>({location:bin,stock:null}));
    }
    for(const [code,master] of Object.entries(s.catalog)){
      if(!map.has(code))map.set(code,{code,stockCT:null,stockAD:null,adRegistered:false,costAD:null,locations:[],barcodes:[]});
      const item=map.get(code);Object.assign(item,{description:master.description||item.description||'',unit:master.unit||item.unit||'UND',barcodes:master.barcodes||[],supplier:master.supplier||'',subfamily:master.subfamily||'',masterStock:master.stock??null,masterCost:master.cost??null});
    }
    for(const code of Object.keys(s.positions))if(!map.has(code))map.set(code,{code,description:'',unit:'UND',stockCT:null,stockAD:null,adRegistered:false,costAD:null,locations:[],barcodes:[]});
    for(const item of map.values()){
      if(Object.prototype.hasOwnProperty.call(s.positions,item.code))item.locations=copy(s.positions[item.code]);
      item.locations=[...new Map(item.locations.filter(bin=>location(bin.location)).map(bin=>[location(bin.location),{location:location(bin.location),stock:bin.stock??null}])).values()].sort((a,b)=>a.location.localeCompare(b.location));
    }
    return [...map.values()].sort((a,b)=>a.code.localeCompare(b.code));
  }
  function resolve(data,input){const value=clean(input);const rows=catalog(data);const exact=rows.find(row=>row.code.toUpperCase()===value.toUpperCase());if(exact)return exact;const hits=rows.filter(row=>(row.barcodes||[]).includes(value));if(hits.length>1)throw new Error('El código de barras corresponde a varios artículos. Selecciona el código de material.');return hits[0]||null;}
  const active=data=>state(data).cycles.find(cycle=>cycle.id===state(data).activeCycleId)||null;
  function snapshot(item,reason,required){return {code:item.code,description:item.description||'',unit:item.unit||'UND',stockCT:item.stockCT,stockAD:item.stockAD,adRegistered:!!item.adRegistered,costAD:item.costAD,masterStock:item.masterStock??null,expectedLocations:item.locations.length?copy(item.locations):[{location:'',stock:null}],required:!!required,reasons:reason?[reason]:[]};}
  function createCycle(data,report,options={}){
    const s=ensure(data),previous=active(data);
    if(previous&&previous.status==='open'&&previous.sourcePair===pair(data))return previous;
    if(previous&&previous.status==='open'){previous.status='archived';previous.archivedAt=stamp(options);}
    const cycle={id:options.id||uid(),sourcePair:pair(data),sourceDates:{ct:data.sources?.ct?.asOf||null,ad:data.sources?.ad?.asOf||null},createdAt:stamp(options),createdBy:options.user||'',status:'open',items:{},records:[],audit:[],processed:{}};
    const lookup=new Map(catalog(data).map(item=>[item.code,item]));
    for(const row of priorities(report).mandatory){const item=lookup.get(row.code);if(item)cycle.items[row.code]=snapshot(item,row.reason,true);}
    s.cycles.push(cycle);s.activeCycleId=cycle.id;return cycle;
  }
  function addItem(data,code,options={}){
    const cycle=active(data);if(!cycle||cycle.status!=='open')throw new Error('Inicia un ciclo de conteo abierto.');
    if(!cycle.items[code]&&cycle.sourcePair!==pair(data))throw new Error('El ciclo pertenece a otro corte. Inicia un nuevo ciclo antes de añadir artículos nuevos.');
    if(!cycle.items[code]){const item=catalog(data).find(row=>row.code===code);if(!item)throw new Error('Artículo no encontrado en Inventory Data.');if(item.unitConflict)throw new Error('El artículo tiene unidades incompatibles. Corrige la fuente antes de contarlo.');cycle.items[code]=snapshot(item,options.reason||'Conteo libre',!!options.required);}
    else if(options.required){cycle.items[code].required=true;if(options.reason&&!cycle.items[code].reasons.includes(options.reason))cycle.items[code].reasons.push(options.reason);}
    return cycle.items[code];
  }
  function recordsFor(cycle,code){return cycle.records.filter(record=>record.code===code&&!record.voided);}
  function itemReport(cycle,code){
    const item=cycle.items[code],records=recordsFor(cycle,code);
    const missing=item.expectedLocations.filter(bin=>!records.some(record=>record.expectedLocation===bin.location)).map(bin=>bin.location||'Sin ubicación asignada');
    const complete=records.length>0&&missing.length===0;
    const quantity=records.length?round(records.reduce((sum,record)=>sum+record.quantity,0)):null;
    const referenceStock=item.stockCT??item.stockAD??item.masterStock??null;
    const referenceSource=item.stockCT!==null?'CT':item.stockAD!==null?'AD':'Maestro';
    const referenceDifference=complete&&referenceStock!==null?round(quantity-referenceStock):null;
    const ctDifference=complete&&item.stockCT!==null?round(quantity-item.stockCT):null;
    const adDifference=complete&&item.stockAD!==null?round(quantity-item.stockAD):null;
    return {...copy(item),records:copy(records),missing,complete,quantity,referenceStock,referenceSource,referenceDifference,ctDifference,adDifference,processed:!!cycle.processed[code],readyForAD:complete&&quantity>0&&!item.adRegistered&&!!cycle.processed[code]};
  }
  function cycleReport(cycle){return cycle?Object.keys(cycle.items).map(code=>itemReport(cycle,code)):[];}
  function tasks(cycle){
    if(!cycle)return [];
    return Object.values(cycle.items).flatMap(item=>item.expectedLocations.map(bin=>({code:item.code,description:item.description,expectedLocation:bin.location,required:item.required,reasons:item.reasons,complete:recordsFor(cycle,item.code).some(record=>record.expectedLocation===bin.location)}))).sort((a,b)=>Number(b.required)-Number(a.required)||(a.expectedLocation||'\uffff').localeCompare(b.expectedLocation||'\uffff')||a.code.localeCompare(b.code));
  }
  function recordCount(data,input,options={}){
    const cycle=active(data);if(!cycle||cycle.status!=='open')throw new Error('El ciclo no está abierto.');
    const item=cycle.items[input.code];if(!item)throw new Error('Añade el artículo al ciclo antes de contar.');
    const expectedLocation=location(input.expectedLocation);
    if(!item.expectedLocations.some(bin=>bin.location===expectedLocation))throw new Error('La ubicación de referencia no pertenece a este ciclo.');
    if(!Number.isFinite(input.quantity)||input.quantity<0)throw new Error('Indica una cantidad física válida, mayor o igual que cero.');
    if(!input.confirmed)throw new Error('Confirma la ubicación del artículo o escribe la correcta.');
    const actualLocation=location(input.actualLocation);if(!actualLocation)throw new Error('El artículo no tiene ubicación: escribe la ubicación real.');
    if(recordsFor(cycle,input.code).some(row=>row.actualLocation===actualLocation)&&!input.additional)throw new Error('Ya hay un conteo en esta ubicación. Confirma que se trata de otra caja o corrige el registro existente.');
    const record={id:options.id||uid(),code:input.code,expectedLocation,actualLocation,locationConfirmed:true,locationCorrected:expectedLocation!==actualLocation,quantity:input.quantity,note:clean(input.note),at:stamp(options),by:options.user||'',voided:false};
    cycle.records.push(record);delete cycle.processed[input.code];
    if(record.locationCorrected){
      const s=ensure(data),existing=s.positions[input.code]||copy(item.expectedLocations);
      s.positions[input.code]=[...new Map(existing.filter(bin=>bin.location&&bin.location!==expectedLocation&&bin.correctedFrom!==expectedLocation).concat([{location:actualLocation,stock:null,correctedFrom:expectedLocation}]).map(bin=>[bin.location,bin])).values()];
      cycle.audit.push({id:uid(),recordId:record.id,code:record.code,type:'ubicacion',from:expectedLocation,to:actualLocation,reason:record.note||'Ubicación real confirmada en conteo',at:record.at,by:record.by});
    }
    return record;
  }
  function correctRecord(data,recordId,quantity,reason,options={}){
    const cycle=active(data);if(!cycle||cycle.status!=='open')throw new Error('El ciclo no está abierto.');
    const record=cycle.records.find(row=>row.id===recordId&&!row.voided);if(!record)throw new Error('Conteo no encontrado.');
    if(!clean(reason))throw new Error('Indica el motivo de la corrección.');
    if(quantity!==null&&(!Number.isFinite(quantity)||quantity<0))throw new Error('Cantidad física no válida.');
    cycle.audit.push({id:uid(),recordId,code:record.code,type:quantity===null?'anulacion':'cantidad',from:record.quantity,to:quantity,reason:clean(reason),at:stamp(options),by:options.user||''});
    if(quantity===null)record.voided=true;else record.quantity=quantity;
    delete cycle.processed[record.code];return record;
  }
  function processItem(data,code,options={}){const cycle=active(data);if(!cycle||cycle.status!=='open')throw new Error('El ciclo no está abierto.');if(!itemReport(cycle,code).complete)throw new Error('Completa todas las ubicaciones antes de procesar este artículo.');cycle.processed[code]={at:stamp(options),by:options.user||''};}
  function closeCycle(data,options={}){const cycle=active(data);if(!cycle||cycle.status!=='open')throw new Error('El ciclo no está abierto.');const rows=cycleReport(cycle);if(rows.some(row=>row.required&&(!row.complete||!row.processed)))throw new Error('Hay conteos obligatorios pendientes de completar o procesar.');if(rows.some(row=>row.records.length&&(!row.complete||!row.processed)))throw new Error('Completa y procesa los artículos ya iniciados.');cycle.status='closed';cycle.closedAt=stamp(options);return cycle;}
  function importMaster(data,matrix,mapping,kind,options={}){
    if(mapping.code===undefined)throw new Error('Selecciona la columna de código.');
    if(kind==='positions'&&mapping.location===undefined)throw new Error('Selecciona la columna de ubicación.');
    const parse=options.parseNumber||Number,start=options.headerRow??0,updates={},positions={},seen=new Map();let rows=0;
    for(let index=start+1;index<matrix.length;index++){
      const row=matrix[index];if(!row.some(value=>clean(value)))continue;
      const rawCode=row[mapping.code];if(typeof rawCode==='number'&&!Number.isSafeInteger(rawCode))throw new Error('Código numérico no seguro: guárdalo como texto.');const code=clean(rawCode).toUpperCase();if(!code)throw new Error('Fila '+(index+1)+': falta el código.');
      const get=field=>mapping[field]===undefined?'':row[mapping[field]];
      if(kind==='positions'){
        const bin=location(get('location'));if(!bin)throw new Error('Fila '+(index+1)+': falta la ubicación.');
        const stock=clean(get('stock'))===''?null:parse(get('stock'));if(stock!==null&&!Number.isFinite(stock))throw new Error('Stock de ubicación no válido.');
        const id=key(code,bin),prior=seen.get(id);seen.set(id,{code,location:bin,stock:prior?(prior.stock===null||stock===null?null:prior.stock+stock):stock});
      }else{
        const old=updates[code]||state(data).catalog[code]||{},item={...copy(old),code};
        for(const field of ['description','supplier','subfamily'])if(mapping[field]!==undefined)item[field]=clean(get(field));
        if(mapping.barcode!==undefined)item.barcodes=[...new Set([...(item.barcodes||[]),clean(get('barcode'))].filter(Boolean))];
        for(const field of ['stock','cost'])if(mapping[field]!==undefined){const value=clean(get(field))===''?null:parse(get(field));if(value!==null&&(!Number.isFinite(value)||(field==='cost'&&value<0)))throw new Error('Fila '+(index+1)+': '+field+' no válido.');if(updates[code]&&updates[code][field]!==undefined&&updates[code][field]!==value)throw new Error('El maestro repite '+code+' con '+field+' distinto. Usa CT/AD para sumar partidas de stock.');item[field]=value;}
        updates[code]=item;
      }
      rows++;
    }
    if(!rows)throw new Error('No se encontraron filas para importar.');
    for(const row of seen.values())(positions[row.code]||=[]).push({location:row.location,stock:row.stock});
    const s=ensure(data);Object.assign(s.catalog,updates);Object.assign(s.positions,positions);
    return {rows,articles:Object.keys(kind==='positions'?positions:updates).length,locations:seen.size};
  }
  return {state,ensure,pair,priorities,catalog,resolve,active,createCycle,addItem,itemReport,cycleReport,tasks,recordCount,correctRecord,processItem,closeCycle,importMaster,location};
});

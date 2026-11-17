(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.TMSEUR1Domain=api;})(typeof self!=='undefined'?self:globalThis,function(){
 'use strict';
 const VERSION=1;
 const kinds={certificate:'Certificado EUR.1',proposal:'Datos EUR.1',invoice:'Factura',packing:'Packing list',hs:'Resumen HS',complement:'Complementario',unknown:'Por identificar',archive:'ZIP original'};
 const norm=value=>String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[_./()\-]+/g,' ').replace(/\s+/g,' ').trim();
 const text=value=>String(value?.v??value?.value??value?.text??value??'').trim();
 const value=field=>field?.value??null;
 const field=(v,source,extra={})=>v===null||v===undefined||v===''?null:{value:v,kind:source?.cachedFormula?'cached_formula':'declared',source,...extra};
 const calculated=(v,sources)=>({value:v,kind:'calculated',sources:sources.filter(Boolean)});
 const sourceLabel=s=>!s?'Sin fuente':s.file+(s.page?' · pág. '+s.page:'')+(s.sheet?' · '+s.sheet:'')+(s.cell?' · '+s.cell:s.row?' · fila '+s.row:'')+(s.method==='ocr'?' · OCR':s.method==='mixed'?' · texto y OCR':'')+(s.method==='manual_review'?' · revisión manual':'');
 const sourcesOf=fields=>fields.flatMap(f=>f?.sources||[f?.source]).filter(Boolean);
 const numericKeys=new Set(['quantity','price','amount','unitNet','net','unitGross','gross','packageGross']);
 const names={quantity:'cantidad',price:'precio',amount:'importe',net:'peso neto',gross:'peso bruto',description:'descripción',origin:'origen',hs:'HS',packageId:'identificador de bulto',cee:'marca CEE',exporter:'exportador',consignee:'destinatario',taxId:'identificación fiscal',address:'dirección',invoice:'factura',date:'fecha',expediente:'expediente',container:'contenedor',seal:'precinto'};
 const headers={
  reference:['REFERENCIA','ARTICULO','CODIGO ARTICULO','CODIGO','REF','ARTICLE','ITEM CODE','ITEM','REFERENCE'],description:['DESCRIPCION','NOMBRE','DESCRIPTION','DENOMINACION'],
  quantity:['CANTIDAD','CANT','QTY','QUANTITY','UNIDADES','UDS'],price:['PRECIO','PRECIO UNITARIO','UNIT PRICE','PRICE'],amount:['TOTAL','IMPORTE','AMOUNT','LINE TOTAL'],
  origin:['ORIGEN','PAIS','PAIS ORIGEN','COUNTRY','ORIGIN'],hs:['HS CODE','HS','ARANCELARIO','PARTIDA ARANCELARIA','COMMODITY CODE','TARIFF CODE'],
  unitNet:['KG NETO UND','KG NETO UD','PESONETO','NET UNIT','UNIT NET WEIGHT'],net:['KG NETO TOTAL','PESO NETO TOTAL','NET WEIGHT','NET TOTAL','NETO TOTAL'],
  unitGross:['KG BRUTO UND','KG BRUTO UD','UNIT GROSS WEIGHT'],gross:['KG BRUTO TOTAL','PESOBRUTO LINEA','PESO BRUTO LINEA','GROSS LINE','GROSS WEIGHT','GROSS TOTAL'],
  packageGross:['PESOBRUTO','PESO BRUTO BULTO','PACKAGE GROSS'],packageId:['CAJAS ID','BULTO ID','PACKAGE ID','PALLET ID','BULTO','CAJA'],
  packageType:['TIPO CONTENEDOR','TIPO EMBALAJE','PACKAGING','PACKAGE TYPE'],cee:['EU','CEE','UE','EU ORIGIN'],unit:['UNIDAD','UNIT','UDM']
 };
 function number(input,locale='auto'){
  if(typeof input==='number')return Number.isFinite(input)?{value:input,precision:Math.min(9,(String(input).split('.')[1]||'').length)}:{error:'Número no válido'};
  let s=String(input??'').trim().replace(/\u00a0/g,' ').replace(/\s+/g,'').replace(/(?:KG|KGS|EUR|DOP|USD|€|\$|%|UND|UDS)$/i,'');
  if(!s)return {value:null};let negative=false;if(/^\(.*\)$/.test(s)){negative=true;s=s.slice(1,-1);}if(!/^[+-]?[\d.,]+$/.test(s))return {error:'Número ilegible: '+input};
  let decimal=null;const comma=s.lastIndexOf(','),dot=s.lastIndexOf('.');
  if(comma>=0&&dot>=0)decimal=comma>dot?',':'.';
  else if(comma>=0){if(locale==='en'&&/^[-+]?\d{1,3}(,\d{3})+$/.test(s))decimal=null;else if(locale==='auto'&&/^[-+]?\d{1,3},\d{3}$/.test(s))return {error:'Separador ambiguo: '+input};else decimal=',';}
  else if(dot>=0){if(locale==='es'&&/^[-+]?\d{1,3}(\.\d{3})+$/.test(s))decimal=null;else if(locale==='auto'&&/^[-+]?\d{1,3}\.\d{3}$/.test(s))return {error:'Separador ambiguo: '+input};else decimal='.';}
  const precision=decimal?(s.length-s.lastIndexOf(decimal)-1):0;
  if(decimal){const split=s.lastIndexOf(decimal);s=s.slice(0,split).replace(/[.,]/g,'')+'.'+s.slice(split+1);}else s=s.replace(/[.,]/g,'');
  const n=Number(s)*(negative?-1:1);return Number.isFinite(n)?{value:n,precision}:{error:'Número no válido'};
 }
 function identifier(input,type='reference'){
  if(typeof input==='number'&&(!Number.isSafeInteger(input)||input<0))return {error:'Identificador numérico no seguro; revisa el texto original.'};
  const s=String(input??'').trim();if(!s)return {value:null};
  if(type==='hs'){const v=s.replace(/[.\s-]/g,'');return /^\d{4,10}$/.test(v)?{value:v}:{error:'HS ilegible o de longitud no reconocida: '+s};}
  if(type==='packageId')return {value:s.replace(/[.\s]/g,'')};
  return {value:s};
 }
 function detect(raw,name=''){
  const rows=allRows(raw),sample=norm(rows.slice(0,150).map(r=>r.text||r.cells.map(text).join(' ')).join('\n'));
  let kind='unknown',review=false,reason='No se reconoce un encabezado documental.';
  if(/(?:MOVEMENT CERTIFICATE|CERTIFICADO DE CIRCULACION|CERTIFICADO EUR|EUR 1 N[Oº°]|MERCANCIAS.*EUR 1)/.test(sample)&&/(EXPORTER|EXPORTADOR|GROSS|BRUTA|BULTOS|CONSIGNEE|DESTINATARIO)/.test(sample)){kind='certificate';reason='Título y campos del certificado.';}
  else if(/CAJAS ID|PACKING LIST|LISTA DE EMPAQUE/.test(sample)||(/ARANCELARIO/.test(sample)&&/PESOBRUTO LINEA/.test(sample))){kind='packing';reason='Identificadores de bultos y columnas de embalaje.';}
  else if(/DATOS EUR|MERCANCIAS PROPUESTAS.*EUR/.test(sample)){kind='proposal';reason='Encabezado de datos EUR.1.';}
  else if(/FACTURA|COMMERCIAL INVOICE|INVOICE NUMBER/.test(sample)){kind='invoice';reason='Encabezado de factura.';}
  else if(/HS CODE PAIS|GRUPOS HS|RESUMEN HS/.test(sample)&&!/REFERENCIA|ARTICULO/.test(sample)){kind='hs';reason='Resumen por partida, sin referencias de mercancía.';}
  else if(/REFERENCIA|ITEM CODE|ARTICULO/.test(sample)&&/CANTIDAD|QUANTITY|QTY/.test(sample)){
   kind=/DATOS.*EUR/i.test(name)?'proposal':/FACTURA|INVOICE/i.test(name)?'invoice':'unknown';review=true;reason='La tabla de artículos es legible, pero factura y datos EUR.1 tienen el mismo formato. Confirma el tipo; el nombre del archivo solo aporta una pista.';
  }else if(sample.length>50){kind='complement';review=true;reason='Documento legible sin estructura reconocida. Confirma si es complementario.';}
  return {kind,review,reason};
 }
 function allRows(raw){
  if(raw.format==='xlsx')return raw.sheets.flatMap(sheet=>sheet.rows.map(row=>({...row,sheet:sheet.name,text:row.cells.map(text).join(' ')})));
  return (raw.pages||[]).flatMap(p=>p.rows.map((r,i)=>({...r,page:p.page,row:i+1,method:p.method,confidence:p.confidence,text:r.text||r.cells.map(text).join(' ')})));
 }
 function rowSource(doc,row,cell){return {docId:doc.id,file:doc.name,...(row.sheet?{sheet:row.sheet,row:row.row,cell:cell?.cell||undefined}:{page:row.page,row:row.row,method:row.method,confidence:row.confidence}),...(cell?.formulaCached?{cachedFormula:true}:{}),raw:cell?text(cell):row.text};}
 function headerMap(cells){const map={};for(const [key,aliases] of Object.entries(headers)){const index=cells.findIndex(c=>aliases.includes(norm(text(c))));if(index>=0)map[key]=index;}return map;}
 function pdfHeaderCells(cells){const fused=[];for(let i=0;i<cells.length;i++){let match=null;for(let len=Math.min(4,cells.length-i);len>0;len--){const joined=cells.slice(i,i+len).map(text).join(' ');if(Object.values(headers).some(a=>a.includes(norm(joined)))){match={...cells[i],v:joined,len};break;}}if(match){fused.push(match);i+=match.len-1;}else fused.push({...cells[i],v:text(cells[i])});}return fused;}
 function parse(raw,info,forcedKind){
  const detected=detect(raw,info.name),doc={id:info.id,name:info.name,format:raw.format,kind:forcedKind||detected.kind,detected,lines:[],totals:{},meta:{},certificate:{hs:[]},issues:[],raw};
  const rows=allRows(raw);const combined=rows.map(r=>r.text).join('\n');doc.locale=/REFERENCIA|CANTIDAD|EXPORTADOR|BULTOS|PESO|DESCRIPCION/i.test(norm(combined))?'es':/QUANTITY|EXPORTER|GROSS WEIGHT|DESCRIPTION/i.test(combined)?'en':'auto';
  const issue=(message,source)=>doc.issues.push({message,source});
  for(const p of raw.pages||[])if(p.error||p.reviewRequired||p.method==='blank')issue(p.error||'Página '+p.page+': lectura OCR de confianza insuficiente o página sin contenido verificable.',{docId:doc.id,file:doc.name,page:p.page,method:p.method,confidence:p.confidence});
  for(const r of rows)for(const c of r.cells)if(c.problem)issue(c.problem,rowSource(doc,r,c));
  const invoiceMatches=[...combined.matchAll(/\b(?:ESF\d|INV|FACTURA|INVOICE)\s*[-:#]?\s*([A-Z0-9][A-Z0-9/-]{3,30})/gi)].filter(m=>/\d/.test(m[1]));
  if(invoiceMatches.length){const m=invoiceMatches.find(m=>/^ESF/i.test(m[0]))||invoiceMatches[0],r=rows.find(r=>r.text.includes(m[0]));doc.meta.invoice=field(/^ESF/i.test(m[0])?m[0].replace(/\s/g,''):m[1],rowSource(doc,r||rows[0]));}
  const hint=info.name.match(/\b(?:ESF\d|INV)[- ]?[\d-]{4,}/i)?.[0];doc.family=value(doc.meta.invoice)||hint?.replace(/\s/g,'')||null;
  const metaPatterns={exporter:/(?:EXPORTADOR|EXPORTER)\s*[:：]\s*(.+)/i,consignee:/(?:DESTINATARIO|CONSIGNEE|CLIENTE)\s*[:：]\s*(.+)/i,taxId:/(?:NIF|CIF|VAT|RNC|TAX ID)\s*[:：]?\s*([A-Z0-9 -]{5,25})/i,address:/(?:DIRECCION|ADDRESS)\s*[:：]\s*(.+)/i,date:/(?:FECHA|DATE)\s*[:：]\s*(\d{1,4}[/.\-]\d{1,2}[/.\-]\d{1,4})/i,expediente:/\b(H\d{6,12})\b/i,container:/\b([A-Z]{4}\s*\d{7})\b/,seal:/(?:PRECINTO|SEAL)\s*[:：#]?\s*([A-Z0-9-]{3,30})/i};
  for(const [key,re] of Object.entries(metaPatterns))for(const r of rows){const m=r.text.match(re);if(m){doc.meta[key]=field(m[1],rowSource(doc,r));break;}}
  if(doc.kind==='certificate'){parseCertificate(doc,rows);return doc;}
  if(doc.kind==='hs'){
   doc.hsRows=[];let columns=null,lastCode=null;
   for(const r of rows){const labels=r.cells.map(c=>norm(text(c)));if(labels.some(t=>t.includes('HS CODE'))){columns={code:labels.findIndex(t=>t.includes('HS CODE')),net:labels.findIndex(t=>t.includes('NETO TOTAL')||t==='NET WEIGHT'),gross:labels.findIndex(t=>t.includes('BRUTO TOTAL')||t==='GROSS WEIGHT'),amount:labels.findIndex(t=>t==='SUMA DE TOTAL'||t==='TOTAL'||t==='AMOUNT')};continue;}if(!columns)continue;
    const c=r.cells[columns.code],rawCode=text(c);if(!/^\d{4,10}$/.test(rawCode)&&!(/^[A-Z]{2}$/.test(rawCode)&&lastCode))continue;
    const origin=/^[A-Z]{2}$/.test(rawCode);if(!origin)lastCode=rawCode;const group={code:field(lastCode,rowSource(doc,r,c)),origin:origin?field(rawCode,rowSource(doc,r,c)):null,source:rowSource(doc,r)};
    for(const key of ['net','gross','amount'])if(columns[key]>=0){const cell=r.cells[columns[key]],n=number(cell?.v??text(cell),doc.locale);if(n.error)issue(n.error,rowSource(doc,r,cell));else group[key]=field(n.value,rowSource(doc,r,cell),{precision:n.precision});}doc.hsRows.push(group);
   }
   if(!doc.hsRows.length)issue('No se reconocieron las columnas del resumen HS. Su conciliación queda pendiente.',{docId:doc.id,file:doc.name});return doc;
  }
  if(!['invoice','proposal','packing'].includes(doc.kind))return doc;
  if(doc.kind==='packing')for(let i=0;i<Math.min(8,rows.length-1);i++){
   const header=rows[i],next=rows[i+1];for(const [label,key] of [['BULTOS','packageCount'],['PESO','physicalGross']]){
    const index=header.cells.findIndex(c=>norm(text(c))===label);if(index<0||!header.cells.some(c=>norm(text(c))==='VOLUMEN'))continue;const cell=raw.format==='xlsx'?next.cells[index]:next.cells.reduce((best,c)=>!best||Math.abs(c.x-header.cells[index].x)<Math.abs(best.x-header.cells[index].x)?c:best,null),n=number(cell?.v??text(cell),doc.locale);if(n.value!==null&&n.value!==undefined)doc.totals[key]=field(n.value,rowSource(doc,next,cell),{precision:n.precision});
   }
  }
  let map=null,positions=null,currentSheet=null;
  for(const r of rows){
   if(r.sheet!==currentSheet){map=null;positions=null;currentSheet=r.sheet;}
   const headerCells=raw.format==='pdf'?pdfHeaderCells(r.cells):r.cells;let candidate=headerMap(headerCells);
   if(candidate.reference!==undefined&&candidate.quantity!==undefined){map=candidate;positions=raw.format==='pdf'?{...Object.fromEntries(Object.entries(map).map(([k,i])=>[k,headerCells[i].x])),...Object.fromEntries(headerCells.map((c,i)=>['_column'+i,c.x]))}:null;continue;}
   // Native/OCR PDF headers may contain multiple words in separate text runs.
   if(raw.format==='pdf'&&!map){const fused=[];for(let i=0;i<r.cells.length;i++){let match=null;for(let len=Math.min(4,r.cells.length-i);len>0;len--){const joined=r.cells.slice(i,i+len).map(text).join(' ');if(Object.values(headers).some(a=>a.includes(norm(joined)))){match={...r.cells[i],v:joined,len};break;}}if(match){fused.push(match);i+=match.len-1;}}
    candidate=headerMap(fused);if(candidate.reference!==undefined&&candidate.quantity!==undefined){map=candidate;positions=Object.fromEntries(Object.entries(map).map(([k,i])=>[k,fused[i].x]));continue;}}
   if(!map)continue;
   const get=key=>{
    if(map[key]===undefined)return null;if(!positions)return r.cells[map[key]];
    const start=positions[key],all=Object.values(positions).filter(x=>x>start+1),end=all.length?Math.min(...all)-1:Infinity;
    const cells=r.cells.filter(c=>{
     const overlaps=c.x<start&&c.x+(c.width||0)>start,inside=c.x>=start-2&&c.x<end;
     if(!overlaps&&!inside)return false;
     // PDF text runs can span centered column headings. A run covering the
     // next heading belongs there, except combined reference/description or
     // quantity/country runs, which are split explicitly below.
     if(inside&&key!=='reference'&&key!=='quantity'&&c.x+(c.width||0)>end+1&&c.x>start)return false;
     return true;
    });let v=cells.map(text).join(' ');
    if(key==='reference'&&map.description!==undefined){const m=v.match(/^(\S+)\s+(.+)/);if(m)v=m[1];}
    if(key==='description'){const ref=get('reference');if(ref&&v.startsWith(text(ref)+' '))v=v.slice(text(ref).length).trim();}
    if(key==='quantity'&&map.origin!==undefined){const m=v.match(/^([\d.,\s+-]+)\s*([A-Z]{2})$/);if(m)v=m[1].trim();}
    if(key==='origin'&&!/^[A-Z]{2}$/.test(v)){const quantityStart=positions.quantity,nextOrigin=r.cells.find(c=>c.x>=quantityStart-2&&c.x<start+1&&/^[\d.,\s+-]+\s*[A-Z]{2}$/.test(text(c)));const m=nextOrigin&&text(nextOrigin).match(/([A-Z]{2})$/);v=m?.[1]||r.cells.map(text).find(t=>/^[A-Z]{2}$/.test(t))||'';}
    if(key==='cee'&&!/^(?:NO\s+)?(?:CEE|UE|EU)$/.test(norm(v)))v=r.cells.map(text).find(t=>/^(?:NO\s+)?(?:CEE|UE|EU)$/.test(norm(t)))||'';
    return v?{v,x:start}:null;
   };
   const refCell=get('reference'),refText=text(refCell),label=norm(r.text),quantityCell=get('quantity');
   const packageSubtotal=/^TOTAL\b/.test(norm(text(get('packageId'))));
   if(packageSubtotal)continue;
   if(/^(?:TOTAL|SUBTOTAL|SUMA|GRAND TOTAL|TOTAL GENERAL)\b/.test(norm(refText))||(!refText&&/\bTOTAL\b/.test(label))||(!refText&&!text(quantityCell)&&doc.kind!=='packing'&&doc.lines.length&&['amount','net','gross'].some(k=>text(get(k))))){
    // Subtotals are evidence, never merchandise or extra packages.
    for(const key of ['amount','net','gross','packageGross']){const c=get(key);const n=number(c?.v??text(c),doc.locale);if(n.value!==null&&n.value!==undefined)doc.totals[key]=field(n.value,rowSource(doc,r,c),{precision:n.precision});}continue;
   }
   if(!refText)continue;
   if(/^TOTAL\b/.test(label)||/^REFERENCIA|^ARTICULO/.test(norm(refText)))continue;
   const q=number(quantityCell?.v??text(quantityCell),doc.locale);
   if(q.value===null||q.value===undefined){if(/\d/.test(refText))issue('No se pudo leer la cantidad de la referencia '+refText,rowSource(doc,r,quantityCell));continue;}
   const id=identifier(refCell?.v??refText);if(id.error){issue(id.error,rowSource(doc,r,refCell));continue;}
   const line={id:doc.id+':'+(r.sheet||r.page)+':'+r.row,reference:field(id.value,rowSource(doc,r,refCell)),source:rowSource(doc,r),documentId:doc.id};
   for(const key of Object.keys(map)){
    if(key==='reference')continue;const c=get(key);if(!c||text(c)==='')continue;const src=rowSource(doc,r,c),rawValue=c.v??text(c);let parsed;
    if(numericKeys.has(key)){parsed=number(rawValue,doc.locale);if(parsed.error){issue(key+': '+parsed.error,src);continue;}line[key]=field(parsed.value,src,{precision:parsed.precision});}
    else if(key==='hs'||key==='packageId'){parsed=identifier(rawValue,key);if(parsed.error)issue(parsed.error,src);else line[key]=field(parsed.value,src);}
    else line[key]=field(text(c),src);
    if(line[key]&&c.formulaCached)line[key].kind='cached_formula';
   }
   line.quantity=field(q.value,rowSource(doc,r,quantityCell),{precision:q.precision});
   if(!line.amount&&line.price)line.amount=calculated(value(line.quantity)*value(line.price),sourcesOf([line.quantity,line.price]));
   if(!line.net&&line.unitNet)line.net=calculated(value(line.quantity)*value(line.unitNet),sourcesOf([line.quantity,line.unitNet]));
   if(!line.gross&&line.unitGross)line.gross=calculated(value(line.quantity)*value(line.unitGross),sourcesOf([line.quantity,line.unitGross]));
   doc.lines.push(line);
  }
  if(!doc.lines.length)issue('No se reconoció una tabla de artículos con referencia y cantidad. Revisa el tipo o registra las líneas mediante revisión documental.',{docId:doc.id,file:doc.name});
  return doc;
 }
 function parseCertificate(doc,rows){
  for(let i=0;i<rows.length;i++){
   const r=rows[i],t=r.text,n=norm(t),source=rowSource(doc,r);
   if(/BORRADOR|DRAFT|PROVISIONAL/.test(n))doc.certificate.draft=field(true,source);
   if(/^\s*\d{4}(?:[\s,;]+\d{4})+\s*$/.test(t)){for(const code of t.match(/\d{4}/g))doc.certificate.hs.push({code:field(code,source),description:null});continue;}
   let hs=t.match(/(?:\bHS(?:\s*CODE)?|\bPARTIDA(?:\s+ARANCELARIA)?|\bCODIGO\s+ARANCELARIO)\s*[:.-]?\s*(\d[\d. ]{2,12}\d)(.*)/i)||t.match(/^\s*(?:\d{1,2}[.)]\s*)?(\d{4}(?:[ .]\d{2}){0,3}|\d{6,10})\b\s+(.+)/);
   if(hs){const parsed=identifier(hs[1],'hs');if(parsed.value){let desc=hs[2]?.trim()||'';const next=rows[i+1];if(next&&!/^\s*\d{4}\b/.test(next.text)&&!/(?:BULTOS|PESO|BRUTO|GROSS|FACTURA)/i.test(next.text)&&desc.length<20)desc+=' '+next.text;doc.certificate.hs.push({code:field(parsed.value,source),description:field(desc.trim(),source)});}}
   const packs=t.match(/\b(\d+)\s*(?:BULTOS|PALET[S]?|PALLET[S]?|PACKAGES|CAJAS|PACKING CASES)\b/i)||t.match(/(?:BULTOS|PACKAGES|NUMBER OF PACKAGES)\s*[:：]\s*(\d+)/i);
   if(packs)doc.certificate.packages=field(Number(packs[1]),source,{precision:0});
   if(/(?:PESO BRUTO|MASA BRUTA|GROSS (?:WEIGHT|MASS))/.test(n)){
    const candidates=[r,...rows.slice(i+1,i+4)];for(const next of candidates){let raw=next.text.replace(/(?:^\s*9[.)]\s*|PESO BRUTO|MASA BRUTA|GROSS WEIGHT|GROSS MASS|KG|KGS|O CUALQUIER OTRA MEDIDA)/gi,' ');const nums=raw.match(/\b\d[\d.,]*(?:\s+[.,]\s*\d+)?\b/g)||[];if(nums.length===1){const parsed=number(nums[0],doc.locale);if(parsed.value!==null&&parsed.value!==undefined){doc.certificate.gross=field(parsed.value,rowSource(doc,next),{precision:parsed.precision});break;}}}
   }
  }
  if(!doc.certificate.gross){for(const r of rows){const m=r.text.match(/\b(\d[\d.,]*)\s*KG\b/i);if(m){const n=number(m[1],doc.locale);if(n.value!==null&&n.value!==undefined){doc.certificate.gross=field(n.value,rowSource(doc,r),{precision:n.precision});break;}}}}
  if(!doc.certificate.hs.length)doc.issues.push({message:'No se pudieron identificar las partidas HS del certificado.',source:{docId:doc.id,file:doc.name}});
 }
 function group(lines){
  const result=new Map();for(const line of lines){const ref=value(line.reference);if(!ref)continue;if(!result.has(ref))result.set(ref,{reference:ref,lines:[],description:value(line.description)||'',quantity:0,amount:0,net:0,gross:0,known:{quantity:true,amount:true,net:true,gross:true},origins:new Set(),hs:new Set()});const g=result.get(ref);g.lines.push(line);
   for(const k of ['quantity','amount','net','gross']){if(value(line[k])===null)g.known[k]=false;else g[k]+=value(line[k]);}if(value(line.origin))g.origins.add(norm(value(line.origin)));if(value(line.hs))g.hs.add(value(line.hs));
  }return result;
 }
 const EU=new Set(['AT','BE','BG','HR','CY','CZ','DK','EE','FI','FR','DE','GR','EL','HU','IE','IT','LV','LT','LU','MT','NL','PL','PT','RO','SK','SI','ES','SE','UE','EU','CEE']);
 function euOrigin(origin){return EU.has(norm(origin));}
 function packageReport(lines,included){
  const map=new Map(),unknown=[];
  for(const line of lines){const id=value(line.packageId);if(!id){unknown.push(line);continue;}if(!map.has(id))map.set(id,{id,lines:[],included:[],excluded:[],gross:null,grossSources:[],grossConflict:false,cee:new Set()});const p=map.get(id);p.lines.push(line);(included.has(value(line.reference))?p.included:p.excluded).push(value(line.reference));const g=value(line.packageGross);if(g!==null){if(p.gross!==null&&Math.abs(p.gross-g)>.001)p.grossConflict=true;p.gross=g;p.grossSources.push(line.packageGross.source);}if(value(line.cee))p.cee.add(norm(value(line.cee)));}
  const packages=[...map.values()].map(p=>({...p,cee:[...p.cee],included:[...new Set(p.included)],excluded:[...new Set(p.excluded)],category:p.included.length?(p.excluded.length?'mixed':'included'):'excluded'}));
  const whole=packages.filter(p=>p.category==='included'),mixed=packages.filter(p=>p.category==='mixed'),any=packages.filter(p=>p.included.length),whollyCEE=whole.filter(p=>p.lines.every(l=>['CEE','UE','EU','SI','YES','TRUE','1'].includes(norm(value(l.cee)))));
  return {packages,total:packages.length,whole:whole.length,whollyCEE:whollyCEE.length,mixed:mixed.length,any:any.length,unknown:unknown.map(l=>l.source),physicalGross:packages.length&&packages.every(p=>p.gross!==null&&!p.grossConflict)?packages.reduce((s,p)=>s+p.gross,0):null};
 }
 function analyze(documents,options={}){
  const tolerance={quantity:1e-6,money:.01,weight:.001,...options.tolerance};
  const checks=[],findings=[],doubts=[],reconciliation=[],representationChecks=[];let serial=0;
  const check=(name,status,declared,computed,extra={})=>{const row={id:++serial,name,status,declared,computed,difference:typeof declared==='number'&&typeof computed==='number'?computed-declared:null,...extra};checks.push(row);if(status==='difference'||status==='review')findings.push(row);if(status==='incomplete')doubts.push(row);return row;};
  const docs=documents.map(record=>record.extraction?{...record.extraction,classificationConfirmed:record.classification_confirmed,status:record.status,error:record.error}:record);
  for(const doc of docs){
   if(doc.kind==='archive'&&doc.status==='ready')continue;
   if(doc.status==='error'||doc.status==='pending'||doc.status==='stored'||doc.error){const affected={certificate:'identificación de las partes, cobertura HS, bultos y peso del certificado',proposal:'mercancías propuestas, referencias, cantidades, importes, HS y pesos',invoice:'conciliación de mercancías, cantidades, importes, origen, pesos e identificación de las partes',packing:'referencias, contenido de bultos, cantidades y pesos',hs:'resumen HS y sus importes y pesos frente a factura',archive:'integridad y completitud de los documentos del embarque',unknown:'clasificación y alcance de la comparación documental',complement:'datos complementarios del embarque'};check('Lectura: '+doc.name,'incomplete',doc.error||'Documento pendiente',null,{sources:[{docId:doc.id,file:doc.name}],affected:[doc.name],explanation:'Quedan afectadas las comprobaciones de '+(affected[doc.kind]||affected.unknown)+'. El documento debe leerse antes de dar por completa la verificación.'});continue;}
   if(doc.detected?.review&&!doc.classificationConfirmed)check('Clasificación: '+doc.name,'incomplete',kinds[doc.kind],'Tipo pendiente de confirmar',{sources:[{docId:doc.id,file:doc.name}]});
   for(const issue of doc.issues||[])if(!issue.resolved)check('Extracción: '+doc.name,'incomplete',issue.message,null,{sources:[issue.source].filter(Boolean),affected:[doc.name]});
  }
  const ready=docs.filter(d=>d.lines||d.certificate);
  const byKind={};for(const kind of ['certificate','proposal','invoice','packing']){byKind[kind]=ready.filter(d=>d.kind===kind);if(!byKind[kind].length)check('Documento requerido: '+kinds[kind],'incomplete','No disponible',null,{affected:[kinds[kind]]});}
  if(byKind.certificate.length>1)check('Certificado a verificar','incomplete',byKind.certificate.length,'Selecciona un único certificado; clasifica los demás como complementarios');
  function primary(kind){
   if((byKind[kind]||[]).some(d=>!d.family)&&(byKind[kind]||[]).length>1&&new Set((byKind[kind]||[]).map(d=>d.format)).size>1)check('Agrupación de representaciones de '+kinds[kind],'incomplete','Hay documentos sin número de factura o familia confirmado',null,{sources:(byKind[kind]||[]).filter(d=>!d.family).map(d=>({docId:d.id,file:d.name})),explanation:'Revisa el número de factura de cada documento para saber si PDF y Excel representan la misma fuente.'});
   const families=new Map();for(const d of byKind[kind]||[]){const signature=d.family||d.name.replace(/\.(pdf|xlsx)$/i,'').replace(/FACTURA|DATOS\s*EUR.?1|PACKING\s*LIST/gi,'').trim();if(!families.has(signature))families.set(signature,[]);families.get(signature).push(d);}
   const selected=[];for(const family of families.values()){
    family.sort((a,b)=>(b.format==='xlsx')-(a.format==='xlsx')||(b.lines?.length||0)-(a.lines?.length||0));const chosen=family[0];selected.push(chosen);
    for(const other of family.slice(1)){
     const a=group(chosen.lines||[]),b=group(other.lines||[]),refs=new Set([...a.keys(),...b.keys()]);let differences=0,incomplete=!(a.size&&b.size);
     for(const ref of refs){const x=a.get(ref),y=b.get(ref);if(!x||!y){differences++;check('Representación · '+ref,'difference',x?'Presente':'Ausente',y?'Presente':'Ausente',{affected:[ref],sources:[...(x?.lines||[]),...(y?.lines||[])].map(l=>l.source),explanation:chosen.name+' frente a '+other.name});continue;}
      for(const k of ['quantity','amount','net','gross']){if(!x.known[k]||!y.known[k]){if(['quantity','amount'].includes(k)&&kind!=='packing')incomplete=true;continue;}const tol=k==='quantity'?tolerance.quantity:k==='amount'?tolerance.money:tolerance.weight,diff=Math.abs(x[k]-y[k]);if(diff>tol+1e-9){const rounding=[...x.lines,...y.lines].reduce((s,l)=>s+(l[k]?.source?.page&&l[k].precision!==undefined ? 0.5*10**(-l[k].precision) : 0),0);const rounded=diff<=rounding+tol+1e-9;if(!rounded)differences++;check('Representación · '+ref+' · '+names[k],rounded?'rounding':'difference',x[k],y[k],{affected:[ref],sources:sourcesOf([...x.lines,...y.lines].map(l=>l[k])),explanation:chosen.name+' frente a '+other.name+'. Redondeo admisible según los decimales de las líneas PDF: '+rounding+'.'});}}
      for(const [key,setName] of [['origin','origins'],['hs','hs']])if(x[setName].size&&y[setName].size&&[...x[setName]].sort().join()!=[...y[setName]].sort().join()){differences++;check('Representación · '+ref+' · '+names[key],'difference',[...x[setName]].join(', '),[...y[setName]].join(', '),{affected:[ref],sources:sourcesOf([...x.lines,...y.lines].map(l=>l[key]))});}
     }
     const row=check('Representaciones de '+kinds[kind],incomplete?'incomplete':differences?'difference':'pass',chosen.name,other.name,{difference:differences,sources:[{docId:chosen.id,file:chosen.name},{docId:other.id,file:other.name}],explanation:'Se usa una sola representación para sumar mercancía. '+(differences?'Hay diferencias documentales entre ambas.':'Las representaciones no se suman entre sí.')});representationChecks.push(row);
    }
   }return selected;
  }
  const proposals=primary('proposal'),invoices=primary('invoice'),packings=primary('packing'),certificate=byKind.certificate[0],proposedLines=proposals.flatMap(d=>d.lines||[]),invoiceLines=invoices.flatMap(d=>d.lines||[]),packingLines=packings.flatMap(d=>d.lines||[]),proposed=group(proposedLines),invoice=group(invoiceLines),packing=group(packingLines);
  for(const [label,lines,required] of [['Datos EUR.1',proposedLines,['description','quantity','amount','origin','hs','net','gross']],['Factura',invoiceLines,['description','quantity','amount','origin','hs','net','gross']],['Packing',packingLines,['description','quantity','origin','hs','net','packageId','cee']]])for(const key of required){const missing=lines.filter(l=>value(l[key])===null);if(missing.length)check(label+' · campos ausentes: '+names[key],'incomplete',missing.length+' líneas sin valor',null,{affected:missing.map(l=>value(l.reference)),sources:missing.map(l=>l.source)});}
  for(const key of ['exporter','consignee','taxId','address','invoice','date','expediente','container','seal']){
   const fields=[certificate,...byKind.invoice,...byKind.proposal,...byKind.packing].filter(Boolean).map(d=>({doc:d,field:d.meta?.[key]})).filter(x=>value(x.field)!==null);
   const values=[...new Set(fields.map(x=>norm(value(x.field))))];
   const missingIdentity=['exporter','consignee'].includes(key)&&(value(certificate?.meta?.[key])===null||!byKind.invoice.some(d=>value(d.meta?.[key])!==null));
   check('Identificación documental · '+names[key],missingIdentity?'incomplete':fields.length<2?'unavailable':values.length>1?'review':'pass',fields.map(x=>x.doc.name+': '+value(x.field)).join(' | ')||'No consta',values.join(' | ')||null,{sources:sourcesOf(fields.map(x=>x.field)),explanation:missingIdentity?'Falta identificar esta parte en el certificado o en la factura. Coteja el original y completa la revisión.':fields.length<2?'Información no contrastable entre dos fuentes; no se presenta como coincidencia.':'Las diferencias de escritura requieren revisión documental.'});
  }
  const comparedKeys=['quantity','amount','net','gross'];
  function compareGroup(ref,left,right,label,keys){
   if(!right){check(label+': '+ref,'difference',left.quantity,'Referencia ausente',{affected:[ref],sources:left.lines.map(l=>l.source),explanation:'La ausencia se determina por referencia, sin depender de la longitud del HS.'});reconciliation.push({reference:ref,description:left.description,proposal:Object.fromEntries(comparedKeys.map(k=>[k,left.known[k]?left[k]:null])),[label]:null,sources:left.lines.map(l=>l.source)});return;}
   const result={reference:ref,description:left.description,proposal:{},[label]:{},sources:[...left.lines,...right.lines].map(l=>l.source)};
   for(const k of keys){const x=left.known[k]?left[k]:null,y=right.known[k]?right[k]:null;result.proposal[k]=x;result[label][k]=y;const tol=k==='quantity'?tolerance.quantity:k==='amount'?tolerance.money:tolerance.weight;
    if(x===null||y===null){check(label+' · '+ref+' · '+names[k],'incomplete',x,y,{affected:[ref],sources:sourcesOf([...left.lines,...right.lines].map(l=>l[k])),explanation:'Falta un valor legible. No se sustituye por cero.'});}
    else if(Math.abs(x-y)>tol+1e-9){let explanation='Tolerancia: '+tol+(k==='amount'?' de la moneda del documento.':k==='net'||k==='gross'?' kg.':' unidades.');if(k==='net'&&left.quantity===right.quantity&&left.quantity){explanation+=' Misma cantidad: '+left.quantity+'. Diferencia de peso unitario calculada: '+((y-x)/left.quantity)+' kg/unidad.';const a=[...new Set(left.lines.map(l=>value(l.unitNet)).filter(v=>v!==null))],b=[...new Set(right.lines.map(l=>value(l.unitNet)).filter(v=>v!==null))];if(a.length===1&&b.length===1)explanation+=' Peso unitario declarado: '+a[0]+' frente a '+b[0]+' kg/unidad.';}check(label+' · '+ref+' · '+names[k],'difference',x,y,{affected:[ref],sources:sourcesOf([...left.lines,...right.lines].flatMap(l=>k==='net'?[l[k],l.unitNet,l.quantity]:[l[k]])),explanation});}
   }
   if(left.origins.size&&right.origins.size&&[...left.origins].sort().join()!=[...right.origins].sort().join())check(label+' · origen · '+ref,'difference',[...left.origins].join(', '),[...right.origins].join(', '),{affected:[ref],sources:sourcesOf([...left.lines,...right.lines].map(l=>l.origin))});
   if(left.hs.size&&right.hs.size){const a=[...left.hs].sort(),b=[...right.hs].sort();if(a.join()!=b.join()){const same4=[...new Set(a.map(c=>c.slice(0,4)))].sort().join()==[...new Set(b.map(c=>c.slice(0,4)))].sort().join();const prefix=same4&&a.every(c=>b.some(other=>c.startsWith(other)||other.startsWith(c)));check(label+' · HS · '+ref,same4?'review':'difference',a.join(', '),b.join(', '),{affected:[ref],sources:sourcesOf([...left.lines,...right.lines].map(l=>l.hs)),explanation:prefix?'Diferencia de longitud HS; la referencia está presente.':same4?'Diferencia de subpartida, con la misma partida HS4.':'Diferente partida HS4 para la misma referencia.'});}}
   reconciliation.push(result);
  }
  for(const [ref,p] of proposed){compareGroup(ref,p,invoice.get(ref),'Factura',comparedKeys);compareGroup(ref,p,packing.get(ref),'Packing',['quantity','net']);}
  for(const [label,other] of [['Factura',invoice],['Packing',packing]])if(proposed.size&&other.size){const entries=checks.filter(c=>c.name.startsWith(label)),status=entries.some(c=>c.status==='incomplete')?'incomplete':entries.some(c=>c.status==='difference')?'difference':entries.some(c=>c.status==='review')?'review':'pass';check('Mercancías propuestas frente a '+label.toLowerCase(),status,proposedLines.length+' líneas propuestas',proposed.size+' referencias contrastadas',{explanation:'Se agrupan referencias repetidas conservando todas las líneas originales.',sources:[...proposals,...(label==='Factura'?invoices:packings)].map(d=>({docId:d.id,file:d.name}))});}
  for(const d of [...proposals,...invoices,...packings]){const repeated=[...group(d.lines||[])].filter(([,g])=>g.lines.length>1);if(repeated.length)check('Referencias repetidas · '+d.name,'information',repeated.length+' referencias',repeated.reduce((s,[,g])=>s+g.lines.length,0)+' líneas agrupadas',{affected:repeated.map(([ref])=>ref),sources:repeated.flatMap(([,g])=>g.lines.map(l=>l.source)),explanation:'Se han sumado las cantidades por referencia. Se conservan todas las líneas para revisar posibles duplicidades, sin descartarlas automáticamente.'});}
  const euRefs=new Set(invoiceLines.filter(l=>euOrigin(value(l.origin))).map(l=>value(l.reference))),missingProposal=[...euRefs].filter(ref=>!proposed.has(ref));
  if(missingProposal.length)check('Cobertura de referencias con país UE declarado','review',[...euRefs].length,proposed.size,{affected:missingProposal,sources:invoiceLines.filter(l=>missingProposal.includes(value(l.reference))).map(l=>l.source),explanation:'Hay referencias de país UE en la factura fuera de la propuesta. Confirmar si su exclusión es intencional. El país no acredita por sí solo origen preferencial.'});
  for(const line of packingLines){if(!proposed.has(value(line.reference)))continue;const cee=norm(value(line.cee));if(cee==='NO CEE'||cee==='NO UE'||cee==='NON EU')check('Marca de origen en packing','difference','Incluido en datos EUR.1',value(line.cee),{affected:[value(line.reference),value(line.packageId)].filter(Boolean),sources:sourcesOf([line.reference,line.cee]),explanation:'Contradicción documental; requiere aclarar la evidencia de origen, sin decidir su validez legal.'});}
  for(const d of [...proposals,...invoices])for(const k of ['amount','net','gross']){
   const declared=value(d.totals?.[k]),fields=d.lines.map(l=>l[k]);if(declared===null){check('Total declarado '+names[k]+' · '+d.name,'incomplete',null,null,{explanation:'No se identificó un total declarado. El total calculado permanece disponible.',sources:[{docId:d.id,file:d.name}]});continue;}
   if(fields.some(f=>value(f)===null)){check('Total '+names[k]+' · '+d.name,'incomplete',declared,null,{sources:sourcesOf([d.totals[k],...fields])});continue;}
   const sum=fields.reduce((s,f)=>s+value(f),0),diff=sum-declared,tol=k==='amount'?tolerance.money:tolerance.weight;
   const matchesFullInvoice=d.kind==='proposal'&&k==='amount'&&invoiceLines.length&&Math.abs(invoiceLines.reduce((s,l)=>s+(value(l.amount)||0),0)-declared)<=tol;
   check('Total '+names[k]+' · '+d.name,Math.abs(diff)<=tol+1e-9?'pass':'difference',declared,sum,{sources:sourcesOf([d.totals[k],...fields]),explanation:'Suma de sus propias líneas. Tolerancia: '+tol+'.'+(matchesFullInvoice&&Math.abs(diff)>tol?' El total declarado coincide con la factura completa.':'')});
  }
  const hsProposed=[...new Set(proposedLines.map(l=>value(l.hs)?.slice(0,4)).filter(Boolean))].sort(),hsCertificate=[...new Set((certificate?.certificate?.hs||[]).map(h=>value(h.code)?.slice(0,4)).filter(Boolean))].sort(),hsMissing=hsProposed.filter(h=>!hsCertificate.includes(h)),hsExtra=hsCertificate.filter(h=>!hsProposed.includes(h));
  check('Cobertura HS a cuatro dígitos',!hsProposed.length||!hsCertificate.length||proposedLines.some(l=>!value(l.hs))?'incomplete':hsMissing.length||hsExtra.length?'difference':'pass',hsCertificate.join(', '),hsProposed.join(', '),{affected:[...hsMissing,...hsExtra],sources:[...sourcesOf(proposedLines.map(l=>l.hs)),...sourcesOf((certificate?.certificate?.hs||[]).map(h=>h.code))],explanation:'Faltantes: '+(hsMissing.join(', ')||'ninguno')+'. Sobrantes: '+(hsExtra.join(', ')||'ninguno')+'. Se conserva el HS completo por línea.'});
  for(const group of certificate?.certificate?.hs||[]){const code=value(group.code)?.slice(0,4),desc=norm(value(group.description));for(const line of proposedLines.filter(l=>value(l.hs)?.slice(0,4)===code)){const item=norm(value(line.description));const contradictions=[[/\bPVC\b/,/SILICON/,'PVC frente a silicona'],[/PORCELANA|PORCELAIN/,/STONEWARE|\bGRES\b/,'porcelana frente a stoneware/gres']];for(const [a,b,message] of contradictions)if(a.test(desc)&&b.test(item))check('Descripción del grupo '+code,'review',value(group.description),value(line.description),{affected:[value(line.reference)],sources:sourcesOf([group.description,line.description]),explanation:'Contradicción de material: '+message+'. Solicitar aclaración; no se propone una clasificación arancelaria alternativa.'});}}
  const packages=packageReport(packingLines,new Set(proposed.keys())),packDeclared=value(certificate?.certificate?.packages);
  for(const d of packings){const physical=packageReport(d.lines||[],new Set(proposed.keys()));for(const [key,computed] of [['packageCount',physical.total],['physicalGross',physical.physicalGross]])if(value(d.totals?.[key])!==null){const declared=value(d.totals[key]);check((key==='packageCount'?'Bultos declarados del embarque':'Peso físico declarado del embarque')+' · '+d.name,computed===null?'incomplete':Math.abs(declared-computed)<=(key==='packageCount'?0:tolerance.weight)?'pass':'difference',declared,computed,{sources:[...sourcesOf([d.totals[key]]),...d.lines.map(l=>l.source)],explanation:'Se recalcula una sola vez por identificador de bulto. No se compara este total con el peso del certificado parcial.'});}}
  const packStatus=!packingLines.length||packages.unknown.length||packDeclared===null?'incomplete':packages.mixed&&packDeclared===packages.whole?'review':packDeclared===packages.any?'pass':'review';
  check('Bultos cubiertos por el certificado',packStatus,packDeclared,packages.any,{affected:packages.packages.filter(p=>p.category==='mixed').map(p=>p.id),sources:[...sourcesOf([certificate?.certificate?.packages]),...packingLines.map(l=>l.source)],explanation:'Embarque: '+packages.total+'; íntegramente incluidos: '+packages.whole+' (con todas las líneas CEE: '+packages.whollyCEE+'); mixtos: '+packages.mixed+'; con alguna mercancía incluida: '+packages.any+'. '+(packStatus==='review'?'Aclarar el criterio de declaración de los bultos mixtos; no se impone automáticamente modificar la cantidad.':'Se cuentan identificadores únicos, sin subtotales ni líneas.')});
  for(const p of packages.packages)if(p.grossConflict)check('Peso físico del bulto '+p.id,'difference','Pesos repetidos para el mismo identificador','Valores distintos',{affected:[p.id],sources:p.grossSources});
  const grossFields=proposedLines.map(l=>l.gross),grossCalculated=proposedLines.length&&grossFields.every(f=>value(f)!==null)?grossFields.reduce((s,f)=>s+value(f),0):null,grossDeclared=value(certificate?.certificate?.gross),precision=certificate?.certificate?.gross?.precision??0,roundingTolerance=.5*Math.pow(10,-precision),grossDifference=grossCalculated===null||grossDeclared===null?null:grossCalculated-grossDeclared;
  check('Peso bruto de mercancías propuestas',grossDifference===null?'incomplete':Math.abs(grossDifference)<=tolerance.weight?'pass':Math.abs(grossDifference)<=roundingTolerance+tolerance.weight?'rounding':'difference',grossDeclared,grossCalculated,{sources:sourcesOf([certificate?.certificate?.gross,...grossFields]),explanation:'Solo mercancía incluida. Tolerancia técnica '+tolerance.weight+' kg; redondeo a '+precision+' decimales: ±'+roundingTolerance+' kg. Peso físico del embarque: '+(packages.physicalGross??'no disponible')+' kg, mostrado por separado y sin repetir bultos.'});
  for(const d of ready.filter(d=>d.kind==='hs')){
   const summary=new Map(),actual=new Map();for(const row of (d.hsRows||[]).filter(r=>!r.origin)){const code=value(row.code)?.slice(0,4);if(!summary.has(code))summary.set(code,[]);summary.get(code).push(row);}for(const l of invoiceLines){const code=value(l.hs)?.slice(0,4);if(!actual.has(code))actual.set(code,[]);actual.get(code).push(l);}
   for(const code of new Set([...summary.keys(),...actual.keys()]))for(const key of ['amount','net','gross']){const a=summary.get(code)||[],b=actual.get(code)||[],known=a.length&&b.length&&[...a,...b].every(r=>value(r[key])!==null),x=a.reduce((s,r)=>s+(value(r[key])||0),0),y=b.reduce((s,r)=>s+(value(r[key])||0),0);check('Resumen HS4 '+code+' · '+names[key],!known?'incomplete':Math.abs(x-y)<=(key==='amount'?tolerance.money:tolerance.weight)+1e-9?'pass':'difference',a.length?x:null,b.length?y:null,{affected:[code],sources:sourcesOf([...a,...b].map(r=>r[key])),explanation:'Resumen HS del embarque frente a factura completa; los renglones de país no se vuelven a sumar.'});}
  }
  // Required checks remain visible even when all parsed line collections are empty.
  if(!proposedLines.length)check('Lista de mercancías propuestas','incomplete',0,null);
  if(!invoiceLines.length)check('Mercancías de factura','incomplete',0,null);
  if(!packingLines.length)check('Mercancías de packing','incomplete',0,null);
  const conclusion=doubts.length?'Verificación incompleta por documentos o datos ilegibles o ausentes.':findings.length?'Con diferencias que requieren revisión.':'Sin diferencias documentales detectadas.';
  const itemLines=[...proposedLines,...invoiceLines,...packingLines];
  return {version:VERSION,at:new Date().toISOString(),conclusion,tolerance,checks,findings,doubts,certificate:certificate?.certificate||null,coverage:{proposed:[...proposed.keys()],invoiceOnly:[...invoice.keys()].filter(r=>!proposed.has(r)),packingOnly:[...packing.keys()].filter(r=>!proposed.has(r)),repeated:docs.flatMap(d=>[...group(d.lines||[])].filter(([,g])=>g.lines.length>1).map(([ref,g])=>({file:d.name,reference:ref,lineCount:g.lines.length,sources:g.lines.map(l=>l.source)})))},hs:{included:hsProposed,certificate:hsCertificate,missing:hsMissing,extra:hsExtra},packages,reconciliation,representations:representationChecks,lines:itemLines,documents:docs.map(d=>({id:d.id,name:d.name,kind:d.kind,status:d.status,issues:d.issues||[],meta:d.meta||{},totals:d.totals||{},reviews:d.reviews||[],lineCount:d.lines?.length||0})),totals:{proposedLines:proposedLines.length,proposedReferences:proposed.size,proposedAmount:proposedLines.every(l=>value(l.amount)!==null)?proposedLines.reduce((s,l)=>s+value(l.amount),0):null,proposedGross:grossCalculated,certificateGross:grossDeclared},disclaimer:'Comparación documental de las fuentes cargadas. No certifica legalmente el origen preferencial ni autoriza automáticamente el despacho.'};
 }
 return {VERSION,kinds,norm,text,value,field,calculated,sourceLabel,sourcesOf,number,identifier,detect,allRows,parse,parseCertificate,group,euOrigin,packageReport,analyze};
});

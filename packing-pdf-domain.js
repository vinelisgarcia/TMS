(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.TMSPackingPDFDomain=api;})(typeof globalThis!=='undefined'?globalThis:this,function(){
 'use strict';
 const normalize=value=>String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]/g,'');
 const invoiceId=value=>String(value??'').trim().toUpperCase().replace(/\s+/g,'');
 const escape=value=>String(value).replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
 const matches=(text,value)=>!!String(value||'').trim()&&new RegExp('(^|[^A-Z0-9])'+escape(String(value).trim())+'(?=$|[^A-Z0-9])','i').test(text);
 const convert=(value,factor)=>Number((value*factor).toPrecision(14));
 function number(raw){const s=String(raw??'').trim();if(/^\d+$/.test(s))return Number(s);if(/^\d+[.,]\d+$/.test(s)){const [whole,fraction]=s.split(/[.,]/);if(fraction.length===3&&Number(whole)!==0)return null;return Number(s.replace(',','.'));}if(/^\d{1,3}(?:,\d{3})+\.\d+$/.test(s))return Number(s.replace(/,/g,''));if(/^\d{1,3}(?:\.\d{3})+,\d+$/.test(s))return Number(s.replace(/\./g,'').replace(',','.'));return null;}
 function textRows(items){
  const rows=[];items.filter(item=>item.str?.trim()).forEach(item=>{const x=item.transform?.[4]||0,y=item.transform?.[5]||0;let row=rows.find(row=>Math.abs(row.y-y)<=2);if(!row){row={y,cells:[]};rows.push(row);}row.cells.push({text:item.str,x,width:item.width||0});});
  return rows.sort((a,b)=>b.y-a.y).map(row=>{row.cells.sort((a,b)=>a.x-b.x);let text='';row.cells.forEach((cell,index)=>{const prev=row.cells[index-1];text+=(index&&cell.x-prev.x-prev.width>1.5?' ':'')+cell.text;});return {text,cells:row.cells};});
 }
 function invoiceNumbers(text){const found=[];const pattern=/(?:\bINVOICE\s*(?:NO\.?|NUMBER|#)|\bFACTURA\s*(?:NO\.?|NUMERO|N[º°.]|#))\s*[:#-]?\s*([A-Z0-9][A-Z0-9/_.-]{1,45})/gi;for(const match of text.matchAll(pattern)){if(!['DATE','FECHA','NUMBER','NO','NUMERO'].includes(normalize(match[1])))found.push(match[1]);}return [...new Set(found)];}
 const defs={
  COUNTRYOFORIGIN:{field:'PAIS'},COO:{field:'PAIS'},ORIGIN:{field:'PAIS'},ORIGEN:{field:'PAIS'},
  HSCODE:{field:'ARANCELARIO'},HTSCODE:{field:'ARANCELARIO'},TARIFFCODE:{field:'ARANCELARIO'},PARTIDAARANCELARIA:{field:'ARANCELARIO'},
  EAN:{field:'EAN'},UPC:{field:'EAN'},BARCODE:{field:'EAN'},UNITB:{field:'UNITB'},UNITSPERBOX:{field:'UNITB'},
  QTY:{field:'CANTIDAD_FACTURA',factor:1,readOnly:true},QTYSHIPPED:{field:'CANTIDAD_FACTURA',factor:1,readOnly:true},SHIPPEDQUANTITY:{field:'CANTIDAD_FACTURA',factor:1,readOnly:true},
  UNITNETWEIGHTKG:{field:'KG NETO UD',factor:1},UNITNETWTKG:{field:'KG NETO UD',factor:1},KGNETOUD:{field:'KG NETO UD',factor:1},UNITNETWEIGHTLB:{field:'KG NETO UD',factor:.45359237},
  BOXID:{field:'CAJAS_ID'},CARTONID:{field:'CAJAS_ID'},CAJASID:{field:'CAJAS_ID'},PACKAGEID:{field:'CAJAS_ID'},
  PACKAGEGROSSWEIGHTKG:{field:'PESOBRUTO',factor:1},BOXGROSSKG:{field:'PESOBRUTO',factor:1},
  PACKAGETYPE:{field:'TIPO CONTENEDOR'},TIPOCONTENEDOR:{field:'TIPO CONTENEDOR'},
  LENGTHM:{field:'LARGO',factor:1},WIDTHM:{field:'ANCHO',factor:1},HEIGHTM:{field:'ALTO',factor:1},
  LENGTHCM:{field:'LARGO',factor:.01},WIDTHCM:{field:'ANCHO',factor:.01},HEIGHTCM:{field:'ALTO',factor:.01},
  LENGTHIN:{field:'LARGO',factor:.0254},WIDTHIN:{field:'ANCHO',factor:.0254},HEIGHTIN:{field:'ALTO',factor:.0254}
 };
 function valueFor(def,raw){
  const value=String(raw??'').trim();if(!value)return null;
  if(def.factor!==undefined||def.field==='UNITB'){const n=number(value);return n!==null&&n>0?convert(n,def.factor??1):null;}
  if(def.field==='PAIS')return /^[A-Za-z]{2}$/.test(value)?value.toUpperCase():null;
  if(def.field==='ARANCELARIO'){const code=value.replace(/[.-]/g,'');return /^\d{6,10}$/.test(code)?code:null;}
  if(def.field==='EAN')return /^\d{8,14}$/.test(value)?value:null;
  if(def.field==='TIPO CONTENEDOR')return /^(?:BOX(?:ES)?|CARTONS?|PALLETS?|CRATES?|CAJAS?|PALET(?: FITOSANITARIO)?|MEDIO PALET(?: FITOSANITARIO)?|PACKAGES?|BULTOS?|CARTON BOX)$/i.test(value)?value:null;
  return value.length<=70?value:null;
 }
 function inlineFacts(text){
  const facts=[],add=(field,value)=>{if(value!==null&&value!=='')facts.push({field,value});};
  let m=text.match(/\b(?:COUNTRY\s+OF\s+ORIGIN|COO|ORIGIN|ORIGEN|MADE\s+IN)\s*[:=]\s*([A-Z]{2})\b/i);if(m)add('PAIS',m[1].toUpperCase());
  m=text.match(/\b(?:HS\s*(?:CODE)?|HTS\s*(?:CODE)?|TARIFF\s+CODE|PARTIDA\s+ARANCELARIA)\s*[:=]\s*(\d[\d.-]{4,14}\d)\b/i);if(m)add('ARANCELARIO',valueFor({field:'ARANCELARIO'},m[1]));
  m=text.match(/\b(?:EAN|UPC|BARCODE)\s*[:=]\s*(\d{8,14})\b/i);if(m)add('EAN',m[1]);
  m=text.match(/\b(?:UNIT\s+NET\s+(?:WEIGHT|WT)|PESO\s+NETO\s+UNITARIO)\s*[:=]\s*([\d.,]+)\s*(KG|G|LBS?|GRAMS?)\b/i);if(m){const n=number(m[1]);if(n>0)add('KG NETO UD',convert(n,/^KG$/i.test(m[2])?1:/^LBS?$/i.test(m[2])?.45359237:.001));}
  m=text.match(/\b(?:UNITS\s+PER\s+BOX|UNITB)\s*[:=]\s*(\d+)\b/i);if(m&&Number(m[1])>0)add('UNITB',Number(m[1]));
  m=text.match(/\b(?:QTY|QUANTITY\s+SHIPPED|CANTIDAD\s+FACTURA)\s*[:=]\s*([\d.,]+)/i);if(m){const n=number(m[1]);if(n>0)facts.push({field:'CANTIDAD_FACTURA',value:n,readOnly:true});}
  m=text.match(/\b(?:BOX\s+ID|CARTON\s+ID|PACKAGE\s+ID|CAJAS_ID)\s*[:=]\s*([A-Z0-9][A-Z0-9._/-]{0,50})/i);if(m)add('CAJAS_ID',m[1]);
  m=text.match(/\b(?:PACKAGE\s+GROSS\s+WEIGHT|BOX\s+GROSS\s+WEIGHT|PESO\s+BRUTO\s+DEL\s+BULTO)\s*[:=]\s*([\d.,]+)\s*(KG|LBS?)\b/i);if(m){const n=number(m[1]);if(n>0)add('PESOBRUTO',convert(n,/^KG$/i.test(m[2])?1:.45359237));}
  for(const [label,field]of [['LENGTH','LARGO'],['WIDTH','ANCHO'],['HEIGHT','ALTO'],['LARGO','LARGO'],['ANCHO','ANCHO'],['ALTO','ALTO']]){m=text.match(new RegExp('\\b'+label+'\\s*[:=]\\s*([\\d.,]+)\\s*(CM|MM|M|IN|INCHES)\\b','i'));if(m){const n=number(m[1]);if(n>0)add(field,convert(n,{CM:.01,MM:.001,M:1,IN:.0254,INCHES:.0254}[m[2].toUpperCase()]));}}
  m=text.match(/\b(?:PACKAGE\s+TYPE|TIPO\s+CONTENEDOR)\s*[:=]\s*([^;|]{1,40}?)(?=\s+(?:LENGTH|WIDTH|HEIGHT|GROSS|BOX)\b|[;|]|$)/i);if(m)add('TIPO CONTENEDOR',valueFor({field:'TIPO CONTENEDOR'},m[1].trim()));
  return facts;
 }
 function analyze(document,lines){
  const invoiceNos=invoiceNumbers(document.pages.map(page=>page.text).join('\n')),facts=[];
  for(const page of document.pages){
   const pageNos=invoiceNumbers(page.text),invoice=document.invoiceOverride|| (pageNos.length===1?pageNos[0]:invoiceNos.length===1?invoiceNos[0]:'');let header=[];
   page.rows.forEach((row,rowIndex)=>{
    const detected=(row.cells||[]).map(cell=>({...cell,definition:defs[normalize(cell.text)]})).filter(cell=>cell.definition);
    if(detected.length>=2){header=detected.map(cell=>({x:cell.x,definition:cell.definition}));return;}
    const matching=lines.filter(line=>[line.data.ARTICULO,line.source.supplierReference].some(ref=>normalize(ref).length>=3&&matches(row.text,ref)));
    if(!matching.length)return;
    const targets=matching.filter(line=>!invoice||!line.source.invoice||invoiceId(line.source.invoice)===invoiceId(invoice));
    const values=inlineFacts(row.text);if(invoice)values.push({field:'FACTURA',value:invoice});
    header.forEach((column,index)=>{const next=header[index+1];const raw=(row.cells||[]).filter(cell=>cell.x>=column.x-4&&(!next||cell.x<next.x-4)).map(cell=>cell.text).join(' ').trim();const value=valueFor(column.definition,raw);if(value!==null)values.push({field:column.definition.field,value,readOnly:!!column.definition.readOnly});});
    const unique=new Map();values.forEach(value=>unique.set(value.field+'|'+String(value.value),value));
    for(const value of unique.values())facts.push({...value,id:[document.hash,page.page,rowIndex,value.field,String(value.value)].join('|'),docHash:document.hash,file:document.name,page:page.page,row:rowIndex+1,evidence:row.text,invoice,targetIds:targets.map(line=>line.id),reference:[...new Set(matching.map(line=>line.data.ARTICULO))].join(', ')});
   });
  }
  return {invoiceNos,facts};
 }
 return {normalize,number,textRows,inlineFacts,analyze,matches};
});

(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.TMSPackingDomain=api;})(typeof globalThis!=='undefined'?globalThis:this,function(){
 'use strict';
 const columns=['ALBARAN','NPI','PEDIDO_CLIENTE','PEDIDO','CLIENTE','CAJAS_ID','ARTICULO','NOMBRE','EAN','CANTIDAD','PAIS','EU','ARANCELARIO','UNITB','KG NETO UD','NETO TOTAL','PESOBRUTO','PESOBRUTO_LINEA','TIPO CONTENEDOR','LARGO','ANCHO','ALTO','VOLUMEN'];
 const numeric=new Set(['CANTIDAD','UNITB','KG NETO UD','NETO TOTAL','PESOBRUTO','PESOBRUTO_LINEA','LARGO','ANCHO','ALTO','VOLUMEN']);
 const norm=value=>String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]/g,'');
 const aliases={CODIGOCLIENTEDELPRODUCTO:'ARTICULO',DESCRIPCIONDEPRODUCTO:'NOMBRE',DESCRIPCION:'NOMBRE',PACKSLIP:'ALBARAN',CANTIDADSOLICITADA:'CANTIDAD',REFERENCIA:'ARTICULO',CODIGO:'ARTICULO',PESONETOUD:'KG NETO UD',PAISDEORIGEN:'PAIS',PESOBRUTOLINEA:'PESOBRUTO_LINEA',PEDIDODOM:'PEDIDO DOM',PROVEEDOR:'Proveedor',REFERENCIAPROVEEDOR:'Referencia proveedor',FACTURANO:'Factura No.',REFERENCIAEMBARQUE:'Referencia embarque'};
 function number(value){if(typeof value==='number')return Number.isFinite(value)?value:null;const s=String(value??'').trim();if(!s)return null;if(/^-?\d+(?:[.,]\d+)?$/.test(s))return Number(s.replace(',','.'));return null;}
 function parse(grid,file,sheet){
  const rowIndex=grid.findIndex(row=>row.some(cell=>['ARTICULO','CODIGOCLIENTEDELPRODUCTO','REFERENCIA','CODIGO'].includes(norm(cell)))&&row.some(cell=>['CANTIDAD','CANTIDADSOLICITADA'].includes(norm(cell))));
  if(rowIndex<0)throw new Error(file+' · '+sheet+': no se identifican las columnas de artículo y cantidad.');
  const headers=grid[rowIndex].map(value=>columns.find(col=>norm(col)===norm(value))||aliases[norm(value)]||String(value??'').trim()),lines=[],errors=[];
  grid.slice(rowIndex+1).forEach((values,index)=>{
   if(!values.some(value=>value!==''&&value!=null))return;
   const source=Object.create(null);headers.forEach((header,i)=>{if(header)source[header]=values[i]??'';});
   if(/^total\b/i.test(String(source.CAJAS_ID??'')))return;
   const article=String(source.ARTICULO??'').trim(),quantity=number(source.CANTIDAD);
   if(!article){if(/^(INCOTERM\b|TOTAL UNIDADES$)/i.test(String(source.NOMBRE||'')))return;if(quantity!==null&&!source.NOMBRE&&!source.CAJAS_ID&&!source.ALBARAN&&!source['PEDIDO DOM'])return;errors.push({file,sheet,row:rowIndex+index+2,message:'Falta el artículo; la fila no se incorporó.'});return;}
   if(quantity===null||quantity<=0){errors.push({file,sheet,row:rowIndex+index+2,message:'Cantidad ausente, inválida o no positiva para '+article});return;}
   const data=Object.fromEntries(columns.map(col=>[col,numeric.has(col)?number(source[col]):String(source[col]??'').trim()]));data.ARTICULO=article;data.CANTIDAD=quantity;
   lines.push({id:[file,sheet,rowIndex+index+2].join('|'),data,source:{file,sheet,row:rowIndex+index+2,quantity,pedidoDom:String(source['PEDIDO DOM']??''),supplier:String(source.Proveedor??''),supplierReference:String(source['Referencia proveedor']??''),invoice:String(source['Factura No.']??''),shipment:String(source['Referencia embarque']??'')}});
  });
  if(!lines.length&&!errors.length)throw new Error(file+' · '+sheet+': sin líneas de materiales.');
  return {lines,errors};
 }
 function build(input,shipment={}){
  const issues=[],packages=new Map(),lines=input.map(item=>({id:item.id,source:{...item.source},data:{...item.data}}));
  lines.forEach(line=>{
   const d=line.data;if(!d.CLIENTE)d.CLIENTE=shipment.customer||'';
   if(shipment.orderColumn&&line.source.pedidoDom&&!d[shipment.orderColumn])d[shipment.orderColumn]=line.source.pedidoDom;
   if(d['KG NETO UD']!==null&&d['KG NETO UD']>0)d['NETO TOTAL']=d['KG NETO UD']*d.CANTIDAD;
   if(['LARGO','ANCHO','ALTO'].every(key=>d[key]!==null&&d[key]>0))d.VOLUMEN=d.LARGO*d.ANCHO*d.ALTO;
   const missing=['CAJAS_ID','PAIS','ARANCELARIO','UNITB','KG NETO UD','PESOBRUTO','TIPO CONTENEDOR','LARGO','ANCHO','ALTO'].filter(key=>d[key]===''||d[key]===null||(numeric.has(key)&&d[key]<=0));
   if(!d.NOMBRE)missing.push('NOMBRE');
   if(missing.length)issues.push({id:line.id,article:d.ARTICULO,field:missing.join(', '),message:'Datos pendientes'});
   if(d.CAJAS_ID){if(!packages.has(d.CAJAS_ID))packages.set(d.CAJAS_ID,[]);packages.get(d.CAJAS_ID).push(line);}
  });
  const allIds=!!lines.length&&lines.every(line=>line.data.CAJAS_ID);let gross=0,volume=0,grossComplete=allIds,volumeComplete=allIds;
  for(const [id,items] of packages){
   const consistent=key=>new Set(items.map(line=>line.data[key]).filter(value=>value!==null&&value!=='')).size===1;
   const grossOk=consistent('PESOBRUTO')&&items.every(line=>line.data.PESOBRUTO>0),volumeOk=['LARGO','ANCHO','ALTO'].every(key=>consistent(key)&&items.every(line=>line.data[key]>0));
   if(!grossOk||!volumeOk)issues.push({id,article:'',field:'Bulto '+id,message:'Faltan pesos/medidas o hay valores incompatibles entre sus líneas.'});
   if(grossOk)gross+=items[0].data.PESOBRUTO;else grossComplete=false;
   if(volumeOk)volume+=items[0].data.VOLUMEN;else volumeComplete=false;
   const net=items.reduce((sum,line)=>sum+(line.data['NETO TOTAL']||0),0),allNet=items.every(line=>line.data['NETO TOTAL']>0);
   if(grossOk&&allNet){items.forEach(line=>line.data.PESOBRUTO_LINEA=items[0].data.PESOBRUTO*line.data['NETO TOTAL']/net);}
   else items.forEach(line=>line.data.PESOBRUTO_LINEA=null);
  }
  ['reference','container','seal','customer'].forEach(key=>{if(!String(shipment[key]||'').trim())issues.push({id:'embarque',article:'',field:key,message:'Dato del consolidado pendiente'});});
  if(lines.some(line=>line.source.pedidoDom)&&!['NPI','PEDIDO_CLIENTE','PEDIDO'].includes(shipment.orderColumn))issues.push({id:'embarque',article:'',field:'PEDIDO DOM',message:'Selecciona la columna de destino del pedido antes de generar el packing list final.'});
  return {lines,issues,packages:allIds?packages.size:null,gross:grossComplete?gross:null,volume:volumeComplete?volume:null,quantity:lines.reduce((sum,line)=>sum+line.data.CANTIDAD,0),shipment:{...shipment},complete:!!lines.length&&!issues.length};
 }
 return {columns,numeric,number,parse,build};
});

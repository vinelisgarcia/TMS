(function(root,factory){const api=factory(root);if(typeof module==='object'&&module.exports)module.exports=api;else root.InventoryDailyExcel=api;})(typeof globalThis==='object'?globalThis:this,function(root){
 'use strict';
 const headers=['Fecha de conteo','Referencia','Descripción','Ubicación prevista','Ubicación contada','Unidad','Cantidad física','Contado por','Observaciones','ID de fila'];
 function compactWorkbook(ExcelJS,template){
  const book=new ExcelJS.Workbook();book.creator='Hospitality by Alvarez · TMS';book.title='Conteo diario del showroom';book.created=new Date(template.generatedAt);
  const sheet=book.addWorksheet('Conteo',{views:[{state:'frozen',ySplit:1,showGridLines:false}],pageSetup:{orientation:'portrait',paperSize:9,fitToPage:true,fitToWidth:1}});
  sheet.columns=[{header:'Código del material',key:'code',width:32},{header:'Cantidad contada',key:'quantity',width:25}];
  sheet.getRow(1).height=32;sheet.getRow(1).eachCell(cell=>{cell.font={name:'Arial',size:11,bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF183F66'}};cell.alignment={vertical:'middle',wrapText:true};});
  for(const item of template.rows){const row=sheet.addRow({code:item.code});row.height=25;row.getCell(1).numFmt='@';row.getCell(2).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF4D5'}};row.getCell(2).dataValidation={type:'decimal',operator:'greaterThanOrEqual',allowBlank:true,showErrorMessage:true,error:'Indica una cantidad mayor o igual que cero.',formulae:[0]};}
  sheet.autoFilter={from:'A1',to:'B'+(template.rows.length+1)};sheet.pageSetup.printArea='A1:B'+(template.rows.length+1);sheet.pageSetup.printTitlesRow='1:1';
  const meta=book.addWorksheet('_TMS',{state:'veryHidden'});meta.addRows([['TMS_DAILY_COUNT',2],['kind',template.kind],['templateId',template.templateId],['cycleId',template.cycleId],['sourcePair',template.sourcePair],['generatedAt',template.generatedAt],['day',template.day],['ID','Referencia','Ubicación prevista','Unidad'],...template.rows.map(row=>[row.id,row.code,row.expectedLocation||null,row.unit])]);
  return book;
 }
 function createWorkbook(ExcelJS,template){
  if(template.version===2&&template.kind==='showroom')return compactWorkbook(ExcelJS,template);
  const book=new ExcelJS.Workbook(),blue='FF183F66',ink='FF23364B',amber='FFFFF4D5';book.creator='Hospitality by Alvarez · TMS';book.title='Plantilla de conteo físico diario';book.created=new Date(template.generatedAt);
  const sheet=book.addWorksheet('Conteo',{views:[{state:'frozen',xSplit:2,ySplit:7,showGridLines:false}],pageSetup:{orientation:'landscape',paperSize:9,fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.4,bottom:.4,header:.15,footer:.15}},properties:{tabColor:{argb:blue},defaultRowHeight:26}});
  sheet.columns=[16,20,42,22,22,12,18,25,38,40].map((width,index)=>({width,hidden:index===9}));
  const fill=color=>({type:'pattern',pattern:'solid',fgColor:{argb:color}}),font={name:'Arial',size:11,color:{argb:ink}};
  const title=(row,text,size=11)=>{sheet.mergeCells(row,1,row,9);const cell=sheet.getCell(row,1);cell.value=text;cell.font={...font,size,bold:row===2};cell.alignment={vertical:'middle',wrapText:true};};
  sheet.getRow(1).height=8;title(2,'Conteo físico diario',17);sheet.getRow(2).height=34;
  title(3,template.kind==='showroom'?'Inventario del showroom':'Inventario CT / AD');sheet.getRow(3).height=26;
  title(4,'Completa las celdas amarillas. Vacío = sin contar. Escribe 0 cuando hayas confirmado que no hay unidades.');sheet.getRow(4).height=28;
  title(5,'Confirma la fecha, la ubicación contada y el nombre del responsable. Puedes filtrar y repartir las referencias entre el equipo.');sheet.getRow(5).height=30;
  const labels=template.kind==='showroom'?['Físico','Virtual']:['CT','AD'];title(6,`Cortes de referencia: ${labels[0]} ${template.sourceDates.ct||'no cargado'}. ${labels[1]} ${template.sourceDates.ad||'no cargado'}.`);sheet.getRow(6).height=26;
  const head=sheet.getRow(7);head.values=headers;head.height=32;head.eachCell(cell=>{cell.font={...font,bold:true,color:{argb:'FFFFFFFF'}};cell.fill=fill(blue);cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};});
  template.rows.forEach((item,index)=>{
   const row=sheet.getRow(index+8);row.values=[new Date(template.day+'T00:00:00Z'),item.code,item.description||null,item.expectedLocation||null,item.expectedLocation||null,item.unit,null,null,null,item.id];row.height=Math.max(34,Math.ceil((item.description||'').length/42)*14+8,Math.ceil((item.expectedLocation||'').length/22)*14+8);
   for(let column=1;column<=10;column++){const cell=row.getCell(column);cell.font={...font};cell.alignment={vertical:'middle',wrapText:[3,4,5,8,9].includes(column),horizontal:[1,6].includes(column)?'center':column===7?'right':'left'};cell.fill=fill([1,5,7,8,9].includes(column)?amber:index%2?'FFF2F5F9':'FFFFFFFF');}
   for(const column of [2,4,5,10])row.getCell(column).numFmt='@';row.getCell(1).numFmt='dd/mm/yyyy';row.getCell(7).numFmt='#,##0.###;[Red](#,##0.###);0';
   row.getCell(7).dataValidation={type:'decimal',operator:'greaterThanOrEqual',allowBlank:true,showErrorMessage:true,errorTitle:'Cantidad física',error:'Escribe un número mayor o igual que cero.',formulae:[0]};
   if(item.unitConflict){row.getCell(6).font={...font,bold:true,color:{argb:'FFB42318'}};row.getCell(9).value='Revisar unidad antes de contar.';}
  });
  const last=template.rows.length+7;sheet.autoFilter={from:'A7',to:'J'+last};sheet.pageSetup.printTitlesRow='1:7';sheet.pageSetup.printArea='A1:I'+last;sheet.headerFooter.oddFooter='&LHospitality by Alvarez · Conteo físico&R Página &P de &N';
  const meta=book.addWorksheet('_TMS',{state:'veryHidden'});meta.addRows([['TMS_DAILY_COUNT',1],['kind',template.kind],['templateId',template.templateId],['cycleId',template.cycleId],['sourcePair',template.sourcePair],['generatedAt',template.generatedAt],['ID','Referencia','Ubicación prevista','Unidad'],...template.rows.map(row=>[row.id,row.code,row.expectedLocation||null,row.unit])]);
  return book;
 }
 function decodeSimple(book,XLSX){
  const normalize=value=>String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const codes=new Set(['codigo','codigomaterial','codigodelmaterial','codigoarticulo','codigodelarticulo','codigosap','codigomaterialsap','referencia','referenciaarticulo','material','articulo','sku','producto']);
  const quantities=new Set(['cantidad','cantidadcontada','cantidadfisica','conteo','contado','unidadescontadas','fisico','cantidadfisicacontada','conteofisico','totalcontado','unidades','cantidadreal']);
  for(const name of book.SheetNames){
   const sheet=book.Sheets[name],matrix=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false,defval:null,range:0});
   for(let i=0;i<Math.min(matrix.length,20);i++){
    const row=matrix[i]||[],code=row.findIndex(value=>codes.has(normalize(value))),qty=row.findIndex(value=>quantities.has(normalize(value)));
    if(code<0||qty<0||code===qty)continue;
    const formulaRows=[];for(const [address,cell]of Object.entries(sheet))if(!address.startsWith('!')&&cell.f&&XLSX.utils.decode_cell(address).c===qty){const n=XLSX.utils.decode_cell(address).r+1;if(n>i+1)formulaRows.push(n-i);}
    return {matrix:[['Código del material','Cantidad contada'],...matrix.slice(i+1).map(values=>[values[code],values[qty]])],formulaRows};
   }
  }
  throw Error('No se encontraron las columnas Código del material y Cantidad contada. Descarga el formato de ejemplo y utiliza esos encabezados.');
 }
 function decode(book,XLSX){
  const sheet=book.Sheets.Conteo,meta=book.Sheets._TMS;if(!sheet||!meta)throw Error('Usa la plantilla de conteo diario descargada desde el TMS.');
  const matrix=XLSX.utils.sheet_to_json(meta,{header:1,raw:true,defval:null,range:0});if(matrix[0]?.[0]!=='TMS_DAILY_COUNT'||![1,2].includes(matrix[0]?.[1]))throw Error('Formato de plantilla no reconocido.');
  const version=matrix[0][1],values=Object.fromEntries(matrix.slice(1,version===2?7:6).map(row=>[row[0],row[1]]));
  const template={version,...values,rows:matrix.slice(version===2?8:7).filter(row=>row[0]).map(row=>({id:String(row[0]),code:String(row[1]??''),expectedLocation:String(row[2]??''),unit:String(row[3]??'')}))};
  const formulaRows=[];for(const [address,cell]of Object.entries(sheet))if(!address.startsWith('!')&&cell.f&&XLSX.utils.decode_cell(address).c===(version===2?1:6))formulaRows.push(XLSX.utils.decode_cell(address).r+1);
  return {template,matrix:XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:null,range:0}),options:{formulaRows,date1904:!!book.Workbook?.WBProps?.date1904}};
 }
 return {createWorkbook,decode,decodeSimple,load:()=>root.ShowroomExecutiveExcel.load()};
});

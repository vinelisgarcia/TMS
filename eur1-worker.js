'use strict';
importScripts('./vendor/fflate-0.8.2.min.js','./eur1-archive.js','./eur1-domain.js');
const acknowledgements=new Map();let excelLoaded=false;
self.onmessage=async event=>{
 const {id,command,payload}=event.data;
 if(command==='ack'){const pending=acknowledgements.get(payload.key);if(pending){acknowledgements.delete(payload.key);payload.error?pending.reject(Error(payload.error)):pending.resolve();}return;}
 try{
  let result;
  if(command==='unpack'){
   let n=0;result=await TMSEUR1Archive.unpack(payload.buffer,fflate,async entry=>{const key=id+':'+(++n);const ack=new Promise((resolve,reject)=>acknowledgements.set(key,{resolve,reject}));self.postMessage({id,entry:{name:entry.name,buffer:entry.data.buffer},key},[entry.data.buffer]);await ack;});
  }else if(command==='xlsx'){
   const archive=TMSEUR1Archive.inspect(payload.buffer,{files:2048,total:100*1024*1024,file:50*1024*1024,ratio:1500});if(archive.entries.some(e=>/vbaProject\.bin$/i.test(e.name)))throw Error('El libro contiene macros. Guarda una copia XLSX sin macros.');
   let xmlCells=0;await TMSEUR1Archive.unpack(payload.buffer,fflate,entry=>{if(/^xl\/worksheets\/sheet.*\.xml$/i.test(entry.name)){const xml=new TextDecoder().decode(entry.data);for(const m of xml.matchAll(/<row\b[^>]*\br="(\d+)"/g))if(Number(m[1])>50000)throw Error('Una hoja supera 50.000 filas.');for(const m of xml.matchAll(/<c\b[^>]*\br="([A-Z]+)\d+"/g)){if(++xmlCells>150000)throw Error('El libro supera 150.000 celdas.');let col=0;for(const c of m[1])col=col*26+c.charCodeAt(0)-64;if(col>256)throw Error('Una hoja supera 256 columnas.');}}},{files:2048,total:100*1024*1024,file:50*1024*1024,ratio:1500});
   if(!excelLoaded){importScripts('./vendor/exceljs-4.4.0.min.js');excelLoaded=true;}
   const wb=new ExcelJS.Workbook();await wb.xlsx.load(payload.buffer);let count=0;const sheets=[];
   for(const sheet of wb.worksheets){if(sheet.rowCount>50000||sheet.columnCount>256)throw Error('La hoja '+sheet.name+' supera 50.000 filas o 256 columnas. Divide el documento.');const rows=[];
    sheet.eachRow((row,r)=>{const cells=Array.from({length:sheet.columnCount},()=>({v:''}));row.eachCell((cell,c)=>{if(++count>150000)throw Error('El libro supera 150.000 celdas con datos.');let v=cell.value,problem=null,formulaCached=false;
     if(v&&typeof v==='object'&&('formula'in v||'sharedFormula'in v)){formulaCached=true;if(v.result===undefined||v.result===null){v='';problem='Fórmula sin resultado guardado; abre el original, recalcula y guarda una copia. El TMS no ejecuta fórmulas.';}else v=v.result;}
     if(v&&typeof v==='object'){if(v instanceof Date)v=v.toISOString().slice(0,10);else if(v.richText)v=v.richText.map(x=>x.text).join('');else if(v.text!==undefined)v=v.text;else if(v.error){problem='Error guardado en la celda: '+v.error;v='';}else v='';}
     if(typeof v==='number'&&/^0{2,}$/.test(cell.numFmt||''))v=String(v).padStart(cell.numFmt.length,'0');
     cells[c-1]={v:v??'',cell:cell.isMerged?cell.master.address:cell.address,formulaCached,problem};});if(cells.some(c=>c.v!==''||c.problem))rows.push({row:r,cells});});sheets.push({name:sheet.name,rows});
   }
   result=TMSEUR1Domain.parse({format:'xlsx',sheets},payload.info,payload.kind);
  }else if(command==='parse')result=TMSEUR1Domain.parse(payload.raw,payload.info,payload.kind);
  else if(command==='analyze')result=TMSEUR1Domain.analyze(payload.documents,payload.options);
  else throw Error('Etapa no reconocida.');
  self.postMessage({id,result});
 }catch(error){self.postMessage({id,error:error.message});}
};

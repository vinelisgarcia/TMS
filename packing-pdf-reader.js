(function(){
 'use strict';
 let library;
 async function read(buffer,name,onProgress){
  try{library=library||await import('./vendor/pdfjs-6.5.438.min.mjs');}catch(error){throw new Error('No se pudo cargar el lector PDF. Revisa la conexión y vuelve a cargar el TMS.');}
  library.GlobalWorkerOptions.workerSrc=new URL('./vendor/pdfjs-worker-6.5.438.min.mjs',document.baseURI).href;
  const task=library.getDocument({data:new Uint8Array(buffer),isEvalSupported:false,enableXfa:false,useSystemFonts:true});let doc;
  try{
   doc=await task.promise;if(doc.numPages>300)throw new Error('El PDF tiene más de 300 páginas. Divide el consolidado en documentos más pequeños.');const pages=[];
   for(let number=1;number<=doc.numPages;number++){
    onProgress?.(number,doc.numPages);const page=await doc.getPage(number),content=await page.getTextContent(),rows=TMSPackingPDFDomain.textRows(content.items);pages.push({page:number,rows,text:rows.map(row=>row.text).join('\n')});page.cleanup();
   }
   return {name,pages,unreadablePages:pages.filter(page=>!page.text.trim()).map(page=>page.page)};
  }catch(error){if(error.name==='PasswordException')throw new Error('El PDF está protegido con contraseña. Carga una copia que permita su lectura.');throw new Error(error.message||'No se pudo leer el PDF.');}finally{await task.destroy();}
 }
 window.TMSPackingPDFReader={read};
})();

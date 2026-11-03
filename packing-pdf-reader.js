(function(){
 'use strict';
 let library,ocrLibrary;
 const asset=path=>new URL('./vendor/'+path,document.baseURI).href;
 const aborted=()=>new DOMException('Lectura cancelada.','AbortError');
 function withSignal(promise,signal){
  if(!signal)return promise;if(signal.aborted)return Promise.reject(aborted());
  return new Promise((resolve,reject)=>{const stop=()=>reject(aborted());signal.addEventListener('abort',stop,{once:true});Promise.resolve(promise).then(resolve,reject).finally(()=>signal.removeEventListener('abort',stop));});
 }
 function loadOCR(){
  if(window.Tesseract)return Promise.resolve();
  if(!ocrLibrary)ocrLibrary=new Promise((resolve,reject)=>{const script=document.createElement('script');script.src=asset('tesseract-7.0.0/tesseract.min.js');script.onload=resolve;script.onerror=()=>{script.remove();ocrLibrary=null;reject(new Error('No se pudo cargar OCR. Revisa la conexión y vuelve a intentarlo.'));};document.head.appendChild(script);});
  return ocrLibrary;
 }
 // Measure image coverage in PDF points, including images under a text footer.
 function needsOCR(operators,text,area){
  const O=library.OPS,images=new Set([O.paintImageXObject,O.paintInlineImageXObject,O.paintImageMaskXObject]),stack=[];let matrix=[1,0,0,1,0,0],coverage=0,count=0;
  operators.fnArray.forEach((op,index)=>{
   if(op===O.save)stack.push([...matrix]);else if(op===O.restore)matrix=stack.pop()||[1,0,0,1,0,0];
   else if(op===O.transform){const [a,b,c,d,e,f]=matrix,[g,h,i,j,k,l]=operators.argsArray[index];matrix=[a*g+c*h,b*g+d*h,a*i+c*j,b*i+d*j,a*k+c*l+e,b*k+d*l+f];}
   else if(images.has(op)){coverage+=Math.abs(matrix[0]*matrix[3]-matrix[1]*matrix[2]);count++;}
  });
  return !text.trim()||coverage/area>.12||(count>0&&text.length<200)||count>8||/\uFFFD/.test(text)||(text.length<150&&operators.fnArray.length>500);
 }
 function isBlank(canvas){
  const sample=document.createElement('canvas'),scale=Math.min(1,512/Math.max(canvas.width,canvas.height));sample.width=Math.max(1,Math.round(canvas.width*scale));sample.height=Math.max(1,Math.round(canvas.height*scale));
  const ctx=sample.getContext('2d',{willReadFrequently:true});ctx.drawImage(canvas,0,0,sample.width,sample.height);const pixels=ctx.getImageData(0,0,sample.width,sample.height).data;let ink=0;
  for(let i=0;i<pixels.length;i+=4)if(Math.min(pixels[i],pixels[i+1],pixels[i+2])<225)ink++;
  sample.width=sample.height=0;return ink<3;
 }
 function preview(canvas){const image=document.createElement('canvas'),scale=Math.min(1,1100/canvas.width);image.width=Math.round(canvas.width*scale);image.height=Math.round(canvas.height*scale);image.getContext('2d').drawImage(canvas,0,0,image.width,image.height);const url=image.toDataURL('image/jpeg',.8);image.width=image.height=0;return url;}
 async function read(buffer,name,onProgress,settings={}){
  const {signal,onPassword}=settings;let task,worker,workerPromise,pageNumber=0,total=0,renderTask;
  const progress=(stage,value)=>onProgress?.(pageNumber,total,{stage,progress:value});
  const stop=()=>{renderTask?.cancel();task?.destroy().catch(()=>{});worker?.terminate().catch(()=>{});};
  signal?.addEventListener('abort',stop,{once:true});
  try{
   library=library||await withSignal(import('./vendor/pdfjs-6.5.438.min.mjs'),signal);
   if(signal?.aborted)throw aborted();
   library.GlobalWorkerOptions.workerSrc=asset('pdfjs-worker-6.5.438.min.mjs');
   task=library.getDocument({data:new Uint8Array(buffer),isEvalSupported:false,enableXfa:false,useSystemFonts:true,cMapUrl:asset('pdfjs-assets-6.5.438/cmaps/'),cMapPacked:true,standardFontDataUrl:asset('pdfjs-assets-6.5.438/standard_fonts/'),wasmUrl:asset('pdfjs-assets-6.5.438/wasm/')});
   task.onPassword=(update,reason)=>{
    if(!onPassword){task.destroy().catch(()=>{});return;}
    Promise.resolve(onPassword({name,incorrect:reason===library.PasswordResponses.INCORRECT_PASSWORD,signal})).then(password=>{if(password===null||signal?.aborted)stop();else update(password);}).catch(stop);
   };
   const doc=await withSignal(task.promise,signal);total=doc.numPages;if(total>300)throw new Error('El PDF tiene más de 300 páginas. Divídelo en documentos más pequeños.');const pages=[];
   const getWorker=()=>workerPromise||(workerPromise=(async()=>{
    progress('load');await loadOCR();
    const created=await Tesseract.createWorker(['eng','spa'],1,{workerPath:asset('tesseract-7.0.0/worker.min.js'),corePath:asset('tesseract-core-7.0.0/'),langPath:asset('tesseract-languages'),workerBlobURL:false,logger:message=>{if(message.status==='recognizing text')progress('ocr',message.progress);}});
    if(signal?.aborted){await created.terminate();throw aborted();}worker=created;
    await worker.setParameters({preserve_interword_spaces:'1',user_defined_dpi:'216'});return worker;
   })());
   for(pageNumber=1;pageNumber<=total;pageNumber++){
    if(signal?.aborted)throw aborted();progress('text');let page,canvas;const result={page:pageNumber,rows:[],text:'',method:'text',confidence:null,reviewRequired:false};
    try{
     page=await withSignal(doc.getPage(pageNumber),signal);const content=await withSignal(page.getTextContent(),signal),native=TMSPackingPDFDomain.textRows(content.items),text=native.map(row=>row.text).join('\n'),viewport=page.getViewport({scale:1});result.rows=native;result.text=text;
     const operators=await withSignal(page.getOperatorList(),signal);
     if(needsOCR(operators,text,viewport.width*viewport.height)){
      progress('render');const scale=Math.min(3,3500/Math.max(viewport.width,viewport.height),Math.sqrt(7500000/(viewport.width*viewport.height))),view=page.getViewport({scale});canvas=document.createElement('canvas');canvas.width=Math.ceil(view.width);canvas.height=Math.ceil(view.height);
      renderTask=page.render({canvasContext:canvas.getContext('2d'),viewport:view,background:'rgb(255,255,255)'});await withSignal(renderTask.promise,signal);renderTask=null;
      if(isBlank(canvas)){result.method='blank';}
      else{
       result.preview=preview(canvas);result.method=text.trim()?'mixed':'ocr';const engine=await withSignal(getWorker(),signal);progress('ocr',0);const {data}=await withSignal(engine.recognize(canvas,{}, {text:true,blocks:true}),signal),rows=TMSPackingPDFDomain.ocrRows(data);
       // Native text supplements OCR; never discard a legible header or footer.
       const seen=new Set(rows.map(row=>TMSPackingPDFDomain.normalize(row.text)));const extra=native.filter(row=>!seen.has(TMSPackingPDFDomain.normalize(row.text)));
       result.rows=[...rows,...extra];result.text=result.rows.map(row=>row.text).join('\n');result.method=text.trim()?'mixed':'ocr';result.confidence=Math.round(data.confidence||0);result.reviewRequired=result.confidence<75||!data.text?.trim();
       if(!data.text?.trim())result.error='OCR no pudo reconocer el contenido visible. Revisa la imagen o carga una copia más nítida.';
      }
     }
    }catch(error){if(signal?.aborted||error.name==='AbortError')throw aborted();result.error=error.message||'No se pudo leer esta página.';result.reviewRequired=true;}
    finally{if(canvas)canvas.width=canvas.height=0;page?.cleanup();}
    pages.push(result);
   }
   return {name,pages,unreadablePages:pages.filter(page=>page.error).map(page=>page.page)};
  }catch(error){if(signal?.aborted||error.name==='AbortError')throw aborted();if(error.name==='PasswordException')throw new Error('No se pudo abrir el PDF con esa contraseña.');throw new Error(error.message||'No se pudo leer el PDF.');}
  finally{signal?.removeEventListener('abort',stop);await worker?.terminate().catch(()=>{});await task?.destroy().catch(()=>{});}
 }
 window.TMSPackingPDFReader={read};
})();

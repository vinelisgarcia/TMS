(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.TMSEUR1Archive=api;})(typeof self!=='undefined'?self:globalThis,function(){
 'use strict';
 const MB=1024*1024,LIMITS={compressed:50*MB,file:50*MB,total:250*MB,files:100,ratio:200,depth:12};
 function path(name){if(!name||name.includes('\0')||/^(?:[\\/]|[a-z]:)/i.test(name))throw Error('Ruta no permitida en ZIP: '+name);const parts=name.replace(/\\/g,'/').split('/');if(parts.some(p=>p==='..'||p==='.')||parts.length>LIMITS.depth)throw Error('Ruta insegura o demasiado profunda: '+name);return parts.join('/');}
 function decode(bytes,utf8){if(utf8)return new TextDecoder('utf-8',{fatal:true}).decode(bytes);let s=new TextDecoder('utf-8').decode(bytes);return s.includes('\uFFFD')?new TextDecoder('windows-1252').decode(bytes):s;}
 function inspect(buffer,settings={}){
  const limits={...LIMITS,...settings},bytes=buffer instanceof Uint8Array?buffer:new Uint8Array(buffer),v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);if(bytes.length>limits.compressed)throw Error('El archivo supera el límite de '+limits.compressed/MB+' MB.');if(bytes.length<22)throw Error('ZIP incompleto.');
  let end=-1;for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(v.getUint32(i,true)===0x06054b50&&i+22+v.getUint16(i+20,true)===bytes.length){end=i;break;}
  if(end<0)throw Error('No se encuentra el directorio del ZIP.');
  const count=v.getUint16(end+10,true),size=v.getUint32(end+12,true),offset=v.getUint32(end+16,true);if(v.getUint16(end+4,true)||v.getUint16(end+6,true)||count!==v.getUint16(end+8,true)||count===65535||size===0xffffffff||offset===0xffffffff)throw Error('ZIP dividido o ZIP64 no admitido.');if(count>limits.files)throw Error('El ZIP contiene más de '+limits.files+' entradas.');if(offset+size>end)throw Error('Directorio ZIP fuera de límites.');
  const entries=[],names=new Set();let pos=offset,total=0;
  for(let i=0;i<count;i++){
   if(pos+46>end||v.getUint32(pos,true)!==0x02014b50)throw Error('Directorio ZIP dañado.');const flags=v.getUint16(pos+8,true),method=v.getUint16(pos+10,true),crc=v.getUint32(pos+16,true),compressed=v.getUint32(pos+20,true),uncompressed=v.getUint32(pos+24,true),nameLength=v.getUint16(pos+28,true),extraLength=v.getUint16(pos+30,true),commentLength=v.getUint16(pos+32,true),attrs=v.getUint32(pos+38,true),local=v.getUint32(pos+42,true);
   if(pos+46+nameLength+extraLength+commentLength>end)throw Error('Entrada ZIP incompleta.');const name=path(decode(bytes.subarray(pos+46,pos+46+nameLength),!!(flags&2048)));
   if(names.has(name))throw Error('El ZIP repite una ruta: '+name);names.add(name);
   if(flags&1)throw Error('ZIP cifrado: abre el archivo y carga los documentos sin contraseña.');if(![0,8].includes(method))throw Error('Método de compresión no admitido: '+name);if(((attrs>>>16)&0xf000)===0xa000)throw Error('El ZIP contiene un enlace simbólico: '+name);
   if(local+30>offset||v.getUint32(local,true)!==0x04034b50)throw Error('Cabecera local ZIP inválida.');const localFlags=v.getUint16(local+6,true),localMethod=v.getUint16(local+8,true),localNameLen=v.getUint16(local+26,true),localExtra=v.getUint16(local+28,true),start=local+30+localNameLen+localExtra;
   if(start+compressed>offset||localMethod!==method||localFlags!==flags||decode(bytes.subarray(local+30,local+30+localNameLen),!!(flags&2048))!==name)throw Error('Las cabeceras del ZIP no coinciden: '+name);
   if(uncompressed>limits.file||uncompressed>Math.max(1,compressed)*limits.ratio)throw Error('Tamaño o expansión excesiva: '+name);total+=uncompressed;if(total>limits.total)throw Error('El contenido extraído supera '+limits.total/MB+' MB.');
   entries.push({name,compressed,uncompressed,crc,method,start,directory:name.endsWith('/')});pos+=46+nameLength+extraLength+commentLength;
  }
  if(pos!==offset+size)throw Error('Longitud del directorio ZIP inconsistente.');return {entries,total,limits};
 }
 const crcTable=Uint32Array.from({length:256},(_,n)=>{let c=n;for(let i=0;i<8;i++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c>>>0;});
 function crc32(bytes){let c=0xffffffff;for(const n of bytes)c=crcTable[(c^n)&255]^(c>>>8);return (c^0xffffffff)>>>0;}
 async function unpack(buffer,fflate,onEntry,settings={}){
  const bytes=buffer instanceof Uint8Array?buffer:new Uint8Array(buffer),manifest=inspect(bytes,settings);let actualTotal=0;
  // One file at a time, in a Web Worker; output is bounded independently of the
  // central directory. Await the consumer before inflating another document.
  for(const entry of manifest.entries){
   if(entry.directory)continue;const parts=entry.name.split('/');if(parts.includes('__MACOSX')||parts.at(-1)==='.DS_Store')continue;
   const chunks=[];let size=0,error=null;
   const consume=chunk=>{size+=chunk.length;actualTotal+=chunk.length;if(size>manifest.limits.file||size>entry.uncompressed||actualTotal>manifest.limits.total)throw Error('Expansión real del ZIP fuera de límites: '+entry.name);chunks.push(chunk);};
   if(entry.method===0)consume(bytes.slice(entry.start,entry.start+entry.compressed));
   else{const inflate=new fflate.Inflate((chunk)=>consume(chunk));try{for(let p=entry.start;p<entry.start+entry.compressed;p+=16384)inflate.push(bytes.subarray(p,Math.min(p+16384,entry.start+entry.compressed)),p+16384>=entry.start+entry.compressed);}catch(e){error=e;}if(error)throw error;}
   if(size!==entry.uncompressed)throw Error('Tamaño extraído distinto del declarado: '+entry.name);const result=new Uint8Array(size);let pos=0;for(const c of chunks){result.set(c,pos);pos+=c.length;}if(crc32(result)!==entry.crc)throw Error('El archivo no supera la comprobación de integridad: '+entry.name);
   await onEntry({name:entry.name,data:result});
  }
  return manifest.entries.length;
 }
 return {LIMITS,path,inspect,unpack,crc32};
});

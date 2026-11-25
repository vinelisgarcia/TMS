// A per-account acceleration cache; every use is checked against the server revision.
(function(root){
 'use strict';
 const dbName='tms-shared-snapshot-v1',storeName='snapshots',ttl=24*60*60*1000;
 function open(){return new Promise((resolve,reject)=>{if(!root.indexedDB)return reject(Error('Almacenamiento no disponible'));const request=root.indexedDB.open(dbName,1);request.onupgradeneeded=()=>request.result.createObjectStore(storeName);request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);request.onblocked=()=>reject(Error('Almacenamiento ocupado'));});}
 async function operation(mode,fn){const db=await open();try{return await new Promise((resolve,reject)=>{const tx=db.transaction(storeName,mode),request=fn(tx.objectStore(storeName));request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);tx.onabort=()=>reject(tx.error);});}finally{db.close();}}
 async function read(userId){try{const entry=await operation('readonly',store=>store.get(userId));if(!entry||Date.now()-entry.at>ttl||entry.userId!==userId||!Number.isSafeInteger(entry.snapshot?.revision)||!entry.snapshot?.tables)return null;return entry.snapshot;}catch(_){return null;}}
 async function write(userId,snapshot){try{if(!userId||!snapshot?.tables||!Number.isSafeInteger(Number(snapshot.revision)))return;await operation('readwrite',store=>store.put({userId,at:Date.now(),snapshot},userId));}catch(_){} }
 async function remove(userId){try{await operation('readwrite',store=>store.delete(userId));}catch(_){} }
 root.TMSSnapshotCache={read,write,remove};
})(typeof window==='undefined'?globalThis:window);

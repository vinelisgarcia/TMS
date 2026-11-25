// Cached data is never trusted until the server confirms the revision and user.
const {test}=require('node:test'),assert=require('node:assert/strict'),{create}=require('../tms-sync.js');
const copy=x=>structuredClone(x);
test('new tabs validate the shared revision and reuse only the same account’s complete snapshot',async t=>{
 const disk=new Map(),requests=[],db={revision:4,tables:{tms_settings:[{clave:'app_state',valor:{value:'original'}}]}};let forbidden=false;
 const cache={read:async id=>copy(disk.get(id)||null),write:async(id,value)=>disk.set(id,copy(value))};
 const client={rpc:async(name,args)=>{assert.equal(name,'tms_get_snapshot');requests.push(copy(args));if(forbidden)return {error:{code:'42501',message:'Revoked'}};return {data:args.p_known_revision===db.revision?{revision:db.revision,unchanged:true}:copy(db),error:null};}};
 function tab(){let local={};const controller=create({getClient:()=>client,capture:()=>local,applySnapshot:async value=>{local=copy(value.tables)},snapshotCache:cache,pollMs:3600000});t.after(()=>controller.stop());return {controller,get value(){return local.tms_settings?.[0]?.valor.value}};}
 const a=tab();await a.controller.start({id:'user-a'});await new Promise(resolve=>setImmediate(resolve));assert.equal(a.value,'original');assert.equal(requests.at(-1).p_known_revision,null);
 const b=tab();await b.controller.start({id:'user-a'});assert.equal(requests.at(-1).p_known_revision,4);assert.equal(b.value,'original');
 db.revision=5;db.tables.tms_settings[0].valor.value='updated';const c=tab();await c.controller.start({id:'user-a'});assert.equal(requests.at(-1).p_known_revision,4);assert.equal(c.value,'updated');
 const d=tab();await d.controller.start({id:'user-b'});assert.equal(requests.at(-1).p_known_revision,null);assert.equal(d.value,'updated');
 forbidden=true;const e=tab();await assert.rejects(e.controller.start({id:'user-a'}),/Revoked/);assert.equal(e.value,undefined);assert.equal(e.controller.getState().revision,null);
});
test('an unavailable or slow cache falls back to a normal server snapshot',async t=>{
 let local={},known='unset';const controller=create({getClient:()=>({rpc:async(_,args)=>{known=args.p_known_revision;return {data:{revision:1,tables:{tms_settings:[]}},error:null};}}),capture:()=>local,applySnapshot:async snapshot=>{local=snapshot.tables},snapshotCache:{read:()=>new Promise(()=>{}),write:()=>{}},pollMs:3600000});t.after(()=>controller.stop());const start=Date.now();await controller.start({id:'u'});assert.equal(known,null);assert.ok(Date.now()-start<1500);assert.equal(controller.getState().revision,1);
});
test('main, inventory and showroom caches are separate for the same account',async t=>{
 const disk=new Map(),calls=[];const cache={read:async key=>copy(disk.get(key)||null),write:async(key,value)=>disk.set(key,copy(value))};
 const createTab=(namespace,table)=>{let local={};const controller=create({namespace,snapshotCache:cache,getSnapshotRpc:namespace+'_get',getClient:()=>({rpc:async(name,args)=>{calls.push({name,known:args.p_known_revision});return {data:args.p_known_revision===3?{revision:3,unchanged:true}:{revision:3,tables:{[table]:[{value:namespace}]}},error:null};}}),capture:()=>local,applySnapshot:async snapshot=>{local=copy(snapshot.tables)},pollMs:3600000});t.after(()=>controller.stop());return {controller,get local(){return local}};};
 for(const namespace of ['tms','tms_inventory','tms_showroom']){const first=createTab(namespace,namespace);await first.controller.start({id:'same-user'});await new Promise(resolve=>setImmediate(resolve));}
 for(const namespace of ['tms','tms_inventory','tms_showroom']){const second=createTab(namespace,namespace);await second.controller.start({id:'same-user'});assert.equal(calls.at(-1).known,3);assert.equal(second.local[namespace][0].value,namespace);}
});

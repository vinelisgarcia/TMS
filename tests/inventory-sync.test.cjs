const {test}=require('node:test'),assert=require('node:assert/strict');
const {create}=require('../tms-sync.js');
const copy=x=>JSON.parse(JSON.stringify(x));
function storage(){const m=new Map();return {getItem:k=>m.get(k)||null,setItem:(k,v)=>m.set(k,String(v)),removeItem:k=>m.delete(k),m};}
test('inventory and operating sessions use isolated RPCs, realtime tables and recovery pointers',async t=>{
  const local=storage(),tab=storage(),calls=[],channels=[];
  const client={rpc:async(name,args)=>{calls.push({name,args});if(name==='tms_inventory_get_snapshot')return {data:{revision:7,tables:{tms_inventory_state:[{id:true,payload:{sources:{ct:null,ad:null}}}]}},error:null};if(name==='tms_get_snapshot')return {data:{revision:99,tables:{tms_order_lines:[]}},error:null};assert.equal(name,'tms_inventory_save_snapshot');assert.equal(args.p_expected_revision,7);return {data:{revision:8,changed:true},error:null};},channel:()=>({on(type,filter){channels.push(filter.table);return this;},subscribe(){return this;}}),removeChannel:async()=>{}};
  let inv={},main={};
  const inventory=create({namespace:'tms_inventory',getSnapshotRpc:'tms_inventory_get_snapshot',saveSnapshotRpc:'tms_inventory_save_snapshot',realtimeTable:'tms_inventory_sync_state',storage:local,sessionStorage:tab,getClient:()=>client,capture:()=>copy(inv),applySnapshot:s=>{inv=copy(s.tables);}});
  const operational=create({storage:local,sessionStorage:tab,getClient:()=>client,capture:()=>copy(main),applySnapshot:s=>{main=copy(s.tables);}});
  t.after(()=>{inventory.stop();operational.stop();});await inventory.start({id:'same-user'});await operational.start({id:'same-user'});
  assert.deepEqual(channels,['tms_inventory_sync_state','tms_sync_state']);
  inv.tms_inventory_state[0].payload.sources.ct={materials:[{code:'A',quantity:5}]};inventory.markDirty();
  await operational.refresh();assert.ok([...local.m.keys()].some(k=>k.startsWith('tms_inventory_pending:')));
  assert.ok([...tab.m.keys()].some(k=>k.startsWith('tms_inventory_instance:')));assert.ok([...tab.m.keys()].some(k=>k.startsWith('tms_sync_instance:')));
  await inventory.save();assert.equal(calls.some(call=>call.name==='tms_save_snapshot'),false);
  assert.equal(operational.getState().revision,99);assert.equal(main.tms_inventory_state,undefined);
});

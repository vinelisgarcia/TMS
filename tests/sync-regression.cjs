const {test} = require('node:test');
const assert = require('node:assert/strict');
const {create} = require('../tms-sync.js');
test('two devices cannot overwrite a revision saved by the other',async()=>{
  let remote={revision:1,value:'initial'};
  const a=create(),b=create(); a.loaded(1); b.loaded(1);
  const write=value=>async revision=>{
    if(revision!==remote.revision)throw new Error('TMS_CONFLICT');
    remote={revision:remote.revision+1,value};return remote.revision;
  };
  a.changed();await a.save(write('new'));
  b.changed();await assert.rejects(b.save(write('stale')),/TMS_CONFLICT/);
  assert.equal(remote.value,'new');assert.equal(b.state.dirty,true);assert.equal(b.state.conflict,true);
});
test('edits during an in-flight write are saved in order with the next revision',async()=>{
  const sync=create(); sync.loaded(10);sync.changed();
  let release; const barrier=new Promise(resolve=>release=resolve); const revisions=[];
  const write=async revision=>{revisions.push(revision);if(revision===10)await barrier;return revision+1;};
  const first=sync.save(write);sync.changed();const second=sync.save(write);release();
  await Promise.all([first,second]);
  assert.deepEqual(revisions,[10,11]);assert.equal(sync.state.revision,12);assert.equal(sync.state.dirty,false);
});
test('a failed save remains dirty and retries the same expected revision',async()=>{
  const sync=create(); sync.loaded(3);sync.changed();
  await assert.rejects(sync.save(async()=>{throw new Error('offline');}),/offline/);
  assert.equal(sync.state.revision,3);assert.equal(sync.state.dirty,true);assert.equal(sync.state.saving,false);
  await sync.save(async revision=>{assert.equal(revision,3);return 4;});
  assert.equal(sync.state.dirty,false);assert.equal(sync.state.error,'');
});
test('missing migration or unresolved conflict never dispatches a write',async()=>{
  const sync=create();sync.changed();let called=false;
  await assert.rejects(sync.save(async()=>{called=true;}),/0008/);assert.equal(called,false);
  sync.loaded(1);sync.changed();sync.state.conflict=true;
  await assert.rejects(sync.save(async()=>{called=true;}),/otra sesión/);assert.equal(called,false);
});

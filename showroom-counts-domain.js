(function(root,factory){const C=typeof module==='object'&&module.exports?require('./inventory-counts-domain.js'):root.InventoryCountDomain;const api=factory(C);if(typeof module==='object'&&module.exports)module.exports=api;else root.ShowroomCountDomain=api;})(typeof globalThis==='object'?globalThis:this,function(C){
  'use strict';
  const describe=row=>({...row,referenceSource:row.referenceSource==='CT'?'Físico':row.referenceSource==='AD'?'Virtual':row.referenceSource});
  return {...C,createCycle:(data,report,options={})=>C.createCycle(data,report,{...options,differenceConvention:'physical_virtual'}),itemReport:(cycle,code)=>describe(C.itemReport(cycle,code)),cycleReport:cycle=>C.cycleReport(cycle).map(describe)};
});

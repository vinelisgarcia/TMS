(function(root,factory){const C=typeof module==='object'&&module.exports?require('./inventory-counts-domain.js'):root.InventoryCountDomain;const api=factory(C);if(typeof module==='object'&&module.exports)module.exports=api;else root.ShowroomCountDomain=api;})(typeof globalThis==='object'?globalThis:this,function(C){
  'use strict';
  const describe=row=>({...row,referenceSource:row.referenceSource==='CT'?'Físico':row.referenceSource==='AD'?'Virtual':row.referenceSource});
  const tasks=cycle=>C.tasks(cycle).sort((a,b)=>Number(b.required)-Number(a.required)||(a.expectedLocation||'\uffff').localeCompare(b.expectedLocation||'\uffff')||a.code.localeCompare(b.code));
  return {...C,tasks,createCycle:(data,report,options={})=>C.createCycle(data,report,{...options,differenceConvention:'physical_virtual'}),itemReport:(cycle,code)=>describe(C.itemReport(cycle,code)),cycleReport:cycle=>C.cycleReport(cycle).map(describe)};
});

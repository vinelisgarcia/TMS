(function(root,factory){const api=factory(typeof module==='object'&&module.exports?require('./inventory-domain.js'):root.InventoryDomain);if(typeof module==='object'&&module.exports)module.exports=api;else root.ShowroomDomain=api;})(typeof globalThis==='object'?globalThis:this,function(D){
  'use strict';
  const norm=value=>String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const labels=value=>String(value).replace(/\bAD\b/g,'virtual').replace(/\bCT\b/g,'físico');
  function detectLayout(matrix,source){
    // Add the count sheet's aliases locally; CT/AD imports retain their mapping.
    const mapped=source==='ct'?matrix.map(row=>row.map(value=>({id:'Código',cantidadcontada:'Cantidad',ubicacionshowroom:'Ubicación'})[norm(value)]||value)):matrix;
    const layout=D.detectLayout(mapped,source);
    if(source==='ad')layout.suggestedStockTypes=(layout.stockTypes||[]).filter(value=>norm(value)==='almacen');
    return layout;
  }
  function consolidate(matrix,options){
    const layout=detectLayout(matrix,options.source),mapping=options.mapping||layout.columns;
    const snapshot=D.consolidate(matrix,{...options,mapping,headerRow:options.headerRow??layout.headerRow});
    if(options.source==='ct'){
      const headers=matrix[snapshot.headerRow]||[],responsible=headers.findIndex(value=>norm(value)==='responsable');
      const note=headers.findIndex(value=>norm(value)==='diferencias');
      // This template has an unnamed category next to Ubicacion Showroom.
      const area=mapping.location!==undefined&&!String(headers[mapping.location+1]??'').trim()?mapping.location+1:-1;
      for(const material of snapshot.materials)for(const detail of material.details){
        const row=matrix[detail.row-1];
        detail.responsible=responsible<0?'':String(row[responsible]??'').trim();
        detail.area=area<0?'':String(row[area]??'').trim();
        detail.sourceNote=note<0?'':String(row[note]??'').trim();
      }
      // Traceability changes must also be detected when quantities stay equal.
      snapshot.semanticKey=JSON.stringify({stock:snapshot.semanticKey,detail:snapshot.materials.flatMap(row=>row.details.map(detail=>[row.code,detail.location||'',detail.responsible,detail.area,detail.sourceNote])).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b)))});
      for(const material of snapshot.materials)for(const detail of material.details)if(detail.quantity<0){snapshot.valid=false;snapshot.invalidRows.push({row:detail.row,code:material.code,message:'El conteo físico no admite cantidades negativas.'});}
    }
    snapshot.warnings=snapshot.warnings.map(labels);
    return snapshot;
  }
  function compare(physical,virtual){
    const result=D.compare(physical,virtual);
    result.issues=result.issues.map(labels);
    if(!result.ready)return result;
    result.rows=result.rows.map(row=>{
      const missing=!row.ct;
      return {...row,ctQty:missing?null:row.ctQty,difference:missing||row.difference===null?null:row.difference===0?0:-row.difference,signedImpactDop:missing||row.signedImpactDop===null?null:row.signedImpactDop===0?0:-row.signedImpactDop,impactDop:missing?null:row.impactDop,status:missing?'not_counted':row.status,issues:[...row.issues.map(labels),...(missing?['Sin registro físico: no se supone cantidad cero.']:[])]};
    });
    const differences=result.rows.filter(row=>row.difference!==null&&row.difference!==0),valued=differences.filter(row=>row.impactDop!==null);
    result.summary={...result.summary,differences:differences.length,unvalued:differences.length-valued.length,valued:valued.length,notCounted:result.rows.filter(row=>row.status==='not_counted').length,grossImpactDop:Math.round(valued.reduce((sum,row)=>sum+row.impactDop,0)*100)/100,netImpactDop:Math.round(valued.reduce((sum,row)=>sum+row.signedImpactDop,0)*100)/100};
    if(result.summary.notCounted)result.issues.push(`${result.summary.notCounted} artículos del stock virtual no tienen conteo físico. No se consideran faltantes hasta registrar un conteo, incluido cero cuando corresponda.`);
    return result;
  }
  return {...D,detectLayout,consolidate,compare};
});

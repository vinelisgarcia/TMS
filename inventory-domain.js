(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.InventoryDomain = api;
})(typeof globalThis === 'object' ? globalThis : this, function () {
  'use strict';
  const text = value => String(value ?? '').replace(/\u00a0/g,' ').trim();
  const norm = value => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const normalizeCode = value => { if (typeof value === 'number' && (!Number.isSafeInteger(value) || value < 0)) throw new Error('Código numérico no seguro: usa una columna de texto.'); return text(value).toUpperCase(); };
  function normalizeUnit(value) {
    const unit = norm(value);
    if (!unit || ['piece','pieces','pc','pcs','pz','pza','pzas','unidad','unidades','und','uds','ud','un'].includes(unit)) return 'UND';
    if (['kg','kgs','kilogramo','kilogramos'].includes(unit)) return 'KG';
    if (['box','boxes','caja','cajas'].includes(unit)) return 'CAJA';
    return text(value).toUpperCase();
  }
  function parseNumber(value, decimal = 'auto') {
    if (value === null || value === undefined || text(value) === '') return null;
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) throw new Error('Número no válido.');
      return value;
    }
    let raw = text(value).replace(/\s/g,'');
    if (/^\(.*\)$/.test(raw)) raw = '-' + raw.slice(1,-1);
    raw = raw.replace(/^([+-]?)(?:RD\$|DOP|\$)/i,'$1');
    if (!/^[+-]?\d[\d.,]*(?:[eE][+-]?\d+)?$/.test(raw)) throw new Error('Número no válido: ' + text(value));
    if (/[eE]/.test(raw)) {
      if (raw.includes(',')) throw new Error('Notación científica ambigua.');
    } else if (raw.includes('.') && raw.includes(',')) {
      const mark = decimal === 'comma' ? ',' : decimal === 'dot' ? '.' : raw.lastIndexOf(',') > raw.lastIndexOf('.') ? ',' : '.';
      const grouping = mark === ',' ? '.' : ',';
      const [integer,fraction,...extra] = raw.split(mark);
      if (extra.length || !/^\d+$/.test(fraction) || (integer.includes(grouping) && !new RegExp('^[+-]?\\d{1,3}(\\'+grouping+'\\d{3})+$').test(integer))) throw new Error('Separadores numéricos no válidos.');
      raw = integer.split(grouping).join('') + '.' + fraction;
    } else {
      const mark = raw.includes(',') ? ',' : raw.includes('.') ? '.' : null;
      if (mark) {
        if (decimal === 'auto') {
          if (/^[+-]?\d{1,3}([,.]\d{3})+$/.test(raw)) throw new Error('Número ambiguo: elige separador decimal para ' + raw);
          if (raw.split(mark).length !== 2) throw new Error('Separador numérico ambiguo.');
          raw = raw.replace(',','.');
        } else if ((decimal === 'comma' && mark === ',') || (decimal === 'dot' && mark === '.')) {
          if (raw.split(mark).length !== 2) throw new Error('Separador decimal repetido.');
          raw = raw.replace(',','.');
        } else {
          if (!/^[+-]?\d{1,3}([,.]\d{3})+$/.test(raw)) throw new Error('Agrupación de miles no válida.');
          raw = raw.split(mark).join('');
        }
      }
    }
    const parsed = Number(raw);
    if (!Number.isFinite(parsed)) throw new Error('Número fuera de rango.');
    return parsed;
  }
  function parseCsv(input) {
    const source = String(input).replace(/^\uFEFF/,'');
    const first = source.split(/\r?\n/).slice(0,15).join('\n');
    const counts = {';':0,'\t':0,',':0}; let inQuote=false;
    for(let i=0;i<first.length;i++){if(first[i]==='\"'){if(inQuote&&first[i+1]==='\"')i++;else inQuote=!inQuote;}else if(!inQuote&&Object.prototype.hasOwnProperty.call(counts,first[i]))counts[first[i]]++;}
    const scores = Object.entries(counts).map(([delimiter,score])=>({delimiter,score})).sort((a,b)=>b.score-a.score);
    const delimiter = scores[0].delimiter;
    const rows = []; let row = [], field = '', quoted = false;
    for (let i=0;i<source.length;i++) {
      const char=source[i];
      if (char==='"') {
        if (quoted && source[i+1]==='"') {field+='"';i++;}
        else if (!quoted && field !== '') field+=char;
        else quoted=!quoted;
      } else if (char===delimiter && !quoted) {row.push(field);field='';}
      else if ((char==='\n'||char==='\r') && !quoted) {
        row.push(field);rows.push(row);row=[];field='';
        if(char==='\r'&&source[i+1]==='\n')i++;
      } else field+=char;
    }
    if(quoted)throw new Error('CSV con comillas sin cerrar.');
    if(field!==''||row.length){row.push(field);rows.push(row);}
    return rows;
  }
  const aliases = {
    code:['iddeproducto','idproducto','material','codigo','codigomaterial','codigodelarticulo','numerodearticulo','codigodearticulo','codarticulo','articulo','referencia','itemcode','sku','item'],
    description:['descripciondeproducto','descripcionproducto','description','descripcion','descripcionarticulo','nombrematerial','nombre','itemdescription'],
    quantity:['quantity','stock','existencia','existencias','cantidad','stocktotal','stockactual','cantidadstock','central','stockfisico','unidades'],
    totalValue:['valorcalculado','costototal','valordestock','valordelstock','valortotal','valortotaldop','valortotalstock','valorstock','valorstockdop','valortotalexistencias','importetotal','importe','valoracion','valorinventario','valorinventariodop','valorexistencias','valor'],
    unit:['cantidadunidad','unitofmeasurement','unidad','unidaddemedida','um','uom'],
    warehouse:['warehouse','almacen','codigoalmacen','centro','deposito'],
    location:['ubicacion','location','localizacion','codigoubicacion','bin','stocklocation','ubicaciondelarticulo'],
    batch:['batchnumber','batch','lote','numerolote'],
    stockType:['tipodestock','tipostock','stocktype'],
    currency:['valorcalculadomoneda','moneda','currency'],
    availability:['availability','disponible','disponibilidad','stockdisponible']
  };
  function detectLayout(matrix,source) {
    let best={headerRow:0,columns:{},score:0};
    matrix.slice(0,50).forEach((row,index)=>{
      const columns={};
      Object.entries(aliases).forEach(([name,names])=>{
        const col=row.findIndex(value=>names.includes(norm(value)));
        if(col>=0)columns[name]=col;
      });
      const score=(columns.code!==undefined?4:0)+(columns.quantity!==undefined?4:0)+(source==='ad'&&columns.totalValue!==undefined?3:0)+Object.keys(columns).length;
      if(score>best.score)best={headerRow:index,columns,score};
    });
    const preamble=matrix.slice(0,best.headerRow).flat().map(text).join(' ');
    const date=preamble.match(/(?:date|fecha)[^\d]*(\d{2})[./-](\d{2})[./-](\d{4})/i);
    if(date)best.asOf=`${date[3]}-${date[2]}-${date[1]}`;
    if(!best.asOf) {
      for(const row of matrix.slice(0,best.headerRow)) {
        const label=row.findIndex(value=>['ultimaactualizacion','fechadecorte','fecha','date','lastupdate'].includes(norm(value)));
        if(label>=0&&typeof row[label+1]==='number'&&row[label+1]>20000&&row[label+1]<100000) {
          best.asOf=new Date(Date.UTC(1899,11,30)+Math.floor(row[label+1])*86400000).toISOString().slice(0,10);break;
        }
      }
    }
    if(best.columns.stockType!==undefined) {
      best.stockTypes=[...new Set(matrix.slice(best.headerRow+1).map(row=>text(row[best.columns.stockType])).filter(Boolean))];
      best.suggestedStockTypes=source==='ad'?best.stockTypes.filter(value=>norm(value)==='almacenexterno'):[];
    }
    return best;
  }
  const rounded=value=>Number(value.toFixed(10));
  function consolidate(matrix,options) {
    const source=options.source, mapping=options.mapping||detectLayout(matrix,source).columns;
    const headerRow=Number(options.headerRow??detectLayout(matrix,source).headerRow);
    const quantityColumn=options.quantityBasis==='availability'?mapping.availability:mapping.quantity;
    if(!['ct','ad'].includes(source))throw new Error('Fuente no válida.');
    if(mapping.code===undefined||quantityColumn===undefined)throw new Error('Selecciona las columnas de material y cantidad.');
    if(source==='ad'&&mapping.totalValue===undefined)throw new Error('Selecciona el valor TOTAL de AD en DOP, no el costo unitario.');
    const map=new Map(),invalidRows=[],warnings=[],canonical=[],exact=new Map();
    let acceptedRowCount=0,blankRowCount=0,filteredRows=0;
    for(let index=headerRow+1;index<matrix.length;index++) {
      const row=matrix[index];
      if(!row.some(value=>text(value)!=='')){blankRowCount++;continue;}
      const code=normalizeCode(row[mapping.code]);
      if(!code||/^(total|totales|grandtotal|sumatotal)$/i.test(norm(code))){
        if(!code && row.some(value=>text(value)!==''))invalidRows.push({row:index+1,message:'Falta el código de material.'});
        continue;
      }
      const warehouse=mapping.warehouse===undefined?'':text(row[mapping.warehouse]);
      if(options.warehouses&&options.warehouses.length&&!options.warehouses.includes(warehouse)){filteredRows++;continue;}
      const stockType=mapping.stockType===undefined?'':text(row[mapping.stockType]);
      if(options.stockTypes&&options.stockTypes.length&&!options.stockTypes.includes(stockType)){filteredRows++;continue;}
      try {
        const currency=source==='ad'&&mapping.currency!==undefined?text(row[mapping.currency]).toUpperCase():'';
        if(source==='ad'&&mapping.currency!==undefined&&currency!=='DOP')throw new Error('La valoración debe estar en DOP. Moneda de esta fila: '+(currency||'vacía'));
        const quantity=parseNumber(row[quantityColumn],options.decimal||'auto');
        if(quantity===null)throw new Error('Cantidad vacía.');
        const totalValue=source==='ad'?parseNumber(row[mapping.totalValue],options.decimal||'auto'):null;
        const availability=mapping.availability===undefined?null:parseNumber(row[mapping.availability],options.decimal||'auto');
        const unit=normalizeUnit(mapping.unit===undefined?options.unitDefault||'UND':row[mapping.unit]);
        const batch=mapping.batch===undefined?'':text(row[mapping.batch]);
        const description=mapping.description===undefined?'':text(row[mapping.description]);
        const articleLocation=mapping.location===undefined?'':text(row[mapping.location]);
        const record={code,unit,quantity,totalValue,availability,warehouse,batch};
        if(articleLocation)record.location=articleLocation;
        if(stockType)record.stockType=stockType;
        if(currency)record.currency=currency;
        const key=JSON.stringify(record);exact.set(key,(exact.get(key)||0)+1);canonical.push(key);
        if(!map.has(code))map.set(code,{code,description,quantitiesByUnit:{},valueSum:0,missingValueRows:0,availabilitySum:0,availabilityRows:0,rowCount:0,warehouses:new Set(),batches:new Set(),details:[]});
        const material=map.get(code);
        material.quantitiesByUnit[unit]=rounded((material.quantitiesByUnit[unit]||0)+quantity);
        material.valueSum=rounded(material.valueSum+(totalValue??0));
        if(source==='ad'&&totalValue===null)material.missingValueRows++;
        if(availability!==null){material.availabilitySum=rounded(material.availabilitySum+availability);material.availabilityRows++;}
        material.rowCount++;if(warehouse)material.warehouses.add(warehouse);if(batch)material.batches.add(batch);
        material.details.push({row:index+1,quantity,totalValueDop:totalValue,unit,warehouse,batch,...(articleLocation?{location:articleLocation}:{}),...(stockType?{stockType}:{}),...(currency?{currency}:{})});
        acceptedRowCount++;
      } catch(error){invalidRows.push({row:index+1,code,message:error.message});}
    }
    if(!acceptedRowCount)throw new Error('No se encontraron filas de inventario válidas. Revisa los encabezados y las columnas.');
    const materials=[...map.values()].sort((a,b)=>a.code.localeCompare(b.code)).map(material=>{
      const units=Object.keys(material.quantitiesByUnit),unitConflict=units.length!==1;
      return {code:material.code,description:material.description,unit:unitConflict?units.join(' / '):units[0],unitConflict,quantitiesByUnit:material.quantitiesByUnit,quantity:unitConflict?null:material.quantitiesByUnit[units[0]],totalValueDop:source==='ad'&&material.missingValueRows===0?material.valueSum:null,knownValueDop:source==='ad'?material.valueSum:null,missingValueRows:material.missingValueRows,availability:!unitConflict&&material.availabilityRows===material.rowCount?material.availabilitySum:null,rowCount:material.rowCount,warehouses:[...material.warehouses].sort(),batches:[...material.batches].sort(),details:material.details};
    });
    const duplicateExtraRows=[...exact.values()].reduce((sum,count)=>sum+Math.max(0,count-1),0);
    if(duplicateExtraRows)warnings.push(`${duplicateExtraRows} filas idénticas se sumaron. Revisa si son stock real dividido o una repetición del reporte.`);
    if(materials.some(row=>row.unitConflict))warnings.push('Hay materiales con unidades incompatibles. No se suman como si fueran equivalentes.');
    if(source==='ad'&&materials.some(row=>row.missingValueRows))warnings.push('Hay materiales sin valor AD completo. Su diferencia económica quedará sin valorar.');
    if(source==='ad'&&mapping.stockType!==undefined&&(!options.stockTypes||!options.stockTypes.length))warnings.push('Se incluyen todos los tipos de stock del reporte. Revisa que En tránsito no se mezcle con el stock físico en almacén.');
    if(mapping.unit===undefined)warnings.push('El archivo no indica unidad: se usa UND. Confirma que las cantidades son unidades equivalentes.');
    const byUnit={};materials.forEach(row=>Object.entries(row.quantitiesByUnit).forEach(([unit,qty])=>byUnit[unit]=rounded((byUnit[unit]||0)+qty)));
    const knownValueDop=source==='ad'?rounded(materials.reduce((sum,row)=>sum+row.knownValueDop,0)):null;
    canonical.sort();
    const semanticKey=JSON.stringify({source,asOf:options.asOf||'',quantityBasis:options.quantityBasis||'quantity',rows:canonical});
    return {source,fileName:options.fileName||'',fileHash:options.fileHash||'',asOf:options.asOf||'',importedAt:new Date().toISOString(),rawRowCount:acceptedRowCount+invalidRows.length+filteredRows,acceptedRowCount,blankRowCount,filteredRows,repeatedMaterials:materials.filter(row=>row.rowCount>1).length,duplicateExtraRows,materials,totals:{materials:materials.length,quantity:Object.keys(byUnit).length===1?Object.values(byUnit)[0]:null,byUnit,availability:materials.every(row=>row.availability!==null)?rounded(materials.reduce((sum,row)=>sum+row.availability,0)):null,totalValueDop:source==='ad'&&materials.every(row=>row.totalValueDop!==null)?knownValueDop:null,knownValueDop,missingValueMaterials:source==='ad'?materials.filter(row=>row.totalValueDop===null).length:0},stockTypes:options.stockTypes||[],quantityBasis:options.quantityBasis||'quantity',unitDefault:options.unitDefault||'UND',mapping:{...mapping},headerRow,decimal:options.decimal||'auto',warnings,invalidRows,valid:invalidRows.length===0,semanticKey};
  }
  function compare(ct,ad) {
    if(ct?.valid===false||ad?.valid===false)return {ready:false,rows:[],summary:{},issues:['Hay filas no válidas en los reportes. Corrige la carga antes de comparar.']};
    if(!ct||!ad)return {ready:false,rows:[],summary:{},issues:['Carga los reportes CT y AD para calcular diferencias.']};
    const left=new Map(ct.materials.map(row=>[row.code,row])),right=new Map(ad.materials.map(row=>[row.code,row]));
    const rows=[...new Set([...left.keys(),...right.keys()])].map(code=>{
      const c=left.get(code),a=right.get(code),issues=[];
      const unitConflict=!!(c?.unitConflict||a?.unitConflict||(c&&a&&c.unit!==a.unit));
      const ctQty=c?c.quantity:0,adQty=a?a.quantity:0;
      const difference=unitConflict?null:rounded(ctQty-adQty);
      const unitCostDop=a&&!unitConflict&&adQty>0&&a.totalValueDop!==null&&a.totalValueDop>=0?a.totalValueDop/adQty:null;
      if(unitConflict)issues.push('Unidades incompatibles: requiere equivalencia antes de comparar.');
      if(!a)issues.push('No registrado en AD; no hay valoración disponible.');
      if(a&&adQty<=0)issues.push('Cantidad AD no positiva: no se deriva costo unitario.');
      if(a&&a.totalValueDop===null)issues.push('Falta parte de la valoración AD.');
      if(a&&a.totalValueDop<0)issues.push('Valor AD negativo: requiere revisión.');
      const rawImpact=difference===0?0:(difference!==null&&unitCostDop!==null?Math.abs(difference)*unitCostDop:null);
      const impactDop=rawImpact===null?null:Math.round((rawImpact+Number.EPSILON)*100)/100;
      return {key:code,code,description:a?.description||c?.description||'',unit:unitConflict?`${c?.unit||'—'} / ${a?.unit||'—'}`:c?.unit||a?.unit,ctQty,adQty,difference,unitCostDop,impactDop,signedImpactDop:impactDop===null?null:difference===0?0:(difference<0?-impactDop:impactDop),status:unitConflict?'unit_conflict':!a?'only_ct':!c?'only_ad':difference===0?'equal':'difference',issues,ct:c||null,ad:a||null};
    }).sort((a,b)=>(a.status==='equal'?1:0)-(b.status==='equal'?1:0)||(b.impactDop??-1)-(a.impactDop??-1)||Math.abs(b.difference??0)-Math.abs(a.difference??0)||a.code.localeCompare(b.code));
    const different=rows.filter(row=>row.difference!==0),valued=different.filter(row=>row.impactDop!==null);
    return {ready:true,rows,issues:ct.asOf!==ad.asOf?['Las fechas de corte CT y AD no coinciden. Las diferencias pueden incluir movimientos entre ambas fechas.']:[],summary:{materials:rows.length,equal:rows.filter(row=>row.status==='equal').length,differences:different.length,onlyCT:rows.filter(row=>row.status==='only_ct').length,onlyAD:rows.filter(row=>row.status==='only_ad').length,unitConflicts:rows.filter(row=>row.status==='unit_conflict').length,unvalued:different.filter(row=>row.impactDop===null).length,valued:valued.length,grossImpactDop:Math.round(valued.reduce((sum,row)=>sum+row.impactDop,0)*100)/100,netImpactDop:Math.round(valued.reduce((sum,row)=>sum+row.signedImpactDop,0)*100)/100,ctQuantity:ct.totals.quantity,adQuantity:ad.totals.quantity,adValueDop:ad.totals.totalValueDop}};
  }
  return {parseCsv,parseNumber,normalizeCode,normalizeUnit,detectLayout,consolidate,compare};
});

const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs');
const C=require('../inventory-counts-domain.js'),D=require('../inventory-domain.js'),Daily=require('../inventory-daily-domain.js'),Excel=require('../inventory-daily-excel.js'),ExcelJS=require('../vendor/exceljs-4.4.0.min.js');
function fixture(){const ct=D.consolidate([['Material','Quantity','Location'],['A',2,'A01'],['A',3,'A02'],['B',2,'']],{source:'ct',asOf:'2026-11-16'}),ad=D.consolidate([['Codigo','Stock','Valor total DOP'],['A',9,123456789],['AD_ONLY',2,40]],{source:'ad',asOf:'2026-11-16'});ct.semanticHash='CT';ad.semanticHash='AD';const data={sources:{ct,ad}};C.importMaster(data,[['CODE','DESCRIPTION'],['0007','Referencia con ceros']],{code:0,description:1},'articles');return data;}
function planned(kind='ctad'){const data=fixture();return Daily.plan(data,D.compare(data.sources.ct,data.sources.ad),{kind,day:'2026-11-16',templateId:'TEST-'+kind,user:'planner@example.invalid',now:'2026-11-16T12:00:00Z'});}
function matrix(template,entries){return [...Array.from({length:7},()=>[]),...template.rows.map(item=>{const entry=entries[item.code+'|'+item.expectedLocation]||{};return [entry.day||template.day,item.code,item.description,item.expectedLocation,item.expectedLocation,item.unit,entry.quantity??null,entry.operator||'Equipo CT',entry.note||'',item.id];})];}
test('blind plans include the complete union and locations, freeze the counting reference and keep top priorities mandatory',()=>{
 const {data,template}=planned(),cycle=C.active(data);assert.deepEqual([...new Set(template.rows.map(row=>row.code))],['0007','A','AD_ONLY','B']);assert.equal(template.rows.filter(row=>row.code==='A').length,2);assert.equal(cycle.items.A.stockCT,5);assert.equal(cycle.items.A.required,true);assert.equal(cycle.records.length,0);assert.equal(template.rows.find(row=>row.code==='0007').code,'0007');assert.doesNotMatch(JSON.stringify(template),/123456789|stockCT|stockAD|availability|costAD|masterStock/);
});
test('daily upload counts zero, skips blank cells, preserves the operator/count day versus uploader/upload time, and is idempotent',()=>{
 const {data,template}=planned(),parsed=Daily.parse(template,matrix(template,{'A|A01':{quantity:0,day:'15/11/2026'},'B|':{quantity:3,day:'2026-11-15'}}));assert.equal(parsed.rows.length,2);assert.equal(parsed.blank,3);
 const options={kind:'ctad',user:'uploader@example.invalid',fileName:'equipo.xlsx',now:'2026-11-16T16:00:00Z'};assert.equal(Daily.apply(data,template,parsed.rows,options).inserted,2);
 const record=C.active(data).records[0];assert.equal(record.quantity,0);assert.equal(record.countedOn,'2026-11-15');assert.equal(record.countedBy,'Equipo CT');assert.equal(record.by,'uploader@example.invalid');assert.equal(record.at,options.now);
 assert.equal(Daily.apply(data,template,parsed.rows,options).skipped,2);assert.equal(C.active(data).records.length,2);assert.equal(Daily.history(data,'2026-11-15').length,2);
});
test('a later day replaces the reference total at that location while preserving prior daily history',()=>{
 const {data,template}=planned();for(const [day,quantity]of [['2026-11-15',4],['2026-11-16',6]])Daily.apply(data,template,Daily.parse(template,matrix(template,{'B|':{day,quantity}})).rows,{kind:'ctad'});
 assert.equal(C.itemReport(C.active(data),'B').quantity,6);assert.equal(C.active(data).records.length,2);assert.equal(Daily.history(data,'2026-11-15')[0].quantity,4);assert.equal(Daily.history(data,'2026-11-15')[0].current,false);assert.equal(Daily.history(data,'2026-11-16')[0].current,true);
});
test('count dates use Santo Domingo, preserve date-only inputs, and reject a cleared date without writing a record',()=>{
 assert.equal(C.countDay('2026-11-16T03:30:00Z'),'2026-11-15');assert.equal(C.countDay('2026-11-16'),'2026-11-16');const {data}=planned(),before=JSON.stringify(data);
 assert.throws(()=>C.recordCount(data,{code:'B',expectedLocation:'',actualLocation:'',quantity:1,confirmed:true,countedOn:''}),/fecha de conteo válida/);assert.equal(JSON.stringify(data),before);
});
test('wrong module/cycle, invalid dates, missing operators, edited references, duplicate rows and changed reuploads fail atomically',()=>{
 const {data,template}=planned(),valid=matrix(template,{'B|':{quantity:2}}),row=valid.find(values=>values[1]==='B');
 for(const [column,value]of [[0,'2026-02-29'],[7,''],[1,'OTHER'],[6,-1]]){const invalid=structuredClone(valid);invalid.find(values=>values[1]==='B')[column]=value;assert.throws(()=>Daily.parse(template,invalid));}
 assert.throws(()=>Daily.parse(template,[...valid,row]));assert.throws(()=>Daily.parse(template,valid,{formulaRows:[valid.indexOf(row)+1]}));
 const parsed=Daily.parse(template,valid),before=JSON.stringify(data);assert.throws(()=>Daily.apply(data,template,parsed.rows,{kind:'showroom'}),/otro inventario/);assert.equal(JSON.stringify(data),before);
 Daily.apply(data,template,parsed.rows,{kind:'ctad'});const prior=JSON.stringify(data);assert.throws(()=>Daily.apply(data,template,[{...parsed.rows[0],quantity:99}],{kind:'ctad'}),/ya fue cargado/);assert.equal(JSON.stringify(data),prior);
 C.active(data).status='archived';assert.throws(()=>Daily.apply(data,template,parsed.rows,{kind:'ctad'}),/ciclo cerrado/);
});
test('showroom count uploads use an isolated cycle convention and complete guided tasks',()=>{
 const {data,template}=planned('showroom');assert.equal(template.rows.filter(row=>row.code==='A').length,1);const rows=Daily.parse(template,[['Código del material','Cantidad contada'],['A',5],['0007',0],['B',null]]).rows;Daily.apply(data,template,rows,{kind:'showroom',user:'carga@example.invalid'});assert.equal(C.active(data).differenceConvention,'physical_virtual');assert.equal(C.itemReport(C.active(data),'A').complete,true);assert.equal(C.itemReport(C.active(data),'A').quantity,5);assert.equal(C.itemReport(C.active(data),'A').records[0].actualLocation,'');assert.equal(C.itemReport(C.active(data),'A').records[0].countedBy,'carga@example.invalid');assert.equal(Daily.apply(data,template,rows,{kind:'showroom',user:'carga@example.invalid'}).skipped,2);assert.throws(()=>Daily.parse(template,[['Código','Cantidad'],['A',1],['A',2]]));assert.throws(()=>Daily.parse(template,[['Código','Cantidad'],['OTHER',1]]));assert.throws(()=>Daily.parse(template,[['Código','Cantidad'],['A',-1]]));
});
test('styled Excel is blind even in metadata, keeps identifiers as text, physical inputs blank, validation and frozen headers',async()=>{
 for(const kind of ['ctad','showroom']){const {template}=planned(kind),book=Excel.createWorkbook(ExcelJS,template),bytes=await book.xlsx.writeBuffer(),saved=await new ExcelJS.Workbook().xlsx.load(bytes),sheet=saved.getWorksheet('Conteo');
  assert.equal(saved.getWorksheet('_TMS').state,'veryHidden');if(kind==='showroom'){assert.deepEqual(sheet.getRow(1).values.slice(1),['Código del material','Cantidad contada']);assert.equal(sheet.columnCount,2);assert.equal(sheet.getCell('A2').value,'0007');assert.equal(sheet.getCell('B2').value,null);assert.equal(sheet.getCell('B2').dataValidation.formulae[0],0);assert.equal(sheet.views[0].ySplit,1);continue;}assert.equal(sheet.getCell('A7').fill.fgColor.argb,'FF183F66');assert.equal(sheet.getCell('G8').fill.fgColor.argb,'FFFFF4D5');assert.equal(sheet.getCell('G8').value,null);assert.equal(sheet.getCell('G8').dataValidation.formulae[0],0);assert.equal(sheet.getCell('B8').value,'0007');assert.equal(sheet.getCell('B8').type,3);assert.equal(sheet.views[0].ySplit,7);assert.match(String(sheet.autoFilter),/:J\d+$/);assert.ok(sheet.getCell('A8').value instanceof Date);assert.doesNotMatch(JSON.stringify(saved.model),/123456789|stockCT|stockAD|costAD|masterStock/);
  if(process.env.TMS_DAILY_SAMPLE_DIR){fs.mkdirSync(process.env.TMS_DAILY_SAMPLE_DIR,{recursive:true});fs.writeFileSync(process.env.TMS_DAILY_SAMPLE_DIR+'/'+kind+'.xlsx',Buffer.from(bytes));}
 }
});

test('showroom totals do not double count older per-location records and later location counts are not mixed with totals',()=>{
 const {data,template}=planned('showroom');let cycle=C.active(data);
 C.recordCount(data,{code:'A',expectedLocation:'A01',actualLocation:'A01',quantity:2,confirmed:true,countedOn:'2026-11-15'},{user:'test',now:'2026-11-15T10:00:00Z'});
 C.recordCount(data,{code:'A',expectedLocation:'A02',actualLocation:'A02',quantity:3,confirmed:true,countedOn:'2026-11-15'},{user:'test',now:'2026-11-15T10:00:01Z'});
 const rows=Daily.parse(template,[['Código','Cantidad'],['A',8]]).rows;Daily.apply(data,template,rows,{kind:'showroom',user:'test',now:'2026-11-16T12:00:00Z'});
 cycle=C.active(data);assert.equal(C.itemReport(cycle,'A').quantity,8);assert.equal(C.itemReport(cycle,'A').complete,true);assert.ok(C.tasks(cycle).filter(row=>row.code==='A').every(row=>!row.complete));
 C.recordCount(data,{code:'A',expectedLocation:'A01',actualLocation:'A01',quantity:4,confirmed:true,countedOn:'2026-11-17'},{user:'test',now:'2026-11-17T12:00:00Z'});
 assert.equal(C.itemReport(cycle,'A').quantity,4);assert.equal(C.itemReport(cycle,'A').complete,false);
});
test('previous showroom templates remain valid for location based uploads',()=>{
 const {data,template}=planned('showroom');template.version=1;const row=template.rows.find(item=>item.code==='A');const rows=Daily.parse(template,matrix(template,{['A|'+row.expectedLocation]:{quantity:3}})).rows;assert.equal(rows.length,1);Daily.apply(data,template,rows,{kind:'showroom',user:'test'});assert.equal(C.itemReport(C.active(data),'A').quantity,3);assert.equal(C.itemReport(C.active(data),'A').complete,false);
});

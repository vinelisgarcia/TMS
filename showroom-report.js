(function(root,factory){
  const api=factory(root);
  if(typeof module==='object'&&module.exports)module.exports=api;else root.ShowroomExecutiveExcel=api;
})(typeof globalThis==='object'?globalThis:this,function(root){
  'use strict';
  const ink='FF183F66',muted='FF526678',green='FF18744D',red='FFB73838',light='FFF2F5F9';
  const argb=color=>color.length===6?'FF'+color:color;
  const currency='"RD$" #,##0.00',headers=['Referencia','Descripción','Inventario','Valor económico Virtual','Conteo físico','Diferencia'];
  const date=value=>/^\d{4}-\d{2}-\d{2}$/.test(value||'')?value.slice(8,10)+'/'+value.slice(5,7)+'/'+value.slice(0,4):String(value||'Sin fecha');
  let libraryPromise=null;
  function load(){
    if(root.ExcelJS)return Promise.resolve(root.ExcelJS);
    if(libraryPromise)return libraryPromise;
    libraryPromise=new Promise((resolve,reject)=>{
      const script=root.document.createElement('script');script.src='./vendor/exceljs-4.4.0.min.js';
      const fail=()=>{clearTimeout(timer);script.onload=null;script.onerror=null;script.remove();reject(Error('No se pudo cargar el generador del informe. Revisa la conexión y vuelve a exportar.'));};
      const timer=setTimeout(fail,20000);
      script.onerror=fail;script.onload=()=>{if(!root.ExcelJS){fail();return;}clearTimeout(timer);resolve(root.ExcelJS);};root.document.head.appendChild(script);
    }).catch(error=>{libraryPromise=null;throw error;});
    return libraryPromise;
  }
  function createWorkbook(ExcelJS,report){
    const rows=report.rows,book=new ExcelJS.Workbook();
    book.creator='TMS · Álvarez Dominicana';book.title='Informe ejecutivo · Inventario del showroom';book.created=new Date(report.at);
    const sheet=book.addWorksheet('Conciliacion',{views:[{state:'frozen',ySplit:12,showGridLines:false}],pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.3,right:.3,top:.4,bottom:.4,header:.15,footer:.2}},properties:{tabColor:{argb:ink}}});
    sheet.columns=[18,48,16,27,18,16].map(width=>({width}));sheet.properties.defaultRowHeight=23;
    const fill=color=>({type:'pattern',pattern:'solid',fgColor:{argb:argb(color)}}),font={name:'Arial',size:11,color:{argb:ink}};
    function block(row,first,last,value,style={}){
      sheet.mergeCells(row,first,row,last);const cell=sheet.getCell(row,first);cell.value=value;cell.font={...font,...style};cell.alignment={vertical:'middle',wrapText:true};return cell;
    }
    function line(row,value,color=muted,bg){const cell=block(row,1,6,value,{color:{argb:argb(color)}});if(bg)cell.fill=fill(bg);sheet.getRow(row).height=Math.max(24,Math.ceil(String(value).length/145)*16+8);return cell;}
    sheet.getRow(1).height=8;
    block(2,1,6,'Inventario del showroom',{size:17,bold:true});sheet.getRow(2).height=32;
    const generated=new Date(report.at).toLocaleString('es-DO',{timeZone:'America/Santo_Domingo',dateStyle:'short',timeStyle:'short'});
    line(3,'ÁLVAREZ DOMINICANA · INFORME EJECUTIVO   |   Generado: '+generated);
    line(4,'Corte virtual: '+date(report.virtualAsOf)+'   ·   Corte físico: '+date(report.physicalAsOf));
    sheet.getRow(4).getCell(1).border={bottom:{style:'thin',color:{argb:'FFB9CAD9'}}};
    line(5,'Alcance: '+report.scope+' · '+rows.length+' de '+report.totalRows+' referencias'+(report.search?' · Búsqueda: '+report.search:''));
    sheet.getRow(6).height=9;
    const valued=rows.filter(row=>Number.isFinite(row.ad?.totalValueDop)),missingValue=rows.length-valued.length;
    const totalValue=valued.length?Math.round(valued.reduce((sum,row)=>sum+row.ad.totalValueDop,0)*100)/100:null;
    const surplus=rows.filter(row=>row.difference>0).length,shortage=rows.filter(row=>row.difference<0).length;
    const notCounted=rows.filter(row=>row.status==='not_counted').length,conflicts=rows.filter(row=>row.status==='unit_conflict').length;
    [['REFERENCIAS DEL INFORME',rows.length],[(missingValue?'VALOR VIRTUAL CONOCIDO':'VALOR VIRTUAL TOTAL')+' · DOP',totalValue],['CON DIFERENCIAS',surplus+shortage]].forEach(([label,value],i)=>{
      const col=i*2+1,labelCell=block(7,col,col+1,label,{size:10,bold:true,color:{argb:muted}}),valueCell=block(8,col,col+1,value===null?'Sin valoración':value,{size:16,bold:true});
      for(const cell of [labelCell,valueCell]){cell.fill=fill(light);cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};}
      valueCell.numFmt=i===1?currency:'#,##0';
    });
    sheet.getRow(7).height=26;sheet.getRow(8).height=38;
    line(9,'Sobrantes: '+surplus+'   ·   Faltantes: '+shortage+'   ·   Coinciden: '+rows.filter(row=>row.difference===0).length+'   ·   Sin conteo: '+notCounted+'   ·   Sin valor virtual: '+missingValue+(conflicts?'   ·   Unidades incompatibles: '+conflicts:''));
    const notes=['Inventario = cantidad virtual · Valor = total virtual en DOP · Diferencia = físico − virtual.'];
    if(report.virtualAsOf!==report.physicalAsOf)notes.push('Los reportes tienen fechas de corte distintas.');
    if(missingValue)notes.push('El importe mostrado suma solo las referencias con valor virtual conocido.');
    if(notCounted||conflicts)notes.push('Las diferencias no calculables quedan en blanco.');
    const note=line(10,notes.join('\n'),missingValue||notCounted||conflicts||report.virtualAsOf!==report.physicalAsOf?'805D20':muted,missingValue||notCounted||conflicts||report.virtualAsOf!==report.physicalAsOf?'FFF7E8':undefined);
    note.font={...font,size:10,color:note.font.color};sheet.getRow(10).height=notes.length*16+10;sheet.getRow(11).height=9;
    sheet.getRow(12).values=headers;sheet.getRow(12).height=34;
    sheet.getRow(12).eachCell(cell=>{cell.font={...font,bold:true,color:{argb:'FFFFFFFF'}};cell.fill=fill(ink);cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};cell.border={right:{style:'thin',color:{argb:'FFFFFFFF'}}};});
    rows.forEach((item,index)=>{
      const row=sheet.addRow([String(item.code),item.description,item.adQty,item.ad?.totalValueDop??null,item.ctQty,item.difference]);
      row.height=Math.max(30,Math.ceil(String(item.description||'').length/43)*15+8);
      row.eachCell({includeEmpty:true},(cell,col)=>{cell.font={...font};cell.alignment={horizontal:col<=2?'left':'right',vertical:'middle',wrapText:col<=2};if(index%2===0)cell.fill=fill('F6F8FB');
        cell.numFmt=col<=2?'@':col===4?currency:col===6?(Number.isInteger(cell.value)?'+#,##0;-#,##0;0':'+#,##0.###;-#,##0.###;0'):Number.isInteger(cell.value)?'#,##0':'#,##0.###';
      });
    });
    const last=12+rows.length;sheet.autoFilter={from:'A12',to:'F'+last};
    if(rows.length)sheet.addConditionalFormatting({ref:'F13:F'+last,rules:[
      {type:'cellIs',operator:'greaterThan',formulae:[0],style:{font:{color:{argb:green},bold:true},fill:fill('E8F4ED')}},
      {type:'cellIs',operator:'lessThan',formulae:[0],style:{font:{color:{argb:red},bold:true},fill:fill('FCECEC')}}
    ]});
    sheet.pageSetup.printTitlesRow='12:12';sheet.pageSetup.printArea='A1:F'+last;
    sheet.headerFooter.oddFooter='&LÁlvarez Dominicana · Showroom&R Página &P de &N';
    return book;
  }
  return {load,createWorkbook};
});

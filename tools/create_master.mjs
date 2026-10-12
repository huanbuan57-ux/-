import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const workbook=Workbook.create();
const sheet=workbook.worksheets.add('产品资料');
sheet.showGridLines=false;
sheet.getRange('A1:K3').values=[
['标准图号','是否负责','是否启用','中文名称','英文名称','包装模数','单箱数量','纸箱尺寸','工位号','重量说明','资料来源'],
['S111900000182','是','是','高位进气管总成','High level intake pipe assembly',24,2,'1120*250*240','/','以实际称重为准','技术通知.doc 人工完成样例'],
['S111900000294','是','是','谐振腔总成','Resonant cavity assembly',24,24,'600*400*200','/','以实际称重为准','技术通知.doc 人工完成样例']];
sheet.getRange('A1:K3').format.font={name:'Microsoft YaHei',size:11};
sheet.getRange('A1:K1').format={fill:'#244A68',font:{name:'Microsoft YaHei',bold:true,color:'#FFFFFF'},rowHeight:30};
sheet.getRange('A2:K3').format.rowHeight=50;
sheet.getRange('A1:K3').format.wrapText=true;
const widths=[23,12,12,25,43,14,14,24,14,26,32];
'ABCDEFGHIJK'.split('').forEach((col,i)=>sheet.getRange(col+'1:'+col+'3').format.columnWidth=widths[i]);
sheet.getRange('B2:C500').dataValidation={rule:{type:'list',values:['是','否']}};
sheet.getRange('F2:G3').setNumberFormat('0');
const notes=workbook.worksheets.add('维护说明'); notes.getRange('A1').values=[['工位号仅采用批次 Excel 明确值，缺失时使用 /；此处不回填。包装数据以 Excel 优先、资料表回退。']];
notes.getRange('A2').values=[['新增图号后填写是否负责、是否启用、英文名称、包装模数、单箱数量、纸箱尺寸和重量说明。']];
workbook.recalculate();
console.log((await workbook.inspect({kind:'table',range:'产品资料!A1:K3',tableMaxRows:3,tableMaxCols:11})).ndjson);
const image=await workbook.render({sheetName:'产品资料',range:'A1:K3',scale:1,format:'png'});
await fs.writeFile('.word-qa/master.png',new Uint8Array(await image.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(workbook)).save('data/产品资料表.xlsx');

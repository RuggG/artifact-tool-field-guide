import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';

const here=path.dirname(fileURLToPath(import.meta.url));
const out=path.join(here,'overview-probe');
await fs.mkdir(out,{recursive:true});
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(here,'man-group-run/source.xlsx')));
const report=[];
const calls=[
 ['full-with-preview-caps',{kind:'workbook,sheet,table',maxChars:150000,tableMaxRows:4,tableMaxCols:6,tableMaxCellChars:80}],
 ['fee-table-tiny-preview',{kind:'table',sheetId:'Fee engine',maxChars:150000,tableMaxRows:1,tableMaxCols:1,tableMaxCellChars:1}],
 ['fee-table-small-budget',{kind:'table',sheetId:'Fee engine',maxChars:1000,tableMaxRows:1,tableMaxCols:1,tableMaxCellChars:1}],
 ['fee-table-explicit-range',{kind:'table',include:'values',sheetId:'Fee engine',range:'A1:W109',maxChars:1000,tableMaxRows:1,tableMaxCols:1,tableMaxCellChars:1}],
 ['metadata-overview',{kind:'workbook,sheet',maxChars:14000}],
 ['fee-regions-small-preview',{kind:'region',sheetId:'Fee engine',range:'A1:W109',maxChars:40000,tableMaxRows:2,tableMaxCols:3,tableMaxCellChars:20}],
 ['fee-regions-larger-preview',{kind:'region',sheetId:'Fee engine',range:'A1:W109',maxChars:40000,tableMaxRows:4,tableMaxCols:6,tableMaxCellChars:20}]
];
for(const [id,input] of calls){
 const r=await wb.inspect(input);
 const result={ndjson:r.ndjson,recordCount:r.recordCount,truncated:r.truncated,metadata:r.metadata};
 const parsed=r.ndjson.split('\n').filter(Boolean).map(x=>JSON.parse(x));
 const shapes=parsed.map(x=>{
  const base={kind:x.kind,sheet:x.sheet||x.name,address:x.address,rows:x.rows,cols:x.cols};
  if(x.values){
   const actual=wb.worksheets.getItem(x.sheet).getRange(x.address).values;
   Object.assign(base,{returnedRows:x.values.length,returnedCols:Math.max(...x.values.map(a=>a.length)),exactlyMatchesDirectValues:JSON.stringify(actual)===JSON.stringify(x.values)});
  }
  if(x.preview)Object.assign(base,{previewAddress:x.previewAddress,previewRows:x.preview.length,previewCols:Math.max(...x.preview.map(a=>a.length)),longestPreviewText:Math.max(0,...x.preview.flat().filter(v=>typeof v==='string').map(v=>v.length))});
  if(x.message)base.message=x.message;
  return base;
 });
 await fs.writeFile(path.join(out,id+'.json'),JSON.stringify({input,result},null,2));
 const summary={id,characters:r.ndjson.length,recordCount:r.recordCount,truncated:r.truncated,records:shapes};
 report.push(summary);console.log(JSON.stringify(summary));
}
const help=wb.help('workbook.inspect',{include:'index,notes,examples',maxChars:6000});
await fs.writeFile(path.join(out,'inspect-help.txt'),help.ndjson);
await fs.writeFile(path.join(out,'report.json'),JSON.stringify(report,null,2));

// Read-only inspection comparisons. Run with the bundled Node runtime.
// No edits, recalculation or export: the original imported representation is the fixture.
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import {fileURLToPath} from 'node:url';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';
const home=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const source=process.argv[2] || path.join(home,'files/inspection/source.xlsx');
const out=path.join(home,'files/inspection');
await fs.mkdir(out,{recursive:true});
const bytes=await fs.readFile(source),hash=b=>crypto.createHash('sha256').update(b).digest('hex');
if(path.resolve(source)!==path.join(out,'source.xlsx'))await fs.writeFile(path.join(out,'source.xlsx'),bytes);
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const baseline=[];
for(const s of wb.worksheets.items){const r=s.getUsedRange();baseline.push({sheet:s.name,address:r.address,values:r.values,formulas:r.formulas,nativeTables:s.tables.items.length});}
const scopes={workbook:{},sheet:{sheetId:'Fee engine'},range:{sheetId:'Fee engine',range:'C4:H12'}};
const cases=[];
function add(id,group,label,input,extra={}){cases.push({id,group,label,input,...extra});}
for(const [mode,kind,label] of [['map','workbook,sheet','Workbook + sheet records'],['table','table','Value grids'],['combined','workbook,sheet,table','Structure + value grids']]){
 for(const scope of ['workbook','sheet','range'])for(const maxChars of [150000,1000]){
  add(`${mode}-${scope}-${maxChars}`,'core',label,{kind,...scopes[scope],maxChars,tableMaxRows:1,tableMaxCols:1,tableMaxCellChars:1},{mode,scope,budget:maxChars,preview:'1 × 1'});
 }
}
add('combined-workbook-14000','core','Structure + value grids',{kind:'workbook,sheet,table',maxChars:14000,tableMaxRows:1,tableMaxCols:1,tableMaxCellChars:1},{mode:'combined',scope:'workbook',budget:14000,preview:'1 × 1'});
for(const maxChars of [150000,1000])add(`table-sheet-wide-${maxChars}`,'core','Value grids',{kind:'table',sheetId:'Fee engine',maxChars,tableMaxRows:4,tableMaxCols:6,tableMaxCellChars:20},{mode:'table',scope:'sheet',budget:maxChars,preview:'4 × 6'});
for(const scope of ['sheet','range'])for(const [rows,cols] of [[2,3],[4,6]])add(`regions-${scope}-${rows}`,'regions','Detected blocks with previews',{kind:'region',...scopes[scope],range:scopes[scope].range||'A1:W109',maxChars:40000,tableMaxRows:rows,tableMaxCols:cols,tableMaxCellChars:20},{mode:'region',scope,budget:40000,preview:`${rows} × ${cols}`});
for(const scope of ['sheet','range'])for(const limit of [3,300])add(`formulas-${scope}-${limit}`,'formulas',`Formula records · up to ${limit}`,{kind:'formula',...scopes[scope],maxChars:150000,options:{maxResults:limit}},{scope,limit});
add('sheet-only','more','Just the sheet inventory',{kind:'sheet',maxChars:150000});
add('table-formula','more','Values + formula records in one response',{kind:'table,formula',sheetId:'Fee engine',range:'C4:H12',maxChars:150000,options:{maxResults:300}});
add('project-values','more','Select values and formulas fields',{kind:'table',sheetId:'Fee engine',range:'C4:H12',include:'values,formulas',maxChars:150000});
add('exclude-values','more','Describe the grid without its values',{kind:'table',sheetId:'Fee engine',exclude:'values',maxChars:150000});
add('record-search','more','Filter records using search',{kind:'sheet',search:'Fee engine',maxChars:150000});
add('chart-token','more','Check an unsupported kind token',{kind:'sheet,table,formula,chart',sheetId:'Fee engine',range:'C4:H12',maxChars:150000,options:{maxResults:300}});
add('objects','more','Inspect drawing and name records',{kind:'drawing,definedName,thread',maxChars:150000});
add('formulas-small-budget','formulas','Constrain formula records with maxChars',{kind:'formula',sheetId:'Fee engine',maxChars:1000,options:{maxResults:300}},{scope:'sheet',limit:300});
add('table-text-preview','more','Shorten text inside a grid preview',{kind:'table',sheetId:'Fee engine',range:'C4:H12',maxChars:400,tableMaxRows:2,tableMaxCols:2,tableMaxCellChars:10});
add('region-long-text','regions','Allow longer text in region previews',{kind:'region',sheetId:'Fee engine',range:'A1:W109',maxChars:40000,tableMaxRows:2,tableMaxCols:3,tableMaxCellChars:80},{mode:'region',scope:'sheet',budget:40000,preview:'2 × 3 · text 80'});
for(const [id,label,input] of [
 ['match-all','Find literal profit text',{kind:'match',searchTerm:'profit',maxChars:150000,options:{useRegex:false,maxResults:100}}],
 ['match-limit','Return the first three matches',{kind:'match',searchTerm:'profit',maxChars:150000,options:{useRegex:false,maxResults:3}}],
 ['match-offset','Skip three matches with offset',{kind:'match',searchTerm:'profit',offset:3,maxChars:150000,options:{useRegex:false,maxResults:3}}],
 ['match-pattern','Use a text pattern on one sheet',{kind:'match',sheetId:'Fee engine',searchTerm:'management.*fee|fee.*margin',maxChars:150000,options:{useRegex:true,maxResults:100}}],
 ['match-formulas','Search formula text as well as values',{kind:'match',sheetId:'Fee engine',searchTerm:'Assumptions!',maxChars:150000,options:{matchFormulas:true,maxResults:100}}]
])add(id,'search',label,input);
const serialize=r=>({ndjson:r.ndjson,recordCount:r.recordCount,truncated:r.truncated,metadata:r.metadata});
const results=[];
for(const c of cases){
 const start=Date.now();
 try{const response=serialize(await wb.inspect(c.input));results.push({...c,method:'inspect',response,elapsedMs:Date.now()-start});console.log(JSON.stringify({id:c.id,records:response.recordCount,chars:response.ndjson.length,truncated:response.truncated}));}
 catch(e){throw new Error(`${c.id}: ${e.message}`);}
}
const findInput={searchTerm:'profit',options:{maxResults:3}};
results.push({id:'find-cells',group:'search',label:'Use findCells for matches and total count',method:'findCells',input:findInput,response:JSON.parse(JSON.stringify(wb.findCells(findInput)))});
const direct=wb.worksheets.getItem('Fee engine').getRange('C4:H12');
results.push({id:'direct-range',group:'more',label:'Read an explicit range directly',method:'range',input:{sheetId:'Fee engine',range:'C4:H12'},response:{values:direct.values,formulas:direct.formulas}});
const inventory=results.find(c=>c.id==='map-workbook-150000').response.ndjson.split('\n').filter(Boolean).map(JSON.parse);
const anchor=inventory.find(x=>x.kind==='sheet'&&x.name==='Fee engine');
if(anchor?.id){const input={kind:'sheet,table',target:{id:anchor.id,beforeLines:0,afterLines:1},maxChars:150000};results.push({id:'target-context',group:'more',label:'Read a window around a sheet record',method:'inspect',input,response:serialize(await wb.inspect(input))});}
const after=await fs.readFile(source);
if(hash(bytes)!==hash(after))throw new Error('Source file changed');
const packageMeta=JSON.parse(await fs.readFile(path.join(home,'node_modules/@oai/artifact-tool/package.json'),'utf8'));
const receipt={capturedAt:new Date().toISOString(),packageVersion:packageMeta.version,source:'files/inspection/source.xlsx',sourceSha256:hash(bytes),sourceUnchanged:true,baselinePurpose:'Separate direct-range reads for checking returned coverage; not part of any inspect response.',calls:results.length};
await fs.writeFile(path.join(out,'captures.json'),JSON.stringify({receipt,baseline,cases:results},null,2));
console.log(JSON.stringify(receipt));

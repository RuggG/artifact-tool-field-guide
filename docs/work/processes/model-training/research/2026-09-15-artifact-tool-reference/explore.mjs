import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const root=path.dirname(fileURLToPath(import.meta.url));
const out=path.join(root,'files','runs');
await fs.mkdir(out,{recursive:true});
const mode=process.argv[2]||'core';
const runName=mode+(process.env.RUN_LABEL?'-'+process.env.RUN_LABEL:'');
const receipts=[];
const eq=(a,b)=>assert.deepEqual(a,b);
const close=(a,b,tol=1e-9)=>assert.ok(typeof a==='number'&&Math.abs(a-b)<=tol, `${a} != ${b}`);
const val=(s,a)=>s.getRange(a).values[0]?.[0];
function fresh(name='Data'){const w=Workbook.create();return [w,w.worksheets.add(name)];}
async function save(name,value){await fs.writeFile(path.join(out,name),typeof value==='string'?value:JSON.stringify(value,(_,v)=>typeof v==='bigint'?v.toString():v,2));}
async function test(id,description,fn){if(process.env.PROBE_FILTER&&!new RegExp(process.env.PROBE_FILTER).test(id))return;const start=Date.now();try{const detail=await fn();receipts.push({id,description,status:'pass',ms:Date.now()-start,detail});}catch(e){receipts.push({id,description,status:'fail',ms:Date.now()-start,error:String(e),stack:e.stack?.split('\n').slice(0,5)});}await save(runName+'.json',receipts);console.log(JSON.stringify(receipts.at(-1)));}
async function xlsx(w,name){w.recalculate();const f=await SpreadsheetFile.exportXlsx(w);await f.save(path.join(out,name+'.xlsx'));return f;}

if(mode==='formula-worker'){
 const records=(await fs.readFile(path.join(out,'formula-help.ndjson'),'utf8')).trim().split('\n').map(JSON.parse);
 const entry=records[Number(process.argv[3])],results=[];
 for(const formula of entry.examples||[]){
  const start=Date.now();
  try{
   const [w,s]=fresh('Sheet1');const data=w.worksheets.add('Data');
   // Shared synthetic fixture; examples are not rewritten. Errors may be caused
   // by an example's missing context and do not prove lack of function support.
   const grid=Array.from({length:20},(_,i)=>[i+1,(i+1)*2,(i+1)*3,(i+1)*4,(i+1)*5]);
   s.getRange('A1:E20').values=grid;data.getRange('A1:E20').values=grid;
   s.getRange('Z100').formulas=[[formula]];w.recalculate();
   results.push({formula,value:val(s,'Z100'),projection:s.getRange('Z100:AC103').values,ms:Date.now()-start});
  }catch(e){results.push({formula,error:String(e),ms:Date.now()-start});}
 }
 await save('formula-'+String(process.argv[3]).padStart(3,'0')+'.json',{name:entry.name,category:entry.category,examples:results});
 process.exit(0);
}

if(mode==='api-worker'){
 const entries=(await fs.readFile(path.join(out,'api-help.ndjson'),'utf8')).trim().split('\n').map(JSON.parse).filter(x=>x.kind==='api').flatMap(x=>(x.examples||[]).map(example=>({name:x.name,...example})));
 const index=Number(process.argv[3]),entry=entries[index],folder=path.join(out,'api-examples',String(index).padStart(3,'0'));
 await fs.mkdir(folder,{recursive:true});process.chdir(folder);
 const [w,s]=fresh('Sheet1');const d=w.worksheets.add('Data');s.getRange('A1:C6').values=[['Name','Value','Other'],['Alpha',1,2],['Beta',3,4],['Gamma',5,6],['Delta',7,8],['Epsilon',9,10]];d.getRange('A1:C6').copyFrom(s.getRange('A1:C6'));
 for(const name of ['Charts','Summary','Shapes','Revenue']){const extra=w.worksheets.add(name);extra.getRange('A1:C6').copyFrom(s.getRange('A1:C6'));}
 const c=s.charts.add('line',s.getRange('A1:C6'));c.setPosition('F1','N15');
 const logs=[];const quiet={log:(...v)=>logs.push(v.map(x=>typeof x==='string'?x:JSON.stringify(x)).join(' ')),error:(...v)=>logs.push(v.map(String).join(' '))};
 const result={index,name:entry.name,summary:entry.summary,originalCode:entry.code};
 // The help catalogue mixes JavaScript and TypeScript's non-null assertion.
 let rawCode=entry.code;
 const conversion=[];
 if(!rawCode){
  const options=Object.fromEntries(Object.entries(entry).filter(([k,v])=>!['name','summary'].includes(k)&&v!==null));
  if(entry.name==='workbook.help'){const {query,...rest}=options;rawCode='console.log(workbook.help('+JSON.stringify(query)+','+JSON.stringify(rest)+'));';}
  if(entry.name==='workbook.inspect'){if(options.target?.id==='ws/<sheetId>')options.target.id='ws/'+s.id;rawCode='console.log(await workbook.inspect('+JSON.stringify(options)+'));';}
  if(entry.name==='workbook.resolve'){w.comments.setSelf({displayName:'Ruggero Gargiulo'});const th=w.comments.addThread({cell:s.getRange('A1')},'Probe');const id=entry.id.startsWith('ws/')?'ws/'+s.id:entry.id.startsWith('th/')?'th/'+th.id:'ch/'+c.id;rawCode='console.log(workbook.resolve('+JSON.stringify(id)+').id);';}
  conversion.push('Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.');
 }
 if(index===49){rawCode="workbook.help('*',{search:'trendline|trendlines',include:['index','examples']});";conversion.push('Repaired malformed trailing dot and used the documented query/options signature.');}
 const code=rawCode.replace(/\]!(?=[.;,\)\[])/g,']');
 result.adaptations=[...conversion,...(code===rawCode?[]:['Removed TypeScript non-null assertion after array access.'])];
 try{
 const AsyncFunction=Object.getPrototypeOf(async function(){}).constructor;
 await new AsyncFunction('Workbook','SpreadsheetFile','FileBlob','workbook','wb','worksheet','sheet','range','chart','series','fs','outputDir','sourceRange','targetRange','console','{\n'+code+'\n}')(Workbook,SpreadsheetFile,FileBlob,w,w,s,s,s.getRange('A1:C6'),c,c.series.items[0],fs,folder,s.getRange('A1:B2'),s.getRange('A1:B6'),quiet);
 result.execution='completed';try{await xlsx(w,'api-example-'+String(index).padStart(3,'0'));result.export='completed';}catch(e){result.exportError=String(e);}
 }catch(e){result.executionError=String(e);}result.logs=logs;
 await save('api-example-'+String(index).padStart(3,'0')+'.json',result);process.exit(0);
}

if(mode==='discover'){
 const [w]=fresh();
 for(const [name,query,opts] of [
   ['api-help','*',{include:'index,examples,notes',maxChars:2000000}],
   ['formula-help','fx.*',{include:'index,examples,notes',maxChars:2000000}],
   ['enums-help','enum.*',{include:'index,examples,notes',maxChars:1000000}],
 ]){const r=w.help(query,opts);await save(name+'.json',r);await save(name+'.ndjson',r.ndjson);console.log(name,JSON.stringify({count:r.recordCount,truncated:r.truncated,keys:Object.keys(r),chars:r.ndjson.length,first:r.ndjson.split('\n')[0]}));}
}

if(mode==='surface'){
 const [w,s]=fresh();s.getRange('A1:C3').values=[['Name','Qty','Price'],['A',1,2],['B',3,4]];
 const table=s.tables.add('A1:C3',true,'DataTable'),chart=s.charts.add('bar',s.getRange('A1:C3'));
 const objects={workbook:w,worksheets:w.worksheets,worksheet:s,range:s.getRange('A1:B2'),format:s.getRange('A1').format,
 tables:s.tables,table,tableRows:table.rows,tableColumns:table.columns,
 notes:w.notes,comments:w.comments,names:w.names,pivotTables:s.pivotTables,slicers:s.slicers,
 awareness:w.awareness,charts:s.charts,chart,series:chart.series,seriesItem:chart.series.items[0],
 shapes:s.shapes,images:s.images,sparklines:s.sparklineGroups,dataTables:s.dataTables,dataValidations:s.dataValidations,
 conditionalFormats:s.getRange('B2:B3').conditionalFormats,freezePanes:s.freezePanes};
 const surface={};for(const [name,obj] of Object.entries(objects)){if(!obj){surface[name]=null;continue;}const desc=Object.getOwnPropertyDescriptors(Object.getPrototypeOf(obj));
 surface[name]=Object.entries(desc).filter(([k])=>k!=='constructor').map(([key,d])=>({key,kind:typeof d.value==='function'?'method':d.get?'getter':'property',set:!!d.set,arity:typeof d.value==='function'?d.value.length:undefined}));
 }
 await save('public-surface.json',surface);
 const selected={};for(const [name,obj] of Object.entries(objects)){if(!['notes','comments','pivotTables','slicers','awareness','tableRows','tableColumns','dataTables'].includes(name)||!obj)continue;
 for(const prop of surface[name])if(prop.kind==='method')selected[name+'.'+prop.key]=obj[prop.key].toString();}
 await save('targeted-public-methods.json',selected);
 console.log(JSON.stringify(surface));
}

if(mode==='core'){
await test('C01','Workbook create, aliases, worksheet collection lifecycle',async()=>{
 const w=Workbook.create();eq(w.worksheets.getSheetCount(),0);
 const [a,b]=w.worksheets.add(['Alpha','Beta']);eq(w.sheets,w.worksheets);
 eq(w.worksheets.getItem('Alpha').id,a.id);eq(w.worksheets.getItemAt(1).id,b.id);
 w.worksheets.setActiveWorksheet('Beta');eq(w.worksheets.getActiveWorksheet().name,'Beta');
 a.name='Renamed';a.index=1;a.showGridLines=false;a.tabColor='#123456';
 const detail={count:w.worksheets.getSheetCount(),index:w.worksheets.getSheetIndex('Renamed'),first:w.worksheets.getFirst().name,nameAt0:w.worksheets.getSheetNameByIndex(0),id:a.id,sheetId:a.sheetId,tabColor:a.tabColor};
 b.delete();eq(w.worksheets.getSheetCount(),1);return detail;
});
await test('C02','Worksheet duplicate and missing lookups',()=>{
 const [w,s]=fresh();const duplicate=w.worksheets.add('Data');
 return {sameId:duplicate.id===s.id,count:w.worksheets.getSheetCount(),missing:w.worksheets.getItemOrNullObject('Missing')};
});
await test('C03','Typed matrices, one-dimensional rows/columns, dates, literal formula text',()=>{
 const [w,s]=fresh();s.getRange('A1:C2').values=[[12,true,null],[new Date('2026-01-01T00:00:00Z'),'=2+3',"'=2+3"]];
 s.getRange('E1:G1').values=[1,2,3];s.getRange('E2:E4').values=[4,5,6];w.recalculate();
 eq(val(s,'B2'),5);eq(val(s,'C2'),'=2+3');eq(s.getRange('E1:G1').values,[[1,2,3]]);
 return {values:s.getRange('A1:C2').values,formulas:s.getRange('A1:C2').formulas,rows:s.getRange('E1:G4').values};
});
await test('C04','Broadcast values versus formula assignment',()=>{
 const [w,s]=fresh();s.getRange('A1:B2').values=[[7]];eq(s.getRange('A1:B2').values,[[7,7],[7,7]]);
 let formulaError;try{s.getRange('D1:E2').formulas=[['=1+1']];}catch(e){formulaError=String(e);}
 return {formulaError,formulas:s.getRange('D1:E2').formulas,values:s.getRange('D1:E2').values};
});
await test('C05','Single-cell expansion and write resizing',()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1,2],[3,4]];eq(s.getRange('A1:B2').values,[[1,2],[3,4]]);
 s.getRange('D1').write([[5,6],[7,8]]);s.getRange('G1').write({formulas:[['=SUM(A1:B2)']]});eq(val(s,'G1'),10);
 let rejected=false;try{s.getRange('A1').write([[9]],{overwrite:'error'});}catch(e){rejected=String(e);}
 assert.ok(rejected);eq(val(s,'A1'),1);return {rejected,expanded:s.getRange('D1:E2').values};
});
await test('C06','Write clear and resize:none boundary',()=>{
 const [w,s]=fresh();s.getRange('A1:B2').values=[[1,2],[3,4]];s.getRange('A1:B2').format.fill='#ff0000';
 let err;try{s.getRange('D1:E2').write([[1,2,3]],{resize:'none'});}catch(e){err=String(e);}
 s.getRange('A1:B2').write({values:[[9]]},{clear:'all'});return {err,values:s.getRange('A1:B2').values};
});
await test('C07','Index/address/navigation conventions',()=>{
 const [w,s]=fresh();const r=s.getRange('C3:E5');r.values=[[1,2,3],[4,5,6],[7,8,9]];
 eq(s.getCell(2,2).values,[[1]]);eq(r.getCell(1,1).values,[[5]]);eq(r.getRow(1).values,[[4,5,6]]);eq(r.getColumn(1).values,[[2],[5],[8]]);
 return {address:r.address,getAddress:r.getAddress(),rowIndex:r.rowIndex,columnIndex:r.columnIndex,rowCount:r.rowCount,columnCount:r.columnCount,
 relativeA1:r.getRange('A1').address,byIndex:r.getRangeByIndexes(1,1,1,1).address,
 offset:r.offset(1,2).address,offsetAlias:r.getOffsetRange(1,2).address,resize:r.resize(2,2).address,
 resized:r.getResizedRange(1,1).address,resizeAlias:r.getResizeRange(2,2).address,current:r.getCurrentRegion().address,used:s.getUsedRange(true).address};
});
await test('C08','Fill down/right shifts relative and preserves anchored references',()=>{
 const [w,s]=fresh();s.getRange('A1:C3').values=[[1,2,3],[4,5,6],[7,8,9]];
 s.getRange('D1').formulas=[['=A1+$B$1+C$1']];s.getRange('D1:D3').fillDown();
 eq(s.getRange('D1:D3').formulas,[['=A1+$B$1+C$1'],['=A2+$B$1+C$1'],['=A3+$B$1+C$1']]);
 s.getRange('A5').formulas=[['=A1*2']];s.getRange('A5:C5').fillRight();eq(s.getRange('A5:C5').values,[[2,4,6]]);
 return {down:s.getRange('D1:D3').values,right:s.getRange('A5:C5').formulas};
});
await test('C09','R1C1, copy all/values/formulas, fillFrom',()=>{
 const [w,s]=fresh();s.getRange('A1:B2').values=[[2,3],[4,5]];s.getRange('C1:C2').formulasR1C1=[['=RC[-2]+RC[-1]'],['=RC[-2]+RC[-1]']];eq(s.getRange('C1:C2').values,[[5],[9]]);
 s.getRange('D1:D2').copyFrom(s.getRange('C1:C2'),'values');eq(s.getRange('D1:D2').values,[[5],[9]]);
 s.getRange('C1:C2').copyTo(s.getRange('E1:E2'),'formulas');
 s.getRange('A4:B5').copyFrom(s.getRange('A1:B2'),'all');s.getRange('G1:G2').fillFrom(s.getRange('C1:C2'));
 return {r1c1:s.getRange('C1:C2').formulasR1C1,copiedFormulas:s.getRange('E1:E2').formulas,fillFrom:s.getRange('G1:G2').formulas};
});
await test('C10','Clear modes retain appropriate contents/styles; null removes values',()=>{
 const [w,s]=fresh();const r=s.getRange('A1:B2');r.values=[[1,2],[3,4]];r.format.fill='#ff0000';r.clear({applyTo:'formats'});eq(val(s,'A1'),1);
 r.format.fill='#00ff00';r.clear({applyTo:'contents'});eq(val(s,'A1'),null);r.values=[[true]];r.values=null;eq(val(s,'A1'),null);
 r.clear({applyTo:'all'});return {values:r.values};
});
await test('C11','Merge, across merge, worksheet merge aliases',async()=>{
 const [w,s]=fresh();s.getRange('A1:C1').merge();s.getRange('A1').values=[['Merged']];s.getRange('A3:C4').merge(true);s.mergeCells('E1:F2');s.unmergeCells('E1:F2');s.getRange('A3:C4').unmerge();await xlsx(w,'merges');return {protoMerges:w.toProto().sheets?.[0]?.merges};
});
await test('C12','Formatting, theme, sizing, autofit and format matrices',async()=>{
 const [w,s]=fresh();s.getRange('A1:C4').values=[['Metric','Value','Rate'],['Alpha',1234.56,0.12],['Beta',-123,0],['Date',new Date('2026-01-01'),null]];
 w.setColorScheme({name:'Probe',themeColors:{accent1:'#123456',accent2:'#007755',bg1:'#ffffff',tx1:'#000000'}});
 s.getRange('A1:C4').format={font:{name:'Arial',size:11},verticalAlignment:'center',rowHeight:20,columnWidth:18};
 s.getRange('A1:C1').format={fill:'accent1',font:{bold:true,color:'#ffffff'},horizontalAlignment:'center'};
 s.getRange('B2:B3').setNumberFormat('#,##0.00;(#,##0.00)');s.getRange('C2:C3').format.numberFormat=[['0.0%'],['0%']];s.getRange('B4').setNumberFormat('yyyy-mm-dd');
 s.getRange('A1:C4').format.borders={preset:'outside',style:'thin',color:'#999999'};
 s.getRange('A1:C4').format.autofitColumns();s.getRange('A1:C4').format.autofitRows();
 s.getRange('A1:A4').format.columnWidthPx=130;s.getRange('A1:C1').format.rowHeightPx=30;
 s.showGridLines=false;await xlsx(w,'formatting');const layout=await w.export({sheetName:'Data',range:'A1:C4',format:'layout'});await save('formatting-layout.json',await layout.text());
 const png=await w.render({sheetName:'Data',range:'A1:C4',scale:2});await fs.writeFile(path.join(out,'formatting.png'),new Uint8Array(await png.arrayBuffer()));
 return {rowHeight:s.getRange('A1').format.rowHeight,columnWidth:s.getRange('A1').format.columnWidth,numberFormat:s.getRange('B2:C3').format.numberFormat,theme:w.theme};
});
await test('C13','Theme/RGB transformed colors, gradient fills, borders and alignments',async()=>{
 const [w,s]=fresh();s.getRange('A1:F4').values=[['Solid','Theme','Gradient','Borders','Wrapped','Centered'],[1,2,3,4,'long text wraps here',6],[2,3,4,5,6,7],[3,4,5,6,7,8]];
 s.getRange('A2:A4').format.fill={type:'solid',color:{type:'rgb',value:'#ff0000',transform:{opacity:0.5}}};
 s.getRange('B2:B4').format.fill={type:'theme',value:'accent1',transform:{lighten:0.3}};
 s.getRange('C2:C4').format.fill={type:'gradient',stops:[{offset:0,color:'#ffffff'},{offset:1,color:'#2563eb'}],angleDeg:45};
 s.getRange('D2:D4').format.borders={top:{style:'thick',color:'#000000'},bottom:{style:'dashed'},insideHorizontal:{style:'dotted'}};
 s.getRange('E2').format.wrapText=true;s.getRange('F2:G2').format.horizontalAlignment='centerAcrossSelection';await xlsx(w,'fills');return {accepted:true};
});
await test('C14','Freeze rows and columns, unfreeze, worksheet reset',async()=>{
 const [w,s]=fresh();s.getRange('A1:B2').values=[[1,2],[3,4]];s.freezePanes.freezeRows(1);s.freezePanes.freezeColumns(1);await xlsx(w,'freeze');
 s.freezePanes.unfreeze();s.reset({clear:'used',applyTo:'all',deleteTables:true,deleteCharts:true,deleteDrawings:true,deleteSparklines:true});eq(val(s,'A1'),null);return {reset:true};
});
await test('C15','Named range calculation, scopes and deletion',async()=>{
 const [w,s]=fresh();s.getRange('A1:A3').values=[[1],[2],[3]];const n=w.names.addRange('SalesRange','Data!$A$1:$A$3',{description:'Synthetic sales'});
 s.names.addRange('LocalRange','Data!$A$2:$A$3');s.getRange('B1:B2').formulas=[['=SUM(SalesRange)'],['=SUM(LocalRange)']];eq(s.getRange('B1:B2').values,[[6],[5]]);
 eq(w.definedNames,w.names);const address=w.names.getItem('SalesRange').getRange().address;await xlsx(w,'names');w.names.getItem('SalesRange').delete();return {address};
});
await test('C16','Named LAMBDA functions and prefixed variables',async()=>{
 const [w,s]=fresh();w.names.addFunction('AddTax',{formula:'=_xlfn.LAMBDA(_xlpm.amount,_xlpm.amount*1.1)',description:'Add synthetic tax'});s.getRange('A1').formulas=[['=AddTax(100)']];w.recalculate();await xlsx(w,'named-function');return {value:val(s,'A1'),names:(await w.inspect({kind:'definedName'})).ndjson};
});
await test('C17','Threaded cell/range comments, reply, reactions and state',async()=>{
 const [w,s]=fresh();s.getRange('A1:B2').values=[[1,2],[3,4]];w.comments.setSelf({displayName:'Ruggero Gargiulo',initials:'RG'});
 const th=w.comments.addThread({cell:s.getRange('A1')},'Synthetic review');const reply=th.addReply('Checked');reply.toggleReaction('👍');th.resolve();th.reopen();
 w.comments.addThread({range:s.getRange('A2:B2')},'Range review');const ins=await w.inspect({kind:'thread',maxChars:10000});await save('comments-inspect.json',ins);await xlsx(w,'comments');return {inspect:ins.ndjson};
});
await test('C18','Tables create, append, style, header/totals and delete',async()=>{
 const [w,s]=fresh();s.getRange('A1:C3').values=[['SKU','Qty','Price'],['A',2,10],['B',3,20]];const t=s.tables.add('A1:C3',true,'Inventory');
 t.rows.add(null,[['C',4,30]]);eq(t.getDataRows().length,3);t.name='Stock';t.style='TableStyleMedium2';t.showTotals=true;t.showBandedColumns=true;t.showFilterButton=true;
 s.getRange('E1').formulas=[['=SUM(Stock[Qty])']];w.recalculate();close(val(s,'E1'),9);await xlsx(w,'table');
 const detail={address:t.address,headers:t.getHeaderRowRange().values,data:t.getDataRows(),formula:val(s,'E1')};t.delete();eq(s.tables.items.length,0);return detail;
});
await test('C19','Headerless table and overlap acceptance boundary',async()=>{
 const [w,s]=fresh();s.getRange('A1:C3').values=[['A',1,2],['B',3,4],['C',5,6]];const t=s.tables.add('A1:B3',false,'Headerless');let overlap;
 try{s.tables.add('B1:C3',false,'Overlap');overlap='accepted';}catch(e){overlap=String(e);}return {headerless:t.getRange().values,overlap};
});
await test('C20','CSV quote/newline parsing and string typing',async()=>{
 const w=await Workbook.fromCSV('Name,Qty,Note\n"North, Ltd",2,"Line 1\nLine 2"\nSouth,03,plain',{sheetName:'CSV'});const s=w.worksheets.getItemAt(0);
 eq(val(s,'B2'),'2');eq(val(s,'A2'),'North, Ltd');const before=s.getRange('A1:C3').values;s.getRange('B2:B3').values=s.getRange('B2:B3').values.map(r=>[Number(r[0])]);s.getRange('E1').formulas=[['=SUM(B2:B3)']];close(val(s,'E1'),5);
 let append;try{await w.fromCSV('A,B\n1,2',{sheetName:'Second'});append='accepted';}catch(e){append=String(e);}return {before,converted:val(s,'E1'),append};
});
await test('C21','Instance CSV and Markdown import',async()=>{
 const w=Workbook.create();const r=await w.fromCSV('A,B\n1,2',{sheetName:'CSV'});const md=await Workbook.fromMarkdown('| Name | Value |\n| --- | ---: |\n| Alpha | 10 |\n| Beta | 20 |',{sheetName:'MD',format:true});
 await xlsx(md,'markdown-import');return {instanceRange:r.range?.address,instanceSheet:r.sheet?.name,markdown:md.worksheets.getItemAt(0).getUsedRange().values,tableCount:md.worksheets.getItemAt(0).tables.items.length};
});
await test('C22','HTML copy/paste values versus formulas',async()=>{
 const [w,s]=fresh();s.getRange('A1:B2').values=[['Qty','Double'],[5,'=A2*2']];s.getRange('A1:B1').format.font.bold=true;
 const html=w.toHTML(0,'A1:B2',{formulas:true});await save('copy-with-formulas.html',html);const [target,t]=fresh('Target');const result=target.fromHTML(0,html,'C3');
 target.recalculate();return {result,values:t.getRange('C3:D4').values,formulas:t.getRange('C3:D4').formulas,htmlLength:html.length};
});
await test('C23','Serialize/load, validated load, constructor and utilities',()=>{
 const [w,s]=fresh();s.getRange('A1:B1').values=[[3,'=A1*4']];const proto=w.toProto();const loaded=Workbook.load(proto,{validate:true});eq(val(loaded.worksheets.getItem('Data'),'B1'),12);
 const constructed=new Workbook(proto);eq(val(constructed.worksheets.getItem('Data'),'A1'),3);
 eq(w.utils.columnToLetter(27),'AA');eq(w.utils.letterToColumn('AA'),27);return {protoKeys:Object.keys(proto),address:w.utils.toA1String(1,1,3,2),fillRight:w.utils.fillRight([[1]],3),fillDown:w.utils.fillDown([[1]],3)};
});
await test('C24','Cross-sheet dependencies, recalculation, trace and stats',async()=>{
 const [w,s]=fresh('Input');const model=w.worksheets.add('Build');s.getRange('A1:B2').values=[['Base',100],['Growth',0.1]];model.getRange('A1').formulas=[["='Input'!B1*(1+'Input'!B2)"]];close(val(model,'A1'),110);
 s.getRange('B2').values=[[0.2]];w.recalculate();close(val(model,'A1'),120);const trace=w.trace('Build!A1');await save('trace.json',trace);return {trace,stats:w.collectFormulaUsageStats(),badTrace:w.trace('Missing!A1')};
});
await test('C25','Dynamic arrays, spill descriptors and obstruction',async()=>{
 const [w,s]=fresh();s.getRange('A1').formulas=[['=SEQUENCE(3,2)']];w.recalculate();eq(s.getRange('A1:B3').values,[[1,2],[3,4],[5,6]]);
 const before={formulas:s.getRange('A1:B3').formulas,display:s.getRange('A1:B3').displayFormulas,infos:s.getRange('A1:B3').formulaInfos};
 let childEdit;try{s.getRange('B2').values=[[99]];childEdit='accepted';}catch(e){childEdit=String(e);}
 s.getRange('E2').values=[[42]];s.getRange('D1').formulas=[['=SEQUENCE(3,2)']];w.recalculate();return {before,childEdit,blocked:s.getRange('D1:E3').values};
});
await test('C26','Inspection kind families, bounding, regex error search and resolve',async()=>{
 const [w,s]=fresh();s.getRange('A1:B3').values=[['Key','Value'],['East',5],['West','=1/0']];s.tables.add('A1:B3',true,'Regions');
 const snapshots={};for(const kind of ['workbook','sheet','table','region','formula','computedStyle','match','drawing','definedName','thread']){
 const opts={kind,sheetId:'Data',range:'A1:B3',maxChars:2500,tableMaxRows:3,tableMaxCols:2};
 if(kind==='match')Object.assign(opts,{searchTerm:'#DIV/0!',options:{maxResults:5}});
 snapshots[kind]=await w.inspect(opts);}
 const index=await w.inspect({kind:'sheet',include:'id,name'});const records=index.ndjson.split('\n').filter(Boolean).map(JSON.parse);const anchor=records.find(r=>r.id)?.id;assert.ok(anchor);eq(w.resolve(anchor).name,'Data');
 const find=w.findCells({searchTerm:'East',sheetId:'Data',options:{maxResults:2,matchFormulas:true}});
 await save('inspect-kinds.json',snapshots);return {sheet:index.ndjson,find,errorMatch:snapshots.match.ndjson};
});
await test('C27','Error scan versus missing-data and full-column boundaries',()=>{
 const [w,s]=fresh();s.getRange('A1:A4').values=[['x'],['x'],[null],['x']];s.getRange('B1:B2').values=[[1],[2]];
 const formulas=['=COUNTIF(A1:A4,"")','=COUNTIFS(A1:A4,"")','=COUNTBLANK(A1:A4)','=SUMIFS(B:B,A:A,"x")','=SUMIFS(B1:B4,A1:A4,"x")','=ROWS(A:A)','=1/0','=NA()','=MISSINGFUNCTION(1)'];
 s.getRange('D1:D9').formulas=formulas.map(x=>[x]);w.recalculate();return formulas.map((formula,i)=>({formula,value:val(s,'D'+(i+1))}));
});
await test('C28','Recorded patch/replay and patch input envelope',async()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];const clone=Workbook.load(w.toProto());
 const {result,patch,idMap,crdtUpdateV2}=w.record(()=>{s.getRange('A1').values=[[2]];s.getRange('B1').formulas=[['=A1*3']];return 'done';});
 await save('recorded-patch.json',patch);const applied=clone.apply(patch);close(val(clone.worksheets.getItem('Data'),'B1'),6);
 return {result,patch,idMap,crdtBytes:crdtUpdateV2?.length,applied};
});
await test('C29','Collaborative state, CRDT subscription and replica application',async()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];w.hydrateCrdtFromProto();const replica=Workbook.load(w.toProto());replica.hydrateCrdtFromProto();let updates=[];
 const unsubscribe=w.onCrdtUpdateV2((update,origin)=>updates.push({update,origin:String(origin)}));
 const recorded=w.record(()=>{s.getRange('A1').values=[[99]];});unsubscribe();
 if(recorded.crdtUpdateV2)replica.applyCrdtUpdateV2(recorded.crdtUpdateV2,{recalculate:true});
 return {ready:w.isCollaborativeStateReady(),updates:updates.map(x=>({bytes:x.update.length,origin:x.origin})),recordedBytes:recorded.crdtUpdateV2?.length,replicaValue:val(replica.worksheets.getItem('Data'),'A1')};
});
await test('C30','Native export and import preserve values, formulas and update behavior',async()=>{
 const [w,s]=fresh();s.getRange('A1:C3').values=[['Qty','Rate','Value'],[3,4,'=A2*B2'],[5,6,'=A3*B3']];await xlsx(w,'roundtrip');
 const imported=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,'roundtrip.xlsx')));const t=imported.worksheets.getItem('Data');eq(val(t,'C2'),12);t.getRange('A2').values=[[10]];imported.recalculate();eq(val(t,'C2'),40);return {formulas:t.getRange('C2:C3').formulas,values:t.getRange('C2:C3').values};
});
}

if(mode==='features'){
const cfRules=[
 ['cellIs',{operator:'greaterThan',formula:2}],['CellValue',{operator:'between',formula:[2,4]}],
 ['Custom',{formula:'=A2>2'}],['expression',{formula:'=A2>2'}],
 ['colorScale',{colors:['#ffffff','#2563eb'],thresholds:['min','max']}],
 ['dataBar',{color:'#2563eb',gradient:true}],['iconSet',{iconSet:'3Arrows',showValue:true,reverse:false}],
 ...['containsText','notContainsText','beginsWith','endsWith'].map(t=>[t,{text:'a'}]),
 ...['containsBlanks','notContainsBlanks','containsErrors','notContainsErrors','duplicateValues','uniqueValues'].map(t=>[t,{}]),
 ['timePeriod',{timePeriod:'today'}],['top10',{rank:2,bottom:false,percent:false}],
 ['aboveAverage',{aboveAverage:true,equalAverage:true,stdDev:0}],
];
await test('F01','All conditional-format rule families; serialize, render and clear',async()=>{
 const [w,s]=fresh('CF');const details=[];
 for(let i=0;i<cfRules.length;i++){
 const [type,config]=cfRules[i],col=w.utils.columnToLetter(i+1);
 s.getRange(col+'1').values=[[type]];s.getRange(col+'2:'+col+'8').values=[[1],[2],[3],[4],[null],['a'],['=1/0']];
 try{s.getRange(col+'2:'+col+'8').conditionalFormats.add(type,{...config,...(!['colorScale','dataBar','iconSet'].includes(type)?{format:{fill:'#fee2e2',font:{color:'#991b1b'}}}:{})});details.push({type,accepted:true});}catch(e){details.push({type,error:String(e)});}
 }
 s.getRange('A1:T8').format.columnWidth=18;s.getRange('A1:T1').format.wrapText=true;
 s.getRange('V2:V5').values=[[1],[2],[3],[4]];s.getRange('V2:V5').conditionalFormats.addCustom('=V2>2',{fill:'#ff0000'});
 await xlsx(w,'conditional-formats');await save('conditional-formats-cache.json',w.getConditionalFormattingRenderCache('CF'));
 const png=await w.render({sheetName:'CF',range:'A1:J8',scale:1});await fs.writeFile(path.join(out,'conditional-formats.png'),new Uint8Array(await png.arrayBuffer()));
 s.getRange('V2:V5').conditionalFormats.clear();s.getRange('A2:A8').conditionalFormats.deleteAll();return details;
});
await test('F02','All validation types, settings, collection and Office-style assignment',async()=>{
 const [w,s]=fresh('Validation');const rules=[
 {type:'none'},{type:'whole',operator:'between',formula1:1,formula2:10},{type:'decimal',operator:'greaterThan',formula1:0},
 {type:'list',values:['Open','Closed']},{type:'date',operator:'between',formula1:'DATE(2026,1,1)',formula2:'DATE(2026,12,31)'},
 {type:'time',operator:'lessThan',formula1:0.5},{type:'textLength',operator:'lessThanOrEqual',formula1:10},{type:'custom',formula1:'A2>0'}];
 for(let i=0;i<rules.length;i++){const col=w.utils.columnToLetter(i+1);s.getRange(col+'1').values=[[rules[i].type]];s.getRange(col+'2:'+col+'5').dataValidation={rule:rules[i],prompt:{title:'Input',message:'Synthetic test',show:true},errorAlert:{title:'Invalid',message:'Check input',style:['stop','warning','information'][i%3],show:true},ignoreBlanks:true,inCellDropDown:true};}
 s.getRange('J2:J5').dataValidation={list:{source:['Yes','No'],inCellDropDown:true},allowBlank:true};
 s.dataValidations.add({range:'L2:L5',rule:{type:'list',formula1:'Validation!$D$2:$D$3'}});
 s.getRange('B2').values=[[999]];eq(val(s,'B2'),999);await xlsx(w,'validations');return {types:rules.map(r=>r.type),invalidProgrammaticWriteAccepted:true};
});
await test('F03','All sparkline types, range alias, markers, axes and export',async()=>{
 const [w,s]=fresh('Sparklines');s.getRange('A1:D4').values=[['Q1','Q2','Q3','Q4'],[1,2,-1,4],[3,2,4,5],[4,null,2,3]];
 const groups=[];for(const [i,type]of ['line','column','stacked'].entries()){
 const col=w.utils.columnToLetter(i+6);const g=s.sparklineGroups.add({type,targetRange:col+'2:'+col+'4',sourceData:'A2:D4',seriesColor:'#2563eb',negativeColor:'#ff0000',lineWeight:2,displayHidden:true,markers:{show:true,high:true,low:true,first:true,last:true,negative:true},axis:{showAxis:true,rightToLeft:false}});groups.push(g);}
 const extra=s.getRange('J2:J4').sparklines.add('line',s.getRange('A2:D4'),{seriesColor:'#00aa00'});extra.delete();
 groups[0].markers.high=true;groups[0].axis.showAxis=true;groups[0].seriesColor='#007755';await xlsx(w,'sparklines');
 const png=await w.render({sheetName:'Sparklines',range:'A1:I5',scale:2});await fs.writeFile(path.join(out,'sparklines.png'),new Uint8Array(await png.arrayBuffer()));
 s.sparklineGroups.delete(groups[0]);s.sparklineGroups.deleteAll();return {types:['line','column','stacked']};
});
await test('F04','Data table full-range convention and independent two-variable results',async()=>{
 const [w,s]=fresh('Sensitivity');s.getRange('B2:B3').values=[[2],[3]];s.getRange('B5').formulas=[['=B2*10+B3']];
 s.getRange('E4').formulas=[['=B5']];s.getRange('F4:H4').values=[[1,2,3]];s.getRange('E5:E7').values=[[10],[20],[30]];
 const result=s.dataTables.add('E4:H7',{rowInput:'B2',columnInput:'B3'});w.recalculate();
 const actual=s.getRange('F5:H7').values;eq(actual,[[20,30,40],[30,40,50],[40,50,60]]);eq(s.getRange('B2:B3').values,[[2],[3]]);
 await xlsx(w,'data-table');return {resultType:typeof result,actual,formulas:s.getRange('F5:H7').formulas,display:s.getRange('F5:H7').displayFormulas,infos:s.getRange('F5').formulaInfos};
});
await test('F05','Data table older body-only and one-variable documentation',()=>{
 const [w,s]=fresh();s.getRange('B2:B3').values=[[2],[3]];s.getRange('E4').formulas=[['=B2*10+B3']];s.getRange('F4:H4').values=[[1,2,3]];s.getRange('E5:E7').values=[[10],[20],[30]];
 let body,one;try{s.dataTables.add('F5:H7',{rowInput:'B2',columnInput:'B3'});w.recalculate();body=s.getRange('E4:H7').values;}catch(e){body=String(e);}
 const [x,t]=fresh();t.getRange('B2').values=[[2]];t.getRange('E4').formulas=[['=B2*10']];t.getRange('F4:H4').values=[[1,2,3]];
 try{t.dataTables.add('E4:H5',{rowInput:'B2'});x.recalculate();one=t.getRange('E4:H5').values;}catch(e){one=String(e);}return {body,one};
});
await test('F06','Drawing images by SVG, bytes, data URL, path; replace and delete',async()=>{
 const [w,s]=fresh('Images');const svg='<svg xmlns="http://www.w3.org/2000/svg" width="80" height="40"><rect width="80" height="40" fill="#2563eb"/><text x="8" y="26" fill="white">Probe</text></svg>';
 const anchor={from:{row:1,col:1},extent:{widthPx:160,heightPx:80}};const im=s.images.add({svg,anchor,alt:'Synthetic blue label'});
 const bytes=await fs.readFile(path.join(out,'formatting.png'));
 const fromBytes=s.images.add({blob:bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength),contentType:'image/png',anchor:{...anchor,from:{row:7,col:1}}});
 const fromURL=s.images.add({dataUrl:'data:image/png;base64,'+bytes.toString('base64'),anchor:{...anchor,from:{row:13,col:1}}});
 const fromPath=s.images.add({path:path.join(out,'formatting.png'),anchor:{...anchor,from:{row:19,col:1}}});
 im.alt='Updated label';im.anchor={...anchor,from:{row:1,col:4}};fromPath.replace({svg});await xlsx(w,'images');
 const result=w.fromImage(0,{bytes,contentType:'image/png'},'J1:L8');const png=await w.render({sheetName:'Images',range:'A1:L16',scale:1});await fs.writeFile(path.join(out,'images.png'),new Uint8Array(await png.arrayBuffer()));
 const detail={count:s.images.items.length,bytes:fromBytes.bytes?.length,contentType:fromBytes.contentType,fromImage:result};fromURL.delete();return detail;
});
await test('F07','Shapes, text, drawing layout directions and deleteAllDrawings',async()=>{
 const [w,s]=fresh('Shapes');for(const [i,geometry]of ['rect','roundRect','ellipse','textbox'].entries()){const sh=s.shapes.add({geometry,name:'Shape'+i,anchor:{from:{row:1+i*5,col:1},extent:{widthPx:180,heightPx:65}},fill:'#e0e7ff',line:{fill:'#334155',width:1,style:'dashed'}});sh.text=geometry;}
 const layouts=[];for(const direction of ['vertical','horizontal','grid']){s.autoLayoutDrawings(s.shapes.items,{direction,align:'center',columns:2,frame:{startCell:'B2',width:600,height:400},gap:16,padding:20});layouts.push({direction,count:s.shapes.items.length});}
 await xlsx(w,'shapes');const png=await w.render({sheetName:'Shapes',autoCrop:'all',scale:1});await fs.writeFile(path.join(out,'shapes.png'),new Uint8Array(await png.arrayBuffer()));s.deleteAllDrawings();eq(s.shapes.items.length,0);return layouts;
});
const enumRecords=(await fs.readFile(path.join(out,'enums-help.ndjson'),'utf8')).trim().split('\n').map(JSON.parse);
const types=enumRecords.find(x=>x.name==='enum.ChartType').values;
for(const type of types)await test('CH-'+type,'Chart type acceptance, export and render: '+type,async()=>{
 const [w,s]=fresh('Chart');s.getRange('A1:D5').values=[['Category','Series A','Series B','Size'],['Alpha',10,12,5],['Beta',20,18,8],['Gamma',15,22,10],['Delta',25,27,12]];
 const c=s.charts.add(type,s.getRange('A1:C5'));c.title=type;c.setPosition('F2','N16');
 c.titleTextStyle.typeface='Arial';c.titleTextStyle.fontSize=16;c.legend={position:'bottom',textStyle:{typeface:'Arial',fontSize:11}};
 c.xAxis={textStyle:{typeface:'Arial',fontSize:10}};c.yAxis={min:0,numberFormatCode:'0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:10}};
 await xlsx(w,'chart-'+type);let render;try{const png=await w.render({sheetName:'Chart',range:'F2:N16',scale:1});await fs.writeFile(path.join(out,'chart-'+type+'.png'),new Uint8Array(await png.arrayBuffer()));render={bytes:png.size};}catch(e){render={error:String(e)};}
 return {type:c.type,series:c.series.items.map(x=>({name:x.name,formula:x.formula,categoryFormula:x.categoryFormula})),render};
});
await test('F08','Charts range/nonadjacent/row binding, setData, axes, labels, legend, line and data table',async()=>{
 const [w,s]=fresh('Charts');s.getRange('A1:D4').values=[['Month','Revenue','Profit','Cost'],['Jan',100,10,90],['Feb',120,20,100],['Mar',150,30,120]];
 const c=s.charts.add('line',[s.getRange('A1:A4'),s.getRange('B1:B4'),s.getRange('C1:C4')]);c.title='Revenue and profit';c.setPosition('F2','O18');
 c.legend={position:'top',overlay:false,textStyle:{typeface:'Arial',fontSize:12}};c.titleTextStyle.typeface='Arial';c.titleTextStyle.fontSize=16;
 c.xAxis={axisType:'textAxis',title:{text:'Month'},textStyle:{typeface:'Arial',fontSize:11}};
 c.yAxis={min:0,max:200,majorUnit:50,title:{text:'USD'},numberFormatCode:'$0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:11}};
 c.dataLabels={showValue:true,position:'outEnd'};c.dataTable={visible:true,showLegendKey:true};
 c.setData(s.getRange('A1:C4'));c.series.items[0].line={fill:'#2563eb',style:'solid',width:2};c.series.items[1].line={fill:'#ff0000',style:'dashed',width:2};
 s.getRange('B2').values=[[200]];w.recalculate();const before=c.series.items.map(x=>({formula:x.formula,values:x.values}));
 await xlsx(w,'chart-settings');const png=await w.render({sheetName:'Charts',range:'A1:O18',scale:1});await fs.writeFile(path.join(out,'chart-settings.png'),new Uint8Array(await png.arrayBuffer()));
 let image;try{image=await w.chartToImage(0,c.id);await save('chart-to-image.json',image);}catch(e){image=String(e);}
 return {title:c.title.text,before,chartImage:image&&typeof image==='object'?Object.keys(image):image};
});
await test('F09','Raster export selection and output variants',async()=>{
 const [w,s]=fresh();s.getRange('A1:C3').values=[['Label','Value','Double'],['A',2,'=B2*2'],['B',4,'=B3*2']];
 const outputs=[];for(const [name,options]of [
 ['png-index',{sheetIndex:0,range:'A1:C3',headers:true,format:'png',scale:2}],
 ['jpeg-center',{sheet:s,center:'B2',width:400,height:200,format:'jpeg',quality:0.8}],
 ['layout',{sheet:0,range:'A1:C3',format:'layout'}],
 ['xlsx',{format:'xlsx'}]
 ]){const blob=await w.export(options);const bytes=new Uint8Array(await blob.arrayBuffer());await fs.writeFile(path.join(out,'export-'+name+(name==='layout'?'.json':name==='xlsx'?'.xlsx':name.startsWith('jpeg')?'.jpg':'.png')),bytes);outputs.push({name,type:blob.type,size:bytes.length});}return outputs;
});
}

if(mode==='ops'){
 const entries=(await fs.readFile(path.join(out,'api-help.ndjson'),'utf8')).trim().split('\n').map(JSON.parse).filter(x=>x.kind==='op');
 const imageBytes=await fs.readFile(path.join(out,'formatting.png')),dataUrl='data:image/png;base64,'+imageBytes.toString('base64');
 for(const entry of entries)for(const [ix,example]of entry.examples.entries())await test('OP-'+entry.op+'-'+ix,entry.summary,async()=>{
  const w=Workbook.create();for(const name of ['Sheet1','Inventory','Scratch','Report','Scorecard','CF','KPIs','Validation','Charts','Shapes','Images','Data']){
  const s=w.worksheets.add(name);s.getRange('A1:C4').values=[['Name','Qty','Stock'],['Alpha',2,10],['Beta',3,20],['Gamma',4,30]];}
  w.comments.setSelf({displayName:'Ruggero Gargiulo'});const op=structuredClone(example);const adaptations=[];
  if(/^table\.(rows.add|set|remove)$/.test(op.op))w.worksheets.getItem('Inventory').tables.add('A1:C3',true,'InventoryTable');
  if(op.op==='range.unmerge')w.worksheets.getItem('Report').getRange('A1:D1').merge();
  if(op.op==='range.format.clear')w.worksheets.getItem('Scorecard').getRange('A2:C8').format.fill='#ff0000';
  if(op.op==='conditionalformat.clear')w.worksheets.getItem('CF').getRange('A1:A3').conditionalFormats.add('cellIs',{operator:'greaterThan',formula:1,format:{fill:'#ff0000'}});
  if(op.op==='datavalidation.clear')w.worksheets.getItem('Validation').getRange('B2:B10').dataValidation={rule:{type:'whole',formula1:1,operator:'greaterThan'}};
  if(/^sparkline\.(set|remove)$/.test(op.op))w.apply([entries.find(x=>x.op==='sparkline.add').examples[0]]);
  if(/^chart\.(set|remove)$/.test(op.op)){const c=w.worksheets.getItem('Charts').charts.add('line',{from:{row:1,col:1},extent:{widthPx:300,heightPx:200}});c.title='Updated chart';c.categories=['A','B'];const se=c.series.add('Values');se.values=[1,2];se.categories=['A','B'];}
  if(/^shape\.(set|remove)$/.test(op.op))w.worksheets.getItem('Shapes').shapes.add({geometry:'rect',anchor:{from:{row:1,col:1},extent:{widthPx:50,heightPx:50}},fill:'#ff0000'});
  if(op.op==='image.add'){op.props.dataUrl=dataUrl;adaptations.push('Replaced placeholder data URL with local synthetic PNG.');}
  if(/^image\.(set|remove)$/.test(op.op)){const image=w.worksheets.getItem('Images').images.add({dataUrl,anchor:op.target.anchor});op.target.imageId=image.imageId??w.toProto().images[0]?.id;adaptations.push('Resolved actual image asset ID.');}
  if(op.op==='names.remove'){if(op.sheet)w.worksheets.getItem(op.sheet).names.addRange(op.name,op.sheet+'!A1:A3');else w.names.addRange(op.name,'Data!A1:A3');}
  if(/^thread\.(reply|resolve|reopen|remove)$/.test(op.op)){const th=w.comments.addThread({cell:w.worksheets.getItem('Sheet1').getRange('F5')},'Original');if(op.op==='thread.reopen')th.resolve();op.target=op.op==='thread.remove'?'th/'+th.id:th.id;adaptations.push('Resolved actual thread ID.');}
  const before=JSON.stringify(w.toProto());const result=await w.apply([op]);w.recalculate();const changed=before!==JSON.stringify(w.toProto());
  await save('op-'+entry.op+'-'+ix+'.json',{input:op,adaptations,result,changed,after:w.toProto()});
  return {input:op,adaptations,result,changed};
 });
 await test('OP-invalid','Unknown action handling and mutation atomicity',()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];let error;try{w.apply([{op:'range.values.set',target:{sheet:'Data',range:'A1'},values:[[2]]},{op:'invalid.action'}]);}catch(e){error=String(e);}return {error,valueAfterFailure:val(s,'A1')};
 });
}

if(mode==='oracles'){
 const checks=[
 ['SUM','=SUM(A1:A5)',15],['PRODUCT','=PRODUCT(A1:A5)',120],['SUMPRODUCT','=SUMPRODUCT(A1:A5,B1:B5)',550],
 ['SUMIFS','=SUMIFS(B1:B5,A1:A5,">2")',120],['COUNTIFS','=COUNTIFS(A1:A5,">2")',3],['AVERAGEIFS','=AVERAGEIFS(B1:B5,A1:A5,">2")',40],
 ['ROUND','=ROUND(12.345,2)',12.35],['MOD','=MOD(17,5)',2],['POWER','=POWER(2,8)',256],['SQRT','=SQRT(81)',9],
 ['CEILING.MATH','=CEILING.MATH(4.2)',5],['FLOOR.MATH','=FLOOR.MATH(4.8)',4],['ABS','=ABS(-4)',4],
 ['AVERAGE','=AVERAGE(A1:A5)',3],['MEDIAN','=MEDIAN(A1:A5)',3],['MIN','=MIN(A1:A5)',1],['MAX','=MAX(A1:A5)',5],
 ['STDEV.S','=STDEV.S(A1:A5)',Math.sqrt(2.5)],['VAR.P','=VAR.P(A1:A5)',2],['CORREL','=CORREL(A1:A5,B1:B5)',1],
 ['NORM.DIST','=NORM.DIST(0,0,1,TRUE)',0.5],['PERCENTILE.INC','=PERCENTILE.INC(A1:A5,0.5)',3],
 ['IF','=IF(A1=1,"yes","no")','yes'],['AND','=AND(A1=1,A2=2)',true],['OR','=OR(FALSE,TRUE)',true],
 ['IFERROR','=IFERROR(1/0,99)',99],['IFNA','=IFNA(NA(),99)',99],['CHOOSE','=CHOOSE(2,10,20,30)',20],
 ['SWITCH','=SWITCH(2,1,10,2,20,0)',20],['IFS-prefixed','=_xlfn.IFS(A1=1,10,TRUE,20)',10],
 ['INDEX','=INDEX(B1:B5,3)',30],['MATCH','=MATCH(3,A1:A5,0)',3],['VLOOKUP','=VLOOKUP(3,A1:B5,2,FALSE)',30],
 ['XLOOKUP','=XLOOKUP(3,A1:A5,B1:B5)',30],['XLOOKUP-prefixed','=_xlfn.XLOOKUP(3,A1:A5,B1:B5)',30],
 ['OFFSET','=OFFSET(A1,1,1)',20],['INDIRECT','=INDIRECT("B3")',30],['ADDRESS','=ADDRESS(3,2)','$B$3'],
 ['FORMULATEXT','=FORMULATEXT(H1)','=A1*2'],['ISFORMULA','=ISFORMULA(H1)',true],['ISNUMBER','=ISNUMBER(B1)',true],
 ['ISBLANK','=ISBLANK(J1)',true],['ISERROR','=ISERROR(1/0)',true],['TYPE','=TYPE("x")',2],
 ['LEN','=LEN("hello")',5],['LEFT','=LEFT("hello",2)','he'],['MID','=MID("hello",2,3)','ell'],
 ['TRIM','=TRIM("  a  b  ")','a b'],['SUBSTITUTE','=SUBSTITUTE("a-b","-","/")','a/b'],
 ['TEXTJOIN-prefixed','=_xlfn.TEXTJOIN(", ",TRUE,"a","b")','a, b'],['TEXT','=TEXT(0.125,"0.0%")','12.5%'],
 ['DATE','=DATE(2026,1,1)',46023],['YEAR','=YEAR(DATE(2026,1,1))',2026],['MONTH','=MONTH(DATE(2026,9,15))',9],
 ['EOMONTH','=DAY(EOMONTH(DATE(2026,2,2),0))',28],['EDATE','=MONTH(EDATE(DATE(2026,1,1),2))',3],
 ['NETWORKDAYS','=NETWORKDAYS(DATE(2026,9,14),DATE(2026,9,18))',5],['DAYS','=DAYS(DATE(2026,2,1),DATE(2026,1,1))',31],
 ['PMT','=PMT(0,10,1000)',-100],['PV','=PV(0,10,-100)',1000],['FV','=FV(0,10,-100)',1000],
 ['NPV','=NPV(0.1,100,100)',100/1.1+100/1.1**2],['IRR','=IRR({-100,110})',0.1],
 ['XIRR','=XIRR({-100,110},{DATE(2025,1,1),DATE(2026,1,1)})',0.1],
 ['RATE','=RATE(1,0,-100,110)',0.1],['NPER','=NPER(0,-10,100)',10],['SLN','=SLN(100,10,9)',10],
 ['CONVERT','=CONVERT(1,"m","cm")',100],['DEC2BIN','=DEC2BIN(10)','1010'],['BIN2DEC','=BIN2DEC("1010")',10],
 ['IMABS','=IMABS("3+4i")',5],['IMSUM','=IMSUM("1+2i","3+4i")','4+6i'],
 ['LET','=LET(x,2,x*3)',6],['LET-prefixed','=_xlfn.LET(_xlpm.x,2,_xlpm.x*3)',6],
 ['LAMBDA-call','=_xlfn.LAMBDA(_xlpm.x,_xlpm.x*3)(2)',6],
 ['COUNTBLANK','=COUNTBLANK(J1:J5)',5],['COUNTIF-blanks','=COUNTIF(J1:J5,"")',5],
 ];
 const [w,s]=fresh('Oracle');s.getRange('A1:B5').values=[[1,10],[2,20],[3,30],[4,40],[5,50]];s.getRange('H1').formulas=[['=A1*2']];
 s.getRange('L1:O1').values=[['Function','Calculated','Expected','Probe']];
 for(const [i,[name,formula,expected]]of checks.entries())await test('FX-'+name,'Independent fixture: '+formula,()=>{
  const r=i+2;s.getRange('L'+r).values=[[name]];s.getRange('M'+r).formulas=[[formula]];s.getRange('N'+r).values=[[expected]];w.recalculate();
  const actual=val(s,'M'+r);s.getRange('O'+r).values=[[typeof actual===typeof expected&&(typeof expected==='number'?Math.abs(actual-expected)<1e-7:actual===expected)?'MATCH':'DIFFERENCE']];
  return {formula,expected,actual,matches:val(s,'O'+r)==='MATCH'};
 });
 await xlsx(w,'formula-oracles');
}

if(mode==='followup'){
await test('R01','FillFrom correct adjacent extension and table totals boundary',async()=>{
 const [w,s]=fresh();s.getRange('A1:B3').values=[[1,2],[3,4],[5,6]];s.getRange('C1').formulas=[['=A1+B1']];s.getRange('C1:C3').fillFrom(s.getRange('C1'));eq(s.getRange('C1:C3').values,[[3],[7],[11]]);
 const t=s.tables.add('E1:F3',true,'Stock');s.getRange('E1:F3').values=[['Item','Qty'],['A',2],['B',3]];t.rows.add(null,[['C',4]]);
 const before={range:t.address,rows:t.getDataRows()};t.showTotals=true;const after={range:t.address,rows:t.getDataRows(),grid:t.getRange().values};
 let insertError;try{t.rows.add(1,[['Inserted',5]]);}catch(e){insertError=String(e);}
 await xlsx(w,'table-totals-boundary');return {fill:s.getRange('C1:C3').formulas,before,after,insertError,columnsAPI:typeof t.columns,filterAPI:typeof t.autoFilter};
});
await test('R02','Named function formula-versus-lambda configuration',async()=>{
 const [w,s]=fresh();const variants=[
 ['Plain',{lambda:'LAMBDA(amount,amount*1.1)'}],
 ['Prefixed',{lambda:'_xlfn.LAMBDA(_xlpm.amount,_xlpm.amount*1.1)'}],
 ['Equals',{lambda:'=_xlfn.LAMBDA(_xlpm.amount,_xlpm.amount*1.1)'}]];
 const results=[];for(const [i,[name,config]]of variants.entries()){w.names.addFunction(name,config);s.getCell(i,0).formulas=[['='+name+'(100)']];w.recalculate();results.push({name,config,value:val(s,'A'+(i+1))});}await xlsx(w,'named-function-variants');return results;
});
await test('R03','CSV instance import on populated workbook preserves or replaces data',async()=>{
 const [w,s]=fresh('Original');s.getRange('A1').values=[['Keep']];const result=await w.fromCSV('A,B\n1,2',{sheetName:'Added'});return {sheets:w.worksheets.items.map(x=>x.name),originalExists:!w.worksheets.getItemOrNullObject('Original').isNullObject,originalCell:val(s,'A1'),resultSheet:result.sheet.name};
});
await test('R04','One-variable row and column data tables; export and roundtrip',async()=>{
 const [w,s]=fresh('OneVariable');s.getRange('B2').values=[[2]];s.getRange('E4').formulas=[['=B2*10']];s.getRange('F4:H4').values=[[1,2,3]];s.dataTables.add('E4:H5',{rowInput:'B2'});
 s.getRange('E8').formulas=[['=B2*10']];s.getRange('E9:E11').values=[[1],[2],[3]];s.dataTables.add('E8:F11',{columnInput:'B2'});w.recalculate();
 eq(s.getRange('F5:H5').values,[[10,20,30]]);eq(s.getRange('F9:F11').values,[[10],[20],[30]]);await xlsx(w,'one-variable-data-tables');
 const round=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,'one-variable-data-tables.xlsx')));round.recalculate();return {row:s.getRange('F5:H5').values,column:s.getRange('F9:F11').values,reimport:round.worksheets.getItemAt(0).getRange('E4:H5').values};
});
await test('R05','Sparklines supported delete route and export preservation',async()=>{
 const [w,s]=fresh();s.getRange('A1:D4').values=[['Q1','Q2','Q3','Q4'],[1,2,-1,4],[3,2,4,5],[4,null,2,3]];
 const groups=[];for(const [i,type]of ['line','column','stacked'].entries()){const col=w.utils.columnToLetter(i+6);groups.push(s.sparklineGroups.add({type,targetRange:col+'2:'+col+'4',sourceData:'A2:D4',dateAxisRange:'A1:D1',axis:{manualMin:-5,manualMax:10,showAxis:true},markers:{show:true,high:true,low:true,negative:true},seriesColor:'#2563eb'}));}
 const tmp=s.getRange('J2:J4').sparklines.add('line',s.getRange('A2:D4'));s.sparklineGroups.delete(tmp);await xlsx(w,'sparklines');const png=await w.render({sheetName:'Data',range:'A1:I5',scale:2});await fs.writeFile(path.join(out,'sparklines.png'),new Uint8Array(await png.arrayBuffer()));return {count:s.sparklineGroups.getAll().length};
});
await test('R06','Image sources and config replacement without overwriting facade anchor',async()=>{
 const [w,s]=fresh('Images');const bytes=await fs.readFile(path.join(out,'formatting.png'));const svg='<svg xmlns="http://www.w3.org/2000/svg" width="80" height="40"><rect width="80" height="40" fill="#2563eb"/></svg>';
 const configs=[{svg},{blob:bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength),contentType:'image/png'},{dataUrl:'data:image/png;base64,'+bytes.toString('base64')},{path:path.join(out,'formatting.png')}];
 for(const [i,config]of configs.entries()){const im=s.images.add({...config,anchor:{from:{row:i*5,col:1},extent:{widthPx:160,heightPx:80}},alt:'Source '+i});if(i===3)im.replace({svg});}
 const imageResult=w.fromImage(0,{bytes,contentType:'image/png'},'F1:H7');await xlsx(w,'images');
 const png=await w.render({sheetName:'Images',range:'A1:I21',scale:1});await fs.writeFile(path.join(out,'images.png'),new Uint8Array(await png.arrayBuffer()));
 return {count:s.images.items.length,fromImage:imageResult,types:s.images.items.map(i=>i.contentType)};
});
await test('R07','Chart configuration at creation, mutable facade settings and analytics',async()=>{
 const [w,s]=fresh('Charts');s.getRange('A1:C5').values=[['Month','Revenue','Profit'],['Jan',100,10],['Feb',120,20],['Mar',150,30],['Apr',170,35]];
 const c=s.charts.add('line',s.getRange('A1:C5'));c.title='Synthetic revenue and profit';c.setPosition('E2','N19');
 c.titleTextStyle.typeface='Arial';c.titleTextStyle.fontSize=16;c.legend={position:'top',textStyle:{typeface:'Arial',fontSize:12}};
 c.xAxis={title:'Month',textStyle:{typeface:'Arial',fontSize:10}};c.yAxis={min:0,max:250,majorUnit:50,title:'USD',numberFormatCode:'$0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:10}};
 c.dataLabels.showValue=true;c.dataLabels.position='outEnd';c.dataTable.visible=true;c.dataTable.showLegendKey=true;
 const se=c.series.items[0];se.line={fill:'#2563eb',style:'solid',width:2};se.valuesFormatCode='$0';
 se.errorBars={type:'percentage',value:5};se.trendlines.add('linear',{displayEquation:true,displayRSquared:true,forecastForward:1,line:{fill:'#999999',style:'dashed',width:1}});
 se.dataLabelOverrides.add(0).showValue=true;c.series.items[1].line={fill:'#ff0000',style:'dashed',width:2};
 s.getRange('B2').values=[[200]];w.recalculate();eq(se.resolveValues()[0],200);await xlsx(w,'chart-settings');
 const png=await w.render({sheetName:'Charts',range:'A1:N19',scale:1});await fs.writeFile(path.join(out,'chart-settings.png'),new Uint8Array(await png.arrayBuffer()));
 const converted=await w.chartToImage(0,c.id);await save('chart-to-image.json',converted);return {values:se.resolveValues(),imageResult:converted,formula:se.formula};
});
await test('R08','Chart dataLabels and dataTable config-first API',async()=>{
 const [w,s]=fresh();const c=s.charts.add('bar',{from:{row:1,col:1},extent:{widthPx:400,heightPx:240},categories:['A','B'],series:[{name:'Value',values:[10,20]}],dataLabels:{showValue:true},dataTable:{visible:true,showLegendKey:true}});
 c.barOptions.direction='column';c.barOptions.grouping='stacked';c.barOptions.gapWidth=120;c.barOptions.overlap=0;await xlsx(w,'chart-config-first');return {labels:c.dataLabels.showValue,dataTable:c.dataTable.visible};
});
await test('R09','Pivots and slicers, hierarchy operations and native export',async()=>{
 const [w,s]=fresh('Data');const summary=w.worksheets.add('Pivot');s.getRange('A1:C5').values=[['Region','Product','Sales'],['North','A',10],['South','A',20],['North','B',30],['South','B',40]];
 const p=summary.pivotTables.add('SalesPivot',s.getRange('A1:C5'),summary.getRange('A1'));const hier=p.hierarchies.items.map(x=>x.name);
 p.rowHierarchies.add(p.hierarchies.getItem('Region'));p.columnHierarchies.add(p.hierarchies.getItem('Product'));p.dataHierarchies.add(p.hierarchies.getItem('Sales'));w.recalculate();
 const slicer=summary.slicers.add('SalesPivot','Region');await save('pivot-proto.json',w.toProto());await xlsx(w,'pivot-slicer');
 return {hier,grid:summary.getRange('A1:E6').values,pivotKeys:Object.keys(p),slicerKeys:Object.keys(slicer),slicerProperties:Object.getOwnPropertyNames(Object.getPrototypeOf(slicer))};
});
await test('R10','Notes operation route and low-level note persistence',async()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];w.comments.setSelf({displayName:'Ruggero Gargiulo'});
 let applyError;try{w.apply([{op:'note.add',target:{cell:{sheet:'Data',address:'A1'}},body:'Synthetic note'}]);}catch(e){applyError=String(e);}
 const note=w.notes.add({id:'probe-note',sheetId:s.sheetId,ref:'A1',text:'Synthetic note'});await save('note-proto.json',w.toProto());let exported;
 try{await xlsx(w,'notes');exported=true;}catch(e){exported=String(e);}note.delete();return {applyError,exported};
});
await test('R11','Undo/redo, recordAsync and awareness selections',async()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];w.hydrateCrdtFromProto();const recorded=await w.recordAsync(async()=>{s.getRange('A1').values=[[2]];await Promise.resolve();s.getRange('B1').formulas=[['=A1*2']];return 7;});
 const before={canUndo:w.canUndo(),canRedo:w.canRedo(),value:val(s,'A1')};w.undo();const undone={canRedo:w.canRedo(),value:val(s,'A1')};w.redo();const redone=val(s,'A1');
 w.awareness.setPresenceSelections('synthetic-reviewer',[{sheetName:'Data',range:'A1:B2'}],{kind:'collaborator'});const selections=w.awareness.getSelectionsForSheet('Data');w.awareness.clearPresenceSelections('synthetic-reviewer');
 return {result:recorded.result,patch:recorded.patch,before,undone,redone,selections};
});
}

if(mode==='advanced'){
await test('A01','Pivot root collection, native serialization and slicer fields',async()=>{
 const [w,s]=fresh('Data');const summary=w.worksheets.add('Pivot');s.getRange('A1:C5').values=[['Region','Product','Sales'],['North','A',10],['South','A',20],['North','B',30],['South','B',40]];
 const p=w.pivotTables.add('SalesPivot',s.getRange('A1:C5'),summary.getRange('A1'));
 p.rowHierarchies.add(p.hierarchies.getItem('Region'));p.columnHierarchies.add(p.hierarchies.getItem('Product'));p.dataHierarchies.add(p.hierarchies.getItem('Sales'));w.recalculate();
 const detail={grid:summary.getRange('A1:E7').values,hierarchies:p.hierarchies.items.map(x=>x.name),methods:Object.getOwnPropertyNames(Object.getPrototypeOf(p))};
 await save('pivot-proto.json',w.toProto());try{await xlsx(w,'pivot-only');detail.pivotExport='completed';}catch(e){detail.pivotExportError=String(e);}
 try{const sl=summary.slicers.add('SalesPivot','Region');detail.slicerMethods=Object.getOwnPropertyNames(Object.getPrototypeOf(sl));detail.slicerItems=sl.slicerItems?.items;await xlsx(w,'pivot-slicer');detail.slicerExport='completed';}catch(e){detail.slicerError=String(e);}
 await save('pivot-details.json',detail);return detail;
});
await test('A02','Chart analytics and mutable label/table objects',async()=>{
 const [w,s]=fresh('Charts');s.getRange('A1:C5').values=[['Month','Revenue','Profit'],['Jan',100,10],['Feb',120,20],['Mar',150,30],['Apr',170,35]];
 const c=s.charts.add('line',s.getRange('A1:C5'));c.title='Synthetic revenue and profit';c.setPosition('E2','N19');
 c.titleTextStyle.typeface='Arial';c.titleTextStyle.fontSize=16;c.legend={position:'top',textStyle:{typeface:'Arial',fontSize:12}};
 c.xAxis={title:'Month',textStyle:{typeface:'Arial',fontSize:10}};c.yAxis={min:0,max:250,majorUnit:50,title:'USD',numberFormatCode:'$0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:10}};
 c.dataLabels.showValue=true;c.dataLabels.position='outEnd';c.dataTable.visible=true;c.dataTable.showLegendKey=true;
 const se=c.series.items[0];se.line={fill:'#2563eb',style:'solid',width:2};se.errorBars={type:'percentage',value:5};
 se.trendlines.add('linear',{displayEquation:true,displayRSquared:true,forecastForward:1,line:{fill:'#999999',style:'dashed',width:1}});
 se.dataLabelOverrides.add(0).showValue=true;c.series.items[1].line={fill:'#ff0000',style:'dashed',width:2};s.getRange('B2').values=[[200]];w.recalculate();
 const detail={resolvedValues:await se.resolveValues(),formula:se.formula,dataTable:c.dataTable.visible,chartCount:s.charts.items.length};
 await xlsx(w,'chart-settings');const png=await w.render({sheetName:'Charts',range:'A1:N19',scale:1});await fs.writeFile(path.join(out,'chart-settings.png'),new Uint8Array(await png.arrayBuffer()));
 try{detail.converted=await w.chartToImage(0,c.id);detail.chartCountAfter=s.charts.items.length;detail.imageCountAfter=s.images.items.length;}catch(e){detail.convertError=String(e);}return detail;
});
await test('A03','Native note through operation API and note delete',async()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];w.comments.setSelf({displayName:'Ruggero Gargiulo'});
 w.apply([{op:'note.add',target:{cell:{sheet:'Data',address:'A1'}},body:'Synthetic cell note'}]);
 const note=w.notes.items[0];const proto=note.proto;await xlsx(w,'native-note');note.delete();eq(w.notes.items.length,0);return proto;
});
await test('A04','Awareness operation families and correct presence rectangle',async()=>{
 const [w,s]=fresh();const result=w.apply([{op:'presence.selection.set',presenceId:'probe',selections:[{sheetName:'Data',range:'A1:B2'}]}]);const selected=w.awareness.getSelectionsForSheet('Data');
 w.apply([{op:'presence.selection.clear',presenceId:'probe'}]);w.comments.setSelf({displayName:'Ruggero Gargiulo'});const th=w.comments.addThread({cell:s.getRange('A1')},'Probe');
 const reaction=w.apply([{op:'thread.reaction.toggle',target:th.id,reaction:'👍'}]);return {result,selected,reaction};
});
await test('A05','Workbook static session API and common CRDT initialization',async()=>{
 const statics=Object.getOwnPropertyNames(Workbook.session??{});const methods={};for(const key of statics)if(typeof Workbook.session[key]==='function')methods[key]=Workbook.session[key].toString();
 const [w,s]=fresh();return {statics,methods,crdt:{loadInitial:w.loadInitialCrdtStateV2.toString(),subscribe:w.onCrdtUpdateV2.toString(),getDoc:w.getCrdtDoc.toString()}};
});
await test('A06','Database formula family with valid database and criteria',async()=>{
 const [w,s]=fresh();s.getRange('A1:C5').values=[['Region','Product','Sales'],['North','A',10],['South','A',20],['North','B',30],['South','B',40]];s.getRange('E1:E2').values=[['Region'],['North']];
 const cases=[['DSUM',40],['DCOUNT',2],['DCOUNTA',2],['DAVERAGE',20],['DMIN',10],['DMAX',30],['DPRODUCT',300],['DSTDEV',Math.sqrt(200)],['DVAR',200],['DSTDEVP',10],['DVARP',100],['DGET','#NUM!']];
 const data=[];for(const [i,[fn,expected]]of cases.entries()){const formula='='+fn+'(A1:C5,"Sales",E1:E2)';s.getCell(i,6).formulas=[[formula]];w.recalculate();data.push({formula,expected,actual:val(s,'G'+(i+1))});}return data;
});
await test('A07','Table total reserve and recalculation before/after flag',async()=>{
 const [w,s]=fresh();s.getRange('A1:B4').values=[['Item','Qty'],['A',2],['B',3],['C',4]];const t=s.tables.add('A1:B4',true,'Stock');
 s.getRange('D1').formulas=[['=SUM(Stock[Qty])']];const before=val(s,'D1');t.showTotals=true;w.recalculate();const after=val(s,'D1');
 const [x,u]=fresh();u.getRange('A1:B5').values=[['Item','Qty'],['A',2],['B',3],['C',4],['Total',null]];const safe=u.tables.add('A1:B5',true,'Reserved');safe.showTotals=true;u.getRange('D1').formulas=[['=SUM(Reserved[Qty])']];x.recalculate();
 return {before,after,withReservedTotalsRow:val(u,'D1')};
});
}

if(mode==='recovery'){
await test('V01','Replicated CRDT edits from a shared initialization update',async()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];const baseline=[];const off=w.onCrdtUpdateV2(u=>baseline.push(u));w.hydrateCrdtFromProto();off();
 const clone=Workbook.create();for(const update of baseline)clone.applyCrdtUpdateV2(update,{recalculate:true});
 const change=w.record(()=>{s.getRange('A1').values=[[99]];});clone.applyCrdtUpdateV2(change.crdtUpdateV2,{recalculate:true});
 const target=clone.worksheets.getItemOrNullObject('Data');return {baselineUpdates:baseline.map(u=>u.length),sheets:clone.worksheets.items.map(x=>x.name),value:target.isNullObject?null:val(target,'A1')};
});
await test('V02','Correct raw-note config and Excel note preservation',async()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];w.comments.setSelf({displayName:'Ruggero Gargiulo'});
 const note=w.notes.add({id:'synthetic-note',target:{cell:{sheetName:s.name,sheetId:s.sheetId,address:'A1'}},authorId:w.comments.self.id,createdAt:new Date().toISOString(),body:{plainText:'Synthetic cell note'}});
 await xlsx(w,'native-note');const proto=note.proto;note.delete();eq(w.notes.items.length,0);return proto;
});
await test('V03','Unsupported operations return warnings; batches retain earlier changes',()=>{
 const [w,s]=fresh();s.getRange('A1').values=[[1]];const result=w.apply([{op:'range.values.set',target:{sheet:'Data',range:'A1'},values:[[2]]},{op:'invalid.action'},{op:'note.add',body:'Probe',target:{cell:{sheet:'Data',address:'A1'}}}]);return {result,value:val(s,'A1'),notes:w.notes.items.length};
});
await test('V04','Fresh structured formula after table totals flag and saved-file reload',async()=>{
 const [w,s]=fresh();s.getRange('A1:B4').values=[['Item','Qty'],['A',2],['B',3],['C',4]];const t=s.tables.add('A1:B4',true,'Stock');
 s.getRange('D1').formulas=[['=SUM(Stock[Qty])']];const before=val(s,'D1');t.showTotals=true;s.getRange('D2').formulas=[['=SUM(Stock[Qty])+0']];w.recalculate();const after=s.getRange('D1:D2').values;await xlsx(w,'table-total-fresh-formula');
 const x=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,'table-total-fresh-formula.xlsx')));x.recalculate();return {before,after,reimport:x.worksheets.getItemAt(0).getRange('D1:D2').values};
});
await test('V05','Slicer select/clear, pivot refresh, layout and export',async()=>{
 const [w,s]=fresh('Data');const t=w.worksheets.add('Pivot');s.getRange('A1:B5').values=[['Region','Sales'],['North',10],['South',20],['North',30],['South',40]];
 const p=w.pivotTables.add('Regional',s.getRange('A1:B5'),t.getRange('A1'));p.rowHierarchies.add(p.hierarchies.getItem('Region'));p.dataHierarchies.add(p.hierarchies.getItem('Sales'));
 const sl=t.slicers.add('Regional','Region');sl.name='RegionSelector';sl.caption='Region';sl.left=320;sl.top=20;sl.width=180;sl.height=180;
 const before=t.getRange('A1:C5').values;sl.selectItems(['North']);w.recalculate();const selected=t.getRange('A1:C5').values;sl.clearFilters();s.getRange('B2').values=[[100]];p.rebuildCache();w.recalculate();const refreshed=t.getRange('A1:C5').values;
 await xlsx(w,'slicer-filter');return {before,selected,refreshed,layout:p.layout.layoutType};
});
await test('V06','Read-only public static session wrapper types and file APIs',()=>{
 const getters={};for(const key of Object.getOwnPropertyNames(Workbook)){if(['prototype','length','name'].includes(key))continue;getters[key]={type:typeof Workbook[key],source:typeof Workbook[key]==='function'?Workbook[key].toString().slice(0,5000):undefined};}
 return {Workbook:getters,SpreadsheetFile:Object.getOwnPropertyNames(SpreadsheetFile),FileBlob:Object.getOwnPropertyNames(FileBlob)};
});
}

await save(runName+'.json',receipts);

import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import {fileURLToPath} from 'node:url';
import {Workbook, SpreadsheetFile, FileBlob} from '@oai/artifact-tool';

const here=path.dirname(fileURLToPath(import.meta.url));
const dir=path.join(here,'operations-run');
await fs.mkdir(dir,{recursive:true});
const source=path.join(here,'man-group-run/source.xlsx');
const out=path.join(dir,'man-group-operations.xlsx');
const sha=async p=>crypto.createHash('sha256').update(await fs.readFile(p)).digest('hex');
const sourceHash=await sha(source);
const records=[];
let wb,lab,tableObject,chart,scratch,small,lastHandle,reopenBook,reopenSheet;
const cell=(sheet,address)=>({address,value:sheet.getRange(address).values[0][0],formula:sheet.getRange(address).formulas[0][0]});
const grid=(sheet,address)=>({address,values:sheet.getRange(address).values,formulas:sheet.getRange(address).formulas});
const labRead=()=>({margin:cell(lab,'C4'),growth:cell(lab,'C5'),quarters:grid(lab,'B9:F12'),annual:cell(lab,'C15')});
const inspect=async (book,args)=>{const r=await book.inspect(args);return {ndjson:r.ndjson,recordCount:r.recordCount,truncated:r.truncated,metadata:r.metadata};};
const names=async book=>await inspect(book,{kind:'sheet',include:'id,name',maxChars:5000});
const serialize=v=>v===undefined?{javascriptType:'undefined'}:v;
const encode=v=>JSON.stringify(v,(_key,x)=>{if(x===undefined)return {javascriptType:'undefined'};if(x && typeof x.ndjson==='string')return {ndjson:x.ndjson,recordCount:x.recordCount,truncated:x.truncated,metadata:x.metadata};return x;},2);
async function run(id,operation,{before='',after='',handle=''}={}) {
  const evaluate=code=>eval('(async()=>{'+code+'\n})()');
  const r={id,operation,checkInput:{before,after},startedAt:new Date().toISOString()};
  if(before)try{r.before=await evaluate(before);}catch(e){r.beforeError=String(e.message);}
  const t=performance.now();
  try {const value=await evaluate(operation);lastHandle=value;r.libraryReturn=handle?{type:handle,representation:'Object contents are read separately in the agent checks.'}:serialize(value);}
  catch(e){r.operationError={name:e.name,message:e.message};}
  r.elapsedMs=Math.round(performance.now()-t);
  if(after)try{r.after=await evaluate(after);}catch(e){r.afterError=String(e.message);}
  records.push(r);
  await fs.writeFile(path.join(dir,id+'.json'),encode(r));
  await fs.writeFile(path.join(dir,'records.json'),encode(records));
  console.log(JSON.stringify({id,elapsedMs:r.elapsedMs,error:r.operationError||r.afterError||null}));
  return r;
}

await run('new-workbook','scratch = Workbook.create();\nsmall = scratch.worksheets.add("Fee calculation");\nreturn small;',{
  handle:'Worksheet object',after:'return await names(scratch);'});
await run('new-formula',`small.getRange("B2:C5").values = [["Average AUM, USD bn",100],["Annual fee margin, bp",50],["Quarter fraction",0.25],["Fees, USD m",null]];
small.getRange("C5").formulas = [["=C2*1000*C3/10000*C4"]];`,{
  after:'return {cells:grid(small,"B2:C5"),independentExpected:100*1000*50/10000*0.25};'});
await run('new-recalculation','small.getRange("C3").values = [[60]];',{
  before:'return cell(small,"C5");',after:'return {fees:cell(small,"C5"),independentExpected:100*1000*60/10000*0.25,explicitRecalculateCalled:false};'});
await run('import-add-sheet',`wb = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
lab = wb.worksheets.add("Fee lab");
return lab;`,{handle:'Worksheet object',after:'return await names(wb);'});
await run('add-linked-build',`lab.getRange("B2").values = [["Man Group fee laboratory"]];
lab.getRange("B3").values = [["Illustrative extension for tool tests; not a forecast"]];
lab.getRange("B4:C5").values = [["Annual fee margin, bp",56],["Quarterly AUM growth",0.02]];
lab.getRange("B8:F12").values = [["Period","Average AUM, USD bn","Fee margin, bp","Year fraction","Fees, USD m"],["Quarter 1",null,null,0.25,null],["Quarter 2",null,null,0.25,null],["Quarter 3",null,null,0.25,null],["Quarter 4",null,null,0.25,null]];
lab.getRange("C9").formulas = [["='Fee engine'!W18"]];
lab.getRange("C10").formulas = [["=C9*(1+$C$5)"]];
lab.getRange("C10:C12").fillDown();
lab.getRange("D9").formulas = [["=$C$4"]];
lab.getRange("D9:D12").fillDown();
lab.getRange("F9").formulas = [["=C9*1000*D9/10000*E9"]];
lab.getRange("F9:F12").fillDown();
lab.getRange("B15").values = [["Total fees, USD m"]];
lab.getRange("C15").formulas = [["=SUM(F9:F12)"]];`,{
  after:'return {selectedReads:labRead(),sourceAum:cell(wb.worksheets.getItem("Fee engine"),"W18")};'});
await run('format-cells',`lab.showGridLines = false;
lab.getRange("B2:P28").format.font = {name:"Arial",size:10,color:"#19364C"};
lab.getRange("B2").format.font = {name:"Arial",size:14,bold:true,color:"#19364C"};
lab.getRange("B3").format.font.italic = true;
lab.getRange("B4:B5").format.columnWidth = 29;
lab.getRange("C4:C5").format.fill = "#FFF2CC";
lab.getRange("C8:F28").format.columnWidth = 19;
lab.getRange("C5").setNumberFormat("0.0%");
lab.getRange("C9:C12").setNumberFormat("0.00");
lab.getRange("D9:D12").setNumberFormat("0.0");
lab.getRange("E9:E12").setNumberFormat("0.00");
lab.getRange("F9:F12").setNumberFormat("#,##0.00");
lab.getRange("C15").setNumberFormat("#,##0.00");
lab.getRange("B8:F8").format = {fill:"#19364C",font:{name:"Arial",size:10,bold:true,color:"#FFFFFF"},rowHeight:32,wrapText:true};
lab.getRange("B15:C15").format.font.bold = true;
lab.getRange("B8:F12").format.rowHeight = 26;`,{
  after:'return {valuesUnaffected:labRead(),styles:await inspect(wb,{kind:"computedStyle",sheetId:"Fee lab",range:"C4:C5",maxChars:2500})};'});
await run('checkpoint-save',`wb.recalculate();
const file = await SpreadsheetFile.exportXlsx(wb);
return await file.save(out);`,{
  after:'return {file:path.basename(out),sha256:await sha(out),savedMargin:cell((await SpreadsheetFile.importXlsx(await FileBlob.load(out))).worksheets.getItem("Fee lab"),"C4")};'});
await run('edit-without-saving','lab.getRange("C4").values = [[51]];',{
  before:'return {diskSha256:await sha(out),memory:labRead()};',
  after:'return {diskSha256:await sha(out),memory:labRead(),onDisk:cell((await SpreadsheetFile.importXlsx(await FileBlob.load(out))).worksheets.getItem("Fee lab"),"C4")};'});
await run('clear-formulas','return lab.getRange("F9:F12").clear({applyTo:"contents"});',{
  before:'return labRead();',after:'return {selectedReads:labRead(),formatStillPresent:await inspect(wb,{kind:"computedStyle",sheetId:"Fee lab",range:"F9",maxChars:1800})};'});
await run('rebuild-formulas',`lab.getRange("F9").formulas = [["=C9*1000*D9/10000*E9"]];
return lab.getRange("F9:F12").fillDown();`,{before:'return labRead();',after:'return labRead();'});
await run('blank-input','return lab.getRange("C4").clear({applyTo:"contents"});',{
  before:'return labRead();',after:'return labRead();'});
await run('guard-missing-input',`lab.getRange("F9").formulas = [["=IF(ISBLANK($C$4),NA(),C9*1000*D9/10000*E9)"]];
lab.getRange("F9:F12").fillDown();`,{before:'return labRead();',after:'return {selectedReads:labRead(),errorScan:await inspect(wb,{kind:"match",sheetId:"Fee lab",range:"B4:F15",searchTerm:"#N/A",maxChars:2500,options:{maxResults:12}})};'});
await run('zero-input','lab.getRange("C4").values = [[0]];',{
  before:'return labRead();',after:'return labRead();'});
await run('restore-valid-input','lab.getRange("C4").values = [[51]];',{after:'return labRead();'});
await run('add-table',`lab.getRange("B20:D23").values = [["Scenario","Fee margin, bp","AUM multiplier"],["Base",56,1],["Downside",51,0.9],["Upside",60,1.1]];
lab.getRange("B19").values = [["Table example; independent of the calculation above"]];
tableObject = lab.tables.add("B20:D23",true,"FeeScenarioInputs");
lab.getRange("B20:D20").format.font.color = "#FFFFFF";
return tableObject;`,{handle:'Excel Table object',after:'return {tables:lab.tables.items.map(t=>({name:t.name,headers:t.getHeaderRowRange().values,rowReadUsesExplicitRange:true})),values:grid(lab,"B20:D23")};'});
await run('append-table-row','return tableObject.rows.add(null,[["Stress",45,0.8]]);',{
  after:'return {cells:grid(lab,"B20:D24")};'});
await run('delete-table','return tableObject.delete();',{
  before:'return {tableCount:lab.tables.items.length,cells:grid(lab,"B20:D24")};',after:'return {tableCount:lab.tables.items.length,cells:grid(lab,"B20:D24")};'});
await run('recreate-table','tableObject = lab.tables.add("B20:D24",true,"FeeScenarioInputs");\nlab.getRange("B20:D20").format.font.color = "#FFFFFF";\nreturn tableObject;',{
  handle:'Excel Table object',after:'return {tableCount:lab.tables.items.length,cells:grid(lab,"B20:D24")};'});
await run('add-chart',`chart = lab.charts.add("line",[lab.getRange("B8:B12"),lab.getRange("F8:F12")]);
chart.title = "Illustrative quarterly fees (USD m)";
chart.hasLegend = false;
chart.titleTextStyle.fontSize = 12;
chart.titleTextStyle.typeface = "Arial";
chart.yAxis = {numberFormatCode:"0.0",numberFormatSourceLinked:false,textStyle:{typeface:"Arial",fontSize:10}};
chart.xAxis = {axisType:"textAxis",textStyle:{typeface:"Arial",fontSize:10}};
chart.series.items[0].line = {fill:"#1B6870",style:"solid",width:2};
chart.setPosition("H8","P22");
return chart;`,{handle:'Chart object',after:'return {series:chart.series.items.map(s=>({formula:s.formula,categoryFormula:s.categoryFormula})),source:grid(lab,"F9:F12")};'});
await run('render-before',`return await wb.render({sheetName:"Fee lab",range:"B2:P25",scale:1.5,format:"png"});`,{
  handle:'Image Blob',after:'const blob=lastHandle;await fs.writeFile(path.join(dir,"fee-lab-before.png"),new Uint8Array(await blob.arrayBuffer()));return {saved:"fee-lab-before.png",bytes:(await fs.stat(path.join(dir,"fee-lab-before.png"))).size};'});
await run('chart-input-change','lab.getRange("C4").values = [[60]];',{
  before:'return {fees:grid(lab,"F9:F12"),bindings:chart.series.items.map(s=>s.formula)};',after:'return {fees:grid(lab,"F9:F12"),bindings:chart.series.items.map(s=>s.formula)};'});
await run('render-after',`return await wb.render({sheetName:"Fee lab",range:"B2:P25",scale:1.5,format:"png"});`,{
  handle:'Image Blob',after:'const blob=lastHandle;await fs.writeFile(path.join(dir,"fee-lab-after.png"),new Uint8Array(await blob.arrayBuffer()));return {saved:"fee-lab-after.png",bytes:(await fs.stat(path.join(dir,"fee-lab-after.png"))).size};'});
await run('delete-chart','return lab.charts.deleteAll();',{
  after:'return {drawingInspection:await inspect(wb,{kind:"drawing",sheetId:"Fee lab",maxChars:2000}),source:grid(lab,"F9:F12")};'});
await run('recreate-chart',`chart = lab.charts.add("line",[lab.getRange("B8:B12"),lab.getRange("F8:F12")]);
chart.title = "Illustrative quarterly fees (USD m)";
chart.hasLegend = false;
chart.titleTextStyle.fontSize = 12;
chart.titleTextStyle.typeface = "Arial";
chart.yAxis = {numberFormatCode:"0.0",numberFormatSourceLinked:false,textStyle:{typeface:"Arial",fontSize:10}};
chart.xAxis = {axisType:"textAxis",textStyle:{typeface:"Arial",fontSize:10}};
chart.series.items[0].line = {fill:"#1B6870",style:"solid",width:2};
chart.setPosition("H8","P22");
return chart;`,{handle:'Chart object',after:'return {series:chart.series.items.map(s=>({formula:s.formula,categoryFormula:s.categoryFormula}))};'});
await run('final-save',`wb.recalculate();
const file = await SpreadsheetFile.exportXlsx(wb);
return await file.save(out);`,{
  before:'return {diskSha256:await sha(out),memory:labRead()};',
  after:'const reopened=await SpreadsheetFile.importXlsx(await FileBlob.load(out));const sheet=reopened.worksheets.getItem("Fee lab");return {file:path.basename(out),bytes:(await fs.stat(out)).size,sha256:await sha(out),reopened:grid(sheet,"B9:F12"),reopenedMargin:cell(sheet,"C4"),reopenedTableCount:sheet.tables.items.length,sourceUnchanged:sourceHash===await sha(source),checkErrorScan:await inspect(reopened,{kind:"match",sheetId:"Fee lab",range:"B4:F24",searchTerm:"#REF!|#DIV/0!|#VALUE!|#N/A|#NAME\\\\?",options:{useRegex:true,maxResults:20},maxChars:3000})};'});
await run('reopen-and-edit',`reopenBook = await SpreadsheetFile.importXlsx(await FileBlob.load(out));
reopenSheet = reopenBook.worksheets.getItem("Fee lab");
reopenSheet.getRange("C4").values = [[55]];`,{
  after:'return {quarterOne:cell(reopenSheet,"F9"),total:cell(reopenSheet,"C15"),independentExpectedQuarterOne:reopenSheet.getRange("C9").values[0][0]*1000*55/10000*0.25,savedFileStillHasMargin:cell((await SpreadsheetFile.importXlsx(await FileBlob.load(out))).worksheets.getItem("Fee lab"),"C4")};'});
// Save the separately-created workbook, with modest formatting for readable download.
small.getRange('B2:C5').format.font={name:'Arial',size:11};
small.getRange('B2:B5').format.columnWidth=29;
small.getRange('C2:C5').format.columnWidth=18;
small.getRange('C2:C4').format.fill='#FFF2CC';
small.showGridLines=false;
scratch.recalculate();
await (await SpreadsheetFile.exportXlsx(scratch)).save(path.join(dir,'new-fee-workbook.xlsx'));
await fs.writeFile(path.join(dir,'summary.json'),JSON.stringify({sourceHash,sourceUnchanged:sourceHash===await sha(source),recordCount:records.length,failures:records.filter(r=>r.operationError||r.afterError||r.beforeError).map(r=>r.id)},null,2));

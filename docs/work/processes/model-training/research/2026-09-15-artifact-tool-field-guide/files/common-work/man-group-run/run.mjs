import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const dir = path.dirname(fileURLToPath(import.meta.url));
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(dir,'source.xlsx')));
const records = [];
async function capture(id, input) {
  const start = performance.now();
  let result, error;
  try { result = await eval('(async () => {' + input + '\n})()'); }
  catch(e) { error = String(e.stack || e); }
  if(result && typeof result.ndjson === 'string') result = {ndjson:result.ndjson, recordCount:result.recordCount, truncated:result.truncated, metadata:result.metadata};
  const record = {id, input, elapsedMs: Math.round(performance.now()-start), result: result === undefined ? {javascriptReturn:'undefined'} : result};
  if(error) record.error = error;
  records.push(record);
  await fs.writeFile(path.join(dir,id+'.json'),JSON.stringify(record,null,2));
  console.log(JSON.stringify({id,elapsedMs:record.elapsedMs,error:error||null,saved:id+'.json'}));
  return result;
}

await capture('overview', `return await wb.inspect({
  kind: "workbook,sheet,table", maxChars: 14000,
  tableMaxRows: 4, tableMaxCols: 6, tableMaxCellChars: 80
});`);
await capture('sheets', `return await wb.inspect({kind: "sheet", include: "id,name"});`);
await capture('literal', `return await wb.inspect({
  kind: "match", searchTerm: "profit", maxChars: 12000,
  options: {useRegex: false, maxResults: 30}
});`);
await capture('regex', `return await wb.inspect({
  kind: "match", searchTerm: "management.*fee|fee.*margin", maxChars: 12000,
  options: {useRegex: true, maxResults: 30}
});`);
await capture('regions', `return await wb.inspect({
  kind: "region", sheetId: "Summary", range: "C7:M28",
  maxChars: 10000, tableMaxRows: 8, tableMaxCols: 6
});`);
await capture('ranges', `const range = wb.worksheets.getItem("Summary").getRange("K13:M20");
return {sheet: "Summary", range: "K13:M20", values: range.values, formulas: range.formulas};`);
await capture('trace', `return await wb.trace("Fee engine!R27");`);
await capture('render', `const image = await wb.render({sheetName: "Summary", range: "C7:M28", scale: 1.5, format: "png"});
const bytes = new Uint8Array(await image.arrayBuffer());
await fs.writeFile(path.join(dir, "summary-before.png"), bytes);
return {saved: "summary-before.png", mimeType: "image/png", bytes: bytes.length};`);
await capture('edit-context', `const a = wb.worksheets.getItem("Assumptions");
const f = wb.worksheets.getItem("Fee engine");
return {
  assumptions: {range:"R25:W29", values:a.getRange("R25:W29").values, formulas:a.getRange("R25:W29").formulas},
  fees: {range:"R26:W28", values:f.getRange("R26:W28").values, formulas:f.getRange("R26:W28").formulas},
  output: {range:"K13:M14", values:wb.worksheets.getItem("Summary").getRange("K13:M14").values}
};`);

function readState() {
  const a=wb.worksheets.getItem('Assumptions'), f=wb.worksheets.getItem('Fee engine'), s=wb.worksheets.getItem('Summary');
  return Object.fromEntries([
    ['baseFeeMargin',a,'W26'], ['activeFeeMargin',a,'W25'],
    ['q4AverageAum',f,'W18'], ['q4ManagementFees',f,'W27'],
    ['fy2027ManagementFees',s,'M13'], ['fy2027OperatingProfit',s,'M20']
  ].map(([label,sheet,cell]) => [label,{sheet:sheet.name,cell,value:sheet.getRange(cell).values[0][0],formula:sheet.getRange(cell).formulas[0][0]}]));
}
await capture('value-edit', `const before = readState();
wb.worksheets.getItem("Assumptions").getRange("W26").values = [[51]];
return {before, after: readState()};`);
await capture('recalculate', `const before = readState();
const nativeReturn = wb.recalculate();
return {nativeReturn: nativeReturn === undefined ? "undefined" : nativeReturn, before, after: readState()};`);
await capture('formula-edit', `const before = readState();
wb.worksheets.getItem("Assumptions").getRange("W25").formulas = [[
  "=IF(CHOOSE($D$4,ISBLANK(W26),ISBLANK(W27),ISBLANK(W28),ISBLANK(W29)),NA(),CHOOSE($D$4,W26,W27,W28,W29))"
]];
return {before, after: readState()};`);
await capture('fill', `const sheet = wb.worksheets.getItem("Assumptions");
const range = sheet.getRange("R25:W25");
const before = {values: range.values, formulas: range.formulas};
range.fillRight();
return {sheet: "Assumptions", range: "R25:W25", before, after: {values: range.values, formulas: range.formulas}};`);
await capture('copy', `const sheet = wb.worksheets.getItem("Assumptions");
const before = readState();
sheet.getRange("W26").copyFrom(sheet.getRange("R26"), "formulas");
return {source: "Assumptions!R26", destination: "Assumptions!W26", before, after: readState()};`);
await capture('final-recalculation', `wb.recalculate();
return {restoredBase: readState(), summary: {
  range: "K13:M20", values: wb.worksheets.getItem("Summary").getRange("K13:M20").values,
  formulas: wb.worksheets.getItem("Summary").getRange("K13:M20").formulas
}};`);
await capture('error-scan', `return await wb.inspect({
  kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\\\?|#NUM!|#NULL!|#SPILL!|#CALC!", maxChars: 3000,
  options: {useRegex: true, maxResults: 30}
});`);
await capture('render-after', `const image = await wb.render({sheetName: "Summary", range: "C7:M28", scale: 1.5, format: "png"});
const bytes = new Uint8Array(await image.arrayBuffer());
await fs.writeFile(path.join(dir, "summary-after.png"), bytes);
return {saved: "summary-after.png", mimeType: "image/png", bytes: bytes.length};`);
await capture('export', `const output = await SpreadsheetFile.exportXlsx(wb);
const nativeReturn = await output.save(path.join(dir, "man-group-demo.xlsx"));
const stats = await fs.stat(path.join(dir, "man-group-demo.xlsx"));
return {saveReturn: nativeReturn === undefined ? "undefined" : nativeReturn,
  saved: "man-group-demo.xlsx", bytes: stats.size};`);
await fs.writeFile(path.join(dir,'records.json'), JSON.stringify(records,null,2));

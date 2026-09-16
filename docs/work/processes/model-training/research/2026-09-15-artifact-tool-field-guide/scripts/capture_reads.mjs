// Read-only comparisons. Run with the bundled Node runtime; no edits or recalculation.
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import {fileURLToPath} from 'node:url';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';

const home=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const source=path.join(home,'files/inspection/source.xlsx');
const hash=value=>crypto.createHash('sha256').update(value).digest('hex');
const sourceHash=hash(await fs.readFile(source));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const setup='const sheet = wb.worksheets.getItem("Summary");\n';
const scopes=[
 {id:'cell',label:'One cell',where:'Summary!K13',code:setup+'const range = sheet.getRange("K13");',note:'Read one named cell. The ordinary getter still returns one row containing one value.'},
 {id:'block',label:'A rectangle',where:'Summary!K13:M20',code:setup+'const range = sheet.getRange("K13:M20");',note:'Read eight rows and three columns. The blank separator at row 18 stays in the result.'},
 {id:'row',label:'Part of one row',where:'Summary!K13:M13',code:setup+'const range = sheet.getRange("K13:M20").getRow(0);',note:'Read the first row inside the chosen rectangle. Zero means the first row of that rectangle.'},
 {id:'column',label:'Part of one column',where:'Summary!K13:K20',code:setup+'const range = sheet.getRange("K13:M20").getColumn(0);',note:'Read the first column inside the chosen rectangle. This is a bounded eight-cell read.'},
 {id:'index',label:'A cell by row and column number',where:'Summary!K13',code:setup+'const range = sheet.getCell(12, 10);',note:'Read the same K13 cell using zero-based numbers: row 12 is Excel row 13, column 10 is K.'},
 {id:'separate',label:'Several separate areas',where:'Summary!K13:M14 + Fee engine!D10:E12',multi:true,code:'const areas = [\n  {sheet: "Summary", address: "K13:M14"},\n  {sheet: "Fee engine", address: "D10:E12"}\n];\nconst selections = areas.map(area => ({\n  sheet: area.sheet,\n  range: wb.worksheets.getItem(area.sheet).getRange(area.address)\n}));',note:'The script reads two chosen rectangles and collects both results. The cells between them are not requested.'},
 {id:'sheet',label:'A whole tab’s used area',where:'Fee engine',multi:true,code:'const sheet = wb.worksheets.getItem("Fee engine");\nconst selections = [{sheet: sheet.name, range: sheet.getUsedRange()}];',note:'Ask the library how far the tab is used, then read that rectangle. Used areas can include cells with formatting but no values.'},
 {id:'workbook',label:'Every tab’s used area',where:'All six tabs',multi:true,code:'const selections = wb.worksheets.items.map(sheet => ({\n  sheet: sheet.name, range: sheet.getUsedRange()\n}));',note:'The script goes through all six tabs and reads each used rectangle. This returns cell contents; charts, comments and formatting need separate reads.'}
];
const methods=[
 {id:'values',label:'Values only',api:'range.values',why:'See the numbers, text, dates and blanks currently held in the cells.',assembly:'The library returns the values in rows. No formula text is requested.'},
 {id:'formulas',label:'Formulas only',api:'range.formulas',why:'See the calculation written in each position. Cells without a formula return an empty string.',assembly:'The library returns the formulas in rows. It does not include their numerical results.'},
 {id:'both',label:'Values and formulas — two lists',api:'range.values + range.formulas',why:'Get both readings. The same position in each list refers to the same cell.',assembly:'Two library reads. The script puts them into one object with values and formulas fields.'},
 {id:'paired',label:'Value and formula together per cell',api:'range.values + range.formulas, joined by the script',why:'Read each address beside its current value and its formula, without matching up two separate lists.',assembly:'Two library reads. The script pairs matching positions and adds a cell address; this layout is made by the script.'},
 {id:'details',label:'Formula details (advanced)',api:'range.formulas + range.displayFormulas + range.formulaInfos',why:'Inspect the stored formula, the formula shown for the cell, and the library’s extra formula information. These getters are useful for spill or data-table cells.',assembly:'Three library reads, packaged by the script. This ordinary Summary example does not demonstrate a spill or a sensitivity table.'},
 {id:'inspect',label:'Values and formula records through inspect',api:'wb.inspect({kind: "table,formula", ...})',why:'Compare the other reading route: inspect returns a block of values and separate addressed formula records.',assembly:'The library creates these records. The capture explicitly saves its NDJSON and metadata; an inspection table here means a rectangle of cell values.'},
 {id:'search',label:'Matching cells with value and formula',api:'wb.findCells(...)',why:'Ask for cells matching some text. Each match already includes its address, value and formula.',assembly:'The library pairs value and formula for matching cells. This searches for selected cells; it does not read every cell in the area.'}
];
const helpers=`function columnName(index) {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + (n - 1) % 26) + name;
  }
  return name;
}
function cellsTogether(range) {
  const values = range.values;
  const formulas = range.formulas;
  return values.flatMap((row, r) => row.map((value, c) => ({
    address: columnName(range.columnIndex + c) + (range.rowIndex + r + 1),
    value,
    formula: formulas[r][c]
  })));
}`;
const bodies={values:'range.values',formulas:'range.formulas',both:'({values: range.values, formulas: range.formulas})',paired:'cellsTogether(range)',details:'({formulas: range.formulas, displayFormulas: range.displayFormulas, formulaInfos: range.formulaInfos})'};
const AsyncFunction=Object.getPrototypeOf(async function(){}).constructor;
const cases=[];
async function capture(def){
 const output=await new AsyncFunction('wb',def.code)(wb);
 if(output===undefined)throw new Error('Missing output: '+def.id);
 const text=JSON.stringify(output);
 cases.push({...def,output:JSON.parse(text),outputSha256:hash(text)});
 console.log(def.id+' · '+text.length+' characters');
}
for(const method of methods.slice(0,4))for(const scope of scopes){
 const expr=bodies[method.id];
 const code=(method.id==='paired'?helpers+'\n\n':'')+scope.code+'\n'+(scope.multi?
  'return selections.map(({sheet, range}) => ({\n  sheet, range: range.address, result: '+expr+'\n}));':
  'return '+expr+';');
 const selections=scope.multi?await new AsyncFunction('wb',scope.code+'\nreturn selections.map(({sheet,range}) => ({sheet,address:range.address,row:range.rowIndex,col:range.columnIndex,rows:range.rowCount,cols:range.columnCount}));')(wb):await new AsyncFunction('wb',scope.code+'\nreturn [{sheet:sheet.name,address:range.address,row:range.rowIndex,col:range.columnIndex,rows:range.rowCount,cols:range.columnCount}];')(wb);
 await capture({id:method.id+'-'+scope.id,method:method.id,scope:scope.id,where:scope.where,note:scope.note,multi:!!scope.multi,selections,code});
}
const block=scopes.find(s=>s.id==='block');
await capture({id:'details-block',method:'details',scope:'block',where:block.where,note:block.note,selections:cases.find(c=>c.id==='both-block').selections,code:block.code+'\nreturn '+bodies.details+';'});
for(const [scope,input] of [['block',{sheetId:'Summary',range:'K13:M20'}],['sheet',{sheetId:'Fee engine'}],['workbook',{}]]){
 const args={kind:'table,formula',...input,maxChars:1000000};
 await capture({id:'inspect-'+scope,method:'inspect',scope,where:scopes.find(s=>s.id===scope).where,input:args,note:'The size allowance is deliberately large for this comparison. Check the saved response for omissions and previews.',code:'const result = await wb.inspect('+JSON.stringify(args,null,2)+');\nreturn {\n  ndjson: result.ndjson, recordCount: result.recordCount,\n  truncated: result.truncated, metadata: result.metadata\n};'});
}
for(const [scope,input] of [['workbook',{searchTerm:'profit',options:{maxResults:100}}],['sheet',{sheetId:'Fee engine',searchTerm:'Assumptions!',options:{matchFormulas:true,maxResults:100}}]]){
 await capture({id:'search-'+scope,method:'search',scope,where:scope==='workbook'?'All six tabs · search “profit”':'Fee engine · search formulas for “Assumptions!”',input,note:'Only matching cells are returned. The response reports the total number of matches and whether the selected limit cut them off.',code:'return wb.findCells('+JSON.stringify(input,null,2)+');'});
}
if(hash(await fs.readFile(source))!==sourceHash)throw new Error('Source file changed');
const receipt={capturedAt:new Date().toISOString(),packageVersion:JSON.parse(await fs.readFile(path.join(home,'node_modules/@oai/artifact-tool/package.json'),'utf8')).version,source:'files/inspection/source.xlsx',sourceSha256:sourceHash,readOnly:true,recalculated:false,cases:cases.length};
await fs.mkdir(path.join(home,'files/reads'),{recursive:true});
await fs.writeFile(path.join(home,'files/reads/captures.json'),JSON.stringify({receipt,scopes:scopes.map(({code,...s})=>s),methods,cases},null,2)+'\n');
console.log(JSON.stringify(receipt));

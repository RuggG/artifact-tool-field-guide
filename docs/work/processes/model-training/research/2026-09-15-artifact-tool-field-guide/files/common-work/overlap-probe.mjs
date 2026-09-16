import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const file=new URL('operations-run/man-group-operations.xlsx',import.meta.url);
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(file.pathname));
const lab=wb.worksheets.getItem('Fee lab');
const evaluate=code=>eval('(async()=>{'+code+'})()');
const r={id:'overlap-table',operation:'const extra = lab.tables.add("D22:F25", true, "OverlapProbe");\nreturn extra;',checkInput:{before:'return lab.tables.items.map(t => t.name);',after:'return {tableNames:lab.tables.items.map(t=>t.name),existingRange:"B20:D24",requestedRange:"D22:F25",overlap:"D22:D24",saved:false};'}};
r.before=await evaluate(r.checkInput.before);
const t=performance.now();
try{await evaluate(r.operation);r.libraryReturn={type:'Excel Table object',representation:'Live object returned. The added range overlaps existing FeeScenarioInputs at B20:D24.'};}catch(e){r.operationError={name:e.name,message:e.message};r.libraryReturn={javascriptType:'undefined'};}
r.elapsedMs=Math.round(performance.now()-t);
r.after=await evaluate(r.checkInput.after);
await fs.writeFile(new URL('operations-run/overlap-table.json',import.meta.url),JSON.stringify(r,null,2));
console.log(JSON.stringify(r));

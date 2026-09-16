import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import assert from 'node:assert/strict';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const root=path.dirname(fileURLToPath(import.meta.url)),out=path.join(root,'files/runs');
const w=Workbook.create(),s=w.worksheets.add('Model');
s.getRange('A1:C3').values=[['Quantity','Price','Revenue'],[2,10,null],[3,10,null]];
s.getRange('C2').formulas=[['=A2*B2']];s.getRange('C2:C3').fillDown();
const before=structuredClone(w.toProto()),input=s.getRange('B2:B3').values,captures=[];
try{for(const price of [8,10,12]){s.getRange('B2:B3').values=[[price]];w.recalculate();const values=s.getRange('C2:C3').values;assert.deepEqual(values,[[2*price],[3*price]]);captures.push({price,values});}}
finally{s.getRange('B2:B3').values=input;w.recalculate();}
assert.deepEqual(s.getRange('C2:C3').values,[[20],[30]]);
const clone=Workbook.load(before);clone.worksheets.getItem('Model').getRange('B2').values=[[99]];
assert.equal(s.getRange('B2').values[0][0],10);
const findings=[];for(const row of s.getUsedRange().values)for(const v of row)if(typeof v==='string'&&(/^#(?:REF!|DIV\/0!|VALUE!|NAME\?|NUM!|N\/A|SPILL!|CALC!)/.test(v)||/not implemented|Cannot read|is not a function|Invalid array length/i.test(v)))findings.push(v);
assert.deepEqual(findings,[]);
const custom=s.shapes.add({geometry:'custom',customPaths:[{width:120,height:90,commands:[{moveTo:{x:0,y:0}},{lineTo:{x:120,y:0}},{quadBezTo:{x1:100,y1:90,x:60,y:90}},{cubicBezTo:{x1:30,y1:90,x2:0,y2:60,x:0,y:0}},{close:{}}]}],anchor:{from:{row:5,col:0},extent:{widthPx:120,heightPx:90}},fill:'#bfdbfe',line:{fill:'#1e40af',width:2}});
await(await SpreadsheetFile.exportXlsx(w)).save(path.join(out,'workflow-example.xlsx'));
const preview=await w.render({sheetName:'Model',range:'A1:F12',scale:2});await fs.writeFile(path.join(out,'workflow-example.png'),new Uint8Array(await preview.arrayBuffer()));
const imported=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,'workflow-example.xlsx')));imported.recalculate();assert.deepEqual(imported.worksheets.getItem('Model').getRange('C2:C3').values,[[20],[30]]);
await fs.writeFile(path.join(out,'workflow-example.json'),JSON.stringify({captures,restored:s.getRange('B2:B3').values,independentClone:true,errorScan:findings,customShapeId:custom.id,reimportedShapes:imported.worksheets.getItem('Model').shapes.items.length,reimportedValues:imported.worksheets.getItem('Model').getRange('C2:C3').values},null,2));
console.log('Scenario capture/restore, independent clone, error scan, custom path, render/export/reimport checks completed.');

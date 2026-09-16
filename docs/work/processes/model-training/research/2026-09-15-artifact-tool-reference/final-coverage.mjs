import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {spawn} from 'node:child_process';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const root=path.dirname(fileURLToPath(import.meta.url)),out=path.join(root,'files/runs');
const save=async(n,x)=>fs.writeFile(path.join(out,n),JSON.stringify(x,null,2));
if(!process.argv[2]){
 const results=[];
 for(const mode of ['geometry','roundtrip']){
  results.push(await new Promise(resolve=>{
   const child=spawn(process.execPath,[fileURLToPath(import.meta.url),mode],{stdio:['ignore','pipe','pipe']});let stdout='',stderr='',timedOut=false;
   child.stdout.on('data',x=>stdout+=x);child.stderr.on('data',x=>stderr+=x);
   const timer=setTimeout(()=>{timedOut=true;child.kill('SIGKILL');},45000);
   child.on('close',(code,signal)=>{clearTimeout(timer);resolve({mode,code,signal,timedOut,stdout,stderr});});
  }));
 }
 await save('final-coverage.json',results);console.log(JSON.stringify(results));
}else if(process.argv[2]==='geometry'){
 const enums=(await fs.readFile(path.join(out,'enums-help.ndjson'),'utf8')).trim().split('\n').map(JSON.parse);
 const shapes=enums.find(x=>x.name==='enum.ShapeGeometry').values,lines=enums.find(x=>x.name==='enum.LineStyle').values;
 const w=Workbook.create(),results=[];
 for(let start=0;start<shapes.length;start+=40){
  const s=w.worksheets.add('Shapes '+(start/40+1));s.showGridLines=false;s.getRange('A1:P82').format.columnWidthPx=38;s.getRange('A1:P82').format.rowHeightPx=17;
  for(let i=start;i<Math.min(start+40,shapes.length);i++){
   const row=Math.floor((i-start)/4)*8,col=((i-start)%4)*4;
   try{s.getCell(row,col).values=[[shapes[i]]];s.getCell(row,col).format.font.size=8;
    const shape=s.shapes.add({geometry:shapes[i],anchor:{from:{row:row+1,col},extent:{widthPx:130,heightPx:95}},fill:'#bfdbfe',line:{fill:'#1e40af',width:1,style:lines[i%lines.length]}});
    results.push({geometry:shapes[i],lineStyle:lines[i%lines.length],created:true,id:shape.id,sheet:s.name});
   }catch(e){results.push({geometry:shapes[i],error:String(e)});}
  }
 }
 await save('geometry.json',results);
 const file=await SpreadsheetFile.exportXlsx(w);await file.save(path.join(out,'geometry.xlsx'));
 for(const s of w.worksheets.items){try{const b=await w.render({sheetName:s.name,range:'A1:P80',scale:1});await fs.writeFile(path.join(out,s.name.replace(' ','-')+'.png'),new Uint8Array(await b.arrayBuffer()));}catch(e){results.push({sheet:s.name,renderError:String(e)});}await save('geometry.json',results);}
 console.log(JSON.stringify({created:results.filter(x=>x.created).length,errors:results.filter(x=>x.error||x.renderError)}));
}else{
 const results=[];
 for(const name of ['comments','native-note','sparklines','validations','conditional-formats','images','shapes','chart-settings','names','named-function-variants','data-table','one-variable-data-tables','pivot-slicer','slicer-filter','table-total-fresh-formula']){
  try{const w=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,name+'.xlsx')));w.recalculate();
   const snapshot={sheets:w.worksheets.items.map(s=>({name:s.name,charts:s.charts.items.length,images:s.images.items.length,shapes:s.shapes.items.length,tables:s.tables.items.length,sparklines:s.sparklineGroups.getAll().length})),inspection:await w.inspect({kind:'definedName,thread',maxChars:20000})};
   await(await SpreadsheetFile.exportXlsx(w)).save(path.join(out,'reimport-'+name+'.xlsx'));results.push({name,completed:true,...snapshot});
  }catch(e){results.push({name,error:String(e)});}await save('feature-roundtrip.json',results);
 }
 console.log(JSON.stringify(results.map(({name,completed,error})=>({name,completed,error}))));
}

import fs from 'node:fs/promises';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {Workbook} from '@oai/artifact-tool';
import {artifactSession} from '@oai/artifact-tool/artifact-session';
const root=path.dirname(fileURLToPath(import.meta.url)),out=path.join(root,'files/runs'),sessionRoot=path.join(root,'files/local-sessions');
await fs.mkdir(sessionRoot,{recursive:true});
const pkg=path.dirname(fileURLToPath(import.meta.resolve('@oai/artifact-tool')));
const child=spawn(process.execPath,[path.join(pkg,'artifact-session-mcp/server.mjs')],{stdio:['pipe','pipe','pipe']});
const pending=new Map;let id=0,buffer='',stderr='';
child.stderr.on('data',d=>{stderr+=d;});
child.stdout.on('data',d=>{buffer+=d;let n;while((n=buffer.indexOf('\n'))>=0){const line=buffer.slice(0,n);buffer=buffer.slice(n+1);try{const r=JSON.parse(line),p=pending.get(r.id);if(p){clearTimeout(p.timer);pending.delete(r.id);p.resolve(r);}}catch{}}});
function rpc(method,params={}){return new Promise((resolve,reject)=>{const current=++id;const timer=setTimeout(()=>{pending.delete(current);reject(new Error('MCP response timed out'));},30000);pending.set(current,{resolve,reject,timer});child.stdin.write(JSON.stringify({jsonrpc:'2.0',id:current,method,params})+'\n');});}
const results=[];
async function call(label,args,meta){const response=await rpc('tools/call',{name:'artifact_session_run',arguments:args,...meta?{_meta:meta}:{}});results.push({label,input:args,response});await fs.writeFile(path.join(out,'sessions.json'),JSON.stringify(results,null,2));console.log(label,JSON.stringify(response.result?.structuredContent??response.error));return response;}
try{
 results.push({label:'initialize',response:await rpc('initialize',{protocolVersion:'2025-11-25',capabilities:{},clientInfo:{name:'artifact-reference-probe',version:'1'}})});
 const schema=await rpc('tools/list');await fs.writeFile(path.join(out,'session-mcp-schema.json'),JSON.stringify(schema,null,2));
 const base={target:'artifact-reference-probe',sessionRoot,outputDirectory:path.join(sessionRoot,'outputs'),artifactType:'spreadsheet',timeoutMs:10000};
 await call('create-and-commit',{...base,create:true,title:'Synthetic session probe',code:"const s=workbook.worksheets.add('Session'); s.getRange('A1:C2').values=[['Qty','Price','Total'],[3,4,'=A2*B2']]; return {total:s.getRange('C2').values,session:session.info};"});
 await call('resume-and-edit',{...base,input:{qty:5},summary:'Change synthetic quantity',code:"const s=workbook.worksheets.getItem('Session'); const before=s.getRange('C2').values; s.getRange('A2').values=[[input.qty]]; workbook.recalculate(); return {before,after:s.getRange('C2').values};"});
 await call('managed-outputs',{...base,code:"const inspection=await workbook.inspect({kind:'table',range:'Session!A1:C2',fileName:'session-inspect.ndjson'}); const preview=await workbook.render({sheetName:'Session',range:'A1:C2',fileName:'session-preview.png',scale:2}); const exported=await workbook.export({format:'xlsx',fileName:'session-probe.xlsx'}); return {inspection,preview,exported};"});
 await call('throw-after-mutation',{...base,code:"workbook.worksheets.getItem('Session').getRange('A2').values=[[999]]; throw new Error('Intentional rollback probe');"});
 await call('read-after-failed-run',{...base,code:"return {value:workbook.worksheets.getItem('Session').getRange('A2').values};"});
 await call('import-existing-xlsx',{...base,target:path.join(out,'roundtrip.xlsx'),code:"return {sheets:workbook.worksheets.items.map(s=>s.name),values:workbook.worksheets.getItemAt(0).getRange('A1:C3').values};"});
 await call('invalid-parameter',{...base,code:'return 1;',invented:true});
 await call('missing-target-without-task-meta',{code:'return 1;',sessionRoot});
 await call('thread-scoped-default',{sessionRoot,code:"const s=workbook.worksheets.add('Default');s.getRange('A1').values=[[42]];return {value:s.getRange('A1').values};"},{'openai/threadId':'artifact-reference-synthetic-task'});
 await call('execution-timeout',{...base,timeoutMs:500,code:'while(true){}'});
 await call('read-after-timeout',{...base,code:"return {value:workbook.worksheets.getItem('Session').getRange('A2').values};"});
 let directError;try{await artifactSession.run(async()=>({ok:true}));}catch(e){directError=String(e);}
 const staticSurface=Object.getOwnPropertyNames(Workbook).filter(x=>!['length','name','prototype'].includes(x));
 results.push({label:'first-party-wrapper-without-trusted-repl',error:directError,staticSurface,runSignature:Workbook.run?.toString()});
 await fs.writeFile(path.join(out,'sessions.json'),JSON.stringify(results,null,2));
}finally{child.stdin.end();await fs.writeFile(path.join(out,'sessions.stderr.log'),stderr);}

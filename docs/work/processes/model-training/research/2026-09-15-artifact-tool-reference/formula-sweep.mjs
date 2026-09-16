import fs from 'node:fs/promises';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
const root=path.dirname(fileURLToPath(import.meta.url)),out=path.join(root,'files/runs');
const entries=(await fs.readFile(path.join(out,'formula-help.ndjson'),'utf8')).trim().split('\n').map(JSON.parse);
const jobs=entries.map((entry,index)=>({entry,index})),results=[];
async function worker(){for(;;){const job=jobs.shift();if(!job)return;const start=Date.now();
const result=await new Promise(resolve=>{
const child=spawn(process.execPath,[path.join(root,'explore.mjs'),'formula-worker',String(job.index)],{stdio:['ignore','ignore','pipe']});
let stderr='',timedOut=false;child.stderr.on('data',chunk=>{if(stderr.length<2000)stderr+=chunk.toString().slice(0,2000-stderr.length);});
const timer=setTimeout(()=>{timedOut=true;child.kill('SIGKILL');},15000);
child.on('error',error=>{clearTimeout(timer);resolve({error:String(error)});});
child.on('close',(code,signal)=>{clearTimeout(timer);resolve({code,signal,timedOut,stderr});});
});
results.push({name:job.entry.name,index:job.index,ms:Date.now()-start,...result});
await fs.writeFile(path.join(out,'formula-sweep.json'),JSON.stringify(results,null,2));
if(results.length%25===0)console.log('Formula entries completed:',results.length,'/',entries.length);
}}
await Promise.all(Array.from({length:4},()=>worker()));
const all=[];for(let i=0;i<entries.length;i++){try{all.push(JSON.parse(await fs.readFile(path.join(out,'formula-'+String(i).padStart(3,'0')+'.json'),'utf8')));}catch{all.push({name:entries[i].name,category:entries[i].category,worker:results.find(x=>x.index===i)});}}
await fs.writeFile(path.join(out,'formula-results.json'),JSON.stringify(all,null,2));
console.log('Completed',all.length,'formula entries.');

import fs from 'node:fs/promises';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
const root=path.dirname(fileURLToPath(import.meta.url)),out=path.join(root,'files/runs');
const entries=(await fs.readFile(path.join(out,'api-help.ndjson'),'utf8')).trim().split('\n').map(JSON.parse).filter(x=>x.kind==='api').flatMap(x=>(x.examples||[]).map(example=>({name:x.name,...example})));
const selected=process.env.API_INDICES?new Set(process.env.API_INDICES.split(',').map(Number)):null;
const jobs=entries.map((entry,index)=>({entry,index})).filter(x=>!selected||selected.has(x.index)),results=[];
async function worker(){for(;;){const job=jobs.shift();if(!job)return;const start=Date.now();const result=await new Promise(resolve=>{const c=spawn(process.execPath,[path.join(root,'explore.mjs'),'api-worker',String(job.index)],{stdio:['ignore','ignore','pipe']});let stderr='',timedOut=false;c.stderr.on('data',chunk=>{if(stderr.length<1500)stderr+=chunk.toString().slice(0,1500-stderr.length);});const t=setTimeout(()=>{timedOut=true;c.kill('SIGKILL');},15000);c.on('close',(code,signal)=>{clearTimeout(t);resolve({code,signal,timedOut,stderr});});});results.push({name:job.entry.name,index:job.index,ms:Date.now()-start,...result});await fs.writeFile(path.join(out,'api-sweep.json'),JSON.stringify(results,null,2));if(results.length%25===0)console.log(results.length,'/',entries.length);}}
await Promise.all(Array.from({length:4},()=>worker()));const all=[];for(let i=0;i<entries.length;i++){try{all.push(JSON.parse(await fs.readFile(path.join(out,'api-example-'+String(i).padStart(3,'0')+'.json'),'utf8')));}catch{all.push({index:i,...entries[i],worker:results.find(x=>x.index===i)});}}
await fs.writeFile(path.join(out,'api-example-results.json'),JSON.stringify(all,null,2));console.log('Completed',all.length,'API examples.');

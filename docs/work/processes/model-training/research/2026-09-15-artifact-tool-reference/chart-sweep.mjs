import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {spawn} from 'node:child_process';
const root=path.dirname(fileURLToPath(import.meta.url)),out=path.join(root,'files/runs');
const types=(await fs.readFile(path.join(out,'enums-help.ndjson'),'utf8')).trim().split('\n').map(JSON.parse).find(x=>x.name==='enum.ChartType').values;
const runs=[];
for(const id of ['F01','F02','F03','F04','F05','F06','F07',...types.map(x=>'CH-'+x),'F08','F09']){
 const result=await new Promise(resolve=>{
  const child=spawn(process.execPath,[path.join(root,'explore.mjs'),'features'],{cwd:root,env:{...process.env,PROBE_FILTER:'^'+id+'$',RUN_LABEL:'bounded-'+id},stdio:['ignore','pipe','pipe']});
  let stdout='',stderr='',timedOut=false;
  child.stdout.on('data',x=>stdout+=x);child.stderr.on('data',x=>stderr+=x);
  const timer=setTimeout(()=>{timedOut=true;child.kill('SIGKILL');},10000);
  child.on('close',(code,signal)=>{clearTimeout(timer);resolve({id,code,signal,timedOut,stdout,stderr});});
 });runs.push(result);await fs.writeFile(path.join(out,'chart-sweep.json'),JSON.stringify(runs,null,2));
 console.log(JSON.stringify({id,timedOut:result.timedOut,code:result.code}));
}

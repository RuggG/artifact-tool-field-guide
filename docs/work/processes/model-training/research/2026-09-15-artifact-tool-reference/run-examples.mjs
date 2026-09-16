import fs from 'node:fs/promises';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
const root=path.dirname(fileURLToPath(import.meta.url)),out=path.join(root,'files/runs'),dir=path.join(out,'examples');
await fs.mkdir(dir,{recursive:true});
const results=[];
for(const name of ['quick_start_example','formula_trace_and_help','chart_suggestions','inspect_existing_workbooks']){
 const src=path.join(root,'files/sources/package/examples',name+'.ts');let code=await fs.readFile(src,'utf8');const adaptations=[];
 if(code.includes('../../../../oai_js_artifact_tool')){code=code.replace('../../../../oai_js_artifact_tool','@oai/artifact-tool');adaptations.push('Replaced source-repository-only import with installed package import.');}
 const file=path.join(dir,name+'.ts');await fs.writeFile(file,code);
 const args=name==='inspect_existing_workbooks'?[path.join(out,'roundtrip.xlsx')]:name==='chart_suggestions'?[path.join(dir,'charts')]:[];
 const result=await new Promise(resolve=>{let stdout='',stderr='',timeout=false;const c=spawn(process.execPath,['--experimental-strip-types',file,...args],{cwd:dir,stdio:['ignore','pipe','pipe']});c.stdout.on('data',d=>stdout+=d);c.stderr.on('data',d=>stderr+=d);const t=setTimeout(()=>{timeout=true;c.kill('SIGKILL');},20000);c.on('close',(exitCode,signal)=>{clearTimeout(t);resolve({exitCode,signal,timeout,stdout,stderr});});});
 results.push({name,adaptations,...result});await fs.writeFile(path.join(out,'supplied-examples.json'),JSON.stringify(results,null,2));console.log(name,result.exitCode,result.timeout);
}

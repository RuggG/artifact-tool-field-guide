import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
const root = path.dirname(fileURLToPath(import.meta.url));
const deps = process.env.ARTIFACT_DEPENDENCIES || '/Users/ruggerogargiulo/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const pkg = path.join(deps, '@oai/artifact-tool');
const skill = process.env.SPREADSHEET_SKILL || '/Users/ruggerogargiulo/.codex/plugins/cache/openai-primary-runtime/spreadsheets/26.909.12148/skills/spreadsheets';
await fs.mkdir(path.join(root,'files'),{recursive:true});
const entries=[];
async function walk(dir,sourceKind,base=dir) {
  for (const e of await fs.readdir(dir,{withFileTypes:true})) {
    const src=path.join(dir,e.name);
    if(e.isDirectory()) { await walk(src,sourceKind,base); continue; }
    if(!/\.(md|ts|mjs|yaml)$/.test(e.name)) continue;
    const bytes=await fs.readFile(src), rel=path.join('files','sources',sourceKind,path.relative(base,src));
    await fs.mkdir(path.dirname(path.join(root,rel)),{recursive:true});
    await fs.writeFile(path.join(root,rel),bytes);
    entries.push({sourceKind,sourcePath:src,savedPath:rel,bytes:bytes.length,sha256:crypto.createHash('sha256').update(bytes).digest('hex'),headings:bytes.toString().split('\n').filter(x=>/^#{1,6} /.test(x))});
  }
}
await walk(path.join(pkg,'docs/spreadsheets'),'package');
await walk(skill,'skill');
// Preserve the small session adapters; the bundled multi-megabyte clients are
// fingerprinted through a targeted excerpt rather than duplicated wholesale.
for (const rel of ['artifact-session/service.mjs','artifact-session/host-client.mjs','artifact-session-mcp/server.mjs','artifact-session-mcp/worker.mjs']) {
  const src=path.join(pkg,'dist',rel),bytes=await fs.readFile(src),savedPath=path.join('files','sources','runtime',rel);
  await fs.mkdir(path.dirname(path.join(root,savedPath)),{recursive:true});
  await fs.writeFile(path.join(root,savedPath),bytes);
  entries.push({sourceKind:'runtime',sourcePath:src,savedPath,bytes:bytes.length,sha256:crypto.createHash('sha256').update(bytes).digest('hex'),headings:[]});
}
const bundlePath=path.join(pkg,'dist/artifact_tool.mjs'),bundle=await fs.readFile(bundlePath),bundleText=bundle.toString(),customStart=bundleText.indexOf('function vQr('),customEnd=bundleText.indexOf('function TQr(',customStart);
await fs.writeFile(path.join(root,'files','custom-path-source.json'),JSON.stringify({sourcePath:bundlePath,sha256:crypto.createHash('sha256').update(bundle).digest('hex'),purpose:'Targeted explanation of the custom-geometry prerequisite and accepted path commands in this installed version.',excerpt:bundleText.slice(customStart,customEnd)},null,2));
const packageJson=JSON.parse(await fs.readFile(path.join(pkg,'package.json'),'utf8'));
await fs.writeFile(path.join(root,'files','source-manifest.json'),JSON.stringify({capturedAt:new Date().toISOString(),package:packageJson,node:process.version,platform:process.platform,entries},null,2));
await fs.writeFile(path.join(root,'files','source-index.md'),'# Source inventory\n\n'+entries.map(x=>'## '+x.savedPath+'\n\n'+x.headings.join('\n')+'\n').join('\n'));
try {await fs.symlink(deps,path.join(root,'node_modules'),'dir');} catch(e) {if(e.code!=='EEXIST') throw e;}
console.log(JSON.stringify({version:packageJson.version,documents:entries.length,totalBytes:entries.reduce((a,x)=>a+x.bytes,0)}));

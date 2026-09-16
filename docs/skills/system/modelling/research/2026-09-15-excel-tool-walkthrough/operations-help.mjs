import fs from 'node:fs/promises';
import {Workbook} from '@oai/artifact-tool';
const wb=Workbook.create();
const result=wb.help('*',{search:'worksheet\\.(delete|name)|range\\.(insert|delete)|tables\\.add|table\\.(resize|delete)',include:'index,notes,examples',maxChars:8500});
await fs.writeFile(new URL('operations-help.json',import.meta.url),JSON.stringify(result,null,2));
console.log(result.ndjson);

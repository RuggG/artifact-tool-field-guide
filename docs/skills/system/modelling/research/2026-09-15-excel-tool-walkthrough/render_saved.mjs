import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const path=new URL('operations-run/new-fee-workbook.xlsx',import.meta.url);
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path.pathname));
const image=await wb.render({sheetName:'Fee calculation',range:'B2:C5',scale:2,format:'png'});
await fs.writeFile(new URL('operations-run/new-fee-workbook.png',import.meta.url),new Uint8Array(await image.arrayBuffer()));
console.log(JSON.stringify({input:'new-fee-workbook.xlsx',sheet:'Fee calculation',range:'B2:C5',values:wb.worksheets.getItem('Fee calculation').getRange('B2:C5').values}));

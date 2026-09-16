(() => {
'use strict';
const D=window.ArtifactReadsData;if(!D)return;
const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const json=x=>JSON.stringify(x,null,2);
const pre=x=>'<pre class="af-code" tabindex="0"><code>'+esc(typeof x==='string'?x:json(x))+'</code></pre>';
const col=i=>{let s='';for(let n=i+1;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const origin=address=>{const m=address.replace(/\$/g,'').match(/([A-Z]+)(\d+)/);return {row:Number(m[2])-1,col:[...m[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1};};
const value=x=>x===null?'<span class="rx-empty" title="Returned null">null</span>':x===''?'<span class="rx-empty" title="Returned empty string">""</span>':'<span class="'+(typeof x==='number'||(typeof x==='string'&&/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(x))?'rx-number':'rx-text')+'">'+esc(typeof x==='object'?json(x):x)+'</span>';
const scroll=(html,label)=>'<div class="rx-scroll" tabindex="0" role="region" aria-label="'+esc(label)+'">'+html+'</div><p class="rx-scroll-hint">Long results scroll within this box; all returned rows and columns remain available.</p>';
function grid(rows,s,label){return '<h5>'+esc(label)+'</h5>'+scroll('<table class="rx-grid"><thead><tr><th scope="col">Row</th>'+Array.from({length:Math.max(0,...rows.map(r=>r.length))},(_,c)=>'<th scope="col">'+col(s.col+c)+'</th>').join('')+'</tr></thead><tbody>'+rows.map((r,i)=>'<tr><th scope="row">'+(s.row+i+1)+'</th>'+r.map((v,c)=>'<td data-rx-cell="'+col(s.col+c)+(s.row+i+1)+'">'+value(v)+'</td>').join('')+'</tr>').join('')+'</tbody></table>',label);}
function records(rows,keys,label){return scroll('<table class="rx-records"><thead><tr>'+keys.map(([k,l])=>'<th scope="col">'+esc(l)+'</th>').join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr data-rx-record>'+keys.map(([k])=>'<td class="'+(['formula','display','formulaInfo'].includes(k)?'rx-formula':'')+'">'+value(r[k]??null)+'</td>').join('')+'</tr>').join('')+'</tbody></table>',label);}
const cellKeys=[['address','Cell'],['value','Value'],['formula','Formula']];
function coverage(c){
 if(c.selections){const n=c.selections.reduce((n,s)=>n+s.rows*s.cols,0);return '<strong>'+n.toLocaleString('en-US')+' cell positions returned'+(c.selections.length>1?' across '+c.selections.length+' areas':'')+'.</strong> Every position in the chosen rectangle'+(c.selections.length>1?'s':'')+' is present, including blanks. '+(c.scope==='workbook'?'The used areas include formatted padding; the inspection route below returns smaller content rectangles. ':'')+'The values are from this import; this read did not recalculate the workbook.';}
 if(c.method==='search'){const o=c.output;return '<strong>'+o.matches.length+' of '+o.total+' matches returned.</strong> '+(o.truncated?'The request stops at '+o.limit+' matches. The remaining '+(o.total-o.matches.length)+' are absent from this response.':'All matches reported by this search are present.')+' This is a search result, not the whole content of the tab. Values in these match records are text; Raw preserves that difference.';}
 const rr=c.output.ndjson.trim().split('\n').filter(Boolean).map(JSON.parse),g=rr.filter(r=>r.kind==='table'),f=rr.filter(r=>r.kind==='formula');const n=g.reduce((n,r)=>n+r.values.reduce((n,row)=>n+row.length,0),0);
 return '<strong>'+n.toLocaleString('en-US')+' value positions and '+f.length.toLocaleString('en-US')+' formula records returned.</strong> '+(g.some(r=>r.valuesTruncated)||c.output.truncated?'Check the limits in the saved response.':'These recorded grids are complete for their stated rectangles; no preview or omitted-record flag is set.')+' “Table” here means rows and columns of cell values, not a named Excel Table.';
}
function readable(c,area){
 if(c.method==='search')return records(c.output.matches,[['sheet','Tab'],...cellKeys,['match','Matched in']],'Matching cells')+'<details class="rx-door"><summary>Search counts and notices</summary>'+pre(Object.fromEntries(Object.entries(c.output).filter(([k])=>k!=='matches')))+'</details>';
 if(c.method==='inspect'){
  const rows=c.output.ndjson.trim().split('\n').filter(Boolean).map(JSON.parse),grids=rows.filter(r=>r.kind==='table'),formulas=rows.filter(r=>r.kind==='formula');
  return grids.map(r=>'<section class="rx-area"><h4>'+esc(r.sheet+'!'+r.address)+'</h4>'+grid(r.values,origin(r.address),'Cell values · '+r.rows+' rows × '+r.cols+' columns')+'</section>').join('')+'<h5>'+formulas.length+' separate formula records</h5>'+records(formulas,[['sheet','Tab'],['address','Cell'],['formula','Formula']],'Formula records')+'<details class="rx-door"><summary>Response counts and metadata</summary>'+pre(Object.fromEntries(Object.entries(c.output).filter(([k])=>k!=='ndjson')))+'</details>';
 }
 const s=c.selections[area],out=c.multi?c.output[area].result:c.output;
 let h='<h4>'+esc(s.sheet+'!'+s.address)+' <small>'+s.rows+' rows × '+s.cols+' columns</small></h4>';
 if(c.method==='paired')h+=records(out,cellKeys,'Values and formulas by cell');
 else if(['values','formulas'].includes(c.method))h+=grid(out,s,c.method==='values'?'Values':'Formulas');
 else if(c.method==='both')h+=grid(out.values,s,'1. Values')+grid(out.formulas,s,'2. Formulas · same cell positions');
 else if(c.method==='details'){
  const rows=out.formulaInfos.flatMap((r,i)=>r.map((v,j)=>({address:col(s.col+j)+(s.row+i+1),formula:out.formulas[i][j],display:out.displayFormulas[i][j],formulaInfo:v})));
  h+='<p>For these ordinary formulas, the stored and displayed formulas match. Blank cells have no formula information. This capture does not contain a spilled array or sensitivity table.</p>'+records(rows,[['address','Cell'],['formula','Stored formula'],['display','Displayed formula'],['formulaInfo','Formula information']],'Formula details');
 }
 return h+'<p class="rx-caption">'+(c.method==='paired'?'The script adds the cell addresses and pairs values with formulas. ':'Row and column headings are reading aids added by this page. ')+'<code>null</code> means no value; <code>""</code> means an empty string (usually no formula). Values keep the captured precision and are not formatted as they appear in Excel.</p>';
}
function mount(root,{tree}){
 if(!root||root.dataset.rxMounted)return;root.dataset.rxMounted='true';
 const p=new URL(location.href).searchParams;let id=p.get('read')||'paired-block',view=p.get('rview')||'readable',area=0;
 if(!D.cases.some(c=>c.id===id))id='paired-block';if(!['readable','tree','raw'].includes(view))view='readable';
 const current=()=>D.cases.find(c=>c.id===id);
 function remember(){const u=new URL(location.href);u.searchParams.set('read',id);u.searchParams.set('rview',view);u.hash='common-work';history.replaceState(null,'',u);}
 function draw(){
  const c=current(),m=D.methods.find(m=>m.id===c.method),active=document.activeElement,restore=active&&root.contains(active)?active.id||active.dataset.rxView:null;
  root.dataset.rxCase=id;
  root.innerHTML='<div class="rx-heading"><div class="af-kicker">38 executed reading comparisons</div><h4>Choose the cells. Choose what comes back.</h4><p>A value is what a cell contains now. A formula is the rule that calculates it. You can read either one, or have your code put them together.</p></div>'+
   '<div class="rx-controls"><label>What to get back<select id="rx-method">'+D.methods.map(x=>'<option value="'+x.id+'" '+(x.id===c.method?'selected':'')+'>'+esc(x.label)+'</option>').join('')+'</select></label><label>Which cells to read<select id="rx-scope">'+D.scopes.filter(s=>D.cases.some(x=>x.method===c.method&&x.scope===s.id)).map(s=>'<option value="'+s.id+'" '+(s.id===c.scope?'selected':'')+'>'+esc(c.method==='search'?(s.id==='sheet'?'One tab · Fee engine':'All tabs'):c.method==='inspect'?(s.id==='sheet'?'A whole tab’s contents':s.id==='workbook'?'Every tab’s contents':s.label):s.label)+'</option>').join('')+'</select></label></div>'+
   '<section class="rx-request" aria-label="Selected reading request"><h5>The request</h5><p><strong>'+esc(c.where)+'</strong> · '+esc(m.why)+'</p><p>'+esc(c.note)+'</p><code class="cw-api">'+esc(m.api)+'</code><p class="rx-assembly">'+esc(m.assembly)+'</p><details class="rx-door"><summary>Exact executed code</summary><p>The file has already been imported as <code>wb</code>. The capture script runs this function body and saves its complete return.</p>'+pre(c.code)+'</details></section>'+
   '<section class="rx-response" aria-label="Saved reading response"><div class="rx-response-head"><h4>What came back</h4><div class="af-segments" aria-label="Reading response view">'+['readable','tree','raw'].map(v=>'<button type="button" class="af-button'+(v===view?' is-active':'')+'" data-rx-view="'+v+'" aria-pressed="'+(v===view)+'">'+v[0].toUpperCase()+v.slice(1)+'</button>').join('')+'</div></div><p class="rx-coverage'+(c.output.truncated?' rx-limited':'')+'">'+coverage(c)+'</p>'+
   (view==='readable'&&c.multi?'<label class="rx-area-control">View returned area <select id="rx-area">'+c.selections.map((s,i)=>'<option value="'+i+'" '+(area===i?'selected':'')+'>'+esc(s.sheet+'!'+s.address)+'</option>').join('')+'</select></label><p class="rx-caption">All '+c.selections.length+' areas were returned. This display choice shows one at a time; Tree, Raw and the download contain them all.</p>':'')+
   '<div class="rx-body"></div></section>'+
   '<p class="rx-delivery"><strong>What the agent sees:</strong> the library returns this data to the code. It reaches the conversation only if the code prints or returns it, and a large tool message can still be cut short. Saving a complete response to a file does not put the whole file in the conversation.</p>'+
   '<div class="af-record-actions"><button type="button" class="af-button" data-rx-copy="code">Copy executed code</button><button type="button" class="af-button" data-rx-download>Download complete example</button><button type="button" class="af-button" data-rx-copy="link">Copy link</button></div><p class="cw-sources"><a href="scripts/capture_reads.mjs">Capture script</a> · <a href="files/reads/captures.json">All 38 complete captures</a> · <a href="data/reads-checks.json">Data checks</a> · Artifact Tool '+esc(D.receipt.packageVersion)+' · unchanged source workbook</p>';
  const body=root.querySelector('.rx-body');
  if(view==='readable')body.innerHTML=readable(c,area);
  else if(view==='tree'){body.className+=' af-tree';body.append(tree(c.method==='inspect'?c.output.ndjson.trim().split('\n').map(JSON.parse):c.output,'Returned data'));if(c.method==='inspect')body.insertAdjacentHTML('beforeend','<details class="rx-door"><summary>Original response with metadata</summary>'+pre(c.output)+'</details>');}
  else body.innerHTML='<p class="rx-caption">'+(c.method==='inspect'?'Exact NDJSON record text. Each line is one returned record.':'Exact saved return, pretty-printed as JSON. The page has not added addresses or labels to this view.')+'</p>'+pre(c.method==='inspect'?c.output.ndjson:c.output)+(c.method==='inspect'?'<details class="rx-door"><summary>Original response with metadata</summary>'+pre(c.output)+'</details>':'');
  if(restore)(root.querySelector('#'+CSS.escape(restore))||root.querySelector('[data-rx-view="'+CSS.escape(restore)+'"]'))?.focus({preventScroll:true});
 }
 root.addEventListener('change',e=>{const c=current();if(e.target.id==='rx-method'){const list=D.cases.filter(x=>x.method===e.target.value);id=(list.find(x=>x.scope===c.scope)||list.find(x=>x.scope==='block')||list[0]).id;area=0;}else if(e.target.id==='rx-scope'){id=c.method+'-'+e.target.value;area=0;}else if(e.target.id==='rx-area')area=Number(e.target.value);else return;draw();remember();});
 root.addEventListener('click',async e=>{const b=e.target.closest('button');if(!b)return;
  if(b.dataset.rxView){view=b.dataset.rxView;draw();remember();}
  if(b.hasAttribute('data-rx-download')){const url=URL.createObjectURL(new Blob([json({receipt:D.receipt,...current()})],{type:'application/json'})),a=document.createElement('a');a.href=url;a.download='read-'+id+'.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
  if(b.dataset.rxCopy){remember();const text=b.dataset.rxCopy==='code'?current().code:location.href;try{await navigator.clipboard.writeText(text);b.textContent='Copied';}catch{root.querySelector('.af-copy-fallback')?.remove();const box=document.createElement('textarea');box.className='af-copy-fallback';box.value=text;box.setAttribute('aria-label','Select and copy');b.after(box);box.focus();box.select();}}
 });draw();
}
window.ArtifactReadsMount=mount;
})();

(() => {
'use strict';
const D = window.ArtifactFieldGuideData;
if (!D) return;
const esc = value => String(value ?? '').replace(/[&<>"']/g, x => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]));
const json = value => JSON.stringify(value,null,2);
const $ = (root,selector) => root.querySelector(selector);
const tag = (text,tone='') => '<span class="af-tag '+tone+'">'+esc(text)+'</span>';
const button = (label,attr='',active=false) => '<button type="button" '+attr+' class="af-button'+(active?' is-active':'')+'" aria-pressed="'+active+'">'+esc(label)+'</button>';
const short = v => typeof v==='string'&&v.length>180 ? v.slice(0,150)+'… ('+v.length.toLocaleString()+' characters; full value in Raw)' : typeof v==='object' ? json(v) : String(v??'null');
const tone = status => /fail|unimplemented|exception|omitted|discrepancy/i.test(status)?'warn':/native|matched/i.test(status)?'good':'';
const params = () => new URL(location.href).searchParams;
function remember(key,value) {try {const u=new URL(location.href);u.searchParams.set(key,String(value));history.replaceState(null,'',u);}catch{}}
function focusRestorer(root){
  const active=document.activeElement;if(!root.contains(active))return ()=>{};
  const attr=[...active.attributes].find(a=>a.name.startsWith('data-'));
  const selector=active.id?'#'+CSS.escape(active.id):attr?'['+attr.name+'="'+CSS.escape(attr.value)+'"]':null;
  return ()=>{if(selector)$(root,selector)?.focus({preventScroll:true});};
}
function tree(value,name='object',depth=0) {
  if (value!==null&&typeof value==='object') {
    const entries=Object.entries(value);
    return '<details class="af-node" '+(depth<2?'open':'')+'><summary><span class="af-key">'+esc(name)+'</span> '+tag(Array.isArray(value)?'array · '+entries.length:'object · '+entries.length)+'</summary><div class="af-branches">'+entries.map(([k,v])=>tree(v,Array.isArray(value)?'['+k+']':k,depth+1)).join('')+(entries.length?'':'<span class="af-muted">Empty</span>')+'</div></details>';
  }
  return '<div class="af-leaf"><span class="af-key">'+esc(name)+'</span><span class="af-value '+(typeof value==='number'?'number':'')+'">'+esc(short(value))+'</span><small>'+esc(value===null?'null':typeof value)+'</small></div>';
}
function matrix(values, options={}) {
  const width=Math.max(0,...values.map(r=>r.length)), before=options.before||[];
  if(!width)return '<div class="af-empty">No cells yet. The container exists before the data.</div>';
  return '<div class="af-grid-wrap"><table class="af-sheet" aria-label="'+esc(options.label||'Recorded cell values')+'"><colgroup><col style="width:29px">'+Array.from({length:width},()=>'<col>').join('')+'</colgroup><thead><tr><th></th>'+Array.from({length:width},(_,i)=>'<th>'+String.fromCharCode(65+i)+'</th>').join('')+'</tr></thead><tbody>'+values.map((row,r)=>'<tr><th>'+String(r+1)+'</th>'+Array.from({length:width},(_,c)=>{
    const val=row[c],address=String.fromCharCode(65+c)+(r+1),formula=options.formulas?.[r]?.[c]||'',changed=options.changes&&(json(val)!==json(before[r]?.[c])||formula!==(options.beforeFormulas?.[r]?.[c]||''));
    return '<td class="'+(typeof val==='number'?'af-number ':'')+(formula?'af-formula ':'')+(changed?'af-changed ':'')+(r===0&&options.styled?'af-styled ':'')+'">'+(options.selectable?'<button type="button" data-cell="'+address+'" title="'+esc(formula||address)+'">'+esc(val??'')+'</button>':esc(val??''))+'</td>';
  }).join('')+'</tr>').join('')+'</tbody></table></div>';
}
function easyArray(value){return Array.isArray(value)&&(value.every(Array.isArray)||value.every(x=>x===null||typeof x!=='object'));}
function readable(value) {
  if(Array.isArray(value)&&!value.length)return '<span class="af-muted">Empty list</span>';
  if(Array.isArray(value)&&value.every(Array.isArray))return matrix(value);
  if(Array.isArray(value)&&value.every(x=>x===null||typeof x!=='object'))return '<ol class="af-value-list">'+value.map(x=>'<li>'+esc(short(x))+'</li>').join('')+'</ol>';
  if(value!==null&&typeof value==='object')return '<dl class="af-fields">'+Object.entries(value).map(([key,v])=>'<div><dt>'+esc(key)+'</dt><dd>'+(easyArray(v)?readable(v):v!==null&&typeof v==='object'?'<details><summary>'+esc(Array.isArray(v)?v.length+' items':'Inspect fields')+'</summary>'+readable(v)+'</details>':esc(short(v)))+'</dd></div>').join('')+'</dl>';
  return '<div class="af-scalar">'+esc(short(value))+'</div>';
}
function views(active,group) {return '<div class="af-segments" aria-label="Representation">'+['readable','tree','raw'].map(x=>button(x[0].toUpperCase()+x.slice(1),'data-'+group+'-view="'+x+'"',x===active)).join('')+'</div>';}
function codeBlock(code) {return '<pre class="af-code"><code>'+esc(code??'No code in this record.')+'</code></pre>';}
function outputView(value,view,label) {return view==='raw'?codeBlock(json(value)):view==='tree'?'<div class="af-tree">'+tree(value,label)+'</div>':readable(value);}
async function copy(text,b) {try{await navigator.clipboard.writeText(text);const old=b.textContent;b.textContent='Copied';setTimeout(()=>b.textContent=old,1600);}catch{const old=b.parentNode.querySelector('.af-copy-fallback');if(old)old.remove();const input=document.createElement('textarea');input.className='af-copy-fallback';input.value=text;input.setAttribute('aria-label','Select and copy this text');b.parentNode.append(input);input.focus();input.select();}}
function download(record) {const blob=new Blob([json(record)],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=(record.id||'record')+'.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
function openImage(src,label) {
  const dialog=document.createElement('dialog');dialog.className='af-dialog';
  dialog.innerHTML='<div class="af-dialog-head"><strong>'+esc(label)+'</strong><button type="button" aria-label="Close image">Close ×</button></div><img src="'+esc(src)+'" alt="'+esc(label)+'">';
  document.body.append(dialog);dialog.querySelector('button').onclick=()=>dialog.close();dialog.addEventListener('close',()=>dialog.remove());dialog.addEventListener('click',e=>{if(e.target===dialog)dialog.close();});dialog.showModal();
}
function renderRecord(record,view) {
  let visual='';
  if(view==='readable'&&record.image){
    const cropped=record.kind==='shape'&&record.title!=='custom',i=record.atlasIndex;
    const preview=cropped?'<svg class="af-shape-preview" viewBox="'+(40+(i%4)*152)+' '+(20+Math.floor(i/4)*136)+' 148 120" role="img" aria-label="Captured geometry: '+esc(record.title)+'"><image href="'+esc(record.image)+'" width="648" height="1380"/></svg>':'<img src="'+esc(record.image)+'" alt="Captured Artifact Tool rendering: '+esc(record.title)+'" loading="lazy">';
    visual='<figure class="af-capture"><button type="button" data-image="'+esc(record.image)+'" data-image-label="'+esc(record.title)+'">'+preview+'</button><figcaption>'+(cropped?'Original rendered geometry · select for full contact sheet':'Original package render · select to enlarge')+'</figcaption></figure>';
  }
  if(view==='readable'&&record.cells)visual+='<details class="af-observed-cells"><summary>Inspect cells saved after this operation</summary>'+readable(record.cells)+'</details>';
  return visual+'<div class="af-io"><section><div class="af-small-heading">'+(record.kind==='guidance'?'Purpose and application':'Input')+'</div>'+outputView(record.input,view,'input')+'</section><section><div class="af-small-heading">'+(record.kind==='guidance'?'Supplied guidance':'Observed output')+'</div>'+outputView(record.result,view,'output')+'</section></div>';
}
function recordActions(record) {return '<div class="af-record-actions">'+button(record.kind==='guidance'?'Copy document':'Copy code','data-copy-code')+button('Download record','data-download')+button('Copy link','data-copy-link')+(record.source?'<a class="af-source" href="'+esc(record.source)+'">'+(record.kind==='guidance'?'Supplied document':'Original receipt')+' ↗</a>':'')+'</div>';}
function wireCommon(root,getRecord) {root.addEventListener('click',e=>{
  const b=e.target.closest('button');if(!b)return;
  const record=getRecord();
  if(b.hasAttribute('data-copy-code'))copy(record.code||json(record.input),b);
  if(b.hasAttribute('data-copy-link'))copy(location.href,b);
  if(b.hasAttribute('data-download'))download(record);
  if(b.dataset.image)openImage(b.dataset.image,b.dataset.imageLabel);
});}
function initJourney(root) {
  let quantity=Number(params().get('quantity'))||2,step=Number(params().get('step')??2),view='readable',phase='after',selected='D5';
  if(![2,5,10].includes(quantity))quantity=2;if(!Number.isInteger(step)||step<0||step>7)step=2;
  const current=()=>D.walkthrough.runs.find(x=>x.quantity===quantity).steps[step];
  function draw(){
    const restore=focusRestorer(root);
    const run=D.walkthrough.runs.find(x=>x.quantity===quantity),item=current(),state=item[phase],sheet=state.worksheets[0],before=item.before.worksheets[0],total=sheet?.values?.[4]?.[3];
    root.innerHTML='<div class="af-toolbar"><div>'+tag('Recorded execution','good')+' <span class="af-muted">Eight steps. One workbook.</span></div><label>Standard quantity <select id="af-quantity" aria-label="Recorded quantity case">'+[2,5,10].map(q=>'<option '+(quantity===q?'selected':'')+' value="'+q+'">'+q+' units</option>').join('')+'</select></label></div>'+
      '<nav class="af-stages" aria-label="Workbook stages">'+run.steps.map((s,i)=>'<button type="button" data-step="'+i+'" aria-current="'+(i===step?'step':'false')+'" class="'+(i===step?'is-active':'')+'"><span>'+String(i+1).padStart(2,'0')+'</span>'+esc(s.label)+'</button>').join('')+'</nav>'+
      '<div class="af-journey-layout"><div><div class="af-kicker">Step '+(step+1)+' / 8</div><h3 class="af-title">'+esc(item.label)+'</h3><p class="af-intent">'+esc(item.intent)+'</p><div class="af-metrics"><div><small>Worksheets</small><strong>'+state.worksheets.length+'</strong></div><div><small>Total revenue</small><strong>'+(typeof total==='number'?'$'+total:'—')+'</strong></div><div><small>Native objects</small><strong>'+((sheet?.tables?.length||0)+(sheet?.charts?.length||0))+'</strong></div></div>'+
      '<div class="af-small-heading">The call</div>'+codeBlock(item.code)+'<div class="af-record-actions">'+button('Previous','data-prev '+(step===0?'disabled':''))+button('Next step','data-next '+(step===7?'disabled':''))+button('Copy code','data-copy-code')+'</div></div>'+
      '<section class="af-state"><div class="af-state-toolbar">'+views(view,'journey')+'<div class="af-segments">'+button('Before','data-phase="before"',phase==='before')+button('After','data-phase="after"',phase==='after')+'</div></div>'+
      (view==='readable'?'<div class="af-workbook-strip"><span class="af-file-icon">▦</span> Sales workbook <span>'+esc(phase)+' '+esc(item.label.toLowerCase())+'</span></div>'+
        (sheet?matrix(sheet.values,{before:before?.values,formulas:sheet.formulas,beforeFormulas:before?.formulas,changes:phase==='after',selectable:true,styled:step>=3&&phase==='after'}):'<div class="af-empty">The workbook has no worksheets yet.</div>')+
        '<div class="af-cell-read" id="af-cell-read">Select a cell to inspect its value and formula.</div>'+
        (sheet?'<div class="af-objects">'+tag('Sales worksheet')+(sheet.tables.length?tag('Table: SalesData'):'')+(sheet.charts.length?tag('Chart: bar'):'')+'</div>':''):
        outputView({input:item.input,returned:item.returned,workbook:state},view,'record'))+
        '<div class="af-state-foot"><span class="af-change-key"></span> Highlight = changed value or formula. The grid is a readable projection of saved cells.</div></section></div>'+
      (step>=6?'<div class="af-file-output"><div><b>Inspect the package’s own output</b><p>The PNG captures its rendering. The XLSX retains editable workbook objects.</p></div>'+button('Open captured PNG','data-image="files/sales-'+quantity+'.png" data-image-label="Sales case: '+quantity+' units"')+'<a class="af-button" href="files/sales-'+quantity+'.xlsx" download>Download XLSX ↗</a></div>':'');
    if(view==='readable')selectCell(selected);restore();
  }
  function selectCell(address){selected=address;const s=current()[phase].worksheets[0],c=address.charCodeAt(0)-65,r=Number(address.slice(1))-1;const box=$ (root,'#af-cell-read');if(!box||!s)return;root.querySelectorAll('[data-cell]').forEach(x=>x.classList.toggle('is-selected',x.dataset.cell===address));box.innerHTML='<b>Sales!'+esc(address)+'</b><span>Value: '+esc(s.values?.[r]?.[c]??'empty')+'</span><code>'+esc(s.formulas?.[r]?.[c]||'No stored formula')+'</code>';}
  root.addEventListener('change',e=>{if(e.target.id==='af-quantity'){quantity=Number(e.target.value);remember('quantity',quantity);draw();}});
  root.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;
    if(b.dataset.step!==undefined){step=Number(b.dataset.step);remember('step',step);draw();}
    if(b.hasAttribute('data-next')&&step<7){step++;remember('step',step);draw();}
    if(b.hasAttribute('data-prev')&&step>0){step--;remember('step',step);draw();}
    if(b.dataset.journeyView){view=b.dataset.journeyView;draw();}
    if(b.dataset.phase){phase=b.dataset.phase;draw();}
    if(b.dataset.cell)selectCell(b.dataset.cell);
  });
  wireCommon(root,()=>({...current(),id:'walkthrough-'+quantity+'-'+step}));draw();
}
const familyRecords=id=>[...D.recipes,...D.operations,...D.apis].filter(x=>x.family===id);
function initFamilies(root) {
  let family=params().get('family')||'ranges',view='readable',id=params().get('action');
  if(!D.families.some(x=>x.id===family))family='ranges';
  let records=familyRecords(family);if(!records.some(x=>x.id===id))id=records[0]?.id;
  const selected=()=>records.find(x=>x.id===id)||records[0];
  function draw(){
    const restore=focusRestorer(root);
    records=familyRecords(family);if(!records.some(x=>x.id===id))id=records[0]?.id;
    const f=D.families.find(x=>x.id===family),r=selected();
    root.innerHTML='<div class="af-explorer-layout"><nav class="af-family-nav" aria-label="Action families"><div class="af-small-heading">15 action families</div>'+D.families.map((f,i)=>'<button type="button" data-family="'+f.id+'" class="'+(f.id===family?'is-active':'')+'" aria-pressed="'+(f.id===family)+'"><span class="af-family-number">'+String(i+1).padStart(2,'0')+'</span><span>'+esc(f.label)+'</span><small>'+familyRecords(f.id).length+'</small></button>').join('')+'</nav>'+
      '<div class="af-family-main"><div class="af-kicker">'+esc(f.label)+'</div><h3 class="af-title">'+esc(f.intent)+'</h3><p class="af-intent">'+esc(f.why)+'</p><div class="af-object-path">'+f.tree.split(' → ').map(x=>'<span>'+esc(x)+'</span>').join('<i>→</i>')+'</div>'+
      '<details class="af-settings"><summary>Actions and settings <span>'+esc(f.actions)+'</span></summary><div class="af-setting-list">'+f.settings.map(([name,type,meaning])=>'<div><code>'+esc(name)+'</code>'+tag(type)+'<p>'+esc(meaning)+'</p></div>').join('')+'</div></details>'+
      '<label class="af-example-label">Explore an example<select id="af-example" aria-label="Action example">'+records.map(x=>'<option value="'+x.id+'" '+(id===x.id?'selected':'')+'>'+esc(x.kind.toUpperCase()+' · '+x.title)+'</option>').join('')+'</select></label>'+
      '<div class="af-record-heading"><div><h4>'+esc(r.title)+'</h4>'+tag(r.kind)+' '+tag(r.status,tone(r.status))+'</div>'+views(view,'family')+'</div><p class="af-record-summary">'+esc(r.summary)+'</p>'+
      renderRecord(r,view)+(r.notes?.length?'<details class="af-code-door"><summary>Supplied settings and notes</summary><ul>'+r.notes.map(n=>'<li>'+esc(n)+'</li>').join('')+'</ul></details>':'')+'<details class="af-code-door"><summary>Code and original evidence</summary><p>Guided snippets show the essential pattern. Raw input/output and the linked receipt preserve the captured context.</p>'+codeBlock(r.code)+recordActions(r)+'</details></div></div>';restore();
  }
  root.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;if(b.dataset.family){family=b.dataset.family;id=null;remember('family',family);remember('action','');draw();}if(b.dataset.familyView){view=b.dataset.familyView;draw();}});
  root.addEventListener('change',e=>{if(e.target.id==='af-example'){id=e.target.value;remember('action',id);remember('family',family);draw();}});
  wireCommon(root,selected);draw();
}
const categories=[['apis','API examples'],['operations','Operations'],['formulas','Formulas'],['charts','Chart types'],['shapes','Shape types'],['support','Code & guidance']].map(([k,t])=>[k,t,D[k].length]);
function initCatalogue(root) {
  let cat=params().get('catalogue')||'formulas',id=params().get('record'),query='',filter='',page=0,view='readable';if(!categories.some(([k])=>k===cat))cat='formulas';
  const all=()=>D[cat],matches=()=>all().filter(x=>(!filter||x.status===filter)&&(!query||(x.title+' '+x.summary+' '+(x.name||'')+' '+(x.category||'')+' '+(x.code||'')).toLowerCase().includes(query.toLowerCase())));
  if(!all().some(x=>x.id===id))id=cat==='formulas'?'fx.SUM':all()[0].id;
  const selected=()=>all().find(x=>x.id===id)||all()[0];
  function draw(keepFocus=false) {
    const restore=focusRestorer(root);
    const results=matches(),pages=Math.max(1,Math.ceil(results.length/10));if(page>=pages)page=0;
    const list=results.slice(page*10,page*10+10),r=selected();
    root.innerHTML='<div class="af-catalog-tabs" role="group" aria-label="Catalogue">'+categories.map(([key,label,count])=>button(label+' '+count,'data-catalogue="'+key+'"',cat===key)).join('')+'</div>'+
      '<div class="af-catalog-tools"><label class="af-search-label">Search this catalogue<input id="af-catalogue-search" type="search" placeholder="Try OFFSET, chart labels, export…" value="'+esc(query)+'" autocomplete="off"></label><label>Observed outcome<select id="af-outcome"><option value="">All outcomes</option>'+[...new Set(all().map(x=>x.status))].map(x=>'<option '+(filter===x?'selected':'')+'>'+esc(x)+'</option>').join('')+'</select></label></div>'+
      '<div class="af-catalog-layout"><aside class="af-results"><div class="af-results-count" role="status">'+results.length+' of '+all().length+' records</div><div class="af-result-list">'+(list.length?list.map(x=>'<button type="button" data-record="'+esc(x.id)+'" class="'+(x.id===id?'is-active':'')+'" aria-pressed="'+(x.id===id)+'"><b>'+esc(x.title)+'</b><small>'+esc(x.category||x.name||x.kind)+'</small>'+tag(x.status,tone(x.status))+'</button>').join(''):'<div class="af-empty">No records match.<br>'+button('Clear filters','data-clear')+'</div>')+'</div><div class="af-pagination">'+button('←','data-page-prev aria-label="Previous results page" '+(page===0?'disabled':''))+'<span>'+(page+1)+' / '+pages+'</span>'+button('→','data-page-next aria-label="Next results page" '+(page+1===pages?'disabled':''))+'</div></aside>'+
      '<section class="af-catalog-detail" aria-label="Selected catalogue record"><div class="af-record-heading"><div><div class="af-kicker">'+esc(r.category||r.kind)+'</div><h3 class="af-title">'+esc(r.title)+'</h3>'+tag(r.status,tone(r.status))+'</div>'+views(view,'catalogue')+'</div><p>'+esc(r.summary)+'</p>'+
      (cat==='formulas'?'<p class="af-note">Common-fixture result; a returned value is not an independent correctness verdict.</p>':'')+
      renderRecord(r,view)+
      (r.oracle?'<div class="af-oracle"><strong>Independent check · separate fixture</strong>'+tag(r.oracle.matches?'Matched expectation':'Discrepancy',r.oracle.matches?'good':'warn')+'<p><code>'+esc(r.oracle.formula)+'</code></p><div class="af-oracle-values"><span>Expected <b>'+esc(short(r.oracle.expected))+'</b></span><span>Observed <b>'+esc(short(r.oracle.actual))+'</b></span></div></div>':'')+
      (r.notes?.length?'<details class="af-code-door"><summary>Arguments and supplied notes</summary><ul>'+r.notes.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul></details>':'')+
      '<details class="af-code-door"><summary>Code and saved record</summary>'+codeBlock(r.code||json(r.input))+recordActions(r)+'</details></section></div>';
    const tabs=$(root,'.af-catalog-tabs'),activeTab=$(tabs,'.is-active');tabs.scrollLeft=activeTab.offsetLeft-tabs.offsetLeft-(tabs.clientWidth-activeTab.offsetWidth)/2;
    if(!keepFocus)restore();
    if(keepFocus){const input=$(root,'#af-catalogue-search');input.focus();input.setSelectionRange(query.length,query.length);}
  }
  root.addEventListener('keydown',e=>{if(e.target.id==='af-catalogue-search'&&e.key==='Enter'&&matches().length){id=matches()[0].id;remember('catalogue',cat);remember('record',id);draw();$(root,'.af-catalog-detail .af-title').setAttribute('tabindex','-1');$(root,'.af-catalog-detail .af-title').focus({preventScroll:true});}});
  root.addEventListener('input',e=>{if(e.target.id==='af-catalogue-search'){query=e.target.value;page=0;draw(true);}});
  root.addEventListener('change',e=>{if(e.target.id==='af-outcome'){filter=e.target.value;page=0;draw();}});
  root.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;
    if(b.dataset.catalogue){cat=b.dataset.catalogue;query='';filter='';page=0;id=all()[0].id;remember('catalogue',cat);remember('record',id);draw();}
    if(b.dataset.record){id=b.dataset.record;remember('record',id);remember('catalogue',cat);draw();}
    if(b.dataset.catalogueView){view=b.dataset.catalogueView;draw();}
    if(b.hasAttribute('data-page-next')){page++;draw();}if(b.hasAttribute('data-page-prev')){page--;draw();}
    if(b.hasAttribute('data-clear')){query='';filter='';page=0;draw();}
  });
  wireCommon(root,selected);draw();
}
function initBoundaries(root) {
  let id='totals',checkpoint=0,view='readable';const current=()=>D.limits.find(x=>x.id===id);
  function draw(){const restore=focusRestorer(root),r=current();root.innerHTML='<div class="af-boundary-nav" aria-label="Boundary examples">'+D.limits.map(x=>button(x.title,'data-boundary="'+x.id+'"',id===x.id)).join('')+'</div>'+
    '<div class="af-boundary-content"><div class="af-kicker">What the experiment teaches</div><h3 class="af-title">'+esc(r.title)+'</h3><div class="af-contrast"><section><small>The expectation</small><p>'+esc(r.expected)+'</p></section><section><small>The observation</small><p>'+esc(r.observed)+'</p></section></div>'+
    (r.states?'<div class="af-checkpoints" aria-label="Table totals checkpoints">'+r.states.map(([label,val],i)=>'<button type="button" data-checkpoint="'+i+'" aria-pressed="'+(i===checkpoint)+'" class="'+(i===checkpoint?'is-active':'')+'"><small>'+esc(label)+'</small><b>'+val+'</b></button>').join('')+'</div><p class="af-checkpoint-note">Selected checkpoint: <b>'+esc(r.states[checkpoint][0])+'</b> — reported sum <b>'+r.states[checkpoint][1]+'</b>.</p>':'')+
    '<div class="af-explanation"><div><b>Why it happens</b><p>'+esc(r.why)+'</p></div><div><b>The working pattern</b><p>'+esc(r.repair)+'</p></div></div><details class="af-code-door"><summary>Inspect the captured input and output</summary>'+views(view,'boundary')+renderRecord(r,view)+codeBlock(r.code)+'</details></div>';restore();}
  root.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;if(b.dataset.boundary){id=b.dataset.boundary;checkpoint=0;draw();}if(b.dataset.checkpoint!==undefined){checkpoint=Number(b.dataset.checkpoint);draw();}if(b.dataset.boundaryView){view=b.dataset.boundaryView;draw();$(root,'.af-code-door').open=true;}});draw();
}
function init(){document.querySelectorAll('[data-af-mode]').forEach(root=>{if(root.dataset.afMounted)return;root.dataset.afMounted='true';({journey:initJourney,families:initFamilies,catalogue:initCatalogue,boundaries:initBoundaries})[root.dataset.afMode](root);});}
window.ArtifactFieldGuideInit=init;
if(!window.ArtifactFieldGuideSwapListener){window.ArtifactFieldGuideSwapListener=()=>window.ArtifactFieldGuideInit?.();document.addEventListener('hub:main-swapped',window.ArtifactFieldGuideSwapListener);}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();

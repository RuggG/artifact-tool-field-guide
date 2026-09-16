(()=>{
 const root=document.getElementById('operations-reference');if(!root)return;
 const d=JSON.parse(document.getElementById('operations-reference-data').textContent),detail=document.getElementById('operation-detail');let current=null;
 const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const pre=v=>'<pre tabindex="0"><code>'+esc(typeof v==='string'?v:JSON.stringify(v,null,2))+'</code></pre>';
 function tree(value,key='Response',open=true){
  if(key==='ndjson'&&typeof value==='string'){try{return tree(value.trim().split('\n').filter(Boolean).map(x=>JSON.parse(x)),'ndjson · decoded records',false);}catch{}}
  if(typeof value==='string'&&value.length>240){const n=document.createElement('details');n.className='oref-tree';n.innerHTML='<summary>'+esc(key)+' <small>'+value.length+' characters</small></summary>'+pre(value);return n;}
  if(value===null||typeof value!=='object'){const row=document.createElement('div');row.className='oref-leaf';row.innerHTML='<b>'+esc(key)+'</b><code>'+esc(JSON.stringify(value))+'</code>';return row;}
  const n=document.createElement('details');n.className='oref-tree';n.innerHTML='<summary>'+esc(key)+' <small>'+Object.keys(value).length+(Array.isArray(value)?' items':' fields')+'</small></summary><div></div>';let loaded=false;
  const fill=()=>{if(!loaded){loaded=true;const box=n.lastElementChild;Object.entries(value).forEach(([k,v])=>box.append(tree(v,k,false)));}};
  n.addEventListener('toggle',()=>{if(n.open)fill();});if(open){fill();n.open=true;}return n;
 }
 function modes(side,readable,raw,codeValue){
  side.querySelectorAll('[data-mode]').forEach(b=>b.addEventListener('click',()=>{
   side.querySelectorAll('[data-mode]').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));
   const pane=side.querySelector('.oref-pane');pane.replaceChildren();
   if(b.dataset.mode==='tree')pane.append(tree(raw));else pane.innerHTML=b.dataset.mode==='readable'?readable:pre(codeValue??raw);
  }));
 }
 function nav(side,output=false){return '<nav class="oref-modes" aria-label="'+side+' display"><button data-mode="readable" aria-pressed="true">Readable</button>'+(output?'<button data-mode="tree" aria-pressed="false">Object tree</button>':'')+'<button data-mode="code" aria-pressed="false">'+(side==='output'?'Exact output':'Exact code')+'</button></nav>';}
 function example(id){
  const c=d.examples[id],box=document.getElementById('example-detail');if(!c||!box)return;
  box.innerHTML='<p class="oref-example-intent">'+esc(c.intent)+'</p><div class="oref-pair"><section data-side="input"><h4>Executed input</h4>'+nav('input')+'<div class="oref-pane">'+c.inputReadable+'</div></section><section data-side="output"><h4>'+esc(c.outputLabel)+'</h4>'+nav('output',true)+'<div class="oref-pane">'+c.outputReadable+'</div></section></div>';
  modes(box.querySelector('[data-side="input"]'),c.inputReadable,c.input,c.input);
  modes(box.querySelector('[data-side="output"]'),c.outputReadable,c.output);
  if(c.checks){
   box.insertAdjacentHTML('beforeend','<details id="agent-checks" class="oref-checks"><summary>Agent-written checks and what they found</summary><p>These are extra reads and comparisons chosen by the agent. The operation does not automatically return them.</p><section data-side="checks">'+nav('checks',true)+'<div class="oref-pane">'+c.checkReadable+'</div></section></details>');
   modes(box.querySelector('[data-side="checks"]'),c.checkReadable,c.checks,c.checkCode+'\n\n// Helper definitions used by the checks\n'+c.checkHelpers+'\n// sha() and lastHandle are defined in the full executed script.');
  }
  box.insertAdjacentHTML('beforeend','<div class="oref-observation"><b>Observed in this example</b><p>'+esc(c.finding)+'</p></div><p class="oref-note">'+esc(c.note)+'</p>');
  if(c.image)box.insertAdjacentHTML('beforeend','<figure><a href="'+esc(c.image)+'?raw"><img src="'+esc(c.image)+'" alt="Recorded workbook render for '+esc(c.label)+'" loading="lazy"></a><figcaption>Saved image returned by rendering; click for full size.</figcaption></figure>');
  box.insertAdjacentHTML('beforeend','<p class="oref-evidence-links">'+c.links.map(([label,url])=>'<a href="'+esc(url)+'">'+esc(label)+'</a>').join(' · ')+'</p>');
 }
 function select(id,record=null,scroll=false,update=true){
  const op=d.operations.find(x=>x.id===id);if(!op)return;current=id;
  root.querySelectorAll('[data-operation]').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.operation===id)));
  detail.innerHTML='<header><p class="oref-kicker">Operation reference</p><h3>'+esc(op.label)+'</h3><code class="oref-api">'+esc(op.api)+'</code><p class="oref-does">'+esc(op.does)+'</p></header><dl class="oref-contract">'+[['You supply',op.supply],['What it returns',op.returns],['Agent still decides / checks',op.agent]].map(([k,v])=>'<div><dt>'+esc(k)+'</dt><dd>'+esc(v)+'</dd></div>').join('')+'</dl><p class="oref-detail-note">'+esc(op.detail)+'</p><details id="recorded-examples"><summary>Recorded examples <span>'+op.examples.length+'</span></summary><p class="oref-note">Saved inputs and outputs from the Man Group file or a small new workbook. Each example shows its own recorded state; these controls do not execute edits.</p>'+ (op.examples.length>1?'<label class="oref-select">Example / setting<select id="example-choice">'+op.examples.map(key=>'<option value="'+key+'">'+esc(d.examples[key].label)+'</option>').join('')+'</select></label>':'')+'<div id="example-detail"></div></details><a class="oref-back" href="#operations-reference">↑ All operations</a>';
  const chosen=record&&op.examples.includes(record)?record:op.examples[0];example(chosen);
  const chooser=document.getElementById('example-choice');if(chooser){chooser.value=chosen;chooser.addEventListener('change',()=>{example(chooser.value);history.replaceState(null,'','#ref-'+current+'~'+chooser.value.replace(':','-'));});}
  if(record)document.getElementById('recorded-examples').open=true;
  if(update)history.replaceState(null,'','#ref-'+id+(record?'~'+record.replace(':','-'):''));
  if(scroll)detail.scrollIntoView({behavior:'instant',block:'start'});
 }
 root.querySelectorAll('[data-operation]').forEach(b=>b.addEventListener('click',()=>select(b.dataset.operation,null,true)));
 function hash(initial=false){const h=decodeURIComponent(location.hash.slice(1));
  if(h.startsWith('ref-')){const [op,r]=h.slice(4).split('~');const rec=r?r.replace(/^(op|old|lab)-/,'$1:'):null;select(op,rec,!initial,false);if(initial)requestAnimationFrame(()=>detail.scrollIntoView({block:'start'}));return true;}
  if(d.aliases[h]){select(...d.aliases[h],false,false);if(initial)requestAnimationFrame(()=>detail.scrollIntoView({block:'start'}));else detail.scrollIntoView({block:'start'});return true;}return false;
 }
 window.addEventListener('hashchange',()=>hash());if(!hash(true))select('map',null,false,false);
})();

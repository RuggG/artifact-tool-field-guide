(() => {
'use strict';
const D=window.ArtifactCommonWorkData;if(!D)return;
const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const json=x=>JSON.stringify(x,null,2);
const pre=x=>'<pre class="af-code" tabindex="0"><code>'+esc(typeof x==='string'?x:json(x))+'</code></pre>';
const button=(s,attr,active=false)=>'<button type="button" '+attr+' class="af-button'+(active?' is-active':'')+'" aria-pressed="'+active+'">'+esc(s)+'</button>';
function tree(value,key='Response',depth=0){
  if(key==='ndjson'&&typeof value==='string'){
    try{return tree(value.trim().split('\n').filter(Boolean).map(JSON.parse),'ndjson · decoded records',depth);}catch{}
  }
  if(value!==null&&typeof value==='object'){
    const n=document.createElement('details');n.className='af-node';
    const entries=Object.entries(value);n.innerHTML='<summary><span class="af-key">'+esc(key)+'</span> <span class="af-tag">'+entries.length+(Array.isArray(value)?' items':' fields')+'</span></summary><div class="af-branches"></div>';
    let loaded=false;const fill=()=>{if(loaded)return;loaded=true;const box=n.lastElementChild;for(const [k,v] of entries)box.append(tree(v,k,depth+1));if(!entries.length)box.textContent='Empty';};
    n.addEventListener('toggle',()=>{if(n.open)fill();});if(depth===0){fill();n.open=true;}return n;
  }
  const n=document.createElement(typeof value==='string'&&value.length>300?'details':'div');n.className='af-leaf';
  n.innerHTML=n.tagName==='DETAILS'?'<summary>'+esc(key)+' · '+value.length+' characters</summary>'+pre(value):'<span class="af-key">'+esc(key)+'</span><span class="af-value">'+esc(json(value))+'</span>';
  return n;
}
function init(root){
  if(root.dataset.cwMounted)return;root.dataset.cwMounted='true';
  const params=new URL(location.href).searchParams;
  let id=params.get('common')||'map',example=params.get('example'),view='readable';
  if(!D.operations.some(x=>x.id===id))id='map';
  const selected=()=>{const op=D.operations.find(x=>x.id===id);return id!=='ranges'?op:{...op,does:'Read one cell, a chosen area, a whole tab, or every tab. Choose the kind of answer you need below.',supply:'Choose the cells and whether you want their values, formulas, or both.',returns:'Actual cell contents, with several ways to arrange the answer.',agent:'The code can collect several reads and put each value beside its formula.',detail:'These examples use the same unchanged workbook. The controls show actual recorded requests and results; they do not run new calculations.'};};
  const record=()=>D.examples[example];
  function remember(){const u=new URL(location.href);u.searchParams.set('common',id);u.searchParams.set('example',example);u.hash='common-work';history.replaceState(null,'',u);}
  function focusRestore(){const el=document.activeElement;if(!root.contains(el))return()=>{};const a=[...el.attributes].find(x=>x.name.startsWith('data-cw-'));const selector=el.id?'#'+CSS.escape(el.id):a?'['+a.name+'="'+CSS.escape(a.value)+'"]':null;return()=>{if(selector)root.querySelector(selector)?.focus({preventScroll:true});};}
  function draw(){
    const scroll=root.querySelector('.cw-groups')?.scrollLeft||0,previousGroup=root.querySelector('[data-cw-group][aria-pressed="true"]')?.dataset.cwGroup,restore=focusRestore(),earlierOpen=root.querySelector('.ix-earlier')?.open,op=selected();if(!op.examples.includes(example))example=op.examples[0];const r=record();
    root.dataset.cwCurrentGroup=op.group;
    root.innerHTML='<div class="af-toolbar"><div><span class="af-tag good">23 practical actions</span> <span class="af-muted">48 retained examples + 47 inspection comparisons + 38 reading comparisons</span></div><label>Jump to action<select id="af-common-jump" aria-label="Common action">'+D.groups.map(([g,label])=>'<optgroup label="'+esc(label)+'">'+D.operations.filter(x=>x.group===g).map(x=>'<option value="'+x.id+'" '+(x.id===id?'selected':'')+'>'+esc(x.label)+'</option>').join('')+'</optgroup>').join('')+'</select></label></div>'+
      '<div class="cw-group-strip"><button class="cw-scroll" type="button" data-cw-scroll="-1" aria-label="Scroll families left">‹</button><nav class="cw-groups" aria-label="Common work groups" tabindex="0">'+D.groups.map(([g,label])=>button(label,'data-cw-group="'+g+'"',g===op.group)).join('')+'</nav><button class="cw-scroll" type="button" data-cw-scroll="1" aria-label="Scroll families right">›</button></div>'+
      '<div class="cw-layout"><nav class="af-family-nav cw-actions" aria-label="Actions in this group"><div class="af-small-heading">'+esc(D.groups.find(x=>x[0]===op.group)[1])+'</div>'+D.operations.filter(x=>x.group===op.group).map(x=>'<button type="button" data-cw-action="'+x.id+'" aria-pressed="'+(id===x.id)+'" class="'+(id===x.id?'is-active':'')+'">'+esc(x.label)+'</button>').join('')+'</nav>'+
      '<article class="cw-main" aria-label="Selected common action"><div class="af-kicker">Practical starting point</div><h3 class="af-title">'+esc(op.label)+'</h3><code class="cw-api">'+esc(op.api)+'</code><p class="af-intent">'+esc(op.does)+'</p>'+
      '<dl class="cw-contract">'+[['You supply',op.supply],['What it returns',op.returns],['Agent decides and checks',op.agent]].map(([k,v])=>'<div><dt>'+esc(k)+'</dt><dd>'+esc(v)+'</dd></div>').join('')+'</dl><p class="cw-detail">'+esc(op.detail)+'</p>'+
      '<div class="cw-crosslinks"><span class="af-tag">'+(op.commonApi?'API named in the common quick start':op.id==='runner'?'Execution wrapper':'Additional calculation control')+'</span><a href="?family='+op.family+'&action=#families">Explore the full '+esc(({workbooks:'workbooks and sheets',files:'files and rendering',changes:'changes and sessions',inspection:'inspection',ranges:'cells and ranges',formulas:'calculation',formatting:'formatting',tables:'Excel Tables',charts:'charts'})[op.family])+' family ↗</a></div>'+
      (op.group==='inspect'||id==='ranges'?'<div id="'+(id==='ranges'?'cw-reads':'cw-inspection')+'"></div><details class="ix-earlier"><summary>Earlier recordings · '+op.examples.length+' examples retained</summary>':'')+
      '<div class="cw-example"><label class="af-example-label">Recorded example <select id="af-common-example" aria-label="Recorded common example">'+op.examples.map(key=>'<option value="'+key+'" '+(key===example?'selected':'')+'>'+esc(D.examples[key].label)+'</option>').join('')+'</select></label>'+
      '<div class="af-record-heading"><div><h4>'+esc(r.label)+'</h4><p>'+esc(r.intent)+'</p></div><div class="af-segments" aria-label="Example representation">'+['readable','tree','raw'].map(x=>button(x[0].toUpperCase()+x.slice(1),'data-cw-view="'+x+'"',view===x)).join('')+'</div></div>'+
      '<div class="cw-pair"><section><h5>Executed input</h5><div data-cw-pane="input"></div></section><section><h5>'+esc(r.outputLabel)+'</h5><div data-cw-pane="output"></div></section></div>'+
      (r.checks?'<section class="cw-checks"><h5>Agent-written checks</h5><p>Extra reads and comparisons chosen by the agent. The edit or export does not automatically return this report.</p><div data-cw-pane="checks"></div></section>':'')+
      '<div class="cw-finding"><b>What this example teaches</b><p>'+esc(r.finding)+'</p></div><p class="cw-note">'+esc(r.note)+'</p>'+
      (r.image?'<figure class="af-capture"><a href="'+esc(r.image)+'"><img src="'+esc(r.image)+'" alt="Saved Artifact Tool render: '+esc(r.label)+'" loading="lazy"></a><figcaption>Actual engine render from the recorded example. Open for full size.</figcaption></figure>':'')+
      '<div class="af-record-actions">'+button('Copy executed input','data-cw-copy="input"')+button('Download record','data-cw-download')+button('Copy link','data-cw-copy="link"')+'</div><p class="cw-note">Snippets preserve the recorded context; use the full script for imports, setup and helper definitions.</p><p class="cw-sources">'+r.links.map(([label,path])=>'<a href="'+esc(path)+'">'+esc(label)+'</a>').join(' · ')+'</p></div>'+(op.group==='inspect'||id==='ranges'?'</details>':'')+'</article></div>';
    for(const [side,readable,raw] of [['input',r.inputReadable,r.input],['output',r.outputReadable,r.output],['checks',r.checkReadable,r.checks]]){
      const pane=root.querySelector('[data-cw-pane="'+side+'"]');if(!pane)continue;
      if(view==='readable')pane.innerHTML=readable;
      else if(view==='tree'){
        if(side==='input')pane.innerHTML=pre(raw);else{pane.classList.add('af-tree');pane.append(tree(raw,side));}
      }else pane.innerHTML=pre(raw)+(side==='checks'?'<details class="af-code-door"><summary>Exact check code and helpers</summary>'+pre(r.checkCode+'\n\n'+r.checkHelpers)+'<p>See the full script for sha() and lastHandle.</p></details>':'');
    }
    if(op.group==='inspect')window.ArtifactInspectionMount(root.querySelector('#cw-inspection'),{actionId:id,tree,onNavigate:choose});
    if(id==='ranges')window.ArtifactReadsMount(root.querySelector('#cw-reads'),{tree});
    if(earlierOpen&&root.querySelector('.ix-earlier'))root.querySelector('.ix-earlier').open=true;
    const nav=root.querySelector('.cw-groups'),current=nav.querySelector('[aria-pressed="true"]');nav.scrollLeft=scroll;
    const update=()=>{const overflow=nav.scrollWidth>nav.clientWidth+2;root.querySelector('.cw-group-strip').classList.toggle('has-overflow',overflow);root.querySelector('[data-cw-scroll="-1"]').disabled=nav.scrollLeft<2;root.querySelector('[data-cw-scroll="1"]').disabled=nav.scrollLeft+nav.clientWidth>=nav.scrollWidth-2;};
    nav.addEventListener('scroll',update,{passive:true});if(root._cwResize)root._cwResize.disconnect();root._cwResize=new ResizeObserver(update);root._cwResize.observe(nav);update();
    if(previousGroup!==op.group){const nr=nav.getBoundingClientRect(),cr=current.getBoundingClientRect();if(cr.left<nr.left)nav.scrollLeft-=nr.left-cr.left;if(cr.right>nr.right)nav.scrollLeft+=cr.right-nr.right;}
    update();restore();
  }
  function choose(next,inspectionId){id=next;example=null;view='readable';const u=new URL(location.href);u.searchParams.delete('example');if(inspectionId)u.searchParams.set('inspect',inspectionId);else{u.searchParams.delete('inspect');u.searchParams.delete('iview');}history.replaceState(null,'',u);draw();remember();}
  root.addEventListener('change',e=>{if(e.target.id==='af-common-jump')choose(e.target.value);if(e.target.id==='af-common-example'){example=e.target.value;draw();remember();}});
  root.addEventListener('click',async e=>{
    const b=e.target.closest('button');if(!b)return;
    if(b.dataset.cwScroll){const nav=root.querySelector('.cw-groups');nav.scrollBy({left:Number(b.dataset.cwScroll)*nav.clientWidth*.7,behavior:matchMedia('(prefers-reduced-motion: reduce)').matches?'instant':'smooth'});}
    if(b.dataset.cwGroup)choose(D.operations.find(x=>x.group===b.dataset.cwGroup).id);
    if(b.dataset.cwAction)choose(b.dataset.cwAction);
    if(b.dataset.cwView){view=b.dataset.cwView;draw();}
    if(b.hasAttribute('data-cw-download')){const blob=new Blob([json({action:selected(),example:record(),provenance:D.provenance})],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='common-'+id+'-'+example.replace(':','-')+'.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
    if(b.dataset.cwCopy){remember();const text=b.dataset.cwCopy==='link'?location.href:record().input;try{await navigator.clipboard.writeText(text);const old=b.textContent;b.textContent='Copied';setTimeout(()=>b.textContent=old,1600);}catch{root.querySelector('.af-copy-fallback')?.remove();const input=document.createElement('textarea');input.className='af-copy-fallback';input.value=text;input.setAttribute('aria-label','Select and copy');b.parentNode.append(input);input.focus();input.select();}}
  });draw();
}
function mount(){document.querySelectorAll('#af-common').forEach(init);}
window.ArtifactCommonWorkInit=mount;
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount);else mount();
if(!window.ArtifactCommonWorkSwapListener){window.ArtifactCommonWorkSwapListener=()=>window.ArtifactCommonWorkInit?.();document.addEventListener('hub:main-swapped',window.ArtifactCommonWorkSwapListener);}
})();

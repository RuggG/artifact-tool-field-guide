(()=>{'use strict';
 const start=()=>{
  const root=document.getElementById('af-satellites'),data=window.ArtifactSatelliteDocuments;
  if(!root||!data||root.dataset.ready)return;root.dataset.ready='true';
  const reader=root.querySelector('#satellite-reader'),docs=data.documents;
  const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const num=n=>n.toLocaleString('en-US');
  let id='style',view='readable';
  const current=()=>docs.find(d=>d.id===id)||docs.find(d=>d.id==='style');
  function render(){
   const d=current();id=d.id;
   root.querySelectorAll('[data-sd-doc]').forEach(a=>a.setAttribute('aria-current',String(a.dataset.sdDoc===id)));
   const options=data.groups.map(g=>`<optgroup label="${esc(g.title)}">${docs.filter(x=>x.group===g.id).map(x=>`<option value="${x.id}" ${x.id===id?'selected':''}>${esc(x.title)}</option>`).join('')}</optgroup>`).join('');
   const outline=d.outline.map(h=>`<a style="--sd-depth:${Math.max(0,h.level-1)}" href="?doc=${id}&dsection=${encodeURIComponent(h.anchor)}#satellite-reader">${esc(h.label)}</a>`).join('');
   reader.innerHTML=`<div class="sd-reader-head"><div><h3 class="sd-title">${esc(d.title)}</h3><div class="sd-path">${esc(d.path)}</div></div><label>Read another document<select id="sd-select">${options}</select></label></div>
    <div class="sd-metrics"><span><strong>${num(d.tokens)}</strong> tokens</span><span><strong>${num(d.words)}</strong> words</span><span><strong>${num(d.characters)}</strong> characters</span></div>
    <p class="sd-when"><strong>When the agent reads it:</strong> ${esc(d.when)}</p>
    <div class="sd-toolbar"><div class="sd-views" role="group" aria-label="Document view"><button type="button" data-sd-view="readable" aria-pressed="${view==='readable'}">Rendered</button><button type="button" data-sd-view="raw" aria-pressed="${view==='raw'}">Raw Markdown</button></div><div class="sd-links"><a href="#af-satellites" data-sd-back>All documents ↑</a><a href="${esc(d.source)}" target="_blank" rel="noopener">Source file ↗</a><button type="button" data-sd-download>Download .md</button></div></div>
    ${view==='readable'?`<details class="sd-outline"><summary>Contents · ${d.outline.length} headings</summary><nav aria-label="Document contents">${outline}</nav></details><article class="sd-document" aria-label="Full document">${d.html}</article>`:`<pre class="sd-raw" aria-label="Complete original Markdown">${esc(d.markdown)}</pre>`}
    <p class="sd-source-note">Source copy: 15 September 2026. Counts include the original Markdown formatting, spaces and line breaks.</p>`;
   reader.querySelectorAll('.sd-document table').forEach(t=>{const w=document.createElement('div');w.className='sd-table-scroll';w.tabIndex=0;w.setAttribute('role','region');w.setAttribute('aria-label','Scrollable document table');t.before(w);w.append(t)});
  }
  function openParents(el){for(let p=el.parentElement;p;p=p.parentElement)if(p.tagName==='DETAILS')p.open=true}
  function move(section){
   openParents(reader);const target=section?document.getElementById('sd-'+id+'-'+section):reader;
   requestAnimationFrame(()=>{(target||reader).scrollIntoView({block:'start',behavior:'instant'});if(!section)reader.focus({preventScroll:true})});
  }
  function pick(next,{nextView='readable',section='',scroll=true,push=true}={}){
   id=docs.some(d=>d.id===next)?next:'style';view=nextView==='raw'?'raw':'readable';render();
   const u=new URL(location.href);u.searchParams.set('doc',id);u.searchParams.set('dview',view);u.hash=section?'sd-'+id+'-'+section:'satellite-reader';
   if(section)u.searchParams.set('dsection',section);else u.searchParams.delete('dsection');
   if(push)history.pushState(null,'',u);
   if(scroll)move(section);
  }
  root.addEventListener('click',e=>{
   const a=e.target.closest('a'),b=e.target.closest('button');
   if(a&&(!e.ctrlKey&&!e.metaKey&&!e.shiftKey&&!e.altKey)){
    if(a.hasAttribute('data-sd-back')){e.preventDefault();openParents(root);root.scrollIntoView({block:'start',behavior:'instant'});root.querySelector('[data-sd-doc]').focus({preventScroll:true});return}
    const u=new URL(a.href,location.href),next=u.searchParams.get('doc');
    if(u.pathname===location.pathname&&next&&a.target!=='_blank'){e.preventDefault();pick(next,{section:u.searchParams.get('dsection')||''});return}
   }
   if(b?.dataset.sdView){pick(id,{nextView:b.dataset.sdView,scroll:false});reader.querySelector(`[data-sd-view="${view}"]`).focus({preventScroll:true})}
   if(b?.hasAttribute('data-sd-download')){const d=current(),url=URL.createObjectURL(new Blob([d.markdown],{type:'text/markdown;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download=d.path.split('/').pop();a.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}
  });
  root.addEventListener('change',e=>{if(e.target.id==='sd-select')pick(e.target.value)});
  const linked=()=>location.hash==='#satellite-reader'||location.hash.startsWith('#sd-');
  const restore=()=>{const p=new URLSearchParams(location.search);pick(p.get('doc')||'style',{nextView:p.get('dview')||'readable',section:p.get('dsection')||'',push:false,scroll:linked()})};
  window.addEventListener('popstate',restore);restore();
  // Let the browser and the shared page frame finish their initial anchor jump.
  const settle=()=>{if(linked())requestAnimationFrame(()=>move(new URLSearchParams(location.search).get('dsection')||''))};
  if(document.readyState==='complete')settle();else window.addEventListener('load',settle,{once:true});
 };
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start);else start();
})();

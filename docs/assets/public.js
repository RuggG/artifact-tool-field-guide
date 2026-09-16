(()=>{'use strict';
const root=new URL('../',document.currentScript.src),home=new URL('work/processes/model-training/research/2026-09-15-artifact-tool-field-guide/field-guide.html',root);
let readers=new Set();
function rewrite(a){
 if(a.hasAttribute('data-public-raw')||a.hasAttribute('download'))return;
 const original=a.getAttribute('href');if(!original||original.startsWith('#')||original.startsWith('blob:'))return;
 let u;try{u=new URL(original,location.href)}catch{return}
 if(u.hostname==='hub.lynott.co'){u=new URL(u.pathname.replace(/^\//,'')+u.search+u.hash,root);a.href=u.href;}
 if(u.origin!==root.origin)return;
 if(/^\/(work|skills|system)\//.test(u.pathname)){u=new URL(u.pathname.slice(1)+u.search+u.hash,root);a.href=u.href;}
 if(!u.pathname.startsWith(root.pathname))return;
 const p=decodeURIComponent(u.pathname.slice(root.pathname.length));
 if(readers.has(p)&&!u.search.includes('raw')&&!u.search.includes('download')){u.pathname+='.html';a.href=u.href;}
}
function scan(node){if(node.nodeType!==1)return;if(node.matches('a[href]'))rewrite(node);node.querySelectorAll('a[href]').forEach(rewrite);}
fetch(new URL('assets/documents.json',root)).then(r=>{if(!r.ok)throw Error('Document index unavailable');return r.json()}).then(paths=>{readers=new Set(paths);scan(document.body)});
new MutationObserver(records=>{for(const r of records){if(r.type==='attributes')rewrite(r.target);else r.addedNodes.forEach(scan)}}).observe(document.body,{childList:true,subtree:true});
document.addEventListener('click',e=>{const a=e.target.closest('a[href]');if(a)rewrite(a)},true);
window.PublicGuide={root:root.href,home:home.href};
})();

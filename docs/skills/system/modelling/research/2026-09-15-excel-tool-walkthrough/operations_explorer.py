from html import escape as e
from pathlib import Path
import json,sys,base64
sys.path.insert(0,str(Path(__file__).parent))
from man_group_walkthrough import json_tree,code,table,fmt,CSS as BASECSS
from pagekit_components import ComponentResult

CSS='''
.ops{--ink:#234759;--line:#ccdce3;font-size:13px;line-height:1.6}.ops button{font:inherit;cursor:pointer;padding:8px 12px;border:1px solid var(--line);border-radius:5px;background:white;color:var(--ink)}.ops button[aria-pressed=true]{background:#1a6570;color:white;border-color:#1a6570}.ops-groups,.ops-choices{display:flex;flex-wrap:wrap;gap:7px;margin:12px 0}.ops-choices{background:#f3f7f8;padding:10px;border-radius:6px}.ops [hidden]{display:none!important}.ops-panel{border:1px solid var(--line);border-radius:8px;overflow:hidden}.ops-head{padding:22px;background:#f3f8f8}.ops-head h3{font-size:23px;margin:0 0 8px}.ops-head p{margin:8px 0!important}.ops-outcome{border-left:3px solid #23818b;padding-left:12px;font-size:15px}.ops-pair{display:grid;grid-template-columns:1fr 1fr}.ops-side{padding:20px;min-width:0}.ops-side+div{border-left:1px solid var(--line);background:#fafcfc}.ops-side h4{font-size:12px;letter-spacing:.05em;text-transform:uppercase;margin:0 0 12px;color:#28616a}.ops small,.ops-caption{font-size:12px;color:#526b77}.ops pre{white-space:pre-wrap;overflow-wrap:anywhere;max-height:450px;overflow:auto;background:#edf3f5;color:#213e50;padding:13px;font-size:11px;line-height:1.5}.ops details{border:1px solid #d8e3e7;border-radius:5px;padding:8px 10px;margin:10px 0}.ops summary{cursor:pointer}.ops .j-node{padding:6px;margin:5px 0;background:#f1f6f8}.ops .j-leaf{padding:5px 0;overflow-wrap:anywhere}.ops .j-key{font-weight:600;margin-right:10px}.ops .j-leaf code{white-space:pre-wrap;overflow-wrap:anywhere}.ops-tree{max-height:440px;overflow:auto}.ops-lesson{padding:16px 22px;border-top:1px solid var(--line)}.ops .xio-grid-scroll{max-height:350px;overflow:auto}.ops .tk-wrap{overflow:auto;max-width:100%}.ops table.tk,.ops table{width:100%;min-width:0!important;font-size:12px}.ops th,.ops td{white-space:normal!important;overflow-wrap:anywhere;padding:7px}.ops .native-return{background:#edf4f6;padding:12px;margin-top:12px}.ops figure{margin:15px 0}.ops img{width:100%;height:auto}.ops .ops-input-nav{display:flex;gap:6px;margin:10px 0}.ops .ops-input-nav button{font-size:11px;padding:5px 8px}.ops .ops-raw{padding:0 20px}.ops .ops-raw pre{max-height:500px}.ops .ops-side .ops-readable{font-size:13px}.ops select{max-width:100%}@media(max-width:850px){.ops-pair{grid-template-columns:1fr}.ops-side+div{border-left:0;border-top:1px solid var(--line)}}
'''
JS='''(()=>{const root=document.getElementById('operations-lab');if(!root)return;
function choose(id,hash){const p=root.querySelector('#op-'+id);if(!p)return;const g=p.dataset.group;root.querySelectorAll('.ops-panel').forEach(x=>x.hidden=x!==p);root.querySelectorAll('[data-op-group]').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.opGroup===g)));root.querySelectorAll('.ops-choices').forEach(x=>x.hidden=x.dataset.group!==g);root.querySelectorAll('[data-op-id]').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.opId===id)));if(hash)history.replaceState(null,'','#op-'+id);}
root.querySelectorAll('[data-op-group]').forEach(b=>b.onclick=()=>choose(root.querySelector('.ops-choices[data-group="'+b.dataset.opGroup+'"] [data-op-id]').dataset.opId,true));root.querySelectorAll('[data-op-id]').forEach(b=>b.onclick=()=>choose(b.dataset.opId,true));
root.querySelectorAll('[data-op-mode]').forEach(b=>b.onclick=()=>{const s=b.closest('.ops-side');s.querySelectorAll('[data-op-mode]').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));s.querySelectorAll('[data-op-pane]').forEach(x=>x.hidden=x.dataset.opPane!==b.dataset.opMode);});
function hash(){if(location.hash.startsWith('#op-')){const id=location.hash.slice(4);choose(id,false);const p=document.getElementById('op-'+id);let n=p?.parentElement;while(n){if(n.tagName==='DETAILS')n.open=true;n=n.parentElement;}}}window.addEventListener('hashchange',hash);hash();})();'''

def scalar_rows(obj,prefix=''):
    if not isinstance(obj,dict):return []
    out=[]
    for k,v in obj.items():
        if isinstance(v,dict) and 'address' in v and 'value' in v:out.append([e(prefix+k+' · '+v['address']),e(fmt(v['value']))])
        elif v is None or isinstance(v,(str,int,float,bool)):out.append([e(prefix+k),e(fmt(v))])
        elif isinstance(v,dict) and k not in ('metadata','checkErrorScan','errorScan','styles','formatStillPresent','drawingInspection'):
            out+=scalar_rows(v,prefix+k+' / ')
    return out

def observed(c):
    def unpack(x):
        if not isinstance(x,dict):return {}
        return x.get('selectedReads',x.get('memory',x))
    a=unpack(c.get('before',{}));b=unpack(c.get('after',{}));rows=[]
    if 'margin' in b and 'annual' in b:
        for key,label in [('margin','Fee margin in memory (bp)'),('annual','Total fees in memory ($m)')]:
            rows.append([label, e(fmt(a[key]['value'])) if key in a else 'Not read',e(fmt(b[key]['value']))])
        if 'quarters' in b:
            rows.append(['First-quarter fees ($m)',e(fmt(a['quarters']['values'][0][4])) if 'quarters' in a else 'Not read',e(fmt(b['quarters']['values'][0][4]))])
        h=table(['Cell chosen by the agent','Before','After'],rows)
        raw=c.get('after',{})
        if 'onDisk' in raw:h+='<p><b>Saved file:</b> fee margin is still '+e(fmt(raw['onDisk']['value']))+'bp. File checksum '+('unchanged' if c['before']['diskSha256']==raw['diskSha256'] else 'changed')+'. The exact checksum is in the Object tree.</p>'
        return h
    if c['id']=='new-recalculation':return table(['Selected observation','Before','After'],[['Quarterly fees ($m)',e(fmt(c['before']['value'])),e(fmt(b['fees']['value']))],['Expected from separate arithmetic','—',e(fmt(b['independentExpected']))],['Explicit recalculate() called','—','No']])
    if c['id']=='delete-table':return table(['Selected observation','Before','After'],[['Formal Excel Tables',str(a['tableCount']),str(b['tableCount'])],['Underlying cell values','Present','Unchanged' if a['cells']['values']==b['cells']['values'] else 'Changed']])
    if c['id']=='overlap-table':return table(['Agent-specified or observed item','Value'],[['Existing Table range','B20:D24'],['Requested Table range','D22:F25'],['Overlap chosen for this test','D22:D24'],['Table names returned by our read',e(', '.join(b['tableNames']))],['Experiment saved to XLSX','No']])
    if c['group']=='tables':
        cells=b.get('cells',b.get('values',{})).get('values',[])
        h=('<p>Formal Table count: <b>'+str(b['tableCount'])+'</b></p>') if 'tableCount' in b else ''
        if cells:h+=table([str(x) for x in cells[0]],[[e(fmt(v)) for v in row] for row in cells[1:]])
        if h:return h
    if 'ndjson' in b:
        records=[json.loads(x) for x in b['ndjson'].splitlines() if x]
        return table(['Tab','Occupied range'],[[e(x.get('name','')),e(x.get('range',''))] for x in records if x.get('kind')=='sheet'])
    rows=scalar_rows(c.get('after',{}))
    labels={'file':'Saved file','bytes':'File size (bytes)','sha256':'Saved file checksum','reopenedMargin · C4':'Saved fee margin (bp)','reopenedTableCount':'Formal Tables after reopening','sourceUnchanged':'Original source unchanged','quarterOne · F9':'Fees after reopening and editing ($m)','total · C15':'Total fees after reopening and editing ($m)','independentExpectedQuarterOne':'Expected first-quarter fees ($m)','savedFileStillHasMargin · C4':'Margin still on disk (bp)','explicitRecalculateCalled':'Explicit recalculate() called','independentExpected':'Expected fees from separate arithmetic ($m)','saved':'Saved image'}
    rows=[[e(labels.get(k,k)),v] for k,v in rows if not ('Sha256' in k or k=='sha256' or k.endswith('/ address'))]
    return table(['Selected observation after the operation','Value'],rows[:12]) if rows else '<p>Open the Object tree to inspect the requested grids, inventory or chart bindings.</p>'

def render(payload,*,body,context):
    cards=payload['cards'];groups=[('files','Memory and files'),('create','Create and extend'),('clear','Clear and rebuild'),('missing','Missing inputs'),('tables','Excel Tables'),('charts','Charts')]
    default='edit-without-saving';h='<div id="operations-lab" class="ops"><nav class="ops-groups" aria-label="Operation families">'
    for g,label in groups:h+=f'<button data-op-group="{g}" aria-pressed="{str(g=="files").lower()}">{label}</button>'
    h+='</nav>'
    for g,label in groups:
        h+=f'<nav class="ops-choices" data-group="{g}" aria-label="{label} examples"'+(' hidden' if g!='files' else '')+'>'
        for c in cards:
            if c['group']==g:h+=f'<button data-op-id="{c["id"]}" aria-controls="op-{c["id"]}" aria-pressed="{str(c["id"]==default).lower()}">{e(c["label"])}</button>'
        h+='</nav>'
    ex=[]
    for c in cards:
        id=c['id'];node='op-'+id
        h+=f'<article id="{node}" class="ops-panel" data-group="{c["group"]}"'+(' hidden' if id!=default else '')+'>'
        h+='<header class="ops-head"><h3>'+e(c['label'])+'</h3><p>'+e(c['intent'])+'</p><p class="ops-outcome">'+e(c['outcome'])+'</p></header><div class="ops-pair">'
        h+='<div class="ops-side"><h4>Library operation</h4><nav class="ops-input-nav"><button data-op-mode="readable" aria-pressed="true">Readable input</button><button data-op-mode="code" aria-pressed="false">Exact code</button></nav>'
        h+='<div data-op-pane="readable" class="ops-readable"><p>'+e(c['intent'])+'</p><p class="ops-caption">This code changes or reads the workbook. The separate checks on the right were also chosen and written by the agent.</p></div><div data-op-pane="code" hidden>'+code(c['operation'])+'</div>'
        native=c['libraryReturn']
        if native=={'javascriptType':'undefined'}:ret='No returned report (JavaScript undefined). For assignments, this is the result of the statement block. The workbook may still have changed.'
        elif isinstance(native,dict) and 'type' in native:ret='Returns '+('an ' if native['type'][0].lower() in 'aeiou' else 'a ')+native['type']+'. This label describes the live object; it is not a serialization of that object’s contents.'
        else:ret=json.dumps(native,ensure_ascii=False)
        h+='<div class="native-return"><b>What this operation returns</b><p>'+e(ret)+'</p></div></div>'
        h+='<div class="ops-side"><h4>Agent-written checks</h4><nav class="ops-input-nav"><button data-op-mode="readable" aria-pressed="true">Readable results</button><button data-op-mode="code" aria-pressed="false">Exact check code</button><button data-op-mode="tree" aria-pressed="false">Object tree</button></nav>'
        h+='<div data-op-pane="readable"><p class="ops-caption">These readings and comparisons are additional script work. They are not automatically returned by the edit or export.</p>'
        h+=observed(c)
        if c.get('before'):
            before_rows=scalar_rows(c['before'])[:12]
            h+='<details><summary>Selected readings before the operation</summary>'+(table(['Observation','Value'],before_rows) if before_rows else json_tree(c['before'],'Before'))+'</details>'
        h+='</div><div data-op-pane="code" hidden><b>Before</b>'+code(c['checkInput']['before'] or '// No before check requested')+'<b>After</b>'+code(c['checkInput']['after'] or '// No after check requested')+'<details><summary>Helpers written by the agent</summary>'+code(payload['helpers'])+'<p>sha() computes a file checksum; lastHandle is the object returned by the preceding operation. Full definitions are in the executed script.</p></details></div>'
        h+='<div data-op-pane="tree" class="ops-tree" hidden>'+json_tree({'before':c.get('before'),'after':c.get('after')},'Agent-selected observations')+'</div></div></div>'
        h+='<div class="ops-lesson">'+e(c['lesson'])+'</div>'
        if id in ('render-before','render-after'):
            image='operations-run/fee-lab-'+('before' if id=='render-before' else 'after')+'.png';raw=(Path(__file__).parent/image).read_bytes()
            h+='<figure><a href="'+image+'?raw"><img src="data:image/png;base64,'+base64.b64encode(raw).decode()+'" alt="Actual Fee lab render at '+('51' if id=='render-before' else '60')+' basis points" /></a><figcaption>Actual library render. Open the image to inspect it at full size.</figcaption></figure>'
        script='overlap-probe.mjs' if id=='overlap-table' else 'operations-run.mjs'
        h+='<div class="ops-raw"><details id="'+node+'-raw"><summary>Exact combined recording</summary>'+code(c)+'<p>Intent, outcome and lesson are authored descriptions added to the execution record. Live-object return types and undefined values are explicitly labelled.</p></details><p><a href="operations-run/'+id+'.json?raw">Original execution record</a> · <a href="'+script+'">Full executed script</a></p></div></article>'
        actions=[{'action':'click','selector':'[data-op-group="'+c['group']+'"]'},{'action':'click','selector':'[data-op-id="'+id+'"]'}]
        states=[{'id':'default','label':c['label']+' explanation and selected results','status':'supported','actions':actions}]
        for state,selector in [('input-code','.ops-side:first-child [data-op-mode="code"]'),('check-code','.ops-side:nth-child(2) [data-op-mode="code"]'),('tree','.ops-side:nth-child(2) [data-op-mode="tree"]')]:
            states.append({'id':state,'label':c['label']+' '+state,'status':'supported','actions':actions+[{'action':'click','selector':'#'+node+' '+selector}]})
        ex.append({'id':node,'kind':'operation-and-agent-checks','label':c['label'],'evidence':{'format':'json','data':c},'states':states})
    h+='</div>'
    return ComponentResult(h,dependencies=[{'id':'ops-css','kind':'css','content':CSS},{'id':'ops-js','kind':'js','content':JS}],exhibits=ex)

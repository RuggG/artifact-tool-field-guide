from html import escape as e
from pathlib import Path
import json, re, base64
from pagekit_components import ComponentResult

HERE=Path(__file__).parent
TERMS=json.loads((HERE/'evidence'/'terms.json').read_text())
def terms_for(c):
    keys=TERMS['cards'][c['id']]
    labels={'overview':'Workbook, tab, table and returned records','regions':'Range, region and preview','runner':'Script, library and tool call','ranges':'Cells, grids, values and formulas','trace':'Dependencies and missing formulas'}
    label=labels.get(c['id'],'Terms used in this example')
    return '<details class="xio-terms" id="xio-'+c['id']+'-terms"><summary>'+e(label)+'</summary><dl>'+''.join('<div><dt>'+e(TERMS['definitions'][k][0])+'</dt><dd>'+e(TERMS['definitions'][k][1])+'</dd></div>' for k in keys)+'</dl></details>'
def fmt(x):
    if x is None:return '—'
    if isinstance(x,bool):return str(x).lower()
    if isinstance(x,float):return f'{x:,.4f}'.rstrip('0').rstrip('.')
    return str(x)
def code(x):return '<pre tabindex="0"><code>'+e(x if isinstance(x,str) else json.dumps(x,indent=2,ensure_ascii=False))+'</code></pre>'
def table(headers,rows):
    return '<div class="xio-grid-scroll"><table'+(' class="xio-compact-table"' if len(headers)<=4 else '')+'><thead><tr>'+''.join('<th>'+e(h)+'</th>' for h in headers)+'</tr></thead><tbody>'+''.join('<tr>'+''.join('<td>'+str(v)+'</td>' for v in row)+'</tr>' for row in rows)+'</tbody></table></div>'
def col(n):
    s=''
    while n:n,m=divmod(n-1,26);s=chr(65+m)+s
    return s
def matrix(values,address='A1',formulas=None,labels=None):
    m=re.match(r'([A-Z]+)(\d+)',address);r0=int(m[2]);c0=0
    for c in m[1]:c0=c0*26+ord(c)-64
    width=max(map(len,values),default=0)
    headers=['Cell row']+[col(c0+j) for j in range(width)]
    if labels:headers=['Metric']+['K · FY2026','L · H1 2027','M · FY2027']
    rows=[]
    for i,row in enumerate(values):
        cells=[e(labels[i] if labels else str(r0+i))]
        for j,value in enumerate(row):
            cell='<span title="'+e(str(value),quote=True)+'">'+e(fmt(value))+'</span>'
            f=formulas[i][j] if formulas and i<len(formulas) and j<len(formulas[i]) else None
            if f:cell+='<details class="xio-cell-formula"><summary>Formula</summary><code>'+e(f)+'</code></details>'
            cells.append(cell)
        rows.append(cells)
    return table(headers,rows)

def json_tree(obj,key='Response',depth=0):
    if not isinstance(obj,(dict,list)):
        value=json.dumps(obj,ensure_ascii=False)
        return '<div class="j-leaf"><span class="j-key">'+e(str(key))+'</span><code>'+e(value)+'</code></div>'
    count=len(obj);kind='fields' if isinstance(obj,dict) else 'items'
    pairs=obj.items() if isinstance(obj,dict) else enumerate(obj)
    children=''.join(json_tree(v,k,depth+1) for k,v in pairs)
    return '<details class="j-node"'+(' open' if depth==0 else '')+'><summary><span class="j-key">'+e(str(key))+'</span><small>'+str(count)+' '+kind+'</small></summary><div>'+children+'</div></details>'

def dependency(n,depth=0):
    label='<b>'+e(n.get('cell',''))+'</b><span class="dep-value">'+e(fmt(n.get('value')))+'</span>'
    detail='<code>'+e(n['formula'])+'</code>' if n.get('formula') else '<small>Formula not returned · may be an input or an import limitation</small>'
    children=n.get('params',[])
    if not children:return '<div class="dep-leaf">'+label+detail+'</div>'
    return '<details class="dep-node"'+(' open' if depth==0 else '')+'><summary>'+label+'<small>'+str(len(children))+' dependencies</small></summary>'+detail+'<div>'+''.join(dependency(x,depth+1) for x in children)+'</div></details>'

LABELS={'baseFeeMargin':'Base fee margin · bp','activeFeeMargin':'Active fee margin · bp','q4AverageAum':'Q4 2027 average AUM · $bn','q4ManagementFees':'Q4 2027 management fees · $m','fy2027ManagementFees':'FY2027 management fees · $m','fy2027OperatingProfit':'FY2027 operating profit · $m'}
INPUT_GUIDES={
    'runner':('Run the saved experiment script with the bundled JavaScript runtime.',[
        ('Program','Node.js → run.mjs'),('Working folder','man-group-run, alongside the copied source.xlsx'),
        ('Initial wait','Up to 1,000 milliseconds before returning control if the process is still running'),
        ('Output allowance','Request up to 1,000 tokens of command output in this call')]),
    'overview':('Ask for a broad map of the imported workbook and previews of its contents.',[
        ('Scope','The whole workbook'),('Include','Workbook, sheet and table records'),
        ('Response cap requested','14,000 characters'),('Fallback preview requested','4 rows × 6 columns per table; 80 characters per cell. Used when table values would otherwise be too large.')]),
    'sheets':('List the sheets so the agent can choose where to look next.',[
        ('Scope','The whole workbook'),('Record type','Sheet'),('Fields requested','Internal ID and name')]),
    'literal':('Find cells containing the text “profit”.',[
        ('Scope','The whole workbook'),('Search text','profit'),('Matching mode','Literal text; regular expressions disabled'),
        ('Limits requested','30 matches; 12,000 response characters')]),
    'regex':('Find labels mentioning management fees or fee margins using a text pattern.',[
        ('Scope','The whole workbook'),('Pattern','management.*fee|fee.*margin'),
        ('Pattern meaning','“management” followed by “fee”, OR “fee” followed by “margin”. The .* permits any intervening characters.'),
        ('Matching mode','Regular expression; this is a text search, not a semantic search'),
        ('Limits requested','30 matches; 12,000 response characters')]),
    'regions':('Identify populated regions and request small previews within a selected part of Summary.',[
        ('Sheet','Summary'),('Search area','C7:M28'),('Record type','Region'),
        ('Preview caps requested','8 rows × 6 columns'),('Response cap requested','10,000 characters')]),
    'ranges':('Read the values and formula text from one exact rectangle.',[
        ('Sheet','Summary'),('Cells','K13:M20 · 8 rows × 3 columns'),
        ('Financial context','FY2026, H1 2027 and FY2027; management fees through operating profit'),
        ('Read','Two grids: values and formulas. Also return the sheet name and cell range.')]),
    'trace':('Ask the library to follow the calculation dependencies of one cell.',[
        ('Starting cell','Fee engine!R27'),('Financial context','Q3 2026 net management fees'),
        ('Request','Return the dependency tree from this cell')]),
    'render':('Render part of the workbook as a PNG and save the resulting image bytes.',[
        ('Sheet and area','Summary!C7:M28'),('Image settings','PNG, at scale 1.5'),
        ('Save as','summary-before.png, in the experiment folder'),
        ('Record','The saved filename, PNG media type and byte count'),
        ('Then view','A separate tools.view_image call opens that saved PNG for Codex to see')]),
    'value-edit':('Replace a single forecast assumption with the number 51.',[
        ('Target','Assumptions!W26'),('Financial meaning','Q4 2027 base management-fee margin, in basis points'),
        ('Write','51 as a cell value, replacing the existing cell content'),
        ('Agent-written check','I chose six cells and wrote readState() to read them before and after. The edit does not return them automatically.')]),
    'recalculate':('Ask the imported workbook to recalculate, then compare the same cells again.',[
        ('Scope','The workbook object in memory'),('Calculation call','wb.recalculate(), with no additional parameters'),
        ('Agent-written check','I chose six cells and wrote readState() to read them before and after. recalculate() does not return this comparison.'),
        ('Record','The calculation call’s return value as well as those readings; undefined is saved as the text “undefined”')]),
    'formula-edit':('Write the existing scenario-selection rule explicitly into the active fee-margin cell.',[
        ('Target','Assumptions!W25'),('Scenario selector','Assumptions!D4 chooses the first, second, third or fourth scenario'),
        ('Scenario inputs','W26, W27, W28 and W29, in that order'),
        ('Formula rule','Return the selected input. If that selected input is blank, return #N/A.'),
        ('Agent-written check','I chose six cells and wrote readState() to read them before and after. The formula edit does not choose or return them.')]),
    'fill':('Fill the active fee-margin formula to the right across six cells.',[
        ('Sheet','Assumptions'),('Range','R25:W25'),('Source','The leftmost cell, R25'),
        ('Copy behaviour','Extend the formula across the range, adjusting relative references while keeping $D$4 fixed'),
        ('Check','Read values and formula text for all six cells before and after')]),
    'copy':('Copy the original base fee-margin formula back into the edited assumption cell.',[
        ('Source','Assumptions!R26'),('Destination','Assumptions!W26'),('Copy mode','Formulas'),
        ('Source formula','=$D$50 — its absolute reference stays fixed when copied'),
        ('Agent-written check','I chose six cells and wrote readState() to read them before and after. These are additional reads.')]),
    'export':('Convert the in-memory workbook to an Excel file and save it locally.',[
        ('Source','The workbook after the preceding demonstration steps'),('Format','Excel .xlsx'),
        ('Save as','man-group-demo.xlsx, in the experiment folder'),
        ('Record','The save call’s return value, filename and file size in bytes')])
}

def readable_input(c):
    summary,items=INPUT_GUIDES[c['id']]
    h='<p class="xio-input-action">'+e(summary)+'</p><dl class="xio-input-fields">'
    h+=''.join('<div><dt>'+e(k)+'</dt><dd>'+e(v)+'</dd></div>' for k,v in items)+'</dl>'
    if c.get('helper'):
        cells=['Assumptions!W26 · base fee margin','Assumptions!W25 · active fee margin','Fee engine!W18 · Q4 average AUM','Fee engine!W27 · Q4 management fees','Summary!M13 · FY2027 management fees','Summary!M20 · FY2027 operating profit']
        h+='<p><b>Why six?</b> I selected these particular cells for this experiment. Six is not a tool setting or an automatically discovered dependency count.</p><details class="record-fold"><summary>The six cells I selected</summary><ul>'+''.join('<li>'+e(x)+'</li>' for x in cells)+'</ul><p>My helper records each cell’s address, value and formula. This sample does not verify every dependency.</p></details>'
    h+='<p class="xio-caption">Plain-English reading of the actual input. Requested limits describe the request, not a guarantee of what the library returns.</p>' if c['id'] in ('overview','literal','regex','regions') else '<p class="xio-caption">Plain-English reading of the actual input. Switch to Exact input for the executed code or tool arguments.</p>'
    return h

def readable(c):
    r=c['result'];id=c['id'];records=c.get('records',[])
    if id=='runner':
        initial=r['initialToolResult'];final=r['completionToolResult']
        return table(['Stage','What came back'],[['Initial call',e('Still running · session '+str(initial['session_id']))],['Later poll',e('Completed · exit code '+str(final['exit_code']))]])+'<details class="xio-raw"><summary>Read all printed output</summary>'+code(final['output'])+'</details>'
    if id=='overview':
        h='<div class="result-head"><b>6 sheets</b><span>5 records returned</span><strong class="limit">Response truncated</strong></div>'
        h+='<p class="xio-table-definition"><b>What does tables: 1 mean?</b> For Summary, this response packages A1:M57 as one table-shaped grid. That rectangle contains several financial sections. The file has <b>no formal Excel Tables</b>; the reported count describes the inspection output.</p>'
        for x in records:
            if x['kind']=='notice':h+='<p class="result-notice">'+e(x['message'])+'</p>';continue
            title=('Table response' if x['kind']=='table' else x['kind'].capitalize())+' · '+str(x.get('name',x.get('sheet',x.get('id',''))))
            if 'values' in x:
                h+='<details class="record-fold"><summary>'+e(title)+'<small>'+str(x['rows'])+' rows × '+str(x['cols'])+' columns</small></summary>'+matrix(x['values'],x['address'].split(':')[0])+'</details>'
            else:h+='<details class="record-fold"><summary>'+e(title)+'</summary>'+json_tree(x)+'</details>'
        return h
    if id=='sheets':return table(['Position','Sheet','Used range','Internal ID'],[[str(x['index']+1),e(x['name']),e(x['range']),'<code>'+e(x['id'])+'</code>'] for x in records])
    if id in ('literal','regex'):
        groups={};h=''
        for x in records:
            if x['kind']=='notice':h+='<p class="result-notice">'+e(x['message'])+'</p>'
            else:groups.setdefault(x.get('sheet',''),[]).append(x)
        h='<p class="result-count">'+str(sum(map(len,groups.values())))+' cell matches across '+str(len(groups))+' sheets</p>'+h
        for name,items in groups.items():
            h+='<details class="record-fold"'+(' open' if name=='Summary' else '')+'><summary>'+e(name)+'<small>'+str(len(items))+' matches</small></summary>'+table(['Cell','Matched value'],[[e(x['address']),e(str(x.get('value')))] for x in items])+'</details>'
        return h
    if id=='regions':
        return ''.join('<details class="record-fold"><summary>'+e(x['address'])+'<small>'+str(x['rows'])+' × '+str(x['cols'])+' · '+str(x['nonEmpty'])+' populated</small></summary><p class="xio-caption">Returned preview: '+e(x['previewAddress'])+'</p>'+matrix(x['preview'],x['previewAddress'].split(':')[0])+'</details>' for x in records)
    if id=='ranges':return matrix(r['values'],'K13',r['formulas'],['13 · Net management fees','14 · Performance fees','15 · Seed gains','16 · Other income','17 · Core net revenue','18 · Blank row','19 · Operating costs','20 · Operating profit'])+'<p class="xio-caption">USD millions. Click Formula beneath a value to inspect its calculation.</p>'
    if id=='trace':return '<div class="dep-tree">'+dependency(r)+'</div>'
    if id=='render':return table(['Saved image','Size'],[[e(r['saved']),f"{r['bytes']:,} bytes"]])+'<p>The actual pixels are shown below. Saving a PNG and viewing it are separate operations.</p>'
    if id in ('value-edit','recalculate','formula-edit','copy'):
        rows=[]
        for key,label in LABELS.items():
            a=r['before'][key];b=r['after'][key];changed=a['value']!=b['value']
            rows.append([e(label)+'<small class="cell-locator">'+e(a['sheet']+'!'+a['cell'])+'</small>',e(fmt(a['value'])),'<span class="'+('changed' if changed else '')+'">'+e(fmt(b['value']))+'</span>'])
        return table(['Readback','Before','After'],rows)+'<p class="xio-caption">Highlighted numbers changed. Expand the object tree for the exact formula returned at each stage.</p>'
    if id=='fill':
        rows=[]
        for j in range(6):
            a=r['before']['formulas'][0][j];b=r['after']['formulas'][0][j]
            rows.append([col(18+j)+'25',e(fmt(r['after']['values'][0][j])),e('Formula present' if a else 'No formula returned'),'<details class="xio-cell-formula"><summary>Read copied formula</summary><code>'+e(b)+'</code></details>'])
        return table(['Cell','Value after','Before','After'],rows)
    if id=='export':
        v=c['verification'];return table(['File result','Observed'],[['Saved file',e(r['saved'])],['Size',f"{r['bytes']:,} bytes"],['Formula/literal differences',str(len(v['formulaOrLiteralChanges']))],['Cached values outside tolerance',str(len(v['cachedValueChangesAboveTolerance']))],['Annotations',str(v['notesBefore'])+' before / '+str(v['notesAfter'])+' after'],['Original download', 'SHA-256 unchanged'],['Native Excel opening','Not tested in this run']])
    return json_tree(r)

CSS=(HERE/'man-group.css').read_text()+'''
.xio-head h3{font-size:22px!important}.xio-purpose{font-size:15px;line-height:1.65;margin:12px 0 5px}.xio-return-summary{font-size:13px;line-height:1.6;margin:9px 0;color:#405b70}.xio-finding{border-left:3px solid #6e92a9;background:#edf4f8;padding:12px 14px;margin:14px 0 0;font-size:14px;line-height:1.6}.xio-pair{grid-template-columns:minmax(0,.86fr) minmax(0,1.14fr)}.xio-side{padding:17px;overflow:hidden}.xio-view-nav{display:flex;gap:5px;margin-bottom:12px;flex-wrap:wrap}.xio-view-nav button{border:1px solid #cbd9e3;background:white;color:#375a75;padding:7px 9px;border-radius:4px;font-size:11px;cursor:pointer}.xio-view-nav button[aria-pressed=true]{background:#e7f0f6;border-color:#61869f;color:#173c57;font-weight:700}.xio-view[hidden]{display:none!important}.xio-view{font-size:12px;line-height:1.55;max-height:540px;overflow:auto;padding:1px}.xio-grid-scroll{max-height:420px}.xio-grid-scroll table{font-size:11px!important;width:100%}.xio-grid-scroll td,.xio-grid-scroll th{white-space:normal;vertical-align:top;min-width:60px;text-align:left;font-variant-numeric:tabular-nums;padding:8px}.xio-grid-scroll td:first-child{color:#284d67;min-width:95px}.xio-grid-scroll th{position:sticky;top:0;z-index:1}.xio-grid-scroll td code{white-space:pre-wrap;overflow-wrap:anywhere}.cell-locator{display:block;font-size:9px;color:#657e90}.changed{font-weight:700;color:#0b5f75;background:#dff2f4;padding:2px 3px}.xio-cell-formula{margin:4px 0!important;padding:0!important;border:none!important;background:transparent!important;box-shadow:none!important}.xio-cell-formula>summary{display:list-item!important;padding:2px 0!important;font-size:10px!important;color:#52778e;cursor:pointer}.xio-cell-formula code{font-size:10px!important;background:transparent!important}.record-fold,.j-node,.dep-node{border:0!important;box-shadow:none!important;background:transparent!important;margin:6px 0!important;padding:0!important}.record-fold>summary,.j-node>summary,.dep-node>summary{display:list-item!important;cursor:pointer;padding:10px 8px!important;background:#edf3f7;border-radius:4px;font-size:12px!important;font-weight:500!important;line-height:1.5;overflow-wrap:anywhere}.record-fold>summary>small,.j-node>summary>small,.dep-node>summary>small{color:#667d8e;font-size:10px;margin-left:10px;display:inline-block}.record-fold[open]>summary{margin-bottom:8px}.record-fold .xio-grid-scroll{max-height:330px}.result-head{display:flex;gap:7px;flex-wrap:wrap;align-items:center;padding:3px 0 14px}.result-head>*{padding:4px 6px;background:#edf3f7;border-radius:4px;font-size:11px}.result-head .limit{background:#fcf1d9;color:#715320;font-weight:500}.result-notice{background:#fbf5e6;border-left:3px solid #b6995a;padding:10px;font-size:12px;line-height:1.5}.result-count{margin:0 0 8px!important;font-size:12px;color:#46677f}.j-node>div{padding-left:13px;border-left:1px solid #d6e3ec;margin-left:5px}.j-leaf{padding:6px 0;display:grid;grid-template-columns:minmax(85px,.4fr) minmax(0,1fr);gap:8px;border-bottom:1px solid #edf2f6}.j-key{color:#315b77;overflow-wrap:anywhere}.j-leaf code{background:transparent;font:10px/1.6 ui-monospace,monospace;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;color:#395467}.dep-node>div{padding-left:15px;border-left:1px solid #d1e0e9;margin-left:4px}.dep-node>code,.dep-leaf>code{display:block;background:transparent;font:10px/1.6 ui-monospace,monospace;color:#466477;white-space:normal;overflow-wrap:anywhere;padding:6px}.dep-leaf{padding:7px 4px;border-bottom:1px solid #e2ebf1}.dep-leaf>b{color:#2c5571;overflow-wrap:anywhere}.dep-value{float:right;font-variant-numeric:tabular-nums;padding-left:8px;color:#143f59}.dep-leaf>small{display:block;color:#748896;font-size:9px}.xio-output pre{max-height:500px}.xio-image-block img{width:100%;height:auto}.xio-helper{margin-top:12px!important}.xio-caption{line-height:1.6!important}.xio-setup{font:11px/1.6 ui-monospace,monospace}
.xio .xio-grid-scroll .tablewrap{margin:0!important;max-width:100%;width:100%}.xio .xio-grid-scroll table.xio-compact-table{min-width:0!important;width:100%!important;table-layout:auto!important;margin:0!important}.xio .xio-grid-scroll table.xio-compact-table th,.xio .xio-grid-scroll table.xio-compact-table td{min-width:0!important;white-space:normal!important;overflow-wrap:anywhere;width:auto!important}.xio .xio-grid-scroll table.xio-compact-table td:first-child{min-width:75px!important}
.xio .xio-grid-scroll td>span{white-space:nowrap}
.xio-nav-label{font-size:11px;font-weight:700;color:#4d7089;margin:16px 0 7px}.xio-nav{margin-bottom:9px}.xio-nav button{display:block}
.xio-input-view[hidden]{display:none!important}.xio-input-view{font-size:12px;line-height:1.6;max-height:540px;overflow:auto;padding:1px}.xio-input-action{margin:0 0 12px!important;font-size:13px;color:#244b65}.xio-input-fields{margin:0}.xio-input-fields>div{padding:10px 0;border-bottom:1px solid #dde7ee}.xio-input-fields dt{font-size:10px;font-weight:700;color:#55748a;margin:0 0 3px}.xio-input-fields dd{margin:0;color:#243f53;overflow-wrap:anywhere}.xio-input-view ul{padding-left:20px}.xio-input-view li{margin:5px 0}
.xio-terms{margin:0!important;border:0!important;border-bottom:1px solid #d4e1ea!important;border-radius:0!important;padding:0!important;box-shadow:none!important;background:#f6f9fb!important}.xio-terms>summary{display:list-item!important;padding:12px 22px!important;font-size:12px!important;line-height:1.5!important;color:#355c78!important;cursor:pointer}.xio-terms>dl{padding:0 22px 16px;margin:0;display:grid;grid-template-columns:1fr 1fr;gap:14px 24px}.xio-terms dt{font-size:12px;color:#244b65;font-weight:700;margin-bottom:4px}.xio-terms dd{font-size:12px;line-height:1.6;margin:0;color:#405970;overflow-wrap:anywhere}.xio-table-definition{font-size:12px;line-height:1.6;border-left:3px solid #6e92a9;background:#edf4f8;padding:10px;margin:0 0 12px!important}
@media(max-width:650px){.xio-terms>dl{grid-template-columns:1fr}}
@media(max-width:850px){.xio-pair{grid-template-columns:1fr}.xio-side+div{border-left:0;border-top:1px solid var(--x-line)}}
@media print{.xio-view-nav{display:none}.xio-view[data-view=readable],.xio-input-view[data-input-pane=readable]{display:block!important;max-height:none}.xio-view[data-view=tree],.xio-view[data-view=raw],.xio-input-view[data-input-pane=exact]{display:none!important}.xio-grid-scroll{max-height:none}}
'''
JS='''(()=>{
const root=document.getElementById('excel-calls');if(!root)return;
function choose(id,hash){const panel=document.getElementById('xio-'+id);if(!panel)return;root.querySelectorAll('.xio-panel').forEach(p=>p.hidden=p!==panel);root.querySelectorAll('[data-xio-case]').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.xioCase===id)));if(hash)history.replaceState(null,'','#xio-'+id);}
root.querySelectorAll('[data-xio-case]').forEach(b=>b.addEventListener('click',()=>choose(b.dataset.xioCase,true)));
root.querySelectorAll('[data-result-view]').forEach(b=>b.addEventListener('click',()=>{const p=b.closest('.xio-panel'),v=b.dataset.resultView;p.querySelectorAll('.xio-view').forEach(x=>x.hidden=x.dataset.view!==v);p.querySelectorAll('[data-result-view]').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));}));
root.querySelectorAll('[data-input-view]').forEach(b=>b.addEventListener('click',()=>{const p=b.closest('.xio-input'),v=b.dataset.inputView;p.querySelectorAll('.xio-input-view').forEach(x=>x.hidden=x.dataset.inputPane!==v);p.querySelectorAll('[data-input-view]').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));}));
function fromHash(){if(location.hash.startsWith('#xio-')){const id=location.hash.slice(5);choose(id,false);const panel=document.getElementById('xio-'+id);if(panel){let n=panel.parentElement;while(n){if(n.tagName==='DETAILS')n.open=true;n=n.parentElement;}}}}
window.addEventListener('hashchange',fromHash);fromHash();
})();'''

def render(payload,*,body,context):
    if body.strip()=='flow':
        labels=[('1 · Codex chooses','What do I need to know?','The agent selects an operation and its settings.'),('2 · Command tool','Run the script','Node executes the JavaScript program.'),('3 · Artifact Tool','Operate on the workbook','The library reads, edits, traces, renders or calculates.'),('4 · Codex reads','Interpret the evidence','The script prints or saves results for the agent to inspect.')]
        html='<div id="excel-call-layers" class="xio-flow">'+''.join('<div><em>'+e(n)+'</em><b>'+e(a)+'</b><small>'+e(b)+'</small></div>' for n,a,b in labels)+'</div>'
        return ComponentResult(html,dependencies=[{'id':'man-group-call-style','kind':'css','content':CSS}],exhibits=[{'id':'excel-call-layers','kind':'diagram','label':'Request to returned evidence','evidence':{'format':'json','data':labels},'states':[{'id':'default','label':'Layer map','status':'supported','actions':[]}]}])
    cards=payload['cards'];h='<div class="xio" id="excel-calls">'
    groups=[('Execution',['runner']),('Discover · same inspect function',['sheets','overview','literal','regex','regions']),('Read specific evidence',['ranges','trace','render']),('Edit, calculate and save',['value-edit','formula-edit','fill','copy','recalculate','export'])]
    for label,ids in groups:
        h+='<div class="xio-nav-label">'+e(label)+'</div><nav class="xio-nav" aria-label="'+e(label)+'">'
        for cid in ids:
            c=next(c for c in cards if c['id']==cid);h+='<button type="button" data-xio-case="'+c['id']+'" aria-controls="xio-'+c['id']+'" aria-pressed="'+str(cid=='sheets').lower()+'">'+e(c['label'])+'</button>'
        h+='</nav>'
    h+='<p class="xio-caption">To compare settings of the same function, use the <a href="#inspection-settings">inspection settings comparison above</a>.</p>';ex=[]
    for i,c in enumerate(cards):
        id='xio-'+c['id'];h+='<article class="xio-panel" id="'+id+'"'+(' hidden' if c['id']!='sheets' else '')+' aria-label="'+e(c['title'])+'"><header class="xio-head"><div class="xio-badge">MAN GROUP · '+e(c['kind'])+'</div><h3>'+e(c['title'])+'</h3><p class="xio-purpose"><b>Why make this call? </b>'+e(c['purpose'])+'</p><p class="xio-return-summary"><b>What comes back? </b>'+e(c['returns'])+'</p><p class="xio-finding"><b>In this run: </b>'+e(c['finding'])+'</p></header>'
        h+=terms_for(c)
        h+='<div class="xio-pair"><div class="xio-side xio-input"><h4>Actual input</h4><nav class="xio-view-nav" aria-label="Input format">'
        for view,label in [('readable','Readable'),('exact','Exact input')]:h+='<button type="button" data-input-view="'+view+'" aria-controls="'+id+'-input-'+view+'" aria-pressed="'+('true' if view=='readable' else 'false')+'">'+label+'</button>'
        h+='</nav><div class="xio-input-view" data-input-pane="readable" id="'+id+'-input-readable">'+readable_input(c)+'</div><div class="xio-input-view" data-input-pane="exact" id="'+id+'-input-exact" hidden>'+code(c['input'])
        if c['id']!='runner':h+='<p class="xio-caption">Executed JavaScript inside the capture helper. “return” passes the result to the recorder that saves the JSON.</p>'
        if c.get('helper'):h+='<details class="xio-raw xio-helper"><summary>What does readState() read?</summary>'+code(c['helper'])+'</details>'
        if c.get('extraInput'):h+='<details class="xio-raw"><summary>Separate image-view tool input</summary>'+code(c['extraInput'])+'<p class="xio-caption">tools.view_image returns the image shown below to Codex.</p></details>'
        wrapped=c['id'] in ('ranges','render','value-edit','recalculate','formula-edit','fill','copy','export')
        h+='</div></div><div class="xio-side xio-output"><h4>'+('Script output and added checks' if wrapped else 'Library response' if c['id']!='runner' else 'Command-tool response')+'</h4>'
        if wrapped:
            h+='<p class="xio-table-definition"><b>This output is assembled by the script.</b> '+('I selected six cells, read them before and after, and built this table. The operation does not return those cells or certify the whole calculation chain.' if c.get('helper') else 'The script explicitly reads values, saves bytes, or performs comparisons. These fields are not an automatic validation report from the operation.')+'</p>'
        h+='<nav class="xio-view-nav" aria-label="Result format">'
        for view,label in [('readable','Readable'),('tree','Object tree'),('raw','Exact JSON')]:h+='<button type="button" data-result-view="'+view+'" aria-pressed="'+('true' if view=='readable' else 'false')+'">'+label+'</button>'
        h+='</nav><div class="xio-view" data-view="readable">'+readable(c)+'</div>'
        obj=c['result']
        if 'records' in c:obj={'records':c['records'],**{k:v for k,v in c['result'].items() if k!='ndjson'}}
        h+='<div class="xio-view" data-view="tree" hidden>'+('<p class="xio-caption">NDJSON decoded into records for browsing. The Exact JSON tab retains the original response string.</p>' if 'records' in c else '')+json_tree(obj)+'</div><div class="xio-view" data-view="raw" hidden>'+code(c['result'])+'</div></div></div>'
        if c.get('image'):
            raw=(HERE/c['image']).read_bytes();h+='<div class="xio-image-block"><h4>Actual rendered Man Group image</h4><a href="'+c['image']+'"><img class="xio-image" src="data:image/png;base64,'+base64.b64encode(raw).decode()+'" alt="Man Group Summary C7:M28: historical and forecast AUM, fees, profit and EPS" /></a><p class="xio-caption">Click for the original full-resolution PNG.</p></div>'
        h+='<div class="xio-notes">'+''.join('<p>'+e(n)+'</p>' for n in c['notes'])+'</div><div class="xio-links">'+''.join('<a href="'+e(x['href'])+'"'+(' download' if x['href'].endswith('.xlsx') else '')+'>'+e(x['label'])+'</a>' for x in c['links'])+'</div></article>'
        actions=[{'action':'click','selector':'[data-xio-case="'+c['id']+'"]'}]
        states=[{'id':'readable','label':c['label']+' readable result','status':'supported','actions':actions}]
        for view in ['tree','raw']:states.append({'id':view,'label':c['label']+' '+view,'status':'supported','actions':actions+[{'action':'click','selector':'#'+id+' [data-result-view="'+view+'"]'}]})
        states.append({'id':'exact-input','label':c['label']+' exact input','status':'supported','actions':actions+[{'action':'click','selector':'#'+id+' [data-input-view="exact"]'}]})
        states.append({'id':'terms','label':c['label']+' term definitions','status':'supported','actions':actions+[{'action':'click','selector':'#'+id+'-terms > summary'}]})
        ex.append({'id':id,'kind':'input-output','label':c['title'],'evidence':{'format':'json','data':c},'states':states})
    h+='</div>'
    return ComponentResult(h,dependencies=[{'id':'man-group-call-style','kind':'css','content':CSS},{'id':'man-group-call-switch','kind':'js','content':JS}],exhibits=ex)

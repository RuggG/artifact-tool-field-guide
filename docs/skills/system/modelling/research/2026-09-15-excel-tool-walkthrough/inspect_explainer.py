from html import escape as e
import json
from pagekit_components import ComponentResult

def pre(x):return '<pre tabindex="0"><code>'+e(x if isinstance(x,str) else json.dumps(x,indent=2,ensure_ascii=False))+'</code></pre>'
def grid(headers,rows):return '<div class="lab-scroll"><table><thead><tr>'+''.join('<th>'+e(x)+'</th>' for x in headers)+'</tr></thead><tbody>'+''.join('<tr>'+''.join('<td>'+e(str(x))+'</td>' for x in row)+'</tr>' for row in rows)+'</tbody></table></div>'
def tree(x,key='Response',depth=0):
 if not isinstance(x,(dict,list)):return '<div class="lab-leaf"><b>'+e(str(key))+'</b><code>'+e(json.dumps(x,ensure_ascii=False))+'</code></div>'
 pairs=x.items() if isinstance(x,dict) else enumerate(x)
 return '<details class="lab-tree"'+(' open' if depth==0 else '')+'><summary>'+e(str(key))+' <small>'+str(len(x))+(' fields' if isinstance(x,dict) else ' items')+'</small></summary><div>'+''.join(tree(v,k,depth+1) for k,v in pairs)+'</div></details>'
def col(n):
 s=''
 while n:n,m=divmod(n-1,26);s=chr(65+m)+s
 return s
def matrix(x,address):
 import re
 m=re.match(r'([A-Z]+)(\d+)',address);c0=0
 for c in m[1]:c0=c0*26+ord(c)-64
 r0=int(m[2]);width=max(map(len,x),default=0)
 def f(v):return '—' if v is None else (f'{v:,.4f}'.rstrip('0').rstrip('.') if isinstance(v,float) else str(v))
 return grid(['Row']+[col(c0+i) for i in range(width)],[[r0+i]+[f(v) for v in row] for i,row in enumerate(x)])

GROUPS=[('workbook','1 · Whole workbook','List the tabs, or also ask for their cell values.'),('table','2 · One tab’s values','Same table request and same 1 × 1 preview setting. Change only the character allowance.'),('regions','3 · Blocks + previews','Same tab and selected area. Change only the size of each block’s preview.'),('search','4 · Find cells','Same search function. Supply ordinary text or a regular expression.')]
HINTS={'map':'6 tab descriptions; no cell values','broad':'2 tab descriptions; 1 full value grid','all':'6 tab descriptions; 6 full value grids','tab-full':'All 109 × 23 cell values','tab-small':'A 1 × 1 preview: blank cell A1','blocks-small':'6 blocks; 2 × 3 cells from each','blocks-large':'Same 6 blocks; 4 × 6 cells from each','search-text':'28 matching cells','search-pattern':'30 matching cells + limit notice'}

def settings(p):
 a=p['input'];rows=[('What to return · kind',a['kind']),('Where to look',a.get('sheetId','All six sheets'))]
 if a.get('range'):rows.append(('Selected area · range',a['range']))
 rows.append(('Response allowance · maxChars',f"{a['maxChars']:,} characters"))
 if 'tableMaxRows' in a:
  rows.append(('Preview rows × columns',str(a['tableMaxRows'])+' × '+str(a['tableMaxCols'])))
  rows.append(('Text per preview cell',str(a['tableMaxCellChars'])+' characters'))
  rows.append(('When preview sizes apply','Fallback when table values are too large' if p['group'] in ('workbook','table') else 'To each region’s returned preview'))
 if 'searchTerm' in a:
  rows.extend([('Search term',a['searchTerm']),('Regular expression',str(a['options']['useRegex']).lower()),('Match limit',str(a['options']['maxResults']))])
 return '<dl class="lab-settings">'+''.join('<div><dt>'+e(k)+'</dt><dd>'+e(v)+'</dd></div>' for k,v in rows)+'</dl>'

def readable(p,sheets):
 records=p['records'];h=''
 if p['group']=='workbook':
  rows=[]
  for s in sheets:
   name=s['name'];desc=any(x['kind']=='sheet' and x.get('name')==name for x in records);data=next((x for x in records if x['kind']=='table' and x.get('sheet')==name),None)
   values=('Full '+str(len(data['values']))+' × '+str(len(data['values'][0]))+' grid') if data else ('Not requested' if p['id']=='map' else 'Not returned')
   rows.append([name,s['range'] if desc else 'Not returned',values])
  h+=grid(['Tab','Description: occupied range','Cell values returned'],rows)
  h+='<p class="lab-caption">All six tab names are shown for orientation. “Not returned” means absent from this captured response.</p>'
  if p['id']=='broad':h+='<p class="lab-warning"><b>Why it stops here:</b> Summary’s grid is 6,602 characters; Assumptions’ is 10,062. Each fits below 14,000 alone, but both together exceed the response allowance. The saved answer stops before Assumptions’ values and reports nine omitted records.</p>'
 elif p['group']=='table':
  x=next(x for x in records if x['kind']=='table')
  h+=grid(['What the size refers to','Actual size'],[['Sheet area described',x['address']+' · '+str(x['rows'])+' rows × '+str(x['cols'])+' columns'],['Cell values actually returned',str(len(x['values']))+' × '+str(len(x['values'][0]))+' grid'],['Values cut to a preview?',str(x.get('valuesTruncated',False)).lower()]])
  if p['id']=='tab-small':h+='<p class="lab-warning"><b>A1 is blank.</b> The returned values are [[null]]. The 109 × 23 dimensions still describe the whole area; they do not say that all those values were returned.</p>'
  else:h+='<p class="lab-caption">Every returned value matched a direct read of Fee engine!A1:W109. Formulas, formatting and comments are not included in this value grid.</p>'
 elif p['group']=='regions':
  h+=grid(['Block found','Full size','Preview returned'],[[x['address'],str(x['rows'])+' × '+str(x['cols']),x['previewAddress']] for x in records if x['kind']=='region'])
  h+='<p class="lab-caption">The same six blocks are returned in both settings. Preview size changes how much of each block’s content you see.</p>'
 elif p['group']=='search':
  matches=[x for x in records if x['kind']!='notice']
  h+='<p><b>'+str(len(matches))+' matching cells</b>. These are candidate locations, not an answer to a financial question.</p>'
  h+=grid(['Tab','Cell','Matched content'],[[x.get('sheet',''),x.get('address',''),x.get('value','')] for x in matches])
 for x in records:
  if x['kind']=='notice':h+='<p class="lab-warning">'+e(x['message'])+'</p>'
  if x['kind']=='table':
   values=x['values'];addr=x.get('valuesPreviewAddress',x['address']).split(':')[0]
   h+='<details class="lab-data"><summary>Read returned values · '+e(x['sheet'])+' · '+str(len(values))+' × '+str(len(values[0]))+'</summary>'+matrix(values,addr)+'</details>'
  if x['kind']=='region':h+='<details class="lab-data"><summary>Read preview · '+e(x['previewAddress'])+'</summary>'+matrix(x['preview'],x['previewAddress'].split(':')[0])+'</details>'
 return h

CSS='''
.lab{border:1px solid #cfdee8;border-radius:8px;margin:20px 0 28px;background:white;overflow:hidden;color:#183b52}.lab-top{padding:20px;background:#edf4f8}.lab-top p{font-size:14px;line-height:1.6;margin:5px 0}.lab-function{font:700 18px ui-monospace,monospace;color:#254f6d}.lab-groups{display:grid;grid-template-columns:repeat(4,1fr);gap:7px;padding:16px 18px 0}.lab button{font-family:inherit;cursor:pointer;border:1px solid #bccfdd;background:#f6f9fb;border-radius:5px;color:#244d69;padding:9px;font-size:12px}.lab button[aria-pressed=true]{background:#214c6a;color:white;border-color:#214c6a}.lab button:focus-visible,.lab summary:focus-visible{outline:3px solid #75b9df;outline-offset:2px}.lab-presets{padding:14px 18px 18px;display:flex;flex-wrap:wrap;gap:7px;align-items:center}.lab-group-note{margin:12px 18px 0!important;font-size:13px;color:#476679}.lab-group[hidden],.lab-panel[hidden],.lab-pane[hidden]{display:none!important}.lab-panel>header{padding:18px 20px;border-top:1px solid #d7e4ec;background:#f6fafc}.lab-panel>header h3{font-size:20px!important;line-height:1.35!important;margin:0 0 9px!important}.lab-panel>header p{font-size:13px;margin:7px 0;line-height:1.6}.lab-stats{display:flex;gap:8px;flex-wrap:wrap;font-size:11px;margin-top:10px}.lab-stats span{padding:5px 7px;background:#e6eff5;border-radius:4px}.lab-io{display:grid;grid-template-columns:minmax(0,.8fr) minmax(0,1.2fr)}.lab-side{min-width:0;padding:17px}.lab-side+div{border-left:1px solid #d7e4ec}.lab-side h4{font-size:11px!important;text-transform:uppercase;letter-spacing:.06em;margin:0 0 9px!important}.lab-formats{display:flex;gap:5px;flex-wrap:wrap;margin-bottom:12px}.lab-formats button{font-size:11px;padding:6px 8px}.lab-settings{margin:0}.lab-settings>div{border-bottom:1px solid #e0e9f0;padding:9px 0}.lab-settings dt{font-size:10px;color:#5a7488;font-weight:700;margin-bottom:3px}.lab-settings dd{margin:0;font-size:12px;overflow-wrap:anywhere;line-height:1.6}.lab-pane{font-size:12px;line-height:1.6;max-height:650px;overflow:auto;padding:1px}.lab-caption{font-size:11px!important;color:#607788;line-height:1.6!important}.lab-warning{padding:10px;border-left:3px solid #a38a50;background:#faf5e9;font-size:12px;line-height:1.6}.lab pre{font:11px/1.6 ui-monospace,monospace!important;white-space:pre-wrap!important;overflow-wrap:anywhere;background:#edf3f7!important;padding:12px!important;border:0;margin:0;max-height:600px;overflow:auto}.lab pre code{font:inherit!important;white-space:inherit;background:transparent!important;color:#284c66;padding:0!important}.lab-scroll{overflow:auto;max-height:350px;max-width:100%}.lab-scroll .tablewrap{margin:0!important;width:100%!important}.lab-scroll table{border-collapse:collapse;width:100%!important;min-width:0!important;font-size:11px!important;margin:0!important;table-layout:auto!important}.lab-scroll td,.lab-scroll th{padding:7px;border:1px solid #dce7ef;text-align:left!important;line-height:1.5;min-width:0!important;white-space:normal!important;overflow-wrap:anywhere}.lab-scroll th{background:#edf4f8;color:#395c74}.lab-data,.lab-tree{border:0!important;box-shadow:none!important;margin:8px 0!important;padding:0!important;background:transparent!important}.lab-data>summary,.lab-tree>summary{display:list-item!important;cursor:pointer;font-size:12px!important;line-height:1.5!important;padding:10px!important;background:#edf3f7}.lab-data .lab-scroll table{min-width:max-content!important}.lab-data .lab-scroll td{white-space:nowrap!important}.lab-tree>div{border-left:1px solid #cddde8;margin-left:6px;padding-left:12px}.lab-tree small{color:#6a8293}.lab-leaf{display:grid;grid-template-columns:minmax(65px,.4fr) minmax(0,1fr);gap:6px;padding:5px 0;border-bottom:1px solid #e6eef3}.lab-leaf b{font-weight:400;overflow-wrap:anywhere}.lab-leaf code{font:10px/1.5 ui-monospace,monospace;white-space:pre-wrap;overflow-wrap:anywhere;background:transparent}.lab-links{padding:12px 20px;border-top:1px solid #d7e4ec;font-size:11px;display:flex;gap:15px;flex-wrap:wrap}.api-map{font-size:13px;line-height:1.6}.api-map a{white-space:normal}.lab-end-note{font-size:12px;color:#557184;padding:0 20px 18px}.lab-static{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:15px 0}.lab-static>div{padding:14px;background:#edf4f8;border:1px solid #d0e0eb}.lab-static p{font-size:13px;line-height:1.6;margin:5px 0}.lab-static b{color:#254e6a}.lab-token-note{font-size:11px;color:#607788;margin-top:10px}
.lab-presets{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr))}.lab-presets button{text-align:left;line-height:1.4;align-self:stretch}.lab-presets button span{display:block;font-weight:700}.lab-presets button small{display:block;margin-top:5px;font-size:10px;line-height:1.5;opacity:.85}
@media(max-width:800px){.lab-io{grid-template-columns:1fr}.lab-side+div{border-left:0;border-top:1px solid #d7e4ec}.lab-groups{grid-template-columns:1fr 1fr}.lab-static{grid-template-columns:1fr}}
'''
JS='''(()=>{const root=document.getElementById('inspect-lab');if(!root)return;
function preset(id){root.querySelectorAll('.lab-panel').forEach(x=>x.hidden=x.dataset.preset!==id);root.querySelectorAll('[data-lab-preset]').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.labPreset===id)));}
function group(id){root.querySelectorAll('[data-lab-group]').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.labGroup===id)));root.querySelectorAll('.lab-group').forEach(x=>x.hidden=x.dataset.group!==id);const first=root.querySelector('.lab-group[data-group="'+id+'"] [data-lab-preset]');if(first)preset(first.dataset.labPreset);}
root.querySelectorAll('[data-lab-group]').forEach(b=>b.addEventListener('click',()=>group(b.dataset.labGroup)));
root.querySelectorAll('[data-lab-preset]').forEach(b=>b.addEventListener('click',()=>preset(b.dataset.labPreset)));
root.querySelectorAll('[data-lab-format]').forEach(b=>b.addEventListener('click',()=>{const side=b.closest('.lab-side');side.querySelectorAll('[data-lab-format]').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));side.querySelectorAll('.lab-pane').forEach(x=>x.hidden=x.dataset.format!==b.dataset.labFormat);}));
})();'''

def render(payload,*,body,context):
 presets=payload['presets'];sheets=payload['sheets']
 h='<div class="lab" id="inspect-lab"><div class="lab-top"><div class="lab-function">wb.inspect(settings)</div><p><b>Every option below calls this same function.</b> Change what it returns, where it looks and how much output it allows.</p><p class="lab-caption">These are recorded calls on the Man Group file. The controls switch between saved results; they do not rerun the model.</p></div><nav class="lab-groups" aria-label="Inspection question">'
 for id,label,note in GROUPS:h+='<button type="button" data-lab-group="'+id+'" aria-pressed="'+str(id=='workbook').lower()+'">'+label+'</button>'
 h+='</nav>'
 for id,label,note in GROUPS:
  h+='<div class="lab-group" data-group="'+id+'"'+(' hidden' if id!='workbook' else '')+'><p class="lab-group-note">'+e(note)+'</p><nav class="lab-presets" aria-label="Recorded settings for '+e(label)+'">'
  for p in presets:
   if p['group']==id:h+='<button type="button" data-lab-preset="'+p['id']+'" aria-controls="lab-'+p['id']+'" aria-pressed="'+str(p['id']=='map').lower()+'"><span>'+e(p['label'])+'</span><small>'+e(HINTS[p['id']])+'</small></button>'
  h+='</nav></div>'
 ex=[]
 for p in presets:
  id=p['id'];h+='<article id="lab-'+id+'" class="lab-panel" data-preset="'+id+'"'+(' hidden' if id!='map' else '')+'><header><h3>'+e(p['outcome'])+'</h3><p>'+e(p['change'])+'</p><div class="lab-stats"><span>'+f"{p['characters']:,}"+' characters</span><span>≈ '+f"{p['tokens']:,}"+' tokens</span><span>'+str(p['result']['recordCount'])+' returned records</span><span>'+('Whole response cut short' if p['result']['truncated'] else 'No whole-response cutoff')+'</span></div></header><div class="lab-io">'
  for side in ['input','output']:
   formats=[('readable','Readable'),('exact','Exact input')] if side=='input' else [('readable','Readable'),('tree','Object tree'),('exact','Exact JSON')]
   h+='<div class="lab-side" data-side="'+side+'"><h4>'+('Settings passed to inspect' if side=='input' else 'What this call returned')+'</h4><nav class="lab-formats" aria-label="'+side+' format">'
   for v,label in formats:h+='<button type="button" data-lab-format="'+v+'" aria-pressed="'+str(v=='readable').lower()+'">'+label+'</button>'
   h+='</nav>'
   for v,label in formats:
    contents=settings(p) if side=='input' and v=='readable' else ('<p class="lab-caption">The actual settings object passed to wb.inspect.</p>'+pre(p['input']) if side=='input' else (readable(p,sheets) if v=='readable' else (tree({'records':p['records'],**{k:x for k,x in p['result'].items() if k!='ndjson'}}) if v=='tree' else pre(p['result']))))
    h+='<div class="lab-pane" data-format="'+v+'"'+(' hidden' if v!='readable' else '')+'>'+contents+'</div>'
   h+='</div>'
  h+='</div><div class="lab-links"><a href="'+p['file']+'">Full saved call</a><a href="'+('man-group-run/run.mjs' if p['file'].startswith('man-group-run') else 'overview-probe.mjs')+'">Executed script</a></div></article>'
  actions=[{'action':'click','selector':'[data-lab-group="'+p['group']+'"]'},{'action':'click','selector':'[data-lab-preset="'+id+'"]'}]
  states=[{'id':'readable','label':p['label'],'status':'supported','actions':actions}]
  for side,v in [('input','exact'),('output','tree'),('output','exact')]:states.append({'id':side+'-'+v,'label':p['label']+' '+side+' '+v,'status':'supported','actions':actions+[{'action':'click','selector':'#lab-'+id+' [data-side="'+side+'"] [data-lab-format="'+v+'"]'}]})
  ex.append({'id':'lab-'+id,'kind':'input-output','label':p['outcome'],'evidence':{'format':'json','data':p},'states':states})
 h+='<p class="lab-end-note">Token estimates count the returned NDJSON using o200k_base, excluding input, other metadata and tool-message overhead. Different captures can have slightly different counts because internal IDs change. This is not an exact GPT-6 billing count.</p></div>'
 return ComponentResult(h,dependencies=[{'id':'inspect-lab-style','kind':'css','content':CSS},{'id':'inspect-lab-controls','kind':'js','content':JS}],exhibits=ex)

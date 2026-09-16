"""Capability-first reference; recorded examples are supporting, lazily displayed evidence."""
from html import escape as e
from pathlib import Path
import json,sys
sys.path.insert(0,str(Path(__file__).parent))
from pagekit_components import ComponentResult
from man_group_walkthrough import readable as old_readable, readable_input, code
from operations_explorer import observed
from inspect_explainer import readable as inspect_readable, settings
H=Path(__file__).parent

def load(name):return json.loads((H/'evidence'/name).read_text())

def examples():
 out={};opdata=load('operations-cards.json')
 for c in opdata['cards']:
  native=c['libraryReturn']
  if native=={'javascriptType':'undefined'}:
   ret='No returned report (JavaScript undefined). This is the return of the recorded code block; the workbook can still have changed.'
  elif 'type' in native:
   ret='Returned: '+native['type']+'. This is our label for the live object, not a complete serialization of its contents.'
  else:ret=json.dumps(native,ensure_ascii=False)
  id='op:'+c['id'];image=''
  if c['id'] in ('render-before','render-after'):
   image='operations-run/fee-lab-'+('before' if c['id']=='render-before' else 'after')+'.png'
  out[id]=dict(id=id,label=c['label'],intent=c['intent'],finding=c['outcome'],
   inputReadable='<p>'+e(c['intent'])+'</p>',input=c['operation'],outputLabel='Operation return',
   outputReadable='<p>'+e(ret)+'</p>',output=native,
   checkReadable=observed(c),checkCode='// Before\n'+(c['checkInput']['before'] or '// No check requested')+'\n\n// After\n'+(c['checkInput']['after'] or '// No check requested'),
   checks={k:c[k] for k in ('before','after') if k in c},checkHelpers=opdata['helpers'],
   note=c['lesson'],image=image,
   links=[['Execution record','operations-run/'+c['id']+'.json?raw'],['Full script','overlap-probe.mjs' if c['id']=='overlap-table' else 'operations-run.mjs']])
 old=load('man-group-cards.json')
 for c in old['cards']:
  wrapped=c['id'] in ('value-edit','formula-edit','recalculate','fill','copy','export','render','ranges')
  label='Script output — includes added reads/checks' if wrapped else ('Command-tool result' if c['id']=='runner' else 'Library response')
  checknote='The exact input includes the agent’s wrapper. The displayed report is assembled by that script; it is not automatically returned by the edit, calculation or export.' if wrapped else ''
  if c['id']=='ranges':
   label='Cell getters, packaged by the script'
   checknote='The values and formulas come from native getters. The script bundles them together; the readable grid adds metric and period captions for orientation.'
  if c['id']=='render':
   label='Script output after saving the image'
   checknote='Rendering returns an image blob. The script saves its bytes and assembles the filename/size report shown here.'
  if c['id']=='export':
   label='Script output and independent file checks'
   checknote='The save returns no report. The script records the filename and size; separate file-reading checks supply the comparison in the readable display.'
  out['old:'+c['id']]=dict(id='old:'+c['id'],label=c['label']+' · original model',intent=c['purpose'],finding=c['finding'],inputReadable=readable_input(c),input=c['input'],outputLabel=label,outputReadable=old_readable(c),output=c['result'],checkReadable='',checks=None,checkCode='',checkHelpers=c.get('helper',''),note=checknote+' '+(' '.join(c['notes'])),image=('man-group-run/summary-before.png' if c['id']=='render' else ''),links=[[x['label'],x['href']] for x in c['links']])
 lab=load('inspect-lab.json')
 for c in lab['presets']:
  out['lab:'+c['id']]=dict(id='lab:'+c['id'],label=c['label'],intent=c['change'],finding=c['outcome'],inputReadable=settings(c),input='await wb.inspect('+json.dumps(c['input'],indent=2)+');',outputLabel='Library response',outputReadable=inspect_readable(c,lab['sheets']),output=c['result'],checks=None,checkCode='',checkReadable='',checkHelpers='',note='Size of the returned NDJSON text: '+f"{c['characters']:,}"+' characters; approximately '+f"{c['tokens']:,}"+' tokens using o200k_base. This is not a billing count or the size of the whole outer tool response. NDJSON means one JSON record per line.',image='',links=[['Execution record',c['file']+'?raw']])
 return out

def render(payload,*,body,context):
 cards=examples();used={x for op in payload['operations'] for x in op['examples']}
 for id in used:assert id in cards,id
 aliases={}
 for op in payload['operations']:
  for id in op['examples']:
   prefix,key=id.split(':');aliases.setdefault(('xio-' if prefix=='old' else 'op-' if prefix=='op' else 'lab-')+key,[op['id'],id])
 aliases.update({'xio-overview':['grids','lab:broad'],'xio-literal':['search','lab:search-text'],'xio-regex':['search','lab:search-pattern'],'inspection-settings':['map',None],'calls':['map',None],'operations':['map',None],'layers':['map',None]})
 h='<div id="operations-reference" class="oref"><div class="oref-map" aria-label="Operations catalogue">'
 for id,label in payload['groups']:
  h+='<section class="oref-family"><h3>'+e(label)+'</h3><div>'
  for op in payload['operations']:
   if op['group']==id:h+='<button type="button" data-operation="'+op['id']+'" aria-pressed="'+str(op['id']=='map').lower()+'">'+e(op['label'])+'</button>'
  h+='</div></section>'
 h+='</div>'+''.join('<span id="ref-'+op['id']+'" class="oref-anchor"></span>'+''.join('<span id="ref-'+op['id']+'~'+x.replace(':','-')+'" class="oref-anchor"></span>' for x in op['examples']) for op in payload['operations'])+'<article id="operation-detail" class="oref-detail" aria-label="Selected operation"></article>'
 data=dict(operations=payload['operations'],examples={k:v for k,v in cards.items() if k in used},aliases=aliases)
 h+='<script id="operations-reference-data" type="application/json">'+json.dumps(data,ensure_ascii=False).replace('<','\\u003c')+'</script></div>'
 states=[dict(id='map',label='Catalogue and default operation',status='supported',actions=[])]
 for op in payload['operations']:
  states.append(dict(id='reference-'+op['id'],label=op['label'],status='supported',actions=[dict(action='click',selector='[data-operation="'+op['id']+'"]')],selector='#operation-detail'))
 # All recorded examples use one UI. Exercise each example once, then representative input/output/check modes.
 for op in payload['operations']:
  for example in op['examples']:
   actions=[dict(action='click',selector='[data-operation="'+op['id']+'"]'),dict(action='click',selector='#recorded-examples > summary')]
   if len(op['examples'])>1:actions.append(dict(action='select',selector='#example-choice',value=example))
   states.append(dict(id='example-'+op['id']+'-'+example.replace(':','-'),label=op['label']+' / '+cards[example]['label'],status='supported',actions=actions,selector='#example-detail'))
 for name,button in [('exact-input','[data-side="input"] [data-mode="code"]'),('output-tree','[data-side="output"] [data-mode="tree"]'),('exact-output','[data-side="output"] [data-mode="code"]')]:
  states.append(dict(id=name,label=name,status='supported',actions=[dict(action='click',selector='[data-operation="ranges"]'),dict(action='click',selector='#recorded-examples > summary'),dict(action='click',selector=button)],selector='#example-detail'))
 for name,button in [('checks-readable',None),('checks-code','[data-side="checks"] [data-mode="code"]'),('checks-tree','[data-side="checks"] [data-mode="tree"]')]:
  actions=[dict(action='click',selector='[data-operation="values"]'),dict(action='click',selector='#recorded-examples > summary'),dict(action='click',selector='#agent-checks > summary')]
  if button:actions.append(dict(action='click',selector=button))
  states.append(dict(id=name,label=name,status='supported',actions=actions,selector='#example-detail'))
 return ComponentResult(h,dependencies=[dict(id='operations-reference-css',kind='css',content=(H/'operations-reference.css').read_text()),dict(id='operations-reference-js',kind='js',content=(H/'operations-reference.js').read_text())],exhibits=[dict(id='operations-reference',kind='capability-reference',label='Excel file operations',evidence=dict(format='json',data=payload),states=states)])

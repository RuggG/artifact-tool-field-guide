"""Page-local, evidence-declaring adapter for the interactive Artifact Tool guide."""
import json
from pathlib import Path
from pagekit_components import ComponentResult

HERE=Path(__file__).resolve().parent
MODES={'workbook-journey':'journey','action-explorer':'families','catalogue-browser':'catalogue','boundary-lab':'boundaries'}

def state(id,label,*actions):return {'id':id,'label':label,'status':'supported','actions':list(actions)}
def click(selector):return {'action':'click','selector':selector}
def render(payload,*,body,context):
    cid=context.id;mode=MODES[cid]
    html=f'<div class="af-guide" id="af-{mode}" data-af-mode="{mode}"><p style="padding:24px">Loading the saved examples. This guide contains captured data and does not call the workbook engine.</p><noscript>Enable JavaScript for the interactive views. The surrounding explanation and linked Markdown reference remain readable.</noscript></div>'
    states=[state('default','Default readable view')]
    if mode=='journey':
        states += [state('raw','Raw captured step',click('[data-journey-view="raw"]')),state('tree','Workbook object tree',click('[data-journey-view="tree"]')),state('quantity','Ten-unit captured case',{'action':'select','selector':'#af-quantity','value':'10'}),state('reopened','Verified exported/reopened result',click('[data-step="7"]'))]
        evidence=payload['walkthrough']
    elif mode=='families':
        states += [state('chart','Chart family',click('[data-family="charts"]')),state('settings','Family settings',click('#af-families .af-settings>summary')),state('tree','Input/output trees',click('[data-family-view="tree"]')),state('raw','Raw input and return value',click('[data-family-view="raw"]')),state('session','Persistent session',click('[data-family="changes"]'),{'action':'select','selector':'#af-example','value':'session-managed-outputs'})]
        evidence={k:payload[k] for k in ['families','recipes','operations']}
    elif mode=='catalogue':
        states += [state('offset','OFFSET result and independent check',{'action':'fill','selector':'#af-catalogue-search','value':'OFFSET'},click('[data-record="fx.OFFSET"]')),state('charts','Native chart comparison',click('[data-catalogue="charts"]')),state('shapes','Shape catalogue',click('[data-catalogue="shapes"]')),state('support','Supplied scripts and guidance',click('[data-catalogue="support"]')),state('empty','No search results',{'action':'fill','selector':'#af-catalogue-search','value':'zzz-no-such-function'})]
        evidence={k:payload[k] for k in ['apis','formulas','charts','shapes','support']}
    else:
        states += [state('fresh','Fresh structured total',click('[data-checkpoint="2"]')),state('validation','Validation boundary',click('[data-boundary="validation"]')),state('reexport','Slicer re-export boundary',click('[data-boundary="reexport"]'))]
        evidence=payload['limits']
    safe=json.dumps(payload,ensure_ascii=False,separators=(',',':')).replace('<','\\u003c').replace('\u2028','\\u2028').replace('\u2029','\\u2029')
    return ComponentResult(html,dependencies=[{'id':'artifact-guide-data','kind':'js','content':'window.ArtifactFieldGuideData='+safe+';'}, {'id':'artifact-guide-css','kind':'css','path':str(HERE/'explorer.css')},{'id':'artifact-guide-js','kind':'js','path':str(HERE/'explorer.js')}],exhibits=[{'id':'af-'+mode,'kind':'interactive-reference','label':{'journey':'Recorded workbook lifecycle','families':'Action families and input/output explorer','catalogue':'Complete searchable capability catalogues','boundaries':'Observed limitations and working patterns'}[mode],'evidence':{'format':'json','data':evidence,'source':'data/field-guide.json'},'states':states}])

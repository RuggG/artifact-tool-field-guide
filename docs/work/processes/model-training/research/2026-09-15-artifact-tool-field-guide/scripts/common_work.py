"""Page-local adapter for the practical selection and its retained recordings."""
import json
from pathlib import Path
from pagekit_components import ComponentResult

HERE = Path(__file__).resolve().parent


def render(payload, *, body, context):
    inspection = json.loads((HERE.parent / 'data/inspection.json').read_text())
    reads = json.loads((HERE.parent / 'data/reads.json').read_text())
    def state(id, label, *actions):
        return dict(id=id,label=label,status='supported',actions=list(actions))
    def click(selector):
        return dict(action='click',selector=selector)
    states = [state('default','Sheet inventory from the Man Group workbook')]
    for id,label in [('values','An edit and separate checks'),('trace','Formula dependency tree'),('table','A native Excel Table'),('render','Recorded image output')]:
        states.append(state(id,label,dict(action='select',selector='#af-common-jump',value=id)))
    choose = lambda selector, value: dict(action='select',selector=selector,value=value)
    states += [state('worked-'+action, label, choose('#af-common-jump', action), choose('#af-common-example', example)) for action,example,label in [('map','old:sheets','Sheet inventory field selection'),('regions','old:regions','Summary financial blocks'),('search','lab:search-pattern','Pattern search across all tabs'),('ranges','old:ranges','Summary values and formulas')]]
    states += [
        state('grid','Full Fee engine grid',choose('#af-common-jump','grids')),
        state('preview','One-cell preview with outer truncated false',choose('#af-common-jump','grids'),choose('#ix-budget','1000')),
        state('range','A complete narrow range',choose('#af-common-jump','grids'),choose('#ix-scope','range')),
        state('regions','Detected blocks and preview footprints',choose('#af-common-jump','regions')),
        state('search','Match limits and notices',choose('#af-common-jump','search'),choose('#ix-preset','match-limit')),
        state('formulas','Formula record limits',choose('#ix-mode','formulas')),
        state('raw','Complete inspection return',click('[data-ix-view="raw"]')),
        state('tree','Expandable inspection response',click('[data-ix-view="tree"]')),
    ]
    for entry in states:
        if entry['id'] in {'grid','preview','range','regions','search','formulas','raw','tree'}:
            entry['selector'] = '#cw-inspection .ix-response'
    states.append(dict(state('coverage','All six sheet coverage maps'), selector='#cw-inspection .ix-coverage'))
    states.append(dict(state('coverage-one','Focus one coverage map without changing the request',choose('#ix-sheet','Fee engine')), selector='#cw-inspection .ix-coverage'))
    for method in ['values','formulas','both','paired','details','inspect','search']:
        states.append(dict(state('read-'+method,'Read cells: '+method,choose('#af-common-jump','ranges'),choose('#rx-method',method)),selector='#cw-reads .rx-response'))
    for scope in ['cell','row','column','index','separate','sheet','workbook']:
        states.append(dict(state('read-scope-'+scope,'Read cells: '+scope,choose('#af-common-jump','ranges'),choose('#rx-method','paired'),choose('#rx-scope',scope)),selector='#cw-reads .rx-response'))
    for view in ['raw','tree']:
        states.append(dict(state('read-'+view,'Reading response: '+view,choose('#af-common-jump','ranges'),click('[data-rx-view="'+view+'"]')),selector='#cw-reads .rx-response'))
    reads_safe=json.dumps(reads,ensure_ascii=False,separators=(',',':')).replace('<','\\u003c').replace('\u2028','\\u2028').replace('\u2029','\\u2029')
    safe=json.dumps(payload,ensure_ascii=False,separators=(',',':')).replace('<','\\u003c').replace('\u2028','\\u2028').replace('\u2029','\\u2029')
    inspect_safe=json.dumps(inspection,ensure_ascii=False,separators=(',',':')).replace('<','\\u003c').replace('\u2028','\\u2028').replace('\u2029','\\u2029')
    html='<div class="af-guide af-common" id="af-common"><p>Loading the examples.</p><noscript>Enable JavaScript to explore the examples; the surrounding explanation and source links remain readable.</noscript></div>'
    return ComponentResult(html,dependencies=[dict(id='artifact-common-data',kind='js',content='window.ArtifactCommonWorkData='+safe+';'),dict(id='artifact-inspection-data',kind='js',content='window.ArtifactInspectionData='+inspect_safe+';'),dict(id='artifact-reads-data',kind='js',content='window.ArtifactReadsData='+reads_safe+';'),dict(id='artifact-reads-css',kind='css',path=str(HERE/'reads.css')),dict(id='artifact-reads-js',kind='js',path=str(HERE/'reads.js')),dict(id='artifact-guide-css',kind='css',path=str(HERE/'explorer.css')),dict(id='artifact-common-css',kind='css',path=str(HERE/'common-work.css')),dict(id='artifact-inspection-css',kind='css',path=str(HERE/'inspection.css')),dict(id='artifact-inspection-js',kind='js',path=str(HERE/'inspection.js')),dict(id='artifact-common-js',kind='js',path=str(HERE/'common-work.js'))],exhibits=[dict(id='af-common',kind='interactive-reference',label='23 common actions, 47 inspection and 38 reading comparisons',evidence=dict(format='json',data=dict(common=payload,inspection=inspection,reads=reads),source='data/common-work.json; data/inspection.json; data/reads.json'),states=states)])

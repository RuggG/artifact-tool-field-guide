#!/usr/bin/env python3
"""Check the public file closure, captured links, and browser interactions."""
import argparse
import json
import re
from pathlib import Path
from urllib.parse import urljoin,urlsplit,unquote
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

ROOT=Path(__file__).resolve().parents[1]
DOCS=ROOT/'docs'
PREFIX='/artifact-tool-field-guide/'
GUIDE='work/processes/model-training/research/2026-09-15-artifact-tool-field-guide/field-guide.html'
WALK='skills/system/modelling/research/2026-09-15-excel-tool-walkthrough/walkthrough.html'


def check_file_links():
    checked=set();missing=set();private=set()
    def check(rel,url):
        if not url or url.startswith(('#','data:','blob:','mailto:','javascript:')):return
        u=urlsplit(urljoin('https://public.test'+PREFIX+rel,url))
        if u.hostname=='hub.lynott.co':private.add((rel,url));return
        if u.hostname!='public.test':return
        path=unquote(u.path)
        if not path.startswith(PREFIX):missing.add((rel,url));return
        dest=DOCS/path[len(PREFIX):]
        if dest.is_dir():dest=dest/'index.html'
        checked.add(str(dest.relative_to(DOCS)))
        if not dest.is_file():missing.add((rel,url))
    def nested(rel,x,key=''):
        if isinstance(x,dict):
            for k,v in x.items():nested(rel,v,k)
        elif isinstance(x,list):
            if key=='links':
                for a in x:
                    if isinstance(a,list) and len(a)==2 and isinstance(a[1],str):check(rel,a[1])
            for v in x:nested(rel,v,key)
        elif isinstance(x,str):
            if key in ['image','href','src','sourceHref','rawHref']:check(rel,x)
            if '<a ' in x or '<img ' in x:
                for tag in BeautifulSoup(x,'html.parser').find_all(['a','img']):check(rel,tag.get('href') or tag.get('src'))
    for p in DOCS.rglob('*.html'):
        rel=p.relative_to(DOCS).as_posix();soup=BeautifulSoup(p.read_text(),'html.parser')
        for tag in soup.find_all(['a','img','link','script']):check(rel,tag.get('href') or tag.get('src'))
        for s in soup.find_all('script',attrs={'data-pk-dependency':True}):
            code=s.string or ''
            m=re.match(r'window\.[\w]+\s*=\s*(\{.*\})\s*;?\s*$',code,re.S)
            if m:
                nested(rel,json.loads(m.group(1)))
    return {'checked_files':len(checked),'missing':sorted(missing),'private_links':sorted(private)}


def browser(base):
    errors=[];bad=[];hub=[];states=[];links=set();shots=ROOT/'qa/screenshots';shots.mkdir(parents=True,exist_ok=True)
    with sync_playwright() as pw:
        b=pw.chromium.launch();ctx=b.new_context(viewport={'width':1440,'height':1000});p=ctx.new_page()
        p.on('pageerror',lambda e:errors.append(str(e)))
        p.on('response',lambda r:bad.append((r.status,r.url)) if r.status>=400 else None)
        p.on('request',lambda r:hub.append(r.url) if 'hub.lynott.co' in r.url or '/_hub/' in r.url or '/_cmt/' in r.url else None)
        def capture_links():links.update(p.eval_on_selector_all('a[href],img[src]','els=>els.map(e=>e.href||e.src)'))
        def visit(path):
            p.goto(base+path,wait_until='networkidle');capture_links()
        def reveal(loc):
            loc.evaluate('(e)=>{for(let p=e.parentElement;p;p=p.parentElement)if(p.tagName==="DETAILS")p.open=true}');loc.scroll_into_view_if_needed()
        visit('');assert p.title()=="Artifact Tool: OpenAI's Modeling Engine",p.title()
        assert p.locator('main h1').first.inner_text()=="Artifact Tool: OpenAI's Modeling Engine"
        p.screenshot(path=str(shots/'desktop.png'))
        # All practical actions and their recorded examples, including rendered links.
        ops=p.locator('#af-common-jump option').evaluate_all('els=>els.map(e=>e.value)')
        for op in ops:
            sel=p.locator('#af-common-jump');reveal(sel);sel.select_option(op)
            earlier=p.locator('.ix-earlier')
            if earlier.count():earlier.evaluate('(e)=>e.open=true')
            examples=p.locator('#af-common-example option').evaluate_all('els=>els.map(e=>e.value)')
            for example in examples:
                sel=p.locator('#af-common-example');reveal(sel);sel.select_option(example);capture_links();states.append('common:'+op+':'+example)
            for view in ['raw','tree','readable']:
                loc=p.locator('[data-cw-view="'+view+'"]');reveal(loc);loc.click();capture_links()
        # Exercise declared states in the original component manifest.
        visit(GUIDE)
        inventory=p.locator('#pagekit-component-inventory').evaluate('(e)=>JSON.parse(e.textContent)')
        for comp in [inventory]:
            for exhibit in comp.get('exhibits',[]):
                for state in exhibit.get('states',[]):
                    for act in state.get('actions',[]):
                        loc=p.locator(act['selector']).first
                        if not loc.count():raise AssertionError('Missing control '+act['selector']+' in '+exhibit['id']+':'+state['id'])
                        reveal(loc)
                        if act['action']=='click':loc.click()
                        elif act['action']=='select':loc.select_option(act['value'])
                        elif act['action']=='fill':loc.fill(act['value'])
                    capture_links();states.append(exhibit['id']+':'+state['id'])
        # Read every satellite document; the embedded reader keeps its full content.
        visit(GUIDE+'#satellite-documents')
        doc_ids=p.evaluate('window.ArtifactSatelliteDocuments.documents.map(x=>x.id)')
        for d in doc_ids:
            visit(GUIDE+'?doc='+d+'#satellite-reader');assert p.locator('#af-satellites').inner_text().strip();states.append('document:'+d)
        visit(WALK);assert p.locator('main').inner_text().strip();p.screenshot(path=str(shots/'walkthrough.png'));states.append('walkthrough')
        visit(GUIDE+'?common=map&inspect=sheet-only&iview=raw#common-work');assert p.locator('#af-common').inner_text().strip();p.screenshot(path=str(shots/'inspection.png'))
        p.set_viewport_size({'width':390,'height':844});visit(GUIDE);p.screenshot(path=str(shots/'mobile.png'));assert p.evaluate('document.documentElement.scrollWidth <= innerWidth+2'), 'Mobile horizontal overflow'
        # A source document must render and its raw file must remain available.
        visit('work/processes/model-training/research/2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md.html');assert p.locator('article h1').count();states.append('markdown-reader')
        for url in links:
            u=urlsplit(url)
            if 'hub.lynott.co' in u.netloc:hub.append(url)
            if u.netloc==urlsplit(base).netloc and u.path.startswith(PREFIX):
                dest=DOCS/unquote(u.path[len(PREFIX):])
                if dest.is_dir():dest=dest/'index.html'
                if not dest.is_file():bad.append(('missing dynamic link',url))
        ctx.close();b.close()
    return {'states':states,'state_count':len(states),'links':len(links),'errors':errors,'bad_responses':bad,'hub_requests':hub}


if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--base',default='http://127.0.0.1:8864'+PREFIX);ap.add_argument('--static-only',action='store_true');args=ap.parse_args()
    report={'static':check_file_links()}
    if not args.static_only:report['browser']=browser(args.base)
    out=ROOT/'qa';out.mkdir(exist_ok=True);(out/'browser-results.json').write_text(json.dumps(report,indent=2));print(json.dumps(report,indent=2))
    assert not report['static']['missing'] and not report['static']['private_links']
    if 'browser' in report:assert not any(report['browser'][k] for k in ['errors','bad_responses','hub_requests'])

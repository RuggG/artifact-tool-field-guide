#!/usr/bin/env python3
"""Export the complete field guide as a standalone GitHub Pages publication.

The public output is generated from saved source pages. Hub files are never edited.
"""
import argparse
import hashlib
import html
import json
import posixpath
import re
import shutil
from pathlib import Path
from urllib.parse import urlsplit, unquote

import markdown
from bs4 import BeautifulSoup, Comment

BASE = '/artifact-tool-field-guide/'
GUIDE = 'work/processes/model-training/research/2026-09-15-artifact-tool-field-guide'
REFERENCE = 'work/processes/model-training/research/2026-09-15-artifact-tool-reference'
WALK = 'skills/system/modelling/research/2026-09-15-excel-tool-walkthrough'
HOME = GUIDE + '/field-guide.html'
REPO = 'https://github.com/RuggG/artifact-tool-field-guide'
TITLE = "Artifact Tool: OpenAI's Modeling Engine"
OUT = Path(__file__).resolve().parents[1] / 'docs'
READABLE = {'.md', '.json', '.ndjson', '.txt', '.log', '.py', '.js', '.mjs', '.ts', '.css', '.yaml'}


def nav():
    return f'''<header class="public-top"><a class="public-brand" href="{BASE}">{html.escape(TITLE)}</a><nav aria-label="Public guide"><a href="{BASE}{HOME}#common-work">Explore</a><a href="{BASE}{HOME}#satellite-documents">Documents</a><a href="{BASE}{REFERENCE}/README.md.html">Reference</a><a href="{BASE}{WALK}/walkthrough.html">Walkthrough</a><a href="{REPO}">GitHub</a></nav></header>'''


def public_urls(text):
    return text.replace('https://hub.lynott.co/', BASE).replace('http://hub.lynott.co/', BASE)


def decorate_links(soup, rel):
    for el in soup.find_all(['a', 'img', 'script', 'link', 'source']):
        for key in ('href', 'src'):
            if not el.has_attr(key):
                continue
            url = public_urls(el[key])
            if url.startswith(('/work/', '/skills/', '/system/')):
                url = BASE + url.lstrip('/')
            el[key] = url


def export_html(raw, rel):
    soup = BeautifulSoup(raw, 'html.parser')
    main = soup.find('main')
    # Some captured HTML files are paste fixtures, not publications.
    if not main:
        return public_urls(raw)
    # Internal publication metadata stays in the Hub edition.
    for panel in main.select('.pk-page-details, .version-history-top, #version-history, .version-rounds'):
        panel.decompose()
    for controls in main.select('.reader-controls'):
        if not controls.find(True):
            controls.decompose()
    styles = '\n'.join(str(s) for s in soup.head.find_all('style'))
    title = soup.title.get_text() if soup.title else Path(rel).name
    # Preserve the local reading interactions, not the private Hub services.
    frame_scripts = []
    for s in soup.find_all('script'):
        code = s.string or s.get_text()
        if s.find_parent('main'):
            continue
        if ('/* HubFrame.onDerive(fn)' in code or '// reader chrome: explorable flow' in code
                or '/* the outline and the scroll-spy read the page' in code):
            frame_scripts.append(str(s))
    rail = soup.select_one('aside.hub-rail')
    rail_html = ''
    if rail:
        # The first card is the generated page outline and Expand all control.
        first = rail.select_one('.ctx-card')
        if first:
            rail_html = '<aside class="public-outline hub-rail" aria-label="On this page">' + str(first) + '</aside>'
    attrs = ' '.join(f'{k}="{html.escape(str(v), quote=True)}"' for k,v in (soup.body.attrs if soup.body else {}).items()
                     if k in ('data-docid', 'data-docver'))
    result = f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="description" content="An interactive guide to Artifact Tool for Excel: actions, actual inputs and outputs, formula reference, spreadsheet skills and workflows."><title>{html.escape(title)}</title>{styles}<link rel="stylesheet" href="{BASE}assets/public.css"></head><body {attrs}>{nav()}<div class="public-layout"><div class="hub-stage">{main}</div>{rail_html}</div>{''.join(frame_scripts)}<script src="{BASE}assets/public.js" defer></script></body></html>'''
    result = public_urls(result)
    clean = BeautifulSoup(result, 'html.parser')
    for c in clean.find_all(string=lambda s:isinstance(s, Comment)):
        c.extract()
    decorate_links(clean, rel)
    return str(clean)


def reader(rel, text):
    path = Path(rel)
    title = path.name
    if path.suffix == '.md':
        content = markdown.markdown(text, extensions=['tables','fenced_code','toc','sane_lists'])
        soup = BeautifulSoup(content, 'html.parser')
        heading = soup.find('h1')
        if heading:
            title = heading.get_text()
        decorate_links(soup, rel)
        content = str(soup)
    else:
        try:
            if path.suffix == '.json':
                text = json.dumps(json.loads(text), ensure_ascii=False, indent=2)
        except (ValueError,TypeError):
            pass
        content = f'<h1>{html.escape(path.name)}</h1><pre class="public-code"><code>{html.escape(text)}</code></pre>'
    return f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)} · Artifact Tool field guide</title><link rel="stylesheet" href="{BASE}assets/public.css"></head><body class="public-reader">{nav()}<main class="public-document"><p class="public-filebar"><a data-public-raw href="{html.escape(path.name)}">Raw file</a> · <a data-public-raw download href="{html.escape(path.name)}">Download</a></p><article>{public_urls(content)}</article></main><script src="{BASE}assets/public.js" defer></script></body></html>'''


def source_file(roots, rel):
    for root in roots:
        p = root / rel
        if p.is_file():
            return p


def main():
    ap=argparse.ArgumentParser();ap.add_argument('--hub-root',type=Path,required=True);ap.add_argument('--extra-root',type=Path,action='append',default=[]);args=ap.parse_args()
    roots=[args.hub_root]+args.extra_root
    OUT.mkdir(parents=True,exist_ok=True)
    copied=[]
    for folder in [GUIDE,REFERENCE,WALK]:
        root=next((r for r in roots if (r/folder).is_dir()),None)
        if root is None:raise SystemExit('Missing source folder: '+folder)
        for src in (root/folder).rglob('*'):
            if not src.is_file() or '__pycache__' in src.parts or src.suffix=='.pyc':continue
            rel=src.relative_to(root).as_posix();dest=OUT/rel;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(src,dest);copied.append(rel)
    # Collect any additional real files linked from the saved Markdown documents.
    seen=set();queue=list(copied);missing=[]
    while queue:
        rel=queue.pop()
        if rel in seen:continue
        seen.add(rel);p=OUT/rel
        if p.suffix not in {'.md','.html'}:continue
        raw=p.read_text(errors='replace')
        if p.suffix=='.md':raw=markdown.markdown(raw,extensions=['tables','fenced_code'])
        soup=BeautifulSoup(raw,'html.parser');region=soup.find('main') or soup
        for a in region.find_all(['a','img']):
            u=a.get('href') or a.get('src') or '';bits=urlsplit(u)
            if bits.scheme and bits.netloc!='hub.lynott.co':continue
            if not bits.path or bits.path.startswith('data:'):continue
            target=unquote(bits.path)
            if target.startswith('/'):target=target.lstrip('/')
            else:target=posixpath.normpath(posixpath.join(posixpath.dirname(rel),target))
            if target.startswith('../'):continue
            if (OUT/target).is_file():continue
            src=source_file(roots,target)
            if src:
                dest=OUT/target;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(src,dest);copied.append(target);queue.append(target)
            else:missing.append({'from':rel,'url':u,'target':target})
    # Export all pages from their original source bytes, retaining original URL paths.
    for rel in copied:
        p=OUT/rel
        if p.suffix=='.html':p.write_text(export_html(p.read_text(),rel))
        elif p.suffix in READABLE:
            text=p.read_text(errors='replace');Path(str(p)+'.html').write_text(reader(rel,text))
    assets=OUT/'assets';assets.mkdir(exist_ok=True)
    shutil.copy2(Path(__file__).with_name('public.css'),assets/'public.css')
    shutil.copy2(Path(__file__).with_name('public.js'),assets/'public.js')
    readable=[rel for rel in copied if (OUT/(rel+'.html')).is_file()]
    (assets/'documents.json').write_text(json.dumps(readable,separators=(',',':')))
    (OUT/'.nojekyll').touch()
    (OUT/'index.html').write_text(f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(TITLE)}</title><meta name="description" content="Explore Artifact Tool for Excel, with recorded inputs and outputs and complete supporting documentation."><meta http-equiv="refresh" content="0;url={HOME}"><script>location.replace({json.dumps(HOME)}+location.search+location.hash)</script></head><body><a href="{HOME}">{html.escape(TITLE)}</a></body></html>''')
    report={'files':len(copied),'readers':len(readable),'bytes':sum(p.stat().st_size for p in OUT.rglob('*') if p.is_file()),'additional':sorted(set(copied)-set(x for x in copied if any(x.startswith(y+'/') for y in [GUIDE,REFERENCE,WALK]))),'unresolved_source_links':missing}
    (Path(__file__).resolve().parents[1]/'export-report.json').write_text(json.dumps(report,indent=2))
    print(json.dumps(report,indent=2))


if __name__=='__main__':main()

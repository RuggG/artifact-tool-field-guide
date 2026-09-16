"""A grouped source catalogue with a full-width, offline Markdown reader."""
import html
import json
from pathlib import Path
from pagekit_components import ComponentResult

HERE = Path(__file__).resolve().parent
def render(payload, *, body, context):
    esc = html.escape
    groups = []
    for group in payload['groups']:
        rows = []
        for doc in payload['documents']:
            if doc['group'] != group['id']:
                continue
            rows.append(f'''<a class="sd-entry" href="?doc={doc['id']}#satellite-reader" data-sd-doc="{doc['id']}">
                <span class="sd-entry-title">{esc(doc['title'])}</span><span class="sd-count">{doc['tokens']:,} <small>tokens</small></span>
                <span class="sd-entry-summary">{esc(doc['summary'])}</span></a>''')
        groups.append(f'<section class="sd-group"><h3>{esc(group["title"])}</h3><p>{esc(group["summary"])}</p><div>{"".join(rows)}</div></section>')
    safe = json.dumps(payload, ensure_ascii=False, separators=(',', ':')).replace('<', '\\u003c').replace('\u2028', '\\u2028').replace('\u2029', '\\u2029')
    html_out = f'''<div id="af-satellites" class="sd-library">
      <div class="sd-overview"><strong>1 main skill + 15 supporting documents</strong><span>{payload['totals']['tokens']:,} tokens in total</span></div>
      <p class="sd-count-note">Counts cover the complete Markdown, including code and metadata. The total is the size of this document collection; the agent reads supporting documents as needed.</p>
      <div class="sd-directory">{''.join(groups)}</div>
      <section id="satellite-reader" class="sd-reader" aria-label="Document reader" tabindex="-1"><p>Select a document to read its full text here.</p></section>
      <noscript><p>The reader needs JavaScript. <a href="../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md">Open the saved main skill</a> to follow its source links.</p></noscript>
    </div>'''
    states = [dict(id='directory',label='All documents with token counts',status='supported',actions=[],selector='#af-satellites .sd-directory')]
    for doc in payload['documents']:
        states.append(dict(id=doc['id'],label=doc['title'],status='supported',actions=[dict(action='click',selector=f'[data-sd-doc="{doc["id"]}"]')],selector='#satellite-reader'))
    states.append(dict(id='raw',label='Complete original Markdown',status='supported',actions=[dict(action='click',selector='[data-sd-view="raw"]')],selector='#satellite-reader'))
    return ComponentResult(html_out, dependencies=[
        dict(id='artifact-satellites-data',kind='js',content='window.ArtifactSatelliteDocuments='+safe+';'),
        dict(id='artifact-satellites-css',kind='css',path=str(HERE/'satellites.css')),
        dict(id='artifact-satellites-js',kind='js',path=str(HERE/'satellites.js')),
    ], exhibits=[dict(id='af-satellites',kind='document-library',label='Spreadsheet skill and 15 satellite documents',
                     evidence=dict(format='json',data=payload,source='data/satellite-documents.json'),states=states)])

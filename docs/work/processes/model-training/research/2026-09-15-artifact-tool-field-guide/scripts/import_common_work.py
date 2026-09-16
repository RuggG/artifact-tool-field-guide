"""Preserve the Man Group guide's delivered recordings in this guide's own home.

Usage: python3 scripts/import_common_work.py /path/to/excel-tool-walkthrough
This imports saved evidence; it does not execute any workbook code.
"""
import hashlib
import json
import re
import shutil
import sys
from html.parser import HTMLParser
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
SOURCE_URL = 'https://hub.lynott.co/skills/system/modelling/research/2026-09-15-excel-tool-walkthrough/walkthrough.html'


class Capture(HTMLParser):
    active = False
    data = ''
    def handle_starttag(self, tag, attrs):
        if tag == 'script' and dict(attrs).get('id') == 'operations-reference-data':
            self.active = True
    def handle_endtag(self, tag):
        if tag == 'script':
            self.active = False
    def handle_data(self, text):
        if self.active:
            self.data += text


class SafeFragment(HTMLParser):
    def handle_starttag(self, tag, attrs):
        assert tag in {'p','b','strong','em','i','span','div','table','thead','tbody','tr','th','td','code','pre','small','br','details','summary','ul','ol','li','dl','dt','dd'}, tag
        assert all(k in {'class','style','title','open','colspan','rowspan','tabindex'} for k, v in attrs), attrs
        assert not any('url(' in (v or '').lower() for k, v in attrs), attrs


def main(source):
    parser = Capture()
    page = source / 'walkthrough.html'
    parser.feed(page.read_text())
    data = json.loads(parser.data)
    ref = json.loads((source / 'evidence/operations-reference.json').read_text())
    assert data['operations'] == ref['operations']
    data.pop('aliases', None)
    data['groups'] = ref['groups']
    family = dict(create='workbooks', import_='files', save='files', runner='changes', map='inspection', grids='inspection', regions='inspection', search='inspection', ranges='ranges', trace='inspection', recalculate='formulas', sheet='workbooks', values='ranges', formulas='formulas', fill='ranges', format='formatting', clear='ranges', table='tables', chart='charts', render='files')
    family.update({'import':family.pop('import_'), 'table-row':'tables','table-delete':'tables','chart-delete':'charts'})
    receipts = []

    def preserve(relative):
        path = relative.split('?')[0]
        original = (source / path).resolve()
        assert original.is_relative_to(source.resolve()) and original.is_file(), original
        destination = HERE / 'files/common-work' / path
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(original, destination)
        receipts.append({'path':path,'sha256':hashlib.sha256(original.read_bytes()).hexdigest()})
        return 'files/common-work/' + relative

    for op in data['operations']:
        op['family'] = family[op['id']]
        op['commonApi'] = op['id'] not in ('runner', 'recalculate')
    for ex in data['examples'].values():
        for key in ('inputReadable', 'outputReadable', 'checkReadable'):
            SafeFragment().feed(ex[key])
            ex[key] = ex[key].replace('Exact input', 'Raw').replace('Object tree', 'Tree')
        ex['note'] = ex['note'].replace('Object tree', 'Tree')
        if ex['outputLabel'] == 'Operation return':
            ex['outputLabel'] = 'Recorded code-block return'
        ex['links'] = [[label,preserve(path)] for label,path in ex['links']]
        if ex['image']:
            ex['image'] = preserve(ex['image'])
    # Local guide copy; keep the captured input/output unchanged.
    notes = {'op:add-linked-build': 'The example sets the assumptions, financial logic and cell references '
                            'explicitly.',
     'op:clear-formulas': 'Clearing contents leaves the cells in place; deleting a row or tab changes '
                          'the workbook structure.',
     'op:overlap-table': 'The API accepts overlapping table ranges. Check their addresses before '
                         'creating tables to avoid conflicts in Excel.',
     'old:runner': 'The program runs in Node. Messages such as “saved: overview.json” come from the '
                   'script’s recording helper.',
     'old:render': 'Rendering returns an image blob. The script saves it, then reports the filename '
                   'and size. The image shows the workbook’s current values.',
     'old:export': 'Saving returns no report; the script adds the filename, size and file comparisons. '
                   'Shared-formula import gaps remain in this copy. Checks compare formula text, '
                   'literal cells, sheet order, annotation count and cached numbers using 1e-8 '
                   'absolute / 1e-10 relative tolerance.',
     'old:ranges': 'Values and formulas come from separate native reads. The script pairs them; the '
                   'readable view adds metric and period labels from the workbook and rounds the '
                   'numbers. Tree and Raw keep the full captured precision.',
     'lab:search-pattern': 'The pattern searches across all tabs. The response reports when the match '
                           'limit is reached; matching cells still need interpretation.'}
    for key, note in notes.items():
        data["examples"][key]["note"] = note
    for ex in data["examples"].values():
        ex["note"] = re.sub(r"Size of the returned NDJSON text: (.*?) characters; approximately (.*?) tokens using o200k_base\. This is not a billing count or the size of the whole outer tool response\. NDJSON means one JSON record per line\.", r"Returned NDJSON: \1 characters · approximately \2 tokens.", ex["note"])
        ex["inputReadable"] = ex["inputReadable"].replace("Plain-English reading of the actual input. Requested limits describe the request, not a guarantee of what the library returns.", "Requested limits: compare these with the returned previews below.")
    for op in data["operations"]:
        if op["id"] == "sheet":
            op["detail"] = "Adding a tab creates an empty sheet. The Fee lab example then writes a schedule alongside the six existing sheets."
    for path in ('operations-run/man-group-operations.xlsx','operations-run/new-fee-workbook.xlsx','operations-run/independent-verification.json'):
        preserve(path)
    data['provenance'] = {'sourceUrl':SOURCE_URL,'sourceVersion':7,'packageVersion':'2.8.59','evidenceDate':'2026-09-15','integrationDate':'2026-09-16','sourceHtmlSha256':hashlib.sha256(page.read_bytes()).hexdigest(),'actions':len(data['operations']),'uniqueExamples':len(data['examples']),'exampleAssignments':sum(len(x['examples']) for x in data['operations'])}
    (HERE / 'data/common-work.json').write_text(json.dumps(data, ensure_ascii=False, indent=2)+'\n')
    receipt = {**data['provenance'],'files':sorted({x['path']:x for x in receipts}.values(),key=lambda x:x['path'])}
    (HERE / 'data/common-work-receipt.json').write_text(json.dumps(receipt,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(data['provenance'],indent=2))


if __name__ == '__main__':
    main(Path(sys.argv[1]))

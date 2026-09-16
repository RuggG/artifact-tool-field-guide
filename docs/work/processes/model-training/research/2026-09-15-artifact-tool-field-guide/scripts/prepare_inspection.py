"""Derive coverage from saved returns and check it against separate direct reads."""
import hashlib
import json
import re
from collections import Counter
from pathlib import Path

HOME = Path(__file__).resolve().parents[1]
raw = json.loads((HOME / 'files/inspection/captures.json').read_text())
baseline = {s['sheet']: s for s in raw['baseline']}
checks = []


def check(name, condition):
    assert condition, name
    checks.append(name)


def position(address):
    address = address.split('!')[-1].replace('$', '').split(':')[0]
    col, row = re.fullmatch(r'([A-Z]+)(\d+)', address).groups()
    n = 0
    for c in col:
        n = n * 26 + ord(c) - 64
    return int(row) - 1, n - 1


def source_value(sheet, address, row=0, col=0, field='values'):
    r, c = position(address)
    sr, sc = position(baseline[sheet]['address'])
    rows = baseline[sheet][field]
    return rows[r - sr + row][c - sc + col]


def letters(col):
    value = ''
    while col >= 0:
        value = chr(65 + col % 26) + value
        col = col // 26 - 1
    return value


def records(case):
    r = case['response']
    if case['method'] == 'range':
        return [dict(kind='directRange', sheet='Fee engine', address='C4:H12', **r)]
    return [json.loads(line) for line in r['ndjson'].splitlines() if line] if 'ndjson' in r else r['matches']


byid = {c['id']: c for c in raw['cases']}
full = records(byid['table-workbook-150000'])
extents = {r['sheet']: r['address'] for r in full}
cases = []
for case in raw['cases']:
    result = case['response']
    rows = records(case)
    if 'ndjson' in result:
        check(case['id'] + ': record count', len(rows) == result['recordCount'])
    coverage = []
    truncated_text = 0
    for name, base in baseline.items():
        found = [r for r in rows if r.get('sheet', r.get('name')) == name]
        entry = dict(sheet=name, extent=extents[name], structure=any(r['kind'] == 'sheet' for r in found),
                     grids=[], previews=[], formulas=[], matches=[], valueCells=0, previewCells=0)
        for r in found:
            kind = r['kind']
            if kind == 'directRange':
                rr, cc = position(r['address'])
                entry['formulas'] += [letters(cc + ci) + str(rr + ri + 1) for ri, row in enumerate(r['formulas']) for ci, f in enumerate(row) if f]
            if kind in ('table', 'directRange') and 'values' in r:
                values = r['values']
                address = r.get('valuesPreviewAddress', r['address'])
                count = sum(map(len, values))
                entry['valueCells'] += count
                entry['grids'].append(dict(address=address, described=r['address'], rows=len(values),
                                           cols=max(map(len, values), default=0), cells=count,
                                           isPreview=r.get('valuesTruncated', False)))
                for ri, rr in enumerate(values):
                    for ci, value in enumerate(rr):
                        actual = source_value(name, address, ri, ci)
                        if value != actual:
                            check(case['id'] + ': clipped text matches source', isinstance(value, str) and isinstance(actual, str)
                                  and value.endswith('...') and actual.startswith(value[:-3]))
                            truncated_text += 1
                check(case['id'] + ': grid agrees with direct read', all(
                    value == source_value(name, address, ri, ci) or
                    (isinstance(value, str) and isinstance(source_value(name, address, ri, ci), str) and value.endswith('...')
                     and source_value(name, address, ri, ci).startswith(value[:-3]))
                    for ri, rr in enumerate(values) for ci, value in enumerate(rr)))
            if kind == 'region' and 'preview' in r:
                values = r['preview']; address = r['previewAddress']
                entry['previewCells'] += sum(map(len, values))
                entry['previews'].append(dict(address=address, described=r['address'], rows=len(values), cols=max(map(len, values), default=0)))
                check(case['id'] + ': preview dimensions', r['previewRows'] == len(values) and r['previewCols'] == max(map(len, values)))
                for ri, rr in enumerate(values):
                    for ci, value in enumerate(rr):
                        actual = source_value(name, address, ri, ci)
                        if isinstance(actual, str) and value != actual:
                            check(case['id'] + ': region text prefix', value.endswith('...') and actual.startswith(value[:-3]))
                            truncated_text += 1
            if kind == 'formula':
                entry['formulas'].append(r['address'])
                check(case['id'] + ': formula ' + r['address'], r['formula'] == source_value(name, r['address'], field='formulas'))
            if kind == 'match':
                entry['matches'].append(r['address'])
                actual = source_value(name, r['address'])
                matches = r['value'] == actual
                if isinstance(actual, (int, float)) and isinstance(r['value'], str):
                    matches = float(r['value']) == actual
                check(case['id'] + ': match ' + r['address'], matches)
        coverage.append(entry)
    notices = list(result.get('metadata', {}).get('notices', result.get('notices', [])))
    notices += [r['message'] for r in rows if r['kind'] == 'notice']
    stats = dict(recordTypes=dict(Counter(r['kind'] for r in rows)), recordCount=len(rows),
                 characters=len(result.get('ndjson', json.dumps(result, ensure_ascii=False))),
                 valueCells=sum(s['valueCells'] for s in coverage), previewCells=sum(s['previewCells'] for s in coverage),
                 formulaCells=sum(len(s['formulas']) for s in coverage), matchCells=sum(len(s['matches']) for s in coverage),
                 previewGrids=sum(g['isPreview'] for s in coverage for g in s['grids']), shortenedTextCells=truncated_text,
                 omittedRecords=result.get('truncated', False) if case['method'] == 'inspect' else None,
                 notices=notices)
    case = {k: v for k, v in case.items() if k != 'elapsedMs'}
    case.update(analysis=dict(coverage=coverage, **stats))
    cases.append(case)


def rec(id):
    return records(byid[id])


check('Source SHA-256 unchanged', hashlib.sha256((HOME / raw['receipt']['source']).read_bytes()).hexdigest() == raw['receipt']['sourceSha256'])
check('No native Excel Tables in fixture', all(s['nativeTables'] == 0 for s in baseline.values()))
check('Large allowance ignores 1x1 fallback', len(rec('table-sheet-150000')[0]['values']) == 109 and len(rec('table-sheet-150000')[0]['values'][0]) == 23)
check('Small allowance returns one blank cell', rec('table-sheet-1000')[0]['values'] == [[None]])
check('Preview and outer omission are distinct', rec('table-sheet-1000')[0]['valuesTruncated'] and not byid['table-sheet-1000']['response']['truncated'])
check('A narrow range fits small allowance', rec('table-range-1000') == rec('table-range-150000'))
check('Full grid unaffected by preview dimensions', rec('table-sheet-150000') == rec('table-sheet-wide-150000'))
check('Formula maxResults ignored in tested range', rec('formulas-range-3') == rec('formulas-range-300') and len(rec('formulas-range-3')) == 13)
check('Formula maxChars does omit records', byid['formulas-small-budget']['response']['truncated'])
check('include formulas does not add formula records', rec('project-values') == rec('table-range-150000'))
check('Search offset skips first three matches', rec('match-offset')[:3] == rec('match-all')[3:6])
check('findCells retains total and limit', byid['find-cells']['response']['total'] == 28 and byid['find-cells']['response']['truncated'])
check('Inspect match limit has its own notice', not byid['match-limit']['response']['truncated'] and any('Cell search truncated' in r.get('message', '') for r in rec('match-limit')))
data = dict(receipt=raw['receipt'], sheets=[dict(sheet=s, extent=extents[s], directReadExtent=b['address'], nativeTables=b['nativeTables']) for s,b in baseline.items()], cases=cases)
(HOME / 'data/inspection.json').write_text(json.dumps(data, ensure_ascii=False, separators=(',', ':')))
(HOME / 'data/inspection-checks.json').write_text(json.dumps(dict(passed=len(checks), checks=checks, sourceSha256=raw['receipt']['sourceSha256'], captureSha256=hashlib.sha256((HOME/'files/inspection/captures.json').read_bytes()).hexdigest()), indent=2))
print(json.dumps(dict(cases=len(cases), checks=len(checks), gridCells=sum(r['rows']*r['cols'] for r in full), formulaRecords=len(rec('formulas-sheet-3')))))

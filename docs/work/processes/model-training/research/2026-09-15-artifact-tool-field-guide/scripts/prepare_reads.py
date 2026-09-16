"""Check fresh read-only captures against the separately saved v4 direct reads."""
import hashlib
import json
import re
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
D = json.loads((ROOT/'files/reads/captures.json').read_text())
B = {s['sheet']:s for s in json.loads((ROOT/'files/inspection/captures.json').read_text())['baseline']}
checks = 0
def check(value, note):
    global checks
    assert value, note
    checks += 1
def origin(address):
    m=re.match(r'([A-Z]+)(\d+)',address.replace('$','')); col=0
    for char in m[1]: col=col*26+ord(char)-64
    return int(m[2])-1,col-1
def letter(c):
    s=''; c+=1
    while c: c,r=divmod(c-1,26); s=chr(r+65)+s
    return s
def expected(s,field):
    b=B[s['sheet']];r,c=s['row'],s['col']
    return [row[c:c+s['cols']] for row in b[field][r:r+s['rows']]]
check(len(D['cases'])==38,'38 executions')
check(hashlib.sha256((ROOT/D['receipt']['source']).read_bytes()).hexdigest()==D['receipt']['sourceSha256'],'Original source unchanged')
for c in D['cases']:
    o=c['output']; method=c['method']; name=c['id']
    if c.get('selections'):
        for i,s in enumerate(c['selections']):
            vals,forms=expected(s,'values'),expected(s,'formulas')
            out=o[i]['result'] if c.get('multi') else o
            if method=='values': check(out==vals,name)
            if method=='formulas': check(out==forms,name)
            if method=='both': check(out==dict(values=vals,formulas=forms),name)
            if method=='paired':
                check(len(out)==s['rows']*s['cols'],name+' all positions')
                for j,cell in enumerate(out):
                    r,col=divmod(j,s['cols'])
                    check(cell==dict(address=letter(s['col']+col)+str(s['row']+r+1),value=vals[r][col],formula=forms[r][col]),name+' '+cell['address'])
            if method=='details':
                check(out['formulas']==forms,name+' stored')
                check(out['displayFormulas']==forms,name+' displayed')
                for r,row in enumerate(out['formulaInfos']):
                    for col,info in enumerate(row):
                        check(info==dict(kind='stored',formula=forms[r][col],display=forms[r][col],isEditable=True) if forms[r][col] else info is None,name+' info')
    elif method=='inspect':
        records=[json.loads(x) for x in o['ndjson'].splitlines()]
        check(len(records)==o['recordCount'] and not o['truncated'],name+' envelope')
        for rec in records:
            r,col=origin(rec['address']);b=B[rec['sheet']]
            if rec['kind']=='table':
                check(not rec.get('valuesTruncated'),name+' full values')
                check(rec['values']==[row[col:col+rec['cols']] for row in b['values'][r:r+rec['rows']]],name+' grid equality')
            elif rec['kind']=='formula':check(rec['formula']==b['formulas'][r][col],name+' formula equality')
            else:raise AssertionError('Unexpected record '+str(rec))
    elif method=='search':
        check(len(o['matches'])==min(o['limit'],o['total']),name+' match count')
        check(o['truncated']==(o['total']>o['limit']),name+' truncation')
        for match in o['matches']:
            r,col=origin(match['address']);b=B[match['sheet']]
            check(('='+match['formula'] if match['formula'] else '')==b['formulas'][r][col],name+' formula pairing')
            v=b['values'][r][col]
            check(float(match['value'])==v if isinstance(v,(int,float)) else match['value']==v,name+' value pairing')
C={c['id']:c for c in D['cases']}
for m in ('values','formulas','both','paired'):check(C[m+'-cell']['output']==C[m+'-index']['output'],m+' cell vs index')
check(sum(s['rows']*s['cols'] for s in C['both-workbook']['selections'])==14697,'Used areas including formatting')
check(sum(r['rows']*r['cols'] for r in map(json.loads,C['inspect-workbook']['output']['ndjson'].splitlines()) if r['kind']=='table')==8383,'Inspection content rectangles')
check(C['search-sheet']['output']['total']==329 and len(C['search-sheet']['output']['matches'])==100,'Search limit example')
check(C['paired-block']['output'][15:18]==[dict(address=a,value=None,formula='') for a in ['K18','L18','M18']],'Blank separator retained')
(ROOT/'data/reads.json').write_text(json.dumps(D,ensure_ascii=False,separators=(',',':'))+'\n')
(ROOT/'data/reads-checks.json').write_text(json.dumps(dict(assertions=checks,passed=True,cases=38,sourceSha256=D['receipt']['sourceSha256'],reference='files/inspection/captures.json',notes=['All getter results compared with a separately captured baseline.','Inspection formulas and grids checked against baseline.','Search values and formulas checked against baseline; original string representations retained.','Used-area padding and blank separators retained.']),indent=2)+'\n')
print(f'{checks} assertions passed; 38 captures prepared')

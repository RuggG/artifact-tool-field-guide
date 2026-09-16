from pathlib import Path
import json,sys
HERE=Path(__file__).parent
sys.path.insert(0,str(HERE/'tokenizer-deps'))
import tiktoken
enc=tiktoken.get_encoding('o200k_base')
def load(p):return json.loads((HERE/p).read_text())
specs=[
 ('map','workbook','Tab list','overview-probe/metadata-overview.json','All six tabs, with names and occupied ranges. No cell values.','kind selects workbook and sheet descriptions only.'),
 ('broad','workbook','Add values · 14k limit','man-group-run/overview.json','Summary values arrive in full, then the response stops after the Assumptions description.','Adding table requests cell grids as well. The 14,000-character allowance cuts this response short.'),
 ('all','workbook','Add values · 150k limit','overview-probe/full-with-preview-caps.json','All six tabs and all six value grids arrive.','Compared with the previous option, only maxChars changes: 14,000 → 150,000. The requested 4 × 6 preview is unchanged.'),
 ('tab-full','table','150k-character allowance','overview-probe/fee-table-tiny-preview.json','Fee engine returns its entire 109 × 23 value grid.','The table fits within this allowance, so the 1 × 1 fallback preview is not used.'),
 ('tab-small','table','1k-character allowance','overview-probe/fee-table-small-budget.json','Only cell A1 arrives as a 1 × 1 preview. It is blank.','Only maxChars changes. The full table is too large for this allowance, so the preview settings take effect.'),
 ('blocks-small','regions','2 × 3 previews','overview-probe/fee-regions-small-preview.json','Six blocks are identified, each with a 2-row × 3-column preview.','kind: region asks for blocks and samples of their cells, rather than one full sheet grid.'),
 ('blocks-large','regions','4 × 6 previews','overview-probe/fee-regions-larger-preview.json','The same six blocks, now with larger 4-row × 6-column previews.','Only tableMaxRows and tableMaxCols change. The sheet, selected area and character allowance stay the same.'),
 ('search-text','search','Text: profit','man-group-run/literal.json','Candidate cells containing profit, including several different profit measures.','kind: match searches cells. Here useRegex is false, so profit is treated as ordinary text.'),
 ('search-pattern','search','Pattern: fees / margin','man-group-run/regex.json','Cells matching management…fee OR fee…margin. A notice reports the match limit.','The search term becomes a pattern, and useRegex is set to true. The agent still has to interpret the matches.')
]
original_args={
 'broad':dict(kind='workbook,sheet,table',maxChars=14000,tableMaxRows=4,tableMaxCols=6,tableMaxCellChars=80),
 'search-text':dict(kind='match',searchTerm='profit',maxChars=12000,options=dict(useRegex=False,maxResults=30)),
 'search-pattern':dict(kind='match',searchTerm='management.*fee|fee.*margin',maxChars=12000,options=dict(useRegex=True,maxResults=30))
}
presets=[]
for id,group,label,file,outcome,change in specs:
 c=load(file);r=c['result'];args=original_args.get(id,c['input'])
 assert isinstance(args,dict)
 presets.append(dict(id=id,group=group,label=label,file=file,input=args,result=r,records=[json.loads(x) for x in r['ndjson'].splitlines() if x],outcome=outcome,change=change,characters=len(r['ndjson']),tokens=len(enc.encode(r['ndjson']))))
sheets=[x for x in presets[0]['records'] if x['kind']=='sheet']
(HERE/'evidence/inspect-lab.json').write_text(json.dumps(dict(sheets=sheets,presets=presets),indent=2,ensure_ascii=False))
print('Prepared',len(presets),'recorded inspect configurations')

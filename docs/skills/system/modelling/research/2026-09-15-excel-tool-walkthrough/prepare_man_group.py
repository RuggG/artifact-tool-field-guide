from pathlib import Path
import json

HERE = Path(__file__).parent
RUN = HERE / 'man-group-run'
def read(name): return json.loads((RUN / (name + '.json')).read_text())
def card(id, label, title, purpose, returns, finding, notes=(), kind='JavaScript spreadsheet library'):
    record = read(id)
    return dict(id=id, label=label, title=title, purpose=purpose, returns=returns,
                finding=finding, notes=list(notes), kind=kind, input=record['input'],
                result=record['result'], elapsedMs=record['elapsedMs'],
                links=[{'label':'Full captured record','href':f'man-group-run/{id}.json'},
                       {'label':'Executed JavaScript','href':'man-group-run/run.mjs'}])

runner = read('runner-tool-record')
cards=[dict(id='runner',label='Script runner',title='Run a script that works with the workbook',
    purpose='Start the JavaScript program that imports the Man Group file and makes the spreadsheet calls below.',
    returns='The command tool returns printed text, process status and, when the program is still running, a session ID. It does not automatically return the workbook or every intermediate calculation.',
    finding='The first call returned session 22618. A later poll returned exit code 0 and the script’s record-saving messages. Each JSON record was saved separately and read back for this page.',
    kind='Built-in Codex command tool',input=json.dumps(runner['executionInput'],indent=2),
    result=runner,notes=['The messages such as “saved: overview.json” were printed by our recording helper. They are not independent guarantees from Excel.','The program runs in Node. No Excel window or connected add-in session is involved in these calls.'],
    links=[{'label':'Actual command and tool results','href':'man-group-run/runner-tool-record.json'}, {'label':'Executed script','href':'man-group-run/run.mjs'}])]
cards += [
card('overview','Broad workbook read','Request tab descriptions and cell values together',
    'Ask for workbook and sheet descriptions plus cell values. This broad request was intended as an overview, but its returned full value grid makes it large.',
    'A workbook record, sheet records and table-shaped cell data, plus metadata telling us whether the response was cut short.',
    'It reports six sheets. This broad request returns five records, including a notice that nine further lines were omitted. The separate sheet call below gives the complete tab list.',
    ['The six “tables” reported here are inspection representations of sheet data. The XLSX contains no native Excel Table objects.','The requested four-row/six-column table preview is a fallback, not an unconditional cap. The Summary values fit individually, so this call returns all 57 × 13 values. The later response-wide cutoff omits other records.']),
card('sheets','Sheet summaries','List the tabs and their occupied ranges',
    'Choose the relevant tab. For this task, Summary shows the answer, Assumptions owns the drivers and Fee engine does the calculations.',
    'One record per tab, including its name, internal ID, position and used rectangle. This call does not explain what the tabs mean.',
    'All six tabs are returned: Summary, Assumptions, Fee engine, Disclosures, Checks and Readme.',
    ['The used rectangle can include empty cells and layout space. It is not a count of populated cells.','Although the request says include id,name, this runtime also returns index, range and address.']),
card('literal','Text search','Find cells containing “profit”',
    'Turn an imprecise request about profit into a list of candidate labels and locations.',
    'Matching cells with sheet, address, value, formula field and whether the match came from the value or formula.',
    '“profit” returns 28 matches, including operating profit, profit after tax, statutory profit and explanatory notes. The agent must choose the intended metric from that context.',
    ['This is text matching. The call does not decide which definition of profit the user intended.']),
card('regex','Pattern search','Search several fee-related word patterns at once',
    'Find management-fee and fee-margin labels even when words appear in different phrases.',
    'The same cell-match records as text search, using the supplied regular expression.',
    'The expression management.*fee|fee.*margin returns 30 cell matches and an additional limit notice. Both parts of the pattern were chosen by the agent.',
    ['A regular expression matches text patterns. It does not provide a financial synonym dictionary or semantic index.']),
card('regions','Populated regions','Locate the populated blocks within Summary',
    'Break a known section into manageable blocks so the agent can inspect the relevant part next.',
    'Region addresses, dimensions, counts and preview matrices. The preview address tells us exactly which part of a wider block was returned.',
    'Within C7:M28 it finds four blocks: headers/AUM, revenue, operating costs/profit, and the PBT-to-EPS bridge. Previews stop at column H even though each region extends to M.',
    ['The returned text counts include numeric cells rendered as text in this inspection view. They do not prove those cells are stored as text in Excel.']),
card('ranges','Values and formulas','Read a precise rectangle of values and formulas',
    'Inspect the forecast revenue and operating-profit section after locating it.',
    'Two aligned two-dimensional arrays: current values and formula text. The first item in each is K13; rows and columns preserve the requested order.',
    'K13:M20 covers FY2026, H1 2027 and FY2027. The read shows the displayed numbers and the Fee engine references behind them.',
    ['Row labels and period captions in the readable grid are added from the same workbook for orientation; the exact rectangular call returns neither automatically.','Numbers in the readable display are rounded. The tree and raw JSON retain their full captured precision.']),
card('trace','Dependency tree','Trace one management-fee result back to its inputs',
    'Explain how Q3 2026 management fees were calculated and which assumptions they depend on.',
    'A nested tree: each node has a cell, formula, value and a params array of dependencies. Expand a branch to follow it upstream.',
    'Fee engine!R27 is $359.478m: $256.77bn average AUM × 1,000 × 56 / 10,000 × 0.25. The fee-rate branch reaches the selected case and source margin.',
    ['The average-AUM node R18 returns formula: null and no children, although a formula exists in the file. That branch is incomplete because this cell is a shared-formula follower.','The first attempt used Excel-style quotes around the sheet name and returned null. The successful trace call uses Fee engine!R27 without those quotes. The failed attempt is retained in the evidence folder.']),
card('render','Rendered image','Turn a selected worksheet rectangle into an image',
    'Let the agent check the model visually: layout, formatting, clipping and whether the numbers are readable.',
    'render returns an image blob. The script saves its bytes as a PNG; a separate view_image tool then makes those pixels visible to the agent.',
    'This is the actual Summary C7:M28 image rendered from your Man Group workbook before edits.',
    ['The file name and byte count are printed by our wrapper after saving the blob. They are not the image itself.','The screenshot shows the current values. A plausible-looking image does not establish that every formula recalculates correctly.']),
card('value-edit','Cell edit','Change one forecast assumption from 56bp to 51bp',
    'Test whether lowering the Q4 2027 base fee margin reaches the model’s earnings outputs.',
    'The assignment itself has no standalone success report. We explicitly read six relevant cells before and after, including each exposed formula.',
    'Assumptions!W26 changes to 51, but W25 stays at 56 and Q4 fees remain $406.716m. The missing shared-formula dependency prevents the change from flowing through.',
    ['The edit is confined to a disposable copy. The original Downloads file remains byte-for-byte unchanged.','readState is our helper that reads these six cells; its full definition is available below the code and in the executed script.']),
card('recalculate','Recalculation','Ask the calculation engine to refresh the workbook',
    'Check whether an explicit recalculation resolves the unchanged result after the assumption edit.',
    'recalculate returns JavaScript undefined. The useful evidence comes from the explicit before-and-after cell reads.',
    'It makes no difference here: the input is 51bp, the active assumption remains 56bp and fees are unchanged. Recalculation cannot follow a dependency missing from the imported calculation graph.',
    ['This is Artifact Tool recalculation, not Microsoft Excel’s native calculation engine.','A second recalculation was run after restoring the original assumption, before the final checks and export.']),
card('formula-edit','Formula edit','Restore the formula behind the active fee margin',
    'Test whether giving the importer the missing formula explicitly restores the connection from the scenario input.',
    'After assigning formula text, we read the same six cells again to see which values actually move.',
    'The active margin becomes 51bp. Q4 management fees fall to $370.402m and FY2027 fees to $1,531.994m. Operating profit remains $792.754m, so restoring this one formula does not fix the whole dependency chain.',
    ['The formula is the existing scenario-selection rule translated to W25, not a new financial assumption or a hardcoded result.','This is a targeted demonstration of the import issue, not a complete repair of the model.']),
card('fill','Formula fill','Fill the active-assumption formula across the forecast quarters',
    'Show how the script copies a formula pattern from R25 across to W25 while adjusting relative references.',
    'We read the row’s values and formula strings before and after fillRight. The method itself does not print these arrays.',
    'All six quarterly active-assumption cells now expose formulas. R26 becomes S26, T26 and so on; the absolute case selector $D$4 stays fixed.',
    ['The values remain 56, 56, 56, 56, 56 and 51. This operation restores the imported formulas for the row; it does not change the scenario rules.']),
card('copy','Formula copy','Copy the original base formula back into the edited cell',
    'Restore the original assumption after the temporary 51bp test, while demonstrating a precise formula-only copy.',
    'copyFrom copies the formula from R26 into W26. We then read the input and downstream outputs again.',
    'W26 becomes =$D$50 and returns to 56bp. Q4 and FY2027 management fees return to their original values.',
    ['The “formulas” option copies the calculation rule, not a frozen value or the source formatting.']),
card('export','Export','Write the workbook object back to an XLSX file',
    'Produce a file that can be downloaded and opened after the in-memory inspection and edits.',
    'exportXlsx creates a file object; save writes it to disk and returns undefined. The script then checks the saved file’s size.',
    'The saved demo is 106,380 bytes. Independent readback confirms the original sheet order, formulas, cell values and 270 annotations remain, after the temporary assumption change was reversed.',
    ['The workbook still has the broader shared-formula import limitation described above. This export is a demonstration copy, not a claimed complete repair.','Verification compared formula text and literal cells, cached numbers within 1e-8 absolute / 1e-10 relative tolerance, sheet order and annotation count. Native Excel was not opened for this run.'])
]
for c in cards:
    if 'ndjson' in c['result']:
        c['records'] = [json.loads(line) for line in c['result']['ndjson'].splitlines() if line.strip()]
    if c['id']=='render':
        c['image']='man-group-run/summary-before.png'
        c['extraInput']=json.dumps({'path':str(RUN/'summary-before.png')},indent=2)
    if c['id'] in ('value-edit','recalculate','formula-edit','copy'):
        script=(RUN/'run.mjs').read_text()
        c['helper']=script[script.index('function readState()'):script.index("await capture('value-edit'")]
    if c['id']=='export':
        c['verification']=read('verification')
        c['links'] += [{'label':'Download demonstration XLSX','href':'man-group-run/man-group-demo.xlsx'}, {'label':'Independent file checks','href':'man-group-run/verification.json'}]

(HERE/'evidence'/'man-group-cards.json').write_text(json.dumps({'source':'man-group-fees-excel.xlsx','cards':cards},indent=2))
print(f'Prepared {len(cards)} Man Group examples')

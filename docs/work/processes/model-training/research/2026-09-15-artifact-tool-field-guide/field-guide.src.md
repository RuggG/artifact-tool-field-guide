---
title: Artifact Tool: OpenAI's Modeling Engine
class: system-doc
type: Interactive field guide
subtitle: Explore the workbook engine through real inputs, readable objects and recorded results
primary_author: RG
agent_assistance:
  - Codex — experiments, interactive design and verification
docver: 8
history:
  - v8 2026-09-16 — Simplified the introduction and organised inspection and reading examples by purpose
  - v7 2026-09-16 — Added the main skill and all 15 satellite documents, with token counts and a full-text Markdown reader
  - v6 2026-09-16 — Added 38 reading comparisons, values and formulas paired by cell, whole-tab and all-tab reads, and full-width results
  - v5 2026-09-16 — Made actual response records the default read, separated metadata, added all-sheet coverage and scrolling family navigation
  - v4 2026-09-16 — Added 47 inspection comparisons on one unchanged workbook, with cell coverage, preview limits and complete saved response records
  - v3 2026-09-16 — Integrated 23 practical actions and 48 recorded Man Group examples, with separate calls, returns and agent checks
  - v2 2026-09-15 — Added a source-linked manual of workbook design, best practices and end-to-end workflows
  - v1 2026-09-15 — New interactive guide grounded in the installed package and recorded experiments
stub:
  - An interactive guide to Artifact Tool's action families, object structure, inputs, outputs, workflows and verified limits.
sections: fold
default_open: sections
tags:
  - Artifact Tool
  - Spreadsheets
  - Developer reference
---

**Artifact Tool turns code into a workbook, calculated results, pictures and Excel files.** Start with [Common work](#common-work) for practical jobs on an existing model, or follow [one small workbook](#walkthrough) from creation to export. The [full action families](#families) show the wider possibilities; [Best practices](#best-practices) explains how to turn those capabilities into a useful workbook. [Satellite documents](#satellite-documents) lets you read the main skill and its supporting guidance in full.

The practical examples use a [Man Group forecasting model](https://hub.lynott.co/skills/system/modelling/research/2026-09-15-excel-tool-walkthrough/walkthrough.html#ref-map) to show the tool’s inputs, outputs and behaviour.

## Start with the work you will do most often {#common-work} {toc=Common work}

*A practical selection: open a model, understand it, change it, check it and save a usable result.*

Start with **Inspect and locate** for an existing file; use **Open and save → Create a workbook** for a fresh build.

For an ordinary edit: **import → inspect → read the relevant cells → change values or formulas → check calculations and a rendered preview → export, save and reopen**.

::: wide
```component common-work-browser
renderer: scripts/common_work.py
data: data/common-work.json
```
:::

### Read the cells you actually want — One cell, a whole tab or every tab {#reading-choices}

**Read and calculate → Read values and formulas** has two choices: which cells to read, and what to get back. Choose **Values only** for their contents, **Formulas only** for their calculation rules, or **Value and formula together per cell** to see both beside an address. The last layout is assembled by the script from two ordinary reads. The exact code and unaltered captured output are available beside every example.

A range is simply the cells you name, such as `K13:M20`. Getting that range creates a way for the code to refer to those cells; it does not add cells or content. Reading it returns all eight rows and three columns, including the blank row 18. A one-cell read is the same idea on a smaller area. The example using `getCell(12, 10)` reads the same K13 cell using numbers that start at zero.

**To read a whole tab**, the code asks for its used area and reads the cells in it. **To read every tab**, it repeats that for each tab and collects the results. In this workbook those six rectangles contain **14,697 cell positions**, including blank cells with formatting. The inspection route returns **8,383 positions** in smaller rectangles around content. Those are different ways of choosing the area; neither counts every possible empty cell in Excel. The saved all-tab reads contain all six results, and a clearly labelled display selector lets you inspect each one.

Values and formulas are cell contents. These reads do not also return charts, comments or all formatting. Formula results are from the imported workbook; this read-only experiment does not establish that they have been freshly recalculated. Raw also preserves differences in representation: for example, `findCells` returns values as text in these match records, while range reads return numerical values as numbers.

### Know what reaches the agent — Structure, cells, previews and output limits {#inspection-visibility}

**Loading a workbook into the engine does not place it all in the agent's conversation.** The code asks for a view of the workbook, receives a result, and chooses what to print or return. Printing only `result.recordCount` gives the agent a count. Printing `result.ndjson` gives it the records, subject to the surrounding tool's output limit. Saving a complete JSON file requires a later read before its contents reach the conversation. This experiment saved complete record payloads, printed short progress summaries and read selected evidence for analysis.

`inspect` is one function with several configurations. `kind` is a comma-separated list of **record types**. These are the engine's tokens or keywords; they are unrelated to language-model token counts. A request can combine types, but the response still depends on its scope, supported types, existing objects and size allowance. Its `ndjson` property is text containing one JSON object per line. Raw displays that NDJSON directly, with long lines wrapped for reading. Tree expands the decoded records. Readable shows all returned records by default, grouped by type. Metadata and the original response envelope have their own disclosures.

| Request | What the recorded call returns | What this does not establish |
| --- | --- | --- |
| `kind: "workbook,sheet"` | Workbook summary and sheet inventory | No cell-value grid; a sheet address alone does not expose its cells |
| `kind: "table"` | Value grids for the selected scope, in full or as previews | Formula text, formatting and other objects need their own reads |
| `kind: "workbook,sheet,table"` | A combined stream of structure and value records | A small allowance can omit later records or reduce a grid to a preview |
| `kind: "region"` | Detected block addresses, dimensions, population counts and previews | A block's whole rectangle is not included merely because its address appears |
| `kind: "formula"` | A record for each included formula cell | Formula records do not contain every cell or prove the import preserved every original formula |
| `kind: "table,formula"` | Value grids plus separate formula records | The two representations can overlap in the cells they cover |
| `kind: "match"` with `searchTerm` | Matching cells, values/formulas and relevant notices | Matching a label is not a full read of its row, sheet or financial meaning |
| `findCells(...)` | A separate object with matches, total count, offset, limit and truncation | Its result shape and truncation flag differ from `inspect` |
| `getRange(...).values` and `.formulas` | Two complete matrices for an explicitly chosen rectangle | The surrounding script can still choose to expose only part of them |

**Here, `table` means an inspection value grid.** Direct collection reads found zero native named Excel Tables in the supplied six-sheet workbook. Inspection nevertheless reports six table grids. Creating a native Excel Table is a separate feature in Work with Tables.

The broad full-value request returned **8,383 cell positions across six content rectangles**, checked against direct range reads. This includes blanks inside the rectangles; it does not enumerate Excel's unused rows and columns. `getUsedRange()` also included formatting beyond some of those content extents in this file. The map's grey background comes from the separate full reference read, so it must not be mistaken for information present in a smaller response.

### Read each limit separately — A false outer flag is not proof of completeness {#inspection-limits}

The recorded Fee engine request makes the distinction concrete. With `maxChars: 150000`, `kind: "table"` returned **109 × 23 = 2,507 values**, despite fallback settings of one row and one column. Changing only `maxChars` to 1,000 returned **`[[null]]`, the blank cell A1**. The record still described `A1:W109`; `valuesTruncated` was true, while the response's outer `truncated` was false.

| Layer | Evidence to read | Observed behavior |
| --- | --- | --- |
| Whole inspection response | Outer `truncated`, notice records and `recordCount` | Records can be omitted. Notices themselves count as records. The allowance applies to NDJSON, not every character in an outer tool message |
| Included value grid | `valuesTruncated`, `valuesPreviewAddress`, actual `values` dimensions | A record can fit by returning only a preview. Described dimensions can remain much larger |
| Region preview | `previewAddress`, `previewRows`, `previewCols`, `preview` | The tested region settings return samples at the top left of each detected block |
| Text within a preview | Actual strings and `tableMaxCellChars` | A cell can be present but its text shortened, such as `Case se...` |
| Cell search | Search-limit notice; for `findCells`, its `total`, `limit`, `offset` and `truncated` | `inspect` returned three matches and a search-limit notice while its outer `truncated` remained false |
| Delivery to the agent | What the script prints/returns, then the tool's output allowance | A complete library result can be reduced to a count, summary or clipped message |
| This guide's display | All returned records by default; optional record filters, expandable trees and scroll areas | Display controls do not change the executed request. Coverage defaults to all sheets; the download preserves the original saved payload |

**Preview settings are not universal hard caps.** `tableMaxRows` and `tableMaxCols` acted as fallbacks for table value grids: a grid that fit arrived in full. The region calls used the chosen preview dimensions. With the same 1,000-character allowance, narrowing the request to `C4:H12` returned the full 54-position rectangle. A targeted read can therefore expose more useful content than a small preview of an entire tab.

### Choose settings for the question — Tested behavior and useful follow-up reads {#inspection-settings}

| Setting or route | Meaning and evidence from version 2.8.59 |
| --- | --- |
| `sheetId` and `range` | Scope supported output types to a sheet or rectangle. Sheet inventory records still describe the sheet; they do not become cell grids |
| `include` / `exclude` | Choose existing record fields. `include: "values,formulas"` did not add formulas to the table record; use the formula kind or the formulas getter. `exclude: "values"` returned only grid metadata |
| `search` | Filter inspection records. Searching sheet records for “Fee engine” returned the matching sheet record |
| `searchTerm` | Search cell content. `useRegex` accepts a text pattern; `matchFormulas` adds formula text to the match search. Matching values can be serialized as strings |
| `options.maxResults` | Limited cell matches. It did **not** cap formula records: limits 3 and 300 both returned 13 formulas in C4:H12, and 898 on Fee engine. `maxChars` did constrain formula output |
| `offset` | Skip preceding cell matches. The recorded offset 3 returned the next three matches, checked against the complete 28-match result |
| `target.id`, `beforeLines`, `afterLines` | Request context around a record anchor. The recorded sheet anchor plus one following line included its entire value-grid record. These are record lines, not worksheet rows |
| `chart` versus `drawing` | Although the archived inspection example uses `chart`, this runtime reported it as an unknown kind and ignored it. The recorded `drawing` request returned chart/shape descriptors. Inspect metadata exposes the discrepancy |
| `resolve` and `help` | Resolve an inspection anchor to an editable object; use help to discover call contracts. Examples remain in the [full inspection family](?family=inspection&action=#families) |

A useful sequence is **inventory → locate blocks or labels → read the selected range's values and formulas → trace a result if needed**. If the question is “show the whole tab”, request its value grid with adequate space, inspect its actual coverage, and request formula or object information separately when relevant. Completeness is always relative to the requested representation and the imported workbook state.

Sources and reproduction: [all 47 captured calls and responses](files/inspection/captures.json), [read-only capture script](scripts/capture_inspection.mjs), [coverage derivation](scripts/prepare_inspection.py), [2,150 checks](data/inspection-checks.json), [archived inspection documentation](../2026-09-15-artifact-tool-reference/files/sources/package/api/references/inspect-help.md) and [quick-start inspection examples](../2026-09-15-artifact-tool-reference/files/sources/package/API_QUICK_START.md). The workbook checksum was unchanged. These checks verify recorded coverage and selected contrasts; they do not certify native Microsoft Excel behavior.

### Understand the selection — Common work, documentation and the full reference {#common-coverage}

**These counts describe different things.** A documentation heading can contain many calls; a practical action can combine several calls; a recorded example shows one specific execution. They should not be added together or treated as competing measures of completeness.

| Surface | Count | What is being counted |
| --- | --- | --- |
| Package quick start, “High-Value + Common” | 14 headings | Broad API topics, excluding its closing runnable example |
| Spreadsheet skill's quick start | 15 headings | A related but different grouping: Notes and Merging cells have headings; Sparklines appears elsewhere |
| This practical selection | 23 actions in six groups | Common jobs, from opening a model to checking and saving edits |
| Supporting example collection | 48 examples | Includes overlapping demonstrations; the explorer groups them by purpose |
| Inspection comparisons | 47 read-only calls | Variations of inspection, cell search and direct reading on one unchanged workbook; not 47 new operations |
| Reading comparisons | 38 read-only calls | Values, formulas and combined reads for different cell selections |
| Full action explorer below | 15 families, 191 examples | API calls, structured operations and guided workflows across the wider library |

**21 of the 23 practical actions use API features named in the package's common section.** This is a feature-level correspondence, not a claim that the quick start gives every argument or validates every example. Detailed inspection modes are also explained earlier in that document. The two remaining entries are **the script runner**, which executes the program outside the library, and **explicit recalculation**, a library control documented elsewhere. Neither is thereby unimportant or discouraged.

The practical selection is **not the documentation's complete recommendation list**. Common capabilities that do not get a standalone action here remain available in the wider guide:

| When the task needs… | Continue with… |
| --- | --- |
| Allowed inputs, dropdowns or rule-based highlighting | [Validation and conditional formatting](?family=rules&action=#families) |
| An inserted image or a small chart inside cells | [Images and sparklines](?family=drawings&action=#families); rendering a PNG is a separate job |
| Discussion threads or native cell notes | [Comments and notes](?family=comments&action=#families) |
| Help discovering supported calls and settings | [Inspection and discovery](?family=inspection&action=#families) |
| Merge, freeze or other sheet layout controls | [Cells and ranges](?family=ranges&action=#families) and [workbooks and sheets](?family=workbooks&action=#families) |

Sources: [package quick start](../2026-09-15-artifact-tool-reference/files/sources/package/API_QUICK_START.md), [skill quick start](../2026-09-15-artifact-tool-reference/files/sources/skill/artifact_tool_docs/API_QUICK_START.md), [preserved selection and examples](data/common-work.json), [source hashes and retained files](data/common-work-receipt.json).

### Read the labels correctly — An action, an API call and a workflow {#common-language}

An **action** is the job you want done, such as “write a formula.” An **API call** is how code asks the library to do it: `sheet.getRange("C5").formulas = [["=C3*C4"]]`. Here `getRange("C5")` obtains a reference to existing cell C5; it does not create another cell or insert content. The assignment writes the formula.

The full reference also uses **operation** for a specific library interface: a structured command passed to `workbook.apply(...)`. Direct API calls and structured operations are two ways of asking the library to act. The practical section predominantly uses direct calls. **Guided** labels an example assembled to teach a workflow; **session** describes a wrapper that keeps or resumes workbook state. Those words are not four equivalent kinds of Excel feature.

Adding a worksheet creates a named, initially empty tab. Writing its values or formulas is a separate step. A **range** is a reference to a rectangle of cells; shifting or growing that reference changes which cells later calls address. Copying formulas, inserting cells and moving values are separate changes. These distinctions make the smaller practical examples easier to connect to the full API.

### Learn from the checks — Import, missing inputs and saved files {#common-lessons}

**Imported formulas need a behavior check.** In the original Man Group example, some shared formulas arrived as saved values. Changing the fee-margin assumption from 56 to 51 basis points therefore did not update all dependent outputs. Repairing one formula did not repair the rest of the chain. Open **Read and calculate → Trace dependencies** and **Edit cells and sheets → Write formulas** to see the captured evidence. A dependency trace describes the imported representation; it cannot reveal a formula that the import did not retain.

**Blank and zero need different treatment when the model requires it.** The fee-lab examples show a blank rate producing zero until the agent writes an explicit missing-input guard. A deliberate zero remains a valid input. Open **Write formulas** for the guard and **Clear cell contents** for the blank and **Write or replace values** for the deliberate zero. This is a modelling decision implemented in formula code, not an automatic completeness check.

**Memory and the saved file are separate.** Open **Write or replace values → Edit without saving**: the in-memory fee margin changes while the saved file and its checksum stay unchanged. Export creates a file object; saving writes it to disk. The save call returns no verification report. Reopening the file and comparing selected cells are additional checks written by the agent.

The [saved Man Group copy with Fee lab](files/common-work/operations-run/man-group-operations.xlsx), [small new fee workbook](files/common-work/operations-run/new-fee-workbook.xlsx) and [independent file checks](files/common-work/operations-run/independent-verification.json) are retained with this guide. The original supplied file was left unchanged. Opening these outputs in native Microsoft Excel was not part of the recorded test.

## Follow one workbook through eight steps {#walkthrough} {toc=Walkthrough}

*See how source facts become formulas, a chart and a checked Excel file.*

Choose a quantity case and move through the stages. The highlighted cells show what changed. Switch to the object tree or raw data to connect the visible result to its structure.

::: wide
```component workbook-journey
renderer: scripts/explorer.py
data: data/field-guide.json
```
:::

### Read the Inspect step — Two calls and a recursive dependency tree {#journey-inspect}

That step makes two separate calls. **`inspect({kind: ...})` requests record types; `trace("Sales!D5")` follows the chosen result back to its inputs.** The older step's displayed input groups their settings for convenience; `trace` is not an option passed to that `inspect` call.

For D5's `=SUM(D2:D3)`, the trace's `params` array contains D2 and D3. D2's `=B2*C2` then has its own `params` containing B2 and C2. The nesting follows the calculation upstream. In this fixture, the typed input cells have `formula: null` and `params: []`: there is no further formula dependency to follow. It does not list every cell that would change downstream if an input changed.

**The old step retained inspection metadata and the trace, but omitted the inspection record payload during serialization.** Its `chart` kind was also ignored, as its metadata notice records. The separate workbook state shown in Readable is a saved cell projection, not that inspect response. Use the [new inspection comparisons](?common=map&inspect=map-workbook-150000#common-work) for complete captured NDJSON records, including the supported `drawing` route and the unknown-token example.

## Find the action that matches your intent {#families} {toc=Action families}

*The family explains why an object exists; its examples show the concrete input and output.*

A **facade** is an editable JavaScript object, such as a range or chart. An **operation** is a structured command that describes a change. Both can act on the same workbook, but their return values and supported behavior differ. Choose a family, then an example. Inputs, readable results, trees and code stay together.

::: wide
```component action-explorer
renderer: scripts/explorer.py
data: data/field-guide.json
```
:::

## Build a workbook people can understand and maintain {#best-practices} {toc=Best practices}

*The design principles, formatting choices and working methods behind the API calls.*

**Start with the reader’s question, give each calculation one home, preserve the source facts and check the result in the form you will deliver.** A compact task may fit on one sheet; a larger model needs a clear route from inputs through calculations to outputs. The aim is a workbook that remains understandable when someone changes an input, adds a period or returns next month.

This section synthesizes the **saved spreadsheet skill, its style and domain guides, the library quick start and feature references, and the supplied example scripts**. These are the documents captured for this study on 15 September 2026; “saved” describes the study copy, not an inactive skill. Read the complete skill documents in [Satellite documents](#satellite-documents). **Design defaults** guide judgment; **API contracts** describe how calls behave; **observed corrections** identify where our version 2.8.59 experiments qualify the documentation. A library example demonstrates syntax—it does not prescribe your tabs, colors or business model.

| Stage | The decision to make | What should exist before moving on |
| --- | --- | --- |
| Frame | What question, audience, refresh cycle and target application? | A clear output and the inputs it actually needs |
| Organize | Where do sources, assumptions, calculations and outputs belong? | A simple dependency flow; roles may share a sheet |
| Build | Which facts are typed inputs, and which results are formulas? | Traceable calculations with consistent periods and keys |
| Present | Which labels, formats, tables or charts help the reader act? | Readable sheets with clear units and editable inputs |
| Verify | Does the result stay correct after a meaningful change? | Independent checks, inspected formulas and readable renders |
| Deliver | Does the saved file preserve what matters? | The requested workbook, with tested behavior and limitations stated |

For a first build, start with [sheet structure](#practice-structure) and [the authoring workflow](#practice-workflows). For an existing file, start with [preservation](#practice-preservation). For presentation decisions, open [formatting](#practice-formatting) and [charts](#practice-charts). Each topic below explains **what to do, why it matters and how it looks in practice**.

### Follow the right design authority — User instructions, references and defaults {#practice-authority}

**Use this order: the user’s instructions → the supplied or selected reference → relevant domain conventions → general defaults.** A request to match an existing workbook outweighs a generic preference for different colors or tab order. Subject matter alone is not a template: a healthcare company’s valuation is a financial model, while its appointment register is an operational record.

For a new workbook without a reference or visual direction, the skill offers the template picker **when that capability is available**. A declined, unavailable or failed picker is a reason to continue with a sensible design. Template browsing alone does not authorize creating a workbook. Save an uploaded reference as a reusable template only when the user chooses that option.

When reconstructing a screenshot, retain real numeric and date values. Separate visible inputs from rows that clearly calculate sums, ratios or other repeated relationships. Match the visible design without treating compression artifacts or zoom as intentional typography. A static recreation of a calculated total will stop working as soon as the inputs change.

**Keep capability and recommendation separate.** A documented formula may be inappropriate for a simple schedule; an example’s three sheets do not mean your task needs three sheets. Likewise, the presence of a feature in help does not establish that it calculates or survives Excel export.

Sources: [spreadsheet skill](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [template selection](../2026-09-15-artifact-tool-reference/files/sources/skill/references/template-elicitation.md), [image references](../2026-09-15-artifact-tool-reference/files/sources/skill/references/image-references.md), [library quick start](../2026-09-15-artifact-tool-reference/files/sources/package/API_QUICK_START.md).

### Organize roles before adding tabs — Logical flow and reading order are different {#practice-structure}

**The logical flow is sources and assumptions → calculations → outputs. These roles are not mandatory worksheets.** A focused workbook answers one question or supports one workflow; its row count does not determine how many sheets it needs. Keep meaningful calculations visible instead of shrinking text, discarding useful rows or hiding workings to force a one-sheet result.

| Workbook shape | A sensible starting arrangement | Why it works |
| --- | --- | --- |
| Small sales calculation | One “Sales” sheet: title and result above; inputs and line calculations below | The question, editable facts and answer fit together |
| Analysis of an imported extract | “Campaigns” for the analysis; “Source” for the unchanged extract | The original record remains available while cleaning and calculations are separate |
| Recurring forecast | “Summary” → “Assumptions” → “Revenue build” → “Sources” | Readers see the answer and drivers first; formulas still flow from sources to builds to summary |
| Larger shared model | Optional Cover first; outputs, assumptions, builds and sources; optional Checks and ReadMe last | Navigation and documentation earn space when complexity or reuse requires them |

These are illustrative layouts, not required sheet names. For a new larger workbook, the skill places **outputs first, assumptions within easy reach, then builds and sources**. A useful Cover comes first; a justified Checks or ReadMe comes last. Preserve an existing workbook’s organization unless the task calls for changing it.

**Give every role a clear boundary.** Keep original extracts unchanged; put cleaning and mapping in prepared areas. Assumptions hold forecast drivers. Builds own meaningful steps such as headcount × compensation and opening balance + movements = closing balance. Outputs link to completed results. Do not duplicate a forecast on the Summary sheet or route finished output values back through Assumptions.

A Checks sheet is a terminal diagnostic area: **other sheets should not depend on it, even for a status or output gate**. Validation required to calculate correctly belongs with the owning input or build. Add a Cover or ReadMe only when it helps a complex, shared or recurring workflow; neither is an automatic requirement for a simple workbook.

Sources: [spreadsheet structure guidance](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [new-workbook workflow](../2026-09-15-artifact-tool-reference/files/sources/skill/workflows/create_workflows.md).

### Name things for the reader — Navigation, labels and provenance {#practice-naming}

**Names should explain purpose.** “Campaigns”, “Appointments”, “Revenue build” and “Forecast variance” tell the reader what to expect. Preserve established names in an existing file. The skill prefers “Forecast review”, “Forecast variance” and “Sensitivity” for those specific jobs; that is an editorial convention, not an API restriction on other words.

Use concise titles and business labels. Include the period, population and units where a reader needs them: “Active customers” must not quietly mean every customer ever recorded. A workflow that needs human input should expose an editable field and say what remains to be supplied. “Missing input: forecast rate” is more useful than a decorative green status badge or an unexplained error code.

Keep one authoritative case selector and link its selected-case label near the top of dependent sheets. Use consistent headers and a small, useful frozen area for long lists; preserve identifying columns when a schedule scrolls horizontally. Compact covers and dashboards may need no freezing at all.

**Cite sources where the facts enter the workbook.** Put a single source above its input block; place multiple sources adjacent to the relevant data, separated by a blank column. Avoid repeating citations on every output or adding a source tab solely to house links. When a separate source area is unsuitable, the skill permits a native note on a hardcoded input. Ordinary explanatory text in a “Notes” column is different from a native note or threaded comment; preserve the distinction and use discussion threads when requested.

Sources: [spreadsheet skill](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [shared style](../2026-09-15-artifact-tool-reference/files/sources/skill/style_guidelines.md), [comments, notes and names API](../2026-09-15-artifact-tool-reference/files/sources/package/api/references/comments-notes-names.spec.md).

### Preserve facts and types — A display format does not repair the data {#practice-inputs}

**Keep raw records intact and transform them in a prepared area.** Retain stable identifiers, original categories, units and the source’s level of detail—for example, one row per transaction rather than an unexplained mixture of transactions and monthly totals. Derived cleaning columns should make the transformation reviewable.

| Intended meaning | Store | Display | Common mistake |
| --- | --- | --- | --- |
| A 12.5% rate | Number `0.125` | `0.0%` → 12.5% | Text “12.5%”, or number 12.5 formatted as a percentage |
| An amount | Number `1500000` | A currency format, or an explicitly scaled presentation | Rounding or dividing the stored value merely to shorten its display |
| A sortable date | A real date value | An explicit date format | Text that looks like a date but sorts alphabetically |
| A postal code or SKU | Text such as “00123” | Preserve its characters | Losing leading zeros by treating an identifier as a quantity |
| An unknown measurement | A defined missing value | Blank or an explicit missing label, as appropriate | Substituting zero and making the result look complete |

Prefer the library’s CSV importer to handwritten CSV splitting. **CSV fields import as strings**, so convert intended number and date columns before calculations; formatting them does not change their type. Preserve true identifiers as text.

For range writes, use row-by-column matrices with dimensions matching the target. Block writes are easier to audit and more efficient than repeated cell calls. If expanding deliberately from an anchor, use the documented single-cell or `write(matrix)` route. A literal label beginning with `=` needs a leading apostrophe so it does not become a formula. Create all sheets that formulas will reference before writing those formulas.

Sources: [library quick start](../2026-09-15-artifact-tool-reference/files/sources/package/API_QUICK_START.md), [range API](../2026-09-15-artifact-tool-reference/files/sources/package/api/references/ranges.spec.md), [spreadsheet skill](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md).

### Make calculations easy to follow — One owner, useful steps and deliberate anchors {#practice-formulas}

**Keep business logic in formulas and calculate reusable results once.** A reader should be able to trace revenue from units and price, rather than finding a hardcoded result or several slightly different copies of the same calculation. Separate useful business steps; do not create a helper cell for every trivial expression.

| Formula pattern | What it means | What to check when copying |
| --- | --- | --- |
| `=C8*C$3` | A row value times the driver for this period | The driver row stays fixed; the period column moves |
| `=C8*$B$3` | A row value times one global driver | Both driver coordinates stay fixed |
| `=SUMIFS(Amount,Month,C$4,Item,$A8)` | Sum records matching this period and item | Period header row and item label column remain anchored |
| `='Revenue build'!E14` | Read a finished result from its owning schedule | The output links to the correct period and metric |

The named ranges in these examples are illustrative and must be defined over aligned records. Use keyed lookups when source and destination row orders differ: anchoring a cell does not align two different item lists. Choose exact versus approximate matching deliberately, and decide whether duplicate keys should be rejected, aggregated or selected according to a defined rule.

The skill favors `SUM` for totals, direct multiplication for two factors, `PRODUCT` for several factors and `SUMPRODUCT` for aligned weighted calculations. `PRODUCT` ignores blanks and text, so validate required inputs. It favors `SUMIFS`, `COUNTIFS` and `AVERAGEIFS` for new conditional calculations; an existing valid single-condition formula does not need cosmetic replacement. For many editable rules, a mapping table can be clearer than a long nested `IF`, provided it preserves ordering, boundaries and gaps.

**Use the simplest supported formula that preserves the meaning.** The skill does not recommend introducing `LET`, array/spill formulas, `MAP`, `REDUCE` or `LAMBDA` into ordinary new work. That is a design default, not an instruction to remove a user-required or existing formula. If a spill formula is appropriate and verified, the library says to write only its anchor cell. Scalar formulas can be seeded and filled down or right; a one-cell formula matrix does not broadcast like a scalar value assignment.

**Observed correction:** the skill mentions `OFFSET` and `INDIRECT`, but both were unimplemented in this version’s independent checks. Use a verified equivalent such as direct references or an appropriate `CHOOSE` design. A help entry is not a correctness test. The quick start also documents compatibility prefixes for some functions and warns about blank criteria in `COUNTIF`/`COUNTIFS`; our blank-count probe returned 0 where 5 was expected. For an unconditional blank count, it recommends `COUNTBLANK`.

Sources: [formula construction guidance](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [quick-start formula limits](../2026-09-15-artifact-tool-reference/files/sources/package/API_QUICK_START.md), [recorded formula checks](../2026-09-15-artifact-tool-reference/README.md). Explore the exact receipts in [Catalogues](#catalogues) and [Boundaries](#boundaries).

### Treat time as part of the calculation — Dates, periods and aggregation {#practice-periods}

**Use the task’s as-of date and calendar.** Use `TODAY()` only when a live clock is intended. Preserve supplied years and fiscal definitions; do not invent a missing year. Label actual and forecast periods explicitly, and keep historical facts unchanged when switching forecast cases.

Within each time scale, run periods chronologically. If several scales are useful, the skill orders broader periods on the left and finer periods on the right: annual, quarterly, monthly, weekly. Separate scale blocks with an unfilled spacer, but keep a continuous monthly timeline together across its actual/forecast boundary. Use only the scales the task needs. A week-of label normally starts on Monday; retain a supplied week-ending convention and use a holiday calendar only when one is provided.

**Different quantities need different aggregation.** Sum flows such as revenue; use the closing observation for balances; rebuild ratios from their numerator and denominator. Adding monthly customer balances or averaging conversion percentages without their populations changes the meaning. Do not count an entire cross-month week in both months.

For daily records feeding a monthly total, inclusive month start and exclusive next-month start also handle timestamps on the last day:

```text
=SUMIFS(Amount,Date,">="&C$4,Date,"<"&EDATE(C$4,1),Item,$A8)
```

Here `C4` contains the real first date of the month, and all named ranges cover the same records. Use bounded ranges with equal extents. The library warns that whole-column references may be truncated to populated extents, creating mismatched ranges or surprising row counts. Test first, middle and later periods, plus the next period if extension is promised.

Sources: [dates and periods in the spreadsheet skill](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [bounded-range caveats](../2026-09-15-artifact-tool-reference/files/sources/package/API_QUICK_START.md).

### Build cases around one selector — Live assumptions, snapshots and sensitivity {#practice-cases}

**One selector should drive one active forecast.** Group assumptions by driver, with the active selection first and aligned case rows beneath it. A numeric selector with validation can choose among Base and Downside; dependent builds read the active assumption, and each sheet displays the linked selected-case name.

```text
Active assumption in I21: =CHOOSE($B$3,I22,I23)
Build's driver cell:       ='Assumptions'!I21
```

In this illustration, `B3` is 1 for Base and 2 for Downside. Historical actuals remain fixed. A missing value in the selected case must be visible; a missing value in an unused case should not automatically block the active forecast. Forecast calculations must not feed back into the assumptions that drive them.

| Need | Appropriate pattern | Essential safeguard |
| --- | --- | --- |
| Work with a different forecast case | Change the authoritative selector and calculate the same build | All linked outputs and case labels reflect the selected case |
| Compare saved results from several cases | Set inputs, calculate, capture labeled results, restore inputs in `finally` | Record which assumptions were captured and how to refresh; two labels must not both point at the current live result |
| Native Excel what-if sensitivity | Use a native data table when explicitly requested or required by the template | Include its corner, headers and body; verify input mapping and native export |

A captured case is a snapshot, not a second live model. Circular formulas that retain their own previous value require an explicit design and verified iteration, convergence and preservation; the skill does not offer them as a default comparison method.

**Observed correction:** the supplied data-table guidance describes two-variable support, but the study successfully exercised one-variable tables in both orientations. Use the recorded examples for the tested range shapes, and keep native Excel behavior marked unverified. Do not silently substitute a static grid when the task requires a native feature.

Sources: [scenario guidance](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [supplied data-table guide](../2026-09-15-artifact-tool-reference/files/sources/skill/artifact_tool_docs/DATA_TABLES.md), [data-table contract](../2026-09-15-artifact-tool-reference/files/sources/package/references/data-tables.spec.md), [scenario and sensitivity experiments](../2026-09-15-artifact-tool-reference/README.md).

### Keep missing information visible — Zero, errors and overrides have different meanings {#practice-missing}

**Zero means a measured or deliberately chosen zero.** Missing, not applicable, unavailable and failed calculations are different states. Decide how each should appear before introducing fallback formulas. A convenient-looking zero can materially misstate a total or chart.

Use `IFERROR` only when the fallback is understood. Blanket `IFERROR(...,0)` can conceal a broken reference. `IFNA` can isolate a missing lookup key, but finding the key still does not prove its value is populated: a reference to an empty cell may become zero. Similarly, `SUMIFS` returning zero does not prove matching records existed. Add a coverage check when the difference matters.

If the task requires an optional override, preserve the base calculation and treat a blank override differently from an explicit zero. With a validated numeric override in `C8` and the base result in `B8`:

```text
=IF(C8="",B8,C8)
```

A blank uses the base; zero overrides it with zero. Check invalid types and preserve the intended rounding. Do not add override machinery when the workflow does not need it.

**Warnings should explain a real condition and the next action.** Keep required-input checks near the owning input or calculation. Completing one prerequisite must not hide another unfinished prerequisite. Data validation improves the editing interface, but our API write of 999 was not blocked by a restrictive rule: validate code-supplied inputs in the surrounding script too.

Sources: [missing-data, error and override guidance](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [validation contract](../2026-09-15-artifact-tool-reference/files/sources/package/references/data-validations.spec.md), [validation experiment](../2026-09-15-artifact-tool-reference/README.md).

### Format for reading and editing — Hierarchy, space and typed number formats {#practice-formatting}

**Use formatting to distinguish roles without overwhelming the numbers.** Follow a supplied workbook’s style. For a new workbook, the shared defaults favor a restrained palette, dark readable text, useful whitespace and light structural borders.

| Element | Documented starting point | Why |
| --- | --- | --- |
| Sheet opening | A blank top row, concise left-aligned title and thin rule | Orients the reader without a large banner; do not insert presentation rows into a raw extract |
| Working grids | Unmerged headers and cells; hide gridlines for a new presentation sheet | Preserves sorting, editing and copying while reducing visual noise |
| Borders and fill | Dark headers with white text; light separators; continuous section bands | Establishes hierarchy without boxing every body cell; blank gutters remain unfilled |
| Alignment | Text left, numbers right, headers centered | Makes labels readable and values comparable; top-align long wrapped descriptions |
| Type | One verified font family; usually two sizes, at most three | Cells, theme and charts should agree; the largest size stays within six points of the body size |
| Row and column sizing | Consistent row heights, useful padding, local width adjustments | Fix clipping and `####` without shrinking the font or auto-fitting an entire existing sheet |
| Inputs and warnings | Restrained input highlighting; warnings visibly distinct with a legend where needed | Editable cells and calculated problems must not be confused |
| Notes | A concise notes column after a blank spacer, or a separate area for long explanations | Long commentary should not inflate every row of a numeric table |

The documented font fallback is Helvetica Neue → Helvetica → Arial → Aptos, using a family actually available to the renderer and target application. A 10-point body and 14-point title illustrate the scale; they are not universal fixed settings. Freeze only the headers and identifiers needed for the sheet’s scrolling behavior.

Use explicit Excel number formats. `#,##0` suits whole counts; `0.0%` suits many analytical rates. Currency usually needs whole units unless cents matter; per-share figures and small percentage differences may justify more precision. Keep the underlying precision intact. Format codes use Excel’s invariant conventions; do not substitute locale punctuation into the code to imitate the screen.

For larger models, tab colors can distinguish outputs and assumptions; grouped source tabs and divider tabs may help a genuinely large workbook. Skip that machinery for compact work. The skill discourages unnecessary oversized KPI cards, merged working grids, decorative status badges and inline bars unless requested. User or template requirements still come first.

Sources: [shared style guidelines](../2026-09-15-artifact-tool-reference/files/sources/skill/style_guidelines.md), [number-format guidance](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [formatting API](../2026-09-15-artifact-tool-reference/files/sources/package/api/references/formatting.spec.md), [pixel and Excel sizing](../2026-09-15-artifact-tool-reference/files/sources/package/references/styles.spec.md).

### Give every chart a question — Selection, binding, placement and export {#practice-charts}

**A chart earns its place by answering a distinct question.** Keep exact values in a table when precision matters, and use a native editable chart when the workbook requires one. A preview image alone does not establish an editable Excel chart.

| Reader’s question | Usual starting point | What to preserve |
| --- | --- | --- |
| Which categories are largest? | Sorted bar or column chart | Meaningful categories, consistent units and a zero baseline |
| How did it change over time? | Chronological line chart | Actual dates, honest gaps and clearly explained axis bounds |
| How is a quantity distributed? | Histogram or box plot when supported; otherwise a disclosed, formula-backed summary | Bin definitions, sample coverage and missing observations |
| Are two variables related? | Scatter chart | Aligned observations and the meaning of each axis |
| What makes up the whole? | Bar or table; a pie/doughnut only for a few rough shares | The denominator and completeness of the whole |

**Bind series to worksheet cells.** Categories and values must align; derived chart helpers should contain formulas, so changing an input updates the chart’s source. Use helpers when reshaping, grouping dates, shortening labels or handling a verified rendering issue—not as a duplicate hardcoded copy of a working table. With nonadjacent ranges, use aligned one-column ranges of equal height. Check whether the first row is really a header.

Place the chart near related data in a reserved blank area. Set its position and dimensions explicitly; check labels, legends and neighboring tables at normal zoom. Keep business colors consistent across comparable charts, use a single color for a single series, and avoid 3D effects. Bars and areas normally start at zero; a restricted or logarithmic line/scatter axis needs a clear reason. Show missing observations as missing rather than false zeros, and disclose omitted or aggregated observations.

**Axis formatting is separate from cell formatting.** If source values are dollars, `$0.0,,"M"` can display 1,500,000 as $1.5M. If source values already represent millions, use `$0.0"M"` instead. Label the unit and avoid scaling twice. The quick start uses `numberFormatSourceLinked: false` for custom axis formats; verify the rendered ticks. Charts have separate text styles and use pixel sizing, while cell text sizes use points.

For a line series, the chart guidance explicitly sets the stroke:

```text
series.line = { fill: color, style: "solid", width: 2 };
```

Setting a fill alone may not preserve the desired line in Excel. Bind data before applying series styles: `setData` replaces series while preserving chart-level titles and axes. Inspect each series’ value and category formulas; `chart.categories` may be empty even when categories are correctly bound to cells.

**Observed limit:** 16 of 25 chart tokens produced native Excel charts in this study. Nine were omitted, including histogram, box-and-whisker, waterfall, combo and pareto. Use the chart catalogue’s per-type receipt, not a blanket “supported” claim. A formula-backed bin table and bar chart can answer a distribution question, but disclose the change if a native histogram was requested. Some nested chart settings also failed export when replaced with plain objects; use the corrected facade-field examples in the action explorer.

Sources: [chart design guidance](../2026-09-15-artifact-tool-reference/files/sources/skill/features/charts.md), [quick-start chart details](../2026-09-15-artifact-tool-reference/files/sources/package/API_QUICK_START.md), [chart suggestions example](../2026-09-15-artifact-tool-reference/files/sources/package/examples/chart_suggestions.ts), [chart export experiments](../2026-09-15-artifact-tool-reference/README.md).

### Choose features for their job — Tables, validation, names and annotations {#practice-features}

**A table, pivot and what-if data table solve different problems.** A native table organizes a rectangular list with headers, filtering and structured references. A pivot summarizes records. A what-if data table evaluates a formula against different inputs. Pick the object whose behavior the workflow needs.

| Feature | Useful when | Construction and verification habit |
| --- | --- | --- |
| Native table | A record list needs structured columns and filters | Give it a unique explicit name, preserve headers and check its range for overlap |
| Defined name | A reused assumption or range becomes clearer with a stable name | Check its scope and reference after insertions or edits; naming does not validate the referenced data |
| Data validation | People must choose from defined options or enter valid values | Set meaningful prompts and rules; separately validate programmatic writes |
| Conditional formatting | A real threshold, exception or comparison should be visible | Use the correct range, rule order and business threshold; keep missing values distinguishable |
| Notes or comments | Provenance or requested collaboration belongs with cells | Keep notes distinct from discussion threads; set the visible author before creating a thread |
| Sparklines | A compact trend is genuinely useful | Verify native settings; the quick start warns that date-axis ranges are not serialized and manual bounds need care |
| Images and shapes | A reference or explanation benefits from a drawing | Reserve space and use the documented anchor/creation methods; inspect export preservation |

**Observed corrections:** turning on a table’s totals row reclassified its last data row in our probe; it did not append a fresh row. Reserve or append the required structure and check formulas after a fresh calculation. Pivot and slicer preservation also varied on re-export. Formula summaries can be a practical alternative when the requested task permits them; they do not fulfill an explicit requirement for a native pivot or slicer.

The image probe likewise showed that replacing an anchor with a plain object could break export. The broader lesson is to follow the actual object contract rather than assuming every nested facade can be replaced by JSON. Use the family examples for tested creation and update shapes.

Sources: [library quick start](../2026-09-15-artifact-tool-reference/files/sources/package/API_QUICK_START.md), [tables API](../2026-09-15-artifact-tool-reference/files/sources/package/api/references/tables.spec.md), [defined names](../2026-09-15-artifact-tool-reference/files/sources/package/references/defined-names.spec.md), [conditional formatting](../2026-09-15-artifact-tool-reference/files/sources/package/references/conditional-formatting.spec.md), [feature experiments](../2026-09-15-artifact-tool-reference/README.md).

### Pick the workflow before writing code — Create, edit, answer or persist {#practice-workflows}

**The surrounding script should make the work repeatable and its effects clear.** The spreadsheet skill uses the bundled JavaScript package through the configured dependency runtime and keeps one patchable, rerunnable `.mjs` builder in a writable working directory. Reuse the installed dependencies; do not modify the package to make an example run.

| Workflow | Suggested sequence | Result to retain |
| --- | --- | --- |
| New workbook | Plan → create referenced sheets → block-write inputs → formulas → style and charts → calculate → inspect → render → export | One builder and the requested workbook, with checks tied to its outputs |
| Existing workbook | Import → summarize and locate → inspect formulas and render → make focused edits → check affected dependencies → export and compare | The revised file with its required structure and native features preserved |
| Read-only question | Locate labels and period → inspect value and formula → trace back to inputs → independently reconcile → explain | An answer with its basis; no unintended edit or export |
| Repeated case capture | Save original inputs → apply case → calculate → collect labeled results → restore in `finally` | A snapshot with captured assumptions and a refresh method |
| Persistent session | Create/resume session → targeted reads or mutations → inspect state → render/export | Workbook state across calls and the returned saved-file descriptors |

The skill also supplies an **operation marker** for its authoring workflow: run it once immediately before the first create/edit, using the documented operation, count and format arguments. Skip it for a read-only question. It records the start of work; it does not create or export a workbook. This is workflow support around the library, not an Excel feature.

Use `workbook.record(...)` when the caller needs the returned patch, object-ID map or collaborative update, rather than wrapping every action without a purpose. For a long build, checkpoint exports can isolate which block introduced a serialization problem. After a failure, read the error, look up the exact feature, patch the smallest relevant section and continue from known state. The quick start recommends a bounded help lookup and, if necessary, one reformulation; its examples are alternatives, not a checklist to execute for every task.

**Persistent-session caution from the study:** managed render/export calls return file descriptors, unlike the ordinary in-memory Blob-like results. A failed callback was rolled back in the tested case, but a timeout can leave the mutation outcome uncertain. Inspect the resumed state before retrying. The [Changes & sessions family](#families) and Code & guidance catalogue expose the actual wrappers and receipts.

Sources: [spreadsheet authoring workflow](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [create workflow](../2026-09-15-artifact-tool-reference/files/sources/skill/workflows/create_workflows.md), [edit workflow](../2026-09-15-artifact-tool-reference/files/sources/skill/workflows/edit_workflows.md), [read-only guidance](../2026-09-15-artifact-tool-reference/files/sources/skill/references/read_only_qna.md), [recorded edits API](../2026-09-15-artifact-tool-reference/files/sources/package/api/API_DOCS.md), [marker script](../2026-09-15-artifact-tool-reference/files/sources/skill/container_tools/mark_artifact_operation_started.mjs).

### Borrow the examples’ methods — Inspect narrowly, trace causes and style after binding {#practice-examples}

**The supplied scripts teach working patterns as well as calls.** Their sample filenames, local import paths and fixture labels need adapting to the installed runtime and the actual task. Their complete captured code and execution receipts are available under Code & guidance in [Catalogues](#catalogues).

| Supplied example | Reusable method | What not to infer from it |
| --- | --- | --- |
| Quick start | Build rectangular inputs and formulas, apply presentation, render and export | Its particular layout is not a universal workbook template |
| Inspect existing workbooks | Start with a bounded workbook summary, then sheets, a relevant region, formulas and drawings | Dumping every cell is not necessary to understand a workbook; choose the intended input file explicitly |
| Formula trace and help | Follow an output through its precedents, cap the printed tree, compare with an independent expectation and tolerance | A sample Checks tab or “OK” label is not required in every deliverable |
| Chart suggestions | Choose a chart from the data question, bind to cells, then style and position it | A suggested chart type or good preview is not proof of native export support |

For example, the trace script compares a compounded result close to **121** with an independently entered expectation of **121**, allowing a **0.001** numerical tolerance. It explains the tiny floating-point difference through the input revenue and growth rate. The lesson is to tie tolerance to the quantity being checked and expose the chain of causes, not to copy 0.001 into every financial or scientific test.

For broad discovery, use bounded `inspect` results. Once the target rectangle is known, direct formula getters return its exact matrix. A dependency trace can grow large: cap the displayed depth and node count, and explicitly say when the view is truncated. Keep full underlying evidence available when the truncated portion could affect the answer.

Sources: [quick-start example](../2026-09-15-artifact-tool-reference/files/sources/package/examples/quick_start_example.ts), [inspection example](../2026-09-15-artifact-tool-reference/files/sources/package/examples/inspect_existing_workbooks.ts), [trace and help example](../2026-09-15-artifact-tool-reference/files/sources/package/examples/formula_trace_and_help.ts), [chart suggestions example](../2026-09-15-artifact-tool-reference/files/sources/package/examples/chart_suggestions.ts).

### Preserve an existing workbook deliberately — Scope edits and check what must survive {#practice-preservation}

**Inspect before changing.** Use a compact sheet/table/formula/drawing inventory to locate the actual objects, then read exact target formulas and render the relevant view. Treat the existing file as the design reference. Values-only work should not quietly become a redesign.

Preserve sheet names and order, formulas, styles, merged regions, validation, conditional formatting, defined names, protection, hidden rows and columns, freeze panes, links, annotations and calculation settings where applicable. Edit the requested cells and affected dependencies. When adding records or periods, extend any chart, formula, table, validation or conditional-formatting range that is supposed to include them.

Avoid sheet-wide auto-fit or broad formatting resets for a narrow edit. Notes and overrides must remain attached to the correct stable record after sorting or refreshing; position alone may not be a safe identity. Report unrelated pre-existing errors rather than silently repairing them beyond the task.

**Separate engine preservation from application preservation.** Compare relevant sheet and object inventories before and after; inspect the changed render and dependent views. Reopening an exported file in Artifact Tool is useful, and examining its XML can verify specific native objects. Neither establishes full Microsoft Excel behavior. If a required native feature cannot be preserved, explain the specific gap before treating a substitute as equivalent.

Sources: [edit workflow](../2026-09-15-artifact-tool-reference/files/sources/skill/workflows/edit_workflows.md), [inspection example](../2026-09-15-artifact-tool-reference/files/sources/package/examples/inspect_existing_workbooks.ts), [import/export reference](../2026-09-15-artifact-tool-reference/files/sources/package/api/references/import-export.md), [round-trip experiments](../2026-09-15-artifact-tool-reference/README.md).

### Verify meaning, behavior and delivery — Three different kinds of evidence {#practice-verification}

**A correct-looking cell, a readable render and a saved XLSX answer different questions.** Check all three in proportion to the work. Formulas calculate automatically, but the documented workflow explicitly recalculates after edits before final checks and export; calculate again if a subsequent change affects the result.

| Check | Concrete test | Failure it can catch |
| --- | --- | --- |
| Meaning | Independently recompute headline metrics from their underlying inputs, using the right units, period and population | A plausible formula that measures the wrong thing |
| Dependencies | Trace key outputs to their labeled assumptions; examine copied formulas at first, middle and later rows/periods | Mis-anchored references, wrong cases and hidden duplicate logic |
| Behavior | Change representative inputs; test blank, zero, duplicate keys, cutoff dates and overrides; restore temporary changes | Stale summaries, swallowed errors or an extension that does not work |
| Completeness | Scan standard Excel errors plus textual exceptions and unimplemented results | Failures that an error-token search alone misses |
| Presentation | Render every sheet for a new broad build; inspect changed and affected views for a narrow edit | Clipping, `####`, unreadable labels, chart overlap and misleading colors |
| Saved artifact | Export, reopen, and check relevant native objects and formulas; test the intended application when available | Export omissions, stale values or round-trip losses |

Use exact comparisons for identifiers, categories and counts. For calculated numbers, choose an absolute or relative tolerance appropriate to their units and scale. An independent check should not simply repeat the formula under test. If the workbook promises another record, period or case can be added, exercise that extension.

Inspect at normal zoom with the cells unselected: selection can hide the intended fill or contrast. Check wrapped labels, row heights, freeze panes, chart axes and legends. Fix the relevant issue and recheck the affected view; repeated renders of unchanged areas do not add evidence.

Deliver the requested artifact rather than several unexplained “final” variants. State material limits specifically: in this study, native Microsoft Excel behavior and account-backed Google Sheets remain untested. A number-producing function or zero-error XML parse is not a general compatibility certificate.

Sources: [verification guidance](../2026-09-15-artifact-tool-reference/files/sources/skill/SKILL.md), [trace example](../2026-09-15-artifact-tool-reference/files/sources/package/examples/formula_trace_and_help.ts), [chart checks](../2026-09-15-artifact-tool-reference/files/sources/skill/features/charts.md), [recorded boundary cases](#boundaries).

### Apply domain conventions where they belong — Finance, marketing, healthcare and science {#practice-domains}

**The workbook’s function determines which conventions apply.** Keep the common discipline—typed facts, traceable calculations and readable outputs—then add only the domain rules relevant to the task.

| Domain | Core calculation discipline | Presentation and workflow implications |
| --- | --- | --- |
| Financial models | Preserve accounting basis, currency, signs and fiscal periods. Separate revenue from collections, expense from cash and capital expenditure from depreciation. Reconcile opening balances, movements and closing balances without hidden plugs. | Working cells may use blue hardcodes, black same-sheet formulas, green cross-sheet links and red external links; output presentation stays dark. Use accounting negatives, meaningful zero/dash formats and appropriate per-share/FTE precision. |
| Marketing and advertising | Keep the original extract, observation period and population. Aggregate the numerator and denominator before calculating a combined rate. Use actual supplied targets. | Label campaign/channel metrics and definitions clearly; do not invent performance scores or a dashboard full of unsupported flags. |
| Healthcare records | Preserve source records, measurement units, codes and whether rows describe patients, encounters or another unit. Apply only supplied protocols and thresholds. | Make workflow fields and unresolved prerequisites explicit. Do not invent universal normal ranges or clinical conclusions from formatting. |
| Scientific research | Keep raw observations flat and typed: one observation per row, one variable per column. Separate processing, retain units and document the relevant protocol. | Use explicit flag columns rather than color or comments alone. Preserve reproducibility for required simulations; include the requested uncertainty and checks without inventing extra analyses. |

In finance, keep historical actuals separate from scenario-driven forecast assumptions. Paid full-time equivalents are not automatically headcount; an already monthly pay figure must not be divided by twelve again. Reconcile variance amounts and percentages to the correct base, explain favorable versus unfavorable signs, and show an unavailable percentage when its denominator makes the comparison undefined. A comparables task does not by itself require a discounted-cash-flow or leveraged-buyout model.

The financial style guide uses parenthesized negatives, often one decimal for rates and two for per-share values; it preserves useful fractional staffing precision. Its month and fiscal-period labels must reflect the source calendar. Financial check differences are neutral around zero and become visibly exceptional beyond an appropriate tolerance; missing or untested inputs do not count as a passed check. These conventions should not automatically color or reshape a marketing, appointment or laboratory workbook.

Sources: [financial models](../2026-09-15-artifact-tool-reference/files/sources/skill/domain_guidance/financial_models.md), [marketing and advertising](../2026-09-15-artifact-tool-reference/files/sources/skill/domain_guidance/marketing_advertising.md), [healthcare](../2026-09-15-artifact-tool-reference/files/sources/skill/domain_guidance/healthcare.md), [scientific research](../2026-09-15-artifact-tool-reference/files/sources/skill/domain_guidance/scientific_research.md).

## Read the skill and its satellite documents {#satellite-documents} {toc=Satellite documents}

*The original instructions behind the workbook: what each document covers, its size and its complete text.*

**The main spreadsheet skill is the entry point. Its satellite documents provide the detail for particular tasks.** Start with **Style and layout**, **Finance guidance** and the **Creating** or **Editing** workflow. The task-specific guides explain charts, screenshots, templates and read-only questions. Technical references cover the API and specialist features; the three other domain guides are included for reference.

Every document below shows its **token count before you open it**. Select a name to read the complete Markdown as headings, paragraphs, lists, tables and code. The reader also shows word and character counts, a contents list, the original Markdown and a download. Links between these documents keep you in the reader.

These are the study's saved source documents, captured on **15 September 2026**. They were checked against the installed skill on **16 September 2026**. Supporting documents are read when applicable; the combined size does not mean the entire collection is loaded on every task. Template bundles and the library's wider API reference are separate resources, beyond this skill-document collection.

::: wide
```component satellite-documents-browser
renderer: scripts/satellites.py
data: data/satellite-documents.json
```
:::

## Browse the complete captured catalogues {#catalogues} {toc=Catalogues}

*Search every supplied API example and formula, plus charts, shapes, scripts and workflow guidance.*

“Executed” means a call ran. It does not certify every option or prove the result correct. Formula examples use a common numeric fixture; independent expected-result checks, where available, are shown separately. Chart previews and native Excel preservation are reported as different outcomes. **Code & guidance** contains the four executed package examples, the operation marker, eleven supplied guidance documents and three study harnesses; their origins and execution status are labeled.

::: wide
```component catalogue-browser
renderer: scripts/explorer.py
data: data/field-guide.json
```
:::

## Check the boundaries that can change the answer {#boundaries} {toc=Boundaries}

*A successful call, a plausible number and a saved file prove different things.*

These cases explain why the surrounding code matters. A wrapper can validate inputs, retain a checkpoint, inspect warnings, compare independent expectations and reopen an export. The useful workflow is **edit → calculate → inspect → render → export → reopen**, with checks at the boundary that matters.

::: wide
```component boundary-lab
renderer: scripts/explorer.py
data: data/field-guide.json
```
:::

## Keep the working pattern and its evidence together {#working-pattern} {toc=Working pattern}

**Use the design guidance and the execution evidence together.** [Best practices](#best-practices) explains how to organize and construct a workbook; the families and catalogues show the actual call shapes, settings and recorded results. For a new workbook, keep one rerunnable builder. For an existing workbook, inspect first and make focused edits.

For persistent work, a session stores workbook state between calls. Its render/export methods return saved-file descriptors; ordinary in-memory calls return Blob-like data. A timed-out mutation can have an uncertain outcome, so reopen and inspect before retrying. The Changes & sessions family contains the actual session inputs and returned objects. In Code & guidance, inspect the scenario, session and isolated-worker scripts to see how these calls become repeatable workflows. Guidance records cover creation, focused edits, templates, style and domain-specific checks; their complete archived text sits behind the code door.

#### Evidence and scope — What was executed and what remains unverified {#evidence-scope}

The preceding research exercised 39 operation commands through 42 examples, 126 API examples, 494 formula examples, 25 chart types and 190 shape types. A separate 77-case formula check matched 72 independent expectations. The native-file audit parsed 195 generated workbooks, with 15 selected import/re-export comparisons. This page adds three freshly captured workbook cases, each checked across eight stages.

The original [complete Markdown reference](../2026-09-15-artifact-tool-reference/README.md) holds the full settings, supplied contracts, corrected usage and experiment ledger. This page's [data manifest](data/MANIFEST.md), [source receipt](data/source-receipt.json) and [walkthrough data](data/walkthrough.json) explain the frozen inputs and how to reproduce them. The source documents and runnable probes remain with the reference.

Native Microsoft Excel behavior and account-backed Google Sheets were not tested. Rendering and XML inspection do not establish complete application compatibility. The workbook grids in this guide are readable projections of saved cells; the captured PNGs show the package's own rendering.

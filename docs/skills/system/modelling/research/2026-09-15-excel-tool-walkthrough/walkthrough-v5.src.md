---
title: Codex and Excel, with Man Group
class: system-doc
docver: 5
primary_author: RG
agent_assistance:
  - Codex
subtitle: What each operation returns, and how its settings change the answer
history:
  - v5 2026-09-15 — Operation map and measured inspection-settings comparison
  - v4 2026-09-15 — Definitions beside the calls, including table versus worksheet
  - v3 2026-09-15 — Readable inputs alongside exact code and arguments
  - v2 2026-09-15 — Man Group rerun and interactive response explorer
  - v1 2026-09-15 — Literal input/output walkthrough
stub:
  - Follow one workbook from discovery through a fee-margin edit and export, with the actual inputs and outputs at each step.
---

**Codex chooses what to ask the workbook, runs code to get it, then interprets the result.** A tab list, a block preview and a text search can all come from the **same function, `inspect`, with different settings**. Exact cell reads, formula tracing, rendering and edits are other operations.

The examples use your Man Group file. Start with the [settings comparison](#inspection-settings): every choice shows its actual input and recorded output. The [individual call examples](#calls) retain the detailed evidence. Nothing on this page runs live code or changes the model.

## Which operation gives you what? {#layers}

```component call-flow
renderer: man_group_walkthrough.py
data: evidence/man-group-cards.json
---
flow
```

The outer **command tool** starts a JavaScript program using Node.js. Inside that program, **Artifact Tool** (`@oai/artifact-tool`) supplies the functions below. They operate on the imported file in memory. These are library calls inside code, rather than separate Excel MCP tools or actions in an Excel window. MCP is a protocol for exposing tools to an agent.

| What the agent wants | Operation inside the script | What it gets |
| --- | --- | --- |
| Find its way around | `wb.inspect(settings)` | Depending on settings: sheet descriptions, cell values, block previews or search matches. Compare them below. |
| Read specified cells exactly | `range.values` and `range.formulas` | Two aligned grids: current values and formula text. [Example](#xio-ranges). |
| Follow a calculation | `wb.trace(cell)` | A tree of the result and its upstream inputs. [Example](#xio-trace). |
| See the layout | `wb.render(settings)` | An image of selected cells. A separate image-view tool displays it to Codex. [Example](#xio-render). |
| Change cells or copy formulas | Set `.values` / `.formulas`; use `fillRight` / `copyFrom` | A changed workbook object. The script must read cells again to show what changed. [Example](#xio-value-edit). |
| Refresh calculations | `wb.recalculate()` | Re-evaluated formulas in the library’s calculation engine. A separate read checks the results. [Example](#xio-recalculate). |
| Produce the Excel file | `SpreadsheetFile.exportXlsx(wb)`, then `.save(...)` | A saved `.xlsx` file. [Example](#xio-export). |

A **workbook** is the complete file. A **worksheet**, or **sheet**, is the grid opened by a **tab**, such as Summary. A **cell** is one position, such as W26. A **range** selects cells: K13:M20 means the rectangle from K13 to M20, inclusive.

#### Import and recording — the setup behind every example

The original file was copied unchanged from Downloads. The script imports that copy into a workbook object named `wb`:

```text
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const wb = await SpreadsheetFile.importXlsx(
  await FileBlob.load(path.join(dir, 'source.xlsx'))
);
```

Each example’s JavaScript runs inside our `capture()` helper — small code we wrote to save the evidence. It saves the returned object, input code and elapsed time. For inspection results it reads `ndjson` (one JSON record per line), `recordCount` (number of returned items), `truncated` (whether the response was cut short) and `metadata` (information about the response). The helper and all cell-read code are visible in the [executed script](man-group-run/run.mjs). The surrounding JSON fields are our recording format, not an additional spreadsheet API.

## Same inspect function, different settings {#inspection-settings}

The settings answer three questions: **what to return** (`kind`), **where to look** (`sheetId` and `range`), and **how much to return** (`maxChars` and preview sizes). `kind` is the output choice within one function; it is not a different tool.

| Setting for kind | What comes back | What it tells the agent |
| --- | --- | --- |
| `workbook,sheet` | Workbook counts and one description per tab | Which tabs exist and their occupied ranges; no cell values. |
| `workbook,sheet,table` | Those descriptions **plus cell-value grids** | Potentially every value in each occupied rectangle. It can be large. |
| `region` | Located blocks, their dimensions and small cell previews | How a selected area breaks into populated blocks, and a sample of each block’s contents. |
| `match` | Cells matching supplied text or a pattern | Candidate locations for the agent to inspect and interpret. |

Here **table** means a grid of values returned by inspection. A formal **Excel Table** is a named feature defined inside a sheet. This workbook has none. **Region** means a populated block identified by the inspection call; it is not another tab.

```component inspection-settings-lab
renderer: inspect_explainer.py
data: evidence/inspect-lab.json
---
lab
```

**The preview setting has a catch.** For `kind: table`, the tested runtime uses the row/column preview sizes as a fallback when the values would be too large. They are not unconditional caps. A table that fits can be returned in full. For the tested `region` calls, each returned block’s preview follows the selected dimensions.

**Two kinds of “cut short” are different.** `valuesTruncated: true` means a table’s values were reduced to a preview; the table’s full dimensions can still be reported. The outer `truncated: true` means whole response records were omitted. The small Fee engine preview has the first flag; the original broad workbook read has the second.

The agent still interprets the data. A list of sheets is a structural overview, not a written explanation of the business model. A search for “profit” gives candidate cells, not a decision about which profit measure is appropriate.

## Individual calls and exact evidence {#calls}

These are separate examples of using the operations above. The five discovery examples all use `inspect` with different settings. The controls show **Readable** explanations, the **Exact input**, and either readable results, an expandable **Object tree**, or **Exact JSON**. JSON is structured text containing named fields and lists.

```component call-explorer
renderer: man_group_walkthrough.py
data: evidence/man-group-cards.json
---
explorer
```

#### What the edit sequence reveals — formula import and recalculation {#findings}

**Reading a result and having a working calculation chain are different things.** The file contains 495 shared-formula follower cells: Excel stores those formulas compactly by referring to a neighbouring master formula. In this library version, some are imported with their saved values but without formula text or working upstream dependencies. The dependency tree stops at those cells.

The temporary edit makes the consequence visible. Changing Q4 2027’s base fee margin from 56bp to 51bp initially leaves management fees unchanged, even after `recalculate()`. Restoring the existing active-assumption formula makes fees change, while operating profit remains unchanged because other dependencies are still missing. Those outcomes are shown in the [cell edit](#xio-value-edit), [recalculation](#xio-recalculate) and [formula edit](#xio-formula-edit) examples. This is a limitation of the imported model in this library, not evidence that native Excel cannot calculate it.

**Inspection responses also need interpretation.** The original broad read is truncated; the dedicated sheet list is complete. The settings comparison above makes that difference visible. The agent has to notice whether values were sampled or whole records omitted before choosing its next call.

The original download is unchanged. The temporary assumption was restored before export; independent file reads found no formula/literal differences or cached numeric differences above tolerance, preserved sheet order and 270 annotations. The exported copy remains a demonstration, with the broader import limitation still present. Native Excel opening was not tested in this run.

[Original copied workbook](man-group-run/source.xlsx) · [Exported demonstration](man-group-run/man-group-demo.xlsx) · [File verification](man-group-run/verification.json) · [Evidence manifest](MANIFEST.md)

Runtime: `@oai/artifact-tool` 2.8.59, workspace bundle 26.909.12148. Everything here is local; no changes were pushed to main.

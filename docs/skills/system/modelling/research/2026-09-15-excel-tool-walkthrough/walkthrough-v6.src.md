---
title: How Codex works with Excel files
class: system-doc
docver: 6
primary_author: RG
agent_assistance:
  - Codex
subtitle: Real operations, agent-written checks, and the instructions that shape both
history:
  - v6 2026-09-15 — Workbook lifecycle, 27 additional operations and explicit verification provenance
  - v5 2026-09-15 — Operation map and measured inspection-settings comparison
stub:
  - Follow the code from an agent’s request through an in-memory workbook to a saved Excel file.
---

**Codex writes a program that works on a workbook in memory, then saves an Excel file.** The library supplies the operations. The spreadsheet skill supplies working instructions. The agent chooses the formulas, reads, checks and sequence of actions. Understanding all three explains much more than looking at a single “tool response”.

These examples use the supplied Man Group workbook and a small new workbook. The new **Fee lab** tab is an illustrative testing area. Its assumptions are chosen to make behaviour observable, not to forecast Man Group. Every example has readable inputs and outputs, exact code and expandable evidence. The page does not execute spreadsheet code in your browser.

## From instructions to a saved file {#layers}

| Part | What it does | What it does not decide for you |
| --- | --- | --- |
| **Spreadsheet skill** | Written guidance the agent reads: inspect first, preserve an existing model, use formulas, test changes, review images and export. | It is not an automatic checker running inside every operation. |
| **Agent’s script** | Chooses cell addresses, financial rules, tests and when to save. Runs JavaScript through Codex’s command tool. | Its before/after report is only as complete as the checks it actually requested. |
| **Artifact Tool library** | Reads and changes workbook objects, calculates formulas, searches, traces, renders and exports. | It does not infer the correct financial meaning of an input or certify the whole model after an edit. |

The functions here come from **`@oai/artifact-tool`**, a JavaScript library. They are not separate Excel MCP tools and do not control the Microsoft Excel window. **MCP** is a protocol for exposing tools to an agent; the outer tool in these examples is a command runner that starts the script.

| Starting point | Working stage | When the Excel file changes |
| --- | --- | --- |
| **New workbook** | `Workbook.create()`, add tabs, write values and formulas. | After `exportXlsx(...)` creates the file object and `.save(...)` writes it. |
| **Existing workbook** | Import the XLSX, then edit its in-memory representation. | After exporting and saving the revised copy. Editing memory alone leaves the existing file unchanged. |

An Excel window already open on your computer is not connected to this working copy. Open the saved output to see its changes. An agent can save repeatedly as it works. Direct control of a live Excel session is a different route, outside these examples.

**The six-cell checks in the original examples were written by me.** I selected the base margin, active margin, AUM, quarterly fees, annual fees and operating profit. `readState()` reads those addresses before and after an operation. The edit and `recalculate()` do not return those six rows automatically. Six is neither a tool setting nor a complete dependency count. The revised examples show **Library operation** and **Agent-written checks** separately.

#### Instructions that materially shape the work — part of the agent’s input {#guidance}

The installed skill and API guide do more than explain syntax. They influence the agent’s chosen workflow and model design. These are the most relevant instructions for interpreting the examples. The [saved documentation excerpts](guidance.md) preserve the source wording, version and original locations.

| Guidance the agent reads | Work the agent must actually perform | Example or implication |
| --- | --- | --- |
| **Inspect before editing; bound large reads.** | Choose areas, character limits and searches, then notice truncation. | The broad workbook read can consume thousands of tokens without describing the model’s economics. [Compare settings](#inspection-settings). |
| **Preserve the existing model and keep edits targeted.** | Read nearby formulas and formatting; extend rather than arbitrarily redesign. | Adding Fee lab leaves the original six tabs intact. This behaviour comes partly from instructions. |
| **Test input changes and affected outputs.** | Select representative cells, write reads and decide expected results. | Our six-cell comparison and separate arithmetic are custom code implementing this general instruction. |
| **Distinguish blank from zero.** | Choose the model’s rule for missing information, write it and test it. | A blank rate initially produces zero; an explicit guard makes it #N/A, while deliberate zero remains zero. [Experiment](#op-blank-input). |
| **Use readable formulas and a coherent build.** | Choose formula structure, common assumptions and relative/absolute references. | The skill discourages introducing LET/LAMBDA and prefers one active forecast build. These are design instructions, not proof those features are impossible in the engine. |
| **Bind charts to cells; render and review.** | Set source references, render an image, open that image and judge it. | Creating a chart returns an object; rendering returns pixels. Neither step automatically performs visual review. |
| **Recalculate before final checks and export.** | Call the engine, inspect results and save. | Ordinary edits already trigger calculation in the new examples. Final recalculation is an additional workflow instruction. |
| **Check that formal Table ranges do not overlap.** | Inspect existing ranges before creating another Table. | The API accepts our deliberately overlapping Table in memory. The agent must enforce this documented constraint. [Experiment](#op-overlap-table). |
| **Verify saved or native features when relevant.** | Reopen the file, compare it, and use the intended application when its behaviour matters. | We check saved formulas, values, tables and chart parts. Native Microsoft Excel opening is not established here. |

User instructions and the supplied workbook take precedence over these defaults. An agent may follow the guidance well or poorly. A repeated modelling choice should therefore be attributed to the instructions, the agent’s judgment or the library only after inspecting the evidence.

## Create, change, delete and rebuild {#operations}

The 27 recorded steps cover six families below. Each distinguishes the **operation’s own return** from the **additional observations requested by the agent**. “Undefined” means no returned report; it does not mean that nothing happened.

```component operations-laboratory
renderer: operations_explorer.py
data: evidence/operations-cards.json
---
operations
```

**What these experiments establish.** Fresh formulas update after an input edit. Unsaved edits do not reach the file. Clearing a formula can leave a plausible zero total. Deleting a formal Table leaves its cells behind. Deleting a chart leaves its source calculations behind. Rebuilding either is another explicit instruction sequence.

The added schedule works after saving and reopening. The earlier Man Group shared-formula import issue remains in the original model: some existing formulas import as saved values without a working dependency chain. A new working calculation does not repair that issue, and this is not evidence of a calculation defect in native Excel.

[Download Man Group with the Fee lab](operations-run/man-group-operations.xlsx?download) · [Download the newly created workbook](operations-run/new-fee-workbook.xlsx?download) · [Independent file verification](operations-run/independent-verification.json) · [Full analysis notes](LEARNINGS.md)

## Reading and inspecting a workbook {#inspection-settings}

The same `inspect` function supplies several kinds of evidence. A tab inventory, a grid of values, a populated-block preview and a text search are different requests—not four separate tools. The returned **values** are cell results; exact formula text can be read separately through `.formulas`.


The settings answer three questions: **what to return** (`kind`), **where to look** (`sheetId` and `range`), and **how much to return** (`maxChars` and preview sizes). `kind` is the output choice within one function; it is not a different tool.

| Setting for kind | What comes back | What it tells the agent |
| --- | --- | --- |
| `workbook,sheet` | Workbook counts and one description per tab | Which tabs exist and their occupied ranges; no cell values. |
| `workbook,sheet,table` | Those descriptions **plus cell-value grids** | Potentially every value in each occupied rectangle. It can be large. |
| `region` | Located blocks, their dimensions and small cell previews | How a selected area breaks into populated blocks, and a sample of each block’s contents. |
| `match` | Cells matching supplied text or a pattern | Candidate locations for the agent to inspect and interpret. |

Here **table** means a grid of values returned by inspection. A formal **Excel Table** is a named feature defined inside a sheet. The original six-tab Man Group file has none. The new Fee lab experiment deliberately creates one. **Region** means a populated block identified by the inspection call; it is not another tab.

```component inspection-settings-lab
renderer: inspect_explainer.py
data: evidence/inspect-lab.json
---
lab
```

**The preview setting has a catch.** For `kind: table`, the tested runtime uses the row/column preview sizes as a fallback when the values would be too large. They are not unconditional caps. A table that fits can be returned in full. For the tested `region` calls, each returned block’s preview follows the selected dimensions.

**Two kinds of “cut short” are different.** `valuesTruncated: true` means a table’s values were reduced to a preview; the table’s full dimensions can still be reported. The outer `truncated: true` means whole response records were omitted. The small Fee engine preview has the first flag; the original broad workbook read has the second.

The agent still interprets the data. A list of sheets is a structural overview, not a written explanation of the business model. A search for “profit” gives candidate cells, not a decision about which profit measure is appropriate.


### Original step-by-step sequence — exact calls and added checks {#calls}


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

[Original copied workbook](man-group-run/source.xlsx?download) · [Exported demonstration](man-group-run/man-group-demo.xlsx?download) · [File verification](man-group-run/verification.json) · [Evidence manifest](MANIFEST.md)

Runtime: `@oai/artifact-tool` 2.8.59, workspace bundle 26.909.12148. These are recorded experiments, not live workbook controls.

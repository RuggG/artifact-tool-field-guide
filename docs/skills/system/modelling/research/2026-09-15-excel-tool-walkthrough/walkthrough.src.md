---
title: What Codex can do with Excel files
class: system-doc
docver: 7
primary_author: RG
agent_assistance:
  - Codex
subtitle: An operations reference, with inspectable inputs, outputs and agent-written checks
history:
  - v7 2026-09-15 — Reorganised by capability, with one reference and supporting recorded examples
  - v6 2026-09-15 — Workbook operations, persistence and verification provenance
stub:
  - Choose an operation to understand its inputs, returns and what the agent must still decide.
---

**The agent works with a workbook in memory, using a JavaScript library called Artifact Tool.** It can inspect cells, write formulas, calculate, format, create charts and save an `.xlsx` file. This page maps those operations and separates what the library does from the checks the agent writes around it.

**Create or import → read and change in memory → export and save a file.** An existing file changes only when saved. These examples do not control a live Microsoft Excel window; that is a separate interaction route.

## Operations {#capabilities}

Choose what you want to do. Each entry explains **what you supply, what comes back, and what the agent still decides**. Open its recorded examples for readable inputs and outputs, exact code, and expandable data. Different settings of `inspect` are grouped by the information they return.

```component operations-reference
renderer: operations_reference.py
data: evidence/operations-reference.json
---
reference
```

## What the agent adds {#guidance}

**An operation is not the whole workflow.** The spreadsheet skill is a local instruction document the agent reads. It recommends inspection, targeted edits, formula checks, visual review and final export. The agent translates that guidance into specific code. The library does not automatically perform all those checks.

| Library capability | Guidance the agent reads | Agent’s own work |
| --- | --- | --- |
| Read, search and trace | Inspect before editing; keep reads manageable. | Choose ranges and search terms, notice missing data, interpret financial meaning. |
| Write values and formulas | Preserve the model; test input changes and affected results; distinguish blank from zero. | Choose assumptions and formulas, select outputs to check, calculate expected answers. |
| Create Tables and charts | Avoid overlapping Tables; bind charts to cells and review a render. | Check existing ranges and decide whether the layout or chart is useful. |
| Recalculate, export and save | Recalculate and verify the finished output. | Read results, scan errors, reopen or compare files, and use native Excel when its behavior matters. |

The **six-cell before/after reports** in some examples were our choice: the script reads six particular addresses. They are not an automatic dependency report, a tool limit, or proof that every formula works. The same distinction applies to our file comparisons and separate arithmetic checks. Exact check code is available beside the examples that use it.

[Read the saved skill and documentation excerpts](guidance.md) for the instructions and their provenance. [Run the script](#ref-runner) explains the outer Codex tool call: the spreadsheet functions themselves are library calls, not separate Excel MCP tools.

## Evidence and scope {#evidence}

The reference covers operations exercised in the supplied **Man Group** workbook and a small new workbook, using Artifact Tool **2.8.59**. It is a tested selection, not a complete API catalogue. The added **Fee lab** sheet is an illustrative calculation, not a new Man Group forecast. The controls display saved recordings rather than executing spreadsheet edits.

**One material limit:** some shared formulas in the original Man Group file imported as saved values without working dependencies. That affects what reads, traces and recalculation can establish. The [dependency example](#ref-trace) and [original-model input edit](#ref-values~old-value-edit) show it directly. It is a limitation of this imported representation, not evidence of a defect in native Excel.

#### Files and detailed findings — downloads, checks and analysis {#files}

[Download Man Group with the Fee lab](operations-run/man-group-operations.xlsx?download) · [Download the small new workbook](operations-run/new-fee-workbook.xlsx?download)

[Independent saved-file verification](operations-run/independent-verification.json) · [Detailed analysis notes](LEARNINGS.md) · [Evidence manifest](MANIFEST.md)

The original download remains unchanged. Independent file reads checked preserved formulas, literals, saved values, sheet order and selected workbook features. Native Microsoft Excel opening was not tested in these operations experiments. The earlier export and the later Fee lab output are separate copies, with their own saved verification records.

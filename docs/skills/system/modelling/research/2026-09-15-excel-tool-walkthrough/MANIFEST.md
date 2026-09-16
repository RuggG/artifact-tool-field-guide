# Evidence manifest

The original sequence and inspection comparisons use the supplied `/Users/ruggerogargiulo/Downloads/man-group-fees-excel.xlsx`, copied to [source.xlsx](man-group-run/source.xlsx). Source SHA-256: `b7ff5c92bf7562fa6fe3f253fdf7b40fd5f6659e8638310314c6b035ef02c181`. The Downloads original remains unchanged.

[The executed script](man-group-run/run.mjs) imports the copy with `@oai/artifact-tool` 2.8.59 and records actual requests/results in [records.json](man-group-run/records.json) and individual JSON files. [The command-tool receipt](man-group-run/runner-tool-record.json) preserves the real invocation, initial running-session response and completion response. Elapsed milliseconds inside the recorder cover each library operation, not research or agent thinking time.

[The page data](evidence/man-group-cards.json) adds explanations to these captures. Readable tables round numeric display to at most four decimals; the object trees and raw response preserve full precision. NDJSON is decoded into records for the tree view and retained as the original string in Exact JSON. Worksheet labels added to the rectangular read are identified as presentation context. Mutations have explicit before/after reads; they do not pretend that the setter returns a success report. The embedded PNG is the unmodified library render, also viewed through the actual image-view tool.

All edits were temporary diagnostics in the copy. `copyFrom` restores the original base assumption before final recalculation and export. [Independent verification](man-group-run/verification.json) compares source/export formulas and literal cells, numeric caches (absolute tolerance 1e-8 or relative 1e-10), sheet order and annotation count. There are no formula/literal differences, no cache differences outside tolerance, and 270 annotations in each file. ZIP integrity checks pass. This does not establish native Excel behavior; native Excel was not opened in this run.

The imported workbook contains 495 shared-formula follower cells (30 on Assumptions, 465 on Fee engine). The trace and edit demonstrations preserve this observed library limitation rather than silently repairing all formulas. The initial quoted-sheet trace attempt is retained separately in [its receipt](man-group-run/trace-quoted-sheet-attempt.json); the displayed successful call uses the unquoted sheet name accepted by this API.

Readable input descriptions are presentation added by `man_group_walkthrough.py`; Exact input retains the captured JavaScript or command arguments. No spreadsheet operations were rerun for this display change.

Page source: `walkthrough.src.md`; page-local renderer: `man_group_walkthrough.py`; supporting local CSS: `man-group.css`; card preparation: `prepare_man_group.py`. The older `evidence/cards.json`, `prepare.py`, `walkthrough.py` and prior assets belong to version 1 and are not used by the current page. The version 6 publication adds the page and its supporting evidence to the Hub reader.

Terminology is authored in [terms.json](evidence/terms.json), beside the calls where it matters. The table/worksheet distinction was checked against the original XLSX package and captured overview; [the verification](evidence/terminology-verification.json) records zero formal Excel Table definitions and zero worksheet table references, alongside the actual inspection counts and Summary rectangle. The page describes the observed inspection behavior without asserting a general segmentation algorithm. Exact call inputs and outputs are unchanged.

Version 5 adds [the inspect settings dataset](evidence/inspect-lab.json), prepared by prepare_inspect_lab.py and displayed by inspect_explainer.py. Nine presets bind actual calls to comparisons of output kind, character allowance and preview size. The original fifteen captures remain intact. The read-only [probe script](overview-probe.mjs) also ran a larger 4 × 6 region preview for the paired comparison. Every preview-size and allowance claim was checked against the recorded results. Current token estimates are computed from each saved response; internal ID changes can move them slightly between imports.


## Version 6: operations and instruction provenance

The new [operations script](operations-run.mjs) runs 26 steps over a copy of Man Group and a separately-created small workbook. [Execution records](operations-run/records.json) separate the operation code, its return type/value, and agent-selected before/after observations. Returned live objects are described by type rather than claimed to be serialized whole. Undefined values have explicit recording markers. Public inspection fields, including NDJSON, are copied explicitly into the saved records.

[The new page data](evidence/operations-cards.json) adds authored intent, outcome and interpretation. Readable summaries use selected observations; Object tree and exact code preserve the fuller evidence. They do not imply that every dependency was checked. [Saved guidance](guidance.md) and [its provenance](evidence/guidance-provenance.json) distinguish instructions from runtime features.

[Independent output verification](operations-run/independent-verification.json) uses read-only openpyxl 3.1.5 and ZIP inspection. The original six sheets retain formulas/literals, cached values within the stated tolerance, number formats and their order. Fee lab is the only new sheet, with one formal Excel Table and an additional chart. The original Downloads file is byte-identical. No native Microsoft Excel session was opened in these experiments.

The initial recording defects and the corrected final run are described in [the analysis notes](LEARNINGS.md). Initial records remain in operations-run/recording-attempt-1. The current cards use the final rerun. The browser renders a saved demonstration; it never executes these workbook-edit scripts.

Runtime dependencies are not vendored into this publication. To rerun, use the installed spreadsheet skill and its bundled runtime, with a local node_modules symlink. Builders consume already-saved JSON; no workbook rerun is necessary to rebuild the page.

A separate [overlap probe](overlap-probe.mjs) brings the additional operation examples to 27. It is not saved into either downloadable XLSX. The [review record](review.json) distinguishes the full-page QA pass from the final overlap example’s manual desktop and phone review.


## Version 7: operations reference

The page now uses `operations_reference.py`, `operations-reference.js`, `operations-reference.css` and `evidence/operations-reference.json`. The catalogue is editorially organised by capability; its examples use the existing `operations-cards.json`, `man-group-cards.json` and `inspect-lab.json` evidence without rerunning or altering the workbook experiments. Some readable presentation helpers are reused from the earlier renderers. Object-tree expansion is rendered on demand in the browser. The source and evidence remain available; the previous source is preserved as `walkthrough-v6.src.md`.

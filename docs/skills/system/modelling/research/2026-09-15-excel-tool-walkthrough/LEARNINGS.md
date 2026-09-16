# Codex, Artifact Tool and Excel: formed view

Evidence date: 15 September 2026. Runtime: Artifact Tool 2.8.59, workspace bundle 26.909.12148. Supplied Man Group file is an example fixture, not independently validated company research.

## Three things shape the result

The library exposes workbook operations. The spreadsheet skill supplies instructions about how the agent should build and check a model. The agent interprets the task and writes code using those operations. A useful comparison of modelling environments must distinguish all three. A preference for one active forecast build or particular formulas can originate in the skill; it is not automatically an engine restriction.

The agent-written verification is real executed work, but is not a hidden automatic feature of the edit call. Cell selection, separate arithmetic, before/after comparisons, error search, image viewing and saved-file comparisons are explicit choices. Our original HTML combined these too closely with operation returns. The revised page separates them.

## Changes and persistence

The new standalone workbook calculates 125m in quarterly fees and changes to 150m after editing the input, without a manual recalculation. The added Fee lab formulas also update immediately. An explicit final recalculate is recommended by the skill before verification/export; it is not required after every ordinary edit in this runtime.

Saving a checkpoint then changing its in-memory margin from 56bp to 51bp leaves the file checksum and its saved 56bp value unchanged. The final export has a 60bp margin. A new import can be edited to 55bp and recalculates to 399.45356938289785m in the first quarter, while the file still has 60bp. No native Excel window participates.

## Meaning comes from the model and its author

Clearing all four fee formulas leaves SUM returning zero. Clearing the fee input while retaining the formulas also returns zero. These are examples of apparently ordinary values hiding missing calculations or assumptions. The agent has to choose a suitable check; searching for formula errors alone would miss them.

An explicit IF(ISBLANK(input),NA(),calculation) rule makes a blank input visible as #N/A, while an entered zero still produces valid zero fees. The skill recommends distinguishing blank from zero. It does not automatically insert this rule; the agent decides whether the distinction is appropriate for the model.

## Structural operations

Adding a worksheet, writing blocks of values, writing a seed formula and filling down, setting formats, clearing contents and rebuilding all worked in the new lab. Clearing content preserves formatting. The rebuild is code, not a native undo operation. We did not establish row/column insertion, sheet deletion or undo API coverage.

A named Excel Table can be created over cells and appended with a row. Deleting that Table object leaves the underlying cells intact; recreating the Table includes the appended row. This is materially different from the informal table-shaped inspection records in the original six-sheet file, which has no formal Excel Tables.

A range-backed chart uses cell references, updates its rendered values when the input changes, can be deleted without clearing the source cells, and can be recreated against those cells. Render returns an image blob. Saving and viewing the image are additional work. Independent ZIP checks identify two chart parts in the output: the pre-existing chart and the added lab chart.

A separate, deliberately overlapping Table request was accepted in a new in-memory import. The existing Table occupies B20:D24 and the added one D22:F25, sharing D22:D24. This matches the guide’s warning that overlap is not rejected by the API. The invalid experiment was not exported; the downloadable workbook retains only the valid Table.

## Original Man Group import issue remains a distinct finding

Fresh calculations working does not negate the original shared-formula import issue. The original copy contains 495 shared-formula follower cells; some read as saved values with missing formula/dependency information in this version. The prior edit tests demonstrated that effect. Adding an independent Fee lab does not repair those original formulas. Claims about this runtime are narrower than claims about Excel itself.

## Verification and recording

The final 26-step sequence and a separate overlap probe completed without runtime exceptions. Independent reads preserve the six original sheets' formulas/literals, cached values within tolerance and number formats. Downloads original is byte-identical. The new lab, formal Table and extra chart are intentional additions. Native Excel opening is not tested by these runs.

Our first recorder omitted non-enumerable inspection text and incorrectly treated a table-row helper as a Range. These were recording-code defects, not spreadsheet calculation failures. We corrected the reads, explicitly captured NDJSON and reran the affected demonstration. The initial raw JSON records remain in operations-run/recording-attempt-1. One intervening script-edit syntax error occurred before workbook execution; it was corrected before the final run. The published cards use the final records and exact executed code.

The extra checks reflect a documented practice but not one universally prescribed script. The page exposes their inputs so the reader can see precisely what was and was not tested.


## Capability reference — editorial review, 15 September 2026

Version 7 groups the existing evidence into 23 operations across six families. It replaces three overlapping explorers with one catalogue. An operation’s contract (input, own return, agent decisions) is the default read; recorded examples, variants, exact input/output and checks are supporting depth. The distinction between native returns and script-assembled observations is preserved, including the original wrapped examples. No spreadsheet experiment or result was changed in this editorial pass.

The interpretation is that the operative unit is often a library operation plus an agent-selected verification routine, informed by the skill. These are separable sources of behavior. A quiet mutation return does not certify success; a rich script report does not mean the mutation generated that report.

# Installed spreadsheet guidance: saved excerpts

Captured 15 September 2026 from workspace bundle 26.909.12148. These are instructions supplied to the agent, not executable guarantees. Line numbers refer to the original installed files. Full source-file SHA-256 values identify the edition.

## Verification

Source: `/Users/ruggerogargiulo/.codex/plugins/cache/openai-primary-runtime/spreadsheets/26.909.12148/skills/spreadsheets/SKILL.md`

SHA-256: `499172cadf77be41b6ec7d87981463aec0859e10aba9cc0b864143e42955474b`

Lines 327–333:

```text
### Verification Rules
Use Artifact Tool to verify requested features and results within the authorized changes and their affected dependencies. Match coverage to the scope, complexity and risk. Report unrelated pre-existing defects without repairing them. Reuse checks for unchanged content and keep authoring-only tests out of the delivered workbook.

After completing all edits, call `workbook.recalculate()` once before the final checks below and export. If you make further edits, recalculate again before repeating affected checks and exporting.
```js
workbook.recalculate();
```
```

Lines 364–368:

```text
3. Verify applicable recalculation in the intended engine. Test representative input changes and boundaries in a disposable copy or restore every temporary edit before delivery. Include blank versus zero, missing/duplicate keys, period cutoffs, overrides and rounding. For cases, change the selector and a later-period driver. Confirm the same build and linked outputs update while actuals remain unchanged. A blank unselected input must not block a valid active case; selecting that case must expose the missing input. Verify any agreed comparison refresh and stale-state behavior separately. Report any engine checks that could not be performed.

For workflows, check that required human inputs have editable fields and that completion guidance accounts for every prerequisite. Complete one prerequisite while leaving another open and confirm the remaining action stays visible. For input-driven rankings and action lists, change an input that should alter the order or included records and verify the list updates. Verify affected charts, status text, validation and conditional formatting react to edits. A saved value, static matrix or unchanged PASS cell is not recalculation proof.

4. Render sheets/ranges to verify visual output. Skip only when the rendered view and its data/formula dependencies are unchanged:
```

Lines 385–387:

```text
6. Inspect the saved file when an affected feature or export concern requires it. Verify requested or preserved native features in the intended engine, including any explicitly required Data Table input/output behavior. Check iteration and capture behavior separately when used.

Finalize only after successful export and the applicable checks. Report what was performed and any remaining limitations. Formula text, a preview and a successful export do not establish native-application behavior.
```

## Formula and model design

Source: `/Users/ruggerogargiulo/.codex/plugins/cache/openai-primary-runtime/spreadsheets/26.909.12148/skills/spreadsheets/SKILL.md`

SHA-256: `499172cadf77be41b6ec7d87981463aec0859e10aba9cc0b864143e42955474b`

Lines 129–140:

```text

## Formulas

Apply these rules to newly added or edited formulas and their affected dependencies. Follow the user's preferences and supplied template; preserve unrelated formulas and layout during narrow edits. Design formulas to support the workbook structure above: the reader should be able to follow the inputs, useful calculation steps and final answer.

### Formula Construction

- Use direct references, familiar functions and meaningful intermediate calculations. Follow [Build Structure and Formula Flow](#build-structure-and-formula-flow) to show the work; do not hide an entire build in one dense formula or add trivial helpers just to make formulas shorter.
- Keep raw data, editable assumptions, mappings and business rules in labeled cells or tables. Mathematical, index and control constants may remain in formulas. Keep calculated results as formulas so they update with their inputs.
- Fixed cutoffs or categories from the user's request can appear directly in formulas when result labels state the rule. For example, label `COUNTIFS(B2:B100,">1000")` as `Invoices over $1,000`, without adding an input cell for `1000`. Use one labeled input cell when the cutoff is user-adjustable or serves as a shared assumption across different calculations.
- Calculate a shared result once and reuse it when the inputs, period, units, rounding and overrides match. Keep independent reconciliations independent.
- Use consistent formulas across comparable rows and periods, while preserving intentional differences such as [historical versus forecast logic](domain_guidance/financial_models.md#periods-assumptions-and-scenarios), one-off adjustments and overrides.
```

Lines 175–189:

```text
- **Formula choices to avoid:** do not introduce `LET`, array/spill formulas, `MAP`, `REDUCE` or `LAMBDA`. Use familiar formulas and labeled intermediate steps. Normal range arguments in functions such as `SUMIFS` and `SUMPRODUCT` remain appropriate, as do the lookup, `INDIRECT`, `OFFSET` and `CHOOSE` patterns below. Preserve required existing/template behavior and do not rewrite unrelated formulas during a narrow edit. Formula length alone is not the test: the reader must be able to understand and extend the calculation.
- **Sensitivity analysis:** use a native What-If Data Table only for an explicitly requested native sensitivity analysis or required existing/template behavior, when supported. Do not introduce `TABLE` into an ordinary forecast or case comparison, or manufacture a second varying input with a metric selector. Excel supports one or two varying inputs; the current Artifact Tool supports only two-variable tables, with both input cells on the table’s worksheet. Read [Data Tables](artifact_tool_docs/DATA_TABLES.md) before creating one. Two inputs test one output across their combinations; use separate tables for additional outputs. If the requested native design is unsupported, explain the limitation before agreeing on a formula-based design or change-input/recalculate/restore process. Label captured results and their refresh method. Ordinary case comparisons follow the single-build and comparison boundary above.

An Excel Table, PivotTable and What-If Data Table are different features. Check the chosen tool and destination's support. If a requested native feature cannot be created or preserved, explain the limitation before substituting a formula or static result. Keep API setup and feature-specific execution details in the relevant tool reference.

### Scalable Formulas and Brief Explanations

Use the patterns below when they make recurring updates easier without hiding the calculation. Choose the simplest approach that supports the actual update workflow, not just the current snapshot.
- **Assumption and Case selection:** prefer one numeric Case selector with labeled case names and `CHOOSE` or `OFFSET` to select the active assumptions. `INDEX/MATCH` or `XLOOKUP` remain valid when they fit the layout. In each driver group, put Active Selection above its case inputs, sharing the same period header. For example, with Case in B3 and two case values in I22:I23, active I21 can be `=CHOOSE($B$3,I22,I23)`; the matching build input is simply `='Assumptions'!I21`. Anchor and validate the selector. Do not repeat the choice in the build or maintain a second editable copy of the drivers. The simple CHOOSE example assumes validated numeric case inputs. Otherwise, test the selected source value before a reference can turn a blank into zero. Preserve a valid zero. A missing unselected case must not block the active case. In an agreed comparison, mark only the affected case and dependent deltas unavailable. Keep necessary validation local to the driver and reuse it. OFFSET and INDIRECT are volatile, so keep references bounded and consider recalculation cost.
- **New monthly source tabs:** if the workflow receives a separate tab in the same format each month, a visible month-to-tab registry and bounded `INDIRECT` references can support new periods without rewriting the reference pattern. Register the new tab and extend the summary periods or bounded ranges when needed. Validate the expected layout, tab names and source coverage; quote and escape sheet names correctly. For a new workflow without that constraint, one source table with a Month column may be simpler.
- **Explain recurring updates:** when a less familiar formula materially improves the workbook, add a short explanation near its control or in the existing guide: why it helps, what the user can change and how to extend it safely. For example: “Add the new month tab in the same layout and register its name in Setup. Extend the summary period and ranges if needed; the formulas keep the same reference pattern.” Keep this brief; do not add comments to every formula or create a new instruction tab for one note.

### Missing Inputs, Errors and Overrides

- Do not invent missing source data or substitute a different metric. If required data is absent, leave the result unavailable and state the specific missing input beside its data or setting and briefly in the response. For a rate, preserve the requested numerator, denominator, population and period; do not substitute another available denominator.
```

## Preserving an existing workbook

Source: `/Users/ruggerogargiulo/.codex/plugins/cache/openai-primary-runtime/spreadsheets/26.909.12148/skills/spreadsheets/workflows/edit_workflows.md`

SHA-256: `f75917488c75b5a7b9e194c57131163d54ad23405802944e1d59eece990d27d8`

Lines 1–24:

```text
# Editing an existing spreadsheet

Treat existing workbook as the reference: preserve its structure, formula patterns, formatting, terminology, and navigation and extend nearby conventions unless the user explicitly requests a change.

User requests always take priority over any rules in this file.

## Safety rules
- Do not add, remove, rename, reorder, or split tabs unless requested or required.
- Before modifying: ALWAYS study and match the existing format, style and conventions when making edits by rendering and viewing the image. Read related values and formulas.

## Important guidelines
- Prioritize consistency unless it conflicts with user request: Ensure existing formulas, layouts, structures, and patterns are consistent. For example, if asked to add another column or row to a table and there is conditional formatting applied to the whole table, it should extend to the new column or rows as well.
- Keep edits targeted unless a broader change is clearly necessary. Exceptions are when there's dependencies, e.g. a dynamic chart that is based on the range of values in a table and a new row is added, the chart should also update.
- Change only requested cells and directly affected formulas/charts. Preserve unrelated tabs, formulas, formatting, validations, named ranges, comments, protection, hidden/grouped rows/columns, freeze panes and chart content. Do not add sheets, rows, columns or helpers unless requested.
- Never overwrite formatting for spreadsheets with established formats, unless requested or to extend an added range.
- Preserve native tables, structured references, pivot sources, shared formulas, filters, external links, INDIRECT routing and calculation/iteration settings. Do not flatten formulas or rebuild unrelated features for convenience.
- For visual fixes, start with the smallest plausible local change. Do not apply sheet-wide autofit, wrapping, or restyling unless requested.

## Formula Rules

For formula edits, follow the user's request first; otherwise preserve valid existing conventions before applying defaults for new formulas.

- Preserve valid existing formula conventions, structure and references to editable inputs unless the user requests otherwise. Look at a couple examples in the requested edit area before making changes.
- If there are any errors with the original workbook, unrelated to the task at hand, do not fix them arbitrarily. Instead, summarize the issues to the user and ask if they want them fixed.
```

## Automatic calculation and bounded discovery

Source: `/Users/ruggerogargiulo/.codex/plugins/cache/openai-primary-runtime/spreadsheets/26.909.12148/skills/spreadsheets/artifact_tool_docs/API_QUICK_START.md`

SHA-256: `2795943c6a356a6bb94767526acb53a6aec8599be1d15fdbe54458bc799d5f4a`

Lines 33–45:

```text
## Build Patterns
- Prefer block writes (`range.values`, `range.formulas`) over per-cell loops. Normally match the matrix shape to the target range (e.g. "D4:M4" → 1×10). To intentionally expand a larger matrix from an anchor, target a single cell or use range.write(matrix).
- `range.values = [[value]]` repeats that value across the range; range.formulas does not broadcast a single formula. For merged ranges, write to the top-left cell or `mergedRange.values = [[value]]` where `mergedRange` is a single merged range.
- Seed scalar formulas once, then `fillDown()` / `fillRight()`. For dynamic-array formulas (like `SEQUENCE`, `UNIQUE`, `FILTER`, `SORT`, `VSTACK`, `HSTACK`), write only the anchor cell and let the result spill after.
- Use `range.displayFormulas` plus `range.formulaInfos` when you need to understand a spill child or a data-table output cell.
- Setting cell values or formulas automatically recalculates dependent formulas; no manual recalculation is needed after ordinary edits, including batches (exception: creating data tables). Call `workbook.recalculate()` once after all edits, before final verification and export. If you make further edits, recalculate again before rechecking and exporting.
- Prefer real `Date` objects for sortable/charted/formula date columns.
- Number and date formats must be applied explicitly (for example `yyyy-mm-dd`).
- Use JSON-serializable values for non-Date cells: `string | number | boolean | null`.
- If a cell is intended to display literal text that begins with `=`, write it as a value prefixed with a single quote (for example `'=B2*C2`). This includes formula descriptions, validation examples, and labels; do not write these cells through `range.formulas`.
- Create every worksheet referenced by formulas before writing any cross-sheet formulas.
- Verify with `await workbook.inspect(...)`; use `workbook.help(...)` only when the quick surface below is insufficient.
- `render` can be used to examine an existing workbook visually and for visual verifications.
```

Lines 81–87:

```text
## Reading existing/imported workbooks
- On existing/imported workbooks, get a compact summary via `inspect` to understand what already exists and where.
- Prefer `inspect(...)` for workbook understanding and discovery across broad areas.
- Prefer direct getters like `range.formulas` when you already know the target range and need the exact rectangular formula matrix.
- If formula locations are unknown, prefer `inspect({ kind: "formula", ... })` over reading `range.formulas` across a very large area.
- Prefer to set `maxChars`, `tableMaxRows`, `tableMaxCols`, and/or `maxResults` to prevent large dumps of data.
- For suspicious or high-impact outputs, use `workbook.trace("Sheet!A1")` to audit the dependency tree from final output/check cell back to source cells. Trace output can be large, so summarize by depth/node count before logging.
```

Lines 274–282:

```text
- When adding new tables, set explicit unique names (`TasksTable`, `SummaryTable`).
- Do not overlap tables; the API does not reject overlaps. Check existing table ranges in the initial compact `inspect` summary before adding one.
- `const table = sheet.tables.add("A1:H200", true, "TasksTable")`
- `table.rows.add(null, [[...], ...])`, `table.getDataRows()`, `table.getHeaderRowRange()`
- Read tables: `sheet.tables.items` -> `Table[]`
- Set + Getters: `table.name`, `table.style`, `table.showHeaders`
- Toggles for table utilities (set/get): `table.showTotals`, `table.showBandedColumns = true`, `table.showFilterButton`
- `table.delete()`

```

## Charts and verification

Source: `/Users/ruggerogargiulo/.codex/plugins/cache/openai-primary-runtime/spreadsheets/26.909.12148/skills/spreadsheets/features/charts.md`

SHA-256: `f85e0fa9ba2505f9ee3dd9c4529d1c529b0ce7c9c5e95a919409025c0217c470`

Lines 11–16:

```text
These are defaults, not restrictions; choose what makes the comparison or relationship clearest:

- Comparisons/rankings: bar or column; sort rankings by descending value.
- Time trends: line in chronological order; area only when the filled volume adds meaning.
- Distributions: histogram for shape; box-and-whisker for median, spread, and outliers. If unsupported in the available API, use formula-backed bins with bars or a summary table.
- Relationships between numeric variables: scatter.
```

Lines 37–43:

```text
- Use the workbook or supplied brand palette consistently for meaning; match the same business item across charts and tables, and preserve meaningful warning/status colors. Keep single-series charts one color unless highlighting meaningful differences. Use a contrasting neutral or palette color for reference/target series; add a dashed line when it helps distinguish the reference. For line charts, set the line stroke explicitly; `series.fill` alone may color a preview without setting the exported line. With the artifact tool, use `series.line = { fill: color, style: 'solid', width: 2 };` (or `style: 'dashed'` for a reference). Avoid unnecessary gradients, heavy borders, excessive gridlines, and decorative effects unless requested; retain useful reference lines.

## Verify and repair

- Before creation, verify source ranges/formulas, headers, categories, series names/values, orientation, and point counts. Repeat these checks after creation or editing, and verify expected chart count and bindings. Exclude headers from values; catch stale/disconnected ranges, formula errors, and categories misread as series.
- Render and inspect at normal zoom: correct values/units, readable titles, ticks, labels, and legends; no blank charts, clipping, crowding, duplicates, or overlaps.
- Verify saved/exported series colors and line strokes against the intended palette using the exported file's chart properties or target application; a correctly colored authoring preview is not enough.
```
